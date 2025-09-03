"""Core DOCX to DITA conversion logic.

Extracted from orlando_toolkit.core.converter.docx_to_dita module
for use in the DOCX plugin architecture.
"""

from __future__ import annotations

import logging
import os
import time
import uuid
import warnings
from datetime import datetime
from typing import Any, Dict, Optional, Callable

from docx import Document  # type: ignore
from lxml import etree as ET  # type: ignore

from orlando_toolkit.core.models import DitaContext
# Removed ConfigManager import - plugin manages its own configuration

from ..utils.docx_parser import extract_images_to_context, iter_block_items
from ..utils.style_analyzer import build_style_heading_map
from .formatting_helpers import (
    STYLE_MAP,
    add_orlando_topicmeta,
)
from .structure_analyzer import (
    build_document_structure,
    determine_node_roles,
    generate_dita_from_structure,
    _add_content_to_topic,
)

logger = logging.getLogger(__name__)


def convert_docx_to_dita_internal(file_path: str, metadata: Dict[str, Any], 
                                 plugin_config: Dict[str, Any] = None,
                                 progress_callback: Optional[Callable[[str], None]] = None) -> DitaContext:
    """Convert a DOCX file into an in-memory DitaContext.

    Processes Word document structure, formatting, images, and metadata
    to generate Orlando-compliant DITA topics and ditamap.
    
    This is the internal conversion function used by the DOCX plugin.
    
    Args:
        file_path: Path to the DOCX file to convert
        metadata: Conversion metadata and configuration options
        plugin_config: Plugin-specific configuration settings
        
    Returns:
        DitaContext containing the complete DITA archive data
        
    Raises:
        Exception: If conversion fails
    """
    logger.info("Starting DOCX->DITA conversion (plugin)...: %s", file_path)

    context = DitaContext(metadata=dict(metadata))

    try:
        if progress_callback:
            progress_callback("Loading DOCX file...")
        logger.info("Loading DOCX file...")
        doc = Document(file_path)
        all_images_map_rid = extract_images_to_context(doc, context)

        if progress_callback:
            progress_callback("Extracting images...")
        logger.info("Extracting images...")

        map_root = ET.Element("map")
        map_root.set("{http://www.w3.org/XML/1998/namespace}lang", "en-US")
        context.ditamap_root = map_root

        map_title = ET.SubElement(map_root, "title")
        map_title.text = metadata.get("manual_title", "Document Title")

        add_orlando_topicmeta(map_root, context.metadata)

        # Get plugin-specific configuration
        if plugin_config is None:
            plugin_config = {}
        docx_conversion_config = plugin_config.get('docx_conversion', {})
        
        # Extract color conversion rules for formatting helpers
        color_rules = plugin_config.get('docx_color_conversion', {})

        # Unified structural inference flag (backward-compatible):
        # Priority order for enabling: metadata.enable_structural_style_inference →
        # metadata.use_structural_analysis → metadata.use_enhanced_style_detection →
        # Note: Legacy structural analysis configuration is no longer used.
        # The intelligent classifier handles all style detection internally.

        # Intelligent style classification (replaces all legacy detection methods)
        if progress_callback:
            progress_callback("Detecting and analyzing document styles...")
        
        style_t0 = time.perf_counter()
        style_heading_map = build_style_heading_map(doc, plugin_config)
        style_ms = int((time.perf_counter() - style_t0) * 1000)
        
        # Apply user overrides from STYLE_MAP (user mappings take precedence)
        legacy_added = 0
        for k, v in STYLE_MAP.items():
            if k not in style_heading_map:
                style_heading_map[k] = v
                legacy_added += 1

        # User override mapping (highest priority): allow overrides and additions
        if isinstance(metadata.get("style_heading_map"), dict):
            user_overrides = metadata["style_heading_map"]  # type: ignore[arg-type]
            added_or_overridden = 0
            for k, v in user_overrides.items():
                try:
                    prev = style_heading_map.get(k)
                    if prev != v:
                        added_or_overridden += 1
                    style_heading_map[k] = v
                except Exception:
                    continue
            logger.debug(f"Applied {added_or_overridden} user style overrides/additions")

        # Final style detection summary
        logger.info(
            "Style detection: intelligent=%s legacy_fill=%s total=%s | detection_ms=%sms",
            len(style_heading_map) - legacy_added, legacy_added, len(style_heading_map), style_ms,
        )
        logger.debug(f"Final style heading map: {len(style_heading_map)} styles detected")
        if logger.isEnabledFor(logging.DEBUG):
            for style_name, level in sorted(style_heading_map.items(), key=lambda x: (x[1], x[0])):
                logger.debug(f"  Level {level}: '{style_name}'")

        # Prepare topic generation (two-pass). Actual building starts after structure analysis below.
        logger.debug("Preparing topic generation (two-pass)...")

        # ======================================================================
        # TWO-PASS APPROACH: Build structure first, then generate DITA
        # ======================================================================

        # Pass 1: Build complete document structure
        if progress_callback:
            progress_callback("Analyzing document structure...")
        logger.info("Analyzing document structure...")
        root_nodes = build_document_structure(doc, style_heading_map, all_images_map_rid)

        # Pass 2: Determine section vs module roles
        if progress_callback:
            progress_callback("Determining section/module roles...")
        logger.info("Determining section/module roles...")
        determine_node_roles(root_nodes)

        # Pass 3: Generate DITA topics and map structure
        if progress_callback:
            progress_callback("Building topics...")
        logger.info("Building topics...")
        heading_counters = [0] * 9

        generate_dita_from_structure(
            root_nodes,
            context,
            context.metadata,
            all_images_map_rid,
            map_root,
            heading_counters,
            color_rules,
        )

        # ------------------------------------------------------------------
        # Fallback: no topics detected → create a single topic hosting all
        # content, titled after the DOCX filename (without extension).
        # ------------------------------------------------------------------
        if not context.topics:
            fallback_title = os.path.splitext(os.path.basename(file_path))[0] or "Document"
            file_name = f"topic_{uuid.uuid4().hex[:10]}.dita"
            topic_id = file_name.replace(".dita", "")

            from .formatting_helpers import create_dita_concept
            
            concept_root, conbody = create_dita_concept(
                fallback_title,
                topic_id,
                context.metadata.get("revision_date", datetime.now().strftime("%Y-%m-%d")),
            )

            # Feed all block items; underlying helper filters empties and
            # handles paragraphs, lists, images, and tables uniformly.
            all_blocks = [blk for blk in iter_block_items(doc)]
            _add_content_to_topic(conbody, all_blocks, all_images_map_rid, color_rules)

            # Reference in ditamap
            topicref = ET.SubElement(
                map_root,
                "topicref",
                {"href": f"topics/{file_name}", "locktitle": "yes"},
            )
            topicref.set("data-level", "1")

            topicmeta_ref = ET.SubElement(topicref, "topicmeta")
            navtitle_ref = ET.SubElement(topicmeta_ref, "navtitle")
            navtitle_ref.text = fallback_title
            critdates_ref = ET.SubElement(topicmeta_ref, "critdates")
            _rev_date_fb = context.metadata.get("revision_date") or datetime.now().strftime("%Y-%m-%d")
            ET.SubElement(critdates_ref, "created", date=_rev_date_fb)
            ET.SubElement(critdates_ref, "revised", modified=_rev_date_fb)
            ET.SubElement(topicmeta_ref, "othermeta", name="tocIndex", content="1")
            ET.SubElement(topicmeta_ref, "othermeta", name="foldout", content="false")
            ET.SubElement(topicmeta_ref, "othermeta", name="tdm", content="false")

            context.topics[file_name] = concept_root

    except Exception as exc:
        logger.error("Conversion failed: %s", exc, exc_info=True)
        raise

    if progress_callback:
        progress_callback("Conversion finished.")
    logger.info("Conversion finished.")
    return context