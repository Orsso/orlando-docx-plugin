"""Intelligent Word style classification for DITA document structure analysis.

This module provides intelligent classification of Word styles to distinguish between 
true heading styles (which should become DITA topics) and body text styles (which 
should remain as content). Uses semantic analysis and comprehensive style databases
to handle international Word documents and edge cases gracefully.
"""

from __future__ import annotations

from typing import Dict, Optional, Any
from docx.document import Document
from docx.enum.style import WD_STYLE_TYPE
import re
import logging

logger = logging.getLogger(__name__)


def build_style_heading_map(doc: Document, plugin_config: Optional[Dict[str, Any]] = None) -> Dict[str, int]:
    """Return a mapping of style names to heading levels using intelligent classification.
    
    This function identifies which Word styles should be treated as headings in the 
    document structure. It uses semantic analysis to distinguish between true heading
    styles and body text formatting styles.
    
    Args:
        doc: Word document to analyze
        plugin_config: Optional plugin configuration (reserved for future use)
        
    Returns:
        Dict mapping style names to heading levels (1-9) for legitimate headings only
    """
    logger.debug("Starting intelligent style classification")
    
    # Built-in Word styles that should never be headings
    builtin_non_heading_styles = {
        'Normal', 'No Spacing', 'Quote', 'Intense Quote', 'Subtle Emphasis',
        'Intense Emphasis', 'Strong', 'Subtle Reference', 'Intense Reference',
        'Book Title', 'List Paragraph', 'Body Text', 'Body Text 2', 'Body Text 3',
        'Body Text Indent', 'Body Text First Indent', 'Body Text First Indent 2',
        'List', 'List 2', 'List 3', 'List Bullet', 'List Bullet 2', 'List Bullet 3',
        'List Continue', 'List Continue 2', 'List Continue 3', 'List Number',
        'List Number 2', 'List Number 3', 'Caption', 'Figure Caption', 'Table Caption',
        'Image Caption', 'Footnote Text', 'Footnote Reference', 'Endnote Text',
        'Endnote Reference', 'Bibliography', 'Index 1', 'Index 2', 'Index 3',
        'Index 4', 'Index 5', 'Index 6', 'Index 7', 'Index 8', 'Index 9',
        'TOC 1', 'TOC 2', 'TOC 3', 'TOC 4', 'TOC 5', 'TOC 6', 'TOC 7', 'TOC 8',
        'TOC 9', 'TOC Heading', 'Table of Contents', 'Header', 'Footer', 'Page Number',
        'Compact', 'Plain Text', 'HTML Preformatted', 'Document Map', 'Hyperlink',
        'FollowedHyperlink', 'Macro Text', 'Block Text', 'Date', 'Salutation',
        'Signature', 'Closing', 'Sans Interligne', 'Texte Normal', 'Citation', 'Legende'
    }
    
    # Pattern for built-in headings
    builtin_heading_pattern = re.compile(r"^Heading\s+(\d+)$", re.IGNORECASE)
    # Pattern for numeric-leading styles, e.g. "1.2 Chapitre", "2-3 Section"
    numeric_leading_pattern = re.compile(r"^\s*(\d+)([.\-]\d+){0,8}\b")
    
    heading_map = {}
    total_paragraph_styles = 0
    
    try:
        for style in doc.styles:
            if style.type == WD_STYLE_TYPE.PARAGRAPH and style.name:
                total_paragraph_styles += 1
                style_name = style.name
                style_lower = style_name.lower().strip()
                
                # 1. Built-in Word heading styles (highest confidence)
                builtin_match = builtin_heading_pattern.match(style_name)
                if builtin_match:
                    level = int(builtin_match.group(1))
                    if 1 <= level <= 9:
                        heading_map[style_name] = level
                        logger.info(f"Style Classification: '{style_name}' -> HEADING Level {level} "
                                   f"(Built-in Word heading style)")
                        continue
                
                # 2. Numeric-leading styles (common corporate heading conventions)
                num_match = numeric_leading_pattern.match(style_name)
                if num_match:
                    # Level is number of numeric segments (cap to 6)
                    matched = num_match.group(0)
                    segments = re.split(r"[.\-]", re.sub(r"^\s+", "", matched))
                    segments = [s for s in segments if s.isdigit()]
                    level = max(1, min(6, len(segments)))
                    heading_map[style_name] = level
                    logger.info(
                        f"Style Classification: '{style_name}' -> HEADING Level {level} (Numeric-leading pattern)"
                    )
                    continue

                # 3. Known built-in non-heading styles (exclude immediately)
                if style_name in builtin_non_heading_styles:
                    logger.debug(f"Style Classification: '{style_name}' -> BODY_TEXT "
                                f"(Known Word built-in style)")
                    continue
                
                # Check lowercase for localized styles
                is_builtin_non_heading = False
                for builtin in builtin_non_heading_styles:
                    if builtin.lower() == style_lower:
                        logger.debug(f"Style Classification: '{style_name}' -> BODY_TEXT "
                                    f"(Localized built-in style)")
                        is_builtin_non_heading = True
                        break
                
                if is_builtin_non_heading:
                    continue
                
                # 4. Strong exclusion patterns (never headings)
                strong_exclusions = ['caption', 'footnote', 'endnote', 'toc', 'bibliography', 
                                   'format', 'case', 'overview', 'break', 'divider', 'separator']
                if any(pattern in style_lower for pattern in strong_exclusions):
                    logger.debug(f"Style Classification: '{style_name}' -> BODY_TEXT "
                                f"(Strong exclusion pattern)")
                    continue
                
                # 5. Medium exclusion patterns
                medium_exclusions = ['quote', 'emphasis', 'strong', 'list', 'continue', 'text']
                if any(pattern in style_lower for pattern in medium_exclusions):
                    logger.debug(f"Style Classification: '{style_name}' -> BODY_TEXT "
                                f"(Medium exclusion pattern)")
                    continue
                
                # 6. International heading patterns (high confidence)
                international_patterns = [
                    'titre', 'titulo', 'uberschrift', 'haupt', 'sous-titre',
                    'chapitre', 'partie', 'sous-chapitre', 'sous chapitre', 'sous-section', 'sous section'
                ]
                if any(pattern in style_lower for pattern in international_patterns):
                    level = _infer_heading_level(style_name)
                    heading_map[style_name] = level
                    logger.info(f"Style Classification: '{style_name}' -> HEADING Level {level} "
                               f"(International heading pattern)")
                    continue
                
                # 7. Custom heading detection
                if 'heading' in style_lower:
                    # "Heading" overrides "style" when both present
                    if 'style' in style_lower:
                        level = _infer_heading_level(style_name)
                        heading_map[style_name] = level
                        logger.info(f"Style Classification: '{style_name}' -> HEADING Level {level} "
                                   f"(Custom heading style)")
                    # But exclude if it has text indicators
                    elif 'text' in style_lower:
                        logger.debug(f"Style Classification: '{style_name}' -> BODY_TEXT "
                                    f"(Contains 'heading' but also 'text')")
                    else:
                        level = _infer_heading_level(style_name)
                        heading_map[style_name] = level
                        logger.info(f"Style Classification: '{style_name}' -> HEADING Level {level} "
                                   f"(Custom style with 'heading' in name)")
                    continue
                
                # 8. Organizational patterns
                org_patterns = ['department head', 'section header', 'chapter title', 'part title']
                if any(pattern in style_lower for pattern in org_patterns):
                    level = _infer_heading_level(style_name)
                    heading_map[style_name] = level
                    logger.info(f"Style Classification: '{style_name}' -> HEADING Level {level} "
                               f"(Organizational heading pattern)")
                    continue
                
                # 9. Title/header patterns (if no exclusions)
                if any(pattern in style_lower for pattern in ['title', 'header']):
                    level = _infer_heading_level(style_name)
                    heading_map[style_name] = level
                    logger.info(f"Style Classification: '{style_name}' -> HEADING Level {level} "
                               f"(Title/header pattern)")
                    continue
                
                # 10. Check for explicit outline level in style
                try:
                    if hasattr(style, '_element'):
                        outline_vals = style._element.xpath("./w:pPr/w:outlineLvl/@w:val", 
                                                          namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                        if outline_vals:
                            outline_level = int(outline_vals[0])
                            if 0 <= outline_level <= 8:
                                level = outline_level + 1
                                heading_map[style_name] = level
                                logger.info(f"Style Classification: '{style_name}' -> HEADING Level {level} "
                                           f"(Explicit outline level)")
                                continue
                except Exception:
                    pass
                
                # 11. Default: unknown styles are body text (conservative)
                logger.debug(f"Style Classification: '{style_name}' -> BODY_TEXT "
                            f"(Unknown style, conservative default)")
                        
    except Exception as e:
        logger.warning(f"Error analyzing document styles: {e}")
    
    logger.info(f"Style Classification Summary: {len(heading_map)} heading styles identified "
                f"from {total_paragraph_styles} paragraph styles")
    
    # Log final heading map summary
    for style_name, level in sorted(heading_map.items()):
        logger.debug(f"Final heading map: '{style_name}' -> Level {level}")
    
    return heading_map


def _detect_builtin_heading_level(style_name: str) -> int | None:
    """Detect level for built-in Word heading styles (Heading 1, Heading 2, etc.)."""
    if not style_name:
        return None
    match = re.match(r"^Heading\s+(\d+)$", style_name, re.IGNORECASE)
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            pass
    return None


def _infer_heading_level(style_name: str) -> int:
    """Infer heading level for custom heading styles."""
    style_lower = style_name.lower()
    
    # Look for explicit numbers first
    number_match = re.search(r'(\d+)', style_name)
    if number_match:
        try:
            level = int(number_match.group(1))
            if 1 <= level <= 9:
                return level
        except ValueError:
            pass
    
    # Semantic level indicators
    if any(word in style_lower for word in ['title', 'main', 'principal', 'primary']):
        return 1
    elif any(word in style_lower for word in ['section', 'chapter', 'part']):
        return 2
    elif any(word in style_lower for word in ['subsection', 'subchapter']):
        return 3
    else:
        return 2  # Default level for custom headings