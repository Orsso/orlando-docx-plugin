"""DOCX Document Handler for Orlando Toolkit Plugin System.

This module provides the main DocumentHandler implementation for converting
Microsoft Word documents (.docx) to DITA format using the plugin architecture.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List

from orlando_toolkit.core.models import DitaContext
from orlando_toolkit.core.plugins.interfaces import DocumentHandlerBase

logger = logging.getLogger(__name__)


class DocxDocumentHandler(DocumentHandlerBase):
    """Document handler for Microsoft Word (.docx) files.
    
    Implements the DocumentHandler protocol to provide DOCX to DITA
    conversion capabilities as a plugin service.
    """
    
    def can_handle(self, file_path: Path) -> bool:
        """Return True if this handler can process the file.
        
        Args:
            file_path: Path to the file to be checked
            
        Returns:
            True if file has .docx extension, False otherwise
        """
        return file_path.suffix.lower() == '.docx'
    
    def convert_to_dita(self, file_path: Path, metadata: Dict[str, Any]) -> DitaContext:
        """Convert DOCX file to DitaContext.
        
        Args:
            file_path: Path to the source DOCX document
            metadata: Conversion metadata and configuration options
            
        Returns:
            DitaContext containing the complete DITA archive data
            
        Raises:
            Exception: If conversion fails
        """
        self.validate_file_exists(file_path)
        
        # Import the actual conversion function
        from .docx_conversion_logic import convert_docx_to_dita_internal
        
        logger.info("Starting DOCX->DITA conversion via plugin: %s", file_path)
        
        # Get plugin configuration if available
        plugin_config = getattr(self, '_plugin_config', {})
        return convert_docx_to_dita_internal(str(file_path), metadata, plugin_config)
    
    def get_supported_extensions(self) -> List[str]:
        """Return list of supported file extensions.
        
        Returns:
            List containing '.docx' extension
        """
        return ['.docx']
    
    def get_conversion_metadata_schema(self) -> Dict[str, Any]:
        """Return JSON schema for DOCX conversion metadata fields.
        
        Returns:
            JSON schema dictionary for DOCX-specific metadata
        """
        return {
            "$schema": "http://json-schema.org/draft-07/schema#",
            "type": "object",
            "properties": {
                "manual_title": {
                    "type": "string",
                    "description": "Title for the generated DITA map"
                },
                "manual_code": {
                    "type": "string",
                    "description": "Manual code for Orlando metadata"
                },
                "manual_reference": {
                    "type": "string",
                    "description": "Manual reference for Orlando metadata"
                },
                "revision_date": {
                    "type": "string",
                    "format": "date",
                    "description": "Revision date (YYYY-MM-DD format)"
                },
                "revision_number": {
                    "type": "string",
                    "description": "Revision number for the document"
                },
                "style_heading_map": {
                    "type": "object",
                    "description": "Custom mapping of style names to heading levels",
                    "additionalProperties": {
                        "type": "integer",
                        "minimum": 1,
                        "maximum": 6
                    }
                },
                "enable_structural_style_inference": {
                    "type": "boolean",
                    "description": "Enable structural analysis for heading detection"
                },
                "min_following_paragraphs": {
                    "type": "integer",
                    "minimum": 1,
                    "description": "Minimum paragraphs required for structural inference"
                },
                "generic_heading_match": {
                    "type": "boolean",
                    "description": "Enable generic heading name pattern matching"
                }
            }
        }