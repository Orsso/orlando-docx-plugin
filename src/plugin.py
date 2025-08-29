"""Main DOCX Plugin Implementation.

This module provides the primary plugin class that implements the BasePlugin
interface and registers DOCX document handling services with the Orlando
Toolkit plugin system.
"""

from __future__ import annotations

import logging
from typing import Any, Dict, Optional

from orlando_toolkit.core.plugins.base import BasePlugin, AppContext
from orlando_toolkit.core.plugins.interfaces import UIExtension
from orlando_toolkit.core.plugins.marker_providers import MarkerProvider

logger = logging.getLogger(__name__)


class DocxConverterPlugin(BasePlugin, UIExtension):
    """Main DOCX converter plugin for Orlando Toolkit.
    
    This plugin provides complete DOCX to DITA conversion capabilities
    including document handling, UI extensions, and style marker support.
    """
    
    def __init__(self, plugin_id: str, metadata: 'PluginMetadata', plugin_dir: str) -> None:
        """Initialize the DOCX converter plugin.
        
        Args:
            plugin_id: Unique identifier for this plugin instance
            metadata: Validated plugin metadata
            plugin_dir: Path to plugin directory
        """
        super().__init__(plugin_id, metadata, plugin_dir)
        self._document_handler: Optional[Any] = None
        self._marker_provider: Optional[Any] = None
        
        # Load plugin-specific configuration
        try:
            self.load_config()
            self.log_debug("Plugin configuration loaded successfully")
        except Exception as e:
            self.log_warning(f"Failed to load plugin configuration: {e}")
            # Continue with default configuration
        
    def get_name(self) -> str:
        """Get human-readable plugin name."""
        return "DOCX Converter"
    
    def get_description(self) -> str:
        """Get plugin description."""
        return "Convert Microsoft Word documents to DITA format with advanced style analysis"
    
    def _get_default_config(self) -> Dict[str, Any]:
        """Get default configuration for the DOCX plugin.
        
        Returns:
            Default configuration dictionary
        """
        return {
            'docx_conversion': {
                'enable_structural_style_inference': True,
                'min_following_paragraphs': 3,
                'generic_heading_match': True
            },
            'heading_filter': {
                'max_active_styles': 5,
                'default_exclusions': [],
                'panel': {
                    'show_by_default': True,
                    'position': 'right'
                }
            }
        }
    
    def on_activate(self) -> None:
        """Called when plugin services should be registered."""
        super().on_activate()
        
        try:
            # Register DocumentHandler service
            from .services.docx_handler import DocxDocumentHandler
            self._document_handler = DocxDocumentHandler()
            
            # Pass plugin configuration to handler
            self._document_handler._plugin_config = self.config
            
            if self.app_context and hasattr(self.app_context, 'service_registry'):
                self.app_context.service_registry.register_service(
                    'DocumentHandler', self._document_handler, self.plugin_id
                )
                self.log_info("Registered DOCX DocumentHandler service")
            
            # Register UI extensions
            if self.app_context and hasattr(self.app_context, 'ui_registry'):
                try:
                    # Register heading filter panel factory
                    self.app_context.ui_registry.register_panel_factory(
                        'heading_filter', self.create_heading_filter_panel, self.plugin_id
                    )
                    self.log_info("Registered heading filter panel factory")
                except AttributeError:
                    self.log_warning("UI registry does not support panel factory registration")
                
                try:
                    # Register marker provider
                    self._marker_provider = DocxStyleMarkerProvider()
                    self.app_context.ui_registry.register_marker_provider(
                        self._marker_provider, self.plugin_id
                    )
                    self.log_info("Registered DOCX style marker provider")
                except AttributeError:
                    self.log_warning("UI registry does not support marker provider registration")
                
        except Exception as e:
            self.log_error(f"Failed to activate plugin: {e}")
            raise
    
    def on_deactivate(self) -> None:
        """Called when plugin should cleanup resources and unregister services."""
        super().on_deactivate()
        
        try:
            # Unregister services
            if self.app_context and hasattr(self.app_context, 'service_registry'):
                if self._document_handler:
                    self.app_context.service_registry.unregister_service(
                        'DocumentHandler', self.plugin_id
                    )
                    self._document_handler = None
                    self.log_info("Unregistered DOCX DocumentHandler service")
            
            # Unregister UI extensions
            if self.app_context and hasattr(self.app_context, 'ui_registry'):
                try:
                    self.app_context.ui_registry.unregister_panel_factory(
                        'heading_filter', self.plugin_id
                    )
                    self.log_info("Unregistered heading filter panel factory")
                except AttributeError:
                    pass
                
                try:
                    if self._marker_provider:
                        self.app_context.ui_registry.unregister_marker_provider(
                            self.plugin_id
                        )
                        self._marker_provider = None
                        self.log_info("Unregistered DOCX style marker provider")
                except AttributeError:
                    pass
                    
        except Exception as e:
            self.log_warning(f"Error during deactivation: {e}")
    
    # UI Extension interface implementation
    
    def get_extension_info(self) -> Dict[str, Any]:
        """Get information about this UI extension."""
        return {
            'supported_components': ['panel_factory', 'marker_provider'],
            'display_name': 'DOCX UI Extensions',
            'description': 'DOCX-specific heading analysis and filtering UI components'
        }
    
    def register_ui_components(self, ui_registry: Any) -> None:
        """Register UI components with the UI registry."""
        try:
            ui_registry.register_panel_factory(
                'heading_filter', self.create_heading_filter_panel, self.plugin_id
            )
            
            if not self._marker_provider:
                self._marker_provider = DocxStyleMarkerProvider()
            ui_registry.register_marker_provider(self._marker_provider, self.plugin_id)
            
        except Exception as e:
            self.log_error(f"Failed to register UI components: {e}")
            raise
    
    def unregister_ui_components(self, ui_registry: Any) -> None:
        """Unregister UI components from the UI registry."""
        try:
            ui_registry.unregister_panel_factory('heading_filter', self.plugin_id)
            ui_registry.unregister_marker_provider(self.plugin_id)
            self._marker_provider = None
            
        except Exception as e:
            self.log_warning(f"Error unregistering UI components: {e}")
    
    def get_panel_factories(self) -> Dict[str, Any]:
        """Get panel factories provided by this extension."""
        return {
            'heading_filter': self.create_heading_filter_panel
        }
    
    def get_marker_providers(self) -> Dict[str, Any]:
        """Get marker providers provided by this extension."""
        if not self._marker_provider:
            self._marker_provider = DocxStyleMarkerProvider()
        return {
            'docx_styles': self._marker_provider
        }
    
    def create_heading_filter_panel(self, parent, **kwargs):
        """Factory method to create heading filter panel instances.
        
        Args:
            parent: Parent widget for the panel
            **kwargs: Additional arguments passed to panel constructor
            
        Returns:
            HeadingFilterPanel instance
        """
        from .ui.heading_filter import HeadingFilterPanel
        return HeadingFilterPanel(parent, **kwargs)


class DocxStyleMarkerProvider(MarkerProvider):
    """Marker provider for DOCX style-based markers in the UI."""
    
    def get_provider_name(self) -> str:
        """Get the name of this marker provider."""
        return "DOCX Style Markers"
    
    def get_supported_marker_types(self) -> list[str]:
        """Get list of marker types this provider supports."""
        return ["style_marker"]
    
    def create_marker(self, marker_type: str, config: Dict[str, Any]) -> Optional[Any]:
        """Create a marker of the specified type.
        
        Args:
            marker_type: Type of marker to create
            config: Configuration for the marker
            
        Returns:
            Marker instance or None if type not supported
        """
        if marker_type == "style_marker":
            return DocxStyleMarker(config)
        return None
    
    def get_marker_config_schema(self, marker_type: str) -> Optional[Dict[str, Any]]:
        """Get configuration schema for a marker type.
        
        Args:
            marker_type: Type of marker
            
        Returns:
            JSON schema for marker configuration or None
        """
        if marker_type == "style_marker":
            return {
                "type": "object",
                "properties": {
                    "style_name": {"type": "string"},
                    "color": {"type": "string"},
                    "visible": {"type": "boolean"},
                    "level": {"type": "integer", "minimum": 1, "maximum": 9}
                },
                "required": ["style_name"]
            }
        return None


class DocxStyleMarker:
    """Marker implementation for DOCX style visualization."""
    
    def __init__(self, config: Dict[str, Any]):
        """Initialize style marker with configuration.
        
        Args:
            config: Marker configuration
        """
        self.style_name = config.get("style_name", "")
        self.color = config.get("color", "#000000")
        self.visible = config.get("visible", True)
        self.level = config.get("level", 1)
    
    def get_marker_type(self) -> str:
        """Get the type of this marker."""
        return "style_marker"
    
    def is_visible(self) -> bool:
        """Check if marker should be visible."""
        return self.visible
    
    def get_color(self) -> str:
        """Get marker color."""
        return self.color
    
    def get_style_name(self) -> str:
        """Get associated style name."""
        return self.style_name
    
    def get_level(self) -> int:
        """Get heading level."""
        return self.level