# DOCX Converter Plugin

**⚠️ STANDALONE REPOSITORY**: This is a standalone plugin repository designed for installation via Orlando Toolkit's Plugin Management system.

Convert Microsoft Word documents to DITA format with advanced style analysis and UI extensions.

## Overview

The DOCX Converter Plugin extends Orlando Toolkit with Microsoft Word document import capabilities. It provides complete DOCX to DITA conversion including structure analysis, image extraction, and style-based markup generation.

**Key Capabilities:**
- **Document Conversion**: Full DOCX to DITA transformation with topic structure generation
- **Style Analysis**: Intelligent heading detection and style mapping
- **Image Processing**: Automatic image extraction and format optimization
- **UI Extensions**: Heading filter panel and style marker visualization
- **Plugin Integration**: Seamless integration with Orlando Toolkit's plugin architecture

## Installation

### Automatic Installation (Recommended)

1. **Install Orlando Toolkit** and launch the application
2. **Open Plugin Management** from the splash screen
3. **Enter GitHub Repository URL**: `https://github.com/organization/orlando-docx-plugin`
4. **Click "Import Plugin"** to download and install automatically
5. **Activate Plugin** to enable DOCX conversion functionality

The plugin system will automatically:
- Download the plugin from GitHub
- Install Python dependencies from `requirements.txt`
- Install to the correct plugins directory
- Register the plugin with Orlando Toolkit

### Manual Installation (Development)

For development purposes only:
1. Clone this repository to your local development environment
2. Install plugin dependencies: `pip install -r requirements.txt`
3. Use Orlando Toolkit's "Install from Directory" option in Plugin Management

### Plugin Dependencies

The plugin requires these Python packages (automatically installed during import):
- **python-docx**: Core DOCX file parsing and structure analysis
- **lxml**: XML processing for DITA generation
- **Pillow**: Image processing and format conversion
- **requests**: HTTP requests for potential future online resources

### Dependency Details

- **python-docx**: Core DOCX file parsing and structure analysis
- **lxml**: XML processing for DITA generation
- **Pillow**: Image processing and format conversion
- **requests**: HTTP requests for potential future online resources

## Features

### Document Conversion

**DOCX Processing:**
- Paragraph and heading structure extraction
- Table conversion with DITA table markup
- List processing (numbered and bulleted)
- Cross-reference and hyperlink preservation
- Style-based content categorization

**DITA Generation:**
- Topic-based structure with proper DITA hierarchy
- DITAMAP creation with navigation structure
- Image references and media handling
- Metadata extraction and mapping

**Image Handling:**
- Automatic image extraction from DOCX
- Format optimization (JPEG, PNG, GIF support)
- Size and quality optimization
- Proper DITA image reference generation

### Style Analysis

The plugin provides sophisticated style analysis capabilities:

**Heading Detection:**
- Automatic heading level recognition
- Style-based heading classification
- Custom style mapping configuration
- Heading hierarchy validation

**Style Mapping:**
- Built-in style templates for common Word styles
- Custom style rule configuration
- Style inheritance and cascading
- Formatting preservation

### UI Extensions

**Heading Filter Panel:**
- Visual heading style management
- Level-based style grouping
- Real-time style visibility toggling
- Style occurrence statistics
- Color-coded style indicators

**Style Markers:**
- Scrollbar marker integration
- Visual style identification in document tree
- Priority-based marker layering
- Customizable marker colors

## Usage

### Basic Conversion

1. **Launch Orlando Toolkit** with DOCX plugin installed
2. **Import DOCX File**: Use the "Import from DOCX" button on the splash screen
3. **Configure Conversion**: Set document metadata and conversion options
4. **Review Structure**: Use the heading filter to analyze document structure
5. **Convert**: Process the document to generate DITA archive
6. **Export**: Save the resulting DITA package

### Advanced Configuration

**Conversion Settings:**
- Document title and metadata
- Topic naming conventions
- Image processing options
- Style mapping rules

**Style Management:**
- Filter heading styles by level
- Toggle style visibility in tree view
- Configure style-to-DITA mapping
- Manage style color coding

### Workflow Example

```
1. Select DOCX file → 2. Analyze structure → 3. Configure styles → 4. Convert to DITA
   [user-manual.docx]   [15 headings found]    [H1-H6 mapping]     [DITA archive]
```

## Plugin Architecture

### Core Components

**DocumentHandler Implementation:**
- `DocxDocumentHandler`: Main conversion service
- Implements `DocumentHandlerBase` interface
- Provides `can_handle()`, `convert_to_dita()`, `get_supported_extensions()` methods

**UI Extension Points:**
- `HeadingFilterPanel`: Right panel extension for heading management
- `DocxStyleMarkerProvider`: Marker provider for style visualization
- Integration with Orlando Toolkit's UI registry system

**Service Architecture:**
- `DocxConversionLogic`: Core conversion algorithms
- `StructureAnalyzer`: Document structure analysis
- `DitaBuilder`: DITA XML generation
- `DocxParser`: Word document parsing utilities
- `StyleAnalyzer`: Style detection and mapping

### Plugin Lifecycle

1. **Discovery**: Plugin detected via `plugin.json` manifest
2. **Loading**: Plugin class instantiated and validated
3. **Activation**: Services registered with Orlando Toolkit
4. **Service Registration**: DocumentHandler and UI components registered
5. **Ready**: Plugin available for document conversion
6. **Deactivation**: Clean shutdown and resource cleanup

## Configuration

### Plugin Manifest (plugin.json)

```json
{
  "name": "docx-converter",
  "version": "1.0.0",
  "display_name": "DOCX Converter",
  "description": "Convert Microsoft Word documents to DITA",
  "orlando_version": ">=2.0.0",
  "plugin_api_version": "1.0",
  "category": "pipeline",
  "entry_point": "src.plugin.DocxConverterPlugin"
}
```

### Style Configuration

Create custom style mappings by configuring the plugin's style detection rules:

- **Heading Styles**: Map Word heading styles to DITA topic types
- **Content Styles**: Configure paragraph and inline style handling
- **Table Styles**: Define table formatting preservation rules
- **List Styles**: Set numbered and bulleted list processing

## Troubleshooting

### Common Issues

**Plugin Not Loading:**
- Verify all dependencies are installed
- Check plugin.json syntax and validity
- Review Orlando Toolkit logs for error messages
- Ensure plugin directory structure is correct

**Conversion Errors:**
- Validate DOCX file integrity (open in Microsoft Word)
- Check for corrupted images or embedded objects
- Verify sufficient disk space for image extraction
- Review conversion metadata and settings

**Style Detection Issues:**
- Examine heading style consistency in source document
- Configure custom style mapping rules
- Use heading filter panel to analyze detected styles
- Verify style inheritance and cascading

**Image Processing Problems:**
- Check image format compatibility (JPEG, PNG, GIF)
- Verify image file integrity and accessibility
- Ensure Pillow library is properly installed
- Review image size and optimization settings

### Performance Optimization

**Large Document Handling:**
- Monitor memory usage during conversion
- Consider document section-based processing
- Optimize image compression settings
- Use incremental conversion for very large files

**Style Processing:**
- Cache style analysis results
- Optimize style detection algorithms
- Use efficient style mapping structures
- Minimize redundant style calculations

## Development

### Plugin Structure

```
orlando-docx-plugin/
├── plugin.json           # Plugin manifest
├── requirements.txt      # Dependencies
├── src/
│   ├── plugin.py        # Main plugin class
│   ├── services/        # Conversion services
│   │   ├── docx_handler.py
│   │   ├── docx_conversion_logic.py
│   │   ├── structure_analyzer.py
│   │   └── formatting_helpers.py
│   ├── ui/              # UI extensions
│   │   └── heading_filter.py
│   └── utils/           # Utilities
│       ├── docx_parser.py
│       ├── dita_builder.py
│       └── style_analyzer.py
└── tests/               # Plugin tests
```

### Testing

Run plugin tests to verify functionality:

```bash
# Run plugin unit tests
python -m pytest orlando-docx-plugin/tests/

# Run integration tests with Orlando Toolkit
python -m pytest tests/integration/test_docx_plugin.py
```

### Extending the Plugin

**Adding New Features:**
1. Implement new service classes in `src/services/`
2. Register services in plugin's `on_activate()` method
3. Add UI components in `src/ui/` if needed
4. Update plugin manifest with new capabilities
5. Add comprehensive tests for new functionality

**Custom Style Handlers:**
1. Extend `StyleAnalyzer` class with new detection logic
2. Add configuration options in plugin settings
3. Implement UI controls for new style options
4. Test with variety of document styles and formats

## Version Compatibility

- **Orlando Toolkit**: 2.0.0+
- **Plugin API**: 1.0
- **Python**: 3.8+
- **Microsoft Word**: 2007+ (DOCX format)

## Related Documentation

- **Main Application**: [Orlando Toolkit](https://github.com/organization/orlando-toolkit) - Application overview and installation
- **Plugin Development**: [Plugin Development Guide](https://github.com/organization/orlando-toolkit/blob/main/docs/PLUGIN_DEVELOPMENT_GUIDE.md) - Complete developer documentation
- **Architecture Overview**: [Architecture Documentation](https://github.com/organization/orlando-toolkit/blob/main/docs/architecture_overview.md) - System architecture details
- **Configuration**: [Configuration Guide](https://github.com/organization/orlando-toolkit/blob/main/orlando_toolkit/config/README.md) - Application configuration

## Support

- **Bug Reports**: Submit issues via this repository or the main [Orlando Toolkit repository](https://github.com/organization/orlando-toolkit/issues)
- **Feature Requests**: Use GitHub Discussions for enhancement ideas
- **Documentation**: Refer to [Plugin Development Guide](https://github.com/organization/orlando-toolkit/blob/main/docs/PLUGIN_DEVELOPMENT_GUIDE.md)
- **Community**: Join discussions in the Orlando Toolkit community

## License

This plugin is distributed under the same MIT license as Orlando Toolkit. See the main project LICENSE file for details.