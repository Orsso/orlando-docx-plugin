# DOCX Converter Plugin

**⚠️ DEPENDANCE REPOSITORY**: Plugin for Orlando Toolkit.

Convert Microsoft Word documents to DITA format with advanced style analysis and UI extensions.

## Features

- **Document Conversion**: DOCX to DITA transformation with topic structure generation
- **Style Analysis**: Intelligent heading detection and style mapping  
- **Image Processing**: Automatic extraction and format optimization
- **UI Extensions**: Heading filter panel and style marker visualization

## Installation

### Automatic Installation (Recommended)

1. Install Orlando Toolkit and launch application
2. Open Plugin Management from splash screen
3. Enter repository URL: `https://github.com/orsso/orlando-docx-plugin`
4. Click "Import Plugin" to download and install
5. Activate plugin to enable DOCX conversion

### Manual Installation (Development)

1. Clone repository to local environment
2. Install dependencies: `pip install -r requirements.txt` 
3. Use "Install from Directory" in Orlando Toolkit Plugin Management

## Quick Start

1. Launch Orlando Toolkit with plugin installed
2. Use "Import from DOCX" button on splash screen
3. Configure conversion settings and style mappings
4. Review structure using heading filter panel
5. Convert to generate DITA archive

## Dependencies

- **python-docx**: DOCX parsing and structure analysis
- **lxml**: XML processing for DITA generation
- **Pillow**: Image processing and format conversion
- **requests**: HTTP requests for external resources

## Documentation

- **[Architecture](docs/architecture.md)**: Plugin structure and component overview
- **[API Reference](docs/api.md)**: Classes, methods, and data structures
- **[Development Guide](docs/development.md)**: Setup, structure, and extension points
- **[Usage Guide](docs/usage.md)**: Conversion workflow and troubleshooting

## Plugin Structure

```
src/
├── plugin.py              # Main plugin entry point
├── services/              # Core conversion services  
│   ├── docx_handler.py    # Document handler implementation
│   ├── docx_conversion_logic.py  # Conversion algorithms
│   └── structure_analyzer.py     # Document structure analysis
├── ui/                    # UI components
│   └── heading_filter.py  # Heading filter panel
└── utils/                 # Utility modules
    ├── dita_builder.py    # DITA XML generation
    └── style_analyzer.py  # Style detection
```

## Compatibility

- **Orlando Toolkit**: 1.2.0+
- **Plugin API**: 1.0  
- **Python**: 3.8+
- **Microsoft Word**: 2007+ (DOCX format)

## Support

- **Issues**: Submit via repository issues or main [Orlando Toolkit repository](https://github.com/orsso/orlando-toolkit/issues)
- **Documentation**: [Plugin Development Guide](https://github.com/orsso/orlando-toolkit/blob/main/docs/PLUGIN_DEVELOPMENT_GUIDE.md)

## License

MIT License - same as Orlando Toolkit main project.