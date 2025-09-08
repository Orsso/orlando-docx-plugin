# Development Guide

## Project Structure

```
docx_plugin/
├── plugin.py              # Main plugin entry point
├── services/              # Core conversion services
│   ├── docx_handler.py    # Document handler implementation
│   ├── docx_conversion_logic.py  # Conversion algorithms
│   ├── structure_analyzer.py     # Document structure analysis
│   └── formatting_helpers.py     # Style processing utilities
├── ui/                    # UI components
│   └── heading_filter.py  # Heading filter panel
└── utils/                 # Utility modules
    ├── docx_parser.py     # Word document parsing
    ├── dita_builder.py    # DITA XML generation
    ├── style_analyzer.py  # Style detection
    └── color_utils.py     # Color processing
```

## Development Setup

1. **Clone repository**
2. **Install dependencies**: `pip install -r requirements.txt`
3. **Install in Orlando Toolkit**: Use "Install from Directory" option
4. **Test with sample documents**

## Key Implementation Points

### Document Handler Registration
```python
def on_activate(self):
    self.framework.document_service.register_handler(DocxDocumentHandler())
    # UI panel and markers are registered by the plugin via UIExtension hooks
```

### Style Detection Logic
Located in `docx_plugin/utils/style_analyzer.py`:
- Built-in "Heading N" detection (1-9)
- Numeric-leading style patterns (e.g., "1.2 Section")
- International heading terms (e.g., "Titre", "Kapitel")
- Explicit outline level if defined on the style

### DITA Generation
Located in `docx_plugin/utils/dita_builder.py` and service layer:
- Topic structure creation (concept topics)
- Ditamap topicref/topichead generation with tocIndex
- Image reference handling and media extraction
- Metadata preservation (manual_code, manual_reference, revision)

## Testing

**Unit Tests**: Test individual components in isolation
**Integration Tests**: Test with Orlando Toolkit framework
**Document Tests**: Test with various DOCX samples

## Extension Points

### Custom Style Handlers
1. Extend `StyleAnalyzer` class
2. Implement custom detection logic
3. Add configuration options
4. Register with plugin system

### New UI Components
1. Create component in `src/ui/`
2. Implement Orlando Toolkit UI interface
3. Register in plugin's `on_activate()`
4. Add necessary CSS/styling