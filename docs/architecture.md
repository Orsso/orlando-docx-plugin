# Plugin Architecture

## Core Components

### Document Handler
- **DocxDocumentHandler**: Primary conversion service implementing `DocumentHandlerBase`
- **Key Methods**: `can_handle()`, `convert_to_dita()`, `get_supported_extensions()`
- **Integration**: Registers with Orlando Toolkit's document processing pipeline

### Service Layer
```
DocxConversionLogic
├── StructureAnalyzer    # Document structure analysis
├── DitaBuilder         # DITA XML generation
├── DocxParser          # Word document parsing
└── StyleAnalyzer       # Style detection and mapping
```

### UI Extensions
- **HeadingFilterPanel**: Right panel for heading management
- **DocxStyleMarkerProvider**: Style visualization in scrollbar
- **Integration**: Orlando Toolkit UI registry system

## Plugin Lifecycle

1. **Discovery**: `plugin.json` manifest detection
2. **Loading**: Plugin class instantiation and validation
3. **Activation**: Service registration with Orlando Toolkit
4. **Registration**: DocumentHandler and UI components registered
5. **Ready**: Available for document conversion
6. **Deactivation**: Clean shutdown and resource cleanup

## Data Flow

```
DOCX Input → DocxParser → StructureAnalyzer → StyleAnalyzer → DitaBuilder → DITA Output
                                     ↓
                            HeadingFilterPanel (UI feedback)
```

## Extension Points

- **Document Handlers**: Implement `DocumentHandlerBase` for new formats
- **Style Analyzers**: Custom style detection logic
- **UI Panels**: Additional management interfaces
- **Markers**: Custom visualization providers