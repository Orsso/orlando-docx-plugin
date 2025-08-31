# API Reference

## Core Classes

### DocxDocumentHandler
Main document conversion service.

**Methods:**
- `can_handle(file_path: str) -> bool`: File format validation
- `convert_to_dita(file_path: str, **kwargs) -> DitaPackage`: Primary conversion
- `get_supported_extensions() -> List[str]`: Returns `['.docx']`

### DocxConversionLogic
Core conversion algorithms.

**Methods:**
- `process_document(doc: Document) -> DitaStructure`: Document processing
- `extract_images(doc: Document) -> List[ImageRef]`: Image extraction
- `generate_topics(structure: DitaStructure) -> List[Topic]`: Topic generation

### StructureAnalyzer
Document structure analysis.

**Methods:**
- `analyze_headings(doc: Document) -> HeadingStructure`: Heading detection
- `build_hierarchy(headings: List[Heading]) -> TopicTree`: Topic hierarchy
- `validate_structure(tree: TopicTree) -> ValidationResult`: Structure validation

### StyleAnalyzer
Style detection and mapping.

**Methods:**
- `detect_heading_styles(doc: Document) -> List[StyleDef]`: Style detection
- `map_styles_to_dita(styles: List[StyleDef]) -> StyleMapping`: DITA mapping
- `get_style_hierarchy() -> Dict[str, int]`: Style level mapping

## UI Components

### HeadingFilterPanel
Heading management interface.

**Properties:**
- `heading_styles: List[HeadingStyle]`: Available heading styles
- `visible_levels: Set[int]`: Currently visible heading levels

**Methods:**
- `update_filter(levels: Set[int])`: Update visibility filter
- `refresh_styles()`: Reload style information

### DocxStyleMarkerProvider
Style visualization markers.

**Methods:**
- `get_markers(document: Document) -> List[Marker]`: Generate style markers
- `update_markers(styles: List[StyleDef])`: Update marker display

## Data Structures

### DitaPackage
```python
@dataclass
class DitaPackage:
    topics: List[Topic]
    ditamap: DitaMap
    images: List[ImageRef]
    metadata: DocumentMetadata
```

### Topic
```python
@dataclass
class Topic:
    id: str
    title: str
    content: str
    level: int
    parent_id: Optional[str]
```

### StyleDef
```python
@dataclass
class StyleDef:
    name: str
    level: int
    font_size: int
    is_bold: bool
    color: Optional[str]
```