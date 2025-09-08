# API Reference

## Core Classes

### DocxDocumentHandler
Main document conversion service.

**Methods:**
- `can_handle(file_path: Path) -> bool`: File format validation
- `handle_document(file_path: str) -> DitaContext`: App-facing entry point
- `convert_to_dita(file_path: Path, metadata: Dict[str, Any], progress_callback: Optional[Callable[[str], None]]) -> DitaContext`: Primary conversion
- `get_supported_extensions() -> List[str]`: Returns `['.docx']`

### DocxConversionLogic
Core conversion algorithms (`convert_docx_to_dita_internal`).

**Responsibilities:**
- Load DOCX, build style heading map, extract images
- Build hierarchical structure and generate DITA topics and map
- Apply metadata (title, revision_date, codes) and tocIndex

### StructureAnalyzer
Document structure analysis.

**Key functions:**
- `build_document_structure(doc, style_heading_map, all_images_map_rid) -> List[HeadingNode]`
- `determine_node_roles(nodes) -> None`
- `generate_dita_from_structure(nodes, context, metadata, all_images_map_rid, parent_element, heading_counters, color_rules) -> None`

### StyleAnalyzer
Style detection and mapping.

**Key function:**
- `build_style_heading_map(doc: Document, plugin_config: Optional[Dict[str, Any]] = None) -> Dict[str, int]`

## UI Components

### HeadingFilterPanel
Heading management interface.

**Public API:**
- `set_data(headings_count, occurrences_by_style, style_levels, current_exclusions)`
- `update_status(text)`
- `clear_selection()`
- `toggle_style_visibility(style, visible)`
- `get_visible_styles()` / `get_selected_style()`

### DocxStyleMarkerProvider
Style visualization markers.

**Description:**
- Provides style-based markers for the scrollbar visualization

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