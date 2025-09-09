# Orlando DOCX Plugin

A plugin for [Orlando Toolkit](https://github.com/orsso/orlando-toolkit) that converts Microsoft Word (DOCX) files to DITA topics with style-aware structure and in-app heading analysis.

Note: This plugin is in early development. Please review and validate all outputs.
## Features

- **Convert DOCX to DITA**: Generates DITA topics from Word documents
- **Style-aware headings**: Built-in and inferred heading detection and mapping
- **Image extraction**: Saves embedded images and relinks them in topics
- **In-app tools**: Heading Filter panel and style markers in the Structure tab

## Installation

**Via Orlando Toolkit Plugin Manager:**

1. Open Plugin Management from Orlando Toolkit splash screen
2. Enter this repository URL: `https://github.com/orsso/orlando-docx-plugin`
3. Click "Import Plugin" to install
4. Activate the plugin

## Usage

**Import from DOCX**: Choose a `.docx` file → the plugin converts it to one or more DITA topics based on detected headings.

```
MyManual.docx
└── ➜ topics/
    ├── topic_abc123.dita
    ├── topic_def456.dita
    └── media/ (extracted images referenced by topics)
```

**Heading analysis**: Use the Heading Filter panel in the Structure tab to inspect detected headings and fine-tune mapping.

## Output

- **DITA Topics**: `.dita` files with sections derived from document structure
- **Media Files**: Extracted images saved in a `media/` folder and referenced from topics
- **Metadata**: Optional fields (title, manual code/reference, revision) added to topic metadata when provided

## Requirements

- **Orlando Toolkit**: 1.2.0+
- **Python**: 3.8+

## Documentation

- **[Architecture](docs/architecture.md)**
- **[API Reference](docs/api.md)**
- **[Development Guide](docs/development.md)**
- **[Usage Guide](docs/usage.md)**

## License

MIT License.
