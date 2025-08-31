# Usage Guide

## Basic Conversion Workflow

1. **Import DOCX**: Use "Import from DOCX" in Orlando Toolkit
2. **Analyze Structure**: Review heading detection results
3. **Configure Styles**: Adjust heading level mappings if needed
4. **Convert**: Generate DITA package
5. **Export**: Save converted content

## Heading Filter Panel

### Features
- **Level Filtering**: Show/hide heading levels 1-6
- **Style Statistics**: Count of each heading style found
- **Color Coding**: Visual identification of style types
- **Real-time Updates**: Dynamic filtering as you work

### Usage
- **Toggle Levels**: Click level buttons (H1-H6) to show/hide
- **View Stats**: See heading counts per level
- **Filter Content**: Use checkboxes to control visibility

## Style Management

### Automatic Detection
- **Font Size**: Larger fonts detected as higher-level headings
- **Style Names**: Word style names mapped to DITA levels
- **Formatting**: Bold, color, and other attributes considered

### Manual Override
- **Custom Mapping**: Configure style-to-level assignments
- **Style Rules**: Define detection patterns
- **Hierarchy Validation**: Ensure proper heading sequence

## Conversion Options

### Document Metadata
- **Title**: Document title for DITAMAP
- **ID**: Unique identifier generation
- **Author**: Author information preservation
- **Subject**: Topic categorization

### Processing Settings
- **Image Handling**: Extract and optimize images
- **Table Conversion**: Preserve table structure in DITA
- **Link Processing**: Maintain cross-references
- **List Formatting**: Convert numbered and bulleted lists

## Troubleshooting

### Common Issues
- **Missing Headings**: Check style consistency in source document
- **Image Problems**: Verify image format compatibility
- **Style Conflicts**: Review heading hierarchy in filter panel
- **Conversion Errors**: Check DOCX file integrity

### Solutions
- **Remap Styles**: Use heading filter to adjust level assignments
- **Check Source**: Validate DOCX file in Microsoft Word
- **Review Logs**: Check Orlando Toolkit console for error details
- **Clean Document**: Remove problematic formatting from source