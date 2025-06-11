# PowerPoint to Slidev Converter

A Python script that converts PowerPoint (.pptx) files exported from Google Slides into beautiful [Slidev](https://sli.dev) presentations.

## Features

- ✅ **Converts PowerPoint slides to Slidev markdown**
- ✅ **Extracts and preserves images** from slides with proper sizing
- ✅ **Maintains hierarchical bullet point structure**
- ✅ **Handles multiword titles correctly** (fixes newline issues)
- ✅ **Escapes special characters** for Vue compatibility (no more parsing errors)
- ✅ **Professional Seriph theme** with custom styling
- ✅ **Clean bullet formatting** (no double bullets)
- ✅ **Proper title slide placement** (no empty first slide)
- ✅ **Support for tables** (manual restoration with custom CSS styling)
- ✅ **YouTube video integration** (manual replacement with `<Youtube>` component)
- ✅ **Reusable bio slides** (importable across presentations)

## Installation

### Prerequisites

1. **Python 3.7+** 
2. **python-pptx library**

```bash
pip install python-pptx
```

### Install Slidev (if not already installed)

```bash
npm install -g @slidev/cli
npm install -g @slidev/theme-seriph
```

## Usage

### Basic Conversion

```bash
python slidev_converter.py --pptx "your-presentation.pptx" --output ./output
```

### Command Line Options

- `--pptx`: Path to the PowerPoint file to convert (required)
- `--batch`: Directory containing multiple PowerPoint files for batch conversion
- `--output`: Output directory (default: `./slidev-presentations`)

### Examples

**Convert a single presentation:**
```bash
python slidev_converter.py --pptx "My Presentation.pptx" --output ./slides
```

**Batch convert multiple presentations:**
```bash
python slidev_converter.py --batch ./presentations --output ./converted-slides
```

**Convert with default settings:**
```bash
python slidev_converter.py --pptx "presentation.pptx"
```

## Workflow

### 1. Prepare Your Google Slides

1. Create your presentation in Google Slides
2. Export as PowerPoint: **File → Download → Microsoft PowerPoint (.pptx)**

### 2. Run the Converter

```bash
python slidev_converter.py --pptx "exported-presentation.pptx" --output ./my-slidev
```

The converter will:
- Extract all slides and convert them to markdown
- Save images from slides as PNG files
- Generate a Slidev-compatible markdown file
- Apply the Seriph theme with professional styling

### 3. View Your Presentation

```bash
cd ./my-slidev
slidev your-presentation-name.md
```

Your presentation will open at `http://localhost:3030`

## Reusable Bio Slides

The project includes `ken-kousen-bio.md` as an example of creating reusable slides that can be imported into any presentation:

```markdown
---
src: ./ken-kousen-bio.md
---
```

This approach allows you to:
- Maintain consistent contact information across presentations
- Update book links and social media in one place
- Create a library of reusable slide components

## Output Structure

After conversion, you'll get:

```
output-directory/
├── presentation-name.md          # Main Slidev file
├── slide_2_image_1.png          # Extracted images
├── slide_5_image_1.png
└── ...
```

## Manual Fixes You May Need

### Tables
Tables from Google Slides may not convert properly. You'll need to recreate them in markdown:

```markdown
| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| Data 1   | Data 2   | Data 3   |
```

### YouTube Videos
Embedded YouTube videos become static images. Replace with Slidev's YouTube component:

```markdown
<Youtube id="VIDEO_ID" width="500" height="300" />
```

### Custom Layouts
For special slide layouts, you may need to add custom CSS or use Slidev's layout system.

## Customization

### Themes
To change the theme, edit the `generate_frontmatter()` function in `slidev_converter.py`:

```python
theme: seriph  # Change to: default, apple-basic, etc.
```

### Image Sizing
Modify image dimensions in the `convert_slide_to_markdown()` function:

```python
style="max-width: 80%; max-height: 400px; margin: 20px auto; display: block;"
```

### Background Images
Update the background URL in `generate_frontmatter()`:

```python
background: https://source.unsplash.com/1920x1080/?your-topic
```

## Troubleshooting

### "python-pptx not installed"
```bash
pip install python-pptx
```

### "Slidev command not found"
```bash
npm install -g @slidev/cli
```

### Images not displaying
- Check that image files are in the same directory as the .md file
- Verify image paths in the markdown are correct (should be `./filename.png`)

### Theme not applying
```bash
npm install -g @slidev/theme-seriph
```

### Vue parsing errors
The converter automatically escapes `<` and `>` characters, but if you see errors, check for any remaining unescaped angle brackets in your content.

## Contributing

Feel free to submit issues or pull requests to improve the converter!

## License

This project is open source. Feel free to use and modify as needed.