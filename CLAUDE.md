# CLAUDE.md - Project Context & Development Notes

## Project Overview

This is a PowerPoint to Slidev converter that transforms .pptx files (exported from Google Slides) into beautiful Slidev presentations. The project was developed to solve the challenge of converting existing Google Slides presentations to the modern Slidev format while preserving content, images, and formatting.

## Key Files

- `slidev_converter.py` - Main converter script
- `README.md` - User documentation  
- `output/` - Generated Slidev presentations and extracted images
- `Integrating AI in Java Projects.pptx` - Example input file

## Architecture & Design Decisions

### Core Classes

1. **SlideContent** (dataclass)
   - Stores slide data: title, content, slide_type, notes, images
   - Images field contains list of filenames for extracted images

2. **SlidevConverter** 
   - Main conversion logic
   - Handles PowerPoint parsing and Slidev generation

3. **MCPSlidevConverter**
   - Wrapper class for MCP integration (future use)
   - Provides batch processing capabilities

### Key Technical Decisions

#### Text Processing
- **Multiword titles**: Use `' '.join(text.strip().split())` to normalize whitespace and remove newlines
- **Angle bracket escaping**: Replace `<` with `&lt;` and `>` with `&gt;` to prevent Vue parsing errors
- **Hierarchical bullets**: Store content as tuples `(level, text)` to preserve indentation

#### Image Handling
- Extract images using `shape.image.blob` from python-pptx
- Save as PNG files with naming convention: `slide_{num}_image_{count}.png`
- Use HTML `<img>` tags with CSS styling instead of markdown for better control:
  ```html
  <img src="./filename.png" style="max-width: 80%; max-height: 400px; margin: 20px auto; display: block;" />
  ```

#### Theme & Styling
- **Seriph theme**: Professional theme good for technical presentations
- **Custom CSS**: Added for table headers and image sizing
- **Background**: Unsplash integration for dynamic backgrounds

### Slide Generation Logic

1. **Frontmatter**: Contains theme, background, fonts, and metadata
2. **Title slide**: Flows directly from frontmatter (no separator)
3. **Content slides**: Use `<v-clicks>` for progressive disclosure
4. **Images**: Added after content with proper styling
5. **Closing slide**: Auto-generated thank you slide

## Known Issues & Limitations

### Conversion Losses
- **Tables**: Often lost in Google Slides → PowerPoint export, need manual recreation
- **YouTube videos**: Become static images, need manual replacement with `<Youtube>` component
- **Complex layouts**: May need manual adjustment
- **Animations**: Not preserved (Slidev uses different animation system)

### Manual Fixes Required
1. **Tables**: Recreate using markdown table syntax
2. **YouTube videos**: Replace image with `<Youtube id="VIDEO_ID" />` component
3. **Special formatting**: May need custom CSS

## Development History & Bug Fixes

### Major Issues Solved

1. **Title slide placement**: Initially created empty slide with background only
   - **Fix**: Skip separator before first slide so frontmatter flows into title

2. **Double bullet display**: Showed both markdown bullets and emoji
   - **Fix**: Use proper markdown bullets, removed emoji duplicates

3. **Vue parsing errors**: Angle brackets in content caused errors
   - **Fix**: Escape `<` and `>` characters in all text content

4. **Lost hierarchical structure**: Flat bullet lists instead of nested
   - **Fix**: Parse paragraph levels from PowerPoint and preserve with indentation

5. **Image sizing issues**: Images displayed at full size, breaking layout
   - **Fix**: Added CSS styling with max-width and max-height constraints

6. **Theme not applying**: Default theme looked basic
   - **Fix**: Switched to Seriph theme, added proper theme installation

## Code Patterns & Best Practices

### Error Handling
- Wrap image extraction in try/catch blocks
- Graceful fallbacks for missing text frames
- Print warnings for failed operations

### Content Processing
```python
# Text normalization pattern
title = ' '.join(slide.shapes.title.text.strip().split())

# Angle bracket escaping pattern  
text = text.replace('<', '&lt;').replace('>', '&gt;')

# Hierarchical content pattern
content.append((level, text))  # Store as tuple
```

### Image Management
- Store image data temporarily during parsing
- Save all images at once after slide processing
- Use descriptive filenames with slide numbers

## Testing & Validation

### Test File
- `Integrating AI in Java Projects.pptx` - 60 slides with various content types
- Contains: text, bullets, images, tables, YouTube video
- Good test case for edge cases and complex content

### Validation Checklist
- [ ] Title displays correctly on first slide
- [ ] Bullet hierarchies preserved  
- [ ] Images extracted and sized properly
- [ ] No Vue parsing errors
- [ ] Theme applies correctly
- [ ] Tables render (if manually added)
- [ ] YouTube videos work (if manually replaced)

## Future Enhancements

### Potential Improvements
1. **Better table detection**: Attempt to preserve table structure
2. **Video link detection**: Recognize YouTube URLs in slide notes
3. **Layout detection**: Identify special slide layouts
4. **Font preservation**: Maintain custom fonts where possible
5. **Animation mapping**: Convert PowerPoint animations to Slidev equivalents

### Code Quality
- Add type hints throughout
- Create unit tests for core functions
- Add logging for debugging
- Configuration file for themes/settings

## Usage Notes for Future Claude Sessions

### Common User Requests
1. **Theme changes**: Modify `generate_frontmatter()` function
2. **Image sizing**: Adjust CSS in `convert_slide_to_markdown()`
3. **Background images**: Change URL in frontmatter generation
4. **Table formatting**: Add custom CSS styling
5. **YouTube integration**: Replace images with `<Youtube>` components

### Debugging Tips
- Check for Vue parsing errors in browser console
- Verify image file paths are relative (`./filename.png`)
- Ensure theme is installed (`npm install -g @slidev/theme-name`)
- Look for unescaped angle brackets in content

### Performance Considerations
- Large image files may slow rendering
- Many images per slide can cause layout issues
- Complex tables may need manual optimization

## Integration Points

### Slidev Features Used
- Frontmatter configuration
- v-clicks for animations
- YouTube component
- Custom CSS styling
- Theme system
- Markdown tables

### Dependencies
- `python-pptx`: PowerPoint file parsing
- `pathlib`: File path handling
- `re`: Text processing and filename sanitization