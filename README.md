# PowerPoint Resizer & Organizer

Automate the tedious task of resizing PowerPoint presentations while maintaining perfect organization and alignment. No more manually readjusting elements!

## Features

- **Automatic Resizing**: Resize presentations to any dimensions (default: 36x48 inches for posters)
- **Smart Reorganization**: Intelligently scales and repositions all elements proportionally
- **Perfect Alignment**: Grid-snapping ensures nothing is off by even a pixel
- **Maintains Layout**: Preserves the relative positioning and relationships between elements
- **Font Scaling**: Automatically adjusts text sizes to maintain readability
- **Handles All Elements**: Works with text boxes, shapes, images, tables, and more

## Installation

1. **Install Python** (if not already installed):
   - Download from [python.org](https://www.python.org/downloads/)
   - Python 3.7 or higher required

2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

   Or directly:
   ```bash
   pip install python-pptx
   ```

## Usage

### Basic Usage

Resize a PowerPoint to 36x48 inches (poster size):

```bash
python resize_powerpoint.py input.pptx output.pptx
```

### Custom Dimensions

Specify custom width and height in inches:

```bash
python resize_powerpoint.py input.pptx output.pptx --width 24 --height 36
```

### Advanced Options

```bash
# Disable grid alignment for faster processing (may be slightly less precise)
python resize_powerpoint.py input.pptx output.pptx --no-grid

# Full help
python resize_powerpoint.py --help
```

## How It Works

1. **Loads** your PowerPoint presentation
2. **Calculates** scale factors based on original and target dimensions
3. **Resizes** the slide canvas to target dimensions
4. **Scales** all elements proportionally (position, size, fonts)
5. **Aligns** everything to a grid for perfect precision
6. **Saves** the reorganized presentation

## Example Workflow

```bash
# 1. Place your PowerPoint file in this directory
# 2. Run the script
python resize_powerpoint.py my_poster.pptx my_poster_36x48.pptx

# 3. Open the output file - everything will be perfectly organized!
```

## What Gets Resized

- ✓ Text boxes and their fonts
- ✓ Shapes and their dimensions
- ✓ Images and their placement
- ✓ Tables and cells
- ✓ Charts and diagrams
- ✓ Groups of objects
- ✓ Position and spacing between all elements

## Tips for Best Results

1. **Save a backup** of your original file before processing
2. **Use .pptx format** (not .ppt) - the modern PowerPoint format
3. **Check the output** - while the script is very accurate, you may want to fine-tune specific elements
4. **Grid alignment** is enabled by default for pixel-perfect precision
5. **Font scaling** maintains text proportions but you can adjust afterward if needed

## Troubleshooting

**"Could not load PowerPoint file"**
- Ensure the file is .pptx format (not .ppt)
- Check that the file isn't corrupted
- Make sure the file isn't open in PowerPoint

**"Elements seem misaligned"**
- Try running without `--no-grid` to enable precise grid alignment
- Check if your original presentation had overlapping elements

**"Fonts are too small/large"**
- The script scales fonts proportionally
- You can manually adjust after processing if needed

## Requirements

- Python 3.7+
- python-pptx library

## License

MIT License - Feel free to use and modify!

## Need Help?

If you encounter any issues or need custom modifications, feel free to open an issue or reach out!

---

**No more tedious manual readjustments!** 🎉
