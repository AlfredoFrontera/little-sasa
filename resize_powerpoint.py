#!/usr/bin/env python3
"""
PowerPoint Resizer and Organizer
Automatically resize PowerPoint presentations to 36x48 inches and reorganize all elements
to maintain perfect layout and alignment.
"""

from pptx import Presentation
from pptx.util import Inches
import argparse
import os
import sys


class PowerPointResizer:
    def __init__(self, input_file, output_file, target_width=36, target_height=48):
        """
        Initialize the PowerPoint Resizer

        Args:
            input_file: Path to input PowerPoint file
            output_file: Path to output PowerPoint file
            target_width: Target width in inches (default: 36)
            target_height: Target height in inches (default: 48)
        """
        self.input_file = input_file
        self.output_file = output_file
        self.target_width = Inches(target_width)
        self.target_height = Inches(target_height)

    def calculate_scale_factors(self, original_width, original_height):
        """
        Calculate scale factors for resizing
        Maintains aspect ratio awareness and calculates both width and height scales
        """
        width_scale = self.target_width / original_width
        height_scale = self.target_height / original_height
        return width_scale, height_scale

    def resize_and_reposition_shape(self, shape, width_scale, height_scale):
        """
        Resize and reposition a shape based on scale factors
        Handles position, size, and maintains relative positioning
        """
        try:
            # Scale position (top-left corner)
            if hasattr(shape, 'left') and hasattr(shape, 'top'):
                shape.left = int(shape.left * width_scale)
                shape.top = int(shape.top * height_scale)

            # Scale size
            if hasattr(shape, 'width') and hasattr(shape, 'height'):
                shape.width = int(shape.width * width_scale)
                shape.height = int(shape.height * height_scale)

            # Handle text frame font sizes
            if hasattr(shape, 'text_frame'):
                self.scale_text_frame(shape.text_frame, min(width_scale, height_scale))

        except Exception as e:
            print(f"Warning: Could not fully resize shape: {e}")

    def scale_text_frame(self, text_frame, scale_factor):
        """
        Scale font sizes in text frame to maintain readability
        """
        try:
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.font.size:
                        # Scale font size
                        new_size = int(run.font.size * scale_factor)
                        run.font.size = new_size
        except Exception as e:
            print(f"Warning: Could not scale text: {e}")

    def align_shapes_to_grid(self, shapes, grid_size=Inches(0.1)):
        """
        Snap shapes to a grid for perfect alignment
        This ensures nothing is off by even a tiny bit
        """
        for shape in shapes:
            try:
                if hasattr(shape, 'left') and hasattr(shape, 'top'):
                    # Snap to grid
                    shape.left = int(shape.left / grid_size) * grid_size
                    shape.top = int(shape.top / grid_size) * grid_size

                if hasattr(shape, 'width') and hasattr(shape, 'height'):
                    # Snap dimensions to grid
                    shape.width = int(shape.width / grid_size) * grid_size
                    shape.height = int(shape.height / grid_size) * grid_size
            except Exception as e:
                print(f"Warning: Could not align shape to grid: {e}")

    def organize_slide(self, slide, width_scale, height_scale, align_to_grid=True):
        """
        Organize all elements on a slide
        Resize and reposition all shapes while maintaining layout
        """
        # Process all shapes on the slide
        for shape in slide.shapes:
            self.resize_and_reposition_shape(shape, width_scale, height_scale)

        # Optionally align everything to a grid for perfect alignment
        if align_to_grid:
            self.align_shapes_to_grid(slide.shapes)

    def process_presentation(self, align_to_grid=True):
        """
        Main processing function
        Resizes the presentation and reorganizes all slides
        """
        print(f"Loading presentation: {self.input_file}")

        # Load the presentation
        try:
            prs = Presentation(self.input_file)
        except Exception as e:
            print(f"Error: Could not load PowerPoint file: {e}")
            return False

        # Get original dimensions
        original_width = prs.slide_width
        original_height = prs.slide_height

        print(f"Original size: {original_width / Inches(1):.2f} x {original_height / Inches(1):.2f} inches")
        print(f"Target size: {self.target_width / Inches(1):.2f} x {self.target_height / Inches(1):.2f} inches")

        # Calculate scale factors
        width_scale, height_scale = self.calculate_scale_factors(original_width, original_height)

        print(f"Scale factors: Width={width_scale:.3f}, Height={height_scale:.3f}")

        # Resize the slide dimensions
        prs.slide_width = self.target_width
        prs.slide_height = self.target_height

        # Process each slide
        total_slides = len(prs.slides)
        print(f"\nProcessing {total_slides} slide(s)...")

        for i, slide in enumerate(prs.slides, 1):
            print(f"  Processing slide {i}/{total_slides}...")
            self.organize_slide(slide, width_scale, height_scale, align_to_grid)

        # Save the modified presentation
        print(f"\nSaving to: {self.output_file}")
        try:
            prs.save(self.output_file)
            print("✓ Successfully saved!")
            return True
        except Exception as e:
            print(f"Error: Could not save PowerPoint file: {e}")
            return False


def main():
    parser = argparse.ArgumentParser(
        description='Resize and reorganize PowerPoint presentations to 36x48 inches',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic usage - resize to 36x48 inches
  python resize_powerpoint.py input.pptx output.pptx

  # Custom dimensions
  python resize_powerpoint.py input.pptx output.pptx --width 24 --height 36

  # Disable grid alignment (faster but less precise)
  python resize_powerpoint.py input.pptx output.pptx --no-grid
        """
    )

    parser.add_argument('input', help='Input PowerPoint file (.pptx)')
    parser.add_argument('output', help='Output PowerPoint file (.pptx)')
    parser.add_argument('--width', type=float, default=36,
                        help='Target width in inches (default: 36)')
    parser.add_argument('--height', type=float, default=48,
                        help='Target height in inches (default: 48)')
    parser.add_argument('--no-grid', action='store_true',
                        help='Disable grid alignment (elements may be slightly off)')

    args = parser.parse_args()

    # Validate input file
    if not os.path.exists(args.input):
        print(f"Error: Input file '{args.input}' does not exist")
        sys.exit(1)

    if not args.input.endswith('.pptx'):
        print("Warning: Input file should be .pptx format")

    # Create resizer and process
    resizer = PowerPointResizer(
        args.input,
        args.output,
        args.width,
        args.height
    )

    align_to_grid = not args.no_grid
    success = resizer.process_presentation(align_to_grid=align_to_grid)

    if success:
        print("\n✓ All done! Your presentation has been resized and organized.")
        sys.exit(0)
    else:
        print("\n✗ Failed to process presentation")
        sys.exit(1)


if __name__ == '__main__':
    main()
