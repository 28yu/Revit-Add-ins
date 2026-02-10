#!/usr/bin/env python3
"""
Generate filled_region_32.png icon for Tools28 Revit Add-in
Design: 4 small squares (4x4px) with hatching in 2x2 layout
"""

from PIL import Image, ImageDraw

# Create 32x32 image with transparent background
img = Image.new('RGBA', (32, 32), (255, 255, 255, 0))
draw = ImageDraw.Draw(img)

# Parameters (matching IconGenerator.html with squareSize = 8 / 2 = 4px at 1x resolution)
square_size = 4  # 4px squares
line_width = 1   # 0.5px × 2
gap = 2          # gap between squares

# Calculate starting position for 2x2 grid (centered)
total_size = square_size * 2 + gap
start_x = (32 - total_size) // 2
start_y = (32 - total_size) // 2

# Draw 4 squares in 2x2 layout
positions = [
    (start_x, start_y),                                    # top-left
    (start_x + square_size + gap, start_y),                # top-right
    (start_x, start_y + square_size + gap),                # bottom-left
    (start_x + square_size + gap, start_y + square_size + gap)  # bottom-right
]

for x, y in positions:
    # Draw hatching (45-degree diagonal lines)
    # Draw multiple diagonal lines from top-left to bottom-right
    for i in range(-square_size, square_size * 2, 2):  # spacing of 2px (1.2px × 2 ≈ 2)
        # Line from (x+i, y) to (x+i+square_size, y+square_size)
        draw.line(
            [(x + i, y), (x + i + square_size, y + square_size)],
            fill=(0, 0, 0, 255),
            width=1
        )

    # Draw square border
    draw.rectangle(
        [x, y, x + square_size - 1, y + square_size - 1],
        outline=(0, 0, 0, 255),
        width=1
    )

# Save the icon
img.save('Resources/Icons/filled_region_32.png', 'PNG')
print("✓ Icon generated: Resources/Icons/filled_region_32.png")
print(f"  Size: {img.size}")
print(f"  Mode: {img.mode}")
