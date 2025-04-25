from PIL import Image, ImageDraw
import os


def create_icon():
    """Create a simple icon for the application."""
    # Create a 256x256 image with a transparent background
    img = Image.new("RGBA", (256, 256), color=(0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Draw a PDF icon (simple representation)
    # Draw the page
    draw.rectangle(
        [(48, 24), (208, 232)], fill=(255, 255, 255), outline=(100, 100, 100), width=3
    )

    # Draw the folded corner
    draw.polygon(
        [(168, 24), (208, 64), (168, 64)],
        fill=(230, 230, 230),
        outline=(100, 100, 100),
        width=3,
    )

    # Draw "PDF" text
    draw.rectangle([(72, 96), (184, 160)], fill=(220, 50, 50))

    # Draw text "METRICS" at the bottom
    draw.rectangle([(64, 180), (192, 210)], fill=(50, 50, 220))

    # Save the icon
    img.save(
        "app_icon.ico",
        format="ICO",
        sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)],
    )
    print(f"Icon created: {os.path.abspath('app_icon.ico')}")


if __name__ == "__main__":
    try:
        from PIL import Image, ImageDraw

        create_icon()
    except ImportError:
        print("Error: Pillow library is required. Install it with 'pip install pillow'")
