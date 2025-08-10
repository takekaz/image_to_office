
import os
import random
from PIL import Image, ImageDraw, ImageFont

def generate_color_chart(width, height, block_size=100):
    """Generates a background with a color chart."""
    img = Image.new('RGB', (width, height))
    draw = ImageDraw.Draw(img)

    for y in range(0, height, block_size):
        for x in range(0, width, block_size):
            r = random.randint(0, 255)
            g = random.randint(0, 255)
            b = random.randint(0, 255)
            draw.rectangle([x, y, x + block_size, y + block_size], fill=(r, g, b))
    return img

def generate_image(image_id, output_dir="img"):
    width, height = 1024, 768
    
    # Try to load MS Gothic font, fallback to a common sans-serif font
    try:
        # This path might vary depending on the system.
        # Common locations for MS Gothic on Windows: "C:/Windows/Fonts/msgothic.ttc"
        # On Linux, MS Gothic might not be present by default. We'll use a generic font.
        font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
        font = ImageFont.truetype(font_path, 24)
    except IOError:
        print(f"Warning: Font not found at {font_path}. Using default Pillow font.")
        font = ImageFont.load_default()

    img = generate_color_chart(width, height)
    draw = ImageDraw.Draw(img)

    # Text data: (position, color, length, label)
    text_configs = [
        ((50, 100), "yellow", 3, "3-digit random"),
        ((100, 200), "white", 4, "4-digit random"),
        ((150, 300), "green", 5, "5-digit random"),
        ((500, 400), "red", 6, "6-digit random")
    ]

    for (x, y), color, length, _ in text_configs:
        random_number = str(random.randint(10**(length-1), 10**length - 1))
        
        # Get text size for background
        # Older Pillow versions might need getsize, newer ones use textbbox
        try:
            bbox = draw.textbbox((x, y), random_number, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
        except AttributeError:
            text_width, text_height = draw.textsize(random_number, font=font)

        # Draw black background rectangle for text
        padding = 5 # Adjust padding as needed
        draw.rectangle([x - padding, y - padding, x + text_width + padding, y + text_height + padding], fill=(0, 0, 0))
        
        # Draw text
        draw.text((x, y), random_number, font=font, fill=color)

    output_path = os.path.join(output_dir, f"sample_image_{image_id:02d}.jpeg")
    img.save(output_path, "jpeg")
    print(f"Generated {output_path}")

if __name__ == "__main__":
    output_directory = "img"
    os.makedirs(output_directory, exist_ok=True)
    for i in range(1, 11):
        generate_image(i, output_directory)
