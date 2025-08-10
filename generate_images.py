
import os
from PIL import Image, ImageDraw, ImageFont
import random

def get_random_color():
    return (random.randint(0, 255), random.randint(0, 255), random.randint(0, 255))

def get_font(font_path, size):
    try:
        return ImageFont.truetype(font_path, size)
    except IOError:
        print(f"Font not found at {font_path}. Please ensure MS Gothic is installed or specify a different font.")
        # Fallback to a default font if MS Gothic is not found
        return ImageFont.load_default()

def generate_image(image_id, width=1024, height=768):
    img = Image.new('RGB', (width, height))
    draw = ImageDraw.Draw(img)

    # Generate a background with 100x100px color charts
    for y in range(0, height, 100):
        for x in range(0, width, 100):
            color = get_random_color()
            draw.rectangle([x, y, x + 100, y + 100], fill=color)

    # Define font and text colors
    font_path = "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"
    font_size = 24
    font = get_font(font_path, font_size)

    # Draw text with black background
    text_configs = [
        ((50, 100), str(random.randint(100, 999)), "yellow"),
        ((100, 200), str(random.randint(1000, 9999)), "white"),
        ((150, 300), str(random.randint(10000, 99999)), "green"),
        ((500, 400), str(random.randint(100000, 999999)), "red"),
    ]

    for (x, y), text, color in text_configs:
        # Get text size to draw black background
        bbox = draw.textbbox((x, y), text, font=font)
        draw.rectangle(bbox, fill="black")
        draw.text((x, y), text, font=font, fill=color)

    img.save(f"img/image_{image_id:02d}.jpeg", "jpeg")

if __name__ == '__main__':
    if not os.path.exists("img"):
        os.makedirs("img")

    for i in range(10):
        generate_image(i)
    print("Generated 10 JPEG images in the 'img' directory.")
