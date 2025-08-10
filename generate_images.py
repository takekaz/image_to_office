
from PIL import Image, ImageDraw, ImageFont
import random
import os

def generate_random_color():
    return (random.randint(0, 255), random.randint(0, 255), random.randint(0, 255))

def get_font(size):
    try:
        # Try to load MS Gothic font
        return ImageFont.truetype("/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf", size)
    except IOError:
        # Fallback to a default font if MS Gothic is not found
        print("IPA Gothic font not found or cannot be loaded. Using a default font.")
        try:
            return ImageFont.truetype("arial.ttf", size) # Common default font on Windows/Linux
        except IOError:
            return ImageFont.load_default() # Fallback to a basic built-in font


output_folder = "img"
os.makedirs(output_folder, exist_ok=True)

for i in range(10):
    img_width, img_height = 1024, 768
    img = Image.new('RGB', (img_width, img_height))
    draw = ImageDraw.Draw(img)

    # Generate colored background charts
    chart_size = 100
    for x in range(0, img_width, chart_size):
        for y in range(0, img_height, chart_size):
            color = generate_random_color()
            draw.rectangle([x, y, x + chart_size, y + chart_size], fill=color)

    # Text parameters
    font_size = 24
    font = get_font(font_size)

    # Text 1: Yellow, 3-digit random number at (50,100)
    text1 = str(random.randint(100, 999))
    text1_color = (255, 255, 0)  # Yellow
    text1_bg_color = (0, 0, 0)  # Black
    text1_position = (50, 100)
    draw.rectangle([text1_position[0], text1_position[1], text1_position[0] + draw.textlength(text1, font=font), text1_position[1] + font_size], fill=text1_bg_color)
    draw.text(text1_position, text1, font=font, fill=text1_color)

    # Text 2: White, 4-digit random number at (100,200)
    text2 = str(random.randint(1000, 9999))
    text2_color = (255, 255, 255)  # White
    text2_bg_color = (0, 0, 0)  # Black
    text2_position = (100, 200)
    draw.rectangle([text2_position[0], text2_position[1], text2_position[0] + draw.textlength(text2, font=font), text2_position[1] + font_size], fill=text2_bg_color)
    draw.text(text2_position, text2, font=font, fill=text2_color)

    # Text 3: Green, 5-digit random number at (150,300)
    text3 = str(random.randint(10000, 99999))
    text3_color = (0, 255, 0)  # Green
    text3_bg_color = (0, 0, 0)  # Black
    text3_position = (150, 300)
    draw.rectangle([text3_position[0], text3_position[1], text3_position[0] + draw.textlength(text3, font=font), text3_position[1] + font_size], fill=text3_bg_color)
    draw.text(text3_position, text3, font=font, fill=text3_color)

    # Text 4: Red, 6-digit random number at (500,400)
    text4 = str(random.randint(100000, 999999))
    text4_color = (255, 0, 0)  # Red
    text4_bg_color = (0, 0, 0)  # Black
    text4_position = (500, 400)
    draw.rectangle([text4_position[0], text4_position[1], text4_position[0] + draw.textlength(text4, font=font), text4_position[1] + font_size], fill=text4_bg_color)
    draw.text(text4_position, text4, font=font, fill=text4_color)

    img.save(os.path.join(output_folder, f"image_{i+1:02d}.jpeg"))
    print(f"Generated image_{i+1:02d}.jpeg")

print("Image generation complete.")
