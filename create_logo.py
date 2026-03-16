from PIL import Image, ImageDraw, ImageFont

def create_logo():
    # Base canvas
    w, h = 1000, 450
    # Background color: #16527A (Dark blue/teal)
    img = Image.new('RGB', (w, h), (22, 82, 122))
    d = ImageDraw.Draw(img, 'RGBA')

    # Central Oval Outline
    cx, cy = w // 2, 170
    oval_w, oval_h = 600, 300
    
    # Draw thicker white oval
    d.ellipse([cx - oval_w//2, cy - oval_h//2, cx + oval_w//2, cy + oval_h//2], outline=(255, 255, 255, 255), width=10)
    
    # Erase the bottom-left gap for the swoosh entry
    d.polygon([(100, 150), (450, 150), (450, 400), (100, 400)], fill=(22, 82, 122, 255))
    
    # Redraw oval with arc to leave a precise gap
    # Angle 0 is 3 o'clock, 90 is 6 o'clock
    d.arc([cx - oval_w//2, cy - oval_h//2, cx + oval_w//2, cy + oval_h//2], start=-140, end=115, fill=(255, 255, 255, 255), width=10)

    # 3 Sweeping Swooshes forming an 'A'
    # Base swoosh (rightmost curve of swoop)
    swoosh1 = [(320, 300), (520, 150), (620, 60), (660, 40), (700, 30), (660, 60), (540, 170), (360, 300)]
    d.polygon(swoosh1, fill=(255, 255, 255, 255))
    
    # Middle swoosh
    swoosh2 = [(260, 270), (460, 150), (620, 50), (660, 60), (500, 160), (300, 270)]
    d.polygon(swoosh2, fill=(255, 255, 255, 255))
    
    # Left swoosh
    swoosh3 = [(200, 240), (400, 150), (580, 70), (620, 80), (440, 160), (240, 240)]
    d.polygon(swoosh3, fill=(255, 255, 255, 255))
    
    # The right leg of the A bridging the swooshes down
    right_leg = [(620, 60), (680, 150), (740, 230), (690, 230), (640, 150), (580, 80)]
    d.polygon(right_leg, fill=(255, 255, 255, 255))
    
    # Horizontal crossbar of 'A'
    d.polygon([(500, 150), (650, 150), (660, 170), (490, 170)], fill=(255, 255, 255, 255))

    # The Blue Star/Cross
    star_cx, star_cy = 670, 40
    d.polygon([(star_cx, star_cy - 40), (star_cx + 10, star_cy - 10), (star_cx + 40, star_cy), (star_cx + 10, star_cy + 10), (star_cx, star_cy + 40), (star_cx - 10, star_cy + 10), (star_cx - 40, star_cy), (star_cx - 10, star_cy - 10)], fill=(54, 150, 203, 255))
    
    # Small sparkles
    d.ellipse([720, 20, 730, 30], fill=(255, 255, 255, 255))
    d.ellipse([640, -10, 650, 0], fill=(255, 255, 255, 255))

    # Text "PathAxiom"
    try:
        font_large = ImageFont.truetype("arialbd.ttf", 90)
        font_a = ImageFont.truetype("arialbd.ttf", 120)
    except:
        font_large = ImageFont.load_default()
        font_a = font_large
        
    d.text((50, 310), "Path", fill=(255, 255, 255, 255), font=font_large)
    d.text((370, 290), "A", fill=(12, 150, 165, 255), font=font_a)  # Teal large 'A'
    d.text((460, 310), "xiom", fill=(255, 255, 255, 255), font=font_large)

    # Broken White underline
    d.line([(50, 410), (360, 410)], fill=(255, 255, 255, 255), width=6)
    d.line([(450, 410), (950, 410)], fill=(255, 255, 255, 255), width=6)
    
    # End sparkle on line
    d.polygon([(950, 395), (955, 405), (965, 410), (955, 415), (950, 425), (945, 415), (935, 410), (945, 405)], fill=(255, 255, 255, 255))

    # Resize to required size
    img = img.resize((472, 221), Image.LANCZOS)

    img.save("logo.png", dpi=(300,300))

    print("✅ High-fidelity PathAxiom logo generated as logo.png")

if __name__ == '__main__':
    create_logo()
