from PIL import Image, ImageDraw
import numpy as np

img = Image.open('/Users/anhpt/Code/CheckFormKLTN/Check KLTN/Logo TLU.jpg').convert('RGBA')

# The logo is inside an oval. We can create an oval mask.
# Or we can make near-white pixels transparent.
data = np.array(img)
r, g, b, a = data.T
white_areas = (r > 240) & (g > 240) & (b > 240)
data[...][white_areas.T] = (255, 255, 255, 0)

img_out = Image.fromarray(data)

# To be safe and "chỉ lấy phần logo trong phần bầu dục bỏ phần nền trắng ở ngoài"
# Let's create an ellipse mask
w, h = img.size
mask = Image.new('L', (w, h), 0)
draw = ImageDraw.Draw(mask)
# Usually logos have a bit of padding. We can draw an ellipse that fits the bounds.
# Let's try drawing an ellipse slightly smaller than the full bounds, or just exactly the bounds.
# A better way is to find the bounding box of non-white pixels and then draw an ellipse.
# Let's see bounding box:
non_white = ~white_areas.T
coords = np.argwhere(non_white)
if len(coords) > 0:
    y0, x0 = coords.min(axis=0)
    y1, x1 = coords.max(axis=0)
    print(f"Bounding box of non-white: {x0}, {y0}, {x1}, {y1}")
    draw.ellipse((x0, y0, x1, y1), fill=255)
    img_out.putalpha(mask)
    # crop to bounding box
    img_out = img_out.crop((x0, y0, x1, y1))
    
img_out.save('/Users/anhpt/Code/CheckFormKLTN/web_static/logo.png')
print("Saved logo.png")
