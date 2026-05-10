from PIL import Image
import os

img_path = r'd:\repos\tracer-study-2025\assets\gambar\running_worker.jpg'
out_path = r'd:\repos\tracer-study-2025\assets\gambar\running_worker.png'

if os.path.exists(img_path):
    img = Image.open(img_path).convert("RGBA")
    datas = img.getdata()

    newData = []
    for item in datas:
        # If the pixel is very light (white/near-white), make it transparent
        if item[0] > 245 and item[1] > 245 and item[2] > 245:
            newData.append((255, 255, 255, 0))
        else:
            newData.append(item)

    img.putdata(newData)
    img.save(out_path, "PNG")
    print(f"Saved transparent PNG to {out_path}")
else:
    print("Source image not found.")
