from PIL import Image
import os

def resize_logo(input_path, output_path, size=(640, 480)):
    if not os.path.exists(input_path):
        print(f"Error: {input_path} does not exist.")
        return
    
    with Image.open(input_path) as img:
        # LANCZOS 필터를 사용하여 고품질 리사이즈 적용
        resized_img = img.resize(size, Image.Resampling.LANCZOS)
        resized_img.save(output_path, "PNG")
        print(f"Resized image saved successfully: {output_path}")

if __name__ == "__main__":
    src = r"C:\Users\레인써클\.gemini\antigravity-ide\brain\706fd70d-cee0-48f3-8529-96e39d506b84\media__1781597285531.png"
    dst = r"C:\Users\레인써클\.gemini\antigravity-ide\brain\706fd70d-cee0-48f3-8529-96e39d506b84\resized_emp_logo.png"
    resize_logo(src, dst)
