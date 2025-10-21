import os
import sys
from PIL import Image
from io import BytesIO
import pytesseract

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Configure Tesseract path
pytesseract.pytesseract.tesseract_cmd = resource_path(os.path.join("Tesseract-OCR", "tesseract.exe"))

class CaptchaHandler:
    target_color = (102, 102, 102)

    def get_captcha_from_image(self, target_image):
        """Process CAPTCHA image and return text"""
        try:
            image_data = BytesIO(target_image)
            image = Image.open(image_data).convert("RGB")
            width, height = image.size

            white_image = Image.new("RGB", (width, height), "white")

            for x in range(width):
                for y in range(height):
                    pixel = image.getpixel((x, y))
                    if pixel == self.target_color:
                        white_image.putpixel((x, y), pixel)

            return pytesseract.image_to_string(white_image).replace(" ", "").strip()
        except Exception as e:
            print(f"CAPTCHA processing error: {e}")
            return ""