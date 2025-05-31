import openai
import cv2
import os
import requests
import datetime
import telebot
from PIL import Image
import easyocr
from config import BOT_TOKEN, GROUP_CHAT_ID
from validate_points_data import get_bullet_points_data  # Import the new function


class OCRPointsHandler:
    def __init__(self, bot, group_chat_id):
        self.bot = bot
        self.group_chat_id = group_chat_id
        self.reader = easyocr.Reader(['en', 'hi'])  # Supports English and Hindi

    def preprocess_image(self, image_path):
        """
        Preprocesses image to enhance contrast and clarity for OCR.
        """
        img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)  
        img = cv2.resize(img, (1080, 1080))  
        img = cv2.GaussianBlur(img, (5, 5), 0)  
        _, img = cv2.threshold(img, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)  
        processed_path = "processed_" + image_path
        cv2.imwrite(processed_path, img)
        return processed_path

    def extract_text_from_image(self, file_id):
        """
        Extracts bullet points from an image using OpenAI Vision API.
        """
        local_file_path = None
        try:
            # ✅ Step 1: Get Image from Telegram
            file_info = self.bot.get_file(file_id)
            file_url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_info.file_path}"
            local_file_path = f"temp_image_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.jpg"

            # ✅ Step 2: Download Image
            response = requests.get(file_url, stream=True)
            response.raise_for_status()
            with open(local_file_path, "wb") as f:
                f.write(response.content)

            # ✅ Step 3: Preprocess Image
            processed_image_path = self.preprocess_image(local_file_path)

            # ✅ Step 4: Extract Bullet Points
            extracted_text = self.get_text_from_openai(processed_image_path)
            
            if extracted_text:
                print("Extracted Points:", extracted_text)
                return extracted_text
            else:
                raise ValueError("No bullet points detected in the image.")

        except requests.exceptions.RequestException as e:
            error_message = f"Failed to download file: {e}"
            print(error_message)
            self.bot.send_message(self.group_chat_id, error_message)
            return ""

        except Exception as e:
            error_message = f"Error during OCR processing: {str(e)}"
            print(error_message)
            self.bot.send_message(self.group_chat_id, error_message)
            return ""

        finally:
            # Cleanup Temporary File
            if local_file_path and os.path.exists(local_file_path):
                os.remove(local_file_path)

    def get_text_from_openai(self, image_path):
        """
        Uses OpenAI Vision API to extract **structured bullet points** from an image.
        """
        try:
            with open(image_path, "rb") as image_file:
                response = openai.ChatCompletion.create(
                    model="gpt-4o",
                    messages=[
                        {"role": "system", "content": "You are an AI expert that extracts key bullet points accurately."},
                        {"role": "user", "content": [
                            {"type": "text", "text": (
                            "Extract **key bullet points** from the given image."
                            "Do not extract explanations or unnecessary text."
                            "Maintain the exact structure and order of points."
                            "If a point is present in both Hindi and English, prioritize Hindi and ignore English."
                            "Extract **only meaningful information**, skipping irrelevant details."
                            "\n\n### **Expected Output Format (JSON)**:"
                            "\n```json"
                            "\n["
                            "\n  {"
                            "\n    \"title\": \"Title of the slide (if present)\","
                            "\n    \"points\": ["
                            "\n      \"• Point 1 exactly as in the image\","
                            "\n      \"• Point 2 exactly as in the image\","
                            "\n      \"• Point 3 exactly as in the image\""
                            "\n    ]"
                            "\n  }"
                            "\n]"
                            "\n```"
                        )},
                            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{self.encode_image(image_path)}"}}
                        ]}
                    ],
                    max_tokens=2000
                )

            extracted_text = response["choices"][0]["message"]["content"]
            return extracted_text

        except Exception as e:
            print(f"OpenAI Vision API Error: {e}")
            return None

    def encode_image(self, image_path):
        """
        Encodes an image file as a Base64 string.
        """
        import base64
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode("utf-8")

    def accumulate_points(self, extracted_text, points_data):
        """
        Accumulates extracted bullet points into points_data.
        """
        try:
            structured_data = get_bullet_points_data(extracted_text)  # Call new function
            points_data.extend(structured_data)  # Append extracted points
        except Exception as e:
            print(f"Error in accumulating bullet points: {e}")
            self.bot.send_message(self.group_chat_id, f"Error processing extracted bullet points: {e}")
