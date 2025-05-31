# import telebot
# import requests
# from ocr import OCRHandler
# from config import BOT_TOKEN, GROUP_CHAT_ID
# from ppt_generator import PPTHandler


# class TelegramBot:
#     def _init_(self):
import telebot
import requests
from ocr import OCRHandler
from config import BOT_TOKEN, GROUP_CHAT_ID
from ppt_generator import PPTHandler
from ai_presentation_generator import AIPresentationGenerator
from ocr_points_handler import OCRPointsHandler 

class TelegramBot:
    def __init__(self):
        """
        Initializes the Telegram bot with OCR and PPT handlers.
        """
        self.bot = telebot.TeleBot(BOT_TOKEN)
        self.ocr_handler = OCRHandler(self.bot, GROUP_CHAT_ID)  # MCQ Handler
        self.ocr_points_handler = OCRPointsHandler(self.bot, GROUP_CHAT_ID)  # Points Handler
        self.question_data = {"Question": [], "Options": [], "Points": []}  # Add Points Data
        self.ppt_handler = PPTHandler()
        self.bullet_points =AIPresentationGenerator()

        # Default presentation settings
        self.presentation_settings = {
            "title": "NEXT LEVEL ACADEMY",
            "topic": "General Questions",
            "teacher_name": "Instructor",
            "ppt_type":"mcq"
        }
        self.register_handlers()


    def generate_ppt(self, chat_id):
        """
        Generates a PowerPoint presentation from accumulated question data and sends it to the user.
        """
        try:
            ppt_type = self.presentation_settings["ppt_type"]
            if ppt_type == "mcq":
                output_file = self.ppt_handler.create_custom_presentation(
                    data=self.question_data,
                    title=self.presentation_settings["title"],
                    topic=self.presentation_settings["topic"],
                    teacher_name=self.presentation_settings["teacher_name"],
                )

            elif ppt_type == "points":
                output_file = self.bullet_points.create_presentation(
                    extracted_data=self.question_data.get("Points",[]),
                    title=self.presentation_settings["title"],
                    topic=self.presentation_settings["topic"],
                    teacher_name=self.presentation_settings["teacher_name"]
                )
            else:
                self.bot.send_message(chat_id, "No data available to generate the presentation.")    

            # Send the generated presentation to the user
            with open(output_file, "rb") as f:
                self.bot.send_document(chat_id, f)

            # Clear question data after generating the presentation
            self.question_data = {key: [] for key in self.question_data}
            self.presentation_settings = {
                "title": "NEXT LEVEL ACADEMY",
                "topic": "General Questions",
                "teacher_name": "Instructor"
            }
            self.bot.send_message(chat_id, "Presentation generated and sent successfully!")
            
        except Exception as e:
            self.bot.send_message(chat_id, f"Error generating the presentation: {str(e)}")

    def register_handlers(self):
        """
        Registers message handlers for the Telegram bot.
        """
        @self.bot.message_handler(commands=['start'])
        def handle_start(message):
            welcome_text = (
                "Welcome to the PPT Generator Bot!\n\n"
                "Commands available:\n"
                "/set_title - Set the presentation title\n"
                "/set_topic - Set the topic\n"
                "/set_teacher - Set the teacher's name\n"
                "/set_type -Set the ppt type(Bullet points or MCQ)\n"
                "/status - Check current settings and question count\n"
                "Send images to add questions\n"
                "Send 'nextlevel' to generate the presentation"
            )
            self.bot.reply_to(message, welcome_text)

        @self.bot.message_handler(commands=['set_title'])
        def handle_set_title(message):
            self.bot.send_message(
                message.chat.id, 
                "Please enter the presentation title:"
            )
            self.bot.register_next_step_handler(message, self.save_title)

    

        @self.bot.message_handler(commands=['set_topic'])
        def handle_set_topic(message):
            self.bot.send_message(
                message.chat.id, 
                "Please enter the topic name:"
            )
            self.bot.register_next_step_handler(message, self.save_topic)



        @self.bot.message_handler(commands=['set_teacher'])
        def handle_set_teacher(message):
            self.bot.send_message(
                message.chat.id, 
                "Please enter the teacher's name:"
            )
            self.bot.register_next_step_handler(message, self.save_teacher)

        @self.bot.message_handler(commands=['set_type'])
        def handle_set_type(message):
            self.bot.send_message(
                message.chat.id, 
                "Please enter the ppt's type(mcq or points):"
            )
            self.bot.register_next_step_handler(message, self.save_type)
            

        @self.bot.message_handler(commands=['status'])
        def handle_status(message):
            status = (
                "Current Settings:\n"
                f"Title: {self.presentation_settings['title']}\n"
                f"Topic: {self.presentation_settings['topic']}\n"
                f"Teacher: {self.presentation_settings['teacher_name']}\n"
                f"Type: {self.presentation_settings['ppt_type']}\n\n"
                f"Questions added: {len(self.question_data['Question'])}"
            )
            self.bot.reply_to(message, status)

     

        @self.bot.message_handler(content_types=['photo'])
        def handle_single_image(message):
            """
            Handles image messages, processes the image using OCR, and updates the question data.
            """
            try:
                photo = message.photo[-1]
                file_id = photo.file_id
                self.bot.send_message(message.chat.id, "Image received. Processing it now...")
                file_info = self.bot.get_file(file_id)
                if not file_info or not file_info.file_path:
                    self.bot.send_message(message.chat.id, "Error: File information could not be retrieved. Please try again.")
                    return

                file_path = f"temp_{file_id}.jpg"

                # Download the image
                self._download_file(file_path, file_info.file_path)

            
                # Use the correct OCR handler
                if self.presentation_settings["ppt_type"] == "mcq":
                    extracted_data = self.ocr_handler.extract_text_from_image(file_id)
                    self.ocr_handler.accumulate_questions(extracted_data, self.question_data)
                    self.bot.send_message(message.chat.id, f"Data extracted and added to the presentation. Total questions: {len(self.question_data['Question'])}")

                elif self.presentation_settings["ppt_type"] == "points":
                    extracted_data = self.ocr_points_handler.extract_text_from_image(file_id)
                    self.ocr_points_handler.accumulate_points(extracted_data, self.question_data["Points"])

                    self.bot.send_message(message.chat.id, f"Data extracted and added to the presentation. Total points: {len(self.question_data['Points'])}")

                    
                else:
                    self.bot.send_message(message.chat.id, "No text detected in the image. Please try another image.")
            except Exception as e:
                self.bot.send_message(message.chat.id, f"Error processing image: {str(e)}")

        @self.bot.message_handler(content_types=['document'])
        def handle_multiple_images(message):
            """
            Handles multiple image uploads, processes them using OCR, and updates the question data.
            """
            try:
                self.bot.send_message(message.chat.id, "Received multiple images. Processing...")
                for document in message.document:
                    file_id = document.file_id
                    file_info = self.bot.get_file(file_id)
                    file_path = f"temp_{file_id}.jpg"

                    # Download and process the image
                    self._download_file(file_path, file_info.file_path)
                    extracted_data = self.ocr_handler.extract_text_from_image(file_path)

                    if extracted_data:
                        self.ocr_handler.accumulate_questions(extracted_data)

                self.bot.send_message(
                    message.chat.id,
                    f"All images processed. Total questions: {len(self.question_data['Question'])}"
                )
            except Exception as e:
                self.bot.send_message(message.chat.id, f"Error processing images: {str(e)}")

        @self.bot.message_handler(func=lambda msg: msg.text.lower() == "nextlevel")
        def handle_nextlevel(message):
            """
            Handles the 'nextlevel' command to generate and send the PowerPoint presentation.
            """
            self.generate_ppt(message.chat.id)
            
    def save_title(self, message):
        """
        Saves the title and confirms it.
        """
        self.presentation_settings["title"] = message.text.strip()
        self.bot.send_message(
            message.chat.id,
            f"✅ Title set successfully: {self.presentation_settings['title']}"
        )

    def save_topic(self, message):
        """
        Saves the topic and confirms it.
        """
        self.presentation_settings["topic"] = message.text.strip()
        self.bot.send_message(
            message.chat.id,
            f"✅ Topic set successfully: {self.presentation_settings['topic']}"
        )

    def save_teacher(self, message):
        """
        Saves the teacher's name and confirms it.
        """
        self.presentation_settings["teacher_name"] = message.text.strip()
        self.bot.send_message(
            message.chat.id,
            f"✅ Teacher's name set successfully: {self.presentation_settings['teacher_name']}"
        )

    def save_type(self, message):
        """
        Saves the ppt's type and confirms it.
        """
        self.presentation_settings["ppt_type"] = message.text.strip().lower()
        self.bot.send_message(
            message.chat.id,
            f"✅ PPT type set successfully: {self.presentation_settings['ppt_type'].upper()}"
        ) 
         

    def _download_file(self, save_path, file_path):
        """
        Downloads a file from Telegram's servers.
        """
        file_url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_path}"
        with open(save_path, "wb") as f:
            f.write(requests.get(file_url).content)

    def start(self):
        """
        Starts polling for Telegram bot updates.
        """
        self.bot.polling(none_stop=True)