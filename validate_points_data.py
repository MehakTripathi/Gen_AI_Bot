import re
import openai
import json
from config import GROUP_CHAT_ID, openai

def clean_json_response(formatted_data):
    """
    Cleans and extracts valid JSON array content from the response.
    Handles:
      - Extra backticks (`json ...`)
      - JSON inside strings
      - Extracting JSON using regex as a fallback
    """
    if not formatted_data:
        print("❌ Empty input received!")
        return None
    
    # Remove leading/trailing backticks (`json ...`)
    formatted_data = re.sub(r"^(`{1,3}json)?|(`{1,3})$", "", formatted_data.strip()).strip()

    # First attempt: Try to parse JSON directly
    try:
        parsed_data = json.loads(formatted_data)
        if isinstance(parsed_data, list):  # Ensure it's a list
            return parsed_data
    except json.JSONDecodeError:
        pass  # If JSON decoding fails, continue to regex method

    # Fallback: Use regex to extract the JSON array
    json_match = re.search(r'(\[\s*\{[\s\S]*?\}\s*\])', formatted_data, re.UNICODE)
    
    if json_match:
        json_content = json_match.group(1).strip()
        try:
            return json.loads(json_content)  # ✅ Properly parse JSON into Python object
        except json.JSONDecodeError as e:
            print(f"❌ JSON Parsing Error: {e}")
            return None
    else:
        print("❌ No valid JSON detected in response.")
        return None
    
def get_bullet_points_data(extracted_text, max_retries=3, max_tokens=500):
    """
    Queries OpenAI to retrieve structured bullet points from extracted text.

    Args:
        extracted_text (str): The text extracted from an image.
        max_retries (int): Maximum number of retries for API calls.
        max_tokens (int): Maximum number of tokens to be returned in the API response.

    Returns:
        list: A structured list of bullet points.
    """
    all_data = []

    prompt = (
        "Extract all **key bullet points** exactly as they appear in the image."
        "Do NOT modify, summarize, or create new text. Maintain original line breaks and formatting."
        "If the text contains structured information, extract it as bullet points."
        "If a title is present, include it in the output."
        "Maintain **original language**, prioritizing Hindi over English if both are present."
        "Ignore unnecessary text like footnotes, page numbers, or explanations."

        "\n\n### **Expected JSON Output Format**:"
        "\n```json"
        "\n["
        "\n  {"
        '    "title": "Title of the slide (if present)",'
        '    "points": ['
        '      "• Bullet point 1 text",'
        '      "• Bullet point 2 text",'
        '      "• Bullet point 3 text"'
        "    ]"
        "  }"
        "]"
        "\n```"
        "\n\nPROCESS THIS TEXT:"
        f"\n\n{extracted_text}"
    )

    attempts = 0
    while attempts < max_retries:
        try:
            response = openai.ChatCompletion.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant that strictly returns JSON."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=max_tokens
            )

            if "choices" not in response or not response["choices"]:
                print("❌ OpenAI API response is empty or incorrect!")
                attempts += 1
                continue

            formatted_data = response["choices"][0]["message"]["content"]
            print(f"📢 OpenAI Response: {formatted_data}")  # Debugging output

            cleaned_data = clean_json_response(formatted_data)
            if not cleaned_data:
                print("No valid JSON detected in response.")
                attempts+=1
                continue

            # If response is a string, convert it to JSON
            if isinstance(cleaned_data, str):
                structured_data = json.loads(cleaned_data)
            elif isinstance(cleaned_data, dict):  
                structured_data = [cleaned_data]  # Convert single dict to list
            elif isinstance(cleaned_data, list):
                structured_data = cleaned_data
            else:
                structured_data = cleaned_data  # Already parsed JSON

            all_data.extend(structured_data)  # Append extracted points
            break  # Exit loop on success

        except (json.JSONDecodeError, Exception) as e:
            print(f"Error on attempt {attempts + 1}: {e}")
            attempts += 1

    if attempts == max_retries:
        print(f"Failed after {max_retries} attempts. Exiting.")

    return all_data
