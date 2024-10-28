import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import os
import re
import json
from googleapiclient.discovery import build
from google.oauth2 import service_account
import requests
from bs4 import BeautifulSoup

pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'  # Update this path

folder_path = r'C:/Users/johna/Desktop/upwork jobs/tennis scrapping/screenshoots'

def resize_and_enhance_image(image, factor):
    image = image.resize((int(image.width * factor), int(image.height * factor)), Image.LANCZOS)
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2)  # Increase contrast
    return image

def apply_image_filters(image):
    image = image.filter(ImageFilter.SHARPEN)
    image = image.filter(ImageFilter.EDGE_ENHANCE_MORE)
    return image

def clean_text(text):
    lines = text.split('\n')
    cleaned_lines = []
    for line in lines:
        line = re.sub(r'(?<=\d)\.(?=\d{3})', '.', line)  # Fix common OCR errors
        line = re.sub(r'[^\d\.]', '', line)  # Remove any characters that are not digits or dots
        cleaned_lines.append(line)
    return '\n'.join(cleaned_lines)

def extract_text_from_image(image_path, max_attempts=10):
    image = Image.open(image_path)
    factor = 2.0  # Start with 2 times the original size
    
    for attempt in range(max_attempts):
        enhanced_image = resize_and_enhance_image(image, factor)
        filtered_image = apply_image_filters(enhanced_image)
        text = pytesseract.image_to_string(filtered_image, config='--psm 6')  # Use page segmentation mode 6 for better table OCR
        cleaned_text = clean_text(text)
        odds = parse_text_data(cleaned_text)
        if validate_extracted_text(odds):
            return cleaned_text
        factor += 0.5  # Increase the size factor for the next attempt

    return cleaned_text

def validate_extracted_text(odds):
    if len(odds) >= 5 and "26.00" in odds:
        differences = [float(odds[i+1]) - float(odds[i]) for i in range(min(len(odds)-1, 4))]
        if all(diff <= 5 for diff in differences):
            return True
    return False

def extract_text_from_images(folder_path, sheet, file_type):
    print(f"Looking for {file_type} images in folder: {folder_path}")
    headers = ["Player", "1+", "3+", "5+", "10+", "15+", "20+", "25+", "30+", "40+"]  # Known headers
    sheet_name = determine_sheet_name(file_type)
    set_headers(headers, sheet, sheet_name)  # Set the headers before processing images
    player_row = 2
    for filename in os.listdir(folder_path):
        if filename.endswith('.png') or filename.endswith('.jpg'):
            if not re.search(file_type, filename, re.IGNORECASE):
                continue
            image_path = os.path.join(folder_path, filename)
            print(f"Processing file: {filename}")
            text = extract_text_from_image(image_path)
            print(f"Extracted text from {filename}: {text[:100]}...")  # Print the first 100 characters of the extracted text
            odds = parse_text_data(text)
            if odds:
                if not validate_extracted_text(odds):
                    print(f"Data from {filename} does not meet criteria but will be added.")
                player_name = name_extraction(image_path)  # Extract player name from the image
                append_to_google_sheet(player_name, odds, headers, player_row, sheet, sheet_name)
                player_row += 1
            else:
                print(f"No valid data parsed from {filename}. Trying enhanced methods.")
                enhanced_text = retry_with_enhanced_methods(image_path)
                odds = parse_text_data(enhanced_text)
                if odds:
                    if not validate_extracted_text(odds):
                        print(f"Data from {filename} does not meet criteria after enhancement but will be added.")
                    player_name = name_extraction(image_path)  # Extract player name from the image
                    append_to_google_sheet(player_name, odds, headers, player_row, sheet, sheet_name)
                    player_row += 1
                else:
                    print(f"No valid data parsed from {filename} even after enhancements.")
        else:
            print(f"Skipping file (not an image): {filename}")

def determine_sheet_name(file_type):
    if re.search(r'\bat\b', file_type, re.IGNORECASE):
        return 'Ace Totals'
    elif re.search(r'\bdf\b', file_type, re.IGNORECASE):
        return 'Double Fouls'
    else:
        return 'Sheet1'  # Default sheet if no specific criteria are met

def retry_with_enhanced_methods(image_path):
    return extract_text_from_image(image_path, max_attempts=10)

def parse_text_data(text):
    lines = text.split('\n')
    odds = []

    for line in lines:
        line = line.strip()
        if re.match(r'^\d+(\.\d+)?$', line) and 0 < float(line) <= 26.0:
            odds.append(line)

    return odds

def set_headers(headers, sheet, sheet_name):
    SPREADSHEET_ID = '1sNwMydSoqasypK-jCEa5oCqsDNuHIgyAccsUjphczY0'  # Your actual spreadsheet ID
    RANGE_NAME = f'{sheet_name}!A1'  # Set headers starting from cell A1

    headers_data = [headers]

    body = {
        'values': headers_data
    }

    print(f"Setting headers in {sheet_name}: {headers}")

    try:
        result = sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME,
            valueInputOption='RAW',
            body=body
        ).execute()
        print(f'{result.get("updatedCells")} header cells updated in {sheet_name}.')
    except Exception as e:
        print(f"An error occurred while setting headers in {sheet_name}: {e}")

def clear_google_sheet(sheet):
    for sheet_name in ['Ace Totals', 'Double Fouls']:
        SPREADSHEET_ID = '1sNwMydSoqasypK-jCEa5oCqsDNuHIgyAccsUjphczY0'  # Your actual spreadsheet ID
        RANGE_NAME = f'{sheet_name}!A1:Z1000'  # Update the range as needed, here assuming max 1000 rows and 26 columns

        body = {}
        try:
            result = sheet.values().clear(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME, body=body).execute()
            print(f"Cleared the Google Sheet {sheet_name}: {result}")
        except Exception as e:
            print(f"An error occurred while clearing the sheet {sheet_name}: {e}")

def append_to_google_sheet(player_name, odds, headers, player_row, sheet, sheet_name):
    SPREADSHEET_ID = '1sNwMydSoqasypK-jCEa5oCqsDNuHIgyAccsUjphczY0'  # Your actual spreadsheet ID
    RANGE_NAME = f'{sheet_name}!A{player_row}:K{player_row}'  # Append data to the correct row

    data = [[player_name] + odds[:len(headers) - 1]]  # Ensure the odds match the header length

    body = {
        'values': data
    }

    print(f"Updating the Google Sheets {sheet_name} with the following data: {data}")

    try:
        result = sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME,
            valueInputOption='RAW',
            body=body
        ).execute()
        print(f'{result.get("updatedCells")} cells updated in {sheet_name}.')
    except Exception as e:
        print(f"An error occurred: {e}")

def name_extraction(image_path):
    def resize_and_enhance_image(image, factor):
        image = image.resize((int(image.width * factor), int(image.height * factor)), Image.LANCZOS)
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2)  # Increase contrast
        enhancer = ImageEnhance.Sharpness(image)
        image = enhancer.enhance(2)  # Increase sharpness
        return image

    def apply_image_filters(image):
        image = image.filter(ImageFilter.SHARPEN)
        image = image.filter(ImageFilter.EDGE_ENHANCE_MORE)
        return image

    def extract_text_from_image(image, max_attempts=10):
        factor = 2.0  # Start with 2 times the original size

        for attempt in range(max_attempts):
            enhanced_image = resize_and_enhance_image(image, factor)
            filtered_image = apply_image_filters(enhanced_image)
            text = pytesseract.image_to_string(filtered_image, config='--psm 6')  # Use page segmentation mode 6 for better table OCR
            text = text.strip()
            if text and re.search(r'[A-Za-z]', text):  # Check if there is any letter in the text
                # Filter out lines that contain numbers
                lines = text.split('\n')
                filtered_lines = [line for line in lines if not re.search(r'\d', line)]
                filtered_text = '\n'.join(filtered_lines)
                if filtered_text.strip() and re.match(r'^[A-Za-z]+\s[A-Za-z]+$', filtered_text):  # Check if it matches "Name Surname"
                    return filtered_text
            factor += 0.5  # Increase the size factor for the next attempt

        return None

    def google_search(query):
        search_url = f"https://www.google.com/search?q={query} tennis"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        response = requests.get(search_url, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')
        first_result = soup.find('h3')
        if first_result:
            # Extract the name from the search result
            name_match = re.match(r'^([A-Za-z]+\s[A-Za-z]+)', first_result.text)
            if name_match:
                return name_match.group(1)
        return None

    def find_name_in_image(image_path):
        image = Image.open(image_path)
        width, height = image.size

        for crop_factor in [0.25, 0.20, 0.15, 0.10]:  # Start with top 25% and decrease
            top_image = image.crop((0, 0, width, int(height * crop_factor)))  # Focus on the top crop_factor% of the image
            name = extract_text_from_image(top_image)
            if name:
                print(f"Found text in {os.path.basename(image_path)} at crop percentage {crop_factor * 100}%: {name}")
                return name
        print(f"No valid text found in {os.path.basename(image_path)}")
        return None

    text = find_name_in_image(image_path)
    if text:
        google_result = google_search(text)
        if google_result and google_result != text:
            print(f"Name '{text}' corrected to '{google_result}' from web search.")
            return google_result
        else:
            return text
    else:
        return "No valid text found"

def main():
    credentials_path = r'C:/Users/johna/Desktop/upwork jobs/tennis scrapping/credentials.json'  # Update with your path
    creds = service_account.Credentials.from_service_account_file(credentials_path, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    print("Credentials loaded successfully.")
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    print("Google Sheets service built successfully.")

    clear_google_sheet(sheet)  # Clear the sheet before starting

    # Process Ace Totals files first
    extract_text_from_images(folder_path, sheet, 'at')

    # Process Double Fouls files next
    extract_text_from_images(folder_path, sheet, 'df')

if __name__ == "__main__":
    main()
    print("done")
