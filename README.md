Tennis Data OCR & Google Sheets Integration

Overview

This project automates the extraction and processing of tennis match data from images. Using Optical Character Recognition (OCR) with Tesseract, the script enhances image quality, extracts both numerical data (like odds or statistics) and textual player names, and then uploads the processed information directly into a Google Sheets spreadsheet. It was designed to handle specific data types such as Ace Totals and Double Fouls by reading images from a designated folder, processing them through multiple enhancement and filtering steps, validating the results, and finally storing them in structured tables.

Features

- Image Enhancement:

 -- Resizing & Contrast Adjustment: Increases image size and contrast to improve OCR accuracy.

 -- Filtering Techniques: Applies sharpening and edge enhancement filters for clearer text recognition.

- OCR and Text Parsing:

 -- OCR Extraction: Uses Tesseract OCR (with a focus on page segmentation mode 6 for tabular data) to extract text from images.

 -- Text Cleaning: Implements regular expressions to fix common OCR errors and remove unwanted characters.

 -- Data Validation: Checks the extracted numerical data (e.g., ensuring the odds are within acceptable ranges and maintaining consistency between consecutive values).

 - Player Name Recognition:

 -- Targeted Cropping: Focuses on the top portion of the image (where names are likely to be found) and reprocesses the cropped area if necessary.

 -- Google Verification: Uses a basic Google search to validate or correct the extracted names, ensuring accurate identification.

 - Google Sheets Integration:

 -- Automatic Updates: Clears previous data and updates the Google Sheets with new rows of extracted data.

 -- Dynamic Sheet Selection: Determines which sheet to update based on specific keywords in the image file names (e.g., using "at" for Ace Totals and "df" for Double Fouls).

Workflow and Functionality

Image Processing Pipeline

1. Enhancement:

The image is resized (starting at 2× the original dimensions) and its contrast is increased. Additional filters such as sharpening and edge enhancement are applied to boost the image clarity.

2. OCR Extraction and Cleanup:

The enhanced image is processed through Tesseract OCR to extract text.

Extracted text is then cleaned using custom regular expression routines to correct common OCR misinterpretations (e.g., removing unwanted characters and adjusting the numeric formats).

3. Data Parsing and Validation:

Extracted text is segmented into lines and parsed to retain only numerical values that fall within a specific range (e.g., above 0 and up to 26.0).

The script validates the extracted data by checking both the count and the consistency of the consecutive numbers.

Player Name Extraction

Name Detection:
The code focuses on the top section of the image (by progressively cropping from the top) to capture potential player names.

Enhancement and Verification:
Similar image enhancement techniques are applied to refine text extraction. Once a candidate name is obtained, the script optionally leverages a Google search to verify or correct the name.

Google Sheets Update Process
Setting Up Headers:
The script first uploads the headers (e.g., "Player", "1+", "3+", etc.) to the appropriate sheet based on the file type.

Appending Data Rows:
Each player's data (name and corresponding odds/numbers) is appended row by row. The project handles data validation issues by adding the data even if not all criteria are met, but it logs warnings when discrepancies arise.

Clearing Existing Data:
Before processing new images, the script clears existing data from specified ranges in the Google Sheets to ensure a fresh start.

Dependencies and Setup
Requirements
Python Libraries:

pytesseract

Pillow

google-api-python-client

google.oauth2

requests

beautifulsoup4

Standard libraries: os, re, json

External Software:

Tesseract-OCR: Make sure Tesseract is installed on your system. Update the script with the correct path to the Tesseract executable.

Google Sheets API:

A valid Google Service Account with credentials stored in a JSON file.

The appropriate permissions to update and clear the target Google Sheets.


Customization

File Type Identification:
Modify the file type detection logic if you need to handle additional categories or use different keywords.

Data Validation Adjustments:

You can alter criteria in the validate_extracted_text function to fine-tune what constitutes valid data.

Image Processing Enhancements:
Adjust the image enhancement factors (e.g., scaling factor, contrast enhancement) as required by the quality of your images.

Future Improvements
Improved OCR Accuracy: Integrate more advanced pre-processing techniques to address diverse image qualities.

Error Handling: Add logging or notification systems to better manage and alert on errors during image processing or Google Sheets updating.

Expanded Data Extraction: Consider additional metrics or more complex data extraction for richer tennis statistics analysis.

License
Include your preferred license here (e.g., MIT License).

Contributing
Contributions are welcome! Feel free to fork the repository, make improvements, and submit pull requests.

Acknowledgements
This project leverages the Tesseract OCR engine and Google Sheets API to streamline the data extraction and analysis process for tennis match statistics.
