import subprocess
import time
import pyautogui
import pyperclip
from openpyxl import load_workbook
from PIL import ImageGrab, ImageOps
import pytesseract
import re
import os

# Global configuration
EXE_PATH = r"C:\InfoVotantes\InfoVotantes.exe"
EXCEL_FILE_PATH = os.path.join(os.path.dirname(__file__), "InfoVotantes.xlsx")
CEDULA = "1117512408"

# Region coordinates for the result area (x, y, width, height)
RESULT_REGION = (932, 334, 305, 314)
# Extra pixels around the result area to avoid clipping
REGION_PADDING = 0

# Disable pyautogui fail-safe
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5


def run_exe():
    """
    Launches the InfoVotantes.exe application.
    Returns the subprocess object for later interaction if needed.
    """
    try:
        print(f"Launching application...")
        process = subprocess.Popen(EXE_PATH)
        time.sleep(2)
        return process
    except FileNotFoundError:
        print(f"Error: Executable not found at {EXE_PATH}")
        return None
    except Exception as e:
        print(f"Error launching application: {e}")
        return None


def extract_text_from_screen(x=0, y=0, width=None, height=None):
    """
    Takes a screenshot of the screen and extracts text using OCR.
    If width/height not specified, captures the entire screen.
    """
    try:
        if width is None or height is None:
            # Capture full screen
            screenshot = ImageGrab.grab()
        else:
            # Capture specific region
            screenshot = ImageGrab.grab(bbox=(x, y, x + width, y + height))
        # Improve contrast for more stable OCR across screens
        preprocessed = ImageOps.grayscale(screenshot)
        preprocessed = ImageOps.autocontrast(preprocessed)

        # Use pytesseract to extract text from the screenshot
        text = pytesseract.image_to_string(preprocessed)
        return text.strip()
    except Exception as e:
        print(f"Error extracting text from screen: {e}")
        return ""


def parse_voting_info(raw_text):
    """
    Parses the raw OCR text and formats it into structured voting information.
    Returns a dictionary with the parsed data.
    """
    lines = [line.strip() for line in raw_text.split('\n') if line.strip()]

    # Debug: Print lines after empty line removal
    print("\n=== LINES AFTER EMPTY LINE REMOVAL ===")
    for idx, line in enumerate(lines):
        print(f"  [{idx}] '{line}'")
    print("======================================\n")
    
    # Post-processing: Fix common OCR errors
    if len(lines) >= 3:
        # Fix Line 2 (Zona): "0" often interpreted as "i¢)"
        lines[2] = lines[2].replace('i¢)', '0').replace('ic)', '0').replace('i©)', '0')
        # Also handle other common "0" misreads
        lines[2] = lines[2].replace('O', '0')  # Capital O to zero
        # Keep only digits in Zona (e.g., "3(" or "3{" -> "3")
        zona_digits = re.findall(r"\d+", lines[2])
        if zona_digits:
            lines[2] = zona_digits[0]

    # Debug: Print lines after post-processing
    print("=== LINES AFTER POST-PROCESSING ===")
    for idx, line in enumerate(lines):
        print(f"  [{idx}] '{line}'")
    print("===================================\n")

    result = {
        'Departamento': '',
        'Municipio': '',
        'Zona': '',
        'Puesto': '',
        'Mesa': '',
        'Direccion': ''
    }

    if len(lines) < 4:
        return result

    # Line 0 -> Departamento
    result['Departamento'] = lines[0]
    
    # Line 1 -> Municipio
    result['Municipio'] = lines[1]
    
    # Line 2 -> Zona (as is)
    result['Zona'] = lines[2]
    
    # Line 3 -> Puesto (first part)
    result['Puesto'] = lines[3]

    if len(lines) < 5:
        return result

    # Line 4: Check if starts with letter or number
    if lines[4] and lines[4][0].isalpha():
        # Line 4 is continuation of Puesto (starts with letter)
        puesto_parts = [result['Puesto'], lines[4]]
        next_index = 5

        # Line 5 can also be continuation of Puesto if it starts with a letter
        if next_index < len(lines) and lines[next_index][0].isalpha():
            puesto_parts.append(lines[next_index])
            next_index += 1

        result['Puesto'] = ' '.join(puesto_parts)

        # Next line after Puesto is Mesa
        if next_index < len(lines):
            result['Mesa'] = lines[next_index]
            # Clean Mesa: keep only digits
            mesa_digits = re.findall(r"\d+", result['Mesa'])
            if mesa_digits:
                result['Mesa'] = mesa_digits[0]
            next_index += 1

            # Everything after Mesa is Direccion
            if next_index < len(lines):
                result['Direccion'] = ' '.join(lines[next_index:])
    else:
        # Line 4 is Mesa (starts with number)
        result['Mesa'] = lines[4]
        # Clean Mesa: keep only digits
        mesa_digits = re.findall(r"\d+", result['Mesa'])
        if mesa_digits:
            result['Mesa'] = mesa_digits[0]

        # Line 5 and onwards is Direccion
        if len(lines) >= 6:
            result['Direccion'] = ' '.join(lines[5:])

    return result


def capture_result_text_with_retry(region, retries=2, delay_seconds=1.5):
    """
    Captures OCR text from a region with retries to allow UI to fully render.
    Returns the best (longest) OCR text observed.
    """
    padded_region = (
        max(0, region[0] - REGION_PADDING),
        max(0, region[1] - REGION_PADDING),
        region[2] + (REGION_PADDING * 2),
        region[3] + (REGION_PADDING * 2)
    )
    best_text = ""
    for attempt in range(retries + 1):
        text = extract_text_from_screen(
            x=padded_region[0],
            y=padded_region[1],
            width=padded_region[2],
            height=padded_region[3]
        )
        if len(text) > len(best_text):
            best_text = text
        if attempt < retries:
            time.sleep(delay_seconds)
    return best_text


def format_voting_info(parsed_data):
    """
    Formats the parsed voting information for display.
    """
    output = []
    output.append(f"Departamento: {parsed_data['Departamento']}")
    output.append(f"Municipio: {parsed_data['Municipio']}")
    output.append(f"Zona: {parsed_data['Zona']}")
    output.append(f"Puesto: {parsed_data['Puesto']}")
    output.append(f"Mesa: {parsed_data['Mesa']}")
    output.append(f"Direccion: {parsed_data['Direccion']}")
    return '\n'.join(output)


def enter_cedula_and_search(cedula):
    """
    Enters the cedula into the text box and presses Enter to search.
    Then reads the results using OCR and returns to the cedula input page.
    """
    try:
        print(f"Pasting cedula: {cedula}")

        # Copy to clipboard and paste
        pyperclip.copy(cedula)
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)

        # Press Enter to search
        print("Searching...")
        pyautogui.press('enter')
        time.sleep(4)  # Wait for search results to load

        # Extract text from results using OCR
        print("Reading results...")
        result_text = capture_result_text_with_retry(RESULT_REGION, retries=2, delay_seconds=1.5)
        
        # Print raw OCR output
        print("\n=== RAW OCR OUTPUT ===")
        print(result_text)
        print("======================\n")
        
        # Parse and print formatted output
        parsed_result = parse_voting_info(result_text)
        formatted_result = format_voting_info(parsed_result)
        print("=== FORMATTED OUTPUT ===")
        print(formatted_result)
        print("========================\n")

        # Press Enter again to go back to cedula input page
        print("Returning to input page...")
        pyautogui.press('enter')
        time.sleep(2)  # Wait for page to load

        print(f"Completed cedula: {cedula}")
        return parsed_result
    except Exception as e:
        print(f"Error: {e}")
        return None


def read_cedulas_from_excel():
    """
    Reads cedulas from Column A of the Excel file starting from row 2.
    Returns a list of cedula values.
    """
    try:
        print(f"Reading cedulas from {EXCEL_FILE_PATH}")
        workbook = load_workbook(EXCEL_FILE_PATH)
        worksheet = workbook.active

        cedulas = []
        row = 2
        while True:
            cell_value = worksheet[f'A{row}'].value
            if cell_value is None:
                break
            cedulas.append(str(cell_value).strip())
            row += 1

        print(f"Found {len(cedulas)} cedulas")
        return cedulas
    except FileNotFoundError:
        print(f"Error: Excel file not found at {EXCEL_FILE_PATH}")
        return []
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []


def write_voting_data_to_excel(row_number, voting_data):
    """
    Writes parsed voting data to Excel columns B-G for the specified row.
    B: Departamento, C: Municipio, D: Zona, E: Puesto, F: Mesa, G: Direccion
    """
    try:
        workbook = load_workbook(EXCEL_FILE_PATH)
        worksheet = workbook.active
        
        worksheet[f'B{row_number}'] = voting_data['Departamento']
        worksheet[f'C{row_number}'] = voting_data['Municipio']
        worksheet[f'D{row_number}'] = voting_data['Zona']
        worksheet[f'E{row_number}'] = voting_data['Puesto']
        worksheet[f'F{row_number}'] = voting_data['Mesa']
        worksheet[f'G{row_number}'] = voting_data['Direccion']
        
        workbook.save(EXCEL_FILE_PATH)
        print(f"✓ Data written to Excel row {row_number}")
        return True
    except Exception as e:
        print(f"Error writing to Excel row {row_number}: {e}")
        return False


def main():
    """
    Main function to orchestrate the automation workflow.
    """
    print("Starting InfoVotantes automation...")

    # Step 1: Read cedulas from Excel
    cedulas = read_cedulas_from_excel()
    if not cedulas:
        print("No cedulas found. Exiting.")
        return

    # Step 2: Run the exe
    process = run_exe()

    if process:
        # Wait for application to be ready
        time.sleep(3)

        # Step 3: Process each cedula
        total_cedulas = len(cedulas)
        for index, cedula in enumerate(cedulas, 1):
            print(f"\nProcessing cedula {index} of {total_cedulas}")
            voting_data = enter_cedula_and_search(cedula)
            
            # Write data to Excel (row starts at 2, so row_number = index + 1)
            if voting_data:
                write_voting_data_to_excel(index + 1, voting_data)
            else:
                print(f"⚠ No data returned for cedula {cedula}")

            time.sleep(2)  # Wait between searches

        print("\nAll cedulas processed")
    else:
        print("Failed to launch application. Exiting.")


if __name__ == '__main__':
    main()