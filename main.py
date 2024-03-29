import os
import time
from tkinter import Tk, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import multiprocessing

def select_folder():
    root = Tk()
    root.withdraw()  # Hide the main window

    folder_path = filedialog.askdirectory(title="Select Folder")

    return folder_path

def upload_and_convert_pdf(file_path):
    # Launch the Chrome browser using Selenium WebDriver
    driver = webdriver.Chrome()
    try:
        # Open ilovepdf.com in the browser if not already opened
        driver.get("https://www.ilovepdf.com/pdf_to_excel")

        # Wait for the file input element to be present
        file_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
        )

        # Upload the PDF file
        file_input.send_keys(file_path)

        # Wait for the conversion process to complete
        convert_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, "processTask"))
        )

        # Scroll the button into view
        driver.execute_script("arguments[0].scrollIntoView();", convert_button)

        # Click the convert button with retry mechanism
        for _ in range(5):
            try:
                convert_button.click()
                break
            except:
                time.sleep(2)  # Wait for 2 seconds before retrying
                continue

        # Wait for the download link to appear with extended timeout
        download_link = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, "download-all-files-button"))
        )

        # Get the download URL of the converted Excel file
        excel_file_url = download_link.get_attribute("href")

        # Download the converted Excel file
        driver.get(excel_file_url)
        print(f"PDF file '{os.path.basename(file_path)}' converted to Excel successfully.")
        time.sleep(10)  # Give some time for the file to download
    except TimeoutException:
        print("Timeout while waiting for conversion process.")
    finally:
        # Close the browser after processing the file
        driver.quit()

def main():
    # Select the folder containing PDF files
    folder_path = select_folder()

    # Get a list of all files in the folder
    files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.lower().endswith('.pdf')]

    # Number of parallel processes
    num_processes = multiprocessing.cpu_count()

    # Create a pool of processes
    pool = multiprocessing.Pool(processes=num_processes)

    try:
        # Map the conversion function to each PDF file
        pool.map(upload_and_convert_pdf, files)
    finally:
        # Close the pool of processes
        pool.close()
        pool.join()

if __name__ == "__main__":
    main()

