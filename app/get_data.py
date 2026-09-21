import os
import requests
import time
import pandas as pd
import io
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

try:
    import win32com.client
except ImportError:
    # pywin32 is Windows-only. The server imports this module (applesauce
    # pulls in the downloaders at the top) but never runs a download, so on a
    # Linux host this stays None and nothing touches it.
    win32com = None
    
import random
from selenium.webdriver.common.action_chains import ActionChains

def _unblock_excel_file(filepath):
    """
    Programmatically opens and re-saves an Excel file, suppressing the
    'overwrite' prompt to fully automate the process.
    """
    excel = None  # Initialize to ensure it's defined for the finally block
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Run in the background

        # --- KEY CHANGE: Disable Excel's pop-up alerts ---
        excel.DisplayAlerts = False

        abs_path = os.path.abspath(filepath)
        workbook = excel.Workbooks.Open(abs_path)

        # Re-saving the file removes "Protected View" and overwrites without prompting
        workbook.SaveAs(abs_path, FileFormat=51) # 51 = .xlsx format

        workbook.Close(SaveChanges=False) # Close without saving again
        print(f"Successfully unblocked and prepared '{os.path.basename(filepath)}'.")
        return True
    except Exception as e:
        print(f"!!! WARNING: Could not unblock the Excel file '{os.path.basename(filepath)}'.")
        print(f"!!! You may need to open it and click 'Enable Editing' manually. Error: {e}")
        return False
    finally:
        # --- CRITICAL: Always clean up the Excel process ---
        if excel:
            excel.DisplayAlerts = True  # Restore alerts for other programs
            excel.Quit()

def _find_header_row(csv_filepath, anchor_text="Part Number"):
    """
    Finds the index of the header row in a CSV by searching for anchor text.
    """
    try:
        # --- CHANGE START: Use utf-8-sig to correctly handle the BOM ---
        with open(csv_filepath, 'r', encoding='utf-8-sig') as f:
        # --- CHANGE END ---
            for i, line in enumerate(f):
                if anchor_text in line:
                    print(f"Header row found at line {i}.")
                    return i
    except Exception as e:
        print(f"Warning: Could not read CSV to find header. Defaulting to 0. Error: {e}")
    return 0 # Default to the first row if not found

def _unblock_csv_file(filepath):
    """
    Programmatically opens and re-saves a CSV file using COM to remove
    'Protected View' and ensure it's readable by pandas.
    """
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        abs_path = os.path.abspath(filepath)
        workbook = excel.Workbooks.Open(abs_path)

        # KEY CHANGE: FileFormat=6 is for CSV, preserving the file type.
        workbook.SaveAs(abs_path, FileFormat=6) 
        
        workbook.Close(SaveChanges=False)
        print(f"Successfully unblocked and prepared CSV: '{os.path.basename(filepath)}'.")
        return True
    except Exception as e:
        print(f"!!! WARNING: Could not unblock the CSV file '{os.path.basename(filepath)}'.")
        print(f"!!! You may need to open it and click 'Enable Editing' manually. Error: {e}")
        return False
    finally:
        if excel:
            excel.DisplayAlerts = True
            excel.Quit()

def _convert_and_unblock_problematic_xls(problem_filepath, target_xlsx_filepath):
    """
    Uses COM to force Excel to open a problematic .xls file (like SpreadsheetML)
    and save it as a proper, clean .xlsx file, bypassing all pop-ups.
    """
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        # --- This is the key: It suppresses the "format mismatch" pop-up ---
        excel.DisplayAlerts = False

        abs_problem_path = os.path.abspath(problem_filepath)
        abs_target_path = os.path.abspath(target_xlsx_filepath)

        print(f"Opening problematic file: {os.path.basename(problem_filepath)}")
        workbook = excel.Workbooks.Open(abs_problem_path)

        # FileFormat=51 is for .xlsx, ensuring a modern, clean format.
        print(f"Re-saving as a valid XLSX file: {os.path.basename(target_xlsx_filepath)}")
        workbook.SaveAs(abs_target_path, FileFormat=51)
        
        workbook.Close(SaveChanges=False)
        print("Conversion successful.")
        return True
    except Exception as e:
        print(f"!!! CRITICAL: Failed during Excel COM conversion. Error: {e}")
        return False
    finally:
        if excel:
            # Restore alerts for other applications and quit Excel.
            excel.DisplayAlerts = True
            excel.Quit()


def download_and_rename_ti_specs():
    """
    Automates the TI spec download and then unblocks the file
    to make it readable by pandas.
    """
    driver = None
    try:
        project_path = os.getcwd()
        new_filename = "ti_specs.xlsx"
        new_filepath = os.path.join(project_path, new_filename)
        print(f"Project path & download location: {project_path}")

        if os.path.exists(new_filepath):
            print(f"An old version of '{new_filename}' was found. Deleting it.")
            os.remove(new_filepath)

        chrome_options = webdriver.ChromeOptions()
        prefs = {"download.default_directory": project_path}
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--start-maximized")

        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--log-level=3") # Suppress console noise

        print("Initializing HEADLESS WebDriver for TI...")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("WebDriver initialized successfully in the background.")
        url = "https://www.ti.com/product-category/passive-discrete/diodes/products.html"
        print(f"Navigating to the correct URL: {url}...")
        driver.get(url)

        try:
            print("Looking for TI cookie banner...")
            cookie_accept_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.ID, "consent_prompt_submit"))
            )
            print("Cookie banner found. Clicking 'Accept All'.")
            cookie_accept_button.click()
            time.sleep(2)
        except TimeoutException:
            print("Cookie banner not found or already accepted. Continuing...")

        files_before = os.listdir(project_path)
        print("Waiting for page components to become stable...")
        time.sleep(8)

        js_command = """
        const downloadButton = document.querySelector('ti-button.ti-selection-tool-action-bar-download');
        if (downloadButton) {
            downloadButton.click();
        } else {
            throw new Error('Could not find the download <ti-button> component.');
        }
        """
        
        try:
            print("Executing targeted click on the <ti-button> component...")
            driver.execute_script(js_command)
            print("JavaScript command executed successfully.")
        except Exception as e:
            print(f"The JavaScript execution failed. Error: {e}")
            raise

        print("Waiting for file to appear on disk...")
        timeout_seconds = 90
        start_time = time.time()
        downloaded_file_path = None
        
        while time.time() - start_time < timeout_seconds:
            files_after = os.listdir(project_path)
            new_files = [f for f in files_after if f not in files_before]
            
            for filename in new_files:
                if filename.lower().endswith('.xlsx') and not filename.lower().endswith('.crdownload'):
                    original_filepath = os.path.join(project_path, filename)
                    try:
                        last_size = os.path.getsize(original_filepath)
                        time.sleep(2)
                        current_size = os.path.getsize(original_filepath)
                        if last_size == current_size and current_size > 0:
                            print(f"\nDownload complete. Original name: {filename}")
                            print(f"Renaming to '{new_filename}'...")
                            os.rename(original_filepath, new_filepath)
                            print("File successfully renamed.")
                            downloaded_file_path = new_filepath
                            break
                    except (OSError, FileNotFoundError):
                        continue
            
            if downloaded_file_path:
                break
            time.sleep(0.5)

        if not downloaded_file_path:
            raise Exception(f"Download did not complete within {timeout_seconds} seconds.")

        _unblock_excel_file(downloaded_file_path)
        
        return downloaded_file_path

    except Exception as e:
        print(f"\nAn error occurred during the TI spec download: {e}")
        return None

    finally:
        if driver:
            print("\nClosing the browser.")
            driver.quit()

def download_and_rename_ti_zener_specs():
    """
    Automates the TI spec download and then unblocks the file
    to make it readable by pandas.
    """
    driver = None
    try:
        project_path = os.getcwd()
        new_filename = "ti_zener_specs.xlsx"
        new_filepath = os.path.join(project_path, new_filename)
        print(f"Project path & download location: {project_path}")

        if os.path.exists(new_filepath):
            print(f"An old version of '{new_filename}' was found. Deleting it.")
            os.remove(new_filepath)

        chrome_options = webdriver.ChromeOptions()
        prefs = {"download.default_directory": project_path}
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--start-maximized")

        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--log-level=3") # Suppress console noise

        print("Initializing HEADLESS WebDriver for TI...")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("WebDriver initialized successfully in the background.")
        url = "https://www.ti.com/product-category/passive-discrete/diodes/zener-diodes/products.html"
        print(f"Navigating to the correct URL: {url}...")
        driver.get(url)

        try:
            print("Looking for TI cookie banner...")
            cookie_accept_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.ID, "consent_prompt_submit"))
            )
            print("Cookie banner found. Clicking 'Accept All'.")
            cookie_accept_button.click()
            time.sleep(2)
        except TimeoutException:
            print("Cookie banner not found or already accepted. Continuing...")

        files_before = os.listdir(project_path)
        print("Waiting for page components to become stable...")
        time.sleep(8)

        js_command = """
        const downloadButton = document.querySelector('ti-button.ti-selection-tool-action-bar-download');
        if (downloadButton) {
            downloadButton.click();
        } else {
            throw new Error('Could not find the download <ti-button> component.');
        }
        """
        
        try:
            print("Executing targeted click on the <ti-button> component...")
            driver.execute_script(js_command)
            print("JavaScript command executed successfully.")
        except Exception as e:
            print(f"The JavaScript execution failed. Error: {e}")
            raise

        print("Waiting for file to appear on disk...")
        timeout_seconds = 90
        start_time = time.time()
        downloaded_file_path = None
        
        while time.time() - start_time < timeout_seconds:
            files_after = os.listdir(project_path)
            new_files = [f for f in files_after if f not in files_before]
            
            for filename in new_files:
                if filename.lower().endswith('.xlsx') and not filename.lower().endswith('.crdownload'):
                    original_filepath = os.path.join(project_path, filename)
                    try:
                        last_size = os.path.getsize(original_filepath)
                        time.sleep(2)
                        current_size = os.path.getsize(original_filepath)
                        if last_size == current_size and current_size > 0:
                            print(f"\nDownload complete. Original name: {filename}")
                            print(f"Renaming to '{new_filename}'...")
                            os.rename(original_filepath, new_filepath)
                            print("File successfully renamed.")
                            downloaded_file_path = new_filepath
                            break
                    except (OSError, FileNotFoundError):
                        continue
            
            if downloaded_file_path:
                break
            time.sleep(0.5)

        if not downloaded_file_path:
            raise Exception(f"Download did not complete within {timeout_seconds} seconds.")

        _unblock_excel_file(downloaded_file_path)
        
        return downloaded_file_path

    except Exception as e:
        print(f"\nAn error occurred during the TI spec download: {e}")
        return None

    finally:
        if driver:
            print("\nClosing the browser.")
            driver.quit()

def download_and_rename_aos_specs():
    """
    Final, production-ready version.
    1. Uses webdriver-manager to automatically handle the driver.
    2. Handles the cookie banner.
    3. Performs a precise scroll to make the button visible.
    4. Uses a robust XPath to locate and click the correct <li> element.
    5. Waits for the download to complete.
    6. Renames the downloaded file to 'aos_specs.xlsx'.
    """
    driver = None
    try:
        # --- 1. Set up Chrome options ---
        project_path = os.getcwd()
        new_filename = "aos_specs.xlsx"
        new_filepath = os.path.join(project_path, new_filename)
        print(f"Project path & download location: {project_path}")

        # --- Clean up any old version of the file first ---
        if os.path.exists(new_filepath):
            print(f"An old version of '{new_filename}' was found. Deleting it.")
            os.remove(new_filepath)

        chrome_options = Options()
        prefs = {"download.default_directory": project_path}
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--log-level=3")

        print("Initializing HEADLESS WebDriver for AOS...")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("WebDriver initialized successfully in the background.")
        # --- 3. Navigate and handle pop-ups ---
        url = "https://www.aosmd.com/products/tvs"
        print(f"Navigating to {url}...")
        driver.get(url)

        try:
            print("Looking for cookie banner...")
            cookie_accept_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "hs-eu-confirmation-button"))
            )
            print("Cookie banner found. Clicking 'Accept All'.")
            cookie_accept_button.click()
            time.sleep(1)
        except TimeoutException:
            print("Cookie banner not found or already accepted. Continuing...")
        
        # --- 4. Precise Scroll ---
        try:
            print("Locating the 'PARAMETRICS' tab to scroll down...")
            parametrics_tab = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.LINK_TEXT, "PARAMETRICS"))
            )
            print("Scrolling to the 'PARAMETRICS' tab...")
            driver.execute_script("arguments[0].scrollIntoView();", parametrics_tab)
            time.sleep(2)
        except TimeoutException:
            print("Could not find the 'PARAMETRICS' tab to scroll to.")
        
        # --- 5. Find and click the download button ---
        files_before = os.listdir(project_path)
        print("Waiting for the download button...")
        xpath_for_download_button = "//li[contains(text(), 'Download results as Excel')]"
        download_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, xpath_for_download_button))
        )
        print("Button found. Clicking...")
        driver.execute_script("arguments[0].click();", download_button)

        # --- 6. Wait for download and perform rename ---
        print("Waiting for download to complete...")
        timeout_seconds = 60
        start_time = time.time()
        
        while time.time() - start_time < timeout_seconds:
            files_after = os.listdir(project_path)
            new_files = [f for f in files_after if f not in files_before]
            
            for filename in new_files:
                if filename.endswith('.xlsx') and not filename.endswith('.crdownload'):
                    original_filepath = os.path.join(project_path, filename)
                    try:
                        # Final check that download is stable
                        last_size = os.path.getsize(original_filepath)
                        time.sleep(1.5)
                        current_size = os.path.getsize(original_filepath)
                        if last_size == current_size and current_size > 0:
                            # --- RENAME THE FILE ---
                            print(f"\nDownload complete. Original name: {filename}")
                            print(f"Renaming to '{new_filename}'...")
                            os.rename(original_filepath, new_filepath)
                            print("File successfully renamed.")
                            return new_filepath # Return the new, predictable path
                    except (OSError, FileNotFoundError):
                        continue

            time.sleep(0.5)

        raise Exception(f"Download did not complete within {timeout_seconds} seconds.")

    except Exception as e:
        print(f"\nAn error occurred: {e}")
        return None

    finally:
        if driver:
            print("\nClosing the browser.")
            driver.quit()

def download_diodes_specs(files_to_download):
    """
    Downloads a dictionary of files directly via their URLs, with a longer timeout.
    """
    print("\n--- Starting Diodes Inc. Spec File Download ---")
    try:
        for filename, url in files_to_download.items():
            print(f"Downloading '{filename}'...")
            
            # Use requests with an increased timeout of 90 seconds
            response = requests.get(url, timeout=90)
            response.raise_for_status()
            
            filepath = os.path.join(os.getcwd(), filename)
            with open(filepath, 'wb') as f:
                f.write(response.content)
            
            print(f"Downloaded '{filename}' successfully.")
            # Unblock the file so pandas can read it
            _unblock_excel_file(filepath)

        print("--- Diodes Inc. Spec File Download Complete ---")
        return True
    except requests.exceptions.RequestException as e:
        print(f"!!! CRITICAL: A network error occurred while downloading Diodes Inc. files. Error: {e}")
        return False
    except Exception as e:
        print(f"!!! CRITICAL: An unexpected error occurred. Error: {e}")
        return False
    
def download_nexperia_specs(files_to_download):
    """
    Automates the download of Nexperia spec files by iterating through a dictionary
    of filenames and their corresponding URLs.
    """
    driver = None
    try:
        project_path = os.getcwd()
        chrome_options = Options()
        prefs = {"download.default_directory": project_path}
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--log-level=3")

        print("Initializing HEADLESS WebDriver for Nexperia batch download...")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("WebDriver initialized successfully.")

        for new_filename, url in files_to_download.items():
            new_filepath = os.path.join(project_path, new_filename)
            if os.path.exists(new_filepath):
                print(f"An old version of '{new_filename}' was found. Deleting it.")
                os.remove(new_filepath)

            print(f"\nNavigating to {url} for '{new_filename}'...")
            driver.get(url)

            try:
                cookie_accept_xpath = "//a[contains(text(), 'Allow all cookies')]"
                cookie_accept_button = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, cookie_accept_xpath))
                )
                driver.execute_script("arguments[0].click();", cookie_accept_button)
                time.sleep(2)
            except TimeoutException:
                print("Cookie banner not found or already accepted.")

            files_before = os.listdir(project_path)
            print("Waiting for the 'Download' button...")
            download_button_selector = "button[data-event='download-excel']"
            download_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, download_button_selector))
            )
            driver.execute_script("arguments[0].click();", download_button)

            print("Waiting for download to complete...")
            timeout_seconds = 60
            start_time = time.time()
            download_successful = False
            
            while time.time() - start_time < timeout_seconds:
                files_after = os.listdir(project_path)
                new_files = [f for f in files_after if f not in files_before]
                
                for filename in new_files:
                    if filename.endswith('.xls') and not filename.endswith('.crdownload'):
                        original_filepath = os.path.join(project_path, filename)
                        try:
                            last_size = os.path.getsize(original_filepath)
                            time.sleep(1.5)
                            current_size = os.path.getsize(original_filepath)
                            if last_size == current_size and current_size > 0:
                                print(f"Download complete. Renaming to '{new_filename}'...")
                                os.rename(original_filepath, new_filepath)
                                download_successful = True
                                break
                        except (OSError, FileNotFoundError):
                            continue
                if download_successful:
                    break
                time.sleep(0.5)

            if not download_successful:
                print(f"!!! WARNING: Download did not complete for '{new_filename}'.")

    except Exception as e:
        print(f"\nAn error occurred during the Nexperia spec download: {e}")
    finally:
        if driver:
            print("\nClosing the browser.")
            driver.quit()

def download_littelfuse_specs(files_to_download):
    """
    Automates the download of Littelfuse spec files by iterating through a dictionary
    of filenames and their corresponding URLs.
    """
    driver = None
    try:
        project_path = os.getcwd()
        chrome_options = webdriver.ChromeOptions()
        prefs = {"download.default_directory": project_path}
        chrome_options.add_experimental_option("prefs", prefs)
        
        # --- START OF FIX: Make Headless Mode More Robust ---
        # 1. Enable Headless Mode
        chrome_options.add_argument("--headless")
        
        # 2. Set a standard User-Agent. This is crucial for preventing sites
        #    from identifying the browser as a bot.
        user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
        chrome_options.add_argument(f'user-agent={user_agent}')

        # 3. Set a common window size. This prevents the site from rendering a
        #    mobile version where the download button might be hidden.
        chrome_options.add_argument("--window-size=1920,1080")
        # --- END OF FIX ---
        
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--log-level=3")

        print("Initializing HEADLESS WebDriver for Littelfuse batch download...")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("WebDriver initialized successfully.")

        for new_filename, url in files_to_download.items():
            csv_filename = new_filename
            xlsx_filename = new_filename.replace('.csv', '.xlsx')
            xlsx_filepath = os.path.join(project_path, xlsx_filename)

            if os.path.exists(xlsx_filepath):
                print(f"An old version of '{xlsx_filename}' was found. Deleting it.")
                os.remove(xlsx_filepath)

            print(f"\nNavigating to {url} for '{csv_filename}'...")
            driver.get(url)

            try:
                # The cookie logic is fine, but we give it a slightly shorter wait
                # as it might not even appear in this more robust headless mode.
                accept_button_xpath = "//button[@id='onetrust-accept-btn-handler']"
                cookie_accept_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, accept_button_xpath))
                )
                driver.execute_script("arguments[0].click();", cookie_accept_button)
                time.sleep(2)
            except TimeoutException:
                print("Cookie banner not found or already accepted.")

            files_before = os.listdir(project_path)
            print("Waiting for the CSV download button...")
            download_button_selector = "button.buttons-csv"
            download_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, download_button_selector))
            )
            driver.execute_script("arguments[0].click();", download_button)

            print("Waiting for download to complete...")
            timeout_seconds = 60
            start_time = time.time()
            download_successful = False
            
            while time.time() - start_time < timeout_seconds:
                files_after = os.listdir(project_path)
                new_files = [f for f in files_after if f not in files_before]
                
                for filename in new_files:
                    if filename.lower().endswith('.csv') and not filename.lower().endswith('.crdownload'):
                        original_filepath = os.path.join(project_path, filename)
                        try:
                            last_size = os.path.getsize(original_filepath)
                            time.sleep(1.5)
                            current_size = os.path.getsize(original_filepath)
                            if last_size == current_size and current_size > 0:
                                print(f"Download complete. Manually cleaning and converting '{filename}'...")

                                with open(original_filepath, 'r', encoding='utf-8-sig') as f:
                                    all_lines = f.readlines()

                                header_row_index = 0
                                for i, line in enumerate(all_lines):
                                    if "Part Number" in line:
                                        header_row_index = i
                                        break
                                
                                clean_csv_lines = all_lines[header_row_index:]
                                clean_csv_string = "".join(clean_csv_lines)
                                
                                # By adding index_col=False, we force pandas to treat the first column
                                # as data, not as the DataFrame's index, which prevents the data shift.
                                temp_df = pd.read_csv(io.StringIO(clean_csv_string), index_col=False)
                                
                                temp_df.to_excel(xlsx_filepath, index=False)
                                os.remove(original_filepath)
                                
                                print("Conversion successful. Unblocking the new Excel file...")
                                _unblock_excel_file(xlsx_filepath)
                                download_successful = True
                                break
                        except (OSError, FileNotFoundError):
                            continue
                
                if download_successful:
                    break
                time.sleep(0.5)

            if not download_successful:
                print(f"!!! WARNING: Download did not complete for '{csv_filename}'.")

    except Exception as e:
        print(f"\nAn error occurred during the Littelfuse spec download: {e}")
    finally:
        if driver:
            print("\nClosing the browser.")
            driver.quit()

def download_jiangsu_specs(files_to_download):
    """
    Automates the download of Jiangsu spec files. This version uses COM to
    force Excel to convert the problematic downloaded file into a valid format.
    """
    driver = None
    try:
        project_path = os.getcwd()
        chrome_options = webdriver.ChromeOptions()
        prefs = {"download.default_directory": project_path}
        chrome_options.add_experimental_option("prefs", prefs)
        
        chrome_options.add_argument("--headless")
        user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
        chrome_options.add_argument(f'user-agent={user_agent}')
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--log-level=3")

        print("Initializing HEADLESS WebDriver for Jiangsu batch download...")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("WebDriver initialized successfully.")

        for new_filename, url in files_to_download.items():
            # --- We now create .xlsx files, so adjust the target path ---
            new_filepath_xlsx = new_filename.replace('.xls', '.xlsx')
            if os.path.exists(new_filepath_xlsx):
                print(f"An old version of '{new_filepath_xlsx}' was found. Deleting it.")
                os.remove(new_filepath_xlsx)

            print(f"\nNavigating to {url} for '{new_filename}'...")
            driver.get(url)

            files_before = os.listdir(project_path)
            print("Waiting for the 'Export' button (导出)...")
            
            export_button_xpath = "//a[contains(text(), '导出')]"
            export_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, export_button_xpath))
            )
            
            driver.execute_script("arguments[0].click();", export_button)
            print("Export button clicked. Waiting for download...")

            timeout_seconds = 90
            start_time = time.time()
            download_successful = False
            
            while time.time() - start_time < timeout_seconds:
                files_after = os.listdir(project_path)
                new_files = [f for f in files_after if f not in files_before]
                
                for filename in new_files:
                    if filename.lower().endswith('.xls') and not filename.lower().endswith('.crdownload'):
                        original_filepath = os.path.join(project_path, filename)
                        try:
                            last_size = os.path.getsize(original_filepath)
                            time.sleep(2)
                            current_size = os.path.getsize(original_filepath)
                            if last_size == current_size and current_size > 0:
                                print(f"Download complete. Converting '{filename}' via Excel...")
                                
                                # --- Use the new conversion function ---
                                conversion_ok = _convert_and_unblock_problematic_xls(original_filepath, new_filepath_xlsx)
                                
                                # Clean up the original downloaded file regardless of success
                                os.remove(original_filepath)
                                
                                if not conversion_ok:
                                    raise Exception("Excel COM conversion failed.")

                                download_successful = True
                                break
                        except Exception as e:
                             print(f"!!! CRITICAL: Failed during file processing for {filename}. Error: {e}")
                             if os.path.exists(original_filepath):
                                 os.remove(original_filepath)
                             continue
                if download_successful:
                    break
                time.sleep(0.5)

            if not download_successful:
                print(f"!!! WARNING: Download did not complete for '{new_filename}'.")

    except Exception as e:
        print(f"\nAn error occurred during the Jiangsu spec download: {e}")
    finally:
        if driver:
            print("\nClosing the browser.")
            driver.quit()      

def download_vishay_specs(files_to_download):
    """
    Automates downloading spec files from Vishay by iterating through a dictionary.
    1. Uses a headless browser window.
    2. Handles multiple URLs and filenames.
    3. Locates the export button using a partial ID match.
    4. Waits for the download to complete, renames the file, and unblocks it.
    """
    driver = None
    try:
        project_path = os.getcwd()
        chrome_options = webdriver.ChromeOptions()
        prefs = {"download.default_directory": project_path}
        chrome_options.add_experimental_option("prefs", prefs)
        
        # --- KEY CHANGES FOR HEADLESS OPERATION ---
        chrome_options.add_argument("--headless")
        user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
        chrome_options.add_argument(f'user-agent={user_agent}')
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--log-level=3")

        print("Initializing HEADLESS WebDriver for Vishay batch download...")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("WebDriver initialized successfully.")

        for new_filename, url in files_to_download.items():
            new_filepath = os.path.join(project_path, new_filename)
            if os.path.exists(new_filepath):
                print(f"An old version of '{new_filename}' was found. Deleting it.")
                os.remove(new_filepath)

            print(f"\nNavigating to {url} for '{new_filename}'...")
            driver.get(url)

            files_before = os.listdir(project_path)
            print("Waiting for the 'Export as MS Excel' button...")
            
            export_button_selector = "button[id*='TabletoExcelButton']"
            export_button = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, export_button_selector))
            )
            
            print("Button found. Clicking to download Excel file...")
            export_button.click()

            print("Waiting for file to appear on disk...")
            timeout_seconds = 90
            start_time = time.time()
            download_successful = False
            
            while time.time() - start_time < timeout_seconds:
                files_after = os.listdir(project_path)
                new_files = [f for f in files_after if f not in files_before]
                
                for filename in new_files:
                    if filename.lower().endswith('.xlsx') and not filename.lower().endswith('.crdownload'):
                        original_filepath = os.path.join(project_path, filename)
                        try:
                            last_size = os.path.getsize(original_filepath)
                            time.sleep(2)
                            current_size = os.path.getsize(original_filepath)
                            if last_size == current_size and current_size > 0:
                                print(f"\nDownload complete. Renaming to '{new_filename}'...")
                                os.rename(original_filepath, new_filepath)
                                print("File successfully renamed. Unblocking...")
                                _unblock_excel_file(new_filepath)
                                download_successful = True
                                break
                        except (OSError, FileNotFoundError):
                            continue
                
                if download_successful:
                    break
                time.sleep(0.5)

            if not download_successful:
                print(f"!!! WARNING: Download did not complete for '{new_filename}'.")

    except Exception as e:
        print(f"\nAn error occurred during the Vishay spec download: {e}")
        return None

    finally:
        if driver:
            print("\nClosing the browser.")
            driver.quit()

def download_semtech_specs():
    """
    Automates downloading the spec file from Semtech's circuit protection page.
    This is a single-purpose function for one URL and is modeled on the robust
    Littelfuse/CSV downloader.
    """
    driver = None
    try:
        project_path = os.getcwd()
        # The final file will be .xlsx
        new_filename = "semtech_specs.xlsx"
        new_filepath = os.path.join(project_path, new_filename)

        if os.path.exists(new_filepath):
            print(f"An old version of '{new_filename}' was found. Deleting it.")
            os.remove(new_filepath)
            
        chrome_options = webdriver.ChromeOptions()
        prefs = {"download.default_directory": project_path}
        chrome_options.add_experimental_option("prefs", prefs)
        
        chrome_options.add_argument("--headless")
        user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
        chrome_options.add_argument(f'user-agent={user_agent}')
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--log-level=3")

        print("Initializing HEADLESS WebDriver for Semtech...")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("WebDriver initialized successfully.")

        url = "https://www.semtech.com/products/circuit-protection#parametric-search"
        print(f"\nNavigating to {url} for '{new_filename}'...")
        driver.get(url)

        try:
            accept_button_xpath = "//button[normalize-space()='Accept All']"
            cookie_accept_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, accept_button_xpath))
            )
            driver.execute_script("arguments[0].click();", cookie_accept_button)
            time.sleep(2)
        except TimeoutException:
            print("Cookie banner not found or already accepted.")

        files_before = os.listdir(project_path)
        print("Waiting for the 'Export Result' (CSV) button...")
        
        download_button_selector = "button.buttons-csv"
        download_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, download_button_selector))
        )
        driver.execute_script("arguments[0].click();", download_button)

        print("Waiting for download to complete...")
        timeout_seconds = 60
        start_time = time.time()
        downloaded_csv_path = None
        
        while time.time() - start_time < timeout_seconds:
            files_after = os.listdir(project_path)
            new_files = [f for f in files_after if f not in files_before]
            
            for filename in new_files:
                if filename.lower().endswith('.csv') and not filename.lower().endswith('.crdownload'):
                    original_filepath = os.path.join(project_path, filename)
                    try:
                        last_size = os.path.getsize(original_filepath)
                        time.sleep(1.5)
                        current_size = os.path.getsize(original_filepath)
                        if last_size == current_size and current_size > 0:
                            print(f"Download of '{filename}' complete.")
                            downloaded_csv_path = original_filepath
                            break
                    except (OSError, FileNotFoundError):
                        continue
            if downloaded_csv_path:
                break
            time.sleep(0.5)

        if not downloaded_csv_path:
            raise Exception(f"CSV download did not complete within {timeout_seconds} seconds.")

        # Convert the downloaded CSV to a clean XLSX file
        print(f"Converting '{os.path.basename(downloaded_csv_path)}' to '{new_filename}'...")
        temp_df = pd.read_csv(downloaded_csv_path, on_bad_lines='skip')
        temp_df.to_excel(new_filepath, index=False)
        os.remove(downloaded_csv_path) # Clean up the temporary CSV
        
        print(f"Conversion successful. Unblocking '{new_filename}'...")
        _unblock_excel_file(new_filepath)
        return new_filepath

    except Exception as e:
        print(f"\nAn error occurred during the Semtech spec download: {e}")
        return None
    finally:
        if driver:
            print("\nClosing the browser.")
            driver.quit()

# --- File: get_data.py ---

