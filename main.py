#pip install selenium
#pip install webdriver_manager
#pip install pdfplumber

#python -m PyInstaller --onefile main.py

import time, os, openpyxl, sys, shutil, pdfplumber, re, getpass
from pathlib import Path
from datetime import datetime, timedelta
from calendar import monthrange
from tkinter import Tk, filedialog
from dateutil.parser import parse

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    StaleElementReferenceException,
    ElementClickInterceptedException,
    TimeoutException
)

base_path = os.path.dirname(os.path.abspath(__file__))

downloads_path = os.path.join(base_path, 'downloads')
sys_downloads = Path.home() / 'Downloads'

downloads_dir = os.path.join(base_path, "downloads")
excel_dir = os.path.join(base_path, "excel")

os.makedirs(downloads_dir, exist_ok=True)
os.makedirs(excel_dir, exist_ok=True)

def get_dates():
    while True:
        user_input = input("" \
        "Enter a single month (MM/YYYY) or a date range within the same month separated by - (MM/DD/YYYY - MM/DD/YYYY): "
        ).strip()
        
        if "-" in user_input:
            start_val, end_val = [vals.strip() for vals in user_input.split("-", 1)]
            try:
                start = datetime.strptime(start_val, "%m/%d/%Y")
                end = datetime.strptime(end_val, "%m/%d/%Y")
            except ValueError:
                print("Something went wrong with your input. Match the format of (MM/DD/YYYY - MM/DD/YYYY)")
                continue
            if start.month != end.month or start.year != end.year:
                print("Range must stay within one month.")
                continue
            return start, end
        else:
            try:
                m, y = [int(vals) for vals in user_input.split("/")]
                last = monthrange(y, m)[1]
                return datetime(y,m,1), datetime(y, m, last)
            except:
                print("Invalid month format. Try again.")

def prompt_excel():
    user_input = input("You will need the exel sheet and make sure it has the dates changed for the month documenting. Type 1 when ready to upload the document: ").strip()
    if user_input.startswith("1"):
        root = Tk()
        root.withdraw()

        root.attributes('-topmost', True)
        root.lift()

        path = filedialog.askopenfilename(
            title = "Choose your excel file (xlsx)",
            initialdir=str(sys_downloads),
            filetypes=[("Excel files", "*.xlsx")]
        )
        root.destroy()

        if not path:
            print("No file chosen. Quiting.")
            sys.exit(1)
        return path
    else:
        print("Try again.")
        prompt_excel()

    # try:
    #     shutil.copy(os.path.join(base_path, 'blank_template.xlsx'), os.path.join(base_path, 'excel'))
    # except Exception as e:
    #     print(f"Error: {e}")
    
    # os.rename(os.path.join(base_path, 'excel', 'blank_template.xlsx'), os.path.join(base_path, 'excel', 'MonthlySummaryReport.xlsx'))
    # return os.path.join(base_path, 'excel', 'MonthlySummaryReport.xlsx')
    
    

    

def finished_download(folder_path, timeout=60):
    seconds = 0
    while True:
        downloading = [filename for filename in os.listdir(folder_path) if filename.endswith('.crdownload')]
        if not downloading:
            break
        time.sleep(1)
        seconds += 1
        if seconds > timeout:
            raise Exception("Downloading a file took too long bruh")

def wait_for_download(folder_path, max_time=60):
    seconds = 0
    while seconds < max_time:
        pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
        if pdf_files:
            return
        time.sleep(1)
        seconds += 1
    raise TimeoutError("PDF file did not download within the timeframe! Seems like the wifi took too long to load.")

def clear_downloads(clear_folder_path,f_type="pdf"):
    for f in os.listdir(clear_folder_path):
        if f.endswith(f_type):
            os.remove(os.path.join(clear_folder_path, f))

def parse_data():
    pdf_path = os.path.join(downloads_path, os.listdir(downloads_path)[0])


    text = ""

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    
    patterns = {
        "Gross Sales": r"Gross Sales \$([\d,]+\.\d{2})",
        "Discounts":        r"- Discounts \$([\d,]+\.\d{2})",
        "Promotions":       r"- Promotions \$([\d,]+\.\d{2})",
        "Refunds":          r"- Refunds \$([\d,]+\.\d{2})",
        "Labor Cost":       r"Labor Cost: \$([\d,]+\.\d{2})",
        "Labor Hours":      r"Labor Hours: ([\d,]+\.\d{2})",
        "Guest Count":      r"Guest Count: (\d+)",
        "Delivery Markup":  r"Delivery Markup\s+\d+\s+\$([\d,]+\.\d{2})"
    }

    results = {}
    for name, pattern in patterns.items():
        m = re.search(pattern, text)
            
        if m:
            val = m.group(1).replace(",", "")
            if name == "Discounts" or name == "Promotions" or name == "Refunds":
                results[name] = -float(val) if "." in val else -int(val)
            else:
                results[name] = float(val) if "." in val else int(val)
        else:
            results[name] = None
    
    cash_over = re.search(r"Cash Over\s*\$([\d,]+\.\d{2})", text)
    if cash_over:
        results["Cash OverShort"] = float(cash_over.group(1).replace(",", ""))
    else:
        cash_short = re.search(r"Cash Short\s*\$([\d,]+\.\d{2})", text)
        if cash_short:
            val = float(cash_short.group(1).replace(",", ""))
            results["Cash OverShort"] = -val
        else:
            results["Cash OverShort"] = None
    
    return results

def write_data(data, excel_path, curr_date=None):
    wb = openpyxl.load_workbook(excel_path)

    mapping = {
        "Gross Sales": ("Sales", "B"),
        "Discounts": ("Discounts", "B"),
        "Promotions": ("Coupons", "B"),
        "Refunds": ("Refunds", "B"),
        "Labor Cost": ("Labor - $", "B"),
        "Guest Count": ("Labor - EPLH", "B"),
        "Labor Hours": ("Labor - EPLH", "C"),
        "Delivery Markup": ("Markup", "B"),
        "Cash OverShort": ("Cash OverShort", "B")
    }
    for key, value in data.items():
        value = 0 if value is None else value

        sheet_name, col = mapping[key]
        sheet = wb[sheet_name]

        start_row = 2 if sheet_name == "Labor - EPLH" else 3
        match_row = None

        for row in range(start_row, sheet.max_row + 1):
            cell_val = sheet[f"A{row}"].value
            if not cell_val:
                continue
            try:
                if isinstance(cell_val, datetime):
                    cell_date = cell_val.date()
                else:
                    cell_date = parse(str(cell_val), dayfirst=False).date()
            except:
                continue
        
            if cell_date == curr_date.date():
                match_row = row
                break

        if match_row is None:
            print("Couldn't find the row for the date")
            continue
        
        sheet[f"{col}{match_row}"] = value


    wb.save(excel_path)

def get_wait(conn_given, conn):
    if conn:
        download_speed = conn.get("downlink")
        # print(f"Wifi speed: {download_speed}mbps")

        base_speed = 1.55
        base_wait = 4.0

        wait_time = base_wait * (base_speed / download_speed)
        wait_time = max(wait_time, 1.0)
        return wait_time
    else:
        return 5


def main():
    start_date, end_date = get_dates()
    current_date = start_date
    input_excel = prompt_excel()
    out_excel   = sys_downloads / "MonthlySummaryReport.xlsx"

    shutil.copy(input_excel, out_excel)
    
    BRINK_USER = input("Brink user: ")
    BRINK_PASS = input("Brink pass: ")
    print("Starting to log in.")

    prefs = {
    "download.default_directory": downloads_path,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
    }
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.add_argument("--disable-web-security")
    chrome_options.add_argument("--headless=new")
    service = Service(executable_path=ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    conn = driver.execute_script(
        "return navigator.connection || navigator.mozConnection || navigator.webkitConnection || null;"
    )
    if conn:
        download_speed = conn.get("downlink")
        print(f"Wifi speed: {download_speed}mbps")

        conn_given = True
    else:
        print("Unable to get wifi speed. Wait time set to 5.")
        conn_given = False

    driver.get('https://admin22.parpos.com/Public/Login')

    WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.ID, "Username")))
    login_url = driver.current_url
    driver.find_element(By.ID, "Username").send_keys(BRINK_USER)
    time.sleep(1)
    driver.find_element(By.ID, "Password").send_keys(BRINK_PASS)
    time.sleep(1)
    driver.find_element(By.ID, "Password").send_keys(Keys.RETURN)

    try:
        WebDriverWait(driver, 20).until(lambda d: d.current_url != login_url)
        print("Login successful")
    except TimeoutException:
        print("Unable to login, maybe the credentials were typed wrong.")
        driver.quit()
        sys.exit(1)

    print("Program is now running. This will take some time, so go ahead and get a quadruple Butterburger while you wait.")

    while current_date <= end_date:
        clear_downloads(downloads_path)
        driver.get('https://admin22.parpos.com/Reports/Report/SalesSummary/')
        print(f"Processing {current_date.strftime('%Y-%m-%d')}...")

        date_input = driver.find_element(By.ID, "DateRangeModel_Date")
        date_input.clear()
        date_str = current_date.strftime('%m/%d/%Y')
        date_input.send_keys(date_str)
        date_input.send_keys(Keys.TAB)
        view_report_button = driver.find_element(By.ID, "run-report")

        view_report_button.click()

        time.sleep(get_wait(conn_given, conn))

        download_button = driver.find_element(By.XPATH, "//*[@title='Export a report and save it to the disk']")
        download_button.click()
        
        wait_for_download(downloads_path)
        
        parsed_data = parse_data()
        write_data(parsed_data, str(out_excel), curr_date=current_date)


        current_date += timedelta(days=1)

    print("All downloads done. MonthlySummaryReport.xlsx in downloads folder.")
    driver.quit()

if __name__ == "__main__":
    main()
    os.system("pause")


