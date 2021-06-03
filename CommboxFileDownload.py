####INSTRUCTIONS
# Install python
# Use pip to install required python packages: selenium, msedge.selenium_tools, pandas
# Add Commbox uname/pword as windows environment variables "COMMBOX_USER", "COMMBOX_PASS" respectivly
# Download MS-Edge WebDriver and add its location to PATH in windows environment variables
# Set the script variables (DownloadFolder and Output file location/name, and adjust dates for the report to download)
# Script is set to run a headless browser, if you want to change this, comment out EdgeOpts.add_argument("--headless")
# Create a new Scheduled task to run with highest privlages and whether user is logged in or not


##-------------IMPORTS-------------##
import time, os, logging, datetime, pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
from msedge.selenium_tools import EdgeOptions, Edge

##-------------VARIABLES-------------##
DownloadFolder = r"C:\Users\Dwayne\Documents\FileDownload\Download" #Temp folder for report download, report will be removed once processed, folder should remain empty
CommboxLoginPage = "https://davidshield.commbox.io/"
InsightsPage = "https://davidshield.commbox.io/insights/objects"
uname = os.getenv('COMMBOX_USER')
pword = os.getenv('COMMBOX_PASS')
ReportStartDate = datetime.datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) - datetime.timedelta(days=14) # sets report start date for 14 days previous
ReportEndDate = datetime.datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) - datetime.timedelta(days=1) # sets report end date to yesterday
OutputFile = r'C:\Users\Dwayne\Documents\FileDownload\CommboxExport_' + ReportStartDate.strftime('%Y-%m-%d') + "_" + ReportEndDate.strftime('%Y-%m-%d') + '.csv' #Final Report Location and Naming convention

##-------------FUNCTIONS-------------##
def format_query_date(date):
    return f'{date:%d-%m-%Y}'

def download_wait(directory, timeout, nfiles=None):
    #Function that waits untill a file is present in the download folder, prevents MS Edge closing before download is complete
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < timeout:
        time.sleep(1)
        dl_wait = False
        files = os.listdir(directory)
        if nfiles and len(files) != nfiles:
            dl_wait = True
        for fname in files:
            if fname.endswith(".crdownload"):
                dl_wait = True
        seconds += 1
    return seconds

def StartEdgeDriver():
    #Opens a headless Microsoft Edge Browser
    #Set all options to run headless but still render required elements
    EdgeOpts = EdgeOptions()
    EdgeOpts.use_chromium = True
    prefs = {"download.default_directory": DownloadFolder}  ##SET DOWNLOAD PATH
    EdgeOpts.add_experimental_option("prefs", prefs)
    EdgeOpts.add_argument("--window-size=1920,1080")
    EdgeOpts.add_argument("--disable-extensions")
    EdgeOpts.add_argument("--proxy-server='direct://'")
    EdgeOpts.add_argument("--proxy-bypass-list=*")
    EdgeOpts.add_argument("--start-maximized")
    EdgeOpts.add_argument("--headless")
    EdgeOpts.add_argument("--disable-gpu")
    EdgeOpts.add_argument("--disable-dev-shm-usage")
    EdgeOpts.add_argument("--no-sandbox")
    EdgeOpts.add_argument("--ignore-certificate-errors")
    driver = Edge(options=EdgeOpts)
    return driver


def LoginCommbox(driver, login_url, insights_url, username, password):
    #Log in to Commbox
    driver.get(login_url)
    login = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "tbEmail"))
    )
    driver.find_element(By.ID, "tbEmail").send_keys(username)
    driver.find_element(By.ID, "tbPassword").send_keys(password)
    driver.find_element(By.ID, "tbPassword").send_keys(Keys.ENTER)
    pageload = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "divStripTabButtonInsights"))
    )
    driver.get(insights_url)
    return driver


def DownloadReport(driver, ReportStartDate, ReportEndDate, DownloadFolder):
    #Navigates to the Insights Page, generates the report and downloads it to the DownloadFolder
    dropdown = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable(
            (By.ID, "ctl00_BodyPlaceHolder_reportsDateRangePicker_selectDatePick")
        )
    )
    ChooseDate = Select(
        driver.find_element(
            By.ID, "ctl00_BodyPlaceHolder_reportsDateRangePicker_selectDatePick"
        )
    )
    ChooseDate.select_by_index(5)
    StartDate = driver.find_element(
        By.ID, "ctl00_BodyPlaceHolder_reportsDateRangePicker_textBoxStartDate"
    )
    driver.execute_script(
        "arguments[0].value = arguments[1]", StartDate, ReportStartDate
    )
    EndDate = driver.find_element(
        By.ID, "ctl00_BodyPlaceHolder_reportsDateRangePicker_textBoxEndDate"
    )
    driver.execute_script("arguments[0].value = arguments[1]", EndDate, ReportEndDate)
    GenerateReport = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "BodyPlaceHolder_CreateReportBtn"))
    )
    GenerateReport.click()
    ExportReport = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "#CreateExlBtn > label"))
    )
    ExportReport.click()
    DownloadReport = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "Download"))
    )
    DownloadReport.click()
    download_wait(DownloadFolder, 30, 1)
    driver.close()

def GetFileName(DownloadFolder):
    #Gets the file name of the report
    os.chdir(DownloadFolder)
    FileList = os.listdir()
    Filename = FileList[0]
    return(Filename)

def LoadAndCleanExcel(Filename):
    #Loads the report into pandas, corrects 1 x Column name and set date format correctly
    df = pd.read_excel(Filename)
    #Setting Date column to date format
    df["Date"] = pd.to_datetime(df["Date"])
    #Removing white space from incorrect column heading
    df = df.rename(columns={"   Average first response time   ":"Average first response time"})
    return(df)

def DeleteReport(DownloadFolder, Filename):
    #Deletes the original downloaded report
    os.chdir(DownloadFolder)
    os.remove(Filename)

def main():
    driver = StartEdgeDriver()
    driver = LoginCommbox(driver, CommboxLoginPage, InsightsPage, uname, pword)
    DownloadReport(driver, format_query_date(ReportStartDate), format_query_date(ReportEndDate), DownloadFolder)
    Filename = GetFileName(DownloadFolder)
    report = LoadAndCleanExcel(Filename)
    #Outputs the report to the Output Folder with the new Filename
    report.to_csv(OutputFile,  encoding='utf-8-sig' , index = False)
    #Deletes downloaded report - required to ensure file is downloaded correctly before MS Edge closes
    DeleteReport(DownloadFolder, Filename)

if __name__ == "__main__":
    main()
