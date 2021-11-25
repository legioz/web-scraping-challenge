from pathlib import Path
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from datetime import timedelta
import webdrivermanager
import time
import os
from dotenv import load_dotenv


load_dotenv()
OUTPUTDIR = Path("output")
browser_lib = Selenium()


def open_website(url):
    driver = webdrivermanager.ChromeDriverManager()
    driver.download_and_install("94.0.4606.61")
    driver_path = Path(driver.link_path).joinpath(driver.driver_filenames.get(driver.get_os_name()))
    options = browser_lib._get_driver_args("chrome")[0]["options"]
    if driver.get_os_name() == "linux":
        options.binary_location = "/usr/bin/chromium-browser"
    browser_lib.set_download_directory(OUTPUTDIR.__str__())
    prefs = {
        "download.default_directory": OUTPUTDIR.resolve(strict=True).__str__(),
        "download.directory_upgrade": True,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    }
    options.add_experimental_option("prefs", prefs)
    browser_lib.open_browser(url, browser="chrome", options=options, executable_path=driver_path)


def get_agencies_elements(name=None):
    if name is not None:
        agency = browser_lib.find_element(
            f"//div[@id='agency-tiles-widget']//span[text()='{name}']/.."
        )
        return agency
    dive_in = browser_lib.find_element("//a[@href='#home-dive-in']").click()
    browser_lib.wait_until_page_contains_element(
        "//a[@href='#home-dive-in' and @aria-expanded='true']"
    )
    browser_lib.wait_until_page_contains_element("//div[@id='agency-tiles-widget']")
    browser_lib.wait_until_element_is_visible("//div[@id='agency-tiles-widget']//a")
    agencies = browser_lib.find_elements(
        "//div[@id='agency-tiles-widget']//a[contains(@href, '/drupal/summary')]/span"
    )
    return list(zip(agencies[::2], agencies[1::2]))


def get_agencies_spending():
    agencies = get_agencies_elements()
    agencies_bills = [(agency[0].text, agency[1].text) for agency in agencies]
    return agencies_bills


def create_agencies_excel(agencies):
    lib_files = Files()
    try:
        lib_files.create_workbook(OUTPUTDIR.joinpath("agencies.xlsx").__str__())
        lib_files.rename_worksheet("Sheet", "Agencies")
        lib_files.append_rows_to_worksheet(agencies)
        lib_files.save_workbook()
    finally:
        lib_files.close_workbook()


def create_individual_investiments_excel(agency_investments):
    lib_files = Files()
    try:
        lib_files.open_workbook(OUTPUTDIR.joinpath("agencies.xlsx").__str__())
        lib_files.create_worksheet("Individual Investiments", agency_investments)
        lib_files.save_workbook()
    finally:
        lib_files.close_workbook()


def get_agency():
    agency_name = os.getenv("AGENCY_NAME")
    if not agency_name:
        raise Exception("Please provide an agency name in the .env file")
    return agency_name


def download_business_case_pdf():
    download_urls = browser_lib.find_elements(
        "//div[@id='investments-table-object_wrapper']//tbody//tr//td[1]//a"
    )
    for url_id in download_urls:
        url = browser_lib.get_element_attribute(url_id, "href")
        browser_lib.execute_javascript(f"window.open('{url}')")
        filename = f"{url_id.text}.pdf"
        browser_lib.switch_window("NEW")
        browser_lib.wait_until_element_is_visible(
            "//div[@id='business-case-pdf']/a", timedelta(minutes=1)
        )
        browser_lib.find_element("//div[@id='business-case-pdf']/a").click()
        while not OUTPUTDIR.joinpath(filename).is_file():
            time.sleep(1)
        browser_lib.close_window()
        browser_lib.switch_window("MAIN")
        print(f" [x] file {filename} downloaded")
    time.sleep(1)
    browser_lib.close_browser()


def get_agency_specific_spending(agency):
    agency = get_agencies_elements(agency).click()
    browser_lib.wait_until_page_contains_element(
        "//div[@id='investments-table-object_length']/label/select",
        timedelta(minutes=1),
    )
    browser_lib.wait_until_element_is_visible(
        "//div[@id='investments-table-object_length']/label/select",
        timedelta(minutes=1),
    )
    browser_lib.set_focus_to_element("//h4[text()='Individual Investments']")
    button_show_all_entries = browser_lib.find_element(
        "//div[@id='investments-table-object_length']/label/select/option[contains(text(),'All')]"
    )
    button_show_all_entries.click()
    browser_lib.wait_until_page_contains_element(
        "//a[@id='investments-table-object_last' and contains(@class, 'disabled')]",
        timedelta(minutes=1),
    )
    browser_lib.wait_until_element_is_visible(
        "//a[@id='investments-table-object_last' and contains(@class, 'disabled')]",
        timedelta(minutes=1),
    )
    investments = browser_lib.find_elements(
        "//div[@id='investments-table-object_wrapper']//tbody//tr//td"
    )
    return investments


def scrapy_specific_agency(agency):
    investments = get_agency_specific_spending(agency)
    rows = []
    row = []
    count = 0
    for td in investments:
        if count < 6:
            count += 1
            row.append(td.text)
        else:
            count = 0
            rows.append(row.copy())
            row.clear()
    return rows


def main():
    try:
        agency = get_agency()
        print(" [x] agency found in environment")
        open_website("https://itdashboard.gov/")
        agencies = get_agencies_spending()
        create_agencies_excel(agencies)
        print(" [x] agencies xlsx created")
        individual_investiments = scrapy_specific_agency(agency)
        create_individual_investiments_excel(individual_investiments)
        print(" [x] agency individual investments sheet created")
        download_business_case_pdf()
    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()
