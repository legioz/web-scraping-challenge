from pathlib import Path
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from datetime import timedelta
from RPA.HTTP import HTTP
from selenium.webdriver import FirefoxProfile
from selenium.webdriver.firefox.options import Options
import webdrivermanager
import time

ROOT_DIR = Path(__file__).resolve(strict=True).parent
OUTPUTDIR = ROOT_DIR.joinpath("output")
browser_lib = Selenium()


def open_website(url):
    driver = webdrivermanager.GeckoDriverManager()
    driver.download_and_install("v0.30.0")
    executable = driver.link_path.joinpath(driver.get_driver_filename()).__str__()
    profile = FirefoxProfile()
    mime_types = "application/pdf"
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", mime_types)
    profile.set_preference(
        "browser.download.dir", OUTPUTDIR.joinpath("pdf").__str__()
    )
    profile.set_preference("pdfjs.disabled", True)
    profile.set_preference("browser.link.open_newwindow", 3)
    profile.set_preference("browser.link.open_newwindow.restriction", 0)
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", mime_types)
    profile.set_preference("plugin.disable_full_page_plugin_for_types", mime_types)
    options = Options()
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.manager.showWhenStarting", False)
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", mime_types)
    options.set_preference(
        "browser.download.dir", OUTPUTDIR.joinpath("pdf").__str__()
    )
    options.set_preference("pdfjs.disabled", True)
    options.set_preference("browser.link.open_newwindow", 3)
    options.set_preference("browser.link.open_newwindow.restriction", 0)
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", mime_types)
    options.set_preference("plugin.disable_full_page_plugin_for_types", mime_types)
    browser_lib.open_browser(url, ff_profile_dir=profile, options=options, executable_path=executable)


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
        lib_files.create_workbook(OUTPUTDIR.joinpath("agencies.xlsx"))
        lib_files.rename_worksheet("Sheet", "Agencies")
        lib_files.append_rows_to_worksheet(agencies)
        lib_files.save_workbook()
    finally:
        lib_files.close_workbook()


def create_individual_investiments_excel(agency_investments):
    lib_files = Files()
    try:
        lib_files.open_workbook(OUTPUTDIR.joinpath("agencies.xlsx"))
        lib_files.create_worksheet("Individual Investiments", agency_investments)
        lib_files.save_workbook()
    finally:
        lib_files.close_workbook()


def get_agency():
    # TODO: get agency name from file
    return "Department of Agriculture"


def download_business_case_pdf(agency):
    browser_lib.set_download_directory(OUTPUTDIR.joinpath("pdf"))
    download_urls = browser_lib.find_elements(
        "//div[@id='investments-table-object_wrapper']//tbody//tr//td[1]//a"
    )
    browser_lib.execute_javascript(
        "Array.from(document.getElementsByTagName('a')).forEach((c)=>{c.target='_blank'})"
    )
    for url_id in download_urls:
        url_id.click()
        filename = f"{url_id.text}.pdf"
        browser_lib.switch_window("NEW")
        browser_lib.wait_until_element_is_visible("//div[@id='business-case-pdf']/a", timedelta(minutes=1))
        browser_lib.find_element("//div[@id='business-case-pdf']/a").click()
        while not OUTPUTDIR.joinpath("pdf").joinpath(filename).is_file():
            time.sleep(1)
        browser_lib.close_window()
        browser_lib.switch_window("MAIN")


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
        open_website("https://itdashboard.gov/")
        agencies = get_agencies_spending()
        create_agencies_excel(agencies)
        agency = get_agency()
        individual_investiments = scrapy_specific_agency(agency)
        create_individual_investiments_excel(individual_investiments)
        download_business_case_pdf(agency)
    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()
