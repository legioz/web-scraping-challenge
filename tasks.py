from pathlib import Path
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from datetime import timedelta


OUTPUTDIR = Path(__file__).resolve(strict=True).parent.joinpath("output")
browser_lib = Selenium()

def open_website(url):
    browser_lib.open_available_browser(url)


def get_agencies_elements(name: str = None):
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


def create_individual_investiments_excel(agency):
    lib_files = Files()
    try:
        lib_files.open_workbook(OUTPUTDIR.joinpath("agencies.xlsx"))
        lib_files.create_worksheet("Individual Investiments", agency)
        lib_files.save_workbook()
    finally:
        lib_files.close_workbook()


def get_agency():
    # TODO: get agency name from file
    return "Department of Agriculture"


def download_business_case_pdf(agency):
    # TODO: 
    pass

def scrapy_specific_agency(agency):
    agency = get_agencies_elements(agency).click()
    browser_lib.wait_until_page_contains_element(
        "//div[@id='investments-table-object_length']/label/select",
        timedelta(minutes=1),
    )
    browser_lib.wait_until_element_is_visible(
        "//div[@id='investments-table-object_length']/label/select",
        timedelta(minutes=1),
    )
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
        "//div[@id='investments-table-object_wrapper']//tr//td"
    )
    return investments


def main():
    try:
        open_website("https://itdashboard.gov/")
        agencies = get_agencies_spending()
        create_agencies_excel(agencies)
        agency = get_agency()
        individual_investiments = scrapy_specific_agency(agency)
        print(individual_investiments)
    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()
