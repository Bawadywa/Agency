# +
import datetime
import time

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.JSON import JSON
from RPA.FileSystem import FileSystem

browser = Selenium()

main_url = "https://itdashboard.gov/"
excel_name = "Report.xlsx"


agency_name = JSON().load_json_from_file("devdata/env.json")["agency_name"]


def wait(seconds: int):
    time.sleep(seconds)


def open_url(url: str):
    browser.open_available_browser(url)
    browser.maximize_browser_window()


def dive_in():
    dive_in_btn = browser.find_element(locator="xpath:/html/body/main/div[1]/div/div/div[3]/"
                                               "div/div/div/div/div/div/div/div/div/a")
    # find dive in button
    browser.click_button(dive_in_btn)
    browser.wait_until_page_contains_element(locator="css:div#agency-tiles-widget span")
    # wait until agencies table appears on the page


def get_agencies() -> dict:
    agencies_table = browser.find_element(locator="id:agency-tiles-widget")
    # find agencies table using id
    agencies_cells = browser.find_elements(locator="css:div.wrapper > div > div", parent=agencies_table)
    # find all cells of the table using css selector

    agencies_dict = {}      # in this dictionary will all values be saved

    for cell in agencies_cells:
        links = browser.find_elements(locator="tag:a", parent=cell)
        # get links, there are 3 of them
        spans = browser.find_elements(locator="tag:span", parent=links[1])
        # in 2 link there is needed information
        agency_name, agency_amount = browser.get_text(spans[0]), browser.get_text(spans[1])
        # get agency name and amount using get_text
        agencies_dict[agency_name] = agency_amount

    return agencies_dict


def fill_excel_agencies(workbook_path: str, data: dict):
    lib = Files()
    workbook = lib.create_workbook(path=workbook_path)

    row, col = 1, 1
    for key, value in data.items():
        # filling worksheet with agency names and amounts
        workbook.set_cell_value(row, col, key)
        workbook.set_cell_value(row, col + 1, value)
        row += 1

    workbook.rename_worksheet("Agencies", workbook.sheetnames[0])
    # rename first worksheet with "Agencies"

    lib.save_workbook(path=workbook_path)


def open_agency(name: str):
    btn = browser.find_element(locator="xpath://span[contains(text(), \"{}\")]/parent::*".format(name))
    # find parent link of the span with agency name to follow the link
    browser.click_link(btn)
    browser.wait_until_page_contains_element(locator="id:investments-table-object_wrapper",
                                             timeout=datetime.timedelta(seconds=10))
    # wait until investments table appears


def show_all_entries():
    browser.select_from_list_by_label("name:investments-table-object_length", "All")
    # select all entries of investments
    wait(10)
    # wait until investments table refreshes


def get_table_header() -> list:
    header_table = browser.find_element(locator="xpath://*[@id=\"investments-table-object_wrapper\"]"
                                                "/div[3]/div[1]/div/table")
    # find header table to get titles for excel file

    data = []
    i = 1
    while True:     # using while true with try except for getting all the rows of the table
        data.append([])
        try:
            for j in range(1, 8):
                data[i - 1].append(browser.get_table_cell(header_table, i, j))
        except Exception:
            data.pop(-1)    # deleting last row because it'll be empty
            break
        i += 1

    header = data[-1]   # last row contains titles
    return header


def get_table_data() -> list:
    table = browser.find_element(locator="id:investments-table-object")
    # find table using id

    data = []
    i = 1
    while True:     # using while true with try except for getting all the rows of the table
        data.append([])
        try:
            for j in range(1, 8):
                data[i-1].append(browser.get_table_cell(table, i, j))
        except Exception:
            data.pop(-1)    # deleting last row because it'll be empty
            break
        i += 1

    result_data = []
    # this for cycle means making result list without rows which are empty (having empty strings - "")
    for row in data:
        if row.count("") == len(row):
            continue
        else:
            result_data.append(row)

    return result_data


def fill_excel_investments(workbook_path: str, worksheet_name: str, header: list, data: list):
    lib = Files()
    workbook = lib.open_workbook(workbook_path)
    # opens already created workbook
    workbook.create_worksheet(worksheet_name)
    # creates new worksheet for investments

    row, col = 1, 1
    for title in header:
        # filling first row of excel as header
        workbook.set_cell_value(row, col, title)
        col += 1

    # filling main data
    for row in range(len(data)):
        for col in range(len(data[row])):
            workbook.set_cell_value(row + 2, col + 1, data[row][col])   
            # using row + 2 because first row is header

    lib.save_workbook(path=workbook_path)


def get_investments_links() -> list:
    table = browser.find_element(locator="id:investments-table-object")
    links = browser.find_elements(locator="tag:a", parent=table)
    return links


def download_pdf(links: list):
    for link in links:
        browser.click_link(link, "CTRL+ALT")
        # click on UII and open webpage in new tab
        browser.switch_window("NEW")
        # switch to new tab
        browser.wait_until_page_contains_element(locator="business-case-pdf")
        pdf_link_parent = browser.find_element(locator="business-case-pdf")
        pdf_link = browser.find_element(locator="tag:a", parent=pdf_link_parent)
        browser.click_link(locator=pdf_link)
        wait(10)
        # wait for pdf to generate
        browser.switch_window("MAIN")
        # switch to main tab
        wait(5)


def first_page(path: str):
    open_url(main_url)
    dive_in()
    fill_excel_agencies(workbook_path=r"{}{}".format(path, excel_name), data=get_agencies())


def second_page(path: str):
    open_agency(agency_name)
    show_all_entries()
    fill_excel_investments(workbook_path=r"{}{}".format(path, excel_name),
                           worksheet_name=agency_name,
                           header=get_table_header(),
                           data=get_table_data())
    download_pdf(links=get_investments_links())


def main(path: str):
    try:
        browser.set_download_directory(path)
        first_page(path)
        second_page(path)
    finally:
        browser.close_all_browsers()
