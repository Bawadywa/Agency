# +
import datetime
import time

import pandas as pd
import re
import PyPDF2

from RPA.Browser.Selenium import Selenium
from RPA.FileSystem import FileSystem
from RPA.JSON import JSON

browser = Selenium()

MAIN_URL = "https://itdashboard.gov/"
EXCEL_NAME = "Report.xlsx"


AGENCY_NAME = JSON().load_json_from_file("devdata/env.json")["agency_name"]
FILE_SYSTEM = FileSystem()


def open_url(url: str):
    browser.open_available_browser(url)
    browser.maximize_browser_window()


def dive_in():
    browser.click_link("#home-dive-in")
    browser.wait_until_page_contains_element(locator="css:div#agency-tiles-widget span",
                                             timeout=datetime.timedelta(seconds=20))
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
    header = ["Agency", "Amount"]
    data_frame = pd.DataFrame(list(data.items()))
    data_frame.columns = header
    data_frame.to_excel(workbook_path, "Agencies", index=False)


def open_agency(name: str):
    browser.click_link(locator="partial link:{}".format(name))
    browser.wait_until_page_contains_element(locator="id:investments-table-object_wrapper",
                                             timeout=datetime.timedelta(seconds=15))
    # wait until investments table appears


def show_all_entries():
    showing_elem = browser.find_element(locator="id:investments-table-object_info")
    showing_text = browser.get_text(showing_elem)   # get showing text to compare in wait until
    browser.select_from_list_by_label("name:investments-table-object_length", "All")
    # select all entries of investments
    browser.wait_until_element_does_not_contain(locator=showing_elem, text=showing_text,
                                                timeout=datetime.timedelta(seconds=15))
    # wait until investments table refreshes


def html_table_to_excel(workbook_path: str, worksheet_name: str):
    data = pd.read_html(browser.get_source(), match="UII")
    table = data[1]
    with pd.ExcelWriter(workbook_path, engine="openpyxl", mode="a") as writer:
        table.to_excel(writer, worksheet_name, index=False)

    return table


def compare_pdf(path: str, file_name: str, record):
    name_pattern, uii_pattern = r"(?<=1\. Name of this Investment:\n \n)[^\n]+",\
                                r"(?<=2\. Unique Investment Identifier \(UII\):\n \n)[^\n]+"
    # using regexp for search

    pdf_reader = PyPDF2.PdfFileReader("{}{}.pdf".format(path, file_name))
    text = pdf_reader.getPage(0).extractText()
    # getting text of the first page
    match_name, match_uii = re.search(name_pattern, text), re.search(uii_pattern, text)
    name, uii = "", ""
    if match_name is not None:
        name = match_name.group(0)
    else:
        return False

    if match_uii is not None:
        uii = match_uii.group(0)
    else:
        return False

    print("-"*5 + "Comparing {}.pdf".format(file_name) + "-"*5)
    if record["UII"] == uii and record["Investment Title"] == name:
        print("Row Index = {}, UII: {}, Name of Investment: {}\n{}"
              .format(record.name + 2, uii, name, record.to_dict()), end="\n\n")
        return True
        # + 2 to show exact row index from excel
    return False


def download_pdf(link: str):
    browser.click_link(link, "CTRL+ALT")
    # click on UII and open webpage in new tab
    browser.switch_window("NEW")
    # switch to new tab
    browser.wait_until_page_contains_element(locator="business-case-pdf",
                                             timeout=datetime.timedelta(seconds=10))
    pdf_link_parent = browser.find_element(locator="business-case-pdf")
    pdf_link = browser.find_element(locator="tag:a", parent=pdf_link_parent)
    browser.click_link(locator=pdf_link)
    # download pdf
    browser.wait_until_page_does_not_contain_element(locator="css:div#business-case-pdf span",
                                                     timeout=datetime.timedelta(seconds=15))
    # wait for pdf to generate
    browser.switch_window("MAIN")
    # switch to main tab


def check_downloaded(path: str, file_name: str):
    try:
        timeout, count = 10, 1
        while count < timeout:
            if FILE_SYSTEM.does_file_exist(path="{}{}.pdf".format(path, file_name)):
                return True
            time.sleep(1)
            count += 1
    except:
        return False


def table_actions(path, table):
    for i in range(len(table)):
        uii = table["UII"][i]
        try:
            link = browser.find_element(locator="link:{}".format(uii))
            # find link
            download_pdf(link)
            # download pdf
            flag_download = check_downloaded(path, uii)
            # check if downloaded
            if flag_download:
                flag_compare = compare_pdf(path, uii, table.loc[i])
                # check compare
                if not flag_compare:
                    print("Error. Values 'UII' and 'Investment Title' are not the same.")

            else:
                continue

        except:
            break


def delete_pdfs(path: str):
    FILE_SYSTEM.remove_files(*list(map(lambda file: file.path,
                                       FILE_SYSTEM.find_files(path + "*.pdf"))))


def first_page(path: str):
    open_url(MAIN_URL)
    dive_in()
    fill_excel_agencies(workbook_path=r"{}{}".format(path, EXCEL_NAME),
                        data=get_agencies())


def second_page(path: str):
    open_agency(AGENCY_NAME)
    show_all_entries()
    table = html_table_to_excel(r"{}{}".format(path, EXCEL_NAME), AGENCY_NAME)
    table_actions(path, table)


def main(path: str):
    try:
        browser.set_download_directory(path)
        delete_pdfs(path)
        first_page(path)
        second_page(path)
    finally:
        browser.close_all_browsers()

