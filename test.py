# +
import datetime
import pandas as pd
import re
import PyPDF2

from RPA.Browser.Selenium import Selenium
from RPA.FileSystem import FileSystem
from RPA.JSON import JSON

browser = Selenium()

main_url = "https://itdashboard.gov/"
excel_name = "Report.xlsx"


agency_name = JSON().load_json_from_file("devdata/env.json")["agency_name"]


def open_url(url: str):
    browser.open_available_browser(url)
    browser.maximize_browser_window()


def dive_in():
    browser.click_link("#home-dive-in")
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
    with pd.ExcelWriter(workbook_path, engine="openpyxl", mode="a") as writer:
        data[1].to_excel(writer, worksheet_name, index=False)


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
        browser.wait_until_page_does_not_contain_element(locator="css:div#business-case-pdf span",
                                                         timeout=datetime.timedelta(seconds=10))
        # wait for pdf to generate
        browser.switch_window("MAIN")
        # switch to main tab
        try:
            browser.wait_until_page_does_not_contain_element(locator="id:main-content",
                                                             timeout=datetime.timedelta(seconds=2))
        except:
            continue
        # wait until file is downloaded


def compare_pdfs_excel(path: str, workbook_path: str, worksheet_name: str):
    pdf_files = FileSystem().find_files(path + r"*.pdf")
    pdf_paths = list(map(lambda file: file.path, pdf_files))
    name_pattern, uii_pattern = r"(?<=1\. Name of this Investment:\n \n)[^\n]+",\
                                r"(?<=2\. Unique Investment Identifier \(UII\):\n \n)[^\n]+"
    pdf_data = []
    for path in pdf_paths:
        pdf_reader = PyPDF2.PdfFileReader(path)
        text = pdf_reader.getPage(0).extractText()
        match_name, match_uii = re.search(name_pattern, text), re.search(uii_pattern, text)
        name, uii = "", ""
        if match_name is not None:
            name = match_name.group(0)

        if match_uii is not None:
            uii = match_uii.group(0)
        pdf_data.append({"name": name, "uii": uii})

    data_frame = pd.read_excel(workbook_path, worksheet_name)
    pd.set_option("display.max_columns", 20)
    records = list(data_frame.to_dict("records"))

    print("-"*5 + "{}".format("Comparing pdf and excel") + "-"*5)
    record_index = 0
    for record in records:
        for index in range(len(pdf_data)):
            if record["UII"] == pdf_data[index]["uii"] and record["Investment Title"] == pdf_data[index]["name"]:
                print("Row Index = {}, Name of Investment: {}, UII: {}\n{}"
                      .format(record_index + 2, *pdf_data[index].values(), record), end="\n\n")
                pdf_data.remove(pdf_data[index])
                break
                # + 2 to show exact row index from excel
        record_index += 1


def delete_pdfs(path: str):
    FileSystem().remove_files(*list(map(lambda file: file.path,
                                        FileSystem().find_files(path + "*.pdf"))))


def first_page(path: str):
    open_url(main_url)
    dive_in()
    fill_excel_agencies(workbook_path=r"{}{}".format(path, excel_name),
                        data=get_agencies())


def second_page(path: str):
    open_agency(agency_name)
    show_all_entries()
    html_table_to_excel(r"{}{}".format(path, excel_name), agency_name)
    download_pdf(links=get_investments_links())
    compare_pdfs_excel(path=path, workbook_path=r"{}{}".format(path, excel_name),
                       worksheet_name=agency_name)


def main(path: str):
    try:
        browser.set_download_directory(path)
        delete_pdfs(path)
        first_page(path)
        second_page(path)
    finally:
        browser.close_all_browsers()
