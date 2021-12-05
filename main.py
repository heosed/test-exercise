from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.PDF import PDF

from config import Agency_Name

import os
import time


dir_path = os.path.dirname(os.path.realpath(__file__))

browser_lib = Selenium()
browser_lib.set_download_directory(directory=f"{dir_path}/output")

lib = Files()
pdf = PDF()

url = "https://itdashboard.gov/"


def main():
    try:
        open_the_website(url)
        click_on_dive_in()
        agencies = get_all_agencies()
        print(agencies)
        table, link = get_agency_info(Agency_Name, agencies)
        links = download_pdfs(table, link)
        compare_pdfs(table, links)
        create_xlsx(table,agencies)

    finally:
        browser_lib.close_all_browsers()


def open_the_website(url):
    browser_lib.open_available_browser(url, headless = True)
    print(f"Opened page {url}.....")


def click_on_dive_in():
    startbutton = browser_lib.get_webelement(
        "xpath://a[@href='#home-dive-in']")
    browser_lib.click_button(startbutton)


def get_all_agencies():
    print("Scraping agencies.....")
    columnlocator = 'xpath://*[@id="agency-tiles-widget"]'
    browser_lib.wait_until_page_contains_element(columnlocator, timeout=20)
    colums = browser_lib.get_webelement(
        columnlocator).find_elements_by_class_name('col-sm-12')
    agencies = {}
    for colum in colums:
        name = colum.find_element_by_class_name('h4.w200').text
        amounts = colum.find_element_by_class_name('h1.w900').text
        link_to_agency = colum.find_element_by_tag_name(
            'a').get_attribute('href')
        agencies[name] = {"amounts": amounts, "link": link_to_agency}
    return agencies


def create_xlsx(table, agencies):
    print("Creating xlsx....")
    sheet = lib.create_workbook(path="output/Agencies.xlsx", fmt='xlsx')
    lib.rename_worksheet("Sheet", "Agencies")

    sheet.set_cell_value(row=1, column="A", value="Agency Name")
    sheet.set_cell_value(row=1, column="B", value="Agency Amount")

    index = 1
    for col in agencies.items():
        index += 1
        sheet.set_cell_value(row=index, column="A", value=col[0])
        sheet.set_cell_value(
            row=index,
            column="B",
            value=col[1].get("amounts"))

    sheet.create_worksheet(name="Investments")
    for row in table:
        if 'UUI_link' in row:
            del row['UUI_link']
        sheet.append_worksheet(content=row, name="Investments")
    lib.save_workbook("output/Agencies.xlsx")


def get_agency_info(name, agencies):
    print("Parsing agency info.....")
    link = agencies.get(name).get('link')
    browser_lib.go_to(link)
    browser_lib.wait_until_page_contains_element(
        "id:investments-table-object_wrapper", timeout=20)

    browser_lib.get_webelement(
        'xpath://*[@id="investments-table-object_length"]/label/select/option[4]').click()
    browser_lib.wait_until_page_does_not_contain_element(
        "xpath://div[@class='dataTables_paginate paging_full_numbers']/a[@class='paginate_button next']",
        timeout=10)
    table = browser_lib.get_webelement("id:investments-table-object")
    table_rows = table.find_elements_by_css_selector(
        "body.page-summary #investments-table-object_wrapper table.dataTable tbody tr")
    excell_table = extract_tables(table_rows)
    return excell_table, link


def extract_tables(table_rows):
    excell_table = []
    for table_row in table_rows:
        try:
            uui_link = table_row.find_element_by_tag_name(
                "a").get_attribute('href')
        except Exception:
            uui_link = "---"
        rows = table_row.find_elements_by_tag_name("td")
        tab = {"UUI": rows[0].text, "UUI_link": uui_link,
               "Bureau": rows[1].text, "Investment Title": rows[2].text,
               "Total Spending": rows[3].text, "Type": rows[4].text,
               "CIO": rows[5].text, "number_of_proj": rows[6].text}
        excell_table.append(tab)
    return excell_table


def download_pdfs(tables, link):
    links = []

    for table in tables:
        if table.get('UUI_link') != '---':
            links.append(table.get('UUI_link'))

    print("Downloading PDFs.....")
    for index, link in enumerate(links):
        try:
            browser_lib.go_to(link)
            browser_lib.wait_until_page_contains_element(
                "class:row.sameHeight", timeout=10)
            browser_lib.get_webelement("id:business-case-pdf").click()
            wait_download = True
            while wait_download:
                pdf_name = link.split("/")[-1]
                if os.path.exists(f"{dir_path}/output/{pdf_name}.pdf"):
                    print(f"{index+1} of {len(links)} done")
                    wait_download = False
        except Exception as e:
            print(e)
            pass
    browser_lib.close_browser()
    print(links)
    return links


def compare_pdfs(table, links):

    print("\nStarted comparing data from values with data from PDF.....")

    index = 1
    for tab in table:
        try:
            uui_link = tab.get("UUI_link")
            if uui_link in links:
                pdf_name = uui_link.split("/")[-1]
                text = pdf.get_text_from_pdf(
                    f"./output/{pdf_name}.pdf", trim=False, pages=1)
                section_a = text.get(1).split("Section")[1]

                UUI_in_pdf = section_a.split("\n")[-2].split(":")[-1]
                inv_title_in_pdf = section_a.split("\n")[-3].split(":")[-1]
                UUI_in_values = tab.get("UUI")
                inv_title_in_values = tab.get('Investment Title')

                print("\n")
                print(f"PDF {index} of {len(links)}")
                print(f"UUI in PDF :{UUI_in_pdf}")
                print(f"UUI in values : {UUI_in_values}")
                print(f"Investment Title in PDF :{inv_title_in_pdf}")
                print(f"Investment Title in values : {inv_title_in_values}")
                index += 1
        except Exception as e:
            print(e)
            pass


if __name__ == "__main__":
    main()
