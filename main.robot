from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.PDF import PDF

from config import Agency_Name

import os
import time


browser_lib = Selenium()
lib = Files()
pdf = PDF()

url = "https://itdashboard.gov/"
dir_path = os.path.dirname(os.path.realpath(__file__))


def main():
    try:
        open_the_website(url)
        click_on_dive_in()
        agencies = get_all_agencies()
        table,link = get_agency_info(Agency_Name,agencies)
        pdfnames = download_pdfs(table,link)
        create_xlsx(table,agencies)
        compare_pdfs(table,pdfnames)

        
    finally:
        browser_lib.close_all_browsers()


def open_the_website(url):
    browser_lib.open_available_browser(url,headless = True)
    print(f"Opened page {url}.....")

def click_on_dive_in():
    startbutton = browser_lib.get_webelement('xpath://*[@id="node-23"]/div/div/div/div/div/div/div/a')
    browser_lib.click_button(startbutton)


def get_all_agencies():
    print("Scraping agencies.....")
    columnlocator = 'xpath://*[@id="agency-tiles-widget"]'
    browser_lib.wait_until_page_contains_element(columnlocator, timeout = 20)
    time.sleep(3)
    colums = browser_lib.get_webelement(columnlocator).find_elements_by_class_name('col-sm-12')
    agencies = {}
    for colum in colums:
        name = colum.find_element_by_class_name('h4.w200').text
        amounts = colum.find_element_by_class_name('h1.w900').text
        link_to_agency = colum.find_element_by_tag_name('a').get_attribute('href')
        agencies[name] = {"amounts" : amounts, "link" : link_to_agency}
    return agencies


def create_xlsx(table,agencies):
    print("Creating xlsx....")
    sheet = lib.create_workbook(path = "output/Agencies.xlsx",fmt='xlsx')
    lib.rename_worksheet("Sheet", "Agencies")

    sheet.set_cell_value(row = 1, column = "A",value = "Agency Name")
    sheet.set_cell_value(row = 1, column = "B",value = "Agency Amount")
    
    index = 1
    for col in agencies.items():
        index += 1
        sheet.set_cell_value(row = index, column = "A",value = col[0])
        sheet.set_cell_value(row = index, column = "B",value = col[1].get("amounts"))

    sheet.create_worksheet(name = "Investments")
    for row in table:
        if 'UUI_link' in row:
            del row['UUI_link']
        sheet.append_worksheet(content = row, name="Investments")
    lib.save_workbook("output/Agencies.xlsx")


def get_agency_info(name,agencies):
    print("Parsing agency info.....")
    link = agencies.get(name).get('link')
    browser_lib.go_to(link)
    browser_lib.wait_until_page_contains_element("id:investments-table-object_wrapper", timeout = 10)
    browser_lib.get_webelement('xpath://*[@id="investments-table-object_length"]/label/select/option[4]').click()
    time.sleep(10)
    table = browser_lib.get_webelement("class:dataTables_scrollBody")
    table_rows = table.find_elements_by_css_selector("body.page-summary #investments-table-object_wrapper table.dataTable tbody tr")

    excell_table = extract_tables(table_rows)
    return excell_table,link

    
def extract_tables(table_rows):
    excell_table = []
    for table_row in table_rows:
        try:
            uui_link = table_row.find_element_by_tag_name("a").text
        except Exception as e:
            uui_link = "---"
        rows = table_row.find_elements_by_tag_name("td")
        uui = rows[0].text
        bureau = rows[1].text
        invest_title = rows[2].text
        spending = rows[3].text
        type = rows[4].text
        cio = rows[5].text
        num_proj = rows[6].text
        temp_table = {"UUI": uui,"UUI_link":uui_link, 
                      "Bureau": bureau, "Investment Title": invest_title, 
                      "Total Spending" : spending, "Type": type, 
                      "CIO":cio,"number_of_proj":num_proj}
        excell_table.append(temp_table)
    return excell_table            

def download_pdfs(tables,link):
    links = []
    for table in tables:
        if table.get('UUI_link') != '---':
            links.append(f"{link}/{table.get('UUI_link')}")
    print("Downloading PDFs.....")
    index = 1
    browser_lib.set_download_directory(directory = f"{dir_path}/output")
    browser_lib.open_available_browser(headless = True)
    for link in links:
        try:
            browser_lib.go_to(link)
            time.sleep(1.5)
            browser_lib.get_webelement('xpath://*[@id="business-case-pdf"]').click()
            time.sleep(4)
            print(f"{index} of {len(links)} done")
            index += 1
        except Exception as e:
            print(e)
            pass
    time.sleep(6)
    browser_lib.close_window()
    return links

def compare_pdfs(table,pdfnames):

    print("\nStarted comparing data from values with data from PDF.....")
    uuis_list = []

    for pdfname in pdfnames:
        uuis_list.append(pdfname[-13:])

    index = 1
    for tab in table:
        try:
            uui = tab.get("UUI")
            if uui in uuis_list:
                text = pdf.get_text_from_pdf(f"./output/{uui}.pdf",trim = False,pages = 1)
                secA = text.get(1).split("Section")[1]

                UUI_in_pdf = secA.split("\n")[-2].split(":")[-1]
                inv_title_in_pdf = secA.split("\n")[-3].split(":")[-1]
                UUI_in_values = tab.get("UUI")
                inv_title_in_values = tab.get('Investment Title')

                print("\n")
                print(f"PDF {index} of {len(uuis_list)}")
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
