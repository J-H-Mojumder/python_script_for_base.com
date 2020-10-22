import time
from selenium import webdriver
import xlsxwriter
page_number = 1
driver = webdriver.Firefox(executable_path=r"E:\geckodriver.exe")
url = "https://www.base.com/blu-ray/pg735/bn10007132/products.htm?filter=a%3a14%3a260339"
row_number = 1
outputWorkbook = xlsxwriter.Workbook("base_links_Blu-ray_comedy.xlsx")
outputsheet = outputWorkbook.add_worksheet()
outputsheet.write("A1", "Category")
outputsheet.write("B1", "Product Name")
outputsheet.write("C1", "Link")

while 1:
    try:
        driver.get(url)
        driver.implicitly_wait(8)
        category = driver.find_elements_by_class_name("search-filters-selected")
        div = driver.find_elements_by_class_name("plist-centered")
        titles = driver.find_elements_by_class_name("title")
    except Exception as e:
        print("Problem in opening link!!")
    else:
        print("link opened")
        for title in titles:
            print("category :"+category[0].text)
            name = title.find_elements_by_tag_name("a")
            product_name = name[0].get_property("title")
            print("Name: "+product_name)
            product_link = name[0].get_property("href")
            print("Link "+product_link)

            outputsheet.write(row_number, 0, category[0].text)
            outputsheet.write(row_number, 1, product_name)
            outputsheet.write(row_number, 2, product_link)
            row_number += 1
        try:
            pagination = div[-1].find_elements_by_tag_name("a")
            print("page number "+str(page_number))
            print(pagination[-1].text)
            next = pagination[-1].text
            pagination[-1].click()
            print("Clicked on next")
            page_number += 1
        except Exception as e:
            print("End!!")
            break
        else:
            current_link = url
            time.sleep(2)
            next_link = driver.current_url
            if next != "Next":
                break;
            else:
                url = next_link
outputWorkbook.close()