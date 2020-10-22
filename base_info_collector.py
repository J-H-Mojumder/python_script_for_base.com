import time
import pyautogui as pyautogui
import requests
from selenium import webdriver
import xlsxwriter
import xlrd

file_location = "C:/Users/Masum's Computer/PycharmProjects/base_com/base_links_Blu-ray_comedy.xlsx"
workBook = xlrd.open_workbook(file_location)
sheet = workBook.sheet_by_index(0)
links = []

#controlling row number of Excel sheet
link_count = 1
category = sheet.cell_value(1,0)
while link_count <= 569:
    links.append(sheet.cell_value(link_count,2))
    link_count+=1
row_counter = 1

driver = webdriver.Firefox(executable_path=r"E:\geckodriver.exe")

outputWorkbook = xlsxwriter.Workbook("base_comedy.xlsx")
outputsheet = outputWorkbook.add_worksheet()
outputsheet.write("A1", "Category")
outputsheet.write("B1", "Product Name")
outputsheet.write("C1", "Release date")
outputsheet.write("D1", "Certificate")
outputsheet.write("E1", "Price")
outputsheet.write("F1", "Product Description")
outputsheet.write("G1", "Image Link")
outputsheet.write("H1", "SKU")
outputsheet.write("I1", "Product Link")
outputsheet.write("J1","Stock Message")
outputsheet.write("K1","Express Shipping")

link_iterator = 0

#iterating several links from the excel sheet deriver by base_link_collector.py
while 1:
    if link_iterator >= 569:
        break
    else:
        try:
            driver.get(links[link_iterator])
            if row_counter == 1:
                time.sleep(2)
            else:
                time.sleep(1)
        except Exception as e:
            print("Problem opening the web link!!")
        else:
            try:
                available = driver.find_elements_by_class_name("not-avail")
                if len(available) > 0:
                    link = links[link_iterator]
                    link_iterator += 1
                    row_counter += 1
                    continue

                subsection = driver.find_elements_by_class_name("sub-section")
                description = driver.find_elements_by_id("tabs-desc")
                main_frame = driver.find_elements_by_id("main_frame")
                iframe = driver.find_elements_by_id("AWIN_CDT")
                stock = driver.find_elements_by_class_name("stock-message")
            except Exception as e:
                print("Mandetory item not found!!")
                break;
            else:
                print(str(len(main_frame)))
                try:
                    name = subsection[0].find_elements_by_tag_name("h1")
                    rrp = subsection[1].find_elements_by_class_name("rrp")
                    temp = driver.find_elements_by_id("mainImage")
                    image_link = temp[0].get_property("src")
                    temp = main_frame[0].find_elements_by_tag_name("div")
                    sku = temp[len(temp)-2].text
                except Exception as e:
                    print("No products!!")
                    print("Error in " + str(row_counter))
                    link = links[link_iterator]
                    link_iterator += 1
                    row_counter += 1
                else:
                    if len(rrp) <= 0:
                        price = driver.find_elements_by_class_name("price")
                    else:
                        print("Else")
                        price = rrp[0].find_elements_by_tag_name("strike")
                        print(price)
                        print(str(len(price)))
                        if len(price) == 0:
                            price = subsection[1].find_elements_by_class_name("price")

                    print("RRP len "+str(len(rrp)))
                    print("Len of Price :"+str(len(price)))
                    product_name = name[0].text
                    product_price = price[0].text
                    product_description = description[0].get_attribute('innerHTML')

                    release = subsection[0].find_elements_by_class_name("reldate")
                    if len(release) > 0:
                        product_release = release[0].text
                    else:
                        product_release = ""

                    certificate = subsection[0].find_elements_by_class_name("certificate")
                    if len(certificate) > 0:
                        product_certificate = certificate[0].text
                    else:
                        product_certificate = ""
                    express = driver.find_elements_by_class_name("adv-blurb")
                    if len(express) > 0:
                        express_shipping = express[0].text
                    else:
                        express_shipping = ""
                    print("Category : "+category)
                    print("Product name : "+product_name)
                    print("Release Date : "+product_release)
                    print("Certificate : "+product_certificate)
                    print("Price : "+product_price)
                    print("Description : "+product_description)
                    print("Image link : "+image_link)
                    print("SKU : "+sku)
                    print("Product URL : "+driver.current_url)
                    print("Stock msg : "+stock[0].text)
                    print("Express Shipping : "+express_shipping)


                    outputsheet.write(row_counter, 0, category)
                    outputsheet.write(row_counter, 1, product_name)
                    outputsheet.write(row_counter, 2, product_release)
                    outputsheet.write(row_counter, 3, product_certificate)
                    outputsheet.write(row_counter, 4, product_price)
                    outputsheet.write(row_counter, 5, product_description)
                    outputsheet.write(row_counter, 6, image_link)
                    outputsheet.write(row_counter, 7, sku)
                    outputsheet.write(row_counter, 8, driver.current_url)
                    outputsheet.write(row_counter, 9, stock[0].text)
                    outputsheet.write(row_counter, 10, express_shipping)

                    
                    print("Row number " + str(row_counter))
                    
                    link = links[link_iterator]
                    link_iterator += 1
                    row_counter += 1
outputWorkbook.close()
