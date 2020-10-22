# Python script using selenium webDriver to scrape base.com (the link collector script)

I have used **selenium webdriver, xlsxwriter and time** library of Python. I have used **geckodriver** as the driver. One have to add his/her geckodriver location in the below mentioned line.
```python
driver = webdriver.Firefox(executable_path=r"E:\geckodriver.exe")
```

One just have to change the URL in the 8th line of the script and replace with his or her destination page. One have to input the link of a specific critaria like one mentioned in the below code part.

```python
url = "https://www.base.com/blu-ray/pg735/bn10007132/products.htm?filter=a%3a14%3a260339"
```
I have used the code part written below to give the driver time to decide if the page has fully been loaded or it needs to take a few seconds (maximum 8 seconds) of time to wait until the page is loaded.

``` python
    idriver.implicitly_wait(8)
```

One have to input the excel file name manually in the below mentioned part of the code.

```python
outputWorkbook = xlsxwriter.Workbook("base_links_Blu-ray_comedy.xlsx")
```

It will grab only 3 columns.

```python
outputsheet.write("A1", "Category")
outputsheet.write("B1", "Product Name")
outputsheet.write("C1", "Link")
```
The output file will be populated with the thubnail's detail links. From there the 'base_info_collector.py' file will take over.

### Please check out the 'base_info_collector.py' file

# Python script using selenium webDriver to scrape base.com (the info collector script)

I have used **pyautogui, time, requests, Selenium  WebDriver, xlsxwriter and xlrd** libraries. One have to set the excel file route which is containing the links which was derived by the 'base_link_collector.py'.
```python
file_location = "C:/Users/Masum's Computer/PycharmProjects/base_com/base_links_Blu-ray_comedy.xlsx"
```
It goes to every link of the excel file and retrives the below mentioned data.

```python
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
```

New output excel file's name needs to be set manually here.

```python
outputWorkbook = xlsxwriter.Workbook("base_comedy.xlsx")
```
