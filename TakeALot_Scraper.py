from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from time import sleep
import tkinter as tk
import xlsxwriter


webdriver = "C:/Users/Brendan/Documents/Python Projects/HelloCoding/chromedriver.exe"
options = Options()


def searchWebsite():

    search_entry = ''
    item_name = []
    item_price = []
    formatted_price = []

# - Grabs the text entered into Tkinter, adds headless arguments to Chrome's webdriver - #

    search_entry = entry1.get()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    driver = Chrome(webdriver, options=options)

# - Opens Takealot, searches the keywords and creates a list of items on the first results page - #

    driver.get('https://www.takealot.com/')
    sleep(8)
    elem = driver.find_element_by_class_name('search-field ')
    elem.send_keys(search_entry)
    elem.send_keys(Keys.RETURN)
    sleep(7)
    search_by_item = driver.find_elements_by_class_name('result-item')

# - Gets the item name and price for each item on the first results page - #

    for item in search_by_item:
        get_name = item.find_element_by_id('pos_link_0').text
        item_name.append(get_name)
        get_price = item.find_element_by_class_name('amount').text
        item_price.append(get_price)
        sleep(3)

    for price in item_price:
        formatted_price.append('R ' + price)

    driver.close()
    master.quit()

    my_dict = dict(zip(item_name, formatted_price))

# - Creates a .xlsx sheet and does some basic formatting - #

    workbook = xlsxwriter.Workbook('SCRAPE_DATA.xlsx')
    cell_format = workbook.add_format({'bold':True, 'size':16})
    cell_format.set_center_across('center_across')
    cell_format2 = workbook.add_format({'align':'right'})
    worksheet = workbook.add_worksheet()

    row = 1
    col = 0

    worksheet.write(row - 1, col, 'ITEM',cell_format)
    worksheet.write(row - 1, col + 1, 'PRICE',cell_format)
    worksheet.set_column('A:A', 1)
    worksheet.set_column('A:A', 50)

# - Writes the information to the above created sheet - #

    for k, v in my_dict.items():
        worksheet.write(row, col, k)
        worksheet.write(row, col + 1, v, cell_format2)
        row += 1

    workbook.close()

# - This creates the window, text box and button that initiates the searchWebsite def - #

master = tk.Tk()

master.title('Takealot Search')

canvas = tk.Canvas(master, width=300, height=200)
canvas.pack()

label1 = tk.Label(master, text='Creates an Excel sheet with the first page results.')
label2 = tk.Label(master, text='Enter your search below (and please be patient)')
entry1 = tk.Entry(master)
button1 = tk.Button(master, text='Go', command=searchWebsite)

canvas.create_window(150, 30, window=label1)
canvas.create_window(150, 60, window=label2)
canvas.create_window(150, 100, window=entry1)
canvas.create_window(150, 140, window=button1)


master.mainloop()


######################################################################################





