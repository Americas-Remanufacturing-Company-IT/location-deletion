import time
from tkinter import *
from tkinter import ttk, filedialog
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
import chromedriver_autoinstaller
import csv
import xlrd


testArr = []
# check chrome stuff
print('Checking Chrome version')
# see if chrome is installed
try:
    currentversion = chromedriver_autoinstaller.utils.get_chrome_version()
except IndexError as e:
    print('Chrome not detected, please install chrome')
    exit()

# if it is, print the version
print(f'Chrome version: {currentversion}')
# make chromedriver path
cdpath = f'{os.getcwd()}\\chromedriver\\'
if not os.path.isdir(cdpath):
    os.mkdir(cdpath)
# install chromedriver for current version
chromedriver_autoinstaller.install(path=cdpath)
cwd = os.getcwd()
options = Options()
prefs = {"download.default_directory" : cwd}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=options)
"""Get users login credential using input"""
root = Tk()
root.geometry("400x200")

entry1 = Entry(root, width=50)
entry1.pack()
entry2 = Entry(root, width=50)
entry2.pack()

options = StringVar()
options.set("Select Option...")

OptionMenu(root, options, "Arc", "Amazon", "Almo","Basco", "BestBuy", "CB", "Galanz", "Hurom", "NewAir", "SB", "TTI", "TLC", "Tineco" ).pack()

def file_select():
    Tk().withdraw()
    file = filedialog.askopenfile(mode='r', filetypes=[('CSV Files', '*.csv')])
    if file:
        global file_name
        filepath = os.path.abspath(file.name)
        file_name = str(filepath)


file_select()

def button_command():
    username = entry1.get()
    password = entry2.get()
    domain = options.get()
    domain_url = f"https://{domain}.arcaugusta.com/"
    driver.get(domain_url)
    driver.find_element(by=By.ID, value="txtUsername").send_keys(username)
    driver.find_element(by=By.ID, value="txtPassword").send_keys(password)
    driver.find_element(by=By.ID, value="btnLogin").click()
    location_page = domain_url + "report/Locations.aspx"
    driver.get(location_page)
    filter()
    get_links()
    os.remove('gvLocations.xls')
    driver.quit()
    exit()


Button(root, text="Login", command=button_command).pack()


def filter():
    """Somewhere in here should be a check to make sure each element was clicked and tells where the error is"""
    wait = WebDriverWait(driver, 15)
    # wait.until(EC.invisibility_of_element_located((By.XPATH, "//*[@id='ctl00_ContentPlaceHolder1_gvLocations_DXPFCForm_DXPWMB-1']")))
    wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnXlsExport")))
    driver.find_element(by=By.ID, value="ctl00_ContentPlaceHolder1_btnXlsExport").click()



def get_links():
    helperArr = []
    final_locations = []
    with open(file_name, "r") as csv_file:
        csv_reader = csv.reader(csv_file)
        for line in csv_reader:
            location_grab = line[0]
            helperArr.append(location_grab)
        repeated = list(set(helperArr))
        repeated.sort()
        for items in repeated:
            final_locations.append(items)

    filename = "gvLocations.xls"
    wait_time = 5

    while not os.path.exists(filename):
        print("File not found. Waiting for", wait_time, "seconds...")
        time.sleep(wait_time)

    wb = xlrd.open_workbook("gvLocations.xls")
    sh = wb.sheet_by_index(0)
    ids = []
    hyperlinks = []
    locations = []

    for row in range(sh.nrows):
        rowValues = sh.row_values(row, start_colx=0, end_colx=2)
        loc_name = rowValues[0]
        link = sh.hyperlink_map.get((row, 0))
        url = '(No URL)' if link is None else link.url_or_path
        hyperlinks.append(url)
        locations.append(loc_name)
    hyperlinks.pop(0)
    locations.pop(0)

    for i, j in enumerate(hyperlinks):
        if hyperlinks[i].__class__.__name__ == "bytes":
            converted = hyperlinks[i].decode('ASCII')
        else:
            converted = hyperlinks[i]
        test_loc = converted
        find_sign = test_loc.find('=')
        location_id = test_loc[find_sign + 1:]
        ids.append(location_id)
    # print(ids)
    # print(locations)

    myArr = []
    resultArr = []

    for i, j in enumerate(locations):
        myArr.append([ids[i], locations[i], 'FALSE', 'Other', locations[i], 'FALSE', 'FALSE', 'FALSE', 'FALSE', 'FALSE', 'FALSE'])

    for arr in myArr:
        for x, y in enumerate(final_locations):
            if y in arr:
                resultArr.append(arr)
    print(resultArr)


    f = filedialog.asksaveasfile(
        defaultextension='.csv',
        filetypes=[
            ('CSV file', '.csv'),
            ('Text file', '.txt'),
            ('All files', '.*'),
        ])


    writer = csv.writer(f)
    writer.writerow(['LocationID', 'NewName', 'Active', 'LocationType', 'NewDescription', 'IsShipping',
                     'IsReceiving', 'IsRepair', 'IsHarvest', 'IsHarvestParentMove', 'IsHarvestComponentMove'])
    
    for x in resultArr:
        writer.writerow(x)


    # for x, y in enumerate(final_locations):
    #     secondArr = [arr for arr in myArr if y in arr][0]
    #     writer.writerow(secondArr)

root.mainloop()

"""Have user select which domain to go to like Basco or Amazon and selenium will login"""

"""We need to be able to go to reports and click on Locations and hit go"""

"""Click on the filter at the bottom"""

"""Change the and to Or"""

"""Make sure it's set to Name contains"""

"""We need to make sure that the spreadsheet has only the locations we need for deactivation"""

"""Find a way to access the column with location names"""

"""Get the amount of columns and loop through each column and add each one to the filter before going to the next"""
# Should add a feature that clicks on the side because the filter won't add until another click has been made.

"""There should be a function that only loops the locations needed like if there is 
IB-ASL-08-01, IB-ASL-08-02, IB-ASL-08-03, etc it should filter for the first one only because the
rest are included."""
"""Once the filter is created click export XLS"""

"""After the xls is exported we need to find a way to open it and run Ben's vbs or something similar in Python."""

"""Once that is ran we need to be able to use the locationrename template"""

"""Copy the IDs from one excel to another and then copy the description to name and desc
 and turn everything else FALSE"""

"""When finished save the file to somewhere else"""

"""We would create a dialog to open the Excel file with the locations that needs to be deleted, create a funciton
to save the download from export to XLS and convert it to an XLSX file and then when completely finished
create a dialog to save it as an csv file anywhere."""
