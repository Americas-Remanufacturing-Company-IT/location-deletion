import time
from tkinter import *
from tkinter import ttk, filedialog
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import chromedriver_autoinstaller
import csv
import xlrd


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
driver.implicitly_wait(20)
"""Get users login credential using input"""
root = Tk()
root.geometry("400x200")

entry1 = Entry(root, width=50)
entry1.pack()
entry2 = Entry(root, width=50)
entry2.pack()

options = StringVar()
options.set("Select Option...")

OptionMenu(root, options, "Arc", "Amazon", "Basco").pack()

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
    return None


Button(root, text="Login", command=button_command).pack()


def filter():
    trial = 'Trial'
    """Somewhere in here should be a check to make sure each element was clicked and tells where the error is"""
    driver.find_element(by=By.CLASS_NAME, value="dxgvFilterBarLink_Glass").click()
    time.sleep(5)
    driver.find_element(by=By.XPATH, value="//a[@class='dxfcGroupType_Glass']").click()
    time.sleep(2)
    driver.find_element(by=By.XPATH, value="//*[@id='ctl00_ContentPlaceHolder1_gvLocations_DXPFCForm_DXPFC_GroupPopup_DXI1_T']/span").click()
    """Bellow needs to be put in a loop and there needs to be a ghost click somewhere"""
    for item, x in enumerate(location_grouping()):
        if item > 0:
            insert = f'[{item + 1}]'
        else:
            insert = ''
        type_value = f'{item + 1}'
        time.sleep(2)
        driver.find_element(by=By.CLASS_NAME, value="dxEditors_fcadd_Glass").click()
        time.sleep(2)
        #Date Exported to Name
        time.sleep(1.5)
        driver.find_element(by=By.XPATH, value=f"//*[@id='ctl00_ContentPlaceHolder1_gvLocations_DXPFCForm_DXPFC']/tbody/tr/td/ul/li/ul/li{insert}/table/tbody/tr/td[2]/table/tbody/tr/td[1]/a").click()
        time.sleep(1.5)
        driver.find_element(by=By.XPATH, value="//*[@id='ctl00_ContentPlaceHolder1_gvLocations_DXPFCForm_DXPFC_FieldNamePopup_DXI4_T']/span").click()
        time.sleep(1.5)
        #Begins with to Contains
        driver.find_element(by=By.XPATH, value=f"//*[@id='ctl00_ContentPlaceHolder1_gvLocations_DXPFCForm_DXPFC']/tbody/tr/td/ul/li/ul/li{insert}/table/tbody/tr/td[2]/table/tbody/tr/td[2]/a").click()
        time.sleep(1.5)
        driver.find_element(by=By.XPATH, value="//*[@id='ctl00_ContentPlaceHolder1_gvLocations_DXPFCForm_DXPFC_OperationPopup_DXI8_T']/span").click()
        time.sleep(1.5)
        #Types in value
        driver.find_element(by=By.XPATH, value=f"//*[@id='ctl00_ContentPlaceHolder1_gvLocations_DXPFCForm_DXPFC_DXValue{type_value}000']").click()
        time.sleep(2)
        driver.find_element(by=By.XPATH, value=f"//*[@id='ctl00_ContentPlaceHolder1_gvLocations_DXPFCForm_DXPFC_DXEdit{type_value}000_I']").send_keys(x)
        time.sleep(2)
        driver.find_element(by=By.CLASS_NAME, value="dxEditors_fcadd_Glass").click()
    time.sleep(2)
    driver.find_element(by=By.CLASS_NAME, value="dxbButton_Glass").click()
    time.sleep(2)
    driver.find_element(by=By.ID, value="ctl00_ContentPlaceHolder1_btnXlsExport").click()
    # time.sleep(15)
    # driver.quit()


def location_grouping():
    with open(file_name, "r") as csv_file:
        csv_reader = csv.reader(csv_file)

        locations = []
        final_locations = []
        for line in csv_reader:
            location_grab = line[0]
            location_split = location_grab.split('-')
            if len(location_split) < 3:
                locations.append('-'.join(location_split))
            elif len(location_split) == 3:
                del location_split[-1]
                locations.append('-'.join(location_split))
            elif len(location_split) > 4:
                del location_split[-1]
                del location_split[-1]
                del location_split[-1]
                locations.append('-'.join(location_split))
            # elif len(location_split) > 2 and location_split[-1].isdigit():
            elif len(location_split) == 4:
                del location_split[-1]
                del location_split[-1]
                locations.append('-'.join(location_split))
            else:
                print('Something went wrong')
                exit()
        repeated = list(set(locations))
        repeated.sort()
        for items in repeated:
            final_locations.append(items)
    return final_locations

def get_links():
    time.sleep(3)
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
    print(ids)
    print(locations)

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
    for i, j in enumerate(locations):
        writer.writerow(
            [ids[i], locations[i], 'FALSE', 'Other', locations[i], 'FALSE', 'FALSE', 'FALSE', 'FALSE', 'FALSE', 'FALSE'])

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
