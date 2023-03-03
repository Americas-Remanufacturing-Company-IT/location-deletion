import os
import csv
import xlrd
import time
from tkinter import *
from tkinter import ttk, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
import chromedriver_autoinstaller
import PySimpleGUI as sg


sg.theme('BlueMono')

screen_width, screen_height = sg.Window.get_screen_size()
POPUP_CENTER_X = int(screen_width / 2)
POPUP_CENTER_Y = int(screen_height / 10)

# Define domain list
domain_list = ["Arc", "Amazon", "Almo", "Basco", "BestBuy", "CB", "Galanz", "Hurom", "NewAir", "SB", "TTI", "TLC", "Tineco"]

# Define GUI layout
layout = [
    [sg.Text('UserName:', size=(20, 1)), sg.Input(key='-UserName-')],
    [sg.Text('Password:', size=(20, 1)), sg.Input(key='-Password-', password_char='*')],
    [sg.Input(key='-File-', readonly=True, disabled_readonly_background_color='light gray'), sg.FileBrowse(file_types=(('CSV Files', '*.csv'),))],
    [sg.Text('Domain', size=(20,1))],
    [sg.Combo(domain_list, key='-Domain-', readonly=True)],
    [sg.Button('Login'), sg.Exit()]
]

# Create GUI window
window = sg.Window('Location Deletion', layout, keep_on_top=True)

def main():
    while True:

        event, values = window.read()

        username = values['-UserName-']
        password = values['-Password-']
        file_name = values['-File-']
        domain = values['-Domain-']
        domain_url = f"https://{domain}.arcaugusta.com/"
        filename = 'gvLocations.xls'
        wait_time = 5

        if event in (sg.WINDOW_CLOSE_ATTEMPTED_EVENT, 'Exit'):
            break
        elif event == 'Login':
            # Check if all fields are filled
            err = False
            for k, v in {'-UserName-' : 'UserName', '-Password-' : 'Password', '-File-' : 'File', '-Domain-' : 'Domain'}.items():
                if not values[k]:
                    sg.popup(f'Please Enter {v}', keep_on_top=True, relative_location=(POPUP_CENTER_X, POPUP_CENTER_Y))
                    err = True
                    break
            if err:
                continue

            # Check chrome version and install chromedriver
            try:
                currentversion = chromedriver_autoinstaller.utils.get_chrome_version()
            except IndexError as e:
                print('Chrome not detected, please install chrome')
                exit()

            print(f'Chrome version: {currentversion}')
            cdpath = f'{os.getcwd()}\\chromedriver\\'
            if not os.path.isdir(cdpath):
                os.mkdir(cdpath)
            chromedriver_autoinstaller.install(path=cdpath)
            cwd = os.getcwd()
            options = Options()
            prefs = {"download.default_directory" : cwd}
            options.add_experimental_option("prefs", prefs)
            driver = webdriver.Chrome(options=options)

            def end():
                driver.quit()
                exit()

            # Login to website
            driver.get(domain_url)
            driver.find_element(by=By.ID, value="txtUsername").send_keys(username)
            driver.find_element(by=By.ID, value="txtPassword").send_keys(password)
            driver.find_element(by=By.ID, value="btnLogin").click()
            location_page = domain_url + "report/Locations.aspx"
            driver.get(location_page)

            # Download location report
            if os.path.exists(filename):
                os.remove(filename)

            wait = WebDriverWait(driver, 15)
            wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnXlsExport")))
            driver.find_element(by=By.ID, value="ctl00_ContentPlaceHolder1_btnXlsExport").click()

            # Extract location data from CSV file
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

            # Wait for location report to download
            while not os.path.exists(filename):
                print("File not found. Waiting for", wait_time, "seconds...")
                time.sleep(wait_time)

            # Extract data from location report
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

            for i in range(len(hyperlinks)):
                if hyperlinks[i].__class__.__name__ == "bytes":
                    converted = hyperlinks[i].decode('ASCII')
                else:
                    converted = hyperlinks[i]
                test_loc = converted
                find_sign = test_loc.find('=')
                location_id = test_loc[find_sign + 1:]
                ids.append(location_id)

            # Filter location data based on input file
            myArr = [[ids[i], locations[i], 'FALSE', 'Other', locations[i], 'FALSE', 'FALSE', 'FALSE', 'FALSE', 'FALSE', 'FALSE'] for i in range(len(locations))]
            resultArr = [arr for arr in myArr if any(x in arr for x in final_locations)]

            # Save filtered location data to output CSV file
            file_path = filedialog.asksaveasfilename(defaultextension='.csv')

            with open(file_path, 'w', newline='') as outfile:
                writer = csv.writer(outfile)
                writer.writerow(['LocationID', 'NewName', 'Active', 'LocationType', 'NewDescription', 'IsShipping',
                                'IsReceiving', 'IsRepair', 'IsHarvest', 'IsHarvestParentMove', 'IsHarvestComponentMove'])

                for x in resultArr:
                    writer.writerow(x)

            sg.popup('Done!', keep_on_top=True, relative_location=(POPUP_CENTER_X, POPUP_CENTER_Y))
            end()

if __name__ == '__main__':
    main()
