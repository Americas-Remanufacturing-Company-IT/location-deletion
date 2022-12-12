import csv
from tkinter import *
from tkinter import ttk, filedialog
import os


def file_select():
    Tk().withdraw()
    file = filedialog.askopenfile(mode='r', filetypes=[('CSV Files', '*.csv')])
    if file:
        global file_name
        filepath = os.path.abspath(file.name)
        file_name = str(filepath)

file_select()

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

print(location_grouping())