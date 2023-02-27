# import csv
# import xlrd

# testArr = ['OR-EE-01-01-F', 'OR-EE-04-05-F', 'OR-EE-10-05-F', 'OR-EE-11-04-F']
# equalsArr = []
# def get_links():
#     wb = xlrd.open_workbook("gvLocations.xls")
#     sh = wb.sheet_by_index(0)
#     ids = []
#     hyperlinks = []
#     locations = []


#     for row in range(sh.nrows):
#         rowValues = sh.row_values(row, start_colx=0, end_colx=2)
#         loc_name = rowValues[0]
#         link = sh.hyperlink_map.get((row, 0))
#         url = '(No URL)' if link is None else link.url_or_path
#         hyperlinks.append(url)
#         locations.append(loc_name)
#     hyperlinks.pop(0)
#     locations.pop(0)

#     for i, j in enumerate(testArr):
#         for k, l in enumerate(locations):
#             if testArr[i] == locations[k]:
#                 equalsArr.append(locations[k])
# get_links()

# print(equalsArr)

arr = ['hello', 'there', 'friends']
helperArr = []
anotherarr = []
for item, x in enumerate(arr):
    helperArr.append(x)
print(helperArr[0])




