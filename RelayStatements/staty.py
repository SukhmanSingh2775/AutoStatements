import xlrd
import openpyxl



 

path = ("/Users/sukhmansingh/Downloads/ExcelStatement.xlsx")

wb = xlrd.open_workbook(path)
sheet = wb.sheet_by_index(1)

load_list_1 = sheet.col_values(1)
load_list_2 = sheet.col_values(2)


prices_column = sheet.col_values(21)

load_list_1 = [i[-4:] for i in load_list_1  if i != "Trip ID"] # Round Trips Loads
load_list_2 = [i[-4:] for i in load_list_2 if i != "Load ID"] # One Way Loads
prices_column = [i for i in prices_column if type(i) != str] # Getting float only for prices
prices_column.pop(-1)

one_way = [load_list_2[i] for i in range(len(load_list_1)) if load_list_1[i] == "" or load_list_1[i] == "-" if load_list_2[i] != ""] #Modified one way loads



prices_one_way = [prices_column[i] for i in range(len(prices_column)) if load_list_1[i] == "" or load_list_1[i] == "-"]
prices_round_trip = [prices_column[i] for i in range(len(prices_column)) if load_list_1[i] and load_list_1[i] != "-"]

load_id = ['SQO1', 'RT9Q', '5Q1R', 'GJD9']
prices = [104.09, 1052, 255.30, 3943.3]


for i in range(len(load_id)):
    if load_id[i] in one_way:
        print("Found it")
        print(load_id[i])
# 


# print(one_way)
# print(round_trip)

# array = ['Sukhman', 'Singh']

# array = [i[2:]  for i in array if i]
# array = [i[:-1] for i in array if i]

# print(array)


# strs = 'Sukhman'

# print(strs[4:]) Remove first 4 MAN
# print(strs[:4])  Get only first 4 SUKH

# print(strs[:-4]) # Removes last 4

# print(strs[-4:]) #Gets last 4






