from xlwings import Workbook, Range

wb = Workbook()
Range('A1').value = 1000    # write operation
Range((2,4)).value = "'x = 2, y = 4 (D4)"

# now you can change the value in the excel sheet and press enter here
wait = input("change value in sheet opened, Press enter to continue")

x = Range('A1').value
print("New value in A1 is {0}".format(x))

# works starting version 0.3.5
# row1 = Range((2,4)).row
# col1 = Range((2,4)).column
# print("Range((2,4)) refers to row {0} and column {1}".format((row1,col1)))
