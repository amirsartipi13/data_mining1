# Reading an excel file using Python
# import xlsxwriter module
import xlrd
import xlsxwriter
import matplotlib.pyplot as plt
import Apriori as ap

# Give the location of the file
loc = ("C:\\Users\\amir\\Desktop\\Data Mining\\Online_Shopping1.xlsx")

# read data set
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# InvoiceNo_Item Dic
InvoiceNo_Item = {}
for row in range(1, sheet.nrows):
    if sheet.cell_value(row, 0) not in InvoiceNo_Item:
        InvoiceNo_Item[sheet.cell_value(row, 0)] = [sheet.cell_value(row, 2)]
    else:
        InvoiceNo_Item[sheet.cell_value(row, 0)].append(sheet.cell_value(row, 2))

# Create xlsx
workbook = xlsxwriter.Workbook('VoiceNo_Item.xlsx')
worksheet = workbook.add_worksheet("My sheet")

# Write InvoiceNo And Item Into xlsx .
row = 0
col = 0
for InvoiceNo in InvoiceNo_Item:
    worksheet.write(row, col, InvoiceNo)
    for item in InvoiceNo_Item[InvoiceNo]:
        col += 1
        worksheet.write(row, col, item)
    col = 0
    row += 1
workbook.close()

# Find Items Number
Item_Number = {}
Total_Item = 0
for row in range(1, sheet.nrows):
    Total_Item += sheet.cell_value(row, 3)
    if sheet.cell_value(row, 2) not in Item_Number:
        Item_Number[sheet.cell_value(row, 2)] = sheet.cell_value(row, 3)
    else:
        Item_Number[sheet.cell_value(row, 2)] += sheet.cell_value(row, 3)

print(Item_Number)
print(Total_Item)

# Make Item_Frequency And Names List
Names = list(Item_Number.keys())
Item_Frequency = []
for item in Item_Number:
    Item_Frequency.append(Item_Number.get(item)/Total_Item)

# Make Categorical Plot
fig, axs = plt.subplots()
axs.bar(Names, Item_Frequency)
fig.suptitle('Item_Frequency')
plt.show()

# L, Sup_List = ap.apriori(InvoiceNo_Item, 0.001)
