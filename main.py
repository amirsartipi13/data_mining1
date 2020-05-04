import xlrd
import xlsxwriter
import matplotlib.pyplot as plt
import Apriori as ap


# read data set
def find_transactions(loc=None, index_sheet=0 ,  sheet=None , save_loc="new_file.xlsx" , save_sheet="new_sheet"):
    """question 1
    this function find transactions from a sheet and save to a new xlsx file
    inputs:
        loc:location of xlsx file
        index_sheet:sheet number of input xlsx file
        sheet:sheet of informations (at least one of the loc or sheet parameters to be sent)
        save_loc:location of output file
        save_sheet:name of output sheet
    outputs:
        it doasn't have any return value but save a xlsx file of transactions
    """
    if sheet is None:
        if loc is None:
            raise Exception("at least one of the loc or sheet parameters to be sent")
        else:
            wb = xlrd.open_workbook(loc)
            sheet = wb.sheet_by_index(index_sheet)

    # InvoiceNo_Item Dic
    InvoiceNo_Item = {}
    for row in range(1, sheet.nrows):
        if sheet.cell_value(row, 2) != '':
            if (sheet.cell_value(row, 0) not in InvoiceNo_Item):
                InvoiceNo_Item[sheet.cell_value(row, 0)] = [sheet.cell_value(row, 2)]
            else:
                InvoiceNo_Item[sheet.cell_value(row, 0)].append(sheet.cell_value(row, 2))

    # Create xlsx
    workbook = xlsxwriter.Workbook(save_loc)
    worksheet = workbook.add_worksheet(save_sheet)

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
    return InvoiceNo_Item


def find_items_count(loc=None , index_sheet=0 ,  sheet=None , save_fig="Item_Frequency" , number_of_best_item=1):
    """question2
    this function find count of each items and print them  and save plot of it
    inputs:
        loc:location of xlsx file
        index_sheet:sheet number of input xlsx file
        sheet:sheet of informations (at least one of the loc or sheet parameters to be sent)
    outputs:
        it doasn't have any return value but print items' count and best seller items and save items plot and best
         seller items plot
    """

    if sheet is None:
        if loc is None:
            raise Exception("at least one of the loc or sheet parameters to be sent")
        else:
            wb = xlrd.open_workbook(loc)
            sheet = wb.sheet_by_index(index_sheet)

    # Find Items Number
    Item_Number = {}
    Total_Item = 0
    for row in range(1, sheet.nrows):
        Total_Item += sheet.cell_value(row, 3)
        if sheet.cell_value(row, 2) not in Item_Number:
            Item_Number[sheet.cell_value(row, 2)] = sheet.cell_value(row, 3)
        else:
            Item_Number[sheet.cell_value(row, 2)] += sheet.cell_value(row, 3)


    print("total item is : " , Total_Item)

    Item_Names = list(Item_Number.keys())
    Item_Frequencies = []
    Item_Number_list = []

    for item in Item_Number:
        Item_Number_list.append((item ,Item_Number.get(item) ))
        Item_Frequencies.append(Item_Number.get(item) / Total_Item)

    best_Item_Names = []
    best_Item_Frequencies = []
    Item_Number_list.sort(key=lambda item:item[1] ,reverse=True) # sort and find best items

    #print Items
    for index , item in enumerate(Item_Number_list):
        print(item[0],"   " , item[1] ,"   ", item[1] / Total_Item)
        if index < number_of_best_item:
            best_Item_Names.append(item[0])
            best_Item_Frequencies.append(item[1]/Total_Item)

    # save figure of items
    fig, axs = plt.subplots(figsize=(250 , 50))
    plt.xticks(rotation=60 , fontsize=20)

    axs.bar(Item_Names, Item_Frequencies)
    fig.suptitle('Item_Frequency', fontsize=30)
    plt.savefig(save_fig)

    # save figure of best items

    fig, axs = plt.subplots(figsize=(100, 50) )
    plt.xticks(rotation=60 ,fontsize=20)
    axs.bar(best_Item_Names, best_Item_Frequencies)
    fig.suptitle('best Item_Frequency' , fontsize=30)
    plt.savefig("best "+save_fig)


if __name__ == '__main__':
    loc = ("test.xlsx")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    Invoice_Item = find_transactions(sheet=sheet, save_loc="VoiceNo_Item.xlsx" , save_sheet="transactions")
#    find_items_count(sheet=sheet , number_of_best_item=10)
    ap.apriori(Invoice_Item.values(), 0.6, 0.03)
    print("finish")
