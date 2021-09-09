import os

import openpyxl
from openpyxl.styles import Font

try:
    from openpyxl.cell import get_column_letter, column_index_from_string
except ImportError:
    from openpyxl.utils import get_column_letter, column_index_from_string


class Solution:
    def __init__(self):
        if os.path.isfile('track_temp.xlsx'):
            self.wb = openpyxl.load_workbook('track_temp.xlsx')
            self.sheet = self.wb['Expenses']
        else:
            self.wb = openpyxl.Workbook()
            self.sheet = self.wb.active
            fontObj1 = Font(name='Times New Roman', bold=True)
            # makes the column wider
            self.sheet.column_dimensions['A'].width = 20
            self.sheet.column_dimensions['B'].width = 20
            self.sheet.column_dimensions['C'].width = 20
            self.sheet.column_dimensions['D'].width = 20
            self.sheet.column_dimensions['E'].width = 20
            # set up the basic structure
            self.sheet.title = 'Expenses'
            self.sheet['A1'].font = fontObj1
            self.sheet['A1'] = 'Spend Categories'
            self.sheet['A2'] = 'Transportation'
            self.sheet['A3'] = 'Retail & Groceries'
            self.sheet['A4'] = 'Entertainment'
            self.sheet['A5'] = 'Restaurants'
            self.sheet['B1'].font = fontObj1
            self.sheet['B1'] = "Transactions"
            self.sheet['C1'].font = fontObj1
            self.sheet['C1'] = "Amount ($)"
            self.sheet['D1'].font = fontObj1
            self.sheet['D1'] = "Budget ($)"
            self.sheet['E1'].font = fontObj1
            self.sheet['E1'] = "Difference ($)"
            self.wb.save('track_temp.xlsx')







    # def setup(self) -> None:
    #     # wb = openpyxl.Workbook()
    #     # sheet = wb.active
    #     # global wb
    #     # global sheet
    #     # wb = openpyxl.Workbook()
    #     # sheet = wb.active
    #     # will utilize this font obj for the headings
    #     fontObj1 = Font(name='Times New Roman', bold=True)
    #
    #     # makes the column wider
    #     sheet.column_dimensions['A'].width = 20
    #     sheet.column_dimensions['B'].width = 20
    #     sheet.column_dimensions['C'].width = 20
    #     sheet.column_dimensions['D'].width = 20
    #     sheet.column_dimensions['E'].width = 20
    #
    #     # set up the basic structure
    #     sheet.title = 'Expenses'
    #     sheet['A1'].font = fontObj1
    #     sheet['A1'] = 'Spend Categories'
    #     sheet['A2'] = 'Transportation'
    #     sheet['A3'] = 'Retail & Groceries'
    #     sheet['A4'] = 'Entertainment'
    #     sheet['A5'] = 'Restaurants'
    #     sheet['B1'].font = fontObj1
    #     sheet['B1'] = "Transactions"
    #     sheet['C1'].font = fontObj1
    #     sheet['C1'] = "Amount ($)"
    #     sheet['D1'].font = fontObj1
    #     sheet['D1'] = "Budget ($)"
    #     sheet['E1'].font = fontObj1
    #     sheet['E1'] = "Difference ($)"
    #
    #
    #
    #
    #     #everything below as a result of track function returns
    #     # if category == "1":
    #     #     # working with columns B - E, row 2
    #     #     #sheet['B2'] = transnum
    #     #     sheet['B2'].value = transnum
    #     #     wb.save('track_temp.xlsx')
    #     # elif category == "2":
    #     #     # working with columns B - E, row 3
    #     #     sheet['B3'] = transnum
    #     #     wb.save('track_temp.xlsx')
    #     # elif category == "3":
    #     #     # working with columns B - E, row 4
    #     #     sheet['B4'] = transnum
    #     #     category = 0
    #     # else:
    #     #     # working with columns B - E, row 5
    #     #     sheet['B5'] = transnum
    #     #     category = 0
    #
    #     # if category == "1":
    #     #     # working with columns B - E, row 2
    #     #     sheet.cell(row=2, column=2).value = input("Enter number of transactions: ")
    #     #     #sheet['B2'] = input("Enter number of transactions: ")
    #     # elif category == "2":
    #     #     # working with columns B - E, row 3
    #     #     sheet['B3'] = input("Enter number of transactions: ")
    #     # elif category == "3":
    #     #     # working with columns B - E, row 4
    #     #     sheet['B4'] = input("Enter number of transactions: ")
    #     # else:
    #     #     # working with columns B - E, row 5
    #     #     sheet['B5'] = input("Enter number of transactions: ")
    #
    #
    #     wb.save('track_temp.xlsx')
    #
    #     #self.updater('track_temp.xlsx', 'Expenses')



    # def input(self) -> None:
    #     # # ask for input here, then feed into updater function
    #     print("Enter digit for the spend category: (1) Transportation, (2) Retail & Groceries, (3) Entertainment "
    #           "and (4) Restaurants ")
    #     category = input("Enter corresponding spend category digit: ")
    #     self.updater('track_temp.xlsx', 'Expenses', category)
    #     # if category == "1":
    #     #     print("You are entering expenses for Transportation")
    #     #     check = input("Proceed Y/N? ")
    #     #     if check == "Y" or check == "y":
    #     #         #self.setup(category)
    #     #         self.updater('track_temp.xlsx', 'Expenses', category)
    #     #         #self.transactions(category)
    #     #     #self.input()
    #     # elif category == "2":
    #     #     print("You are entering expenses for Retail & Groceries")
    #     #     check = input("Proceed Y/N? ")
    #     #     if check == "Y" or check == "y":
    #     #         #self.setup(category)
    #     #         self.updater('track_temp.xlsx', 'Expenses', category)
    #     #         #self.transactions(category)
    #     #     #self.input()
    #     # elif category == "3":
    #     #     print("You are entering expenses for Restaurants")
    #     #     check = input("Proceed Y/N? ")
    #     #     if check == "Y" or check == "y":
    #     #         #self.setup(category)
    #     #         self.updater('track_temp.xlsx', 'Expenses', category)
    #     #         #self.transactions(category)
    #     #     #self.input()
    #     # elif category == "4":
    #     #     print("You are entering expenses for Entertainment")
    #     #     check = input("Proceed Y/N? ")
    #     #     if check == "Y" or check == "y":
    #     #         #self.setup(category)
    #     #         self.updater('track_temp.xlsx', 'Expenses', category)
    #     #         #self.transactions(category)
    #     #     #self.input()
    #     # else:
    #     #     print("invalid entry")
    #     #     #self.input()

    def updater(self):
        # if os.path.isfile(filename):
        #     wb = openpyxl.load_workbook(filename)
        #     sheet = wb[sheetname]
        #     print(sheet['A1'].value)
        # else:
        #     self.__init__()



        # wb = openpyxl.load_workbook(filename)
        # sheet = wb[sheetname]
        #wb.template = False
        #wb.save(filename)
        # print(sheet['A1'].value)
        # print(sheet['B3'].value)
        print("Enter digit for the spend category: (1) Transportation, (2) Retail & Groceries, (3) Entertainment "
              "and (4) Restaurants ")
        category = input("Enter corresponding spend category digit: ")

        # if category == "1":
        #     B2 = input("Enter number of transactions: ")
        #     print(B2)
        #     self.sheet['B2'] = B2
        #     #wb.save(filename)
        # elif category == "2":
        #     B3 = input("Enter number of transactions: ")
        #     print(B3)
        #     self.sheet['B3'] = B3
        #     #wb.save(filename)
        # else:
        #     self.sheet['B4'] = "ok"
        #     print("else")
        #
        # #del category
        #
        # self.wb.save(filename)

        #wb.save(filename)
        # self.sheet['B2'] = input("Enter number of transactions: ")
        # self.sheet['B3'] = input("Enter number of transactions: ")
        #self.wb.save(filename)
        if category == "1":
            # working with columns B - E, row 2
            self.sheet.cell(row=2, column=2).value = input("Enter number of transactions: ")
            #sheet['B2'] = input("Enter number of transactions: ")
        elif category == "2":
            # working with columns B - E, row 3
            self.sheet['B3'] = input("Enter number of transactions: ")
        elif category == "3":
            # working with columns B - E, row 4
            self.sheet['B4'] = input("Enter number of transactions: ")
        elif category == "4":
            # working with columns B - E, row 5
            self.sheet['B5'] = input("Enter number of transactions: ")
        else:
            self.updater()

        self.wb.save('track_temp.xlsx')



    # def transactions(self, category) -> None:
    #     transnum = input("Enter number of transactions: ")
    #     amount = input("Enter amount: ")
    #     budget = input("Enter your budget: ")
    #     self.track(category, transnum, amount, budget)


    # def track(self, category) -> None:
    #     # print(sheet["A1"].value)
    #     # print(sheet['A5'].value)
    #     #     transnum = input("Enter number of transactions: ")
    #     #     amount = input("Enter amount: ")
    #     #     budget = input("Enter your budget: ")
    #     if category == "1":
    #         # working with columns B - E, row 2
    #         sheet['B2'] = input("Enter number of transactions: ")
    #     elif category == "2":
    #         # working with columns B - E, row 3
    #         sheet['B3'] = input("Enter number of transactions: ")
    #     elif category == "3":
    #         # working with columns B - E, row 4
    #         sheet['B4'] = input("Enter number of transactions: ")
    #         category = 0
    #     else:
    #         # working with columns B - E, row 5
    #         sheet['B5'] = input("Enter number of transactions: ")
    #         category = 0
    #
    #     wb.save('track_temp.xlsx')





a = Solution()
a.updater()



