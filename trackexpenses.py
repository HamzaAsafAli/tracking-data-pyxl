import openpyxl
from openpyxl.styles import Font

try:
    from openpyxl.cell import get_column_letter, column_index_from_string
except ImportError:
    from openpyxl.utils import get_column_letter, column_index_from_string


class Solution:
    def setup(self) -> None:
        wb = openpyxl.Workbook()
        sheet = wb.active
        # will utilize this font obj for the headings
        fontObj1 = Font(name='Times New Roman', bold=True)

        # makes the column wider
        sheet.column_dimensions['A'].width = 20
        sheet.column_dimensions['B'].width = 20
        sheet.column_dimensions['C'].width = 20
        sheet.column_dimensions['D'].width = 20
        sheet.column_dimensions['E'].width = 20

        # set up the basic structure
        sheet.title = 'Expenses'
        sheet['A1'].font = fontObj1
        sheet['A1'] = 'Spend Categories'
        sheet['A2'] = 'Transportation'
        sheet['A3'] = 'Retail & Groceries'
        sheet['A4'] = 'Entertainment'
        sheet['A5'] = 'Restaurants'
        sheet['B1'].font = fontObj1
        sheet['B1'] = "Transactions"
        sheet['C1'].font = fontObj1
        sheet['C1'] = "Amount ($)"
        sheet['D1'].font = fontObj1
        sheet['D1'] = "Budget ($)"
        sheet['E1'].font = fontObj1
        sheet['E1'] = "Difference ($)"

        wb.save('track_temp.xlsx')

    def input(self) -> None:
        self.setup()
        # ask for input here, then feed into track function
        print("Enter digit for the spend category: (1) Transportation, (2) Retail & Groceries, (3) Entertainment and (4) Restaurants ")
        category = input("Enter corresponding spend category digit: ")
        if category == "1":
            print("You are entering expenses for Transportation")
            check = input("Proceed Y/N? ")
            if check == "Y" or check == "y":
                self.transactions(category)
            self.input()
        elif category == "2":
            print("You are entering expenses for Retail & Groceries")
            check = input("Proceed Y/N? ")
            if check == "Y" or check == "y":
                self.transactions(category)
            self.input()
        elif category == "3":
            print("You are entering expenses for Restaurants")
            check = input("Proceed Y/N? ")
            if check == "Y" or check == "y":
                self.transactions(category)
            self.input()
        elif category == "4":
            print("You are entering expenses for Entertainment")
            check = input("Proceed Y/N? ")
            if check == "Y" or check == "y":
                self.transactions(category)
            self.input()
        else:
            print("invalid entry")
            self.input()

    def transactions(self, category) -> None:
        transnum = input("Enter number of transactions: ")
        amount = input("Enter amount: ")
        budget = input("Enter your budget: ")
        self.track(category, transnum, amount, budget)


    def track(self, category, transnum, amount, budget) -> None:
        print(category + transnum + amount + budget)



a = Solution()
a.input()





