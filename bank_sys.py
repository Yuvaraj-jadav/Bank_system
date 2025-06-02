import openpyxl as xl # Importing the openpyxl library to handle Excel files
file = xl.load_workbook('bank_accounts.xlsx')  # Load the Excel file
sheet = file.active

# cell_obj = sheet.cell(row=1, column=1)
class bank_account:
    

    def __init__(self,name,acc_no,ifsc_code,bank_balance):
        self.name=f"name of the bank holder: {name}\n"             #instance attributes of object
        self.acc_no=f"Account number of bank holder: {acc_no}\n"
        self.ifsc_code=f"IFSC code of bank: {ifsc_code}\n"
        self.bank_balance=float(bank_balance)
    def dep(self,deposit):     #methods of object
         self.bank_balance += deposit
         c1=sheet.cell(row=2, column=4)  # Accessing the cell in the Excel sheet
         c1.value = self.bank_balance
         file.save('bank_accounts.xlsx')
    def wd(self,withdraw):     #methods of object
        self.bank_balance-= withdraw
        c1=sheet.cell(row=2, column=4)
        c1.value = self.bank_balance
        file.save('bank_accounts.xlsx')
        if self.bank_balance < 0:
            print("Insufficient balance")
            self.bank_balance += withdraw # Prevent overdraft
            c1.value = self.bank_balance # Update the cell value in the Excel sheet
    def cb(self):              #methods of object
        print(self.bank_balance)
    def display(self):
        print(self.name, self.acc_no, self.ifsc_code, self.bank_balance)


n = sheet.cell(row=2, column=1).value  
acc = sheet.cell(row=2, column=2).value  # Read the first line from the file
ifsc =sheet.cell(row=2, column=3).value
bl= sheet.cell(row=2,column=4).value # Read the IFSC code from the file
obj1 = bank_account(n,acc, ifsc,bl)
obj1.dep(2000)  # Deposit 1000
obj1.display() 
