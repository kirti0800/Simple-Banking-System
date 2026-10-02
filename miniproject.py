import openpyxl

workbook = openpyxl.Workbook()
sheet = workbook.active

sheet["A1"] = "Account No"
sheet["B1"] = "Name"
sheet["C1"] = "Balance"

sheet.append([101, "Kirti", 5000])

workbook.save("bank.xlsx")

print("Data saved in Excel!")