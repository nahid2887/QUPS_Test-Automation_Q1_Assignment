import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Sunday"

ws.append(["Keyword", "Longest Option", "Shortest Option"])

data = [
    ("Keyword1", "Dhaka"),
    ("Keyword2", "Saturday"),
    ("Keyword3", "Baby"),
    ("Keyword4", "School"),
    ("Keyword5", "Cricket"),
    ("Keyword6", "Momey"),
    ("Keyword7", "Int"),
    ("Keyword8", "Look"),
    ("Keyword9", "Hello"),
    ("Keyword10", "By")
]

for row in data:
    ws.append([row[0], "", "", row[1]])
wb.save("keywords.xlsx")

print("Excel file 'keywords.xlsx' created successfully!")
