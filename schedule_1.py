import random
from itertools import combinations
from openpyxl import Workbook

# Define the work schedule matrix
schedule = [[0] * 30 for i in range(5)]

# Populate the matrix with random work schedules
for j in range(30):
    available_employees = list(range(5))
    random.shuffle(available_employees)
    for c in combinations(available_employees, 3):
        if all(schedule[k][j] == 0 for k in c):
            for i in c:
                schedule[i][j] = 1
            break

# Create an Excel workbook and sheet
workbook = Workbook()
sheet = workbook.active

# Write the work schedule matrix to the sheet
for i in range(5):
    for j in range(30):
        sheet.cell(row=i + 1, column=j + 1, value=schedule[i][j])

# Save the workbook to a file
workbook.save("work_schedule.xlsx")
