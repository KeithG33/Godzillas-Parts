from openpyxl import load_workbook
import sys

accept_str = False


# First things first, load the sheet
godzilla_parts_workbook = load_workbook('parts_list.xlsx')

if not godzilla_parts_workbook:
    print("Excel sheet didn't load, better figure that out buddy guy")
else: 

    while accept_str == False:
        # Get me inputs
        part_string = input("Ready to go, add part using 'Brand, Part, Purchase $' format:  ")
        part_list = part_string.split(", ")

        # Make sure I didn't feck up
        print("You inputted: \nBrand: " + part_list[0] + "\nPart: " + part_list[1] + "\nPurchase Price: " + part_list[2])

        keep_or_not = input("Do you want to keep this input? Answer y/n\n ")

        if keep_or_not == "y":
            accept_str = True
            break
        elif keep_or_not == "n":
            continue
        elif keep_or_not != "y" or accept_str != "n":
            ("Input error, please type 'y' for yes and 'n' for no...Try again\n ")
        


    brand = part_list[0]
    part = part_list[1]
    price = part_list[2]

    # Got the info we need, time to edit excel sheet. Brand is ColB, part is ColC, price ColD
    ws = godzilla_parts_workbook.active

    MAX_ROW = 20

    for col in range (2,5):
        for row in range(6, MAX_ROW):
            cell = ws.cell(column=col, row=row, value = "")

            if cell.value is None or cell.value == "":
                if col == 2:
                    cell.value = brand
                    break
                elif col == 3:
                    cell.value = part
                    break
                elif col == 4:
                    cell.value = price
                    break

godzilla_parts_workbook.save("parts_list.xlsx")
print ("Workbook Saved")