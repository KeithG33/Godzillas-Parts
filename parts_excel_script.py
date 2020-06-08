from openpyxl import load_workbook
import sys

input_accept = False


godzilla_parts_workbook = load_workbook('parts_list.xlsx')
ws = godzilla_parts_workbook.active


# Lets split this up into methods for different actions.
def add_part():
    global input_accept
    global godzilla_parts_workbook
    global ws

    while input_accept == False:
        # Get me inputs
        part_string = input("Ready to go, add part using 'Brand, Part, Purchase $' format:  ")
        part_list = part_string.split(", ")

        # Make sure I didn't ... up
        print("You inputted: \nBrand: " + part_list[0] + "\nPart: " + part_list[1] + "\nPurchase Price: " + part_list[2])

        keep_or_not = input("Do you want to keep this input? Answer y/n\n ")

        if keep_or_not == "y":
            input_accept = True
            break
        elif keep_or_not == "n":
            continue
        elif keep_or_not != "y" or input_accept != "n":
            ("Input error, please type 'y' for yes and 'n' for no...Try again\n ")

    input_accept = False

    brand = part_list[0]
    part = part_list[1]
    price = part_list[2]

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
    
    # Now sum up price
    ws["F5"] = "=SUM(D6:D2000)"

def peace_out():
    print("Nos vemos cuando gastas mas dinero ;)")
    exit()

def check_damage():
    global ws
    money_spent = ws['F5']
    print("You've spent: ${}".format(money_spent))

def chooser(choice):
    switcher = {
                0: add_part,
                2: check_damage,
                3: peace_out
                }
    func = switcher.get(choice, lambda:'Invalid, choose from the listed choices')
    return func()
    

# Now we run the interface
if not godzilla_parts_workbook:
    print("Excel sheet didn't load, better figure that out buddy guy")
    exit()
else:
    while True:
        action_choice = input("Welcome to the excel-omator...how would you like to proceed?\n\n 1 - Add part \n 2 - See spending total \n 3 - Exit")
        chooser(action_choice)




    

    

    # Got the info we need, time to edit excel sheet. Brand is ColB, part is ColC, price ColD
godzilla_parts_workbook.save("parts_list.xlsx")
print ("Workbook Saved")