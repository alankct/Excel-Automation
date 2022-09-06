import sys
import csv


"""
/datacsv

This command is used to transfer the data from a .csv file into an Excel worksheet. If a csv file is not
given, get_csv_file() is run. Edits the worksheet if the .csv file is succesful, and None if the user
decides to back-track to the main command terminal.
"""

def csv_to_excel(ws, csv_file = None):

    if not csv_file:
        csv_file = get_csv_file()

    if not csv_file:
        # User returned the /back command instead of a .csv file, so they will return to the main terminal
        return None

    csv_data = []
    with open(csv_file) as file_obj:
        reader = csv.reader(file_obj)
        for row in reader:
            csv_data.append(row)

    for row in csv_data:
        ws.append(row)
    
    print("The data was succesfully added to your worksheet")


"""
Checks the validity of the .csv filename, by first making sure the user has not typed /end or /back; ensuring
the program works fluidly. Returns the String 'csv_file' if this file is able to be opened by the program.
"""
def get_csv_file():

    while True:

        csv_file = input(f"Enter the full CSV filename: ")
            
        if csv_file == "/end":
            print("Program ended")
            sys.exit(0)

        elif csv_file == "/back":
            return None

        elif len(csv_file) <= 4 or csv_file[-4:] != ".csv":
            print("This filename is invalid, make sure it ends with .csv")
            continue

        else:
            try:
                open(csv_file)
                return csv_file
            
            except Exception as e:
                print(e)
                continue