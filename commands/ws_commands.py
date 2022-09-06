import sys

WS_COMMANDS = ("/wscreate", "/wschange", "/wslist")

""" 
The worksheet commands are ran at any time to create a new worksheet, change the current active
worksheet, or list out all worksheets. Returns a worksheet, or None if the user back-tracks.
"""          
def worksheet(wb, wscommand):

    ws = None

    match wscommand:
        case "/ws":
            print(f"Worksheet commands: {WS_COMMANDS}")
        case "/wscreate":
            ws = wscreate(wb)
        case "/wschange":
            ws = wschange_to(wb)
        case "/wslist":
            print(wb.sheetnames)
        case _:
            print("Invalid /ws command.")

    return ws


"""
/wscreate

Creates a new worksheet. Checks validity of new worksheet name. Prompts the user for what they want
their active worksheet to be by calling wschange_to(), and returns the worksheet.
"""
def wscreate(wb):

    while True:
        sheet_name = input("Name of new worksheet: ")

        if sheet_name == "/end":
            print("Program ended")
            sys.exit(0)

        if sheet_name == "/back":
            print("Defaulting to current active worksheet")
            return None
        
        # Catches invalid sheet names
        try:
            wb.create_sheet(sheet_name)
        except Exception as e:
            print(e)
            continue

        ws = wschange_to(wb)
        break
        
    return ws

"""
/wschange

Changes active worksheet, or defaults to the active worksheet if the user enters /back. Returns a
worksheet.
"""
def wschange_to(wb):

    while True:

        print(f"All sheets in current workbook: {wb.sheetnames}")
        sheet_name = input("Change active worksheet to: ")

        if sheet_name == "/end":
            print("Program ended")
            sys.exit(0)

        if sheet_name == "/back":
            print("Defaulting to current active worksheet")
            return None
        
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            return ws

        else:
            print("Worksheet does not exist.")