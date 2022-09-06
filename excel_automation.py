from commands import wb_commands, ws_commands, csv_to_excel, wiki_to_csv

ALL_COMMANDS = ("/wb", "/wbcreate", "/wbload", "/ws", "/wscreate", "/wschange", "/wslist", "/datacsv", "/datawiki")
WB_COMMANDS = ("/wb", "/wbcreate", "/wbload")
WS_COMMANDS = ("/wscreate", "/wschange", "/wslist")

def main():
    
    wb = None
    wb_name = ""
    ws = None
    command = ""

    while command != "/end":

        # Workbook commands (these are always ran first until the user creates or loads a workbook)
        if wb == None or command in WB_COMMANDS:
            poss_wb, poss_wb_name  = wb_commands.workbook(command)
            if poss_wb:
                wb, wb_name = poss_wb, poss_wb_name
                ws = wb.active
                ws_statement = f"Currently working on worksheet '{ws.title}' in workbook '{wb_name}'"
                print(ws_statement)
            elif wb == None and poss_wb == None: # The user back-tracked but has not created a wb yet
                print("In order to use this program, create or load a workbook by typing /wbcreate or /wbload")

        # Worksheet commands list the available worksheets, create a new worksheet, or change the active worksheet
        elif command in WS_COMMANDS:
            print(ws_statement)
            poss_ws = ws_commands.worksheet(wb, command)
            if poss_ws:
                ws = poss_ws
            print(ws_statement)

        else:
            match command:
                case "":
                    pass
                case "/back":
                    pass
                case "/save":
                    wb.save(wb_name)
                    print("Saved")
                case "/commands":
                    print(ALL_COMMANDS)
                case "/ws":
                    print(f"Worksheet Commands: {WS_COMMANDS}")
                    print(ws_statement)
                case "/datacsv":
                    print(f"For cleanest results, make sure '{ws.title}' is a blank worksheet. To go back: /back")
                    csv_to_excel.csv_to_excel(ws)
                case "/datawiki":
                    print(f"For cleanest results, make sure '{ws.title}' is a blank worksheet. To go back: /back")
                    csv_file = wiki_to_csv.wikitable_to_csv()
                    if csv_file:
                        csv_to_excel.csv_to_excel(ws, csv_file)
                case _:
                    # If an exact match is not confirmed, this last case will be used
                    print("Command not found. For a list of commands, type /commands")
        
        # Automatically autosaves all changes made on the Excel file after every command is run
        if wb != None:
            wb.save(wb_name)

        # Prompts the user for a command after every iteration
        command = input("Type a command: ") 
    
    print("Program ended")


main()