import sys
import csv
from urllib.request import urlopen
from bs4 import BeautifulSoup

"""
/datawiki

This command is used to transfer the data from a Wikipedia table into an Excel worksheet. If a Wiki 
URL or table is not found, the command prompts the user to re-type the URL. Returns a .csv file with
all the data if a table is succesfuly found, prompting the main terminal to run the /datacsv command
in order to transfer the .csv data to the current worksheet.
"""

def wikitable_to_csv():
    
    while True:

        wiki_url = input("Insert Wikipedia URL: ")

        if wiki_url == "/end":
            print("Program ended")
            sys.exit(0)

        if wiki_url == "/back":
            return None

        # Makes sure the URL is valid and exists
        try:
            html_data = urlopen(wiki_url)
        except Exception as e:
            print(e)
            continue
        
        # Makes sure a Wikipedia Table exists in this URL
        parser = BeautifulSoup(html_data, "html.parser")
        try:
            table = parser.findAll("table", {"class":"wikitable"})[0]
        except:
            print("A Wikipedia table was not found using this URL")
            continue
        
        # Aggregates all the data from the parsed table into a .csv file
        csv_file = "delete__this__.csv"
        rows = table.findAll("tr")
        with open(csv_file, "wt+", newline="") as f:
            writer = csv.writer(f)
            for i in rows:
                row = []
                for cell in i.findAll(["td", "th"]):
                    row.append(cell.get_text())
                writer.writerow(row)

        return csv_file