import gspread
import random
import json
import requests
import os
from collections import defaultdict
from datetime import datetime, timedelta
import random


webhook_url = os.getenv('WEBHOOK_URL')
google_json = os.getenv('GoogleJson')
 
 
# Step 2: Create a temporary file for the Google service account credentials
google_credentials_path = os.path.expanduser('~/repo/emailserver-415706-bae70316794d.json')
with open(google_credentials_path, 'w') as f:
    f.write(google_json)


service_acc = gspread.service_account(google_credentials_path)


# Function to create a worksheet if it doesn't exist
def get_or_create_worksheet(spreadsheet, title, rows=100, cols=20):
    try:
        return spreadsheet.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        print(f"Worksheet '{title}' not found. Creating a new worksheet.")
        return spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)


spreadsheet_emp = service_acc.open_by_key('1oCXrqaPi8IiqWZQux3vmZKpeb2oftSLcURMlLz2D_jY')

# \------------------------------------------------------------------------------------------------------------/
def getWorksheetToDict(spreadsheet: gspread.spreadsheet.Spreadsheet, work_sheet_name, key_column_name, value_column_name):
    try:
        worksheet = spreadsheet.worksheet(work_sheet_name)
        
        # /-------------------------------------------------------------------------------------------------------\
                ## get_all_records return list of dictionary and each dictionary in list represent recored ##
        # \-------------------------------------------------------------------------------------------------------/
        worksheet_record =  worksheet.get_all_records(expected_headers=[key_column_name, value_column_name])
        
        
        # /-------------------------------------------------------------------------------------------------------\
                ## Empty dictionary which will store two columns data in the form of key and value ##
        # \-------------------------------------------------------------------------------------------------------/
        output_dictionary = dict()
        
        # /-------------------------------------------------------------------------------------------------------\
                            ## Loop over all record and store into c_Cases_Notification_Dict ##
        # \-------------------------------------------------------------------------------------------------------/
        for record in worksheet_record:
            output_dictionary[record[key_column_name]] = record[value_column_name]
            
        return output_dictionary  
      
    except Exception as e:
        raise Exception(e)
    

# /----------------------------------------------------------------------------\
                ## END : GET WORKSHEET DATA INTO DICTIONARY ##
# \----------------------------------------------------------------------------/



# Select the worksheets by their names
worksheet_emp = spreadsheet_emp.worksheet("Emp_Names")
# worksheet_seat = spreadsheet_emp.worksheet("Seat_Capacity")
available_sheet_dict = getWorksheetToDict(spreadsheet_emp, "Seat_Capacity", 'RoomNo', "Seat Count")
worksheet_exclusion = spreadsheet_emp.worksheet("Exclusion")

# Fetch the data from the spreadsheets
emp_data = worksheet_emp.get_all_records()
worksheet_exclusion = worksheet_exclusion.get_all_records()


# Get the all the configurations from the configuration sheet
configuration_sheet_name = "Configuration"
key_col_name = "Key"
value_col_name = 'Value'
configuration_sheet = spreadsheet_emp.worksheet(configuration_sheet_name)
configuration_dict  = getWorksheetToDict(spreadsheet_emp, configuration_sheet_name, key_col_name, value_col_name)
# print(configuration_dict)

# Get data from cache
cache_sheet = spreadsheet_emp.worksheet('Cache')
cache_data = cache_sheet.get_all_records()


exclusion_list:list = []
# print(worksheet_exclusion)
for item in worksheet_exclusion:
    exclusion_list.append(item["Name"])

# function to create new cache data 
def updateCacheData():
    project_type_to_name_map: dict[str, list] = {}
    for record in emp_data:
        record_project = record['Project']
        record_name = record['Name']
        if record_project not in project_type_to_name_map.keys():
            project_type_to_name_map[record_project] = []
            project_type_to_name_map[record_project].append(record_name)
        else:
            project_type_to_name_map[record_project].append(record_name)
            
    # print(project_type_to_name_map)
    
    
    # print("---------------------------------------------")
    # print(available_sheet_dict)
    
    room_to_names_map: dict[any, list] = {}
    current_room_size_map: dict = {}
    
    for room_no in available_sheet_dict:
        current_room_size_map[room_no] = 0
        room_to_names_map[room_no] = []
    
    projects_keys: list = list(project_type_to_name_map.keys())
    print(projects_keys)
     
     
    project_type_to_name_map_copy = project_type_to_name_map.copy()
    
    
    room_names: list = list(room_to_names_map.keys())
    print("Room names", room_names)

    
    room_index = 0
    while len(projects_keys) != 0:
        random_project_name = random.choice(projects_keys)
        
        if "Miscellaneous" not in random_project_name:
            group_emp = project_type_to_name_map_copy[random_project_name]
            if len(group_emp) <= abs(current_room_size_map[room_names[room_index]] - available_sheet_dict[room_names[room_index]]):
                for name in group_emp:
                    if current_room_size_map[room_names[room_index]] != available_sheet_dict[room_names[room_index]] and name not in exclusion_list:
                        room_to_names_map[room_names[room_index]].append(name)
                        current_room_size_map[room_names[room_index]] += 1

                if room_index != len(room_names) - 1:
                    room_index += 1
                else:
                    room_index = 0
            else:
                if room_index != len(room_names) - 1:
                    room_index += 1
                else:
                    room_index = 0
                continue
            
            projects_keys.remove(random_project_name)
        else:
            projects_keys.remove(random_project_name)
     
    projects_keys: list = list(project_type_to_name_map.keys())

    while len(projects_keys) != 0:
        random_project_name = random.choice(projects_keys)
        
        if "Miscellaneous" in random_project_name:
            miscellaneous: list = project_type_to_name_map_copy[random_project_name]
            
            
            while len(miscellaneous) != 0:
                random_name = random.choice(miscellaneous)
                if random_name in exclusion_list:
                    miscellaneous.remove(random_name)
                    continue 

                assigned = False  
                room_index = 0  

                # Loop over rooms to find one with available space
                while room_index < len(room_names):
                    # Check if the current room has available capacity
                    available_capacity = available_sheet_dict[room_names[room_index]] - current_room_size_map[room_names[room_index]]
                    
                    if available_capacity > 0:  # If room has space
                        room_to_names_map[room_names[room_index]].append(random_name)
                        current_room_size_map[room_names[room_index]] += 1
                        assigned = True 
                        break 

                    room_index += 1
                if not assigned:
                    print(f"No available room for {random_name}.")
                    miscellaneous.remove(random_name)
                else:
                    miscellaneous.remove(random_name)
        
        # Remove the processed project key
        projects_keys.remove(random_project_name)

    # Clear cache_sheet
    cache_sheet.clear()
    
    cache_header = ['Room No.', 'Names']
    
    cache_sheet.append_row(cache_header)
    
    for key in room_to_names_map:
        names = ""
        cnt = 0
        for name in room_to_names_map[key]:
            if cnt != len(room_to_names_map[key]) - 1:
                names += name + ","
            else:
                names += name 
            cnt += 1
        print(names)
        cache_sheet.append_row([key, names])
        
    return room_to_names_map    
               
# Create cache if it not present or update is true or size of employee increased than previous 
if len(cache_data) == 0 or configuration_dict['update'] == "ON" or len(emp_data) != configuration_dict['prev_emp_size']:
    
    
    print("Create cache is executed....")



# suffling code



# Send notification
allocation = updateCacheData()


 #Send the Adaptive Card to Webhook (Optional - for MS Teams)
adaptive_card_table = {
    "type": "AdaptiveCard",
    "body": [
        {
            "type": "TextBlock",
            "text": f"Seating Arrangements for Today ({datetime.now().strftime('%d-%m-%Y')})",
            "weight": "Bolder",
            "size": "Medium",
            "wrap": True
        },
        {
            "type": "Table",
            "columns": [
                {"width": 1},
                {"width": 1}
            ],
            "rows": [
                {
                    "type": "TableRow",
                    "cells": [
                        {"type": "TableCell", "items": [{"type": "TextBlock", "text": "Room No.", "wrap": True}]},
                        {"type": "TableCell", "items": [{"type": "TextBlock", "text": "Names", "wrap": True}]}
                    ]
                }
            ],
            "spacing": "Small",
            "separator": True,
            "horizontalAlignment": "Center",
            "horizontalCellContentAlignment": "Center"
        }
    ]
}

# Add room allocations to the adaptive card
for room, people in allocation.items():
    adaptive_card_table["body"][1]["rows"].append(
        {
            "type": "TableRow",
            "cells": [
                {"type": "TableCell", "items": [{"type": "TextBlock", "text": room, "wrap": True}]},
                {"type": "TableCell", "items": [{"type": "TextBlock", "text": ", ".join(people), "wrap": True}]}
            ]
        }
    )

# Send the Adaptive Card message
headers = {'Content-Type': 'application/json'}
response = requests.post(webhook_url, headers=headers, data=json.dumps(adaptive_card_table))

if response.status_code in [200, 202]:
    print("Message posted successfully!")
else:
    print(f"Failed to post message. Status code: {response.status_code}")


if configuration_dict["update"] == 'ON':
    configuration_sheet.update_cell(2, 2, 'OFF')
