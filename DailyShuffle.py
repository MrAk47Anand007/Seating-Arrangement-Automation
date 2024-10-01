import gspread
import random
import json
import requests
import os
from collections import defaultdict
from datetime import datetime, timedelta

# Function to create a worksheet if it doesn't exist
def get_or_create_worksheet(spreadsheet, title, rows=100, cols=20):
    try:
        return spreadsheet.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        print(f"Worksheet '{title}' not found. Creating a new worksheet.")
        return spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)

# Step 1: Get the webhook URL and Google Service Account JSON from environment variables
webhook_url = os.getenv('WEBHOOK_URL')
google_json = os.getenv('GoogleJson')
 
 
# Step 2: Create a temporary file for the Google service account credentials
google_credentials_path = os.path.expanduser('~/repo/emailserver-415706-bae70316794d.json')
with open(google_credentials_path, 'w') as f:
    f.write(google_json)


service_acc = gspread.service_account(google_credentials_path)

# Open the spreadsheet by its key
spreadsheet_emp = service_acc.open_by_key('1oCXrqaPi8IiqWZQux3vmZKpeb2oftSLcURMlLz2D_jY')

# Select the worksheets by their names
worksheet_emp = spreadsheet_emp.worksheet("Emp_Names")
worksheet_seat = spreadsheet_emp.worksheet("Seat_Capacity")
worksheet_exclusion = spreadsheet_emp.worksheet("Exclusion")

# Fetch the data from the spreadsheets
emp_data = worksheet_emp.get("A2:B19")
seat_avail = worksheet_seat.get("A1:B4")
exclusion_list = worksheet_exclusion.get_all_records()

# Convert the list of lists into dictionaries
data_dict = {row[0]: row[1] for row in emp_data if len(row) == 2}
seat_dict = {row[0]: int(row[1]) for row in seat_avail[1:] if len(row) == 2}

# Convert exclusion data to a list of names
exclusion_names = [entry['Name'] for entry in exclusion_list]

# Step 2: Group employees by project
project_groups = defaultdict(list)
misc_people = []

# Categorize employees into project groups or misc_people
for name, project in data_dict.items():
    if name not in exclusion_names:
        if "Miscellaneous" in project:
            misc_people.append(name)
        else:
            project_groups[project].append(name)

# Step 3: Shuffle project groups (Merge Shuffle)
def merge_shuffle(groups):
    # Convert dict values (groups) to a list
    group_list = list(groups.values())
    
    # Shuffle the order of groups
    random.shuffle(group_list)
    
    # Flatten the shuffled list into a single list
    shuffled_list = [person for group in group_list for person in group]
    
    return shuffled_list

shuffled_project_people = merge_shuffle(project_groups)
random.shuffle(misc_people)  # Shuffle misc people

# Step 4: Initialize room allocations
room_allocations = {room: [] for room in seat_dict}

# Step 5: Randomize room order and allocate people (keeping groups together)
def allocate_groups_to_rooms(groups, room_allocations, seat_dict):
    random.shuffle(rooms)  # Shuffle rooms to make allocation random
    for group in groups:
        for room in rooms:
            capacity = seat_dict[room]
            if len(room_allocations[room]) + len(group) <= capacity:
                room_allocations[room].extend(group)
                break

rooms = list(seat_dict.keys())
random.shuffle(rooms)

# Step 6: Allocate shuffled project people first (keep groups together)
allocate_groups_to_rooms(list(project_groups.values()), room_allocations, seat_dict)

# Step 7: Allocate misc people
def allocate_misc_to_rooms(misc_people, room_allocations, seat_dict):
    for person in misc_people:
        for room in rooms:
            capacity = seat_dict[room]
            if len(room_allocations[room]) < capacity:
                room_allocations[room].append(person)
                break

allocate_misc_to_rooms(misc_people, room_allocations, seat_dict)

# Step 8: Write the allocations to Google Sheets
def write_allocations_to_sheet(worksheet, allocation):
    # Prepare data for writing
    allocation_data = []
    for room_name, people in allocation.items():
        allocation_data.append([room_name, ', '.join(people)])

    # Append data to the worksheet
    worksheet.append_rows(allocation_data)

    # Optional: Add a timestamp for tracking
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    worksheet.append_row(["Timestamp", timestamp])

# Combine room allocations
allocation = room_allocations

# Write the allocations back to the Google Sheet
allocation_worksheet = get_or_create_worksheet(spreadsheet_emp, "Group_Allocations")
write_allocations_to_sheet(allocation_worksheet, allocation)

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
