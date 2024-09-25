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

# Step 3: Group employees by project
project_groups = defaultdict(list)
misc_people = []

# Categorize employees into project groups or misc_people
for name, project in data_dict.items():
    if name not in exclusion_names:
        if "Miscellaneous" in project:
            misc_people.append(name)
        else:
            project_groups[project].append(name)

# Shuffle remaining misc employees for dynamic allocation
random.shuffle(misc_people)

# Step 4: Initialize room allocations and outside space allocations
room_allocations = {room: [] for room in seat_dict}
outside_space_allocations = {f"{room}(Outside Space)": [] for room in seat_dict}

# Randomize the room order
rooms = list(seat_dict.keys())
random.shuffle(rooms)

# Function to allocate groups to rooms with random room indexing
def allocate_group(group, room_allocations, outside_space_allocations, seat_dict):
    random.shuffle(rooms)  # Shuffle the rooms to make the allocation order random
    for room in rooms:
        capacity = seat_dict[room]
        if len(room_allocations[room]) + len(group) <= capacity:
            room_allocations[room].extend(group)
            return  # Group successfully allocated, exit the function

    # If no room has enough space, assign the group to the outside space of the first available room
    outside_space_allocations[f"{rooms[0]}(Outside Space)"].extend(group)

# Step 5: Load previous allocations from Google Sheets
def load_previous_allocations(worksheet):
    allocations = {}
    records = worksheet.get_all_records()
    for record in records:
        room = record.get("Room No.")
        names = record.get("Names", "")
        allocations[room] = names.split(", ") if names else []
    return allocations

# Step 6: Shuffle groups based on historical data
def shuffle_groups_with_history(project_groups, previous_allocations):
    shuffled_groups = []
    
    # Create a list of all groups
    all_groups = list(project_groups.values())
    
    # Shuffle groups
    random.shuffle(all_groups)

    # Check against previous allocations to avoid repeats
    for group in all_groups:
        if not any(set(group).issubset(set(previous)) for previous in previous_allocations.values()):
            shuffled_groups.append(group)

    return shuffled_groups

# Step 7: Write the allocations to Google Sheets
def write_allocations_to_sheet(worksheet, allocation):
    # Prepare data for writing
    allocation_data = []
    for room_name, people in allocation['rooms'].items():
        allocation_data.append([room_name, ', '.join(people)])

    # Append data to the worksheet
    worksheet.append_rows(allocation_data)

    # Optional: Add a timestamp for tracking
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    worksheet.append_row(["Timestamp", timestamp])

# Step 8: Allocate groups to rooms with consideration of previous allocations
def allocate_groups_to_rooms(project_groups, misc_people, seat_dict, previous_allocations):
    room_allocations = {room: [] for room in seat_dict}
    outside_space_allocations = {f"{room}(Outside Space)": [] for room in seat_dict}

    rooms = list(seat_dict.keys())
    random.shuffle(rooms)

    # Allocate project groups
    for group in project_groups:
        allocate_group(group, room_allocations, outside_space_allocations, seat_dict)

    # Allocate remaining misc employees to rooms
    for room in rooms:
        capacity = seat_dict[room]
        while len(room_allocations[room]) < capacity and misc_people:
            room_allocations[room].append(misc_people.pop())

    # Handle leftover misc employees by placing them in outside spaces
    if misc_people:
        for room in rooms:
            if misc_people:
                outside_space_allocations[f"{room}(Outside Space)"].extend(misc_people)
                misc_people = []

    return {
        "rooms": room_allocations,
        "outside_spaces": outside_space_allocations
    }

# Step 9: Main Execution Logic

# Ensure the "Group_Allocations" worksheet exists or create it
allocation_worksheet = get_or_create_worksheet(spreadsheet_emp, "Group_Allocations")

# Load previous allocations
previous_allocations = load_previous_allocations(allocation_worksheet)

# Shuffle project groups based on previous allocations
shuffled_project_groups = shuffle_groups_with_history(project_groups, previous_allocations)

# Allocate groups to rooms
allocation = allocate_groups_to_rooms(shuffled_project_groups, misc_people, seat_dict, previous_allocations)

# Write the new allocations back to Google Sheets
write_allocations_to_sheet(allocation_worksheet, allocation)

# Send the Adaptive Card to Webhook (Optional - for MS Teams)
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
for room, people in allocation["rooms"].items():
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
