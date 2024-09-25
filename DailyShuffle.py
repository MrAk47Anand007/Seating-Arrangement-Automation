import gspread
import random
import os
import json
import requests
from collections import defaultdict
from datetime import datetime, timedelta

# Step 1: Get the webhook URL and Google Service Account JSON from environment variables
webhook_url = os.getenv('WEBHOOK_URL')
google_json = os.getenv('GoogleJson')

# Step 2: Create a temporary file for the Google service account credentials
google_credentials_path = os.path.expanduser('~/repo/emailserver-415706-bae70316794d.json')
with open(google_credentials_path, 'w') as f:
    f.write(google_json)

# Step 3: Authenticate using the Google service account JSON file
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

# Step 4: Group employees by project
project_groups = defaultdict(list)
for name, project in data_dict.items():
    if name not in exclusion_names:
        project_groups[project].append(name)

# Step 5: Shuffle projects and assign rooms
room_assignments = {room: [] for room in seat_dict.keys()}
projects = list(project_groups.keys())
random.shuffle(projects)

for project in projects:
    people = project_groups[project]
    random.shuffle(people)
    
    # Find the room with the most available space
    target_room = max(room_assignments, key=lambda x: seat_dict[x] - len(room_assignments[x]))
    
    # Calculate how many people we can fit in this room
    available_space = seat_dict[target_room] - len(room_assignments[target_room])
    people_to_assign = min(len(people), available_space)
    
    # Assign people to the room
    room_assignments[target_room].extend(people[:people_to_assign])
    
    # If there are remaining people, try to keep them together in another room
    remaining_people = people[people_to_assign:]
    while remaining_people:
        target_room = max(room_assignments, key=lambda x: seat_dict[x] - len(room_assignments[x]))
        available_space = seat_dict[target_room] - len(room_assignments[target_room])
        people_to_assign = min(len(remaining_people), available_space)
        room_assignments[target_room].extend(remaining_people[:people_to_assign])
        remaining_people = remaining_people[people_to_assign:]

# Step 6: Create or open the 'Today's Arrangement' sheet
arrangement_sheet_name = "Today's Arrangement"

try:
    # Try to load yesterday's arrangement
    worksheet_today = spreadsheet_emp.worksheet(arrangement_sheet_name)
    yesterday_data = worksheet_today.get_all_records()
    
    # Convert yesterday's data into a dictionary for easy lookup
    yesterday_assignments = {room: set(names.split(', ')) for room, names in 
                             [(entry['Room No.'], entry['Names']) for entry in yesterday_data]}
    
    # Step 7: Apply the seating algorithm to avoid repetition from yesterday
    for room, people in room_assignments.items():
        if room in yesterday_assignments:
            # Find people who were in this room yesterday
            repeat_people = set(people) & yesterday_assignments[room]
            if repeat_people:
                # Try to swap these people with those in other rooms
                for other_room, other_people in room_assignments.items():
                    if other_room != room:
                        for person in repeat_people.copy():
                            if person in people:  # Check if person is still in the current room
                                for other_person in other_people:
                                    if other_person not in yesterday_assignments.get(other_room, set()):
                                        # Swap
                                        people[people.index(person)] = other_person
                                        other_people[other_people.index(other_person)] = person
                                        repeat_people.remove(person)
                                        break
                            if not repeat_people:
                                break
                    if not repeat_people:
                        break

except gspread.exceptions.WorksheetNotFound:
    # First day case: No previous data, so keep the current assignments
    worksheet_today = spreadsheet_emp.add_worksheet(title=arrangement_sheet_name, rows="100", cols="10")

# Step 9: Write today's arrangement to the "Today's Arrangement" sheet
arrangement_data = [["Room No.", "Names"]]
for room, people in room_assignments.items():
    arrangement_data.append([room, ", ".join(people)])

# Overwrite or update the existing "Today's Arrangement" sheet
worksheet_today.update("A1", arrangement_data)

# Get today's date in dd-mm-yyyy format adjusted for IST (UTC+5:30)
ist_offset = timedelta(hours=5, minutes=30)
today = (datetime.utcnow() + ist_offset).strftime("%d-%m-%Y")

# Step 10: Prepare the Adaptive Card for Teams notification
adaptive_card_table = {
    "type": "AdaptiveCard",
    "body": [
        {
            "type": "TextBlock",
            "text": f"Seating Arrangements for Today ({today})",
            "weight": "Bolder",
            "size": "Medium",
            "wrap": True
        },
        {
            "type": "Table",
            "columns": [
                {
                    "width": 1
                },
                {
                    "width": 1
                }
            ],
            "rows": [
                {
                    "type": "TableRow",
                    "cells": [
                        {
                            "type": "TableCell",
                            "items": [
                                {
                                    "type": "TextBlock",
                                    "text": "Room No.",
                                    "wrap": True
                                }
                            ]
                        },
                        {
                            "type": "TableCell",
                            "items": [
                                {
                                    "type": "TextBlock",
                                    "text": "Names",
                                    "wrap": True
                                }
                            ]
                        }
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

# Adding rows with room and employee names
for room, people in room_assignments.items():
    adaptive_card_table["body"][1]["rows"].append(
        {
            "type": "TableRow",
            "cells": [
                {
                    "type": "TableCell",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": room,
                            "wrap": True
                        }
                    ]
                },
                {
                    "type": "TableCell",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": ", ".join(people),
                            "wrap": True,
                            "horizontalAlignment": "Left"
                        }
                    ]
                }
            ]
        }
    )

# Step 11: Send Adaptive Card to Webhook
headers = {
    'Content-Type': 'application/json'
}

response = requests.post(webhook_url, headers=headers, data=json.dumps(adaptive_card_table))

# Check if the request was successful
if response.status_code == 200 or 202:
    print("Message posted successfully!")
else:
    print(f"Failed to post message. Status code: {response.status_code}, Response: {response.text}")
