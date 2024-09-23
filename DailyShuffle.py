import gspread
import random
import os
import json
import requests
from collections import defaultdict

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
misc_people = []
for name, project in data_dict.items():
    if name not in exclusion_names:
        if "Miscellaneous" in project:
            misc_people.append(name)
        else:
            project_groups[project].append(name)

# Step 5: Assign rooms based on seat availability
room_assignments = {}
assigned_people = set()

# Assign project groups first
for room, seat_count in seat_dict.items():
    projects = list(project_groups.items())
    random.shuffle(projects)

    room_assignments[room] = []
    remaining_seats = seat_count

    for project, people in projects:
        if len(people) <= remaining_seats:
            room_assignments[room].extend(people)
            remaining_seats -= len(people)
            assigned_people.update(people)
            project_groups[project] = []

# Step 6: Shuffle and assign misc people into rooms with available space
misc_people = [person for person in misc_people if person not in assigned_people]
random.shuffle(misc_people)

for room, people in room_assignments.items():
    remaining_seats = seat_dict[room] - len(people)
    if remaining_seats > 0 and misc_people:
        to_assign = misc_people[:remaining_seats]
        room_assignments[room].extend(to_assign)
        misc_people = misc_people[remaining_seats:]

# Handle exclusions (no changes for these people)
for exclusion in exclusion_names:
    for room, people in room_assignments.items():
        if exclusion in data_dict and exclusion in people:
            room_assignments[room].append(exclusion)

# Create the adaptive card table with title and formatted table rows
adaptive_card = {
    "type": "AdaptiveCard",
    "body": [
        {
            "type": "TextBlock",
            "text": f"Seating Arrangements for Today ({today})",
            "weight": "Bolder",
            "size": "Medium"
        },
        {
            "type": "Table",
            "columns": [
                {
                    "type": "TableColumn",
                    "width": "stretch"
                },
                {
                    "type": "TableColumn",
                    "width": "stretch"
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
                                    "weight": "Bolder",
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
                                    "weight": "Bolder",
                                    "wrap": True
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    ],
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "version": "1.3"
}

# Add room and employee names to the table with text wrapping enabled
for room, people in room_assignments.items():
    adaptive_card["body"][1]["rows"].append(
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
                            "wrap": True
                        }
                    ]
                }
            ]
        }
    )

# Headers for the HTTP POST request
headers = {
    'Content-Type': 'application/json'
}


# Send the adaptive card payload to the webhook
response = requests.post(webhook_url, headers=headers, data=json.dumps(adaptive_card))

# Check if the request was successful
if response.status_code == 200 or 202:
    print("Message posted successfully!")
else:
    print(f"Failed to post message. Status code: {response.status_code}, Response: {response.text}")
