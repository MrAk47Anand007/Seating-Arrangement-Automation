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
misc_people = []
for name, project in data_dict.items():
    if name not in exclusion_names:
        if "Miscellaneous" in project:
            misc_people.append(name)
        else:
            project_groups[project].append(name)

# Shuffle the people in each project group
for project, people in project_groups.items():
    random.shuffle(people)

# Step 5: Assign rooms based on seat availability
room_assignments = {}
assigned_people = set()

# Shuffle the project groups randomly for each room assignment
projects = list(project_groups.items())
random.shuffle(projects)

# Assign people from project groups to rooms
for room, seat_count in seat_dict.items():
    room_assignments[room] = []
    remaining_seats = seat_count

    for project, people in projects:
        if len(people) <= remaining_seats:
            # Assign the people to the room and remove them from the project group
            assigned_people.update(people)
            room_assignments[room].extend(people)
            remaining_seats -= len(people)
            project_groups[project] = []  # Clear the group once assigned
        else:
            # Assign as many people as possible from this project and then stop
            to_assign = people[:remaining_seats]
            assigned_people.update(to_assign)
            room_assignments[room].extend(to_assign)
            project_groups[project] = people[remaining_seats:]  # Keep unassigned people for the next room
            break  # Move to the next room

# Step 6: Assign Miscellaneous people (randomly) to available rooms
random.shuffle(misc_people)
for room, people in room_assignments.items():
    remaining_seats = seat_dict[room] - len(people)
    if remaining_seats > 0 and misc_people:
        to_assign = misc_people[:remaining_seats]
        room_assignments[room].extend(to_assign)
        misc_people = misc_people[remaining_seats:]  # Remove the assigned people from misc list

# Step 7: Write the assignments to the "Today's Arrangement" sheet
arrangement_sheet_name = "Today's Arrangement"
try:
    worksheet_today = spreadsheet_emp.worksheet(arrangement_sheet_name)
    worksheet_today.clear()  # Clear the old arrangement
except gspread.exceptions.WorksheetNotFound:
    worksheet_today = spreadsheet_emp.add_worksheet(title=arrangement_sheet_name, rows="100", cols="10")

arrangement_data = [["Room No.", "Names"]]
for room, people in room_assignments.items():
    arrangement_data.append([room, ", ".join(people)])

worksheet_today.update("A1", arrangement_data)

# Step 8: Prepare Adaptive Card for Microsoft Teams
ist_offset = timedelta(hours=5, minutes=30)
today = (datetime.utcnow() + ist_offset).strftime("%d-%m-%Y")

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

# Adding rows with room and employee names to the card
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
                            "wrap": True
                        }
                    ]
                }
            ]
        }
    )

# Step 9: Send Adaptive Card to Webhook
headers = {'Content-Type': 'application/json'}
response = requests.post(webhook_url, headers=headers, data=json.dumps(adaptive_card_table))

# Check if the request was successful
if response.status_code == 200 or 202:
    print("Message posted successfully!")
else:
    print(f"Failed to post message. Status code: {response.status_code}, Response: {response.text}")
