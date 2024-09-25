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
all_people = []
for name, project in data_dict.items():
    if name not in exclusion_names:
        project_groups[project].append(name)
        all_people.append(name)

# Step 5: Shuffle all people and assign rooms
random.shuffle(all_people)
room_assignments = {room: [] for room in seat_dict.keys()}

person_index = 0
while person_index < len(all_people):
    for room, capacity in seat_dict.items():
        if len(room_assignments[room]) < capacity and person_index < len(all_people):
            room_assignments[room].append(all_people[person_index])
            person_index += 1

# Step 6: Try to keep project members together if possible
for room, people in room_assignments.items():
    project_counts = defaultdict(int)
    for person in people:
        project = data_dict[person]
        project_counts[project] += 1
    
    # If there's a dominant project in the room, try to swap others out
    if project_counts:
        dominant_project = max(project_counts, key=project_counts.get)
        if project_counts[dominant_project] > len(people) / 2:
            for i, person in enumerate(people):
                if data_dict[person] != dominant_project:
                    for other_room, other_people in room_assignments.items():
                        if other_room != room:
                            for j, other_person in enumerate(other_people):
                                if data_dict[other_person] == dominant_project:
                                    # Swap
                                    room_assignments[room][i], room_assignments[other_room][j] = room_assignments[other_room][j], room_assignments[room][i]
                                    break
                    if data_dict[room_assignments[room][i]] == dominant_project:
                        break

# Step 7: Create or open the 'Today's Arrangement' sheet
arrangement_sheet_name = "Today's Arrangement"

try:
    # Try to load yesterday's arrangement
    worksheet_today = spreadsheet_emp.worksheet(arrangement_sheet_name)
    yesterday_data = worksheet_today.get_all_records()
    
    # Convert yesterday's data into a set for easy lookup
    yesterday_assignments = {(entry['Room No.'], name) for entry in yesterday_data for name in entry['Names'].split(', ')}
    
    # Step 8: Apply the seating algorithm to avoid repetition from yesterday
    for room, people in room_assignments.items():
        new_assignment = []
        for person in people:
            if (room, person) in yesterday_assignments:
                # Try to move this person to another room
                for other_room, other_people in room_assignments.items():
                    if other_room != room and len(other_people) > 0:
                        swap_candidate = random.choice(other_people)
                        if (other_room, swap_candidate) not in yesterday_assignments:
                            other_people.remove(swap_candidate)
                            other_people.append(person)
                            new_assignment.append(swap_candidate)
                            break
                else:
                    # If no swap was possible, keep the person in the same room
                    new_assignment.append(person)
            else:
                new_assignment.append(person)
        room_assignments[room] = new_assignment

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
