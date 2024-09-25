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

# Step 5: Assign rooms based on seat availability
room_assignments = {}
assigned_people = set()

# Priority queue (smallest group assigned first)
project_heap = []
for project, people in project_groups.items():
    if people:
        heapq.heappush(project_heap, (len(people), project))

for room, seat_count in seat_dict.items():
    room_assignments[room] = []
    remaining_seats = seat_count

    while project_heap and remaining_seats > 0:
        group_size, project = heapq.heappop(project_heap)

        if group_size <= remaining_seats:
            room_assignments[room].extend(project_groups[project])
            remaining_seats -= group_size
            assigned_people.update(project_groups[project])
        else:
            room_assignments[room].extend(project_groups[project][:remaining_seats])
            assigned_people.update(project_groups[project][:remaining_seats])
            project_groups[project] = project_groups[project][remaining_seats:]
            heapq.heappush(project_heap, (len(project_groups[project]), project))
            remaining_seats = 0

# Step 6: Shuffle and assign miscellaneous people into rooms with available space
misc_people = [person for person in misc_people if person not in assigned_people]
random.shuffle(misc_people)

for room, people in room_assignments.items():
    remaining_seats = seat_dict[room] - len(people)
    if remaining_seats > 0 and misc_people:
        to_assign = misc_people[:remaining_seats]
        room_assignments[room].extend(to_assign)
        misc_people = misc_people[remaining_seats:]

# Step 7: Create or open the 'Today's Arrangement' sheet
arrangement_sheet_name = "Today's Arrangement"

try:
    # Try to load yesterday's arrangement from the "Today's Arrangement" sheet
    worksheet_today = spreadsheet_emp.worksheet(arrangement_sheet_name)
    yesterday_data = worksheet_today.get_all_records()

    # Convert yesterday's data into a dictionary for easy lookup
    yesterday_assignments = {entry['Room No.']: entry['Names'].split(', ') for entry in yesterday_data}

    # Step 8: Apply the seating algorithm to avoid repetition from yesterday
    for room, people in room_assignments.items():
        previous_people = set(yesterday_assignments.get(room, []))
        new_people = [p for p in people if p not in previous_people]
        remaining_people = [p for p in people if p in previous_people]

        # Shuffle the remaining people if necessary to avoid complete repetition
        if len(new_people) == 0 and remaining_people:
            random.shuffle(remaining_people)
            room_assignments[room] = remaining_people
        else:
            room_assignments[room] = new_people + remaining_people

except gspread.exceptions.WorksheetNotFound:
    # First day case: No previous data, so generate from scratch
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
