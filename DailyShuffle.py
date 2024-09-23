import gspread
import random
import os
import requests
from collections import defaultdict

# Step 1: Get the webhook URL and Google Service Account credentials from environment variables
webhook_url = os.getenv('WEBHOOK_URL')
google_service_account_path = os.path.expanduser('~/repo/emailserver-415706-bae70316794d.json')

# Authenticate using the Google service account JSON file path
service_acc = gspread.service_account(google_service_account_path)

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
for name, project in data_dict.items():
    if name not in exclusion_names:
        if "Miscellaneous" in project:
            misc_people.append(name)
        else:
            project_groups[project].append(name)

# Step 4: Assign rooms based on seat availability
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

# Step 5: Shuffle and assign misc people into rooms with available space
misc_people = [person for person in misc_people if person not in assigned_people]
random.shuffle(misc_people)

for room, people in room_assignments.items():
    remaining_seats = seat_dict[room] - len(people)
    if remaining_seats > 0 and misc_people:
        to_assign = misc_people[:remaining_seats]
        room_assignments[room].extend(to_assign)
        misc_people = misc_people[remaining_seats:]

# Step 6: Handle exclusions (no changes for these people)
for exclusion in exclusion_names:
    for room, people in room_assignments.items():
        if exclusion in data_dict and exclusion in people:
            room_assignments[room].append(exclusion)

# Prepare Adaptive Card for Microsoft Teams
adaptive_card = {
    "type": "message",
    "attachments": [
        {
            "contentType": "application/vnd.microsoft.card.adaptive",
            "content": {
                "type": "AdaptiveCard",
                "version": "1.2",
                "body": [
                    {
                        "type": "TextBlock",
                        "text": f"Seating Arrangements for Today ({os.popen('date +%d-%m-%Y').read().strip()})",
                        "weight": "Bolder",
                        "size": "Large"
                    },
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "auto",
                                "items": [{"type": "TextBlock", "text": "Room No", "weight": "Bolder"}]
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [{"type": "TextBlock", "text": "Names", "weight": "Bolder"}]
                            }
                        ]
                    }
                ]
            }
        }
    ]
}

# Add room assignments to the adaptive card
for room, people in room_assignments.items():
    row = {
        "type": "ColumnSet",
        "columns": [
            {
                "type": "Column",
                "width": "auto",
                "items": [{"type": "TextBlock", "text": room}]
            },
            {
                "type": "Column",
                "width": "stretch",
                "items": [{"type": "TextBlock", "text": ", ".join(people)}]
            }
        ]
    }
    adaptive_card['attachments'][0]['content']['body'].append(row)

# Send POST request to the webhook URL
response = requests.post(webhook_url, json=adaptive_card)

# Check for successful response
if response.status_code == 200:
    print("Seating arrangements posted successfully!")
else:
    print(f"Failed to post seating arrangements: {response.status_code}, {response.text}")
