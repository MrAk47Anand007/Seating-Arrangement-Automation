# Seating Arrangement Automation

This Python-based project automates the assignment of seating arrangements for employees based on their project groups and seat availability. The script integrates with **Google Sheets** to fetch employee and seat data, processes seating assignments dynamically, and posts the final seating arrangement to **Microsoft Teams** using a **Power Automate webhook** and **Adaptive Cards**.

## Features

- Fetches employee names, project groups, and seat availability from Google Sheets.
- Assigns employees to rooms dynamically based on project groups and available seats.
- Ensures that employees working on the same project are seated together.
- Randomizes seating for employees assigned to miscellaneous projects.
- Preserves seating for employees listed on the exclusion sheet.
- Posts the seating arrangements to Microsoft Teams using an Adaptive Card.
- Date-specific title in the Teams message (e.g., "Seating Arrangements for Today (dd-mm-yyyy)").

## Technologies Used

- **Python 3.x**
- **gspread** for interacting with Google Sheets API
- **requests** for making HTTP requests to Power Automate
- **Google Sheets API**
- **Microsoft Teams Adaptive Cards**
- **Power Automate Webhook**

## Prerequisites

To run this project locally, you will need:

1. **Python 3.x** installed on your system.
2. **Google Service Account** with access to the Google Sheets used for employee and seating data.
3. A **Power Automate webhook URL** that posts messages to Microsoft Teams.
4. Required Python packages (detailed below).

## Setup Instructions

### 1. Clone the Repository

```bash
git clone https://github.com/your-username/seating-arrangement-automation.git
cd seating-arrangement-automation
```

### 2. Install Required Packages

Create a virtual environment (optional but recommended):

```bash
python -m venv venv
source venv/bin/activate   # On Windows, use `venv\Scripts\activate`
```

Then install the necessary packages:

```bash
pip install -r requirements.txt
```

The `requirements.txt` file should include:

```
gspread
oauth2client
requests
```

### 3. Set Up Google Service Account

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project and enable the **Google Sheets API**.
3. Create a **Service Account** and download the JSON credentials file.
4. Share your Google Sheet with the service account email (something like `your-service-account@project-id.iam.gserviceaccount.com`) to grant access.

### 4. Modify the Python Script

In the Python script, update the following variables:
- **Service Account JSON Path**: Provide the local path to your Google Service Account credentials.
  
    ```python
    service_acc = gspread.service_account("path/to/your/service_account.json")
    ```

- **Google Sheet ID**: Add the unique ID of your Google Sheet containing employee and seating data.
  
    ```python
    spreadsheet_emp = service_acc.open_by_key('your-google-sheet-id')
    ```

- **Webhook URL**: Replace the placeholder webhook URL with your Power Automate webhook URL.
  
    ```python
    webhook_url = "https://your-power-automate-webhook-url"
    ```

### 5. Run the Script

Once everything is set up, run the Python script:

```bash
python daily_shuffle.py
```

This will:
- Fetch data from Google Sheets.
- Shuffle and assign employees to rooms based on the rules.
- Send the seating arrangement to Microsoft Teams via Power Automate.

## Usage

1. **Data Structure**:
    - The Google Sheet should have three tabs:
        - **Emp_Names**: Contains employee names and their corresponding project names.
        - **Seat_Capacity**: Lists the room numbers and available seats.
        - **Exclusion**: Contains the names of employees who should not have their seats shuffled.
    - Example structure of **Emp_Names**:
        ```
        | Name               | Project                 |
        |--------------------|-------------------------|
        | Anand Kale          | Product Development     |
        | Sonu Vishwakarma    | Product Development     |
        | Shubham Wadhawane   | Miscellaneous, UTI      |
        ```

    - Example structure of **Seat_Capacity**:
        ```
        | Room No. | Seat Count |
        |----------|-------------|
        | 203      | 6           |
        | 213      | 6           |
        | 203.1    | 3           |
        ```

    - Example structure of **Exclusion**:
        ```
        | Name                |
        |---------------------|
        | Shubham Wadhawane    |
        | Pooja Kedari         |
        ```

2. **Teams Output**:
    - After running the script, the seating arrangement will be posted to Microsoft Teams with the title:
      **Seating Arrangements for Today (dd-mm-yyyy)**.
    - The arrangement is presented in a 2-column Adaptive Card format with room numbers and names.

## Example

After running the script, you'll get a seating arrangement like:

```
Room 203: Anand Kale, Sonu Vishwakarma, Pooja Kedari
Room 213: Shantanu Gilbile, Rushikesh Sanap, Yash Jadhav
Room 203.1: Shubham Wadhawane, Ajeet Kumar
```

## Contributing

If you'd like to contribute to this project, feel free to submit a pull request or open an issue on GitHub.

## License

This project is licensed under the MIT License.

---

This README file gives a clear overview of the project, explains the setup process, and provides a good starting point for others to use or contribute to your project on GitHub.
