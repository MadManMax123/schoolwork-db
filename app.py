from flask import Flask, render_template, request
import csv
import pandas as pd
import os

app = Flask(__name__)


CSV_FILE = r"C:\Users\satya\OneDrive\My stuff\Personal\Programming\Python\webservers\schoolworkdb\entries.csv"
EXCEL_FILE = r"C:\Users\satya\OneDrive\My stuff\Personal\Programming\Python\webservers\schoolworkdb\entries.xlsx"

COLUMNS = ["Name", "Class", "Section", "Admission Number", "Date", "Day"]

# Add class descriptions dynamically
for i in range(10):
    COLUMNS.append(f"Class{i}")
    COLUMNS.append(f"Description{i}")

COLUMNS += ["HW Given", "Remarks", "Announcements"]

# Ensure CSV and Excel files exist with correct headers
if not os.path.exists(CSV_FILE):
    with open(CSV_FILE, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(COLUMNS)

if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=COLUMNS)
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/', methods=['GET', 'POST'])
def form():
    if request.method == 'POST':
        data = [
            request.form['name'],
            request.form['class'],
            request.form['section'],
            request.form['admission'],
            request.form['date'],
            request.form['day']
        ]
        
        # Collect class-wise data
        for i in range(10):
            data.append(request.form.get(f'class{i}', ''))  # Default to empty string if missing
            data.append(request.form.get(f'description{i}', ''))

        data.append(request.form.get('hw', ''))
        data.append(request.form.get('remarks', ''))
        data.append(request.form.get('announcements', ''))
        
        # Append to CSV
        with open(CSV_FILE, 'a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(data)
        
        # Append to Excel safely
        df = pd.read_excel(EXCEL_FILE)

        # Ensure column consistency before appending
        if len(df.columns) == len(data):  
            df.loc[len(df)] = data
            df.to_excel(EXCEL_FILE, index=False)
        else:
            print("Error: Column mismatch detected!")

        return "Entry recorded successfully!"
    
    subjects = ["Math", "English1", "English2", "6th Subject", "2nd Language", "Biology", "Physics", "Chemistry", "Library", "SUPW", "Games", "History & Civics", "Geography", "CT", "Leadership"]
    return render_template('form.html', subjects=subjects)

if __name__ == '__main__':
    app.run(debug=True)
