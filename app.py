from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)
excel_file = 'counter_data.xlsx'

# Initialize the Excel file if not exists
if not os.path.exists(excel_file):
    df = pd.DataFrame(columns=['Name', 'Count'])
    df.to_excel(excel_file, index=False)


@app.route('/', methods=['GET', 'POST'])
def index():
    total = 0
    message = ""

    if request.method == 'POST':
        name = request.form['name']
        count = int(request.form['count'])

        # Load existing data
        df = pd.read_excel(excel_file)

        # Append new entry
        new_entry = pd.DataFrame([{'Name': name, 'Count': count}])
        df = pd.concat([df, new_entry], ignore_index=True)

        # Save back to Excel
        df.to_excel(excel_file, index=False)

        # Calculate total
        total = df['Count'].sum()
        message = f"Thank you, {name}! The current total count is: {total}"

    else:
        df = pd.read_excel(excel_file)
        total = df['Count'].sum()

    return render_template('form.html', message=message, total=total)


if __name__ == '__main__':
    app.run(debug=True)
