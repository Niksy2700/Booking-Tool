from flask import Flask, render_template, request
from openpyxl import Workbook, load_workbook
import os

app = Flask(__name__)

EXCEL_FILE = 'bookings.xlsx'

# Initialize Excel file if it doesn't exist
if not os.path.exists(EXCEL_FILE):
    wb = Workbook()
    ws = wb.active
    ws.append(["Name", "Room", "Date", "Time", "Duration"])
    wb.save(EXCEL_FILE)

@app.route('/', methods=['GET', 'POST'])
def book_room():
    if request.method == 'POST':
        name = request.form['name']
        room = request.form['room']
        date = request.form['date']
        time = request.form['time']
        duration = request.form['duration']

        # Append to Excel
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        ws.append([name, room, date, time, duration])
        wb.save(EXCEL_FILE)

        return render_template('confirmation.html', name=name, room=room, date=date, time=time, duration=duration)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
