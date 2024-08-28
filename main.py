import pandas as pd
from datetime import datetime, timedelta

def calculate_pay_bulk(shifts):
    rates = {
        'Standard': {'start': None, 'end': None, 'factor': 1.00, 'amount_per_hour': 17.00},
        'Evening': {'start': datetime.strptime("17:30", "%H:%M").time(), 'end': datetime.strptime("22:00", "%H:%M").time(), 'factor': 1.50, 'amount_per_hour': 25.50},
        'Night': {'start': datetime.strptime("22:00", "%H:%M").time(), 'end': datetime.strptime("06:00", "%H:%M").time(), 'factor': 2.00, 'amount_per_hour': 34.00}
    }

    results = []

    for shift in shifts:
        worker = shift['Worker']
        date = shift['Date']
        start_time_str = shift['Start Time']
        end_time_str = shift['End Time']

        start_time = datetime.strptime(f"{date} {start_time_str}", "%d %b %Y %H:%M")
        end_time = datetime.strptime(f"{date} {end_time_str}", "%d %b %Y %H:%M")

        if end_time < start_time:
            end_time += timedelta(days=1)

        evening_hours = 0
        night_hours = 0
        standard_hours = 0

        if start_time.time() < rates['Night']['end']:
            night_end = datetime.combine(start_time.date(), rates['Night']['end'])
            night_hours = (min(end_time, night_end) - start_time).total_seconds() / 3600
            start_time = night_end  # Update start_time to after the night period

        if start_time.time() < rates['Evening']['end'] and end_time.time() > rates['Evening']['start']:
            evening_start = datetime.combine(start_time.date(), rates['Evening']['start'])
            evening_end = datetime.combine(start_time.date(), rates['Evening']['end'])
            evening_hours = (min(end_time, evening_end) - max(start_time, evening_start)).total_seconds() / 3600

        if start_time.time() >= rates['Night']['start'] or end_time.time() <= rates['Night']['end']:
            night_start = datetime.combine(start_time.date(), rates['Night']['start'])
            night_end = datetime.combine(start_time.date() + timedelta(days=1), rates['Night']['end'])
            night_hours += (min(end_time, night_end) - max(start_time, night_start)).total_seconds() / 3600

        total_shift_hours = (end_time - datetime.strptime(f"{date} {shift['Start Time']}", "%d %b %Y %H:%M")).total_seconds() / 3600
        standard_hours = total_shift_hours - evening_hours - night_hours

        total_pay = (standard_hours * rates['Standard']['amount_per_hour']) + \
                    (evening_hours * rates['Evening']['amount_per_hour']) + \
                    (night_hours * rates['Night']['amount_per_hour'])

        results.append({
            'Worker': worker,
            'Date': date,
            'Standard Hours': round(standard_hours, 2),
            'Evening Hours': round(evening_hours, 2),
            'Night Hours': round(night_hours, 2),
            'Pay': round(total_pay, 2)
        })

    df = pd.DataFrame(results)

    file_path = 'worker_pay_details.xlsx'
    df.to_excel(file_path, index=False)

    return file_path

shifts = [
    {'Worker': 'Kate', 'Date': '15 May 2024', 'Start Time': '08:00', 'End Time': '16:00'},
    {'Worker': 'David', 'Date': '15 May 2024', 'Start Time': '15:00', 'End Time': '23:30'},
    {'Worker': 'Kate', 'Date': '16 May 2024', 'Start Time': '11:30', 'End Time': '18:15'},
    {'Worker': 'David', 'Date': '16 May 2024', 'Start Time': '14:30', 'End Time': '22:00'},
    {'Worker': 'Sam', 'Date': '17 May 2024', 'Start Time': '05:30', 'End Time': '14:10'},
]

excel_file_path = calculate_pay_bulk(shifts)

print(f"Generated Excel file: {excel_file_path}")