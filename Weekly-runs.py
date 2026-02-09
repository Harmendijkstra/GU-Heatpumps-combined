# Created by Claude.ai

import subprocess
from datetime import datetime, timedelta

FOLDER_MEETSET = 'Meetset2-Nunspeet'
START_DATE = '2025-05-18'
END_DATE = '2025-06-05'
# start = datetime.strptime(input("Start date (YYYY-MM-DD): "), '%Y-%m-%d')
# end = datetime.strptime(input("End date (YYYY-MM-DD): "), '%Y-%m-%d')
start =  datetime.strptime(START_DATE, '%Y-%m-%d')
end =  datetime.strptime(END_DATE, '%Y-%m-%d')

# Find first Monday
monday = start + timedelta(days=(7 - start.weekday()) % 7)

# Loop through weeks
print("Starting weekly runs from", start.strftime('%Y-%m-%d'), " to ", end.strftime('%Y-%m-%d'))
while monday <= end:
    sunday = monday + timedelta(days=6)
    if sunday <= end:
        print("Processing week:", monday.strftime('%Y-%m-%d'), " to ", sunday.strftime('%Y-%m-%d'))
        subprocess.run(['python', 'HPimport.py', FOLDER_MEETSET, 
                       monday.strftime('%Y-%m-%d'), 
                       sunday.strftime('%Y-%m-%d')])
    monday += timedelta(days=7)