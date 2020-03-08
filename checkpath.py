import os
from pathlib import Path, PureWindowsPath
import datetime 

report_date=str(datetime.date.today())
directory =PureWindowsPath('../weekly_reports/')
filename = 'weekly_report_' + report_date + '.png'
if not os.path.isdir(directory):
    os.mkdir(directory)

file_path = os.path.join(directory, filename)
print(file_path)

