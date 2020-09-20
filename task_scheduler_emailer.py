import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import pandas as pd
import subprocess
import smtplib
import sys
import time

# This script will get all Task Scheduler events that returned a non-zero code.
# Scheduled Tasks when run successfully will produce a return code of 0, anything else is viewed by us as an error.
# However, these tasks are stored in the Windows Event Log as 'Information' events, not errors so we need to use PowerShell to query them.
# NOTE: This script should always be run in a directory containing the powershell script 'task_scheduler_errors.ps1'

# Set global date and time variables
# For dates
today_datetime = datetime.date.today()
today_string = today_datetime.strftime('%m/%d/%Y')
yesterday_string = (today_datetime - datetime.timedelta(days=1)).strftime('%m/%d/%Y')
# For times
now_datetime = datetime.datetime.now()
now_time_string = now_datetime.strftime('%H:%M %p')
# Set the datetime variable to a unique string using todays date and time
date_time_range = f'{yesterday_string} {now_time_string} - {today_string} {now_time_string}'
# Set the sender email and password
sender_email = 'my_gmail_username@gmail.com'
sender_password = 'R34LLYG00DP@SSW0RD'
# Server/PC name that this is running for
server_name = 'MY-1337-SERVER'

# logger takes a string and prints it to the console with a datetime
def logger(info):
	print(f'{datetime.datetime.now()}: {info}')

# get_smtp_connection takes no arguments and returns a boolean and an SMTP server connection
def get_smtp_connection():
	# Connect to Google's SMTP server for email
	try:
		server = smtplib.SMTP('smtp.gmail.com', 587)
		server.ehlo()
		server.starttls()
		server.login(sender_email,sender_password)
		logger('Connnected to smtp.gmail.com!')
		return server
	except Exception as e:
		logger(f'ERROR CONNECTING TO SMTP: {e}')
		return None

# run_powershell_script can take a string representation of date like '09/03/2020'. This defaults to a range for
# logging purposes and calls the default PowerShell script which runs for the last 24 hours.
def run_powershell_script(date_string=date_time_range):
	# Set up powershell script call string
	if date_string == date_time_range:
		subject_date = date_time_range
		ps_call = rf"{os.path.join(os.getcwd(), 'task_scheduler_errors.ps1').replace(' ','` ')}" #DEFAULTS TO NOW - 1 DAY (Last 24 Hrs)
	else:
		subject_date = f'{date_string} 12:00 AM - {today_string} {now_time_string}'
		ps_call = rf"{os.path.join(os.getcwd(), 'task_scheduler_errors.ps1').replace(' ','` ')} {date_string}" #RUNS FOR DATE TO TODAY
	logger(f'Running powershell script for {subject_date}...')
	# Call powershell script within current directory
	process = subprocess.run(['powershell.exe', ps_call],
			capture_output=True, # Capture output to a variable instead of print to cmd
			encoding='UTF-8', # Set encoding to UTF-8
		)
	# Set standard out and standard error to variables
	out = process.stdout
	err = process.stderr
	# Format variables for HTML
	out = out.replace("\n","<br />\n")
	err = err.replace("\n","<br />\n")
	# Set return variable
	out_split = out.split('<br />\n<br />\n')
	# Check for errors
	if err != '':
		logger(f'ERROR running PowerShell script! {err}')
		exit(1)
	else:
		logger(f'Powershell script run for {subject_date} successfully.')
		return out_split, subject_date

logger('----------------------------------')
logger('Task Scheduler Emailer Started')

# Set current directory variable
cur_dir = os.getcwd()
logger(f'Current Directory: {cur_dir}')
# Run powershell script for yesterday to get values up to the current time
powershell_output, subject_date = run_powershell_script()

result_string = '''
<html>
<body>
<div contenteditable>
'''
loop_results = ''
count_failed = 0
for event in powershell_output:
	if event != '':
		count_failed += 1
		task_name = event.split('\\')[1].split('"')[0]
		logger(f'Failed Task Name: {task_name}')
		loop_results += '<p>'
		loop_results += f'<b>Failed Batch Script Task #{count_failed} | Task Name: {task_name}</b><br>'
		loop_results += f'{event}'
		loop_results += '</p><br>'

# Log number of failed tasks
logger(f'Failed Tasks: {count_failed}')

# Add failed task results to the top of the email
if count_failed >= 1:
	result_string += (f'<h3 style="color: red">{count_failed} failed batch script task(s) for {subject_date}</h3>')
	result_string += loop_results
	result_string += (f'<br>')
else:
	result_string += (f'<h3 style="color: green">No failed batch script tasks for {subject_date}!</h3>')

# Set up actual email variables
sent_from = 'A Friend'
send_to = ['first_email@aol.com','second_email@hotmail.com']
message = MIMEMultipart("alternative")
message['From'] = sent_from
message['To'] = ', '.join(send_to)
if count_failed > 0:
	message['Subject'] = (f'{server_name} | [ERRORS FOUND] {count_failed} failed batch script task(s) for {subject_date}')
else:
	message['Subject'] = (f'{server_name} | No failed batch script tasks for {subject_date}')

# Set the email signature
result_string += '''\
--
<br>I am a cool dude
<br><b>I work in a cool place</b>
<br><em>I have a cool job</em>
<br>p. (123) 456-7890
<br>e. <a href='mailto:%s'>%s</a>
'''%(sender_email,sender_email)

# Finish result_string HTML by closing tags
result_string += '''
</div>
</body>
</html>
'''

# Declare the message as HTML using our finished result_string
message.attach(MIMEText(result_string,'html'))

server = None
smtp_counter = 0
while not server:
	smtp_counter += 1
	server = get_smtp_connection()
	try:
		server.sendmail(sent_from, send_to, message.as_string())
		logger(f'Email(s) sent to {send_to} successfully.')
		server.close()
		continue
	except Exception as e:
		if server is not None:
			logger(f'ERROR SENDING EMAILS: {e}')

logger(f'SMTP Connection Attempts: {smtp_counter}')
logger(f'{server_name} Task Scheduler Emailer completed SUCCESSFULLY!')
exit(0)
# Program exits before CSV is created, if we want to create this in the future the code will be here.
# Set the file name. You can set path if you specify full path/name.
cur_time = time.strftime("%m%d%Y-%H%M%S")
filename = f'{server_name}_BATCH_TASK_ERRORS_{cur_time}.csv'
# Use pandas to export data as a CSV to the file above
export_csv = df.to_csv(filename, index=None, header=True)
