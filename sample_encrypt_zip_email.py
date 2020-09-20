"""
This script can be used as a base to read data from a database into a Pandas dataframe,
 manipulate that data, then zip and encrypt it in memory, then send that as an email
 attachment.
"""
import pyodbc, time # pyodbc is used for the database connections, time is for datetime in the filename
import pandas as pd # pandas allows us to use dataframes and other useful tools
import os			# os allows us to interact with the operating system
import smtplib		# allows us to send email over SMTP
import sys          # Used to get CLI args
# For email
from email.encoders import encode_base64
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from mimetypes import guess_type
# For file zip and encryption
from io import BytesIO
from base64 import encodebytes
import pyzipper
# For fake dataframe
import numpy as np
from faker.providers.person.en import Provider

# Prints strings to stdout with a timestamp
def log(log_string):
	print(f'{time.strftime("%m%d%Y-%H%M%S")}: {log_string}')

# Prints errors to stdout for logging, then exits with non-zero exit code
def error(error_string, exception):
	log(f'ERROR: {error_string}')
	log(exception)
	log('EXITING PROGRAM!')
	exit(1)

# Functions for generating fake data using Faker
def random_names(name_type, size):
	"""
	Generate n-length ndarray of person names.
	name_type: a string, either first_names or last_names
	"""
	names = getattr(Provider, name_type)
	return np.random.choice(names, size=size)

def random_genders(size, p=None):
	"""Generate n-length ndarray of genders."""
	if not p:
	    # default probabilities
	    p = (0.49, 0.49, 0.01, 0.01)
	gender = ("M", "F", "O", "")
	return np.random.choice(gender, size=size, p=p)

def random_dates(start, end, size):
	"""
	Generate random dates within range between start and end.
	Adapted from: https://stackoverflow.com/a/50668285
	"""
	# Unix timestamp is in nanoseconds by default, so divide it by
	# 24*60*60*10**9 to convert to days.
	divide_by = 24 * 60 * 60 * 10**9
	start_u = start.value // divide_by
	end_u = end.value // divide_by
	return pd.to_datetime(np.random.randint(start_u, end_u, size), unit="D")

# Check arguments in function to allow args in any order
def arg_checker(argument):
	# Check for test flag
	if argument == 'test':
		test_flag = True
	else:
		test_flag = False
	# Check for noout flag
	if argument == 'noout':
		no_out_flag = True
	else:
		no_out_flag = False

# Checks for command line argument 1
try:
	arg_1 = sys.argv[1]
	arg_checker(arg_1)
except IndexError:
	log('First CLI argument was not inlcuded.')

# Checks for command line argument 2
try:
	arg_2 = sys.argv[2]
	arg_checker(arg_2)
except IndexError:
	log('A second CLI argument was not included.')

# Logs flags
if test_flag:
	log(f'Using test flag. Program will not query a database.')
if no_out_flag:
	log(f'Using no output flag. Program will not write to disk.')

# GMail Connection
GMAIL_USERNAME = 'username@gmail.com'
GMAIL_PASSWORD = 'password123'
GMAIL_DISPLAY_NAME = 'My Username'
def get_smtp_connection():
	# Connect to Google's SMTP server for email
	try:
		server = smtplib.SMTP('smtp.gmail.com', 587)
		server.ehlo()
		server.starttls()
		server.login(GMAIL_USERNAME, GMAIL_PASSWORD)
		log(f'Connnected to smtp.gmail.com!')
		return True, server
	except Exception as e:
		error(f'COULD NOT CONNECT TO SMTP!', e)
		return False, None

# Database Connection
DB_DRIVER = '{ODBC Driver 17 for SQL Server}'
DB_SERVER = '192.168.1.1'
DB_DATABASE = 'dbo'
DB_USERNAME = 'db_username'
DB_PASSWORD = 'db_password'

log('-------------------------------------')
log(f'{__file__} Script Started.')

if not test_flag:
	# Attempt DB Connection
	try:
		connection = pyodbc.connect(driver=DB_DRIVER, server=DB_SERVER, database=DB_DATABASE,
			uid=DB_USERNAME, pwd=DB_PASSWORD)
		log('Connected to database Successfully.')
	except pyodbc.Error as ex:
		sqlstate = ex.args[1]
		error('Could not connect to database!',f'{sqlstate}')

	# Primary Frontline query
	initial_script = '''
	SELECT DISTINCT e.a_employee_number AS 'EmployeeID'
	, CAST(j2.a_job_class_code AS INT) AS 'job_class'
	, CASE
		WHEN CURRENT_TIMESTAMP < DATEFROMPARTS(YEAR(CURRENT_TIMESTAMP), 7, 12)
			THEN FORMAT(DATEADD(YEAR, -1, DATEFROMPARTS(YEAR(CURRENT_TIMESTAMP), 8, 15)), 'MM/dd/yyyy')
		ELSE FORMAT(DATEFROMPARTS(YEAR(CURRENT_TIMESTAMP), 8, 15), 'MM/dd/yyyy')
	END AS 'EvaluationCycleStartDate'
	, CASE
		WHEN CURRENT_TIMESTAMP < DATEFROMPARTS(YEAR(CURRENT_TIMESTAMP), 7, 12)
			THEN FORMAT(DATEFROMPARTS(YEAR(CURRENT_TIMESTAMP), 7, 12), 'MM/dd/yyyy')
		ELSE FORMAT(DATEADD(YEAR, 1, DATEFROMPARTS(YEAR(CURRENT_TIMESTAMP), 7, 12)), 'MM/dd/yyyy')
	END AS 'EvaluationCycleEndDate'
	FROM pr_employee_master e
	LEFT JOIN pr_job_pay j2
		ON j2.a_employee_number = e.a_employee_number
		AND j2.a_job_class_code = e.a_job_code_primary
		AND j2.a_projection = 0
		AND j2.s_inactive = 'A'
		AND j2.a_base_pay = 'Y'
	WHERE (e.e_activity_status IN ('A','B')
	-- Includes employees terminated/inactivated in last 30 days
		OR (e.e_activity_status = 'I'
			AND ((e.e_inactive_date >= DATEADD(DAY, -30, CURRENT_TIMESTAMP))
				OR (e.e_terminated_date >= DATEADD(DAY, -30, CURRENT_TIMESTAMP))))
	-- Includes employees that are marked with PD Only bargain unit
		-- OR (e.e_activity_status = 'I'
		-- 	AND e.a_bargain_primary = 'CHAR')
	)
	AND e.a_projection = 0
	AND j2.a_job_class_code IS NOT NULL
	'''

	try:
		# Set query results to a variable and read using pandas
		log('Running main query...')
		results = pd.read_sql_query(initial_script,connection)
		# Close connections
		connection.close()
		log('Connection closed.')
		# Format the results in a data frame
		main_dataframe = pd.DataFrame(results)
		log('Query results gathered into pandas dataframe.')
	except Exception as e:
		error('Could not gather query results to pandas dataframe.',e)
elif test_flag:
	try:
		log('Creating test dataframe...')
		size = 100
		main_dataframe = pd.DataFrame(columns=['First', 'Last', 'Gender', 'Birthdate'])
		main_dataframe['First'] = random_names('first_names', size)
		main_dataframe['Last'] = random_names('last_names', size)
		main_dataframe['Gender'] = random_genders(size)
		main_dataframe['Birthdate'] = random_dates(start=pd.to_datetime('1940-01-01'), end=pd.to_datetime('2008-01-01'), size=size)
		log('Test dataframe created.')
	except Exception as e:
		error(f'Could not create test dataframe!', e)

# Log current directory
log(f'Current Directory: {os.getcwd()}')
# Set the datetime variable to a unique string using todays date and time
file_datetime = time.strftime("%m%d%Y-%H%M%S")

# For outputting files to disk

if not no_out_flag:
	# Check for needed directories, create them if they do not exist
	if not os.path.exists('OUTPUT'):
		os.makedirs('OUTPUT')
		log('OUTPUT directory created.')
	if not os.path.exists('ARCHIVE'):
		os.makedirs('ARCHIVE')
		log('ARCHIVE directory created.')

	# Set the relative paths for output files
	xlsx_filename = (f'OUTPUT\\Test_File_{file_datetime}.xlsx')
	csv_filename = (f'OUTPUT\\Test_File_{file_datetime}.csv')

	# Check for existing files in the OUTPUT directory, if so, archive them.
	try:
		for file in os.listdir('OUTPUT'):
			# Rename file(s)
			os.rename(f'OUTPUT\\{file}',f'ARCHIVE\\{file}')
			log(f'{file} moved to archive as {file}.')
	except Exception as e:
		log(f'There was a problem moving old files to archive!')
		log(e)

	try:
		# Set needed columns to list
		if test_flag:
			columns = ['First'
				,'Last'
				,'Gender'
				,'Birthdate'
			]
		else:
			columns = ['EmployeeID'
				,'EvalType'
				,'EvaluationCycleStartDate'
				,'EvaluationCycleEndDate'
			]
		# Reorder columns and set dataframe to desired columns only
		main_dataframe = main_dataframe[columns]
		# Use pandas to export data as a CSV to the filename above
		export_csv = main_dataframe.to_excel(xlsx_filename, index=None, header=True)
		log(f'{xlsx_filename} created successfully.')
		main_dataframe.to_csv(csv_filename, index=None, header=True, encoding='utf-8')
		log(f'{csv_filename} created successfully.')
	except Exception as e:
		error(f'{csv_filename} NOT CREATED!', e)

# Start email portion of program

# Set today's date and datetime to variable
date = time.strftime('%m/%d/%Y')
date_time = time.strftime('%m/%d/%Y %H:%M')

# Set the email variables
log(f'Setting email variables...')
# Who will be displayed as the sender of the email?
sent_from = GMAIL_DISPLAY_NAME
# Who will receive the email? Separate multiple recipients with commas.
send_to = [
	'christopher_widak@charleston.k12.sc.us'
	# 'something@gmail.com'
	# ,'something_bad@aol.com'
]
# What is the subject of the email? Use standard formatting.
subject = f'Test Email For {date}'
# Data to zip and attach
attachment_data = main_dataframe.to_csv(index=None)
# What is the encryption password?
zip_password = 'test'
# What is the name of the file to be zipped?
zip_inner_filename = (f'Test_File_{file_datetime}.csv')
# What is the name of the zip file?
zip_filename = (f'Test_File_{file_datetime}.zip')
# What is the heading within the email?
email_heading = f'This email brought to you by {GMAIL_DISPLAY_NAME}'
# What is the subheading within the email?
email_subheading = f'Something useful about the file.'
# What is the body of the email?
email_body = f'Attached is the import file containing information as of {date_time}.'
# What is the email disclaimer?
email_disclaimer = f'If you are not the intended recipient of this email you must delete it immediately. If you have mistakenly downloaded any files attached with this email you must delete them immediately.'
log(f'Email variables set.')

log(f'Setting up email message parts...')
# Set up actual email message
message = MIMEMultipart("alternative")
message['From'] = sent_from
message['To'] = ', '.join(send_to)
message['Subject'] = subject
attachment = MIMEBase('application', 'zip')
# Create encrypted version of CSV and save in memory for email
# Set encryption password
encryption_password = zip_password.encode()
# Set up BytesIO object to store attachment in memory
try:
	log(f'Creating and zipping file in memory...')
	with BytesIO() as file_in_memory:
		# Set up pyzipper to create zipped file
		with pyzipper.AESZipFile(
				file_in_memory
				,'w'
				,compression=pyzipper.ZIP_LZMA) as packet:
			# Set the encryption password
			packet.setpassword(encryption_password)
			# Set the encryption method
			packet.setencryption(pyzipper.WZ_AES, nbits=256)
			# Write the dataframe to CSV with the zip_inner_filename
			packet.writestr(zip_inner_filename, attachment_data.encode())
			# Set the payload of the attachment to the zip file in memory using BytesIO
			attachment.set_payload(encodebytes(file_in_memory.getbuffer()))
			# Add the header so GMail knows that this is the expected filetype
			attachment.add_header('Content-Transfer-Encoding', 'base64')
			# Set the attachment information in the header
			attachment.add_header('Content-Disposition', 'attachment', filename=zip_filename)
			# Attach the file to the message
			message.attach(attachment)
	log(f'File created, zipped and attached.')
except Exception as e:
	error(f'Could not create and attach the file!', e)

log(f'Creating email HTML...')
# Set the email body starting with HTML tags
email_string = '''
<html>
<body>
<div>
'''
# Set email body information
email_string += '''
<h3>%s</h3>
<h4>%s</h4>
'''%(email_heading, email_subheading)
email_string += f'<p>{email_body}</p>'
# Set disclaimer
email_string += '''
<br>
<small><b style="color: red;">%s</b></small>
<br>
'''%(email_disclaimer)
# Set the email signature
email_string += '''\
--
<br>%s
<br><b>I work here</b>
<br><em>Something cool</em>
<br>p. (123) 456-7890
<br>e. <a href='mailto:%s'>%s</a>
'''%(GMAIL_DISPLAY_NAME, GMAIL_USERNAME, GMAIL_USERNAME)
# Finish email_string HTML by closing tags
email_string += '''
</div>
</body>
</html>
'''
# Declare the message as HTML using our finished email_string
message.attach(MIMEText(email_string,'html'))
log(f'Email HTML created.')
# Send email in while loop to ensure sent
smtp_connected = False
smtp_counter = 0
while not smtp_connected:
	smtp_counter += 1
	smtp_connected, server = get_smtp_connection()
	try:
		server.sendmail(sent_from, send_to, message.as_string())
		log(f'Email(s) sent to {send_to} successfully.')
		server.close()
		continue
	except Exception as e:
		smtp_connected = False
		if server is not None:
			error(f'Could not send emails!', e)

log(f'SMTP Connection Attempts: {smtp_counter}')
# EOF
log(f'{__file__} completed successfully!')
