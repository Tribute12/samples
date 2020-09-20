# Sample Programs

First, open a CMD window here and run:
`python -m pip install -r requirements.txt`

This will install the needed packages.

## Sample Encrypt Zip Email

**Note** that the program will not successfully send an email unless you edit the `GMAIL` variables in the application. Use a valid Gmail account and it should work unless you have 2FA enabled. If so you'll need to set your Gmail account to allow less secure apps to access and then create an app password to use here instead of your usual password.

Then to run the program, use:
`python sample_encrypt_zip_email.py test noout`

The command line arguments are not needed, but not including `test` as the first argument will try to run a connection to a database that doesn't exist. The `noout` argument makes it so that files are not saved to the disk. If you'd like files saved to a directory called `OUTPUT` then leave this line off.

## Sample Task Scheduler Emailer

Edit the Gmail Python in the python script to run successfully.

**Note** that the program will not successfully send an email unless you edit the `GMAIL` variables in the application. Use a valid Gmail account and it should work unless you have 2FA enabled. If so you'll need to set your Gmail account to allow less secure apps to access and then create an app password to use here instead of your usual password.

Then to run the program use:
`python task_scheduler_emailer.py`

The PowerShell script, `task_scheduler_errors.ps1` must exist in the same directory as the Python script.
