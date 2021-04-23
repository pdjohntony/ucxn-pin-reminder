# UCXN PIN Reminder
CLI tool that uses the Unity Connection CUPI API to generate mailbox PIN expiration reports and send out PIN expiration warning emails via SMTP

## Table of contents
- [Installation](#installation)
- [Setup](#setup)
- [Usage](#usage)
## Installation
Clone the repository
```
git clone https://github.com/pdjohntony/ucxn-pin-reminder
```
Install the python requirements
```python
pip install -r requirements.txt
```

## Setup

Open the `config.ini` file and fill it out. Most importantly the UCXN server ip/fqdn, credentials, and the SMTP server, from/to addresses.

Optionally you can customize the end user and admin email templates in the `email_assets` folder.

## Usage
```bash
Usage: python pin-reminder.py [OPTION]

Optional Arguments:
  -n, -noemail     generates report but does not send user or admin emails
  -h, -help        display this help and exit
```

`config.ini example`
```ini
[UNITY]
server   = ucxn-1.xyz.com
username = admin
password = 

[SMTP]
server                   = smtp.xyz.com
from_address             = pin-reminder@xyz.com
# days to send expiration emails on, seperate by commas
email_intervals          = 15,5,1,0
# admin email to receive PIN reports, seperate by commas
admin_email              = admin@xyz.com
# specify your email file name located in the "email_assets" folder
# do not include file extension, you need both an html and txt version
# example:
#	/email_assets/
#		user_reminder_template.html
#		user_reminder_template.txt
# config.ini line: user_reminder_email_file = user_reminder_template
admin_report_email_file  = admin_report_template
user_reminder_email_file = user_reminder_template
# specify full file name with file extension for the email attachment
user_reminder_attachment = Changing Your Voicemail PIN.docx

[DEBUG]
# 0 off, 1 on but prints only in log file, 2 on prints to console and log file
debug = 1

[LOGGING]
# the number of days to keep reports
retention_days = 14
```