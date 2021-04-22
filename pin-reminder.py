# -------------------------------------------------#
# Application:  UCXN PIN Reminder
# Author:       Phill Johntony (phidjohn@cdw.com)
# Summary:
# 	Collects mailbox PIN data from Unity Connection
#	Sends an expiration warning email to the end user
#  Creates an Excel report
#  Sends email with report to admins
# ------------------------------------------------#
import sys
import datetime
import time
import re
import configparser
import requests
import json
import pandas
import math
from tqdm import tqdm
import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import logging
import traceback
from urllib3 import disable_warnings
from urllib3.exceptions import InsecureRequestWarning
disable_warnings(InsecureRequestWarning)

#! TO DO LIST
# CLEAN THINGS UP
# ADD IF EMAIL WAS SENT COLUMN TO MAILBOX DF
# CONFIGURE ADMIN EMAIL INTERVALS
# COLLECT EMAIL SENT STATS FOR ADMIN REPORT/EMAIL
# UPDATE README

headers = {
	"content-type": "application/json",
	"accept"      : "application/json",
	"connection"  : "keep-alive"
}

def read_ini(cfg_file_name):
	"""
	Reads config file.

	If successful, returns file contents as a dictionary. Otherwise raise an exception.

	Args:
		cfg_file_name (str): file name to read

	Returns:
		cfg (dict): config file contents
	"""
	try:
		# Check if file exists
		if not os.path.isfile(cfg_file_name): raise Exception("does not exist!")

		config = configparser.ConfigParser()
		config.read(cfg_file_name)
		
		cfg                                       = {}
		cfg["base_url"]                           = config.get('UNITY', 'server')
		cfg["username"]                           = config.get('UNITY', 'username')
		cfg["password"]                           = config.get('UNITY', 'password')
		cfg["smtp_server"]                        = config.get('SMTP', 'server')
		cfg["from_address"]                       = config.get('SMTP', 'from_address')
		cfg["email_intervals"]                    = config.get('SMTP', 'email_intervals')
		cfg["admin_email"]                        = config.get('SMTP', 'admin_email')
		cfg["admin_report_email_file_name"]       = config.get('SMTP', 'admin_report_email_file')
		cfg["user_reminder_email_file_name"]      = config.get('SMTP', 'user_reminder_email_file')
		cfg["user_reminder_attachment_file_name"] = config.get('SMTP', 'user_reminder_attachment')
		cfg["retention_days"]                     = config.get('LOGGING', 'retention_days')
		cfg["debug_lvl"]                          = config.get('DEBUG', 'debug')
		cfg["email_assets_folder_name"]           = "email_assets"
		cfg["reports_folder_name"]                = "reports"
		cfg["logs_folder_name"]                   = "logs"

		return cfg
	except Exception as e:
		logger.error(f"Error in {cfg_file_name} file: {e}")
		sys.exit(1)

def validate_ini(cfg_file_name):
	"""
	Formats and validates config elements

	- Checks for blank values
	- Prefixes https to ucxn server ip/fqdn
	- Creates credential tuple
	- Splits strings with commas into list, then strips leading/trailing whitespace
	- Checks for and creates directories
	- Checks for email assets files
	- Converts retention_days from str to int
	- Changes debug level from default 2 to config value

	Args:
		cfg_file_name (str): file name for print

	Returns:
		cfg (dict): config file contents
	"""
	try:
		for k,v in cfg.items(): # Check for blank values
			if v == "": raise Exception(f"{k} is blank")

		cfg["base_url"]        = "https://" + cfg["base_url"]
		cfg["creds"]           = (cfg["username"], cfg["password"])

		cfg["email_intervals"] = [x.strip() for x in cfg["email_intervals"].split(',')] # splits into list, then strips whitespace
		cfg["admin_email"]     = [x.strip() for x in cfg["admin_email"].split(',')]

		if not os.path.isdir(cfg["email_assets_folder_name"]): os.mkdir(cfg["email_assets_folder_name"])
		if not os.path.isdir(cfg["reports_folder_name"]):      os.mkdir(cfg["reports_folder_name"])

		cfg["admin_report_email_file_fqdn_txt"]  = os.path.join(cfg["email_assets_folder_name"], cfg["admin_report_email_file_name"]+".txt")
		cfg["admin_report_email_file_fqdn_html"] = os.path.join(cfg["email_assets_folder_name"], cfg["admin_report_email_file_name"]+".html")
		if not os.path.isfile(cfg["admin_report_email_file_fqdn_txt"]): raise Exception(f"{cfg['admin_report_email_file_fqdn_txt']} does not exist!")
		if not os.path.isfile(cfg["admin_report_email_file_fqdn_html"]): raise Exception(f"{cfg['admin_report_email_file_fqdn_html']} does not exist!")

		cfg["user_reminder_email_file_fqdn_txt"]  = os.path.join(cfg["email_assets_folder_name"], cfg["user_reminder_email_file_name"]+".txt")
		cfg["user_reminder_email_file_fqdn_html"] = os.path.join(cfg["email_assets_folder_name"], cfg["user_reminder_email_file_name"]+".html")
		if not os.path.isfile(cfg["user_reminder_email_file_fqdn_txt"]): raise Exception(f"{cfg['user_reminder_email_file_fqdn_txt']} does not exist!")
		if not os.path.isfile(cfg["user_reminder_email_file_fqdn_html"]): raise Exception(f"{cfg['user_reminder_email_file_fqdn_html']} does not exist!")

		cfg["user_reminder_attachment_file_fqdn"] = os.path.join(cfg["email_assets_folder_name"], cfg["user_reminder_attachment_file_name"])
		if not os.path.isfile(cfg["user_reminder_attachment_file_fqdn"]): raise Exception(f"{cfg['user_reminder_attachment_file_fqdn']} does not exist!")
		
		cfg["retention_days"] = int(cfg["retention_days"])

		if cfg["debug_lvl"] == "1": # Turn off console debug msgs
			for handler in logger.handlers:
				if type(handler) == logging.StreamHandler:
					handler.setLevel(logging.INFO)

		# for k,v in cfg.items(): logger.debug(f"{k}={v}")
		return cfg
	except ValueError:
		logger.error(f"Error in config file: retention_days must be a number not a string")
		sys.exit(1)
	except Exception as e:
		logger.error(f"Error in {cfg_file_name} file: {e}")
		sys.exit(1)

def init_logger(console_debug_lvl = '1'):
	"""
	Initiates logger

	- Creates log directory if none exists
	- Creates two log handlers
		- One for the log file
		- Another for the console
	- Sets debug level

	Args:
		console_debug_lvl (str): 0 off, 1 on prints only in log file, 2 on prints to log file & console
	"""
	try:
		# Log File Variables
		log_file_dir = "logs"
		log_file_dir = os.path.join(os.getcwd(), log_file_dir)
		log_file_name = 'ucxn-pin-reminder'
		log_file_ext = '.log'
		log_file_date = datetime.datetime.now().strftime("%Y%m%d")
		log_file_time = datetime.datetime.now().strftime("%H%M%S")
		global log_file_fullname
		global log_file_actual
		log_file_fullname = (log_file_name + '-' + log_file_date + '-' + log_file_time + log_file_ext)
		log_file_actual = os.path.join(log_file_dir, log_file_fullname)

		# Create Log File directory if it does not exist
		if not os.path.exists(log_file_dir): os.mkdir(log_file_dir)

		# Global log FILE settings
		log_file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(module)s -> %(message)s')
		log_file_handler   = logging.FileHandler(log_file_actual)
		log_file_handler.setFormatter(log_file_formatter)

		# Global log CONSOLE settings
		log_console_formatter = logging.Formatter('%(asctime)s - %(message)s')
		log_console_handler   = logging.StreamHandler()
		log_console_handler.setFormatter(log_console_formatter)

		if console_debug_lvl == '2':
			# Debug writes to log file AND displays in console
			log_console_handler.setLevel(logging.DEBUG)
			logger.setLevel(logging.DEBUG)
		elif console_debug_lvl == '1':
			# Debug only writes to log file, does not display in console
			log_console_handler.setLevel(logging.INFO)
			logger.setLevel(logging.DEBUG)
		else:
			# Debug is completely off, doesn't write to log file
			log_console_handler.setLevel(logging.INFO)
			logger.setLevel(logging.INFO)

		# Adds configurations to global log
		logger.addHandler(log_file_handler)
		logger.addHandler(log_console_handler)

	except IOError as e:
		errOut = "** ERROR: Unable to create or open log file %s" % log_file_name
		if e.errno == 2:    errOut += "- No such directory **"
		elif e.errno == 13: errOut += " - Permission Denied **"
		elif e.errno == 24: errOut += " - Too many open files **"
		else:
			errOut += " - Unhandled Exception-> %s **" % str(e)
			sys.stderr.write(errOut + "\n")
			traceback.print_exc()

	except Exception:
		traceback.print_exc()

def get_auth_rules():
	"""
	GETs auth rules from UCXN

	If successful, returns response as a list. Otherwise raise an exception.

	Returns:
		authrules (list): with each rule as a dict
	"""
	try:
		url       = f"{cfg['base_url']}/vmrest/authenticationrules"
		response  = requests.get(url=url, auth=cfg["creds"], headers=headers, verify=False)
		if response.status_code != 200: raise Exception(f"Unexpected response from UCXN. Status Code: {response.status_code} Reason: {response.reason}")
		resp_json = response.json()

		authrules = []
		for r in resp_json["AuthenticationRule"]:
			authrules.append({
				"ObjectId"   : r["ObjectId"],
				"DisplayName": r["DisplayName"],
				"MaxDays"    : r["MaxDays"]
			})
		return authrules
	except Exception as e:
		logger.error(f"Error: {e}")
		send_admin_email_error()
		sys.exit(1)

def get_mailboxes():
	"""
	GETs list of mailboxes

	- Initial GET returns total number of mailboxes
	- Calculates how many GETs required to list all mailboxes at 100 per page
	- Performs GET pagination loop

	If successful, returns response as a list. Otherwise raise an exception.

	Returns:
		mailboxes (list): with each mailbox as a dict
	"""
	try:
		#! PAGINATION works, but needs more thorough testing
		url       = f"{cfg['base_url']}/vmrest/users?rowsPerPage=0"
		response  = requests.get(url=url, auth=cfg["creds"], headers=headers, verify=False)
		if response.status_code != 200: raise Exception(f"Unexpected response from UCXN. Status Code: {response.status_code} Reason: {response.reason}")
		resp_json = response.json()
		total_mailboxes = resp_json['@total']
		logger.debug(f"Total Mailboxes: {total_mailboxes}")
		rowsPerPage = 100
		total_pages = math.ceil(int(total_mailboxes) / rowsPerPage)
		logger.debug(f"total_pages = {total_pages}")
		logger.debug("Starting page loop")

		mailboxes = []
		for pageNumber in tqdm(range(total_pages)):
			# print(f"pageNumber = {pageNumber+1}")
			url       = f"{cfg['base_url']}/vmrest/users?rowsPerPage={rowsPerPage}&pageNumber={pageNumber+1}"
			response  = requests.get(url=url, auth=cfg["creds"], headers=headers, verify=False)
			if response.status_code != 200: raise Exception(f"Unexpected response from UCXN. Status Code: {response.status_code} Reason: {response.reason}")
			resp_json = response.json()

			for m in resp_json["User"]:
				mailboxes.append({
					"ObjectId"     : m["ObjectId"],
					"Alias"        : m["Alias"],
					"Display Name" : m["DisplayName"],
					"Extension"    : m["DtmfAccessId"],
					"Email Address": m.get("EmailAddress", "")
				})
		return mailboxes
	except Exception as e:
		logger.error(f"Error: {e}")
		send_admin_email_error()
		sys.exit(1)

def get_pin_data():
	#! Maybe move try/except inside the for loop so only individual mailboxes fail and not all of them?
	"""
	GETs the mailbox PIN data

	- Performs individual GETs for each mailbox to get the PIN data
	- Caclulates PIN expiration dates

	If successful, returns updated mailboxes (list[dict]). Otherwise raise an exception.

	Returns:
		mailboxes (dict)
	"""
	try:
		for m in tqdm(mailboxes):
			url       = f"{cfg['base_url']}/vmrest/users/{m['ObjectId']}/credential/pin"
			response  = requests.get(url=url, auth=cfg["creds"], headers=headers, verify=False)
			if response.status_code != 200: raise Exception(f"Unexpected response from UCXN. Status Code: {response.status_code} Reason: {response.reason}")
			resp_json = response.json()

			for r in authrules:
				if resp_json["CredentialPolicyObjectId"] == r["ObjectId"]:
					m["Auth Rule"]       = r["DisplayName"]
					m["Expiration Days"] = r["MaxDays"]
					break
			m["PIN Doesnt Expire"]  = resp_json["DoesntExpire"]
			m["PIN Must Change"]    = resp_json["CredMustChange"]
			m["Date Last Changed"]  = datetime.datetime.strptime(resp_json["TimeChanged"], "%Y-%m-%d %H:%M:%S.%f")
			m["Expiration Date"]    = m["Date Last Changed"] + datetime.timedelta(days=int(m["Expiration Days"]))
			m["Days Until Expired"] = m["Expiration Date"] - today
			m["Days Until Expired"] = m["Days Until Expired"].days
			m["Date Last Changed"]  = m["Date Last Changed"].date() # Convert datetime to date
			m["Expiration Date"]    = m["Expiration Date"].date()   # Convert datetime to date

		# df = pandas.DataFrame(mailboxes)
		# print(df[['Alias', 'PIN Doesnt Expire', 'PIN Must Change', 'Date Last Changed', 'Expiration Days', 'Expiration Date', 'Days Until Expired']])
		return mailboxes
	except Exception as e:
		logger.error(f"Error: {e}")
		send_admin_email_error()
		sys.exit(1)

def send_user_email():
	#! should update mailbox dict with time last email sent
	"""
	Sends user an expiration email if:

	- PIN Never Expires == false
	- Mailbox has an email address configured
	- If Days Until Expired matches one of the configured email intervals

	If successful, returns updated mailboxes (list[dict]). Otherwise raise an exception.

	Returns:
		mailboxes (dict)
	"""
	for m in tqdm(mailboxes):
		try:
			if m["PIN Doesnt Expire"] == "false" and m["Email Address"] != "":
				if any(str(m["Days Until Expired"]) in s for s in cfg['email_intervals']):
					logger.debug(f"Setting up email for Alias={m['Alias']}")

					if m["Days Until Expired"] > 1:
						days_str = f"{m['Days Until Expired']} days"
					elif m["Days Until Expired"] == 0:
						days_str = "today"
					else:
						days_str = f"{m['Days Until Expired']} day"

					sender    = cfg['from_address']
					receivers = m["Email Address"]

					message            = MIMEMultipart("alternative")
					message["Subject"] = f"{m['Extension']} - Voicemail PIN About to Expire - {m['Expiration Date']}"
					message["From"]    = sender
					message["To"]      = receivers

					text = open(cfg["user_reminder_email_file_fqdn_txt"], "r")
					text = text.read()
					text = text.format(ext=m['Extension'],days=days_str)
					html = open(cfg["user_reminder_email_file_fqdn_html"], "r")
					html = html.read()
					html = html.format(ext=m['Extension'],days=days_str)

					attachment_filename = cfg['user_reminder_attachment_file_name']  # In same directory as script

					# Open file in binary mode
					with open(cfg['user_reminder_attachment_file_fqdn'], "rb") as attachment:
						part_att = MIMEBase("application", "octet-stream") # Add file as application/octet-stream
						part_att.set_payload(attachment.read())            # Email client can usually download this automatically as attachment

					# Encode file in ASCII characters to send by email    
					encoders.encode_base64(part_att)

					# Add header as key/value pair to attachment part
					part_att.add_header(
						"Content-Disposition",
						f"attachment; filename= {attachment_filename}",
					)

					message.attach(MIMEText(text, "plain")) # Add HTML/plain-text parts to MIMEMultipart message
					message.attach(MIMEText(html, "html"))  # The email client will try to render the last part first
					message.attach(part_att)                # Attachment File

					smtpObj = smtplib.SMTP(cfg['smtp_server'])
					smtpObj.sendmail(sender, receivers, message.as_string())         
					logger.debug(f"Successfully sent email to={receivers}")
		except Exception as e:
			logger.error(f"Error: User email was not sent: {e}")

	return mailboxes

def send_admin_email():
	#! Needs to format vars to include number of emails sent and other stats
	"""
	Sends admin email

	Attaches generated report

	"""
	try:
		sender    = cfg['from_address']
		receivers = cfg['admin_email']

		message            = MIMEMultipart("alternative")
		message["Subject"] = "Unity Connection PIN Reminder & Report Tool"
		message["From"]    = sender
		message["To"]      = ", ".join(receivers)

		text = open(cfg["admin_report_email_file_fqdn_txt"], "r")
		text = text.read()
		# text = text.format(ext=m['Extension'],days=days_str)
		html = open(cfg["admin_report_email_file_fqdn_html"], "r")
		html = html.read()
		# html = html.format(ext=m['Extension'],days=days_str)

		attachment_filename = report_filename  # In same directory as script

		# Open file in binary mode
		with open(os.path.join(cfg["reports_folder_name"], attachment_filename), "rb") as attachment:
			part_att = MIMEBase("application", "octet-stream") # Add file as application/octet-stream
			part_att.set_payload(attachment.read())            # Email client can usually download this automatically as attachment

		# Encode file in ASCII characters to send by email    
		encoders.encode_base64(part_att)

		# Add header as key/value pair to attachment part
		part_att.add_header(
			"Content-Disposition",
			f"attachment; filename= {attachment_filename}",
		)

		message.attach(MIMEText(text, "plain")) # Add HTML/plain-text parts to MIMEMultipart message
		message.attach(MIMEText(html, "html"))  # The email client will try to render the last part first
		message.attach(part_att)                # Attachment File

		smtpObj = smtplib.SMTP(cfg['smtp_server'])
		smtpObj.sendmail(sender, receivers, message.as_string())         
		logger.info(f"Admin email successfully sent to: {receivers}")
	except Exception as e:
		logger.error(f"Error: Admin email was not sent: {e}")

def send_admin_email_error():
	"""
	Sends admin email if a crash error occurs

	Attaches log file

	"""
	try:
		sender    = cfg['from_address']
		receivers = cfg['admin_email']

		message            = MIMEMultipart("alternative")
		message["Subject"] = "ERROR - Unity Connection PIN Reminder & Report Tool"
		message["From"]    = sender
		message["To"]      = ", ".join(receivers)

		text = "An error has occurred, see attached log for more details..."
		html = "An error has occurred, see attached log for more details..."

		attachment_filename = log_file_fullname  # In same directory as script

		# Open file in binary mode
		with open(log_file_actual, "rb") as attachment:
			part_att = MIMEBase("application", "octet-stream") # Add file as application/octet-stream
			part_att.set_payload(attachment.read())            # Email client can usually download this automatically as attachment

		# Encode file in ASCII characters to send by email    
		encoders.encode_base64(part_att)

		# Add header as key/value pair to attachment part
		part_att.add_header(
			"Content-Disposition",
			f"attachment; filename= {attachment_filename}",
		)

		message.attach(MIMEText(text, "plain")) # Add HTML/plain-text parts to MIMEMultipart message
		message.attach(MIMEText(html, "html"))  # The email client will try to render the last part first
		message.attach(part_att)                # Attachment File

		smtpObj = smtplib.SMTP(cfg['smtp_server'])
		smtpObj.sendmail(sender, receivers, message.as_string())         
		logger.info(f"Admin error email successfully sent to: {receivers}")
	except Exception as e:
		logger.error(f"Error: Admin error email was not sent: {e}")

def generate_report():
	"""
	Generates PIN report file

	- Creates pandas dataframe from mailboxes (dict)
	- Loads dataframe into ExcelWriter
	- Dynamically adjust all the column lengths
	- Saves as XSLX file

	Returns:
		report_filename (str): filename used for admin email attachment
	"""
	try:
		# Create a Pandas Excel writer using XlsxWriter engine.
		df = pandas.DataFrame(mailboxes)
		del df["ObjectId"]
		report_filename = 'ucxn_voicemail_pin_report_'+datetime.datetime.now().strftime("%Y-%m-%d-%I-%M-%S")+'.xlsx'
		writer = pandas.ExcelWriter(os.path.join(cfg["reports_folder_name"], report_filename), engine='xlsxwriter')
		# Convert the dataframe to an XlsxWriter Excel object.
		df.to_excel(writer, sheet_name='Sheet1', index=False)
		# Dynamically adjust all the column lengths
		for column in df:
			column_length = max(df[column].astype(str).map(len).max(), len(column))
			col_idx = df.columns.get_loc(column)
			writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_length)
		writer.save() # Close the Pandas Excel writer and output the Excel file.
		logger.info(f"Report saved: {report_filename}")
		return report_filename
	except Exception as e:
		logger.error(f"Error: report was not saved: {e}")
		send_admin_email_error()
		sys.exit(1)

def purge_files(retention_days, file_dir, file_ext):
	"""
	Purges files past a certain date
	"""
	try:
		if retention_days > 0:
			logger.debug(f"Purging {file_ext} files in {file_dir} folder older than {retention_days} days...")
			retention_sec = retention_days*86400
			now = time.time()
			for file in os.listdir(file_dir):
				if file.endswith(file_ext):
					file_fullpath = os.path.join(file_dir, file)
					if os.stat(file_fullpath).st_mtime < (now-retention_sec):
						os.remove(file_fullpath)
						logger.debug(f"{file} has been deleted")
	except Exception as e:
		logger.debug("File purge error: " + str(e))

if __name__ == "__main__":
	today = datetime.datetime.today()
	time_start = datetime.datetime.now()

	# Initiate logger
	logger = logging.getLogger('global-log')
	init_logger(console_debug_lvl="2")

	cfg = read_ini("config.ini")
	validate_ini("config.ini")

	logger.info("Step 1 of 6: Getting auth rules...")
	authrules = get_auth_rules()

	logger.info("Step 2 of 6: Getting mailboxes...")
	mailboxes = get_mailboxes()
	
	logger.info("Step 3 of 6: Getting PIN data...")
	get_pin_data()

	logger.info("Step 4 of 6: Sending User Emails...")
	send_user_email()

	logger.info("Step 5 of 6: Saving Report...")
	report_filename = generate_report()

	logger.info("Step 6 of 6: Sending Admin Email...")
	send_admin_email()

	time_end   = datetime.datetime.now()
	time_total = divmod((time_end - time_start).seconds, 60)
	logger.debug(f"Tool Runtime: {time_total[0]} minutes {time_total[1]} seconds")

	purge_files(cfg['retention_days'], cfg["logs_folder_name"], ".log")
	purge_files(cfg['retention_days'], cfg["reports_folder_name"], ".xlsx")