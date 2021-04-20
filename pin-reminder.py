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
from urllib3 import disable_warnings
from urllib3.exceptions import InsecureRequestWarning
disable_warnings(InsecureRequestWarning)

#! TO DO LIST
# CLEAN THINGS UP
# IMPLEMENT LOGGER
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
	try:
		if not os.path.isfile(cfg_file_name): # Check if file exists
			print(f"{cfg_file_name} does not exist!")
			sys.exit(1)

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
		cfg["email_assets_folder_name"]           = "email_assets"
		cfg["reports_folder_name"]                = "reports"
		cfg["logs_folder_name"]                   = "logs"

		return cfg
	except Exception as e:
		print(f"Error in {cfg_file_name} file: {e}")
		sys.exit(1)

def validate_ini(cfg_file_name):
	try:
		for k,v in cfg.items(): # Check for blank values
			if v == "": raise Exception(f"{k} is blank")

		cfg["base_url"]                 = "https://" + cfg["base_url"]
		cfg["creds"]                    = (cfg["username"], cfg["password"])

		cfg["email_intervals"]          = [x.strip() for x in cfg["email_intervals"].split(',')] # splits into list, then strips whitespace
		cfg["admin_email"]              = [x.strip() for x in cfg["admin_email"].split(',')]

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

		cfg["user_reminder_attachment_file_fqdn"]  = os.path.join(cfg["email_assets_folder_name"], cfg["user_reminder_attachment_file_name"])
		if not os.path.isfile(cfg["user_reminder_attachment_file_fqdn"]): raise Exception(f"{cfg['user_reminder_attachment_file_fqdn']} does not exist!")
		
		cfg["retention_days"]           = int(cfg["retention_days"])

		for k,v in cfg.items(): print(f"{k}={v}")
		return cfg
	except ValueError:
		print(f"Error in config file: retention_days must be a number not a string")
		sys.exit(1)
	except Exception as e:
		print(f"Error in {cfg_file_name} file: {e}")
		sys.exit(1)

def get_auth_rules():
	url       = f"{cfg['base_url']}/vmrest/authenticationrules"
	response  = requests.get(url=url, auth=cfg["creds"], headers=headers, verify=False)
	resp_json = response.json()

	authrules = []
	for r in resp_json["AuthenticationRule"]:
		authrules.append({
			"ObjectId"   : r["ObjectId"],
			"DisplayName": r["DisplayName"],
			"MaxDays"    : r["MaxDays"]
		})
	return authrules

def get_mailboxes():
	#! PAGINATION works, but needs more thorough testing
	url       = f"{cfg['base_url']}/vmrest/users?rowsPerPage=0"
	response  = requests.get(url=url, auth=cfg["creds"], headers=headers, verify=False)
	resp_json = response.json()
	total_mailboxes = resp_json['@total']
	print(f"Total Mailboxes: {total_mailboxes}")
	rowsPerPage = 100
	total_pages = math.ceil(int(total_mailboxes) / rowsPerPage)
	print(f"total_pages = {total_pages}")
	print("Starting page loop")

	mailboxes = []
	for pageNumber in tqdm(range(total_pages)):
		# print(f"pageNumber = {pageNumber+1}")
		url       = f"{cfg['base_url']}/vmrest/users?rowsPerPage={rowsPerPage}&pageNumber={pageNumber+1}"
		response  = requests.get(url=url, auth=cfg["creds"], headers=headers, verify=False)
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

def get_pin_data():
	for m in tqdm(mailboxes):
		url       = f"{cfg['base_url']}/vmrest/users/{m['ObjectId']}/credential/pin"
		response  = requests.get(url=url, auth=cfg["creds"], headers=headers, verify=False)
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

def analyze_pin_data():
	for m in tqdm(mailboxes):
		try:
			if m["PIN Doesnt Expire"] == "false" and m["Email Address"] != "":
				if any(str(m["Days Until Expired"]) in s for s in cfg['email_intervals']):
					print(f"\nSetting up email for Alias={m['Alias']}")

					if m["Days Until Expired"] > 1:
						days_str = f"{m['Days Until Expired']} days"
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
					print(f"Successfully sent email to={receivers}")
		except Exception as e:
			print(f"Error: User email was not sent: {e}")

	return mailboxes

def send_admin_email():
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
		print(f"Admin email successfully sent to: {receivers}")
	except Exception as e:
		print(f"Error: Admin email was not sent: {e}")

def purge_files(retention_days, file_dir, file_ext):
	try:
		if retention_days > 0:
			print(f"Purging {file_ext} files in {file_dir} folder older than {retention_days} days...")
			retention_sec = retention_days*86400
			now = time.time()
			for file in os.listdir(file_dir):
				if file.endswith(file_ext):
					file_fullpath = os.path.join(file_dir, file)
					if os.stat(file_fullpath).st_mtime < (now-retention_sec):
						os.remove(file_fullpath)
						print(f"{file} has been deleted")
	except Exception as e:
		print("File purge error: " + str(e))

if __name__ == "__main__":
	cfg = read_ini("config.ini")
	validate_ini("config.ini")

	today = datetime.datetime.today()
	time_start = datetime.datetime.now()
	print(time_start)

	print("Getting auth rules...")
	authrules = get_auth_rules()

	print("Getting mailboxes...")
	mailboxes = get_mailboxes()
	
	print("Getting PIN data...")
	get_pin_data()

	print("Analyzing PIN data...")
	analyze_pin_data()

	# Create a Pandas Excel writer using XlsxWriter as the engine.
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
	print(f"Report saved: {report_filename}")

	time_end = datetime.datetime.now()
	print(time_end)

	send_admin_email()

	purge_files(cfg['retention_days'], cfg["logs_folder_name"], ".log")
	purge_files(cfg['retention_days'], cfg["reports_folder_name"], ".xlsx")