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

headers = {
	"content-type": "application/json",
	"accept"      : "application/json",
	"connection"  : "keep-alive"
}

def read_ini(file_path):
	config = configparser.ConfigParser()
	config.read(file_path)

	cfg = {
		"base_url"       : "https://" + config["UNITY"]["server"],
		"creds"          : (config["UNITY"]["username"], config["UNITY"]["password"]),
		"smtp_server"    : config["SMTP"]["server"],
		"from_address"   : config["SMTP"]["from_address"],
		"email_intervals": config["SMTP"]["email_intervals"].split(","),
		"admin_email"    : config["SMTP"]["admin_email"],
		"retention_days" : int(config["LOGGING"]["retention_days"])
	}
	return cfg

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
		if m["PIN Doesnt Expire"] == "false" and m["Email Address"] != "":
			if any(str(m["Days Until Expired"]) in s for s in cfg['email_intervals']):
				print(f"GOING TO SEND EMAIL TO {m['Alias']}")

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

				text = open("email_reminder.txt", "r")
				text = text.read()
				text = text.format(ext=m['Extension'],days=days_str)
				html = open("email_reminder.html", "r")
				html = html.read()
				html = html.format(ext=m['Extension'],days=days_str)

				attachment_filename = "email_reminder.txt"  # In same directory as script

				# Open file in binary mode
				with open(attachment_filename, "rb") as attachment:
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

				try:
					smtpObj = smtplib.SMTP(cfg['smtp_server'])
					smtpObj.sendmail(sender, receivers, message.as_string())         
					print("Successfully sent email")
				except Exception as e:
					print("Error: unable to send email")
					print(e)

	return mailboxes

def send_admin_email(smtp_server, from_address, admin_email):
	sender    = from_address
	receivers = admin_email

	message = """\
Subject: Hi there

This message is sent from Python."""

	try:
		smtpObj = smtplib.SMTP(smtp_server)
		smtpObj.sendmail(sender, receivers, message)         
		print("Successfully sent email")
	except Exception as e:
		print("Error: unable to send email")
		print(e)

def purge_reports(retention_days):
	try:
		file_ext = ".xlsx"
		if retention_days > 0:
			print(f"Purging XLSX report files older than {retention_days} days...")
			retention_sec = retention_days*86400
			now = time.time()
			for file in os.listdir():
				if file.endswith(file_ext):
					if os.stat(file).st_mtime < (now-retention_sec):
						os.remove(file)
						print(f"{file} has been deleted")
	except Exception as e:
		print("Report cleanup error: " + str(e))

if __name__ == "__main__":
	cfg = read_ini("config.ini")
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
	writer = pandas.ExcelWriter(report_filename, engine='xlsxwriter')
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

	send_admin_email(cfg['smtp_server'], cfg['from_address'], cfg['admin_email'])

	purge_reports(cfg['retention_days'])