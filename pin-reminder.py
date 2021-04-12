import sys
import datetime
import time
import re
import configparser
import requests
import json
import pandas
import math
from urllib3 import disable_warnings
from urllib3.exceptions import InsecureRequestWarning
disable_warnings(InsecureRequestWarning)

headers = {
	"content-type": "application/json",
	"accept"      : "application/json",
	"connection"  : "keep-alive"
}

today = datetime.datetime.today()

def read_ini(file_path):
	config = configparser.ConfigParser()
	config.read(file_path)

	cfg = {
		"base_url": "https://" + config["UNITY"]["server"],
		"creds": (config["UNITY"]["username"], config["UNITY"]["password"])
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
	#! PAGINATION is WIP
	url       = f"{cfg['base_url']}/vmrest/users?rowsPerPage=0"
	response  = requests.get(url=url, auth=cfg["creds"], headers=headers, verify=False)
	resp_json = response.json()
	total_mailboxes = resp_json['@total']
	print(f"Total Mailboxes: {total_mailboxes}")
	rowsPerPage = 100
	total_pages = math.ceil(int(total_mailboxes) / rowsPerPage)
	print(f"total_pages = {total_pages}")
	print("Starting page loop")

	for pageNumber in range(total_pages):
		print(f"pageNumber = {pageNumber+1}")
		url       = f"{cfg['base_url']}/vmrest/users?rowsPerPage={rowsPerPage}&pageNumber={pageNumber+1}"
		response  = requests.get(url=url, auth=cfg["creds"], headers=headers, verify=False)
		resp_json = response.json()

		mailboxes = []
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
	for m in mailboxes:
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

	df = pandas.DataFrame(mailboxes)
	print(df[['Alias', 'PIN Doesnt Expire', 'PIN Must Change', 'Date Last Changed', 'Expiration Days', 'Expiration Date', 'Days Until Expired']])
	return mailboxes

if __name__ == "__main__":
	cfg = read_ini("config.ini")

	print("Getting auth rules...")
	authrules = get_auth_rules()

	print("Getting mailboxes...")
	mailboxes = get_mailboxes()
	
	print("Getting PIN data...")
	get_pin_data()

	df = pandas.DataFrame(mailboxes)
	# Create a Pandas Excel writer using XlsxWriter as the engine.
	writer = pandas.ExcelWriter("report.xlsx", engine='xlsxwriter')
	# Convert the dataframe to an XlsxWriter Excel object.
	df.to_excel(writer, sheet_name='Sheet1', index=False)
	# Get the xlsxwriter workbook and worksheet objects.
	workbook  = writer.book
	worksheet = writer.sheets['Sheet1']
	# Add some cell formats.
	format1 = workbook.add_format({'num_format': '@'})
	# format1.set_num_format('@') # @ - This is text format in excel
	# Set the format but not the column width.
	# worksheet.set_column('A:A', 13)
	# worksheet.set_column('B:B', 17)
	# worksheet.set_column('C:C', 8)
	# worksheet.set_column('D:D', 38)
	# worksheet.set_column('E:E', 13)
	# Close the Pandas Excel writer and output the Excel file.
	writer.save()
	print("Output saved")