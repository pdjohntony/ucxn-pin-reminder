import sys
import datetime
import time
import re
import configparser
import requests
import json
import pandas
from urllib3 import disable_warnings
from urllib3.exceptions import InsecureRequestWarning
disable_warnings(InsecureRequestWarning)

headers = {
	"content-type": "application/json",
	"accept": "application/json",
	"connection": "keep-alive"
}

def read_ini(file_path):
	config = configparser.ConfigParser()
	config.read(file_path)

	cfg = {
		"base_url": "https://" + config["UNITY"]["server"],
		"creds": (config["UNITY"]["username"], config["UNITY"]["password"])
	}
	return cfg

cfg = read_ini("config.ini")

print("Getting auth rules...")

url = f"{cfg['base_url']}/vmrest/authenticationrules"
response = requests.get(url=url, auth=cfg["creds"], headers=headers, verify=False)
resp_json = response.json()

authrules = []
for r in resp_json["AuthenticationRule"]:
	authrules.append({
		"ObjectId"   : r["ObjectId"],
		"DisplayName": r["DisplayName"],
		"MaxDays"    : r["MaxDays"]
	})

# df = pandas.DataFrame(authrules)
# print(df)

print("Getting mailboxes data...")

url = f"{cfg['base_url']}/vmrest/users"
response = requests.get(url=url, auth=cfg["creds"], headers=headers, verify=False)
resp_json = response.json()

mailboxes = []
for m in resp_json["User"]:
	mailboxes.append({
		"ObjectId"    : m["ObjectId"],
		"Alias"       : m["Alias"],
		"DisplayName" : m["DisplayName"],
		"Extension"   : m["DtmfAccessId"],
		"EmailAddress": m.get("EmailAddress", "")
	})

print(f"Total Mailboxes: {str(len(mailboxes))}")
print("Getting mailbox PIN data...")

today = datetime.datetime.today()

for m in mailboxes:
	url = f"{cfg['base_url']}/vmrest/users/{m['ObjectId']}/credential/pin"
	response = requests.get(url=url, auth=cfg["creds"], headers=headers, verify=False)
	resp_json = response.json()

	m["DoesntExpire"]     = resp_json["DoesntExpire"]
	m["CredMustChange"]   = resp_json["CredMustChange"]
	m["TimeChanged"]      = datetime.datetime.strptime(resp_json["TimeChanged"], "%Y-%m-%d %H:%M:%S.%f")
	m["TimeExpired"]      = m["TimeChanged"] + datetime.timedelta(days=180)
	m["DaysUntilExpired"] = m["TimeExpired"] - today
	m["DaysUntilExpired"] = m["DaysUntilExpired"].days
	for r in authrules:
		if resp_json["CredentialPolicyObjectId"] == r["ObjectId"]:
			m["AuthRule"] = r["DisplayName"]
			break

x = datetime.datetime.now()
# x2 = datetime.datetime.strptime()

df = pandas.DataFrame(mailboxes)
print(df[['Alias', 'DoesntExpire', 'CredMustChange', 'TimeChanged', 'TimeExpired', 'DaysUntilExpired']])