import json
import time

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from hidden_var import json_file_name, spreadsheet_url

def excel():
	scope = [
	'https://spreadsheets.google.com/feeds',
	'https://www.googleapis.com/auth/drive',
	]
	credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
	gc = gspread.authorize(credentials)

	# 스프레스시트 문서 가져오기 
	doc = gc.open_by_url(spreadsheet_url)
	return (doc)


def	get_member_dic(doc):
	worksheet = doc.worksheet('SlackMemberID')

	members = dict()
	num = 1
	while True:
		id = worksheet.cell(num, 1).value
		name = worksheet.cell(num, 2).value
		if id == None:
			break
		members[id] = name
		num = num + 1

		if num % 30 == 0:
			time.sleep(60)
	return (members)


def	attend_yes(doc, members, attendees_id):
	worksheet = doc.worksheet('Test')

	print("id print")

	for id in attendees_id:
		name = members[id]
		cell = worksheet.find(name)
		worksheet.update_cell(cell.row, cell.col + 1, "O")
	
def	attend_no(doc, members, no_attendees_id):
	worksheet = doc.worksheet('Test')

	for id in no_attendees_id:
		name = members[id]
		cell = worksheet.find(name)
		worksheet.update_cell(cell.row, cell.col + 1, "X")
		worksheet.update_cell(cell.row, cell.col + 2, "P")
		worksheet.update_cell(cell.row, cell.col + 3, "P")


# 참여 인원을 Parsing 받아오는 파트.
def parse_attendee(doc, dic_mem):
	with open("../auto_check/check_1/2021-08-29.json") as json_file:
		json_data = json.load(json_file)

	# API 1분당 60개의 제한 -> time delay
	time.sleep(30)

	for message in json_data:
		if "[북라톤 참석 인원 조사]" in message["text"]:
			reactions = message["reactions"]
			for i in reactions:
				if i["name"] == 'ok_hand':
					time.sleep(30)
					attend_yes(doc, dic_mem, i["users"])
				else:
					time.sleep(30)
					attend_no(doc, dic_mem, i["users"])


if __name__ == '__main__':
	doc = excel()
	dic_mem = get_member_dic(doc)

	print("dic print")
	print(dic_mem)

	parse_attendee(doc, dic_mem)
