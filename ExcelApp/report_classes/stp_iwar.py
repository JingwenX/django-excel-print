# -*- coding: utf-8 -*
#rid 71 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_issued_watering_assignment/{}/{}/{}'.format(str(params["year"]), str(params["assign_num"]), str(params["item_num"]))
	return base_url


#Issued Watering Assignment
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]
	item_num = params["item_num"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheet = workbook.add_worksheet()

	data = res
	title = 'Issued Watering Assignment'

	#MAIN DATA FORMATING
	format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
	format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
	item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)

	#HEADER
	#write general header and format
	rightmost_idx = 'H'
	stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

	#additional header image
	worksheet.insert_image('F1', stp_config.CONST.ENV_LOGO,{'x_offset':180,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

	#set column width


	col_wid = [32.11, 11.44, 45.11, 12.33, 11.44, 8.89, 8.89,8.89]

	for i in range (0,ord(rightmost_idx)-65):
		worksheet.set_column(chr(i+65)+':'+chr(i+65), col_wid[i])

	#set row
	worksheet.set_row(0,36)
	worksheet.set_row(1,36)
	worksheet.set_row(5,23.4)
	worksheet.set_row(6, 31.2)

	#CREATE MUN LIST
	mun_list = []

	for iid, item in enumerate(data["items"]):
		if not data["items"][iid]["municipality"] in mun_list:
			mun_list.append(data["items"][iid]["municipality"])

	cr = 8
	tag_list  = ["watering_item_id", "rin", "location", "road_side", "broadleaved", "conifers", "other_trees", "total_items"]
	for munidx, mun in enumerate(mun_list):
		worksheet.write('A' + str(cr), "Municipality:"+mun, format_text)
		cr +=1
		title= ["Watering Item No.", "RIN", "Location", "Roadside", "No. of Broadleaved", "No. of Conifers", "No. of Others", "Total No. of Tree"]
		worksheet.write_row('A' + str(cr), title, item_header_format)
		cr += 1

		for idx, val in enumerate(data["items"]):
			"""
			for i in range (0,ord(right_most_idx)-65):
				a = data["items"][idx][tag_list[i]] if "seq_id" in data["items"][idx].keys() else ""
				worksheet.write('A1', a if a is not None else "", format_text)
			cr += 1
			"""
			if data["items"][idx]["municipality"] == mun:
				

				a1 = data["items"][idx]["watering_item_id"]  if "watering_item_id" in data["items"][idx].keys() else ""
				worksheet.write('A' + str(cr), a1 if a1 is not None else "", format_text)

				a2 = data["items"][idx]["rin"] if "rin" in data["items"][idx].keys() else ""
				worksheet.write('B' + str(cr), a2 if a2 is not None else "", format_text)
				
				a3 = data["items"][idx]["location"] if "location" in data["items"][idx].keys() else ""
				worksheet.write('C' + str(cr), a3 if a3 is not None else "", format_text)
				
				a4 = data["items"][idx]["road_side"] if "road_side" in data["items"][idx].keys() else ""
				worksheet.write('D' + str(cr), a4 if a4 is not None else "", format_text)
				
				a5 = data["items"][idx]["broadleaved"] if "broadleaved" in data["items"][idx].keys() else ""
				worksheet.write('E' + str(cr), a5 if a5 is not None else "", format_text)
				
				a6 = data["items"][idx]["conifers"] if "conifers" in data["items"][idx].keys() else ""
				worksheet.write('F' + str(cr), a6 if a6 is not None else "", format_text)

				a7 = data["items"][idx]["other_trees"]  if "other_trees" in data["items"][idx].keys() else ""
				worksheet.write('G' + str(cr), a7 if a7 is not None else "", format_text)

				a8 = data["items"][idx]["total_items"] if "total_items" in data["items"][idx].keys() else ""
				worksheet.write('H' + str(cr), a8 if a8 is not None else "", format_text)
				
				cr += 1
		
		cr += 2
		

	cr += 4


	#====ending=======

	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data