# -*- coding: utf-8 -*
#rid 51 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_contractor_plant_tree/'
	base_url += str(params["year"])
	return base_url


#Contractor Plants Trees
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheet = workbook.add_worksheet()

	data = res
	title = 'Tree Planting Status'
	#year = year #delete

	# MAIN DATA FORMATING
	format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
	format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
	item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)

	#set column width
	worksheet.set_column('A:A', 15.78)
	worksheet.set_column('B:B', 31.44)
	worksheet.set_column('C:C', 12.22)
	worksheet.set_column('D:D', 11.89)
	worksheet.set_column('E:E', 9.78)
	worksheet.set_column('F:F', 11.33)
	worksheet.set_column('G:G', 11.89)
	worksheet.set_column('H:H', 11)

	#set row
	worksheet.set_row(0,36)
	worksheet.set_row(1,36)
	worksheet.set_row(5,23.4)
	worksheet.set_row(6, 31.2)

	#HEADER
	#write general header and format
	rightmost_idx = 'I'
	stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num) #change 2017 to year

	# FILE SPECIFIC FORMATING
	format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
	item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)

	#additional header image
	worksheet.insert_image('F1', r'\\ykr-apexp1\staticenv\apps\199\env_internal_bw.png',{'x_offset':70,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})


	item_fields = ['Tree Planting Detail No.', 'Location', 'Assignment No.', 'Assignment Status', 'Planting Status', 'Planting Start Date', 'Planting End Date', 'Assigned Inspector', 'Status of Inspection']
	worksheet.write_row('A7', item_fields, item_header_format)


	#MAIN DATA
	cr = 8
	for idx, val in enumerate(data["items"]):
		#for idx2, val2 in enumerate(data["items"][idx]):
		#	worksheet.write(chr(idx2+65)+str(idx+8),data["items"][idx][val2], format_text)
		a1 = data["items"][idx]["tree_planting_detail_no"] if "tree_planting_detail_no" in data["items"][idx].keys() else ""
		worksheet.write('A' + str(cr), a1 if a1 is not None else "", format_text)

		a2 = data["items"][idx]["location"] if "location" in data["items"][idx].keys() else ""
		worksheet.write('B' + str(cr), a2 if a2 is not None else "", format_text)

		a3 = data["items"][idx]["assignment_num"] if "assignment_num" in data["items"][idx].keys() else ""
		worksheet.write('C' + str(cr), a3 if a3 is not None else "", format_num)

		a4 = data["items"][idx]["assignment_status"] if "assignment_status" in data["items"][idx].keys() else ""
		worksheet.write('D' + str(cr), a4 if a4 is not None else "", format_text)

		a5 = data["items"][idx]["planting_status"] if "planting_status" in data["items"][idx].keys() else ""
		worksheet.write('E' + str(cr), a5 if a5 is not None else "", format_text)

		a6 = data["items"][idx]["start_date"] if "start_date" in data["items"][idx].keys() else ""
		worksheet.write('F' + str(cr), a6 if a6 is not None else "", format_text)

		a7 = data["items"][idx]["end_date"] if "end_date" in data["items"][idx].keys() else ""
		worksheet.write('G' + str(cr), a7 if a7 is not None else "", format_text)

		a8 = data["items"][idx]["inspector"] if "inspector" in data["items"][idx].keys() else ""
		worksheet.write('H' + str(cr), a8 if a8 is not None else "", format_text)

		a9 = data["items"][idx]["inspection_status"] if "inspection_status" in data["items"][idx].keys() else""
		worksheet.write('I' + str(cr), a9 if a9 is not None else "", format_text)
		cr += 1

	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data