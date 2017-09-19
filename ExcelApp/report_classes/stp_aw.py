# -*- coding: utf-8 -*
#rid 75 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_aw/'
	base_url += str(params["year"])
	return base_url


#Additional Watering Item
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheet = workbook.add_worksheet()

	data = res
	title = 'Summary of Additional Watering Item'

	#MAIN DATA FORMATING
	format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
	format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
	item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)
	subtotal_format = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT)
	subtotal_format_money = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_MONEY)
	subtitle_format = workbook.add_format(stp_config.CONST.SUBTITLE_FORMAT)

	#set column width
	worksheet.set_column('A:A', 12.67)
	worksheet.set_column('B:B', 12.67)
	worksheet.set_column('C:C', 70)
	worksheet.set_column('D:D', 12)
	#worksheet.set_column('J:J', 9.65)

	#set row
	worksheet.set_row(0,36)
	worksheet.set_row(1,36)
	worksheet.set_row(5,23.4)
	#worksheet.set_row(6, 31.2)

	#HEADER
	#write general header and format
	rightmost_idx = 'D'
	stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

	#additional header image
	worksheet.insert_image('C1', stp_config.CONST.ENV_LOGO,{'x_offset':335,'y_offset':22, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})
	
	#COLUMN NAMES
	item_fields = ['Contract Item No.',	'RIN', 'Location', 'Quantity']
	


	cr = 7 #initiate cr
	"""
	#ORI MAIN DATA
	for idx, val in enumerate(data["items"]):
		for idx2, val2 in enumerate(data["items"][idx]):
			if chr(idx2 + 65) == 'F':
				worksheet.write(chr(idx2+65)+str(idx+8),data["items"][idx][val2], format_num)
				#cr += 1
			elif chr(idx2 + 65 ) != 'F' and chr(idx2 + 65 ) != 'J':
				worksheet.write(chr(idx2+65)+str(idx+8),data["items"][idx][val2], format_text)
		cr += 1
			#cr = idx2 #record the last row num
	"""

	##EDITED MAIN DATA
	#loop over to add distinct 
	mun_list = {}
	total = 0

	for iid, item in enumerate(data["items"]):
		if not data["items"][iid]["municipality"] in mun_list:
			mun_list[data["items"][iid]["municipality"]] = 0
	
	#loop over all programs to write 
	for mid, mun in enumerate(mun_list):

		worksheet.merge_range('A' + str(cr) + ':D' + str(cr), str('Municipality: ' + mun), subtitle_format)
		worksheet.write_row('A' + str(cr+1), item_fields, item_header_format)
		cr += 2

		for idx, val in enumerate(data["items"]):
			if  data["items"][idx]["municipality"] == mun:

				a1 = data["items"][idx]["contract_item_num"]  if "contract_item_num" in data["items"][idx].keys() else ""
				worksheet.write('A' + str(cr), a1 if a1 is not None else "", format_text)

				a2 = data["items"][idx]["rin"] if "rin" in data["items"][idx].keys() else ""
				worksheet.write('B' + str(cr), a2 if a2 is not None else "", format_text)
				
				a3 = data["items"][idx]["location"] if "location" in data["items"][idx].keys() else ""
				worksheet.write('C' + str(cr), a3 if a3 is not None else "", format_text)
				
				a4 = data["items"][idx]["quantity"] if "quantity" in data["items"][idx].keys() else ""
				worksheet.write('D' + str(cr), a4 if a4 is not None else 0, format_num)
				
				cr += 1

				mun_list[mun] += a4
				total += a4

		
		worksheet.write('A' + str(cr), "Total:", subtotal_format) #write total
		worksheet.write_row('B' + str(cr)+':C' + str(cr), ["", ""], subtotal_format)
		worksheet.write('D' + str(cr), mun_list[mun], subtotal_format) #write total
		cr +=1
		worksheet.set_row(cr-1,stp_config.CONST.BREAKDOWN_INBETWEEN_HEIGHT)
		cr += 1

	#cr += 4
	cr += 2


	##============MUNICIPALITY SUMMARY==========
	item_fields = ['Municipality', 'Quantity']
	worksheet.write_row('A' + str(cr), item_fields, item_header_format)
	cr += 1
	#MUN SUMMARY CALCULATION

	
	#calculate number and overall
	if total > 0:
		for mid, mun in enumerate(mun_list):
			worksheet.write('A' + str(cr), mun if mun is not None else "", format_text)
			worksheet.write('B' + str(cr), mun_list[mun] if mun_list[mun] is not None else "", format_num)
			cr += 1
		worksheet.write('A' + str(cr), "Total:", subtotal_format)
		worksheet.write('B' + str(cr), total, subtotal_format)



	#====ending=======

	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data