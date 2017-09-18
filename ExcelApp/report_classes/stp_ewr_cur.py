# -*- coding: utf-8 -*
#rid 74 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config


def form_url(params):
	#base_url = 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_extra_work/{}/{}'.format(str(params["year"]), str(params["assign_num"]))
	base_url = 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_extra_work_current/'
	base_url += str(params["year"])
	#print(base_url)
	return base_url


#Extra Work Report
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
	title = 'Extra Work Payment - Current Payment Assignment'

	#MAIN DATA FORMATING
	format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
	format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
	item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)
	item_format_money = workbook.add_format(stp_config.CONST.ITEM_FORMAT_MONEY)
	subtotal_format_text = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_TEXT)
	subtotal_format = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT)
	subtotal_format_money = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_MONEY)
	subtitle_format = workbook.add_format(stp_config.CONST.SUBTITLE_FORMAT)

	#set column width
	rightmost_idx = 'G'
	stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

	#additional header image
	worksheet.insert_image('D1', stp_config.CONST.ENV_LOGO,{'x_offset':-35,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})
	
	item_fields = ["Contract Item No.", "Location", "Description", "Unit", "Quantity", "Unit Price", "Total Cost"]
	#worksheet.write_row('A1', title, format_text)

	col_wid = [13.22, 54.11, 23.67, 7.33, 8.22, 8.56, 19]
	for i in range (0,ord(rightmost_idx)-64):
	#for i in range (0,19):

		worksheet.set_column(chr(i+65)+':'+chr(i+65), col_wid[i])
	#worksheet.set_column('J:J', 9.65)
	worksheet.set_column('G:G', 12)
	#set row
	worksheet.set_row(0,36)
	worksheet.set_row(1,36)
	worksheet.set_row(5,23.4)
	#worksheet.set_row(6, 31.2)

	#make mun list
	mun_list = []
	for iid, item in enumerate(data["items"]):
		if not data["items"][iid]["municipality"] in mun_list:
			mun_list.append(data["items"][iid]["municipality"])

	cr = 7

	total_to_pay = 0 
	total_payment = 0

	tag_list  = ["contract_item_id", "contract_item_num", "location", "municipality", "description", "measurement", "qty_to_pay", "price", "payment"]
	for mid, mun in enumerate(mun_list):
		worksheet.merge_range('A' + str(cr) + ':' + rightmost_idx + str(cr), str('Municipality: ' + mun), subtitle_format) #was format_text
		worksheet.set_row(cr-1,stp_config.CONST.BREAKDOWN_SUBTITLE_HEIGHT)
		worksheet.write_row('A' + str(cr+1), item_fields, item_header_format)
		cr += 2


		mun_total_qty = 0
		mun_total_payment = 0

		for idx, val in enumerate(data["items"]):
			"""
			for i in range (0,ord(right_most_idx)-65):
				a = data["items"][idx][tag_list[i]] if "seq_id" in data["items"][idx].keys() else ""
				worksheet.write('A1', a if a is not None else "", format_text)
			cr += 1
			"""

			if data["items"][idx]["municipality"] == mun:


				a1 = data["items"][idx]["contract_item_num"]  if "contract_item_num" in data["items"][idx].keys() else ""
				worksheet.write('A' + str(cr), a1 if a1 is not None else "", format_text)

				a2 = data["items"][idx]["location"] if "location" in data["items"][idx].keys() else ""
				worksheet.write('B' + str(cr), a2 if a2 is not None else "", format_text)
				
				a3 = data["items"][idx]["description"] if "description" in data["items"][idx].keys() else ""
				worksheet.write('C' + str(cr), a3 if a3 is not None else "", format_text)
				
				a4 = data["items"][idx]["measurement"] if "measurement" in data["items"][idx].keys() else ""
				worksheet.write('D' + str(cr), a4 if a4 is not None else "", format_text)
				
				a5 = data["items"][idx]["qty_to_pay"] if "qty_to_pay" in data["items"][idx].keys() else ""
				worksheet.write('E' + str(cr), a5 if a5 is not None else 0, format_num)
				
				a6 = data["items"][idx]["price"] if "price" in data["items"][idx].keys() else ""
				worksheet.write('F' + str(cr), a6 if a6 is not None else 0, item_format_money)

				a7 = data["items"][idx]["payment"]  if "payment" in data["items"][idx].keys() else ""
				worksheet.write('G' + str(cr), a7 if a7 is not None else 0, item_format_money)
			
				
				cr += 1
				mun_total_qty += data["items"][idx]["qty_to_pay"] if "qty_to_pay" in data["items"][idx].keys() else 0
				mun_total_payment += data["items"][idx]["payment"] if "payment" in data["items"][idx].keys() else 0
				total_to_pay += data["items"][idx]["qty_to_pay"] if "payment" in data["items"][idx].keys() else 0
				total_payment += data["items"][idx]["payment"] if "payment" in data["items"][idx].keys() else 0

		#cr += 1
		worksheet.write('A' + str(cr), "Total:", subtotal_format_text) #write total
		worksheet.write_row('B' + str(cr)+':G' + str(cr), ["", "", "", "", ""], subtotal_format)
		worksheet.write('E' + str(cr), mun_total_qty, subtotal_format) #write total
		worksheet.write('G' + str(cr), mun_total_payment, subtotal_format_money) #write total
		cr += 1
		worksheet.set_row(cr-1,stp_config.CONST.BREAKDOWN_INBETWEEN_HEIGHT)
		cr += 1
		

	#write grand total
	if total_to_pay != 0:
		worksheet.write('A' + str(cr), 'Grand Total:', subtotal_format_text)
		worksheet.write_row('B' + str(cr)+':D' + str(cr), ["", "", ""], subtotal_format)
		worksheet.write('E' + str(cr), total_to_pay, subtotal_format) #write total
		worksheet.write('F' + str(cr), "", subtotal_format) #write total
		worksheet.write('G' + str(cr), total_payment, subtotal_format_money) #write total


	#====ending=======

	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data