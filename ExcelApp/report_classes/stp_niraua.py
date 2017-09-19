# -*- coding: utf-8 -*
#rid 58 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_nursery_aua/{}'.format(str(params["year"]))
	return base_url


#Contract Item Summary - All Items
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
	title = 'Nursery Inspection Requirement - Assigned and Unassigned Species'

	#MAIN DATA FORMATING
	format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
	format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
	item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)
	subtotal_format = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT)

	#set column width
	worksheet.set_column('A:A', 31.33)
	worksheet.set_column('B:B', 28)
	worksheet.set_column('C:C', 31.78)
	worksheet.set_column('D:D', 23.56)
	worksheet.set_column('E:E', 9.11)
	worksheet.set_column('F:F', 12.33)
	worksheet.set_column('G:G', 10.33)
	#worksheet.set_column('J:J', 9.65)

	#set row
	worksheet.set_row(0,36)
	worksheet.set_row(1,36)
	worksheet.set_row(5,23.4)
	worksheet.set_row(6, 31.2)

	#HEADER
	#write general header and format
	rightmost_idx = 'G'
	stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

	#additional header image
	worksheet.insert_image('D1', stp_config.CONST.ENV_LOGO,{'x_offset':135,'y_offset':22, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})
	
	#COLUMN NAMES
	item_fields = ['Stock Type', 'Plant Type', 'Species', 'Qty Required', 'Qty Tagged', 'Qty Substituted', 'Qty Left to Tag']
	worksheet.write_row('A7', item_fields, item_header_format)


	cr = 8 #initiate cr
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
	total_required = 0
	total_tagged = 0
	total_sub = 0
	total_left = 0

	for idx, val in enumerate(data["items"]):
		if data["items"][idx]["table_name"] == "ASSIGNED":
		
			a1 = data["items"][idx]["stock_type"]  if "stock_type" in data["items"][idx].keys() else ""
			worksheet.write('A' + str(cr), a1 if a1 is not None else "", format_text)

			a2 = data["items"][idx]["plant_type"] if "plant_type" in data["items"][idx].keys() else ""
			worksheet.write('B' + str(cr), a2 if a2 is not None else "", format_text)
			
			a3 = data["items"][idx]["species"] if "species" in data["items"][idx].keys() else ""
			worksheet.write('C' + str(cr), a3 if a3 is not None else "", format_text)
			
			a4 = data["items"][idx]["required"] if "required" in data["items"][idx].keys() else ""
			worksheet.write('D' + str(cr), a4 if a4 is not None else "", format_num)
			
			a5 = data["items"][idx]["tagged"] if "tagged" in data["items"][idx].keys() else ""
			worksheet.write('E' + str(cr), a5 if a5 is not None else "", format_num)
			
			a6 = data["items"][idx]["substituted"] if "substituted" in data["items"][idx].keys() else ""
			worksheet.write('F' + str(cr), a6 if a6 is not None else "", format_num)

			a7 = data["items"][idx]["left"] if "left" in data["items"][idx].keys() else ""
			worksheet.write('G' + str(cr), a7 if a7 is not None else "", format_num)

			cr += 1
			merge_bottom_idx = cr - 1
			total_required += data["items"][idx]["required"] if "required" in data["items"][idx].keys() else ""
			total_tagged += data["items"][idx]["tagged"] if "tagged" in data["items"][idx].keys() else ""
			total_sub += data["items"][idx]["substituted"] if "substituted" in data["items"][idx].keys() else ""
			total_left += data["items"][idx]["left"] if "left" in data["items"][idx].keys() else ""
	
	if total_required != 0 or total_tagged != 0 or total_sub != 0 or total_left != 0:
		worksheet.write('A' + str(cr), "Total:", subtotal_format) #write total
		worksheet.write_row('B' + str(cr)+':C' + str(cr), ["", ""], subtotal_format)
		worksheet.write('D' + str(cr), total_required, subtotal_format) #write total
		worksheet.write('E' + str(cr), total_tagged, subtotal_format) #write total
		worksheet.write('F' + str(cr), total_sub, subtotal_format) #write total
		worksheet.write('G' + str(cr), total_left, subtotal_format) #write total

	cr += 4

	cr_unassigned_header = cr

	# item_fields2 = ['Stock Type',	'Plant Type', 'Species', 'Species Substituted For', 'Qty Tagged']
	# worksheet.write_row('A'+str(cr_unassigned_header), item_fields2, item_header_format)
	cr+=1

	sub_table_total = 0

	for idx, val in enumerate(data["items"]):
		if data["items"][idx]["table_name"] == "UNASSIGNED":
			if data["items"][idx]["sub"] >0:
				
				item_fields2 = ['Stock Type',	'Plant Type', 'Species', 'Species Substituted For', 'Qty Tagged']
				worksheet.write_row('A'+str(cr_unassigned_header), item_fields2, item_header_format)
			
				a1 = data["items"][idx]["stock_type"]  if "stock_type" in data["items"][idx].keys() else ""
				worksheet.write('A' + str(cr), a1 if a1 is not None else "", format_text)

				a2 = data["items"][idx]["plant_type"] if "plant_type" in data["items"][idx].keys() else ""
				worksheet.write('B' + str(cr), a2 if a2 is not None else "", format_text)
				
				a3 = data["items"][idx]["species"] if "species" in data["items"][idx].keys() else ""
				worksheet.write('C' + str(cr), a3 if a3 is not None else "", format_text)
				
				a4 = data["items"][idx]["subspecies"] if "subspecies" in data["items"][idx].keys() else ""
				worksheet.write('D' + str(cr), a4 if a4 is not None else "", format_text)
				
				a5 = data["items"][idx]["sub"] if "sub" in data["items"][idx].keys() else ""
				worksheet.write('E' + str(cr), a5 if a5 is not None else "", format_num)
	
				sub_table_total += data["items"][idx]["sub"] if "sub" in data["items"][idx].keys() else 0


				cr += 1

	if sub_table_total != 0:
		worksheet.write('A' + str(cr), "Total:", subtotal_format) #write total
		worksheet.write_row('B' + str(cr)+':D' + str(cr), ["", "", ""], subtotal_format)
		worksheet.write('E' + str(cr), sub_table_total, subtotal_format) #write total

	#====ending=======

	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data