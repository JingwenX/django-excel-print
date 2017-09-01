# -*- coding: utf-8 -*
#rid 58 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

class Report(object):

	def form_url(params):
		base_url = 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_nursery_requirement_aua/{}'.format(str(params["year"]))
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

		#set column width
		worksheet.set_column('A:A', 31.33)
		worksheet.set_column('B:B', 28)
		worksheet.set_column('C:C', 31.78)
		worksheet.set_column('D:D', 23.56)
		worksheet.set_column('E:E', 9.11)
		worksheet.set_column('F:F', 8.89)
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
		worksheet.insert_image('F1', stp_config.CONST.ENV_LOGO,{'x_offset':35,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})
		
		#COLUMN NAMES
		item_fields = ['Stock Type',	'Plant Type', 'Species', 'Qty Required', 'Qty Tagged', 'Qty Substituted', 'Qty Left to Tag']
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


		for idx, val in enumerate(data["items"][0]["assigned"]):

			
			a1 = data["items"][0]["assigned"][idx]["stock_type"]  if "stock_type" in data["items"][0]["assigned"][idx].keys() else ""
			worksheet.write('A' + str(cr), a1 if a1 is not None else "", format_text)

			a2 = data["items"][0]["assigned"][idx]["plant_type"] if "plant_type" in data["items"][0]["assigned"][idx].keys() else ""
			worksheet.write('B' + str(cr), a2 if a2 is not None else "", format_text)
			
			a3 = data["items"][0]["assigned"][idx]["species"] if "species" in data["items"][0]["assigned"][idx].keys() else ""
			worksheet.write('C' + str(cr), a3 if a3 is not None else "", format_text)
			
			a4 = data["items"][0]["assigned"][idx]["required"] if "required" in data["items"][0]["assigned"][idx].keys() else ""
			worksheet.write('D' + str(cr), a4 if a4 is not None else "", format_text)
			
			a5 = data["items"][0]["assigned"][idx]["tagged"] if "tagged" in data["items"][0]["assigned"][idx].keys() else ""
			worksheet.write('E' + str(cr), a5 if a5 is not None else "", format_text)
			
			a6 = data["items"][0]["assigned"][idx]["substituted"] if "substituted" in data["items"][0]["assigned"][idx].keys() else ""
			worksheet.write('F' + str(cr), a6 if a6 is not None else "", format_text)

			a7 = data["items"][0]["assigned"][idx]["left"] if "left" in data["items"][0]["assigned"][idx].keys() else ""
			worksheet.write('G' + str(cr), a7 if a7 is not None else "", format_text)

			cr += 1
			merge_bottom_idx = cr - 1
			

		cr += 4

		item_fields2 = ['Stock Type',	'Plant Type', 'Species', 'Species Substituted For', 'Qty Tagged']
		worksheet.write_row('A'+str(cr), item_fields2, item_header_format)
		cr+=1

		for idx, val in enumerate(data["items"][0]["unassigned"]):
			if data["items"][0]["unassigned"][idx]["substituted"] >0:
			
				a1 = data["items"][0]["unassigned"][idx]["stock_type"]  if "stock_type" in data["items"][0]["unassigned"][idx].keys() else ""
				worksheet.write('A' + str(cr), a1 if a1 is not None else "", format_text)

				a2 = data["items"][0]["unassigned"][idx]["plant_type"] if "plant_type" in data["items"][0]["unassigned"][idx].keys() else ""
				worksheet.write('B' + str(cr), a2 if a2 is not None else "", format_text)
				
				a3 = data["items"][0]["unassigned"][idx]["species"] if "species" in data["items"][0]["unassigned"][idx].keys() else ""
				worksheet.write('C' + str(cr), a3 if a3 is not None else "", format_text)
				
				a4 = data["items"][0]["unassigned"][idx]["subspecies"] if "subspecies" in data["items"][0]["unassigned"][idx].keys() else ""
				worksheet.write('D' + str(cr), a4 if a4 is not None else "", format_text)
				
				a5 = data["items"][0]["unassigned"][idx]["substituted"] if "substituted" in data["items"][0]["unassigned"][idx].keys() else ""
				worksheet.write('E' + str(cr), a5 if a5 is not None else "", format_text)
			

				cr += 1

		#====ending=======

		workbook.close()

		xlsx_data = output.getvalue()
		return xlsx_data