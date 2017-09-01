# -*- coding: utf-8 -*
#rid 53 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

class Report(object):

	def form_url(params):
		base_url = 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_nursery_requirement/'
		base_url += str(params["year"])
		return base_url


	#Nursery Tagging Requirement Report
	def render(res, params):

		rid = params["rid"]
		year = params["year"]
		con_num = params["con_num"]
		assign_num = params["assign_num"]

		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()

		data = res
		title = 'Nuersery Tagging Requirement'

		#MAIN DATA FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)

		#set column width
		worksheet.set_column('A:A', 19.67)
		worksheet.set_column('B:B', 24)
		worksheet.set_column('C:C', 23.89)
		worksheet.set_column('D:D', 15.11)
		worksheet.set_column('E:E', 12.33)
		worksheet.set_column('F:F', 15.56)

		#set row
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.set_row(5,23.4)
		worksheet.set_row(6, 31.2)

		#HEADER
		#write general header and format
		rightmost_idx = 'F'
		stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)


		#additional header image
		worksheet.insert_image('D1', stp_config.CONST.ENV_LOGO, {'x_offset':75,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})
		
		#COLUMN NAMES
		item_fields = ['Stock Type', 'Plant Type', 'Species', 'Qty Required', 'Qty Tagged', 'Qty Left To Tag']
		worksheet.write_row('A7', item_fields, item_header_format)

		#MAIN DATA
		for idx, val in enumerate(data["items"]):
			for idx2, val2 in enumerate(data["items"][idx]):
				if chr(idx2 + 65) <= 'C':
					worksheet.write(chr(idx2+65)+str(idx+8),data["items"][idx][val2], format_text)
				elif chr(idx2 + 65 ) > 'C':

					worksheet.write(chr(idx2+65)+str(idx+8),data["items"][idx][val2], format_num)

		workbook.close()

		xlsx_data = output.getvalue()
		return xlsx_data