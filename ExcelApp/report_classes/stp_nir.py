# -*- coding: utf-8 -*
#rid 6 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

class Report(object):

	def form_url(params):
		base_url = 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_nusery_inspection/'
		base_url += str(params["year"])
		return base_url


	#Nursery Inspection Report
	def render(res, params):

		rid = params["rid"]
		year = params["year"]
		con_num = params["con_num"]
		assign_num = params["assign_num"]

		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()

		data = res
		title = 'Nuersery Inspection Report'
		

		#set column width
		worksheet.set_column('A:A', 11.22)
		worksheet.set_column('B:B', 8.67)
		worksheet.set_column('C:C', 33)
		worksheet.set_column('D:D', 35.22)
		worksheet.set_column('E:E', 12.22)
		worksheet.set_column('F:F', 19.22)
		worksheet.set_column('G:G', 10.89)
		worksheet.set_column('H:H', 10.67)

		#set row
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.set_row(5,23.4)
		worksheet.set_row(6, 31.2)

		#HEADER
		#write general header and format
		rightmost_idx = 'H'
		stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

		# FILE SPECIFIC FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)

		#additional header image
		worksheet.insert_image('F1', r'\\ykr-apexp1\staticenv\apps\199\env_internal_bw.png',{'x_offset':45,'y_offset':13, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})
		
		#COLUMN NAMES
		item_fields = ['Tree Tag Range', 'Tag Color', 'Species', 'Species Substituted For',	'Nursery', 'Stock Type', 'Farm/Lot', 'Status']
		worksheet.write_row('A7', item_fields, item_header_format)

		#MAIN DATA
		for idx, val in enumerate(data["items"]):
			for idx2, val2 in enumerate(data["items"][idx]):
					worksheet.write(chr(idx2+65)+str(idx+8),data["items"][idx][val2], format_text)

		workbook.close()
		xlsx_data = output.getvalue()
		return xlsx_data