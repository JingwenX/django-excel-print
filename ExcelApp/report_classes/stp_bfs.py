# -*- coding: utf-8 -*
#rid 7 
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

class Report(object):

	def form_url(params):
		base_url = 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_costing_bid/'
		base_url += str(params["year"])
		return base_url


	#Bid Form Summary
	def render(res, params):

		rid = params["rid"]
		year = params["year"]
		con_num = params["con_num"]
		assign_num = params["assign_num"]

		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()
		title = 'Bid Form Summary'
		

		data = res

		worksheet.set_column('A:F', 30)
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)

		item_fields = ['Item Number', 'Item', 'Unit', 'Quantity', 'Unit Price', 'Total']

		cr = 7

		#MAIN DATA FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)
		##Hunter's additional formatting
		item_format = workbook.add_format(stp_config.CONST.ITEM_FORMAT)
		title_format = workbook.add_format(stp_config.CONST.TITLE_FORMAT)
		item_format_money = workbook.add_format(stp_config.CONST.ITEM_FORMAT_MONEY)
		subtitle_format = workbook.add_format(stp_config.CONST.SUBTITLE_FORMAT)
		subtotal_format = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT)
		subtotal_format_money = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_MONEY)

		#HEADER
		#write general header and format
		rightmost_idx = 'F'
		stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

		#additional header image
		worksheet.insert_image('E1', stp_config.CONST.ENV_LOGO,{'x_offset':180,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

		#MAIN DATA

		miDict = {'A' : 'A - Tree Planting - Ball and Burlap Trees',
		'B' : 'B - Tree Planting - Potted Perennials and Grass',
		'C' : 'C - Tree Planting - Potted Shrubs',
		'D' : 'D - Transplanting',
		'E' : 'E - Stumping',
		'F' : 'F - Watering',
		'G' : 'G - Tree Maintenance',
		'H' : 'H - Automated Vehicle Locating System'}

		
		items = {'A' : {}, 'B' : {}, 'C' : {}, 'D' : {}, 'E' : {}, 'F' : {}, 'G' : {}, 'H' : {}}

		
		for idx, val in enumerate(data["items"]):
			if not data["items"][idx]["item"] in items[data["items"][idx]["sec"]]:
				items[data["items"][idx]["sec"]].update({data["items"][idx]["item"] : [(data["items"][idx]["ino"] if "ino" in data["items"][idx] else '--'),
					data["items"][idx]["unit"] if "unit" in data["items"][idx] else 'N/A',
					int(data["items"][idx]["quantity"]) if "quantity" in data["items"][idx] else 0,
					data["items"][idx]["up"] if "up" in data["items"][idx] else 0]})
			else:
				items[data["items"][idx]["sec"]][data["items"][idx]["item"]][2] += int(data["items"][idx]["quantity"])

		for idx, val in enumerate(items):
			if items[val]:
				worksheet.merge_range('A' + str(cr) + ':F' + str(cr), miDict[val], subtitle_format)
				worksheet.write_row('A' + str(cr+1) + ':F' + str(cr+1), item_fields, subtitle_format)
				cr += 2
				start = cr
				for idx2, val2 in enumerate(items[val]):
					d = [items[val][val2][0], val2, items[val][val2][1], items[val][val2][2], str(items[val][val2][3]).lstrip(' ')]

					for i, v in enumerate(d):
							d[i] = '$0' if i >= 1 and (d[i] == 0 or d[i] =='0' or d[i] =='$.00') else str(d[i]).lstrip()

					worksheet.write_row('A' + str(cr), d, format_text)
					worksheet.write_formula('F' + str(cr), '=D' + str(cr) + '*E' + str(cr), item_format_money)
					cr += 1

				worksheet.write('A' + str(cr), 'SubTotal: ', subtotal_format)
				worksheet.write('B' + str(cr), '', subtotal_format)
				worksheet.write('C' + str(cr), '', subtotal_format)
				worksheet.write_formula('D' + str(cr), '=SUM(D' + str(start) + ':D' + str(cr-1) + ')', subtotal_format)
				worksheet.write('E' + str(cr), '', subtotal_format)
				worksheet.write_formula('F' + str(cr), '=SUM(F' + str(start) + ':F' + str(cr-1) + ')', subtotal_format_money)

				worksheet.merge_range('A' + str(cr+1) + ':G' + str(cr+1),' ')
				cr += 2


		workbook.close()

		xlsx_data = output.getvalue()
		return xlsx_data