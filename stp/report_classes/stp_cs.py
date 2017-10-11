# -*- coding: utf-8 -*
#rid 3 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'stp_costing_bid/'
	base_url += str(params["year"])
	return base_url


#Costing Summary
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheet = workbook.add_worksheet()
	title = 'Costing Summary by Program'
	

	data = res

	#MAIN DATA FORMATING
	format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
	format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
	item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)	

	title_format = workbook.add_format(stp_config.CONST.TITLE_FORMAT)
	item_format_money = workbook.add_format(stp_config.CONST.ITEM_FORMAT_MONEY)
	subtitle_format = workbook.add_format(stp_config.CONST.SUBTITLE_FORMAT)
	subtitle_format2 = workbook.add_format(stp_config.CONST.SUBTITLE_FORMAT2)
	subtotal_format = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT)
	subtotal_format_text = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_TEXT)
	subtotal_format_money = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_MONEY)

	#SET COLUMN AND ROW
	worksheet.set_column('A:G', 30)
	worksheet.set_row(0,36)
	worksheet.set_row(1,36)
	
	#HEADER
	#write general header and format
	rightmost_idx = 'G'
	stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

	#additional header image
	worksheet.insert_image('F1', stp_config.CONST.ENV_LOGO,{'x_offset':180,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

	#MAIN DATA
	item_fields = ['Item', 'Quantity', 'Last Year Price', 'This Year Estimate', 'This Year Actual', 'Estimated Total', 'Total']

	programs = {'Capital Infrastructure' : [],
	'Infill / Retrofit' : [],
	'EAB Replacement' : []} 

	#separates items by programs
	for idx, val in enumerate(data["items"]):
		if data["items"][idx]["program"] == "Capital Infrastructure":
			programs['Capital Infrastructure'].append(data["items"][idx])
		elif data["items"][idx]["program"] == "Infill / Retrofit":
			programs['Infill / Retrofit'].append(data["items"][idx])
		else:
			programs['EAB Replacement'].append(data["items"][idx])

	cr = 7

	for pid, program in enumerate(programs):
		if programs[program]:
			worksheet.merge_range('A' + str(cr) + ':G' + str(cr), program, subtitle_format)
			worksheet.write_row('A' + str(cr+1), item_fields, item_header_format)
			cr += 2

		miDict = {'A' : 'Tree Planting - Ball and Burlap Trees',
		'B' : 'Tree Planting - Potted Perennials and Grass',
		'C' : 'Tree Planting - Potted Shrubs',
		'D' : 'Transplanting',
		'E' : 'Stumping',
		'F' : 'Watering',
		'G' : 'Tree Maintenance',
		'H' : 'Automated Vehicle Locating System'}

		items = {'A' : {}, 'B' : {}, 'C' : {}, 'D' : {}, 'E' : {}, 'F' : {}, 'G' : {}, 'H' : {}}

		for idx, val in enumerate(programs[program]):
			if "item" in programs[program][idx]:
				if not programs[program][idx]["item"] in items[programs[program][idx]["sec"]]:
					items[programs[program][idx]["sec"]].update({programs[program][idx]["item"] : [int(programs[program][idx]["quantity"]),
						programs[program][idx]["lyp"] if "lyp" in programs[program][idx] else 0,
						programs[program][idx]["pe"] if "pe" in programs[program][idx] else 0,
						programs[program][idx]["up"] if "up" in programs[program][idx] else 0]})
				else:
					items[programs[program][idx]["sec"]][programs[program][idx]["item"]][0] += int(programs[program][idx]["quantity"])

		for idx, val in enumerate(items):
			if items[val]:
				worksheet.merge_range('A' + str(cr) + ':G' + str(cr), miDict[val], subtitle_format2)
				cr += 1
				start = cr
				for idx2, val2 in enumerate(items[val]):
					d = [val2]
					d.extend(items[val][val2])

					#changes all zeros to $0 for currency items
					for i, v in enumerate(d):
						d[i] = '$0' if i >= 1 and (d[i] == 0 or d[i] =='0' or d[i] =='$.00') else str(d[i]).lstrip(' ')

					worksheet.write('A{}'.format(cr), d[0], format_text)
					worksheet.write('B{}'.format(cr), float(d[1]), format_num)
					worksheet.write('C{}'.format(cr), d[2], item_format_money)
					worksheet.write('D{}'.format(cr), d[3], item_format_money)
					worksheet.write('E{}'.format(cr), d[4], item_format_money)
					worksheet.write_formula('F' + str(cr), '=B' + str(cr) + '*D' + str(cr), item_format_money)
					worksheet.write_formula('G' + str(cr), '=B' + str(cr) + '*E' + str(cr), item_format_money)
					cr += 1

				worksheet.write('A' + str(cr), 'SubTotal: ', subtotal_format_text)
				worksheet.write_formula('B' + str(cr), '=SUM(B' + str(start) + ':B' + str(cr-1) + ')', subtotal_format)
				worksheet.write('C' + str(cr), '', subtotal_format)
				worksheet.write('D' + str(cr), '', subtotal_format)
				worksheet.write('E' + str(cr), '', subtotal_format)
				worksheet.write_formula('F' + str(cr), '=SUM(F' + str(start) + ':F' + str(cr-1) + ')', subtotal_format_money)
				worksheet.write_formula('G' + str(cr), '=SUM(G' + str(start) + ':G' + str(cr-1) + ')', subtotal_format_money)

				worksheet.merge_range('A' + str(cr+1) + ':G' + str(cr+1),' ')
				cr += 2

	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data