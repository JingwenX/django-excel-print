# -*- coding: utf-8 -*
#rid 22 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_warranty_lists/'
	base_url += str(params["wtype"])
	return base_url

#Warranty Report Deficiency List Summary
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]
	wtype = params["wtype"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheet = workbook.add_worksheet()

	type = 'Year 1 Warranty' if wtype == '1' else 'Year 2 Warranty' if wtype == '2' else '12 Month Warranty'
	title = 'Warranty Report Deficiency List Summary ' + type

	#MAIN DATA FORMATING
	format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
	format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
	format_num2 = workbook.add_format(stp_config.CONST.FORMAT_NUM2)
	item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)
	##Hunter's additional formatting
	item_format = workbook.add_format(stp_config.CONST.ITEM_FORMAT)
	title_format = workbook.add_format(stp_config.CONST.TITLE_FORMAT)
	item_format_money = workbook.add_format(stp_config.CONST.ITEM_FORMAT_MONEY)
	subtitle_format = workbook.add_format(stp_config.CONST.SUBTITLE_FORMAT)
	subtotal_format = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT)
	subtotal_format_text = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_TEXT)
	subtotal_format_money = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_MONEY)

	#HEADER
	#write general header and format
	rightmost_idx = 'C'
	stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

	#additional header image
	worksheet.insert_image('C1', stp_config.CONST.ENV_LOGO,{'x_offset':90,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

	data = res

	worksheet.set_column('A:C', 50)
	worksheet.set_row(0,36)
	worksheet.set_row(1,36)
	item_fields = ['Deficiency', 'Deficiency Type', 'Total']

	#MAIN DATA
	cr = 7
	regions = {}

	for iid, val in enumerate(data["items"]):
		if str(data["items"][iid]["contractyear"]) == year:
			rKey = str(data["items"][iid].get("municipality")) + ' - ' + str(data["items"][iid].get("contract item")) + ' - ' + str(data["items"][iid].get("road side"))
			if not rKey in regions:
				regions.update({rKey : {data["items"][iid]["deficiency"] : [data["items"][iid].get("deficiency type"), 1]}})
			else:
				if data["items"][iid]["deficiency"] in regions[rKey]:
					regions[rKey][data["items"][iid]["deficiency"]][1] += 1
				else:
					regions[rKey].update({data["items"][iid]["deficiency"]: [data["items"][iid].get("deficiency type"), 1]})
				
	breaks = []
	for reg_id, reg in enumerate(sorted(regions)):
		worksheet.merge_range('A{}:C{}'.format(cr,cr), reg, subtitle_format)
		worksheet.write_row('A{}'.format(cr+1), item_fields, item_header_format)
		cr += 2
		start = cr
		for deficiency in regions[reg]:
			d = [deficiency]
			d.extend(regions[reg][deficiency])
			worksheet.write_row('A{}'.format(cr), [d[0], d[1]], format_text)
			worksheet.write('C{}'.format(cr), d[2], format_num)
			cr += 1
		worksheet.write('A{}'.format(cr), "Subtotal: ", subtotal_format_text)
		worksheet.write('B{}'.format(cr), " ", subtotal_format)
		worksheet.write_formula('C{}'.format(cr), "=SUM(C{}:C{})".format(start, cr-1), subtotal_format)
		breaks.append(cr)
		cr += 2
	
	worksheet.set_h_pagebreaks(breaks)
	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data