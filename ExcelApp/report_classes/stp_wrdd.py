# -*- coding: utf-8 -*
#rid 21 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_warranty_lists/'
	base_url += str(params["wtype"])
	return base_url

#Warranty Report Deficiency List Details
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
	title = 'Warranty Report Deficiency List Details ' + type

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
	rightmost_idx = 'G'
	stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

	#additional header image
	worksheet.insert_image('F1', stp_config.CONST.ENV_LOGO,{'x_offset':120,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

	data = res

	worksheet.set_column('A:G', 25)
	worksheet.set_row(0,36)
	worksheet.set_row(1,36)
	item_fields = ['Tree ID', 'Tag Colour', 'Tag Number', 'Item', 'Health Rating', 'Deficiency', 'Required Repair']

	#MAIN DATA
	cr = 7
	regions = {}

	for iid, val in enumerate(data["items"]):
		if str(data["items"][iid]["contractyear"]) == year:
			rKey = str(str(data["items"][iid].get("municipality")) + ' - ' + str(data["items"][iid].get("contract item")) + ' - ' + str(data["items"][iid].get("road")) + ' - ' + str(data["items"][iid].get("road side")))
			if not rKey in regions:
				regions.update({rKey : [[
					data["items"][iid].get("municipality"),
					data["items"][iid].get("contract item"),
					data["items"][iid].get("road"),
					data["items"][iid].get("between1"),
					data["items"][iid].get("between2"),
					data["items"][iid].get("road side"),
					data["items"][iid].get("tag number"),
					data["items"][iid].get("tag colour"),
					data["items"][iid].get("item"),
					data["items"][iid].get("tree id"),
					data["items"][iid].get("health"),
					data["items"][iid].get("deficiency"),
					data["items"][iid].get("required repair")
					]]})
			else:
				regions[rKey].append([
					data["items"][iid].get("municipality"),
					data["items"][iid].get("contract item"),
					data["items"][iid].get("road"),
					data["items"][iid].get("between1"),
					data["items"][iid].get("between2"),
					data["items"][iid].get("road side"),
					data["items"][iid].get("tag number"),
					data["items"][iid].get("tag colour"),
					data["items"][iid].get("item"),
					data["items"][iid].get("tree id"),
					data["items"][iid].get("health"),
					data["items"][iid].get("deficiency"),
					data["items"][iid].get("required repair")
					])
				
	breaks = []

	for reg_id, reg in enumerate(sorted(regions)):
		worksheet.merge_range('A{}:C{}'.format(cr, cr), "Municipality: " + str(regions[reg][0][0]), item_header_format)
		worksheet.merge_range('A{}:C{}'.format(cr+1, cr+1), "Contract Item No.: " + str(regions[reg][0][1]), item_header_format)
		(worksheet.merge_range('A{}:C{}'.format(cr+2, cr+2), "Regional Road: " + str(regions[reg][0][2]) + " Between " + 
			str(regions[reg][0][3]) + " and " + str(regions[reg][0][4]), item_header_format))
		cr += 2

		worksheet.write_row('A{}'.format(cr+1), item_fields, item_header_format)
		cr += 1

		for tree in regions[reg]:
			worksheet.write_row('A{}'.format(cr), tree[6:], format_text)
			cr += 1

		breaks.append(cr)
		cr += 1
	
	worksheet.set_h_pagebreaks(breaks)

	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data