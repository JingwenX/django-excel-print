# -*- coding: utf-8 -*
#rid 24 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_warranty_replacement/'
	base_url += str(params["wtype"])
	return base_url

#Warranty Report Replacement List
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]
	wtype = params["wtype"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheets = [] 

	type = 'Year 1 Warranty' if wtype == '1' else 'Year 2 Warranty' if wtype == '2' else '12 Month Warranty'
	title = 'Warranty Report Replacement List ' + type

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
	subtotal_format_money = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_MONEY)

	#HEADER
	#write general header and format
	rightmost_idx = 'D'
	data = res
	item_fields = ['Tag Number', 'Species', 'Health Rating', 'Comments']

	#MAIN DATA
	cr = 7
	regions = {}

	for iid, val in enumerate(data["items"]):
		if str(data["items"][iid]["year"]) == year:
			rKey = str(data["items"][iid].get("municipality")) + '-' + str(data["items"][iid].get("contract item")) + '-' + str(data["items"][iid].get("road side"))
			if len(rKey) > 30 and data["items"][iid].get("municipality") == 'Whitchurch-Stouffville':
				rKey = 'WS-' + str(data["items"][iid].get("contract item")) + '-' + str(data["items"][iid].get("road side"))
			if not rKey in regions:
				regions.update({rKey : [[
					data["items"][iid].get("tno"),
					data["items"][iid].get("spec"),
					data["items"][iid].get("hel"),
					' '
					]]})
			else:
				regions[rKey].append([
					data["items"][iid].get("tno"),
					data["items"][iid].get("spec"),
					data["items"][iid].get("hel"),
					' '
					])
				

	for reg_id, reg in enumerate(sorted(regions)):
		worksheets.append(workbook.add_worksheet(reg[:31]))

		worksheets[reg_id].insert_image('C1', stp_config.CONST.ENV_LOGO,{'x_offset':180,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

		worksheets[reg_id].set_column('A:D', 35)
		worksheets[reg_id].set_row(0,36)
		worksheets[reg_id].set_row(1,36)

		stp_config.const.write_gen_title(title, workbook, worksheets[reg_id], rightmost_idx, year, con_num)

		cr = 7

		worksheets[reg_id].merge_range('A{}:D{}'.format(cr,cr), reg, item_header_format)
		worksheets[reg_id].write_row('A{}'.format(cr+1), item_fields, item_header_format)
		cr += 2

		for tree in regions[reg]:
			worksheets[reg_id].write('A{}'.format(cr), tree[0], format_num2)
			worksheets[reg_id].write_row('B{}'.format(cr), tree[1:], format_text)
			cr += 1
		cr += 1
		
	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data