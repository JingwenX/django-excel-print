# -*- coding: utf-8 -*
#rid 26 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	if params["snap"] == '0':
		base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_tree_planting_deficiency_current/'
		base_url += str(params["year"])
	else:
		base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_tree_planting_deficiency_snap/'
		base_url += str(params["snap"])
	return base_url

#Tree Planting Deficiency List
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]
	snap = params["snap"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheets = []

	title = 'Tree Planting Deficiency List' + (' Current' if snap == '0' else ' Snap - ' + str(snap))

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
	data = res
	item_fields = ['Tree ID', 'Tag Number', 'Item', 'Health', 'Deficiency', 'Required Repair']

	#MAIN DATA
	regions = {}

	for iid, val in enumerate(data["items"]):
		if str(data["items"][iid]["contractyear"]) == year:
			rKey = data["items"][iid].get("mun") + '-' + data["items"][iid].get("con") #+ '-' + data["items"][iid].get("rd")]
			if not rKey in regions:
				regions.update({rKey : {data["items"][iid].get("rd") : [[
					data["items"][iid].get("tid"),
					data["items"][iid].get("tno"),
					data["items"][iid].get("item"),
					data["items"][iid].get("health"),
					data["items"][iid].get("def"),
					data["items"][iid].get("rep"),
					data["items"][iid].get("tpuc")
					]]}})
			else:
				if data["items"][iid].get("rd") in regions[rKey]:
					regions[rKey][ data["items"][iid].get("rd")].append([
						data["items"][iid].get("tid"),
						data["items"][iid].get("tno"),
						data["items"][iid].get("item"),
						data["items"][iid].get("health"),
						data["items"][iid].get("def"),
						data["items"][iid].get("rep"),
						data["items"][iid].get("tpuc")
						])
				else:
					regions[rKey].update({data["items"][iid].get("rd") : [[
						data["items"][iid].get("tid"),
						data["items"][iid].get("tno"),
						data["items"][iid].get("item"),
						data["items"][iid].get("health"),
						data["items"][iid].get("def"),
						data["items"][iid].get("rep"),
						data["items"][iid].get("tpuc")
					]]})
				
	for reg_id, reg in enumerate(sorted(regions)):
		worksheets.append(workbook.add_worksheet(reg[:31]))

		stp_config.const.write_gen_title(title, workbook, worksheets[reg_id], rightmost_idx, year, con_num)
		worksheets[reg_id].insert_image('E1', stp_config.CONST.ENV_LOGO,{'x_offset':180,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

		worksheets[reg_id].set_column('A:F', 30)
		worksheets[reg_id].set_row(0,36)
		worksheets[reg_id].set_row(1,36)

		cr = 7

		worksheets[reg_id].merge_range('A{}:F{}'.format(cr,cr), reg, item_header_format)
		cr += 1

		for sid, side in enumerate(regions[reg]):
			worksheets[reg_id].write('A{}'.format(cr), 'RoadSide: ' + side, subtitle_format)
			worksheets[reg_id].write_row('A{}'.format(cr+1), item_fields, item_header_format)
			cr += 2

			start = cr
			for tree in regions[reg][side]:
				worksheets[reg_id].write_row('A{}'.format(cr), tree[0:6], format_text)
				cr += 1
			cr += 1

			worksheets[reg_id].merge_range('A{}:B{}'.format(cr,cr), 'Summary', item_header_format)
			cr += 1
			worksheets[reg_id].write('A{}'.format(cr), 'Number of Deficient Trees: ', subtotal_format)
			worksheets[reg_id].write('B{}'.format(cr), (cr - start - 2), subtotal_format)
			cr += 1
			worksheets[reg_id].write('A{}'.format(cr), 'Total Trees Planted on Contract: ', subtotal_format)
			worksheets[reg_id].write('B{}'.format(cr), regions[reg][side][0][6], subtotal_format)

			cr += 2
		
	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data