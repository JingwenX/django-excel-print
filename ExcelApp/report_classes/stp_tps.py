# -*- coding: utf-8 -*
#rid 9 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_tree_planting/'
	base_url += str(params["year"])
	base_url += '/' + str(params["item_num"])
	return base_url


#Tree Planting Summary
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheet = workbook.add_worksheet()
	title = 'Tree Planting Summary'
	
	data = res

	item_fields = ["Contract Item No.", "Tree Planting Detail No.", "Location", "Activity", "Quantity"]

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

	worksheet.set_column('A:E', 30)
	worksheet.set_row(0,36)
	worksheet.set_row(1,36)

	#HEADER
	#write general header and format
	rightmost_idx = 'E'
	
	#MAIN DATA
	stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

	#additional header image
	worksheet.insert_image('D1', stp_config.CONST.ENV_LOGO,{'x_offset':180,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

	muns = {}
	mtots = {}

	for idx, val in enumerate(data["items"]):
		if (data["items"][idx]["type_id"] in [1,2,3] and ((str(assign_num) == '-1' and not "assignment_num" in data["items"][idx]) or 
			(("assignment_num" in data["items"][idx]) and (str(data["items"][idx]["assignment_num"]) == str(assign_num))))):

			locat = (data["items"][idx].get("regional_road", "--") + ", " + 
				data["items"][idx].get("between_road_1", "--") + " to " + 
				data["items"][idx].get("between_road_2", "--"))

			act = ("Tree Planting" if data["items"][idx]["type_id"] == 1
					else "Stumping" if data["items"][idx]["type_id"] == 2
					else "Transplanting")

			if not data["items"][idx]["municipality"] in muns:
				muns.update({data["items"][idx]["municipality"] : [[
					data["items"][idx].get("contract_item_num", " "),
					data["items"][idx].get("detail_num", " "),
					locat,
					act,
					data["items"][idx].get("quantity", 0)
					]]})
			else:
				muns[data["items"][idx]["municipality"]].append([
					data["items"][idx].get("contract_item_num", " "),
					data["items"][idx].get("detail_num", " "),
					locat,
					act,
					data["items"][idx].get("quantity", 0)
					])

	cr = 7
	
	for mid, mun in enumerate(muns):
		worksheet.merge_range('A{}'.format(cr) + ':E{}'.format(cr), mun, subtitle_format)
		worksheet.write_row('A{}'.format(cr+1), item_fields, item_header_format)
		cr += 2
		for row in muns[mun]:
			mtots[mun] = mtots.get(mun, 0) + row[4]
			worksheet.write_row('A{}'.format(cr), row, format_text)
			cr += 1

		worksheet.merge_range('A{}'.format(cr) + ':D{}'.format(cr), "Total: ", subtotal_format)
		worksheet.write('E{}'.format(cr), mtots.get(mun, 0), subtotal_format)
		cr += 2



	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data