# -*- coding: utf-8 -*
#rid 19 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_warranty_ha/'
	base_url += str(params["wtype"])
	return base_url

#Warranty Report Health Analysis
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
	title = 'Warranty Report Health Analysis ' + type

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
	subtotal_format_text = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_TEXT)
	subtotal_format_money = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_MONEY)

	#HEADER
	#write general header and format
	rightmost_idx = 'E'
	stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

	#additional header image
	worksheet.insert_image('D1', stp_config.CONST.ENV_LOGO,{'x_offset':100,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

	data = res

	worksheet.set_column('A:E', 25)
	worksheet.set_row(0,36)
	worksheet.set_row(1,36)
	item_fields = ['Health', 'Total Trees Inspected', 'Number of Trees Accepted', 'Number of Trees Rejected', 'Number of Trees Missing']

	#MAIN DATA
	cr = 7
	muns = {}

	for idx, val in enumerate(data["items"]):
		if str(data["items"][idx]["contractyear"]) == year:
			if not data["items"][idx]["municipality"] in muns:
				muns.update({data["items"][idx]["municipality"] : {}})

	for mid, mun in enumerate(muns):
		for item_id, item in enumerate(data["items"]):
			if str(data["items"][item_id]["contractyear"]) == year:
				if data["items"][item_id]["municipality"] == mun and "warrantyaction" in data["items"][item_id]: 
					if data["items"][item_id]["health"] not in muns[mun]:
						muns[mun].update({data["items"][item_id]["health"] : [
								data["items"][item_id].get("healthname"),
								1,
								1 if data["items"][item_id].get("warrantyaction") == 'Accept' else 0,
								1 if data["items"][item_id].get("warrantyaction") == 'Reject' else 0,
								1 if data["items"][item_id].get("warrantyaction") == 'Missing Tree' else 0
							]})
					else:
						muns[mun][data["items"][item_id]["health"]][1] += 1
						muns[mun][data["items"][item_id]["health"]][2] += 1 if data["items"][item_id].get("warrantyaction") == 'Accept' else 0
						muns[mun][data["items"][item_id]["health"]][3] += 1 if data["items"][item_id].get("warrantyaction") == 'Reject' else 0
						muns[mun][data["items"][item_id]["health"]][4] += 1 if data["items"][item_id].get("warrantyaction") == 'Missing Tree' else 0

	for mid, mun in enumerate(muns):
		worksheet.merge_range('A{}:E{}'.format(cr,cr), mun, item_header_format)
		worksheet.write_row('A{}'.format(cr+1), item_fields, item_header_format)
		cr += 2
		start = cr
		for hid, hel in enumerate(sorted(muns[mun])):
			d = [str(hel) + ' - ' + str(muns[mun][hel][0])]
			d.extend(muns[mun][hel][1:])
			worksheet.write_row('A{}'.format(cr), d, format_text)
			cr += 1

		worksheet.write('A' + str(cr), 'Totals: ', subtotal_format_text)
		worksheet.write_formula('B' + str(cr), '=SUM(B' + str(start) + ':B' + str(cr-1) + ')', subtotal_format)
		worksheet.write_formula('C' + str(cr), '=SUM(C' + str(start) + ':C' + str(cr-1) + ')', subtotal_format)
		worksheet.write_formula('D' + str(cr), '=SUM(D' + str(start) + ':D' + str(cr-1) + ')', subtotal_format)
		worksheet.write_formula('E' + str(cr), '=SUM(E' + str(start) + ':E' + str(cr-1) + ')', subtotal_format)
		cr += 2
		
	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data