# -*- coding: utf-8 -*
#rid 20 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_warranty_ca/'
	base_url += str(params["wtype"])
	return base_url

#Warranty Report Contract Analysis
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
	title = 'Warranty Report Contract Analysis ' + type

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
	rightmost_idx = 'E'
	data = res
	item_fields = ['Health', 'Total Trees Inspected', 'Number of Trees Accepted', 'Number of Trees Rejected', 'Number of Trees Missing']

	#MAIN DATA
	cons = {}

	for idx, val in enumerate(data["items"]):
		if str(data["items"][idx]["contractyear"]) == year:
			if not data["items"][idx]["contract_item"] in cons:
				cons.update({data["items"][idx]["contract_item"] : {}})



	for cid, con in enumerate(cons):
		for item_id, item in enumerate(data["items"]):
			if str(data["items"][item_id]["contractyear"]) == year:
				if data["items"][item_id]["contract_item"] == con and "warrantyaction" in data["items"][item_id]: 
					if data["items"][item_id]["health"] not in cons[con]:
						cons[con].update({data["items"][item_id]["health"] : [
								data["items"][item_id].get("healthname"),
								1,
								1 if data["items"][item_id].get("warrantyaction") == 'Accept' else 0,
								1 if data["items"][item_id].get("warrantyaction") == 'Reject' else 0,
								1 if data["items"][item_id].get("warrantyaction") == 'Missing Tree' else 0
							]})
					else:
						cons[con][data["items"][item_id]["health"]][1] += 1
						cons[con][data["items"][item_id]["health"]][2] += 1 if data["items"][item_id].get("warrantyaction") == 'Accept' else 0
						cons[con][data["items"][item_id]["health"]][3] += 1 if data["items"][item_id].get("warrantyaction") == 'Reject' else 0
						cons[con][data["items"][item_id]["health"]][4] += 1 if data["items"][item_id].get("warrantyaction") == 'Missing Tree' else 0

	for cid, con in enumerate(sorted(cons)):
		worksheets.append(workbook.add_worksheet(con))

		worksheets[cid].set_column('A:E', 25)
		worksheets[cid].set_row(0,36)
		worksheets[cid].set_row(1,36)

		stp_config.const.write_gen_title(title, workbook, worksheets[cid], rightmost_idx, year, con_num)

		#additional header image
		worksheets[cid].insert_image('D1', stp_config.CONST.ENV_LOGO,{'x_offset':100,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})
		cr = 7

		worksheets[cid].merge_range('A{}:E{}'.format(cr,cr), con, subtitle_format)
		worksheets[cid].write_row('A{}'.format(cr+1), item_fields, item_header_format)
		cr += 2

		for hid, hel in enumerate(sorted(cons[con])):
			d = [str(hel) + ' - ' + str(cons[con][hel][0])]
			d.extend(cons[con][hel][1:])
			worksheets[cid].write('A{}'.format(cr), d[0], format_text)
			worksheets[cid].write_row('B{}'.format(cr), d[1:], format_num)
			cr += 1

		worksheets[cid].write('A' + str(cr), 'Totals: ', subtotal_format_text)
		worksheets[cid].write_formula('B' + str(cr), '=SUM(B9:B' + str(cr-1) + ')', subtotal_format)
		worksheets[cid].write_formula('C' + str(cr), '=SUM(C9:C' + str(cr-1) + ')', subtotal_format)
		worksheets[cid].write_formula('D' + str(cr), '=SUM(D9:D' + str(cr-1) + ')', subtotal_format)
		worksheets[cid].write_formula('E' + str(cr), '=SUM(E9:E' + str(cr-1) + ')', subtotal_format)
		cr += 2
		
	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data