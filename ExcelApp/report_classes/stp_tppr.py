# -*- coding: utf-8 -*
#rid 25 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_tree_planting_payment/'
	base_url += str(params["payno"])
	return base_url

#Tree Planting Payment Report
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]
	payno = params["payno"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheets = []

	title = 'Tree Planting Payment Report - ' + ('current' if payno == '0' else payno)

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
	rightmost_idx = 'D'
	data = res
	item_fields = ['Item', 'Quantity', 'Unit Price', 'Total']

	#MAIN DATA
	cons = {}
	summary = {}

	for iid, val in enumerate(data["items"]):
		if str(data["items"][iid]["contractyear"]) == year:
			if not data["items"][iid]["contractitem"] in cons:
				cons.update({data["items"][iid]["contractitem"] : {data["items"][iid].get("item"): [
					data["items"][iid].get("code") ,
					data["items"][iid].get("qty"),
					data["items"][iid].get("up"),
					data["items"][iid].get("total")
					]}})
			else:
				if not data["items"][iid].get("item") in cons[data["items"][iid]["contractitem"]]:
					cons[data["items"][iid]["contractitem"]].update({data["items"][iid].get("item") : [
						data["items"][iid].get("code"),
						data["items"][iid].get("qty"),
						data["items"][iid].get("up"),
						data["items"][iid].get("total")
						]})
				else:
					cons[data["items"][iid]["contractitem"]][data["items"][iid].get("item")][1] += data["items"][iid].get("qty")

				
	for cid, con in enumerate(sorted(cons)):
		worksheets.append(workbook.add_worksheet(con[:31]))

		stp_config.const.write_gen_title(title, workbook, worksheets[cid], rightmost_idx, year, con_num)
		worksheets[cid].insert_image('C1', stp_config.CONST.ENV_LOGO,{'x_offset':300,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

		worksheets[cid].set_column('A:D', 45)
		worksheets[cid].set_row(0,36)
		worksheets[cid].set_row(1,36)

		cr = 7

		worksheets[cid].merge_range('A{}:D{}'.format(cr,cr), con, item_header_format)
		worksheets[cid].write_row('A{}'.format(cr+1), item_fields, item_header_format)
		cr += 2
		start = cr

		for item in cons[con]:
			if not item in summary:
				summary.update({item : [
					cons[con][item][0],
					cons[con][item][1],
					cons[con][item][2]
					]})
			else:
				summary[item][1] += cons[con][item][1]

			worksheets[cid].write('A{}'.format(cr), item, format_text)
			worksheets[cid].write('B{}'.format(cr), cons[con][item][1], format_num)
			worksheets[cid].write_row('C{}'.format(cr), [cons[con][item][2], cons[con][item][3]], item_format_money)
			cr += 1
		worksheets[cid].write('A{}'.format(cr), "Subtotal: ", subtotal_format_text)
		worksheets[cid].write_formula('B{}'.format(cr), "=SUM(B{}:B{})".format(start, cr - 1), subtotal_format)
		worksheets[cid].write('C{}'.format(cr), " ", subtotal_format_money)
		worksheets[cid].write_formula('D{}'.format(cr), "=SUM(D{}:D{})".format(start, cr - 1), subtotal_format_money)
		cr += 2

	worksheets.append(workbook.add_worksheet('Payment Summary'))

	stp_config.const.write_gen_title(title, workbook, worksheets[-1], rightmost_idx, year, con_num)
	worksheets[-1].insert_image('C1', stp_config.CONST.ENV_LOGO,{'x_offset':180,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

	worksheets[-1].set_column('A:E', 45)
	worksheets[-1].set_row(0,36)
	worksheets[-1].set_row(1,36)

	cr = 7
	worksheets[-1].merge_range('A{}:E{}'.format(cr,cr), 'Summary', item_header_format)
	worksheets[-1].write_row('A{}'.format(cr+1), ['Item Code'] + item_fields, item_header_format)
	cr += 2

	for sid, sitem in enumerate(sorted(summary)):
		d = [sitem]
		d.extend(summary[sitem])
		worksheets[-1].write_row('A{}'.format(cr), [d[1], d[0]], format_text)
		worksheets[-1].write('C{}'.format(cr), d[2], format_num)
		worksheets[-1].write('D{}'.format(cr), d[3], item_format_money)
		worksheets[-1].write_formula('E{}'.format(cr), '=C{}*D{}'.format(cr, cr), item_format_money)
		cr += 1

	worksheets[-1].write('A{}'.format(cr), 'Grand Total: ', subtotal_format_text)
	worksheets[-1].write('B{}'.format(cr), ' ', subtotal_format)
	worksheets[-1].write_formula('C{}'.format(cr), '=SUM(C9:C{})'.format(cr-1), subtotal_format)
	worksheets[-1].write('D{}'.format(cr), ' ', subtotal_format)
	worksheets[-1].write_formula('E{}'.format(cr), '=SUM(E9:E{})'.format(cr-1), subtotal_format_money)
		
	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data