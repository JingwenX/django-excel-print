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

	title = 'Tree Planting Payment Report'

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
	rightmost_idx = 'D'
	data = res
	item_fields = ['Item', 'Quantity', 'Unit Price', 'Total']

	#MAIN DATA
	cons = {}
	summary = {}

	for iid, val in enumerate(data["items"]):
		if str(data["items"][iid]["contractyear"]) == year:
			if not data["items"][iid]["contractitem"] in cons:
				cons.update({data["items"][iid]["contractitem"] : [[
					data["items"][iid].get("item"),
					data["items"][iid].get("qty"),
					data["items"][iid].get("up"),
					data["items"][iid].get("total")
					]]})
			else:
				cons[data["items"][iid]["contractitem"]].append([
					data["items"][iid].get("item"),
					data["items"][iid].get("qty"),
					data["items"][iid].get("up"),
					data["items"][iid].get("total")
					])

			#if not data["items"][iid]["item"] in summary:
			#	summary.update({data["items"][iid]["item"] : [
			#		1,
			#		data["items"][iid].get("up")
			#		]})
			#else:
			#	summary[data["items"][iid]["item"]][0] += 1
				

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

		for contract in cons[con]:
			if not contract[0] in summary:
				summary.update({contract[0] : [
					contract[1],
					contract[2]
					]})
			else:
				summary[contract[0]][0] += contract[1]

			worksheets[cid].write_row('A{}'.format(cr), [contract[0], contract[1]], format_text)
			worksheets[cid].write_row('C{}'.format(cr), [contract[2], contract[3]], item_format_money)
			cr += 1
		worksheets[cid].write('A{}'.format(cr), "Subtotal: ", subtotal_format)
		worksheets[cid].write_formula('B{}'.format(cr), "=SUM(B{}:B{})".format(start, cr - 1), subtotal_format)
		worksheets[cid].write('C{}'.format(cr), " ", subtotal_format_money)
		worksheets[cid].write_formula('D{}'.format(cr), "=SUM(D{}:D{})".format(start, cr - 1), subtotal_format_money)
		cr += 2

	worksheets.append(workbook.add_worksheet('Payment Summary'))

	stp_config.const.write_gen_title(title, workbook, worksheets[-1], rightmost_idx, year, con_num)
	worksheets[-1].insert_image('C1', stp_config.CONST.ENV_LOGO,{'x_offset':180,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

	worksheets[-1].set_column('A:D', 45)
	worksheets[-1].set_row(0,36)
	worksheets[-1].set_row(1,36)

	cr = 7
	worksheets[-1].merge_range('A{}:D{}'.format(cr,cr), 'Summary', item_header_format)
	worksheets[-1].write_row('A{}'.format(cr+1), item_fields, item_header_format)
	cr += 2

	print(summary)

	for sid, sitem in enumerate(sorted(summary)):
		d = [sitem]
		d.extend(summary[sitem])
		worksheets[-1].write_row('A{}'.format(cr), [d[0], d[1]], format_text)
		worksheets[-1].write('C{}'.format(cr), d[2], item_format_money)
		worksheets[-1].write_formula('D{}'.format(cr), '=B{}*C{}'.format(cr, cr), item_format_money)
		cr += 1

	worksheets[-1].write('A{}'.format(cr), 'Grand Total: ', subtotal_format)
	worksheets[-1].write_formula('B{}'.format(cr), '=SUM(B9:B{})'.format(cr-1), subtotal_format)
	worksheets[-1].write('C{}'.format(cr), ' ', subtotal_format)
	worksheets[-1].write_formula('D{}'.format(cr), '=SUM(D9:D{})'.format(cr-1), subtotal_format_money)
		
	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data