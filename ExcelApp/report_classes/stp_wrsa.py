# -*- coding: utf-8 -*
#rid 17 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_warranty_sa/'
	base_url += str(params["wtype"])
	return base_url

#Warranty Report Species Analysis
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
	title = 'Warranty Report Species Analysis ' + type

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
	rightmost_idx = 'E'
	stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

	#additional header image
	worksheet.insert_image('B1', stp_config.CONST.ENV_LOGO,{'x_offset':180,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

	data = res

	worksheet.set_column('A:E', 25)
	worksheet.set_row(0,36)
	worksheet.set_row(1,36)
	item_fields = ['Species', 'Total Trees Inspected', 'Number of Trees Accepted', 'Number of Trees Rejected', 'Number of Trees Missing']
	worksheet.write_row('A7', item_fields, item_header_format)

	#MAIN DATA
	cr = 8
	species = []
	totals = {}

	for sid, spec in enumerate(data["items"]):
		if str(data["items"][sid]["contractyear"]) == year and not data["items"][sid]["species"] in species:
			species.append(data["items"][sid]["species"])

	species.sort()

	for sid, spec in enumerate(species):
		for idx, val in enumerate(data["items"]):
			if data["items"][idx]["species"] == spec and "warrantyaction" in data["items"][idx] and str(data["items"][idx]["contractyear"]) == year:
				temp = [1,1 if data["items"][idx]["warrantyaction"] == 'Accept' else 0,
				1 if data["items"][idx]["warrantyaction"] == 'Reject' else 0,
				1 if data["items"][idx]["warrantyaction"] == 'Missing Tree' else 0]

				if spec in totals:
					totals[spec] = [totals[spec][0] + temp[0], totals[spec][1] + temp[1], totals[spec][2] + temp[2], totals[spec][3] + temp[3]]

				else:
					totals[spec] = [temp[0], temp[1], temp[2], temp[3]]


		worksheet.write('A' + str(cr), species[sid], format_text)
		worksheet.write_row('B' + str(cr), totals[spec], format_text)
		cr += 1

	#FORMULAE AND FOOTERS
	worksheet.write('A' + str(cr), 'Totals: ', format_text)
	worksheet.write_formula('B' + str(cr), '=SUM(B8:B' + str(cr-1) + ')', format_text)
	worksheet.write_formula('C' + str(cr), '=SUM(C8:C' + str(cr-1) + ')', format_text)
	worksheet.write_formula('D' + str(cr), '=SUM(D8:D' + str(cr-1) + ')', format_text)
	worksheet.write_formula('E' + str(cr), '=SUM(E8:E' + str(cr-1) + ')', format_text)

	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data