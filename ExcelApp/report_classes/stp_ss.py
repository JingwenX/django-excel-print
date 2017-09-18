# -*- coding: utf-8 -*
#rid 3 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_contract_detail/'
	base_url += str(params["year"])
	return base_url


#Species Summary
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheet = workbook.add_worksheet()
	data = res
	title = 'Contract Preparation - Species Summary'
	
	
	#MAIN DATA FORMATING
	format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
	format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
	item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)	

	title_format = workbook.add_format(stp_config.CONST.TITLE_FORMAT)
	item_format_money = workbook.add_format(stp_config.CONST.ITEM_FORMAT_MONEY)
	subtitle_format = workbook.add_format(stp_config.CONST.SUBTITLE_FORMAT)
	subtitle_format2 = workbook.add_format(stp_config.CONST.SUBTITLE_FORMAT2)
	subtotal_format = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT)
	subtotal_format_text = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_TEXT)
	subtotal_format_money = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_MONEY)

	#Set column and rows
	worksheet.set_column('A:B', 60)
	worksheet.set_row(0,36)
	worksheet.set_row(1,36)

	#HEADER
	#write general header and format
	rightmost_idx = 'B'
	stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

	#additional header image
	worksheet.insert_image('B1', stp_config.CONST.ENV_LOGO,{'x_offset':180,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

	#MAIN DATA
	item_fields = ['Species', 'Quantity']

	species = {}
	cr = 7
	#print(data["items"])
	for idx, val in enumerate(data["items"]):
		if "species" in data["items"][idx]:
			if not data["items"][idx]["species"] in species:
				species[data["items"][idx]["species"]] = data["items"][idx]["quantity"]
			else:
				species[data["items"][idx]["species"]] += data["items"][idx]["quantity"]

	worksheet.write_row('A' + str(cr), item_fields, item_header_format)
	cr += 1

	start = cr
	for sid, spec in enumerate(species):
		#worksheet.write('A' + str(cr), [spec, species[spec]], format_text)
		worksheet.write('A' + str(cr), spec, format_text)
		worksheet.write('B' + str(cr), species[spec], format_num)

		cr += 1

	worksheet.write('A{}'.format(cr), 'Total: ', subtotal_format_text)
	worksheet.write_formula('B{}'.format(cr), '=sum(B{}:B{})'.format(start, cr - 1), subtotal_format)

	workbook.close()
	
	xlsx_data = output.getvalue()
	return xlsx_data