# -*- coding: utf-8 -*
#rid 101 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_contract_item_detail/'
	base_url += str(params["year"])
	return base_url


#Contract Item Details
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheet = workbook.add_worksheet()

	data = res['items'][0]
	title = 'Contact Item ' + data['contract_item_num']

	# MAIN DATA FORMATING
	format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
	format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
	item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)


	# TODO: SET ROW WIDTH

	worksheet.set_column('A:A', 50)
	worksheet.set_column('B:B', 50)
	worksheet.set_row(0,36)
	worksheet.set_row(1,36)
	worksheet.set_row(5,23.4)

	#HEADER
	#write general header and format
	rightmost_idx = 'B'
	stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, data['year'], con_num) #change 2017 to year

	# FILE SPECIFIC FORMATING
	format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
	item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)

	#additional header image
	worksheet.insert_image('B1', r'\\ykr-apexp1\staticenv\apps\199\env_internal_bw.png',{'x_offset':110,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

	# Master data
	worksheet.write_row('A7', ['Contract Number', con_num],format_text) 
	worksheet.write_row('A8', ['Status', data.get('status', '')],format_text) 
	worksheet.write_row('A9', ['Program', data.get('program', '') + ' ' + data.get('project_type','')],format_text) 
	worksheet.write_row('A10', ['Ownership', data.get('ownership', '')],format_text) 
	worksheet.write_row('A11', ['Municipality', data.get('municipality', '')],format_text) 
	worksheet.write_row('A12', ['Regional Road', data.get('regional_road', '')],format_text) 
	worksheet.write_row('A13', ['Between Road 1', data.get('between_road_1', '')],format_text) 
	worksheet.write_row('A14', ['Between Road 2', data.get('between_road_2', '')],format_text) 
	worksheet.write_row('A15', ['RINs', data.get('rins', '')],format_text) 

	item_fields = ['Contract Detail', 'Quantity']
	worksheet.write_row('A18', item_fields, item_header_format)


	#MAIN DATA
	cr = 19
	for row in data.get('item_details', {}):
		worksheet.write_row('A{}'.format(cr), [row['name'], row['qty']], format_text)
		cr += 1;


	
	cr += 3;
	worksheet.write_row('A{}'.format(cr), ['Comment', 'User'], item_header_format)
	cr += 1;
	for row in data.get('comment_details', {}):
		worksheet.write_row('A{}'.format(cr), [row['comments'], row['name']], format_text)
		cr += 1;


	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data