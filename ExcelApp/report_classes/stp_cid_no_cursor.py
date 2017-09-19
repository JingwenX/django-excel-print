# -*- coding: utf-8 -*
#rid 102 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
    main_data_url = str(stp_config.CONST.API_URL_PREFIX)+ 'bsmart_data/bsmart_data/stp_ws/stp_contract_item_detail_main/'
    main_data_url += str(params["item_num"])
    detail_item_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_contract_item_detail_detail_item/'
    detail_item_url += str(params["item_num"])
    comment_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_contract_item_detail_comment/'
    comment_url += str(params["item_num"])
    url_dict = {}
    url_dict["main_data"] = main_data_url
    url_dict["detail_item"] = detail_item_url
    url_dict["comment"] = comment_url
    return url_dict


#Contract Item Details
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheet = workbook.add_worksheet()

	data = res["main_data"]["items"][0] #changed by removing cursor
	#data = res[0]['main_data']
	#print("res type is "+ str(type(res)))
	#data = {"program":"Kapital Infrastructure","project_type":"Weston Road, Major Mackenzie Dr to Teston Rd","municipality":"Vaughan","regional_road":"Weston Road","between_road_1":"Major Mackenzie Drive West","between_road_2":"Teston Road","rins":"56-10","contract_item_num":"2012 -  044","year":2012,"ownership":"Regional ROW","status":"Active"}
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
	worksheet.insert_image('B1', stp_config.CONST.ENV_LOGO,{'x_offset':110,'y_offset':22, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

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
	data_item_details = res["detail_item"]["items"] #changed by removing cursor
	for rid, row in enumerate(data_item_details):
		#worksheet.write_row('A{}'.format(cr), [row['name'], row['qty']], format_text)
		#worksheet.write_row('A{}'.format(cr), row, format_text)
		worksheet.write('A' + str(cr), data_item_details[rid]['name'], format_text) 
		worksheet.write('B' + str(cr), data_item_details[rid]['qty'], format_num) 
		cr += 1; 


	
	cr += 3;
	worksheet.write_row('A{}'.format(cr), ['Comment', 'User'], item_header_format)
	cr += 1;
	data_comments = res["comment"]["items"] #changed by removing cursor
	for rid, row in enumerate(data_comments):
		#worksheet.write_row('A{}'.format(cr), [row['comments'], row['name']], format_text)
		worksheet.write('A' + str(cr), data_comments[rid]['comments'], format_text) 
		worksheet.write('B' + str(cr), data_comments[rid]['name'], format_text) 
		cr += 1;


	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data