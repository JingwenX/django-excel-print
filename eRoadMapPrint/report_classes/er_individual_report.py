# -*- coding: utf-8 -*
#rid 201 TPE
# test http://127.0.0.1:8000/stp/?rid=201&year=2017&con_num=T01&item_num=1282
# test http://127.0.0.1:8000/stp/?rid=201&year=2017&con_num=T01&item_num=1201
# test http://127.0.0.1:8000/stp/?rid=201&year=2017&con_num=T01&item_num=565


import xlsxwriter
from io import BytesIO
import datetime
from .. import tpe_config

#pdf
import os
import string
import pythoncom
import tempfile
import win32com.client



#====util======


def nextCell(col_letter):
	col_idx = ord(col_letter)-64
	next_col_letter = chr(col_idx+65)
	return next_col_letter

def form_url(params):
	#assign_num = project id pid
	#item_num = report id rid
    description_url = r'http://ykr-dev-apex.devyork.ca/apexenv/'+ 'bsmart_data/tpe/description/'
    description_url += str(params["assign_num"])

    edocs_url = r'http://ykr-dev-apex.devyork.ca/apexenv/' + 'bsmart_data/tpe/edocs/'
    edocs_url += str(params["assign_num"])

    milestones_url = r'http://ykr-dev-apex.devyork.ca/apexenv/' + 'bsmart_data/tpe/milestones/'
    milestones_url += str(params["item_num"])

    report_details_url = r'http://ykr-dev-apex.devyork.ca/apexenv/' + 'bsmart_data/tpe/report_details/'
    report_details_url += str(params["item_num"])

    report_facts_url = r'http://ykr-dev-apex.devyork.ca/apexenv/' + 'bsmart_data/tpe/report_facts/'
    report_facts_url += str(params["assign_num"])

    status_graph_url = r'http://ykr-dev-apex.devyork.ca/apexenv/' + 'bsmart_data/tpe/status_graph/'
    status_graph_url += str(params["item_num"])

    report_info_url = r'http://ykr-dev-apex.devyork.ca/apexenv/' + 'bsmart_data/tpe/report_info/'
    report_info_url += str(params["item_num"])

    url_dict = {}
    url_dict["description"] = description_url
    url_dict["edocs"] = edocs_url
    url_dict["milestones"] = milestones_url
    url_dict["report_details"] = report_details_url
    url_dict["report_facts"] = report_facts_url
    url_dict["status_graph"] = status_graph_url
    url_dict['report_info'] = report_info_url
    return url_dict



#Render Report Card Details
def render(res, params):
	pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheet = workbook.add_worksheet()




	# MAIN DATA FORMATING
	PROJECT_HEADER_FORMAT = {'bold':True,
		'font_name':'Calibri',
		'font_size':20,
		#'border':2, #2 is the value for thick border
		'align':'left',
		'valign':'vcenter',}
	ENV_HEADER_FORMAT = {'bold':True,
		'font_name':'Calibri',
		'font_size':20,
		#'border':2, #2 is the value for thick border
		'align':'center',
		'valign':'top',}
	FORMAT_TEXT_BOLD = {'font_name':'Calibri',
		'font_size':12,
		'align': 'center',
		'valign': 'vcenter',
		'text_wrap': True,
		'bold': True}
	BLUE_TITLE_FORMAT = {'bold':True,
		'font_name':'Calibri',
		'font_size':14,
		#'border':2, #2 is the value for thick border
		'bg_color':'#4169E1',
		'font_color':'white',
		'align':'center',
		'valign':'vcenter',}
	format_text = workbook.add_format(tpe_config.CONST.FORMAT_TEXT)
	format_num = workbook.add_format(tpe_config.CONST.FORMAT_NUM)
	project_header_format = workbook.add_format(PROJECT_HEADER_FORMAT)
	item_header_format = workbook.add_format(tpe_config.CONST.ITEM_HEADER_FORMAT)
	main_header1_format = workbook.add_format(tpe_config.CONST.MAIN_HEADER1_FORMAT)
	main_header2_format = workbook.add_format(tpe_config.CONST.MAIN_HEADER2_FORMAT)
	title_format = workbook.add_format(tpe_config.CONST.TITLE_FORMAT)
	item_header_format = workbook.add_format(tpe_config.CONST.ITEM_HEADER_FORMAT)
	env_header_format = workbook.add_format(ENV_HEADER_FORMAT)

	#new formats
	format_text_bold = workbook.add_format(FORMAT_TEXT_BOLD)

	#title
	format_text_center = workbook.add_format(tpe_config.CONST.FORMAT_TEXT)
	format_text_center.set_align('center')
	format_text_center.set_border(False)


	format_text_bold_14 = workbook.add_format(FORMAT_TEXT_BOLD)
	format_text_bold_14.set_font_size(14)
	format_text_bold_14.set_bold(True)

	format_text_bold_all_border = workbook.add_format(FORMAT_TEXT_BOLD)
	format_text_bold_all_border.set_border(2)

	format_text_l_no_border_no_wrap = workbook.add_format(tpe_config.CONST.FORMAT_TEXT)
	format_text_l_no_border_no_wrap.set_text_wrap(False)
	format_text_l_no_border_no_wrap.set_border(False)
	format_text_l_no_border_no_wrap.set_align('left')

	#colorful
	format_text_bold_all_border_white = workbook.add_format(FORMAT_TEXT_BOLD)
	format_text_bold_all_border_white.set_border(2)

	format_text_bold_all_border_red = workbook.add_format(FORMAT_TEXT_BOLD)
	format_text_bold_all_border_red.set_border(2)
	format_text_bold_all_border_red.set_bg_color('red')

	format_text_bold_all_border_green = workbook.add_format(FORMAT_TEXT_BOLD)
	format_text_bold_all_border_green.set_border(2)
	format_text_bold_all_border_green.set_bg_color('#33cc33')

	format_text_bold_all_border_yellow = workbook.add_format(FORMAT_TEXT_BOLD)
	format_text_bold_all_border_yellow.set_border(2)
	format_text_bold_all_border_yellow.set_bg_color('yellow')

	format_text_bold_all_border_colorful = [format_text_bold_all_border_white, format_text_bold_all_border_red, format_text_bold_all_border_green, format_text_bold_all_border_yellow]

	blue_title_format = workbook.add_format(BLUE_TITLE_FORMAT)
	blue_title_format_bottom_border = blue_title_format.set_bottom(2)

	blue_content_format = workbook.add_format(tpe_config.CONST.FORMAT_TEXT)
	blue_content_format.set_align('left')
	blue_content_format.set_valign('top')
	blue_content_format.set_border(False)

	#table 1: project facts

	format_text_bold_left_border = workbook.add_format(FORMAT_TEXT_BOLD)
	format_text_bold_left_border.set_left(2)
	format_text_bold_left_border.set_align('left')

	format_text_normal_right_border = workbook.add_format(tpe_config.CONST.FORMAT_TEXT)
	format_text_normal_right_border.set_align('left')
	format_text_normal_right_border.set_border(False)
	format_text_normal_right_border.set_border_color('black')
	format_text_normal_right_border.set_right(2)

	format_text_top_border = workbook.add_format(tpe_config.CONST.FORMAT_TEXT)
	format_text_top_border.set_border(False)
	format_text_top_border.set_top(2)
	format_text_top_border.set_border_color('black')

	#table 3:


	format_percentage_bold = workbook.add_format(tpe_config.CONST.FORMAT_TEXT)
	format_percentage_bold.set_border_color('black')
	format_percentage_bold.set_border(2)
	format_percentage_bold.set_align('right')
	format_percentage_bold.set_bold(True)


	format_text_bold_all_border_left = workbook.add_format(tpe_config.CONST.FORMAT_TEXT)
	format_text_bold_all_border_left.set_border(2)
	format_text_bold_all_border_left.set_bold(True)

	format_percentage = workbook.add_format(tpe_config.CONST.FORMAT_TEXT)
	format_percentage.set_border(2)
	format_percentage.set_align('center')
	format_percentage.set_border_color('black')



	format_text_normal_all_border = workbook.add_format(tpe_config.CONST.FORMAT_TEXT)
	format_text_normal_all_border.set_border(2)
	format_text_normal_all_border.set_border_color('black')

	# table 3: total completeness
	format_percentage_grey = workbook.add_format(tpe_config.CONST.FORMAT_TEXT)
	format_percentage_grey.set_border(2)
	format_percentage_grey.set_align('center')
	format_percentage_grey.set_border_color('black')
	format_percentage_grey.set_bg_color('#D3D3D3')
	format_percentage_grey.set_bold(True)
	format_percentage_grey.set_font_size(14)

	format_text_normal_all_border_grey = workbook.add_format(tpe_config.CONST.FORMAT_TEXT)
	format_text_normal_all_border_grey.set_border(2)
	format_text_normal_all_border_grey.set_border_color('black')
	format_text_normal_all_border_grey.set_bg_color('#D3D3D3')
	format_text_normal_all_border_grey.set_bold(True)
	format_text_normal_all_border_grey.set_font_size(14)



	def getColorFormat(color_name):
		i=0
		if color_name == 'G':
			i = 2

		elif color_name == 'Y':
			i = 3

		elif color_name == 'R':
			i = 1

		else:
			i = 0

		return format_text_bold_all_border_colorful[i]

	def merge_bottom_cr(cr, text):
		l = len(text)
		num_row_needed = round(l/90 + 0.5)
		return cr + num_row_needed


	# TODO: SET ROW WIDTH
	worksheet.set_row(0, 40)
	worksheet.set_row(1, 16)
	worksheet.set_row(2, 45.6)
	for i in range(3, 32):
		worksheet.set_row(i, 21)


	worksheet.set_column('A:A', 21.22)
	worksheet.set_column('B:B', 16.44)
	worksheet.set_column('C:C', 12)
	worksheet.set_column('D:D', 3)
	worksheet.set_column('E:E', 21.33)
	worksheet.set_column('H:H', 10)

	#HEADER
	#write general header and format
	data = res["report_facts"]["items"][0] #changed by removing cursor
	title =  'Project: ' + str(data['project name'])
	
	rightmost_idx = 'L'
	worksheet.insert_image('A1',r'\\ykr-apexp1\staticenv\York_Logo.png',{'x_offset':0,'y_offset':0,'x_scale':0.25,'y_scale':0.25})
	worksheet.insert_image('J1', r'\\ykr-apexp1\staticenv\eroadmap_legend.PNG', {'x_offset':20,'y_offset':10, 'x_scale':0.45,'y_scale':0.45})
	#WRITE TITLE
	worksheet.merge_range('A1:'+ rightmost_idx + '1','ENV - Technology Project Status Card', env_header_format)
	title_data = res["report_info"]["items"][0]
	worksheet.merge_range('A2:' + rightmost_idx + '2',title_data['report_name'], format_text_center)
	worksheet.merge_range('A3:'+ rightmost_idx + '3',title, project_header_format)

	from_date =  title_data['from_date'] if 'from_date' in title_data.keys() else 'TBD'
	to_date =  title_data['to_date'] if 'to_date' in title_data.keys() else 'TBD'
	worksheet.write('A4', 'Project Start Date: ' + from_date, format_text_l_no_border_no_wrap)
	worksheet.write('A5', 'Project End Date: ' + to_date, format_text_l_no_border_no_wrap)


	#WRITE FOOTER

	
	#FOOTER
	##LEFT FOOTER

	data2 = res["edocs"]["items"][0]
	edocs_folder = data2['project_folder'] if 'project_folder' in title_data.keys() else 'TBD'
	edocs_footer = 'eDOCS Project Folder: #' + edocs_folder

	##RIGHT FOOTER
	#WRITE DATE PRINTED
	date_format = workbook.add_format({'italic':True})
	date_format.set_align('right')
	now = datetime.datetime.now()
	##date data
	date_printed = 'Date Printed: ' + str(now.day) + '-' + str(now.strftime("%b")) + '-' + str(now.year)
	#worksheet.write(rightmost_idx +'3', date_printed, date_format)
	footer = '&L' + edocs_footer + '&R' + date_printed
	worksheet.set_footer(footer)


	#TITLE
	#worksheet.write('A1', title,format_text) 
	cr = 5

	# FILE SPECIFIC FORMATING
	format_text = workbook.add_format(tpe_config.CONST.FORMAT_TEXT)
	item_header_format = workbook.add_format(tpe_config.CONST.ITEM_HEADER_FORMAT)

	#additional header image
	#TODO: logo image

	#==========TABLE 1: Project facts, including title==========
	cr += 1
	worksheet.merge_range('A' + str(cr) + ':' + 'C' + str(cr), 'Project Facts',  blue_title_format)
	cr += 1

	data1 = res["report_facts"]["items"][0]
	tag_list1 = ["project sponsor", "project manager", "technical specialist", "technical team", "sme"]
	name_list = ['Project Sponsor', 'PM', "Technical Specialist", 'Technical Team', 'SME']

	#worksheet.write_column('A' + str(cr) + ':A' + str(cr + 4), name_list, format_text_bold_left_border)

	for fname in tag_list1:
		fact = data1[fname].split(',') if fname in data1.keys() else "TBD"
		row_count = 0
		for name in fact:
			row_name = name_list[tag_list1.index(fname)]
			worksheet.merge_range('B' + str(cr) + ':' + 'C' + str(cr), name, format_text_normal_right_border)
			if row_count == 0:
				worksheet.write('A' + str(cr), str(row_name), format_text_bold_left_border)
			else:
				worksheet.write('A' + str(cr), "", format_text_bold_left_border)
			cr += 1
			row_count += 1

	worksheet.write_row('A' + str(cr) + ':' + 'C' +str(cr), ["", "", ""], format_text_top_border)
	cr += 1
	#==========TABLE 2: eDOCS==========
	"""
	data2 = res["edocs"]["items"]
	tag_list1 = ["project sponsor", "project manager", "technical team", "sme"]
	name_list = ['Project Sponsor', 'PM', 'Technical Team', 'SME']
	for eid, edoc in enumerate(data2):
		if edoc == 'BRD':
			worksheet.write('A' + str(cr), data2[eid]["edoc_name"],  format_text_bold)
			worksheet.write('B' + str(cr), data2[eid]["edoc_num"], format_text)
			cr += 1
	"""


	#=========TABLE 3: milestones completeness========
	data3 = res["milestones"]["items"]
	left_panel_rightmost_idx = 'C'

	#WRITE DATA
	worksheet.merge_range('A' + str(cr) + ':' + left_panel_rightmost_idx + str(cr), 'Major Activities / Milestones',  blue_title_format)
	cr +=1
	worksheet.merge_range('A' + str(cr) + ':' + 'B' + str(cr), 'Milestone', format_text_bold_all_border_left)
	worksheet.write('C' + str(cr), '% Complete', format_percentage_bold)
	cr += 1


	for mid, milestone in enumerate(data3):
		worksheet.merge_range('A' + str(cr) + ':' + 'B' + str(cr), milestone['name'],  format_text_normal_all_border)
		worksheet.write('C'+str(cr), milestone['percent'] + '%', format_percentage)
		cr += 1
	total_percent =  title_data['total_percent'] if 'total_percent' in title_data.keys() else 'TBD'
	worksheet.merge_range('A' + str(cr) + ':' + 'B' + str(cr), 'Total Completeness',  format_text_normal_all_border_grey)
	worksheet.write('C'+str(cr), str(total_percent) + '%', format_percentage_grey)
	cr += 1
	worksheet.write_row('A' + str(cr) + ':' + 'C' +str(cr), ["", "", ""], format_text_top_border)
	cr += 1

	#===========TABLE 4: Status Graph=========
	data4 = res["status_graph"]["items"]
	cr_right = 5
	rp_leftmost_idx = 'E' #right panel
	rp_rightmost_idx = 'L'

	col_name_list = ['Overall', 'Cost', 'Schedule', 'Scope', 'Team', 'Client', 'Vendor']
	tag_name_list = ['overall', 'cost', 'schedule', 'scope', 'team', 'client', 'vendor']

	worksheet.write(rp_leftmost_idx + '6', 'This Reporting Period', format_text_normal_all_border)
	worksheet.write(rp_leftmost_idx + '7', 'Last Reporting Period', format_text_normal_all_border)

	worksheet.write_row(rp_leftmost_idx + str(cr_right) + ':' + 'L', [''] + col_name_list, format_text_bold_all_border)
	cr_right += 1
	for pid, period in enumerate(data4):
		col_letter = nextCell(rp_leftmost_idx)
		for tag in tag_name_list:
			letter = data4[pid][tag] if tag in data4[pid].keys() else ""
			worksheet.write(str(col_letter) + str(cr_right), letter, getColorFormat(letter))
			col_letter = nextCell(col_letter)
		cr_right += 1


	#===========TABLE 5: Desctiption===========

	cr_right += 1
	data4 = res["description"]["items"][0]
	worksheet.merge_range(rp_leftmost_idx + str(cr_right) + ':' + rp_rightmost_idx  + str(cr_right), 'Project Description / Scope', blue_title_format)
	cr_right += 1
 
	description_text = data4['description'] if 'description' in data4.keys() else 'N/A'
	cr_bottom_line = merge_bottom_cr(cr_right, description_text)
	worksheet.merge_range(rp_leftmost_idx + str(cr_right) + ':' + rp_rightmost_idx  + str(cr_bottom_line), description_text, blue_content_format)
	cr_right = cr_bottom_line + 2

	#==========TABLE 6: report_details=========
	data5 = res["report_details"]["items"]
	report_detail_title_dict = {'MIIR':'Major Issues and Identified Risks',
								'SA':'Significant Accomplishments for this Reporting Period',
								'AP':'Actions Planned for Next Reporting Period',
								'IP':'In Progress Decision Requests & Change Requests'}
	#WRITE DATA
	for did, detail_name in enumerate(report_detail_title_dict):
		worksheet.merge_range(rp_leftmost_idx + str(cr_right) + ':' + rp_rightmost_idx + str(cr_right), report_detail_title_dict[detail_name],  blue_title_format)
		cr_right += 1
		cr_bf_write = cr_right
		for rid, request in enumerate(data5):
			if data5[rid]['req_type'] == detail_name:
				detail = data5[rid]['request']
				cr_bottom_line = merge_bottom_cr(cr_right, detail)
				worksheet.merge_range(rp_leftmost_idx + str(cr_right) + ':' + rp_rightmost_idx + str(cr_bottom_line), detail,  blue_content_format)
				cr_right = cr_bottom_line + 1
		if cr_bf_write == cr_right:
			detail = 'N/A'
			worksheet.merge_range(rp_leftmost_idx + str(cr_right) + ':' + rp_rightmost_idx + str(cr_right), detail,  blue_content_format)
			cr_right += 1
		cr_right += 1
	max_cr = max(cr, cr_right)
	worksheet.print_area('A1:L' + str(max_cr))
	worksheet.set_margins(left=0.6, right=0.6, top=0.2, bottom=0.4)
	worksheet.set_landscape()
	worksheet.set_paper(1) # 11x17 inches
	worksheet.fit_to_pages(1, 1)
	workbook.close()

	"""Excel
	xlsx_data = output.getvalue()
	return xlsx_data
	"""

	#pdf

	xlsx_data = output.getvalue()

	loadExcel = win32com.client.DispatchEx('Excel.Application')

	try:
		with tempfile.NamedTemporaryFile(delete=False) as pd:
			pd.seek(0)
			with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as ef:
				ef.write(xlsx_data)
				ef.seek(0)
				
				Excel = loadExcel.Workbooks.Open(ef.name, 0, False, 2)
				Excel.ExportAsFixedFormat(0, pd.name, 0, True, True)
				Excel.Close(SaveChanges=0)
				loadExcel.Quit()
				ef.close()
			
			pdf = open(str(pd.name) + '.pdf', 'rb')
			pdf_data = pdf.read()
			pdf.close()
			pd.close()
	finally:
		os.unlink(pd.name)
		os.unlink(str(pd.name) + '.pdf')
		os.unlink(str(ef.name))

	pythoncom.CoUninitialize()
	#return the pdf data as response
	return pdf_data
