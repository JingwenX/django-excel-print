#rid 8 STP
# -*- coding: utf-8 -*
import xlsxwriter
from io import BytesIO
import datetime
from .. import tpe_config
import win32com.client
import pythoncom
import tempfile
import os
import string
import datetime

def form_url(params):
	base_url = 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/tpe/roadmap_data/'
	return base_url

#Tree Planting Details
def render(res, params):
	pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
	rid = params["rid"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheet = workbook.add_worksheet()
	title = 'eRoadMap'
	
	data = res

	worksheet.set_margins(left=0.4, right=0.4, top=0.4, bottom=0.4)
	worksheet.set_landscape()
	worksheet.set_paper(3) # 11x17 inches

	worksheet.set_column('A:A', 60)
	worksheet.set_column('B:B', 17)
	worksheet.set_column('C:C', 17)
	worksheet.set_column('D:D', 10)
	worksheet.set_column('E:E', 17)
	worksheet.set_column('F:F', 12)
	worksheet.set_column('G:G', 15)
	worksheet.set_column('H:H', 15)
	worksheet.set_column('I:I', 15)
	worksheet.set_column('J:J', 15)
	worksheet.set_column('K:K', 12)
	worksheet.set_column('L:L', 8)
	worksheet.set_column('M:M', 12)

	item_fields = [ "Project Name", "Project Status", "% Complete", "Proj. No.",	"Scheduled Start Year", "Exp. Start Date mm/dd/yy", "Exp. End Date mm/dd/yy", 
		"Prev. End Date mm/dd/yy", "Project Sponsor", "Performance Plan", "Branch Ranking", "Project Type", "Program"]

	formats = {
		'text' : workbook.add_format(tpe_config.CONST.text),
		'date' : workbook.add_format(tpe_config.CONST.date_format),
		'text_left' : workbook.add_format(tpe_config.CONST.text_left),
		'item_header_format' : workbook.add_format(tpe_config.CONST.item_header_format),
		'subheader' : workbook.add_format(tpe_config.CONST.subheader),
		'OMM_header' : workbook.add_format(tpe_config.CONST.OMM_header),
		'OMM_text_high' : workbook.add_format(tpe_config.CONST.OMM_text_high),
		'OMM_text_high_left' : workbook.add_format(tpe_config.CONST.OMM_text_high_left),
		'EPP_header' : workbook.add_format(tpe_config.CONST.EPP_header),
		'EPP_text_high' : workbook.add_format(tpe_config.CONST.EPP_text_high),
		'EPP_text_high_left' : workbook.add_format(tpe_config.CONST.EPP_text_high_left),
		'BPOS_header' : workbook.add_format(tpe_config.CONST.BPOS_header),
		'BPOS_text_high' : workbook.add_format(tpe_config.CONST.BPOS_text_high),
		'BPOS_text_high_left' : workbook.add_format(tpe_config.CONST.BPOS_text_high_left),
		'IAM_header' : workbook.add_format(tpe_config.CONST.IAM_header),
		'IAM_text_high' : workbook.add_format(tpe_config.CONST.IAM_text_high),
		'IAM_text_high_left' : workbook.add_format(tpe_config.CONST.IAM_text_high_left),
		'SI_header' : workbook.add_format(tpe_config.CONST.SI_header),
		'SI_text_high' : workbook.add_format(tpe_config.CONST.SI_text_high),
		'SI_text_high_left' : workbook.add_format(tpe_config.CONST.SI_text_high_left),
		'CPD_header' : workbook.add_format(tpe_config.CONST.CPD_header),
		'CPD_text_high' : workbook.add_format(tpe_config.CONST.CPD_text_high),
		'CPD_text_high_left' : workbook.add_format(tpe_config.CONST.CPD_text_high_left),
		'DEPARTMENTAL_header' : workbook.add_format(tpe_config.CONST.Departmental_header),
		'DEPARTMENTAL_text_high' : workbook.add_format(tpe_config.CONST.Departmental_text_high),
		'DEPARTMENTAL_text_high_left' : workbook.add_format(tpe_config.CONST.Departmental_text_high_left),
		'dot' : workbook.add_format(tpe_config.CONST.dot),
		'ydot' : workbook.add_format(tpe_config.CONST.ydot),
		'gdot' : workbook.add_format(tpe_config.CONST.gdot),
	}

	#HEADER
	#write general header and format
	rightmost_idx = 'M'
	worksheet.insert_image('A1',r'\\ykr-apexp1\staticenv\York_Logo.png',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})
	worksheet.set_row(0,25)
	worksheet.set_row(1,25)
	worksheet.merge_range('A1:K1', '                             Business Operations & Technology Support', workbook.add_format({'font_size':18,'bold':True,'align':'center','border':False}))
	worksheet.merge_range('A2:K2', '                             2017-2020 ENV IT Projects Roadmap', workbook.add_format({'font_size':18,'bold':True,'align':'center','border':False}))

	now = datetime.datetime.now()
	date_printed = 'Date Printed: ' + str(now.month) + '/' + str(now.day) + '/' + str(now.year)
	worksheet.merge_range('L1:M2', date_printed, formats['date'])
	footer = 'Page &P of &N'
	worksheet.set_footer(footer)

	#MAIN DATA
	projects = {
		'OMM' : [],
		'EPP' : [],
		'BPOS' : [],
		'IAM' : [],
		'SI' : [],
		'CPD' : [],
		'DEPARTMENTAL' : [],

	}

	for idx, val in enumerate(data["items"]):
		if (not data["items"][idx]["branch"] in projects):
			projects.update({data["items"][idx].get("branch") : [[
				data["items"][idx].get("proj_name"),
				data["items"][idx].get("project_status"),
				data["items"][idx].get("status"),
				data["items"][idx].get("text_color"),
				data["items"][idx].get("completeness"),
				#project number
				data["items"][idx].get("start_year"),
				data["items"][idx].get("exp_start_date"),
				data["items"][idx].get("exp_end_date"),
				data["items"][idx].get("prev_end_date"),
				data["items"][idx].get("proj_sponsor"),
				data["items"][idx].get("perf_plan"),
				data["items"][idx].get("rank_branch"),
				data["items"][idx].get("proj_type"),
				data["items"][idx].get("program"),

				]]})
		else:
			projects[data["items"][idx].get("branch")].append([
				data["items"][idx].get("proj_name"),
				data["items"][idx].get("project_status"),
				data["items"][idx].get("status"),
				data["items"][idx].get("text_color"),
				data["items"][idx].get("completeness"),
				#project number
				data["items"][idx].get("start_year"),
				data["items"][idx].get("exp_start_date"),
				data["items"][idx].get("exp_end_date"),
				data["items"][idx].get("prev_end_date"),
				data["items"][idx].get("proj_sponsor"),
				data["items"][idx].get("perf_plan"),
				data["items"][idx].get("rank_branch"),
				data["items"][idx].get("proj_type"),
				data["items"][idx].get("program"),

				])

	breaks = []
	cr = 3
	worksheet.write_row('A3', item_fields, formats['item_header_format'])
	worksheet.merge_range('A4:B4', '2017-2020 ENV IT Project Roadmap', formats['subheader'])
	worksheet.write('C4', data["items"][0].get("total_num"), formats['subheader'])
	worksheet.merge_range('D4:M4',' ', formats['subheader'])
	cr += 2

	for bid, branch in enumerate(projects):
		worksheet.write('A{}'.format(cr), str(branch) + " Project", formats[str(branch) + '_header'])
		worksheet.write('B{}'.format(cr), ' ', formats[str(branch) + '_header'])
		worksheet.write('C{}'.format(cr), len(projects[branch]), formats[str(branch) + '_header'])
		worksheet.merge_range('D{}:M{}'.format(cr,cr),' ', formats[str(branch) + '_header'])
		cr += 1
		start = cr
		for pid, proj in enumerate(projects[branch]):
			worksheet.write('A{}'.format(cr), projects[branch][pid][0], formats[str(branch) + '_text_high_left'] if projects[branch][pid][3] == 0 else formats['text_left'])
			worksheet.write('B{}'.format(cr), u"\u25CB" if projects[branch][pid][2] == 'N' else u"\u25CF", (formats['ydot'] if projects[branch][pid][2] == 'Y' 
				else formats['dot'] if projects[branch][pid][2] == 'N' else formats['gdot'])) #todo: format conditions
			
			worksheet.write('C{}'.format(cr), projects[branch][pid][4], formats[str(branch) + '_text_high'] if projects[branch][pid][3] == 0 else formats['text'])
			worksheet.write('D{}'.format(cr), (cr - start + 1), formats[str(branch) + '_text_high'] if projects[branch][pid][3] == 0 else formats['text'])  

			

			worksheet.write_row('E{}'.format(cr), projects[branch][pid][5:], formats[str(branch) + '_text_high'] if projects[branch][pid][3] == 0 else formats['text'])
			cr += 1

	#worksheet.fit_to_pages(1, math.ceil(cr/50))
	worksheet.fit_to_pages(1, 0)
	workbook.close()

	return_data = output.getvalue()

	loadExcel = win32com.client.gencache.EnsureDispatch('Excel.Application')

	if(params['asPDF'] == '1'):
		try:
			with tempfile.NamedTemporaryFile(delete=False) as pd:
				pd.seek(0)
				with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as ef:
					ef.write(return_data)
					ef.seek(0)
					
					Excel = loadExcel.Workbooks.Open(ef.name, 0, False, 2)
					Excel.ExportAsFixedFormat(0, pd.name, 0, True, True)
					Excel.Close(SaveChanges=0)
					loadExcel.Quit()
					ef.close()
				
				pdf = open(str(pd.name) + '.pdf', 'rb')
				return_data = pdf.read()
				pdf.close()
				pd.close()
		finally:
			os.unlink(pd.name)
			os.unlink(str(pd.name) + '.pdf')
			os.unlink(str(ef.name))

		
	pythoncom.CoUninitialize()

	return return_data