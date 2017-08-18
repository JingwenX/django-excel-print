# -*- coding: utf-8 -*-
import xlsxwriter
import json
from io import BytesIO

#each function holds a different report, dictionary maps each function to the report id
class reports(object):

	def r2(res):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()
		title = 'Summary of Contract Items, Grouped by Area Forester'
		title2 = 'Top Performers, Grouped by Area Forester'

		main_header1_format = workbook.add_format({
			'bold':True,
			'font_name':'Calibri',
			'font_size':12,
			'border':2, #2 is the value for thick border
			'align':'center',
			'valign':'top',
		})

		main_header2_format = workbook.add_format({
			'font_name':'Calibri',
			'font_size':18,
			'font_color':'white',
			'border':2,
			'align':'left',
			'bg_color':'black',
		})

		title_format = workbook.add_format({
			'font_name':'Calibri',
			'font_size':18,
			'font_color':'white',
			'border':2,
			'align':'left',
			'bg_color':'gray',
		})

		header1 = workbook.add_format({
			'bold':True,
			'font_name':'Calibri',
			'font_size':12,
			'border':2,
			'align': 'left',
		})

		item_header_format = workbook.add_format({
			'bold':True,
			'font_name':'Calibri',
			'font_size':12,
			'border':2,
			'align': 'left',
			'bg_color':'gray',
		})

		item_format = workbook.add_format({
			'font_name':'Calibri',
			'font_size':12,
			'align': 'left',
		})

		data = json.loads(res)

		worksheet.set_column('A:F', 25)
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.merge_range('A1:F2','Natural Heritage and Forestry Division, Environmental Services Department',main_header1_format)
		worksheet.insert_image('A1',r'\\ykr-apexp1\staticenv\York_Logo.png',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})
		#worksheet.insert_image('D1','\\ykr-fs1.ykregion.ca\Corp\WCM\EnvironmentalServices\Toolkits\DesignComms\ENVSubbrand\HighRes',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})
		worksheet.merge_range('A4:F4','CON#'+'2014'+'-Street Tree Planting and Establishment Activities',main_header2_format)
		worksheet.merge_range('A5:F5',title,title_format)

		foresters = []
		ft = {}

		for afid, forester in enumerate(data["items"]):
			if not data["items"][afid]["area_forester"] in foresters:
				foresters.append(data["items"][afid]["area_forester"])

		cr = 6 #current row, starting at offset where data begins
		item_fields = ['Contract Item Num', 'Location', 'RINS', 'Description', 'Item', 'Quantity']

		for afid, forester in enumerate(foresters):
			worksheet.merge_range('A' + str(cr) + ':F' + str(cr), str('Area Forester: ' + forester), header1)
			worksheet.write_row('A' + str(cr+1), item_fields, item_header_format)
			cr += 2
			for idx, val in enumerate(data["items"]):
				if data["items"][idx]["area_forester"] == forester:
					if forester in ft:
						ft.update({forester: ft[forester] + list(data["items"][idx].values())[5]})
					else:
						ft.update({forester: list(data["items"][idx].values())[5]})
					worksheet.write_row('A' + str(cr), list(data["items"][idx].values())[0:6], item_format)
					cr += 1
			cr += 1

		worksheet.merge_range('A' + str(cr) + ':F' + str(cr), title2, title_format)
		cr += 1
		item_fields2 = ['Area Forester', 'Top Performer', 'Top Performer %', 'Non Top Performer', 'Non Top Performer %', 'Total']
		worksheet.write_row('A' + str(cr), item_fields2, item_header_format)
		cr += 1

		#hard coding for now, need to redo once API is complete
		for afid, forester in enumerate(foresters):
			worksheet.write_row('A' + str(cr), [forester, 0, '0%', ft[forester], '100%', ft[forester]], item_format)
			cr += 1

		worksheet.write('A' + str(cr), 'Totals: ', item_format)
		worksheet.write('F' + str(cr), sum(list(ft.values())), item_format)


		workbook.close()
		
		xlsx_data = output.getvalue()
		return xlsx_data

	d  =  {'2' : r2}