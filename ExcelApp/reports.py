import xlsxwriter
import json

#each function holds a different report, dictionary maps each function to the report id
class reports(object):

	def r2(res):
		workbook = xlsxwriter.Workbook('r2.xlsx')
		worksheet = workbook.add_worksheet()
		title = 'This is the Title'

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

		data = json.loads(res)

		worksheet.set_column('A:E',35)
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.merge_range('A1:E2','Natural Heritage and Forestry Division, Environmental Services Department',main_header1_format)
		worksheet.insert_image('A1',r'\\ykr-apexp1\staticenv\York_Logo.png',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})
		#worksheet.insert_image('D1','\\ykr-fs1.ykregion.ca\Corp\WCM\EnvironmentalServices\Toolkits\DesignComms\ENVSubbrand\HighRes',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})
		worksheet.merge_range('A4:E4','CON#'+'2014'+'-Street Tree Planting and Establishment Activities',main_header2_format)
		worksheet.merge_range('A5:E5',title,title_format)

		for idx, val in enumerate(data["items"]):
			for idx2, val2 in enumerate(data["items"][idx]):
				worksheet.write(chr(idx2+65)+str(idx+6),data["items"][idx][val2])

		workbook.close()
		return workbook

	d  =  {'2' : r2}