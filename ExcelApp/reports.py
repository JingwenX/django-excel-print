# -*- coding: utf-8 -*-
import xlsxwriter
from io import BytesIO

#each function holds a different report, dictionary maps each function to the report id
class reports(object):

	#Summary of Contract Items by Area Forester
	def r2(res, rid):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()
		title = 'Summary of Contract Items, Grouped by Area Forester'
		title2 = 'Top Performers, Grouped by Area Forester'
		year = '2017'

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

		data = res

		worksheet.set_column('A:F', 25)
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.merge_range('A1:F2','Natural Heritage and Forestry Division, Environmental Services Department',main_header1_format)
		worksheet.insert_image('A1',r'\\ykr-apexp1\staticenv\York_Logo.png',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})
		#worksheet.insert_image('D1','\\ykr-fs1.ykregion.ca\Corp\WCM\EnvironmentalServices\Toolkits\DesignComms\ENVSubbrand\HighRes',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})
		worksheet.merge_range('A4:F4','CON#'+ year +'-Street Tree Planting and Establishment Activities',main_header2_format)
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

	#Costing Summary
	def r6(res, rid):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()

		title = 'Costing Summary'
		year = '2014'

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

		subtitle_format = workbook.add_format({
			'font_name':'Calibri',
			'font_size': 14,
			'font_color':'black',
			'border':2,
			'align':'left',
			'bg_color':'#D3D3D3',
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

		item_format.set_text_wrap()

		item_format_money = workbook.add_format({
			'font_name':'Calibri',
			'font_size':12,
			'align': 'left',
			'num_format': '$#,##0',
		})

		item_format_money.set_text_wrap()

		subtotal_format = workbook.add_format({
			'font_name':'Calibri',
			'font_size': 14,
			'font_color':'black',
			'align':'left',
			'bg_color':'#D3D3D3',
		})

		subtotal_format_money = workbook.add_format({
			'font_name':'Calibri',
			'font_size': 14,
			'font_color':'black',
			'align':'left',
			'bg_color':'#D3D3D3',
			'num_format': '$#,##0',
		})

		data = res

		worksheet.set_column('A:G', 30)
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.merge_range('A1:G2','Natural Heritage and Forestry Division, Environmental Services Department',main_header1_format)
		worksheet.insert_image('A1',r'\\ykr-apexp1\staticenv\York_Logo.png',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})
		#worksheet.insert_image('D1','\\ykr-fs1.ykregion.ca\Corp\WCM\EnvironmentalServices\Toolkits\DesignComms\ENVSubbrand\HighRes',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})
		worksheet.merge_range('A4:G4','CON#'+ year +'-Street Tree Planting and Establishment Activities',main_header2_format)
		worksheet.merge_range('A5:G5',title,title_format)
		worksheet.merge_range('A6:G6',' ')
		item_fields = ['Item', 'Quantity', 'Last Year Price', 'This Year Estimate', 'This Year Actual', 'Estimated Total', 'Total']

		programs = {'Capital Infrastructure' : [],
		'Infill / Retrofit' : [],
		'EAB Replacement' : []} 

		#separates items by programs
		for idx, val in enumerate(data["items"]):
			if data["items"][idx]["program"] == "Capital Infrastructure":
				programs['Capital Infrastructure'].append(data["items"][idx])
			elif data["items"][idx]["program"] == "Infill / Retrofit":
				programs['Infill / Retrofit'].append(data["items"][idx])
			else:
				programs['EAB Replacement'].append(data["items"][idx])

		cr = 7

		#print(programs)

		for pid, program in enumerate(programs):
			if programs[program]:
				worksheet.merge_range('A' + str(cr) + ':G' + str(cr), program, title_format)
				worksheet.write_row('A' + str(cr+1), item_fields, item_header_format)
				cr += 2

			miDict = {'A' : 'Tree Planting - Ball and Burlap Trees',
			'B' : 'Tree Planting - Potted Perennials and Grass',
			'C' : 'Tree Planting - Potted Shrubs',
			'D' : 'Transplanting',
			'E' : 'Stumping',
			'F' : 'Watering',
			'G' : 'Tree Maintenance',
			'H' : 'Automated Vehicle Locating System'}

			#print(programs)
			items = {'A' : {}, 'B' : {}, 'C' : {}, 'D' : {}, 'E' : {}, 'F' : {}, 'G' : {}, 'H' : {}}

			#this is good
			for idx, val in enumerate(programs[program]):
				if "item" in programs[program][idx]:
					if not programs[program][idx]["item"] in items[programs[program][idx]["sec"]]:
						items[programs[program][idx]["sec"]].update({programs[program][idx]["item"] : [int(programs[program][idx]["quantity"]),
							programs[program][idx]["lyp"] if "lyp" in programs[program][idx] else 0,
							programs[program][idx]["pe"] if "pe" in programs[program][idx] else 0,
							programs[program][idx]["up"] if "up" in programs[program][idx] else 0]})
					else:
						items[programs[program][idx]["sec"]][programs[program][idx]["item"]][0] += int(programs[program][idx]["quantity"])

			for idx, val in enumerate(items):
				if items[val]:
					worksheet.merge_range('A' + str(cr) + ':G' + str(cr), miDict[val], subtitle_format)
					cr += 1
					start = cr
					for idx2, val2 in enumerate(items[val]):
						d = [val2]
						d.extend(items[val][val2])
						for i, v in enumerate(d):
							d[i] = '$0' if i >= 2 and d[i] == 0 else d[i]

						worksheet.write_row('A' + str(cr), d, item_format)
						worksheet.write_formula('F' + str(cr), '=B' + str(cr) + '*D' + str(cr), item_format_money)
						worksheet.write_formula('G' + str(cr), '=B' + str(cr) + '*E' + str(cr), item_format_money)
						cr += 1

					worksheet.write('A' + str(cr), 'SubTotal: ', subtotal_format)
					worksheet.write_formula('B' + str(cr), '=SUM(B' + str(start) + ':B' + str(cr-1) + ')', subtotal_format)
					worksheet.write('C' + str(cr), '', subtotal_format)
					worksheet.write('D' + str(cr), '', subtotal_format)
					worksheet.write('E' + str(cr), '', subtotal_format)
					worksheet.write_formula('F' + str(cr), '=SUM(F' + str(start) + ':F' + str(cr-1) + ')', subtotal_format_money)
					worksheet.write_formula('G' + str(cr), '=SUM(G' + str(start) + ':G' + str(cr-1) + ')', subtotal_format_money)

					worksheet.merge_range('A' + str(cr+1) + ':G' + str(cr+1),' ')
					cr += 2


		#print(programs)
		workbook.close()

		xlsx_data = output.getvalue()
		return xlsx_data

	def r7(res, rid):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()

		title = 'Bid Summary'
		year = '2014'

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

		subtitle_format = workbook.add_format({
			'font_name':'Calibri',
			'font_size': 14,
			'font_color':'black',
			'border':2,
			'align':'left',
			'bg_color':'#D3D3D3',
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

		item_format.set_text_wrap()

		item_format_money = workbook.add_format({
			'font_name':'Calibri',
			'font_size':12,
			'align': 'left',
			'num_format': '$#,##0',
		})

		item_format_money.set_text_wrap()

		subtotal_format = workbook.add_format({
			'font_name':'Calibri',
			'font_size': 14,
			'font_color':'black',
			'align':'left',
			'bg_color':'#D3D3D3',
		})

		subtotal_format_money = workbook.add_format({
			'font_name':'Calibri',
			'font_size': 14,
			'font_color':'black',
			'align':'left',
			'bg_color':'#D3D3D3',
			'num_format': '$#,##0',
		})

		data = res

		worksheet.set_column('A:F', 30)
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.merge_range('A1:G2','Natural Heritage and Forestry Division, Environmental Services Department',main_header1_format)
		worksheet.insert_image('A1',r'\\ykr-apexp1\staticenv\York_Logo.png',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})
		#worksheet.insert_image('D1','\\ykr-fs1.ykregion.ca\Corp\WCM\EnvironmentalServices\Toolkits\DesignComms\ENVSubbrand\HighRes',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})
		worksheet.merge_range('A4:G4','CON#'+ year +'-Street Tree Planting and Establishment Activities',main_header2_format)
		worksheet.merge_range('A5:G5',title,title_format)
		worksheet.merge_range('A6:G6',' ')
		item_fields = ['Item Number', 'Item', 'Unit', 'Quantity', 'Unit Price', 'Total']

		cr = 7

		#print(programs)

		for idx, val in enumerate(data["items"]):

			miDict = {'A' : 'A - Tree Planting - Ball and Burlap Trees',
			'B' : 'B - Tree Planting - Potted Perennials and Grass',
			'C' : 'C - Tree Planting - Potted Shrubs',
			'D' : 'D - Transplanting',
			'E' : 'E - Stumping',
			'F' : 'F - Watering',
			'G' : 'G - Tree Maintenance',
			'H' : 'H - Automated Vehicle Locating System'}

			print(data["items"])
			items = {'A' : {}, 'B' : {}, 'C' : {}, 'D' : {}, 'E' : {}, 'F' : {}, 'G' : {}, 'H' : {}}

			#this is good
			for idx, val in enumerate(programs[program]):
				if "item" in programs[program][idx]:
					if not programs[program][idx]["item"] in items[programs[program][idx]["sec"]]:
						items[programs[program][idx]["sec"]].update({programs[program][idx]["item"] : [int(programs[program][idx]["quantity"]),
							programs[program][idx]["lyp"] if "lyp" in programs[program][idx] else 0,
							programs[program][idx]["pe"] if "pe" in programs[program][idx] else 0,
							programs[program][idx]["up"] if "up" in programs[program][idx] else 0]})
					else:
						items[programs[program][idx]["sec"]][programs[program][idx]["item"]][0] += int(programs[program][idx]["quantity"])

			for idx, val in enumerate(items):
				if items[val]:
					worksheet.merge_range('A' + str(cr) + ':G' + str(cr), miDict[val], subtitle_format)
					cr += 1
					start = cr
					for idx2, val2 in enumerate(items[val]):
						d = [val2]
						d.extend(items[val][val2])
						worksheet.write_row('A' + str(cr), d, item_format)
						worksheet.write_formula('F' + str(cr), '=B' + str(cr) + '*D' + str(cr), item_format_money)
						worksheet.write_formula('G' + str(cr), '=B' + str(cr) + '*E' + str(cr), item_format_money)
						cr += 1

					worksheet.write('A' + str(cr), 'SubTotal: ', subtotal_format)
					worksheet.write_formula('B' + str(cr), '=SUM(B' + str(start) + ':B' + str(cr-1) + ')', subtotal_format)
					worksheet.write('C' + str(cr), '', subtotal_format)
					worksheet.write('D' + str(cr), '', subtotal_format)
					worksheet.write('E' + str(cr), '', subtotal_format)
					worksheet.write_formula('F' + str(cr), '=SUM(F' + str(start) + ':F' + str(cr-1) + ')', subtotal_format_money)
					worksheet.write_formula('G' + str(cr), '=SUM(G' + str(start) + ':G' + str(cr-1) + ')', subtotal_format_money)

					worksheet.merge_range('A' + str(cr+1) + ':G' + str(cr+1),' ')
					cr += 2


		#print(programs)
		workbook.close()

		xlsx_data = output.getvalue()
		return xlsx_data

		#Warranty Report Species Analysis
	def r17(res, rid):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()

		type = 'Year 1 Warranty' if rid == '17' else 'Year 2 Warranty' if rid == '18' else '12 Month Warranty'
		title = 'Warranty Report Species Analysis ' + type
		year = '2017'

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

		data = res

		worksheet.set_column('A:E', 25)
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.merge_range('A1:E2','Natural Heritage and Forestry Division, Environmental Services Department',main_header1_format)
		worksheet.insert_image('A1',r'\\ykr-apexp1\staticenv\York_Logo.png',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})
		#worksheet.insert_image('D1','\\ykr-fs1.ykregion.ca\Corp\WCM\EnvironmentalServices\Toolkits\DesignComms\ENVSubbrand\HighRes',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})
		worksheet.merge_range('A4:E4','CON#'+ year +'-Street Tree Planting and Establishment Activities',main_header2_format)
		worksheet.merge_range('A5:E5',title,title_format)
		worksheet.merge_range('A6:E6',' ')
		item_fields = ['Species', 'Total Trees Inspected', 'Number of Trees Accepted', 'Number of Trees Rejected', 'Number of Trees Missing']
		worksheet.write_row('A7', item_fields, item_header_format)


		#MAIN DATA
		cr = 8
		species = []
		totals = {}

		for sid, spec in enumerate(data["items"]):
			if not data["items"][sid]["species"] in species:
				species.append(data["items"][sid]["species"])

		species.sort()

		for sid, spec in enumerate(species):
			for idx, val in enumerate(data["items"]):
				if data["items"][idx]["species"] == spec and "warrantyaction" in data["items"][idx]:
					temp = [1,1 if data["items"][idx]["warrantyaction"] == 'Accept' else 0,
					1 if data["items"][idx]["warrantyaction"] == 'Reject' else 0,
					1 if data["items"][idx]["warrantyaction"] == 'Missing Tree' else 0]

					if spec in totals:
						totals[spec] = [totals[spec][0] + temp[0], totals[spec][1] + temp[1], totals[spec][2] + temp[2], totals[spec][3] + temp[3]]

					else:
						totals[spec] = [temp[0], temp[1], temp[2], temp[3]]


			worksheet.write('A' + str(cr), species[sid], item_format)
			worksheet.write_row('B' + str(cr), totals[spec], item_format)
			cr += 1

		#FORMULAE AND FOOTERS
		worksheet.write('A' + str(cr), 'Totals: ', item_format)
		worksheet.write_formula('B' + str(cr), '=SUM(B8:B' + str(cr-1) + ')', item_format)
		worksheet.write_formula('C' + str(cr), '=SUM(C8:C' + str(cr-1) + ')', item_format)
		worksheet.write_formula('D' + str(cr), '=SUM(D8:D' + str(cr-1) + ')', item_format)
		worksheet.write_formula('E' + str(cr), '=SUM(E8:E' + str(cr-1) + ')', item_format)

		workbook.close()

		xlsx_data = output.getvalue()
		return xlsx_data

	d  =  {'2' : r2, '6' : r6, '7' : r7, '17': r17, '18': r17, '19': r17}