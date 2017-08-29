# -*- coding: utf-8 -*-
import xlsxwriter
from io import BytesIO
import datetime
from . import stp_config
#each function holds a different report, dictionary maps each function to the report id
class reports(object):

	#Species Summary
	def r3(res, rid, year, con_num, asgn_num):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()
		data = res
		title = 'Species Summary'
		
		
		#MAIN DATA FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)	

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

		for idx, val in enumerate(data["items"]):
			if "species" in data["items"][idx]:
				if not data["items"][idx]["species"] in species:
					species[data["items"][idx]["species"]] = data["items"][idx]["quantity"]
				else:
					species[data["items"][idx]["species"]] += data["items"][idx]["quantity"]

		worksheet.write_row('A' + str(cr), item_fields, item_header_format)
		cr += 1

		for sid, spec in enumerate(species):
			#worksheet.write('A' + str(cr), [spec, species[spec]], format_text)
			worksheet.write('A' + str(cr), spec, format_text)
			worksheet.write('B' + str(cr), species[spec], format_num)

			cr += 1


		workbook.close()
		
		xlsx_data = output.getvalue()
		return xlsx_data

	#Top Performers
	def r4(res, rid, year, con_num, asgn_num):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()
		title = 'Top Performers'
		

		#MAIN DATA FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)	
		item_format = workbook.add_format(stp_config.CONST.ITEM_FORMAT)	

		data = res

		worksheet.set_column('A:A', 35)
		worksheet.set_column('B:Q', 7)
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.set_row(7,36)
		
		#HEADER
		#write general header and format
		rightmost_idx = 'Q'
		stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

		#additional header image
		worksheet.insert_image('J1', stp_config.CONST.ENV_LOGO,{'x_offset':180,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

		item_fields1 = ['Infill / Retrofit', 'Capital Infrastructure', 'EAB Replacement', 'Total']
		item_fields2 = ['Top Performer', 'Non Top Performer']
		item_fields3 = ['QTY', '%']

		cr = 7
		rcr = 2 + 64

		worksheet.merge_range('A' + str(cr) + ':A' + str(cr+2), 'Location', item_header_format)
		for idx, val in enumerate(item_fields1):
			worksheet.merge_range(chr(rcr) + str(cr) + ':' + chr(rcr + 3) + str(cr), val, item_header_format)
			for idx2, val2 in enumerate(item_fields2):
				worksheet.merge_range(chr(rcr) + str(cr+1) + ':' + chr(rcr + 1) + str(cr+1), val2, item_header_format)
				for idx3, val3, in enumerate(item_fields3):
					worksheet.write(chr(rcr) + str(cr+2), val3, item_header_format)
					rcr += 1
		cr += 3

		print(data["items"])
		for idx, val in enumerate(data["items"]):
			d = [list(data["items"][idx].values())[0]]
			t = list(data["items"][idx].values())[1:] 

			ex = []
			for i in range(0, 2*len(t), 1):
				ex.insert(i, t[int(i/2)] if i % 2 == 0 else str('{0:.2f}'.format(100*t[int(i/2)]/(t[int(i/2)] + t[int(i/2) - 1]))) + '%' if t[int(i/2)] + t[int(i/2)-1] > 0 else '0%') 

			worksheet.write_row('A' + str(cr), d + ex, format_text)
			cr += 1
			
			

		workbook.close()
		
		xlsx_data = output.getvalue()
		return xlsx_data

	#Costing Summary
	def r6(res, rid, year, con_num, asgn_num):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()
		data = res
		title = 'Costing Summary'
		

		data = res

		#MAIN DATA FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)	

		title_format = workbook.add_format(stp_config.CONST.TITLE_FORMAT)
		item_format_money = workbook.add_format(stp_config.CONST.ITEM_FORMAT_MONEY)
		subtitle_format = workbook.add_format(stp_config.CONST.SUBTITLE_FORMAT)
		subtotal_format = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT)
		subtotal_format_money = workbook.add_format(stp_config.CONST.SUBTOTAL_FORMAT_MONEY)

		#SET COLUMN AND ROW
		worksheet.set_column('A:G', 30)
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		
		#HEADER
		#write general header and format
		rightmost_idx = 'G'
		stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

		#additional header image
		worksheet.insert_image('F1', stp_config.CONST.ENV_LOGO,{'x_offset':180,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

		#MAIN DATA
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

			items = {'A' : {}, 'B' : {}, 'C' : {}, 'D' : {}, 'E' : {}, 'F' : {}, 'G' : {}, 'H' : {}}

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

						#changes all zeros to $0 for currency items
						for i, v in enumerate(d):
							d[i] = '$0' if i >= 1 and (d[i] == 0 or d[i] =='0' or d[i] =='$.00') else str(d[i]).lstrip(' ')

						worksheet.write_row('A' + str(cr), d, format_text)
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

		workbook.close()

		xlsx_data = output.getvalue()
		return xlsx_data

	#Bid Form Summary
	def r7(res, rid, year, con_num, asgn_num):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()
		data = res
		title = 'Bid Form Summary'
		

		data = res

		worksheet.set_column('A:F', 30)
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)

		item_fields = ['Item Number', 'Item', 'Unit', 'Quantity', 'Unit Price', 'Total']

		cr = 7

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
		rightmost_idx = 'F'
		stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

		#additional header image
		worksheet.insert_image('E1', stp_config.CONST.ENV_LOGO,{'x_offset':180,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

		#MAIN DATA

		miDict = {'A' : 'A - Tree Planting - Ball and Burlap Trees',
		'B' : 'B - Tree Planting - Potted Perennials and Grass',
		'C' : 'C - Tree Planting - Potted Shrubs',
		'D' : 'D - Transplanting',
		'E' : 'E - Stumping',
		'F' : 'F - Watering',
		'G' : 'G - Tree Maintenance',
		'H' : 'H - Automated Vehicle Locating System'}

		
		items = {'A' : {}, 'B' : {}, 'C' : {}, 'D' : {}, 'E' : {}, 'F' : {}, 'G' : {}, 'H' : {}}

		
		for idx, val in enumerate(data["items"]):
			if not data["items"][idx]["item"] in items[data["items"][idx]["sec"]]:
				items[data["items"][idx]["sec"]].update({data["items"][idx]["item"] : [(data["items"][idx]["ino"] if "ino" in data["items"][idx] else '--'),
					data["items"][idx]["unit"] if "unit" in data["items"][idx] else 'N/A',
					int(data["items"][idx]["quantity"]) if "quantity" in data["items"][idx] else 0,
					data["items"][idx]["up"] if "up" in data["items"][idx] else 0]})
			else:
				items[data["items"][idx]["sec"]][data["items"][idx]["item"]][2] += int(data["items"][idx]["quantity"])

		for idx, val in enumerate(items):
			if items[val]:
				worksheet.merge_range('A' + str(cr) + ':F' + str(cr), miDict[val], subtitle_format)
				worksheet.write_row('A' + str(cr+1) + ':F' + str(cr+1), item_fields, subtitle_format)
				cr += 2
				start = cr
				for idx2, val2 in enumerate(items[val]):
					d = [items[val][val2][0], val2, items[val][val2][1], items[val][val2][2], str(items[val][val2][3]).lstrip(' ')]

					for i, v in enumerate(d):
							d[i] = '$0' if i >= 1 and (d[i] == 0 or d[i] =='0' or d[i] =='$.00') else str(d[i]).lstrip()

					worksheet.write_row('A' + str(cr), d, format_text)
					worksheet.write_formula('F' + str(cr), '=D' + str(cr) + '*E' + str(cr), item_format_money)
					cr += 1

				worksheet.write('A' + str(cr), 'SubTotal: ', subtotal_format)
				worksheet.write('B' + str(cr), '', subtotal_format)
				worksheet.write('C' + str(cr), '', subtotal_format)
				worksheet.write_formula('D' + str(cr), '=SUM(D' + str(start) + ':D' + str(cr-1) + ')', subtotal_format)
				worksheet.write('E' + str(cr), '', subtotal_format)
				worksheet.write_formula('F' + str(cr), '=SUM(F' + str(start) + ':F' + str(cr-1) + ')', subtotal_format_money)

				worksheet.merge_range('A' + str(cr+1) + ':G' + str(cr+1),' ')
				cr += 2


		workbook.close()

		xlsx_data = output.getvalue()
		return xlsx_data

	#Warranty Report Species Analysis
	def r17(res, rid, year, con_num, asgn_num):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()

		type = 'Year 1 Warranty' if rid == '17' else 'Year 2 Warranty' if rid == '18' else '12 Month Warranty'
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

		#Contractor plant trees (Tree Planting Status)
	def r51(res, rid, year, con_num, asgn_num):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()

		data = res
		title = 'Contractor Plant Tree'
		#year = year #delete

		# MAIN DATA FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)

		#set column width
		worksheet.set_column('A:A', 15.78)
		worksheet.set_column('B:B', 31.44)
		worksheet.set_column('C:C', 12.22)
		worksheet.set_column('D:D', 11.89)
		worksheet.set_column('E:E', 9.78)
		worksheet.set_column('F:F', 11.33)
		worksheet.set_column('G:G', 11.89)
		worksheet.set_column('H:H', 11)

		#set row
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.set_row(5,23.4)
		worksheet.set_row(6, 31.2)

		#HEADER
		#write general header and format
		rightmost_idx = 'I'
		stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num) #change 2017 to year

		# FILE SPECIFIC FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)

		#additional header image
		worksheet.insert_image('F1', r'\\ykr-apexp1\staticenv\apps\199\env_internal_bw.png',{'x_offset':70,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})


		item_fields = ['Tree Planting Detail No.', 'Location', 'Assignment No.', 'Assignment Status', 'Planting Status', 'Planting Start Date', 'Planting End Date', 'Assigned Inspector', 'Status of Inspection']
		worksheet.write_row('A7', item_fields, item_header_format)


		#MAIN DATA
		cr = 8
		for idx, val in enumerate(data["items"]):
			#for idx2, val2 in enumerate(data["items"][idx]):
			#	worksheet.write(chr(idx2+65)+str(idx+8),data["items"][idx][val2], format_text)
			a1 = data["items"][idx]["tree_planting_detail_no"] if "tree_planting_detail_no" in data["items"][idx].keys() else ""
			worksheet.write('A' + str(cr), a1 if a1 is not None else "", format_text)

			a2 = data["items"][idx]["location"] if "location" in data["items"][idx].keys() else ""
			worksheet.write('B' + str(cr), a2 if a2 is not None else "", format_text)

			a3 = data["items"][idx]["assignment_num"] if "assignment_num" in data["items"][idx].keys() else ""
			worksheet.write('C' + str(cr), a3 if a3 is not None else "", format_num)

			a4 = data["items"][idx]["assignment_status"] if "assignment_status" in data["items"][idx].keys() else ""
			worksheet.write('D' + str(cr), a4 if a4 is not None else "", format_text)

			a5 = data["items"][idx]["planting_status"] if "planting_status" in data["items"][idx].keys() else ""
			worksheet.write('E' + str(cr), a5 if a5 is not None else "", format_text)

			a6 = data["items"][idx]["start_date"] if "start_date" in data["items"][idx].keys() else ""
			worksheet.write('F' + str(cr), a6 if a6 is not None else "", format_text)

			a7 = data["items"][idx]["end_date"] if "end_date" in data["items"][idx].keys() else ""
			worksheet.write('G' + str(cr), a7 if a7 is not None else "", format_text)

			a8 = data["items"][idx]["inspector"] if "inspector" in data["items"][idx].keys() else ""
			worksheet.write('H' + str(cr), a8 if a8 is not None else "", format_text)

			a9 = data["items"][idx]["inspection_status"] if "inspection_status" in data["items"][idx].keys() else""
			worksheet.write('I' + str(cr), a9 if a9 is not None else "", format_text)
			cr += 1

		workbook.close()

		xlsx_data = output.getvalue()
		return xlsx_data

	# Nursery Inspection
	def r52(res, rid, year, con_num, asgn_num):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()

		data = res
		title = 'Nuersery Inspection Report'
		

		#set column width
		worksheet.set_column('A:A', 11.22)
		worksheet.set_column('B:B', 8.67)
		worksheet.set_column('C:C', 33)
		worksheet.set_column('D:D', 35.22)
		worksheet.set_column('E:E', 12.22)
		worksheet.set_column('F:F', 19.22)
		worksheet.set_column('G:G', 10.89)
		worksheet.set_column('H:H', 10.67)

		#set row
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.set_row(5,23.4)
		worksheet.set_row(6, 31.2)

		#HEADER
		#write general header and format
		rightmost_idx = 'H'
		stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

		# FILE SPECIFIC FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)

		#additional header image
		worksheet.insert_image('F1', r'\\ykr-apexp1\staticenv\apps\199\env_internal_bw.png',{'x_offset':45,'y_offset':13, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})
		
		#COLUMN NAMES
		item_fields = ['Tree Tag Range', 'Tag Color', 'Species', 'Species Substituted For',	'Nursery', 'Stock Type', 'Farm/Lot', 'Status']
		worksheet.write_row('A7', item_fields, item_header_format)

		#MAIN DATA
		for idx, val in enumerate(data["items"]):
			for idx2, val2 in enumerate(data["items"][idx]):
					worksheet.write(chr(idx2+65)+str(idx+8),data["items"][idx][val2], format_text)

		workbook.close()
		xlsx_data = output.getvalue()
		return xlsx_data

	#Nursery Tagging Requirement
	def r53(res, rid, year, con_num, asgn_num):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()

		data = res
		title = 'Nuersery Tagging Requirement'

		#MAIN DATA FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)

		#set column width
		worksheet.set_column('A:A', 19.67)
		worksheet.set_column('B:B', 24)
		worksheet.set_column('C:C', 23.89)
		worksheet.set_column('D:D', 15.11)
		worksheet.set_column('E:E', 12.33)
		worksheet.set_column('F:F', 15.56)

		#set row
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.set_row(5,23.4)
		worksheet.set_row(6, 31.2)

		#HEADER
		#write general header and format
		rightmost_idx = 'F'
		stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)


		#additional header image
		worksheet.insert_image('D1', stp_config.CONST.ENV_LOGO, {'x_offset':75,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})
		
		#COLUMN NAMES
		item_fields = ['Stock Type', 'Plant Type', 'Species', 'Qty Required', 'Qty Tagged', 'Qty Left To Tag']
		worksheet.write_row('A7', item_fields, item_header_format)

		#MAIN DATA
		for idx, val in enumerate(data["items"]):
			for idx2, val2 in enumerate(data["items"][idx]):
				if chr(idx2 + 65) <= 'C':
					worksheet.write(chr(idx2+65)+str(idx+8),data["items"][idx][val2], format_text)
				elif chr(idx2 + 65 ) > 'C':

					worksheet.write(chr(idx2+65)+str(idx+8),data["items"][idx][val2], format_num)

		workbook.close()

		xlsx_data = output.getvalue()
		return xlsx_data

#Summary of Contract Items - All Items
	def r54(res, rid, year, con_num, asgn_num):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()

		data = res
		title = 'Summary of Contract Items - All Items'

		#MAIN DATA FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)

		#set column width
		worksheet.set_column('A:A', 12.67)
		worksheet.set_column('B:B', 39.89)
		worksheet.set_column('C:C', 12)
		worksheet.set_column('D:D', 11)
		worksheet.set_column('E:E', 30.56)
		worksheet.set_column('F:F', 9)
		worksheet.set_column('G:G', 11.33)
		worksheet.set_column('H:H', 12.67)
		worksheet.set_column('I:I', 13)
		#worksheet.set_column('J:J', 9.65)

		#set row
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.set_row(5,23.4)
		worksheet.set_row(6, 31.2)

		#HEADER
		#write general header and format
		rightmost_idx = 'I'
		stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

		#additional header image
		worksheet.insert_image('G1', stp_config.CONST.ENV_LOGO,{'x_offset':35,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})
		
		#COLUMN NAMES
		item_fields = ['Contract Item No.',	'Location', 'RINs', 'Description', 'Item', 'Quantity', 'Program', 'Municipality', 'Area forester']
		worksheet.write_row('A7', item_fields, item_header_format)


		cr = 0 #initiate cr
		#MAIN DATA
		for idx, val in enumerate(data["items"]):
			for idx2, val2 in enumerate(data["items"][idx]):
				if chr(idx2 + 65) == 'F':
					worksheet.write(chr(idx2+65)+str(idx+8),data["items"][idx][val2], format_num)
				elif chr(idx2 + 65 ) != 'F' and chr(idx2 + 65 ) != 'J':
					worksheet.write(chr(idx2+65)+str(idx+8),data["items"][idx][val2], format_text)
				cr = idx2 #record the last row num

		cr += 4

		##============TOP PERFORMAERS SUB TABLE=============
		#TOP PERFORMERS COLUMN NAMES
		item_fields = ['Contract Item No.',	'Top Performer', 'Top Performer %', 'Non Top Performer', 'Non Top Performer %', 'Total']
		worksheet.write_row( "A" + str(cr), item_fields, item_header_format)
		cr += 1
		
		#TOP PERFORMER CALCULATION
		contract_items = []
		tp = {} # tp = top performers

		for iid, item in enumerate(data["items"]):
			if not data["items"][iid]["contract_item_num"] in contract_items:
				contract_items.append(data["items"][iid]["contract_item_num"])
		
		#calculate number and overall
		tp["Overall"] = {"top_p_qty": 0, "non_top_p_qty" : 0, "total_qty" : 0}

		for cid, item in enumerate(data["items"]):
			# first time having the key
			if not data["items"][cid]["contract_item_num"] in tp.keys(): 
				if data["items"][cid]["top_performer"] == 'Y':
					#for different reports, change contract_item_num to other group by ids
					tp[data["items"][cid]["contract_item_num"]] = {"top_p_qty": data["items"][cid]["quantity"], "non_top_p_qty" : 0, "total_qty" : data["items"][cid]["quantity"]}
					tp["Overall"]["top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]
				elif data["items"][cid]["top_performer"] == 'N':
					tp[data["items"][cid]["contract_item_num"]] = {"top_p_qty": 0, "non_top_p_qty" : data["items"][cid]["quantity"], "total_qty" : data["items"][cid]["quantity"]}
					tp["Overall"]["non_top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]

			# second time having the key
			elif data["items"][cid]["contract_item_num"] in tp.keys():
				if data["items"][cid]["top_performer"] == 'Y':
					tp[data["items"][cid]["contract_item_num"]]["top_p_qty"] += data["items"][cid]["quantity"]
					tp[data["items"][cid]["contract_item_num"]]["total_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]

				elif data["items"][cid]["top_performer"] == 'N':
					tp[data["items"][cid]["contract_item_num"]]["non_top_p_qty"] += data["items"][cid]["quantity"]
					tp[data["items"][cid]["contract_item_num"]]["total_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["non_top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]

		#calculate percentage and write top performers into table
		percent_fmt = workbook.add_format({'num_format': '0.00%','border':True,'border_color':'gray',})
		rowcount = 0
		cr -= 1 # overall is the first row

		for idx, cnum in enumerate(tp):
			if cnum != "Overall":
				worksheet.write("A" + str(idx + cr), cnum, format_text)
				worksheet.write("B" + str(idx + cr), tp[cnum]["top_p_qty"], format_num)
				worksheet.write("C" + str(idx + cr), tp[cnum]["top_p_qty"]/tp[cnum]["total_qty"], percent_fmt)
				worksheet.write("D" + str(idx + cr), tp[cnum]["non_top_p_qty"], format_num)
				worksheet.write("E" + str(idx + cr), tp[cnum]["non_top_p_qty"]/tp[cnum]["total_qty"], percent_fmt)
				worksheet.write("F" + str(idx + cr), tp[cnum]["total_qty"], format_num)

				rowcount = idx

		cr += 1 # overall was the first row
		#write overall data
		worksheet.write("A" + str(rowcount + cr), "Overall", format_text)
		worksheet.write("B" + str(rowcount + cr), tp["Overall"]["top_p_qty"], format_num)
		worksheet.write("C" + str(rowcount + cr), tp["Overall"]["top_p_qty"]/tp["Overall"]["total_qty"], percent_fmt)
		worksheet.write("D" + str(rowcount + cr), tp["Overall"]["non_top_p_qty"], format_num)
		worksheet.write("E" + str(rowcount + cr), tp["Overall"]["non_top_p_qty"]/tp["Overall"]["total_qty"], percent_fmt)
		worksheet.write("F" + str(rowcount + cr), tp["Overall"]["total_qty"], format_num)

		#====ending=======

		workbook.close()

		xlsx_data = output.getvalue()
		return xlsx_data


#Summary of Contract Items by Area Forester
	def r55(res, rid, year, con_num, asgn_num):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()

		data = res
		title = 'Summary of Contract Items, Grouped by Area Forester'
		title2 = 'Top Performers, Grouped by Area Forester'

		#MAIN DATA FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)

		#set column width
		worksheet.set_column('A:A', 11.56)
		worksheet.set_column('B:B', 45.22)
		worksheet.set_column('C:C', 21.22)
		worksheet.set_column('D:D', 11.78)
		worksheet.set_column('E:E', 33.78)
		worksheet.set_column('F:F', 11.44)

		#set row
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.set_row(5,23.4)
		worksheet.set_row(6, 31.2)


		#HEADER
		#write general header and format
		rightmost_idx = 'F'
		stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

		#additional header image
		worksheet.insert_image('D1', stp_config.CONST.ENV_LOGO,{'x_offset':45,'y_offset':13, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})
		

		# main data calculation
		foresters = []
		ft = {}

		for afid, forester in enumerate(data["items"]):
			if not data["items"][afid]["area_forester"] in foresters:
				foresters.append(data["items"][afid]["area_forester"])

		# main data
		cr = 8 #current row, starting at offset where data begins
		item_fields = ['Contract Item No.', 'Location', 'RINs', 'Description', 'Item', 'Quantity']

		for afid, forester in enumerate(foresters):
			worksheet.merge_range('A' + str(cr) + ':F' + str(cr), str('Area Forester: ' + forester), format_text)
			worksheet.write_row('A' + str(cr+1), item_fields, item_header_format)
			cr += 2
			for idx, val in enumerate(data["items"]):
				if data["items"][idx]["area_forester"] == forester:
					if forester in ft:
						ft.update({forester: ft[forester] + list(data["items"][idx].values())[5]})
					else:
						ft.update({forester: list(data["items"][idx].values())[5]})
					worksheet.write_row('A' + str(cr), list(data["items"][idx].values())[0:6], format_text)
					cr += 1
			cr += 1

		worksheet.merge_range('A' + str(6) + ':F' + str(6), title2, format_text)
		cr += 1
		
		##============TOP PERFORMAERS SUB TABLE=============
		#TOP PERFORMERS COLUMN NAMES
		item_fields = ['Area Forester', 'Top Performer', 'Top Performer %', 'Non Top Performer', 'Non Top Performer %', 'Total']
		worksheet.write_row( "A" + str(cr), item_fields, item_header_format)
		cr += 1
		
		#TOP PERFORMER CALCULATION
		contract_items = []
		tp = {} # tp = top performers

		for iid, item in enumerate(data["items"]):
			if not data["items"][iid]["area_forester"] in contract_items:
				contract_items.append(data["items"][iid]["area_forester"])
		
		#calculate number and overall
		tp["Overall"] = {"top_p_qty": 0, "non_top_p_qty" : 0, "total_qty" : 0}

		for cid, item in enumerate(data["items"]):
			# first time having the key
			if not data["items"][cid]["area_forester"] in tp.keys(): 
				if data["items"][cid]["top_performer"] == 'Y':
					#for different reports, change contract_item_num to other group by ids
					tp[data["items"][cid]["area_forester"]] = {"top_p_qty": data["items"][cid]["quantity"], "non_top_p_qty" : 0, "total_qty" : data["items"][cid]["quantity"]}
					tp["Overall"]["top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]
				elif data["items"][cid]["top_performer"] == 'N':
					tp[data["items"][cid]["area_forester"]] = {"top_p_qty": 0, "non_top_p_qty" : data["items"][cid]["quantity"], "total_qty" : data["items"][cid]["quantity"]}
					tp["Overall"]["non_top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]

			# second time having the key
			elif data["items"][cid]["area_forester"] in tp.keys():
				if data["items"][cid]["top_performer"] == 'Y':
					tp[data["items"][cid]["area_forester"]]["top_p_qty"] += data["items"][cid]["quantity"]
					tp[data["items"][cid]["area_forester"]]["total_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]

				elif data["items"][cid]["top_performer"] == 'N':
					tp[data["items"][cid]["area_forester"]]["non_top_p_qty"] += data["items"][cid]["quantity"]
					tp[data["items"][cid]["area_forester"]]["total_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["non_top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]

		#calculate percentage and write top performers into table
		percent_fmt = workbook.add_format({'num_format': '0.00%','border':True,'border_color':'gray',})
		rowcount = 0
		cr -= 1 # overall is the first row

		for idx, cnum in enumerate(tp):
			if cnum != "Overall":
				worksheet.write("A" + str(idx + cr), cnum, format_text)
				worksheet.write("B" + str(idx + cr), tp[cnum]["top_p_qty"], format_num)
				worksheet.write("C" + str(idx + cr), tp[cnum]["top_p_qty"]/tp[cnum]["total_qty"], percent_fmt)
				worksheet.write("D" + str(idx + cr), tp[cnum]["non_top_p_qty"], format_num)
				worksheet.write("E" + str(idx + cr), tp[cnum]["non_top_p_qty"]/tp[cnum]["total_qty"], percent_fmt)
				worksheet.write("F" + str(idx + cr), tp[cnum]["total_qty"], format_num)

				rowcount = idx

		cr += 1 # overall was the first row
		#write overall data
		worksheet.write("A" + str(rowcount + cr), "Overall", format_text)
		worksheet.write("B" + str(rowcount + cr), tp["Overall"]["top_p_qty"], format_num)
		worksheet.write("C" + str(rowcount + cr), tp["Overall"]["top_p_qty"]/tp["Overall"]["total_qty"], percent_fmt)
		worksheet.write("D" + str(rowcount + cr), tp["Overall"]["non_top_p_qty"], format_num)
		worksheet.write("E" + str(rowcount + cr), tp["Overall"]["non_top_p_qty"]/tp["Overall"]["total_qty"], percent_fmt)
		worksheet.write("F" + str(rowcount + cr), tp["Overall"]["total_qty"], format_num)
		
		#====ending=======


		workbook.close()
		
		xlsx_data = output.getvalue()
		return xlsx_data

#Summary of Contract Items by Program
	def r56(res, rid, year, con_num, asgn_num):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()
		data = res
		title = 'Summary of Contract Items, Grouped by Program'
		title2 = 'Top Performers, Grouped by Program'

		#MAIN DATA FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)		
		
		#set column width
		worksheet.set_column('A:A', 12.56)
		worksheet.set_column('B:B', 47.78)
		worksheet.set_column('C:C', 14.11)
		worksheet.set_column('D:D', 15.78)
		worksheet.set_column('E:E', 33.89)
		worksheet.set_column('F:F', 8.99)


		#set row
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.set_row(5,23.4)
		#worksheet.set_row(6, 31.2)

		#HEADER
		#write general header and format
		rightmost_idx = 'F'
		stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

		#additional header image
		worksheet.insert_image('E1', stp_config.CONST.ENV_LOGO,{'x_offset':70,'y_offset':16, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})
	

		foresters = []
		ft = {}

		for afid, forester in enumerate(data["items"]):
			if not data["items"][afid]["program"] in foresters:
				foresters.append(data["items"][afid]["program"])

		# main data
		cr = 7 #current row, starting at offset where data begins
		item_fields = ['Contract Item No.', 'Location', 'RINs', 'Description', 'Item', 'Quantity']
		#loop over all programs
		for afid, forester in enumerate(foresters):
			worksheet.merge_range('A' + str(cr) + ':F' + str(cr), str('Program: ' + forester), format_text)
			worksheet.write_row('A' + str(cr+1), item_fields, item_header_format)
			cr += 2
			for idx, val in enumerate(data["items"]):
				if data["items"][idx]["program"] == forester:
					#if forester in ft:
				    #ft.update({forester: ft[forester] + list(data["items"][idx].values())[6]})
					#else:
				#		ft.update({forester: list(data["items"][idx].values())[6]})
					#worksheet.write_row('A' + str(cr), list(data["items"][idx].values())[0:5], format_text)
					a1 = data["items"][idx]["contract_item_num"]
					worksheet.write('A' + str(cr), a1 if a1 is not None else "", format_text)
					a2 = data["items"][idx]["location"]
					worksheet.write('B' + str(cr), a2 if a2 is not None else "", format_text)
					a3 = data["items"][idx]["rins"]
					worksheet.write('C' + str(cr), a3 if a3 is not None else "", format_text)
					a4 = data["items"][idx]["description"]
					worksheet.write('D' + str(cr), a4 if a4 is not None else "", format_text)
					a5 = data["items"][idx]["item"]
					worksheet.write('E' + str(cr), a5 if a5 is not None else "", format_text)
					a6 = data["items"][idx]["quantity"]
					worksheet.write('F' + str(cr), a6 if a6 is not None else "", format_text)
					cr += 1
			cr += 1


		##============TOP PERFORMAERS SUB TABLE=============
		#TOP PERFORMERS COLUMN NAMES
		item_fields = ['Area Forester', 'Top Performer', 'Top Performer %', 'Non Top Performer', 'Non Top Performer %', 'Total']
		worksheet.write_row( "A" + str(cr), item_fields, item_header_format)
		cr += 1
		
		#TOP PERFORMER CALCULATION
		contract_items = []
		tp = {} # tp = top performers

		for iid, item in enumerate(data["items"]):
			if not data["items"][iid]["program"] in contract_items:
				contract_items.append(data["items"][iid]["program"])
		
		#calculate number and overall
		tp["Overall"] = {"top_p_qty": 0, "non_top_p_qty" : 0, "total_qty" : 0}

		for cid, item in enumerate(data["items"]):
			# first time having the key
			if not data["items"][cid]["program"] in tp.keys(): 
				if data["items"][cid]["top_performer"] == 'Y':
					#for different reports, change contract_item_num to other group by ids
					tp[data["items"][cid]["program"]] = {"top_p_qty": data["items"][cid]["quantity"], "non_top_p_qty" : 0, "total_qty" : data["items"][cid]["quantity"]}
					tp["Overall"]["top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]
				elif data["items"][cid]["top_performer"] == 'N':
					tp[data["items"][cid]["program"]] = {"top_p_qty": 0, "non_top_p_qty" : data["items"][cid]["quantity"], "total_qty" : data["items"][cid]["quantity"]}
					tp["Overall"]["non_top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]

			# second time having the key
			elif data["items"][cid]["program"] in tp.keys():
				if data["items"][cid]["top_performer"] == 'Y':
					tp[data["items"][cid]["program"]]["top_p_qty"] += data["items"][cid]["quantity"]
					tp[data["items"][cid]["program"]]["total_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]

				elif data["items"][cid]["top_performer"] == 'N':
					tp[data["items"][cid]["program"]]["non_top_p_qty"] += data["items"][cid]["quantity"]
					tp[data["items"][cid]["program"]]["total_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["non_top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]

		#calculate percentage and write top performers into table
		percent_fmt = workbook.add_format({'num_format': '0.00%','border':True,'border_color':'gray',})
		rowcount = 0
		cr -= 1 # overall is the first row

		for idx, cnum in enumerate(tp):
			if cnum != "Overall":
				worksheet.write("A" + str(idx + cr), cnum, format_text)
				worksheet.write("B" + str(idx + cr), tp[cnum]["top_p_qty"], format_num)
				worksheet.write("C" + str(idx + cr), tp[cnum]["top_p_qty"]/tp[cnum]["total_qty"], percent_fmt)
				worksheet.write("D" + str(idx + cr), tp[cnum]["non_top_p_qty"], format_num)
				worksheet.write("E" + str(idx + cr), tp[cnum]["non_top_p_qty"]/tp[cnum]["total_qty"], percent_fmt)
				worksheet.write("F" + str(idx + cr), tp[cnum]["total_qty"], format_num)

				rowcount = idx

		cr += 1 # overall was the first row
		#write overall data
		worksheet.write("A" + str(rowcount + cr), "Overall", format_text)
		worksheet.write("B" + str(rowcount + cr), tp["Overall"]["top_p_qty"], format_num)
		worksheet.write("C" + str(rowcount + cr), tp["Overall"]["top_p_qty"]/tp["Overall"]["total_qty"], percent_fmt)
		worksheet.write("D" + str(rowcount + cr), tp["Overall"]["non_top_p_qty"], format_num)
		worksheet.write("E" + str(rowcount + cr), tp["Overall"]["non_top_p_qty"]/tp["Overall"]["total_qty"], percent_fmt)
		worksheet.write("F" + str(rowcount + cr), tp["Overall"]["total_qty"], format_num)
		
		#====ending=======


		workbook.close()
		
		xlsx_data = output.getvalue()
		return xlsx_data

#Summary of Contract Items by Mun
	def r57(res, rid, year, con_num, asgn_num):
		output = BytesIO()
		workbook = xlsxwriter.Workbook(output, {'in_memory': True})
		worksheet = workbook.add_worksheet()
		data = res
		title = 'Summary of Contract Items, Grouped by Program'
		title2 = 'Top Performers, Grouped by Program'

		#MAIN DATA FORMATING
		format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
		format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
		item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)		

		#set column width
		worksheet.set_column('A:A', 12.56)
		worksheet.set_column('B:B', 47.78)
		worksheet.set_column('C:C', 14.11)
		worksheet.set_column('D:D', 15.78)
		worksheet.set_column('E:E', 33.89)
		worksheet.set_column('F:F', 8.99)
		#worksheet.set_column('G:G', 11.33)
		#worksheet.set_column('H:H', 11.89)
		#worksheet.set_column('I:I', 11)
		#worksheet.set_column('J:J', 9.65)

		#set row
		worksheet.set_row(0,36)
		worksheet.set_row(1,36)
		worksheet.set_row(5,23.4)
		#worksheet.set_row(6, 31.2)

		#HEADER
		#write general header and format
		rightmost_idx = 'F'
		stp_config.const.write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num)

		#additional header image
		worksheet.insert_image('E1', stp_config.CONST.ENV_LOGO,{'x_offset':70,'y_offset':16, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

		foresters = []
		ft = {}

		for afid, forester in enumerate(data["items"]):
			if not data["items"][afid]["municipality"] in foresters:
				foresters.append(data["items"][afid]["municipality"])

		# main data
		cr = 7 #current row, starting at offset where data begins
		item_fields = ['Contract Item No.', 'Location', 'RINs', 'Description', 'Item', 'Quantity']
		#loop over all programs
		for afid, forester in enumerate(foresters):
			worksheet.merge_range('A' + str(cr) + ':F' + str(cr), str('Municipality: ' + forester), format_text) #was format_text
			worksheet.write_row('A' + str(cr+1), item_fields, item_header_format)
			cr += 2
			for idx, val in enumerate(data["items"]):
				if data["items"][idx]["municipality"] == forester:
					#if forester in ft:
				    #ft.update({forester: ft[forester] + list(data["items"][idx].values())[6]})
					#else:
				#		ft.update({forester: list(data["items"][idx].values())[6]})
					#worksheet.write_row('A' + str(cr), list(data["items"][idx].values())[0:5], format_text)
					a1 = data["items"][idx]["contract_item_num"]  if "contract_item_num" in data["items"][idx].keys() else ""
					worksheet.write('A' + str(cr), a1 if a1 is not None else "", format_text)

					a2 = data["items"][idx]["location"] if "location" in data["items"][idx].keys() else ""
					worksheet.write('B' + str(cr), a2 if a2 is not None else "", format_text)
					
					a3 = data["items"][idx]["rins"] if "rins" in data["items"][idx].keys() else ""
					worksheet.write('C' + str(cr), a3 if a3 is not None else "", format_text)
					
					a4 = data["items"][idx]["description"] if "description" in data["items"][idx].keys() else ""
					worksheet.write('D' + str(cr), a4 if a4 is not None else "", format_text)
					
					a5 = data["items"][idx]["item"] if "item" in data["items"][idx].keys() else ""
					worksheet.write('E' + str(cr), a5 if a5 is not None else "", format_text)
					
					a6 = data["items"][idx]["quantity"] if "quantity" in data["items"][idx].keys() else ""
					worksheet.write('F' + str(cr), a6 if a6 is not None else "", format_text)
					
					cr += 1
			cr += 1


		##============TOP PERFORMAERS SUB TABLE=============
		#TOP PERFORMERS COLUMN NAMES
		item_fields = ['Area Forester', 'Top Performer', 'Top Performer %', 'Non Top Performer', 'Non Top Performer %', 'Total']
		worksheet.write_row( "A" + str(cr), item_fields, item_header_format)
		cr += 1
		
		#TOP PERFORMER CALCULATION
		contract_items = []
		tp = {} # tp = top performers

		for iid, item in enumerate(data["items"]):
			if not data["items"][iid]["municipality"] in contract_items:
				contract_items.append(data["items"][iid]["municipality"])
		
		#calculate number and overall
		tp["Overall"] = {"top_p_qty": 0, "non_top_p_qty" : 0, "total_qty" : 0}

		for cid, item in enumerate(data["items"]):
			# first time having the key
			if not data["items"][cid]["municipality"] in tp.keys(): 
				if data["items"][cid]["top_performer"] == 'Y':
					#for different reports, change contract_item_num to other group by ids
					tp[data["items"][cid]["municipality"]] = {"top_p_qty": data["items"][cid]["quantity"], "non_top_p_qty" : 0, "total_qty" : data["items"][cid]["quantity"]}
					tp["Overall"]["top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]
				elif data["items"][cid]["top_performer"] == 'N':
					tp[data["items"][cid]["municipality"]] = {"top_p_qty": 0, "non_top_p_qty" : data["items"][cid]["quantity"], "total_qty" : data["items"][cid]["quantity"]}
					tp["Overall"]["non_top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]

			# second time having the key
			elif data["items"][cid]["municipality"] in tp.keys():
				if data["items"][cid]["top_performer"] == 'Y':
					tp[data["items"][cid]["municipality"]]["top_p_qty"] += data["items"][cid]["quantity"]
					tp[data["items"][cid]["municipality"]]["total_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]

				elif data["items"][cid]["top_performer"] == 'N':
					tp[data["items"][cid]["municipality"]]["non_top_p_qty"] += data["items"][cid]["quantity"]
					tp[data["items"][cid]["municipality"]]["total_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["non_top_p_qty"] += data["items"][cid]["quantity"]
					tp["Overall"]["total_qty"] += data["items"][cid]["quantity"]

		#calculate percentage and write top performers into table
		percent_fmt = workbook.add_format({'num_format': '0.00%','border':True,'border_color':'gray',})
		rowcount = 0
		cr -= 1 # overall is the first row

		for idx, cnum in enumerate(tp):
			if cnum != "Overall":
				worksheet.write("A" + str(idx + cr), cnum, format_text)
				worksheet.write("B" + str(idx + cr), tp[cnum]["top_p_qty"], format_num)
				worksheet.write("C" + str(idx + cr), tp[cnum]["top_p_qty"]/tp[cnum]["total_qty"], percent_fmt)
				worksheet.write("D" + str(idx + cr), tp[cnum]["non_top_p_qty"], format_num)
				worksheet.write("E" + str(idx + cr), tp[cnum]["non_top_p_qty"]/tp[cnum]["total_qty"], percent_fmt)
				worksheet.write("F" + str(idx + cr), tp[cnum]["total_qty"], format_num)

				rowcount = idx

		cr += 1 # overall was the first row
		#write overall data
		worksheet.write("A" + str(rowcount + cr), "Overall", format_text)
		worksheet.write("B" + str(rowcount + cr), tp["Overall"]["top_p_qty"], format_num)
		worksheet.write("C" + str(rowcount + cr), tp["Overall"]["top_p_qty"]/tp["Overall"]["total_qty"], percent_fmt)
		worksheet.write("D" + str(rowcount + cr), tp["Overall"]["non_top_p_qty"], format_num)
		worksheet.write("E" + str(rowcount + cr), tp["Overall"]["non_top_p_qty"]/tp["Overall"]["total_qty"], percent_fmt)
		worksheet.write("F" + str(rowcount + cr), tp["Overall"]["total_qty"], format_num)
		
		#====ending=======



		workbook.close()
		
		xlsx_data = output.getvalue()
		return xlsx_data




	d  =  {'3' : r3,
	'4' : r4,
	'6' : r6, 
	'7' : r7, 
	'17': r17, 
	'18': r17, 
	'19': r17,
	'51' : r51, 
	'52' : r52,
	'53' : r53,
	'54' : r54,
	'55' : r55,
	'56' : r56,
	'57' : r57}