# -*- coding: utf-8 -*
#rid 56 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'stp_contract_summary/'
	base_url += str(params["year"])
	return base_url


#by program
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]

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
	subtitle_format = workbook.add_format(stp_config.CONST.SUBTITLE_FORMAT)

	
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
	worksheet.insert_image('E1', stp_config.CONST.ENV_LOGO,{'x_offset':60,'y_offset':22, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})


	#MAIN DATA
	##Making dict with key as group by id and distinct contract item num as value
	foresters = {}

	
	for afid, forester in enumerate(data["items"]):
		if not data["items"][afid]["program"] in foresters.keys():
			# add mun as key
			foresters[data["items"][afid]["program"]] = []
			# append contract item num to []
			if data["items"][afid]["contract_item_num"] not in foresters[data["items"][afid]["program"]]:
				foresters[data["items"][afid]["program"]].append(data["items"][afid]["contract_item_num"])
		elif data["items"][afid]["program"] in foresters.keys():
			# append contract item num to []
			if data["items"][afid]["contract_item_num"] not in foresters[data["items"][afid]["program"]]:
				foresters[data["items"][afid]["program"]].append(data["items"][afid]["contract_item_num"])


	# writing main data
	cr = 7 #current row, starting at offset where data begins
	item_fields = ['Contract Item No.', 'Location', 'RINs', 'Status', 'Item', 'Quantity']
	#loop over all programs
	for afid, forester in enumerate(foresters.keys()):
		worksheet.merge_range('A' + str(cr) + ':F' + str(cr), str('Program: ' + forester), subtitle_format) #was format_text
		worksheet.write_row('A' + str(cr+1), item_fields, item_header_format)
		cr += 2

		for cid, contract_item in enumerate(foresters[forester]):
			merge_top_idx = cr
			location = ''
			rins = ''
			description = ''

			for idx, val in enumerate(data["items"]):
				if data["items"][idx]["program"] == forester and data["items"][idx]["contract_item_num"] == contract_item:

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
					worksheet.write('F' + str(cr), a6 if a6 is not None else "", format_num)
					
					cr += 1
					merge_bottom_idx = cr - 1
					location = a2
					rins = a3
					description = a4


			worksheet.merge_range('A'+ str(merge_top_idx) + ':A' + str(merge_bottom_idx) , contract_item, format_text)
			worksheet.merge_range('B'+ str(merge_top_idx) + ':B' + str(merge_bottom_idx) , location, format_text)
			worksheet.merge_range('C'+ str(merge_top_idx) + ':C' + str(merge_bottom_idx) , rins, format_text)
			worksheet.merge_range('D'+ str(merge_top_idx) + ':D' + str(merge_bottom_idx) , description, format_text)
			worksheet.set_row(cr,stp_config.CONST.BREAKDOWN_INBETWEEN_HEIGHT)
			

		cr += 1


	##============TOP PERFORMAERS SUB TABLE=============
	#TOP PERFORMERS COLUMN NAMES
	item_fields = ['Program', 'Top Performer', 'Top Performer %', 'Non Top Performer', 'Non Top Performer %', 'Total']

	
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
	#cr -= 1 # overall is the first row
	if tp["Overall"]["total_qty"] >0 :

		worksheet.write_row( "A" + str(cr), item_fields, item_header_format)
		#cr += 1
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