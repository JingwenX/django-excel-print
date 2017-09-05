# -*- coding: utf-8 -*
#rid 8 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_tree_planting/'
	base_url += str(params["year"])
	base_url += '/' + str(params["item_num"])
	return base_url

#Tree Planting Details
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheets = []
	title = 'Tree Planting Detail'
	
	data = res

	item_fields = ["Description of Tree Planting Locations", "Mark Type", "Mark Location", "Offset From Mark", 
						"Spacing", "Item", "Quantity", "Hydro", "Comments"]
	item_fields2 = ["Species Summary, This Location", "Number of Trees"]

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
	rightmost_idx = 'I'
	
	#MAIN DATA
	cons_main = {}

	for idx, val in enumerate(data["items"]):
		if (data["items"][idx]["type_id"] in [1,2,3,5,6] and (str(assign_num) == '-1' and not "assignment_num" in data["items"][idx]) or 
			(("assignment_num" in data["items"][idx]) and (str(data["items"][idx]["assignment_num"]) == str(assign_num)))):
			if not data["items"][idx]["detail_num"] in cons_main:
				cons_main.update({data["items"][idx]["detail_num"] : {
					"Municipality" : data["items"][idx]["municipality"] if "municipality" in data["items"][idx] else "none",
					"Regional Road" : data["items"][idx]["regional_road"] if "regional_road" in data["items"][idx] else "none",
					"Between Road 1" : data["items"][idx]["between_road_1"] if "between_road_1" in data["items"][idx] else "none",
					"Between Road 2" : data["items"][idx]["between_road_2"] if "between_road_2" in data["items"][idx] else "none",
					"RINs" : data["items"][idx]["rins"] if "rins" in data["items"][idx] else "none",
					"Contract Item No." : data["items"][idx]["contract_item_num"] if "contract_item_num" in data["items"][idx] else "none",
					"Tree Planting Detail No." : data["items"][idx]["detail_num"] if "detail_num" in data["items"][idx] else "none"
				}}) 

	for idx, val in enumerate(cons_main):
		worksheets.append(workbook.add_worksheet(val))
		worksheets[idx].set_column('A:A', 40)
		worksheets[idx].set_column('B:I', 18)
		worksheets[idx].set_row(0,36)
		worksheets[idx].set_row(1,36)

		stp_config.const.write_gen_title(title, workbook, worksheets[idx], rightmost_idx, year, con_num)

		#additional header image
		worksheets[idx].insert_image('G1', stp_config.CONST.ENV_LOGO,{'x_offset':150,'y_offset':18, 'x_scale':0.5,'y_scale':0.5, 'positioning':2})

		cr = 7
		for idx2, val2 in enumerate(cons_main[val]):
			worksheets[idx].write_row('A' + str(cr), [val2, cons_main[val][val2]], format_text)
			cr += 1

		locs = {#location : [[]]}
		}

		summary = {}

		for i, v in enumerate(data["items"]):
			if (data["items"][i]["type_id"] in [1,2,3,5,6] and ((str(assign_num) == '-1' and not "assignment_num" in data["items"][i]) or 
			(("assignment_num" in data["items"][i]) and (str(data["items"][i]["assignment_num"]) == str(assign_num))))):
				if data["items"][i]["detail_num"] == val:
					loc = data["items"][i]["regional_road"] + ', ' + data["items"][i]["between_road_1"] + ' to ' + data["items"][i]["between_road_2"]
					tItem = ((data["items"][i].get("stock_type", "--") + ' - ' +
						data["items"][i].get("plant_type", "--")  + ' - ' +
						data["items"][i].get("species", "--")) if data["items"][i]["type_id"] == 1
						else data["items"][i]["stump_size"] if data["items"][i]["type_id"] == 2
						else data["items"][i]["transp_dis"] if data["items"][i]["type_id"] == 3
						else 'Supplemental Tree Maintenance' if data["items"][i]["type_id"] == 5
						else 'Extra Work' if data["items"][i]["type_id"] == 6
						else ' ')

					summary[tItem] = summary.get(tItem, 0) + data["items"][i].get("quantity", 0)

					if not loc in locs:
						locs.update({loc : [[ 
							data["items"][i].get("roadside", " "),
							data["items"][i].get("mark_type", " "),
							data["items"][i].get("marking_location", " "),
							data["items"][i].get("offset_from_mark", " "),
							data["items"][i].get("spacing_on_centre", " "),
							tItem,
							data["items"][i]["quantity"] if "quantity" in data["items"][i] else ' ',
							data["items"][i]["hydro"] if "hydro" in data["items"][i] else ' ',
							data["items"][i].get("comments", ' ')
							]]})
					else:
						locs[loc].append([
							data["items"][i].get("roadside", " "), 
							data["items"][i]["mark_type"] if "mark_type" in data["items"][i] else ' ',
							data["items"][i]["marking_location"] if "marking_location" in data["items"][i] else ' ',
							data["items"][i]["offset_from_mark"] if "offset_from_mark" in data["items"][i] else ' ',
							data["items"][i]["spacing_on_centre"] if "spacing_on_centre" in data["items"][i] else ' ',
							tItem,
							data["items"][i]["quantity"] if "quantity" in data["items"][i] else ' ',
							data["items"][i]["hydro"] if "hydro" in data["items"][i] else ' ', 
							data["items"][i].get("comments", ' ')
							])


		for side in ['North', 'South', 'East', 'West', 'Centre Median']:

			for lid, loc in enumerate(locs):
				tLoc = []
				for l in locs[loc]:
					if l[0] == side:
						tLoc.append(l[1:])
						print(tLoc)
				if tLoc:
					cr += 1
					worksheets[idx].write('A' + str(cr), side, item_header_format)
					worksheets[idx].write_row('A' + str(cr+1), item_fields, item_header_format)
					cr += 2

					if not len(tLoc) == 1:
						worksheets[idx].merge_range('A' + str(cr) + ':A' + str(cr + len(tLoc) - 1), loc, format_text)
					else:
						worksheets[idx].write('A' + str(cr), loc, format_text)

					for item in tLoc:
						worksheets[idx].write_row('B' + str(cr), item, format_text)
						cr += 1

		cr += 1
		worksheets[idx].write_row('A' + str(cr), item_fields2, item_header_format)
		cr += 1

		tStart = cr
		for sid, item in enumerate(summary):
			worksheets[idx].write_row('A' + str(cr), [item, summary[item]], format_text)
			cr += 1
		worksheets[idx].write('A'+str(cr), "Total: ", subtotal_format)
		worksheets[idx].write_formula('B' + str(cr), '=SUM(B' + str(tStart) + ':B' + str(cr-1) + ')', format_text)


	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data