# -*- coding: utf-8 -*
#rid 70 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

def form_url(params):
	base_url = str(stp_config.CONST.API_URL_PREFIX) + 'bsmart_data/bsmart_data/stp_ws/stp_issued_watering_assignment/{}/{}/{}'.format(str(params["year"]), str(params["assign_num"]), str(params["item_num"]))
	return base_url


#Issued Watering Assignment
def render(res, params):

	rid = params["rid"]
	year = params["year"]
	con_num = params["con_num"]
	assign_num = params["assign_num"]
	item_num = params["item_num"]

	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheet = workbook.add_worksheet()

	data = res
	title = 'Issued Watering Assignment'

	#MAIN DATA FORMATING
	format_text = workbook.add_format(stp_config.CONST.FORMAT_TEXT)
	format_text_bold =  workbook.add_format(stp_config.CONST.FORMAT_TEXT)
	#format_text_bold.add_format({'bold':True})
	format_num = workbook.add_format(stp_config.CONST.FORMAT_NUM)
	format_text.set_locked(False)
	format_num.set_locked(False)
	item_header_format = workbook.add_format(stp_config.CONST.ITEM_HEADER_FORMAT)
	format_text_lock = workbook.add_format(stp_config.CONST.FORMAT_TEXT_LOCK)
	format_text_lock_hidden = workbook.add_format(stp_config.CONST.FORMAT_TEXT_LOCK_HIDDEN)
	format_num_lock = workbook.add_format(stp_config.CONST.FORMAT_NUM_LOCK)


	#set column width
	right_most_idx = 'T'
	title= ["New or Updated Information", "RIN", "Municipality", "Main Road", "Between Road 1", "Between Road 2", "Roadside", "Location Notes", "Deciduous with Watering Bags", "Conifers with Blue Flagging", "Other Items with Blue Flagging", "Total Items", "Date Watered", "24 Hr Time Watered", "Truck ID", "Total Items Reported", "Total Items Confirmed", "Comments"]
	worksheet.write('A1', "SEQ_ID", format_text_lock_hidden)
	worksheet.write_row('B1', title, format_text)

	col_idx = ['A', 'B', 'C',  'D',   'E',   'F',   'G',   'H',   'I',   'J', 'K',    'L', 'M',   'N',  'O',  'P',  'Q',   'R', 'S']
	col_wid = [0, 16.89, 8.33, 18.33, 24.56, 31.89, 31.89, 10.89, 20.22, 11.22, 12.11, 12.11,7.89, 9.67, 9.89, 8.33, 9.33, 10.33, 10.33]
	for i in range (0,ord(right_most_idx)-65):
	#for i in range (0,19):

		worksheet.set_column(chr(i+65)+':'+chr(i+65), col_wid[i])
	#worksheet.set_column('J:J', 9.65)

	#set row
	worksheet.set_row(0,45)
	worksheet.set_row(1,36)
	worksheet.set_row(5,23.4)
	worksheet.set_row(6, 31.2)

	


	cr = 2
	tag_list  = ["seq_id", "NEW_OR_UPDATED_RIN", "rin", "municipality", "main_road", "between_1", "between_2", "road_side", "LOCATION_NOTES", "broadleaved_(gator_bags)", "conifers", "other_trees", "total_items", "date_watered", "time_watered_(24hr_clock)", "truck_id", "water_count_(total_watered)", "yr_audit_water_count_confirmed", "COMMENTS"]
	for idx, val in enumerate(data["items"]):
		"""
		for i in range (0,ord(right_most_idx)-65):
			a = data["items"][idx][tag_list[i]] if "seq_id" in data["items"][idx].keys() else ""
			worksheet.write('A1', a if a is not None else "", format_text)
		cr += 1
		"""
		a1 = data["items"][idx]["seq_id"]  if "seq_id" in data["items"][idx].keys() else ""
		worksheet.write('A' + str(cr), a1 if a1 is not None else "", format_text_lock_hidden )

		a2 = str(data["items"][idx]["new_or_updated_rin"]) if "new_or_updated_rin" in data["items"][idx].keys() else ""
		if a2 == '1':
			worksheet.write('B' + str(cr), "New" if a2 is not None else "", format_text)
		elif a2 == '2':
			worksheet.write('B' + str(cr), "Updated" if a2 is not None else "", format_text)
		else:
			worksheet.write('B' + str(cr), "", format_text)
		
		a3 = data["items"][idx]["rin"] if "rin" in data["items"][idx].keys() else ""
		worksheet.write('C' + str(cr), a3 if a3 is not None else "", format_text_lock)
		
		a4 = data["items"][idx]["municipality"] if "municipality" in data["items"][idx].keys() else ""
		worksheet.write('D' + str(cr), a4 if a4 is not None else "", format_text_lock)
		
		a5 = data["items"][idx]["main_road"] if "main_road" in data["items"][idx].keys() else ""
		worksheet.write('E' + str(cr), a5 if a5 is not None else "", format_text_lock)


		a6 = data["items"][idx]["between_1"] if "between_1" in data["items"][idx].keys() else ""
		worksheet.write('F' + str(cr), a6 if a6 is not None else "", format_text_lock)

		a7 = data["items"][idx]["between_2"]  if "between_2" in data["items"][idx].keys() else ""
		worksheet.write('G' + str(cr), a7 if a7 is not None else "", format_text_lock)

		a8 = data["items"][idx]["road_side"] if "road_side" in data["items"][idx].keys() else ""
		worksheet.write('H' + str(cr), a8 if a8 is not None else "", format_text_lock)
		
		a9 = data["items"][idx]["LOCATION_NOTES"] if "LOCATION_NOTES" in data["items"][idx].keys() else ""
		worksheet.write('I' + str(cr), a9 if a9 is not None else "", format_text_lock)
		
		a10 = data["items"][idx]["broadleaved_(gator_bags)"] if "broadleaved_(gator_bags)" in data["items"][idx].keys() else ""
		worksheet.write('J' + str(cr), a10 if a10 is not None else "", format_num_lock)
		
		a11 = data["items"][idx]["conifers"] if "conifers" in data["items"][idx].keys() else ""
		worksheet.write('K' + str(cr), a11 if a11 is not None else "", format_num_lock)
		
		a12 = data["items"][idx]["other_trees"] if "other_trees" in data["items"][idx].keys() else ""
		worksheet.write('L' + str(cr), a12 if a12 is not None else "", format_num_lock)

		a13 = data["items"][idx]["total_items"] if "total_items" in data["items"][idx].keys() else ""
		worksheet.write('M' + str(cr), a13 if a13 is not None else "", format_num_lock)
		
		a14 = data["items"][idx]["date_watered"] if "date_watered" in data["items"][idx].keys() else ""
		worksheet.write('N' + str(cr), a14 if a14 is not None else "", format_text)
		
		a15 = data["items"][idx]["time_watered_(24hr_clock)"] if "time_watered_(24hr_clock)" in data["items"][idx].keys() else ""
		worksheet.write('O' + str(cr), a15 if a15 is not None else "", format_text)
		
		a16 = data["items"][idx]["truck_id"] if "truck_id" in data["items"][idx].keys() else ""
		worksheet.write('P' + str(cr), a16 if a16 is not None else "", format_text)
		
		a17 = data["items"][idx]["water_count_(total_watered)"] if "water_count_(total_watered)" in data["items"][idx].keys() else ""
		worksheet.write('Q' + str(cr), a17 if a17 is not None else "", format_num)

		a18 = data["items"][idx]["yr_audit_water_count_confirmed"] if "yr_audit_water_count_confirmed" in data["items"][idx].keys() else ""
		worksheet.write('R' + str(cr), a18 if a18 is not None else "", format_num)
		
		a19 = data["items"][idx]["comments"] if "comments" in data["items"][idx].keys() else ""
		worksheet.write('S' + str(cr), a19 if a19 is not None else "", format_text)
		
		cr += 1


	cr += 4

	worksheet.protect(options = {
	    'objects':               False,
	    'scenarios':             False,
	    'format_cells':          True,
	    'format_columns':        True,
	    'format_rows':           True,
	    'insert_columns':        False,
	    'insert_rows':           False,
	    'insert_hyperlinks':     False,
	    'delete_columns':        False,
	    'delete_rows':           False,
	    'select_locked_cells':   True,
	    'sort':                  False,
	    'autofilter':            False,
	    'pivot_tables':          False,
	    'select_unlocked_cells': True,
		})

	#====ending=======

	workbook.close()

	xlsx_data = output.getvalue()
	return xlsx_data