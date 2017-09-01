# -*- coding: utf-8 -*
#rid 4 STP
import xlsxwriter
from io import BytesIO
import datetime
from .. import stp_config

class Report(object):

	def form_url(params):
		base_url = 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_top_performer/'
		base_url += str(params["year"])
		return base_url


	#Top Perforemers
	def render(res, params):

		rid = params["rid"]
		year = params["year"]
		con_num = params["con_num"]
		assign_num = params["assign_num"]

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