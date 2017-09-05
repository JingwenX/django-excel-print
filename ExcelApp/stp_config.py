import datetime

class const(object):
	# API PREFIX
	API_URL_PREFIX = 'http://ykr-apexp1/ords/'

	# Header and Title format
	
	MAIN_HEADER1_FORMAT = {'bold':True,
		'font_name':'Calibri',
		'font_size':12,
		'border':2, #2 is the value for thick border
		'align':'center',
		'valign':'top',}
	
	MAIN_HEADER2_FORMAT = {'font_name':'Calibri',
		'font_size':18,
		'font_color':'white',
		'border':2,
		'align':'left',
		'bg_color':'black',
			}

	TITLE_FORMAT = {
		'font_name':'Calibri',
		'font_size':18,
		'font_color':'white',
		'border':2,
		'align':'left',
		'bg_color':'gray',
		}
	# Header and Title text

	#Main data format and text
	ITEM_HEADER_FORMAT = {
					'bold':True,
					'font_name':'Calibri',
					'font_size':12,
					'border':2,
					'align': 'center',
					'bg_color':'#C0C0C0',
					'valign': 'vcenter',
					'text_wrap': True,
				}


	ITEM_FORMAT = {
			'font_name':'Calibri',
			'font_size':12,
			'align': 'left',
			}

	FORMAT_TEXT = {'font_name':'Calibri',
		'font_size':12,
		'align': 'left',
		'valign': 'vcenter',
		'text_wrap': True,
		'border': True,
		'border_color':'gray',}

	FORMAT_NUM = {'font_name':'Calibri',
		'font_size':12,
		'align': 'center',
		'valign': 'vcenter',
		'text_wrap': True,
		'border':True,
		'border_color':'gray',}
	FOOTER_PAGE_NUM = '&LPage &P of &N'

	ENV_LOGO = r'\\ykr-apexp1\staticenv\apps\199\env_internal_bw.png'

	##Hunter's format
	SUBTITLE_FORMAT = {
			'font_name':'Calibri',
			'font_size': 14,
			'bold':True,
			'font_color':'black',
			'border':2,
			'align':'left',
			'bg_color':'#C0C0C0',
		}
	ITEM_FORMAT_MONEY = {
			'font_name':'Calibri',
			'font_size':12,
			'align': 'left',
			'num_format': '$#,##0',
			'border' : True,
			'border_color':'gray',
		}
	SUBTOTAL_FORMAT = {
			'font_name':'Calibri',
			'font_size': 14,
			'font_color':'black',
			'align':'left ',
			'border' : True,
			'border_color':'gray',
			#'bg_color':'#D3D3D3',
		}
	SUBTOTAL_FORMAT_MONEY = {
			'font_name':'Calibri',
			'font_size': 14,
			'font_color':'black',
			'align':'left',
			#'bg_color':'#D3D3D3',
			'border' : True,
			'border_color':'gray',
			'num_format': '$#,##0',
		}

	def write_gen_title(title, workbook, worksheet, rightmost_idx, year, con_num):

		MAIN_HEADER1_FORMAT = {'bold':True,
			'font_name':'Calibri',
			'font_size':12,
			'border':2, #2 is the value for thick border
			'align':'center',
			'valign':'top',}
		MAIN_HEADER2_FORMAT = {'font_name':'Calibri',
			'font_size':18,
			'font_color':'white',
			'border':2,
			'align':'left',
			'bg_color':'black',
				}
		TITLE_FORMAT = {
			'font_name':'Calibri',
			'font_size':18,
			'font_color':'white',
			'border':2,
			'align':'left',
			'bg_color':'gray',
			}
		# Header and Title text

		#Main data format and text
		ITEM_HEADER_FORMAT = {
						'bold':True,
						'font_name':'Calibri',
						'font_size':12,
						'border':2,
						'align': 'center',
						'bg_color':'#C0C0C0',
						'valign': 'vcenter',
						'text_wrap': True,
					}


		ITEM_FORMAT = {
				'font_name':'Calibri',
				'font_size':12,
				'align': 'left',
				}

		FORMAT_TEXT = {'font_name':'Calibri',
			'font_size':12,
			'align': 'left',
			'valign': 'vcenter',
			'text_wrap': True,}

		FORMAT_NUM = {'font_name':'Calibri',
			'font_size':12,
			'align': 'center',
			'valign': 'vcenter',
			'text_wrap': True,}
		FOOTER_PAGE_NUM = '&LPage &P of &N'

		# add format
		main_header1_format = workbook.add_format(MAIN_HEADER1_FORMAT)
		main_header2_format = workbook.add_format(MAIN_HEADER2_FORMAT)
		title_format = workbook.add_format(TITLE_FORMAT)
		item_header_format = workbook.add_format(ITEM_HEADER_FORMAT)
		format_text = workbook.add_format(FORMAT_TEXT) #text left align
		format_num = workbook.add_format(FORMAT_NUM) #number center align
		#insert YORKREGION logo
		worksheet.insert_image('A1',r'\\ykr-apexp1\staticenv\York_Logo.png',{'x_offset':10,'y_offset':10,'x_scale':0.25,'y_scale':0.25})

		worksheet.merge_range('A1:'+ rightmost_idx + '2','Natural Heritage and Forestry Division, Environmental Services Department', main_header1_format)
		worksheet.merge_range('A4:' + rightmost_idx + '4', str(con_num) + ' (' + str(year) +') - Street Tree Planting and Establishment Activities', main_header2_format)
		worksheet.merge_range('A5:' + rightmost_idx + '5', title, title_format)
		worksheet.merge_range('A6:' + rightmost_idx + '6',' ')


		#write date printed
		##date format
		date_format = workbook.add_format()
		date_format.set_align('right')
		now = datetime.datetime.now()
		##date data
		date_printed = 'Date Printed: ' + str(now.day) + '-' + str(now.strftime("%b")) + '-' + str(now.year)
		worksheet.write(rightmost_idx +'3', date_printed, date_format)

		#set printer default
		worksheet.set_landscape()
		#worksheet.set_margins({'left':0.25, 'right':0.25, 'top':0.75, 'bottom':0.75})
		worksheet.print_area('A1:H1048576')
		#worksheet.set_v_pagebreaks(6)

		#set footer
		worksheet.set_footer(FOOTER_PAGE_NUM) 

		return
	
CONST = const()