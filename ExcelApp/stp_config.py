"""
def footer_page_num():
	return '&LPage &P of &N'

def format_text():
	return {'font_name':'Calibri',
			'font_size':12,
			'align': 'left',
			'valign': 'vcenter',
			'text_wrap': True}
"""



class const(object):
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
					'align': 'left',
					'bg_color':'#C0C0C0',
					'valign': 'center',
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
	
CONST = const()