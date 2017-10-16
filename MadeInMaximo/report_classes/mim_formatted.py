#rid 8 STP
# -*- coding: utf-8 -*
import xlsxwriter
from io import BytesIO
import datetime
import tempfile
import os
import string
import json

#Made in maximo format generator
def render(params):
	output = BytesIO()
	workbook = xlsxwriter.Workbook(output, {'in_memory': True})
	worksheet = workbook.add_worksheet()
	title = 'IM Formatted'

	data = open(os.getcwd() + r'\MadeInMaximo\dataFiles\mim_data.json', 'r')

	jdata = json.loads(data.read())

	worksheet.set_column('A:J', 20)

	cr = 1
	
	for csc_id, class_subclass in enumerate(jdata):
		if '/' in class_subclass:
			currClass = class_subclass.split('/')[0]
			currSubClass = class_subclass.split('/')[1]
			worksheet.write('A{}'.format(cr), currClass, workbook.add_format({'bg_color': '#278a44'}))
			worksheet.write('B{}'.format(cr), currSubClass, workbook.add_format({'bg_color': '#278a44'}))
		else:
			worksheet.write('A{}'.format(cr), class_subclass, workbook.add_format({'bg_color': '#278a44'}))
			worksheet.write('B{}'.format(cr), '', workbook.add_format({'bg_color': '#278a44'}))


		asc_lookup = 67
		maxcount = 0
		for aid, attr in enumerate(jdata[class_subclass]):
			val_count = 0
			worksheet.write('{}{}'.format(chr(asc_lookup), cr), jdata[class_subclass][attr].get('NAME') + '-text-[n/a]', workbook.add_format({'bg_color': '#278a44'}))
			for id, val in enumerate(jdata[class_subclass][attr]['VALUES']):
				worksheet.write('{}{}'.format(chr(asc_lookup), cr + val_count + 1), val)
				val_count += 1

			maxcount = val_count if val_count > maxcount else maxcount
			asc_lookup += 1

		cr += maxcount + 1

	workbook.close()

	return_data = output.getvalue()
	return return_data