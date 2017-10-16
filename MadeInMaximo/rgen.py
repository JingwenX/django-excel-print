# -*- coding: utf-8 -*-
import xlsxwriter
import json
#from . import report_classes
from .report_classes import *

r_dict = {
        "1":   [mim_formatted, 'newFormat']
	}

#calls the specific report function and returns the excel as a download
class ReportGenerator(object):

    def __init__(self, response):
        self.rid = rid
        self.response = response

    def formExcel(params):
    	#down = reports.reports.d[rid](res, rid, year, con_num, asgn_num)
    	down = r_dict[params["rid"]][0].render(params)
    	return down


    



