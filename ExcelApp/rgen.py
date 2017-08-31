# -*- coding: utf-8 -*-
import xlsxwriter
import json
#from . import report_classes
from .report_classes import *

r_dict = {
		"3": [r_test.Report, 'Species Summary'],
    	"4" : [stp_tp.Report, 'Top Performers'], 
    	"6" : [stp_cs.Report, 'Costing Summary'],
   		"7" : [stp_bfs.Report, 'Bid Form Summary'],
    	"8" : [stp_tpd.Report, 'Tree Planting Details'],
    	"9" : [stp_tps.Report, 'Tree Planting Summary'],
    	"17" : [stp_wrsa1.Report, 'Warranty Report Species Analysis - 1 year warranty'],
    	"18" : [stp_wrsa2.Report, 'Warranty Report Species Analysis - 2 year warranty'],
    	"19" : [stp_wrsa12.Report, 'Warranty Report Species Analysis - 12 months warranty'],
    	"51" : [stp_cpt.Report, 'Contractor Plant Trees'],
    	"52" : [stp_nir.Report, 'Nursery Inspection Report'],
    	"53" : [stp_ntrr.Report, 'Nursery Tagging Requirement Report'],
		"54": [stp_by_ai.Report, 'Contract Item Summary - by All Items'],
    	"55" : [stp_cis_af.Report, 'Contract Item Summary - by Area Forester'],
    	"56" : [stp_cis_p.Report, 'Contract Item Summary - by Program'],
    	"57" : [stp_cis_m.Report, 'Contract Item Summary - by Municipality'],
    	"101": [stp_cid.Report, 'Contract Item Detail'],
	}

#calls the specific report function and returns the excel as a download
class ReportGenerator(object):

    def __init__(self, response):
        self.rid = rid
        self.response = response

    def get_url(params):
    	return r_dict[params["rid"]][0].form_url(params)


    def formExcel(res, params):
    	#down = reports.reports.d[rid](res, rid, year, con_num, asgn_num)
    	down = r_dict[params["rid"]][0].render(res, params)
    	return down


    



