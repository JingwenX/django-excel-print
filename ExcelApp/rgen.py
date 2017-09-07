# -*- coding: utf-8 -*-
import xlsxwriter
import json
#from . import report_classes
from .report_classes import *

r_dict = {
		"3":   [stp_ss, 'Species Summary'],
    	"4" :  [stp_tp, 'Top Performers'], 
    	"6" :  [stp_cs, 'Costing Summary'],
   		"7" :  [stp_bfs, 'Bid Form Summary'],
    	"8" :  [stp_tpd, 'Tree Planting Details'],
    	"9" :  [stp_tps, 'Tree Planting Summary'],
    	"17" : [stp_wrsa, 'Warranty Report Species Analysis'],
    	"18" : [stp_wrma, 'Warranty Report Municipality Analysis'],
    	"19" : [stp_wrha, 'Warranty Report Health Analysis'],
    	"20" : [stp_wrca, 'Warranty Report Contract Analysis'],
    	"21" : [stp_wrdd, 'Warranty Report Deficiency List Details'],
    	"22" : [stp_wrds, 'Warranty Report Deficiency List Summary'],
    	"23" : [stp_wrsl, 'Warranty Report Species List'],
    	"24" : [stp_wrrl, 'Warranty Report Replacement List'],
    	"25" : [stp_tppr, 'Tree Planting Payment Report'],
    	"26" : [stp_tpdl, 'Tree Planting Deficiency List'],
    	"51" : [stp_cpt, 'Contractor Plant Trees'],
    	"52" : [stp_nir, 'Nursery Inspection Report'],
    	"53" : [stp_ntrr, 'Nursery Tagging Requirement Report'],
		"54":  [stp_cis_ai, 'Contract Item Summary - by All Items'],
    	"55" : [stp_cis_af, 'Contract Item Summary - by Area Forester'],
    	"56" : [stp_cis_p, 'Contract Item Summary - by Program'],
    	"57" : [stp_cis_m, 'Contract Item Summary - by Municipality'],
        "58" : [stp_niraua, 'Nursery Tagging Requirement - Assigned and Unassiged Species'],
        "70" : [stp_iwa, 'Issued Watering Assignment'],
        "71" : [stp_iwar, 'Issued Watering Assignment Report'],
        "72" : [stp_ewr, 'Extra Work Payment Report'],
        "73" : [stp_wpr, 'Watering Payment Report'],
    	"101": [stp_cid, 'Contract Item Detail'],
        "102": [stp_cid_no_cursor, 'Contract Item Detail'],
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


    



