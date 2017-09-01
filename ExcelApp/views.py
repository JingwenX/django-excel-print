# -*- coding: utf-8 -*-
from django.shortcuts import render
from django.template import loader
from django.urls import reverse
from django.http import HttpResponse
from django.http import JsonResponse
import requests, io, json, time
from . import rgen
from . import stp_config

#Dictionary mapping report id to restful API
d = {'2' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_contract_item_summary/',
	 '3' : 'http://ykr-apexp1/ords/stp_contract_detail/', #species summary
	 '4' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_top_performer/', #top performers
	 '6' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_costing_bid/',
	 '7' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_costing_bid/',
	 '8' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_tree_planting/',
	 '9' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_tree_planting/',
	 '17' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_warranty_1yr/',
	 '18' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_warranty_2yr/',
	 '19' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_warranty_12mo/',
	 '51' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_contractor_plant_tree/',
	 '52' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_nusery_inspection/',
	 '53' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_nursery_requirement/',
	 '54' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_contract_item_summary/',
	 '55' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_contract_item_summary/',
	 '56' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_contract_summary/',
	 '57' : 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_contract_summary/',
	 '101': 'http://ykr-apexp1/ords/bsmart_data/bsmart_data/stp_ws/stp_contract_item_detail/'}


file_name =	{'2' : 'Contract Item Summary - by Area Forester',
	 '3' : 'Species Summary', 
	 '4' : 'Top Performers', 
	 '6' : 'Costing Summary',
	 '7' : 'Bid Form Summary',
	 '8' : 'Tree Planting Details',
	 '9' : 'Tree Planting Summary',
	 '17' : 'Warranty Report Species Analysis - 1 year warranty',
	 '18' : 'Warranty Report Species Analysis - 2 year warranty',
	 '19' : 'Warranty Report Species Analysis - 12 months warranty',
	 '51' : 'Contractor Plant Trees',
	 '52' : 'Nursery Inspection Report',
	 '53' : 'Nursery Tagging Requirement Report',
	 '54' : 'Contract Item Summary - All Items',
	 '55' : 'Contract Item Summary - by Area Forester',
	 '56' : 'Contract Item Summary - by Program',
	 '57' : 'Contract Item Summary - by Municipality',
	 '101': 'Contract Item Detail'}


#index page for gui
def index(request):
	return render(request, 'ExcelApp/index.html')

#main page for filling the form
def details(request):
	return render(request, 'ExcelApp/main.html')


#returns json for testing
def getReport(request):

	params = {
		'rid':-1,
		'year':-1,
		'con_num':-1,
		'assign_num':-1,
		'item_num':-1,
		'wtype': -1,
	}

	for p in params:
		if p in request.GET:
			params[p] = request.GET[p]

	#if rid in d:
	#	url = d[rid]
	#	#add year
	#	if not rid in ('8', '9'):
	#		if year != -1:
	#			url += str(year)
	#		if item_num != -1:
	#			url += str(item_num)
	#	else:
	#		url += '{}/{}'.format(year, item_num)

		#add
	s = requests.Session()
	#print(rgen.ReportGenerator.get_url(params))
	response = s.get(rgen.ReportGenerator.get_url(params))

	it = json.loads(response.content)
	#test
	#return HttpResponse(response.content)
	content = json.loads(response.content)
	
	pageNum = 1
	while "next" in it:
		response = s.get(rgen.ReportGenerator.get_url(params) + '?page=' + str(pageNum))
		it = json.loads(response.content)
		content["items"].extend(it["items"])
		pageNum += 1

	# TODO: Convert config into json
	file = HttpResponse(rgen.ReportGenerator.formExcel(content, params), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
	file['Content-Disposition'] = 'attachment; filename=' + rgen.r_dict[params["rid"]][1] + '.xlsx'

	return file 

