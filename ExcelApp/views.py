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
d = {'2' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contract_item_summary/',
	 '3' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contract_detail/', #species summary
	 '4' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_top_performer/', #top performers
	 '6' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_costing_bid/',
	 '7' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_costing_bid/',
	 '8' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_tree_planting/',
	 '9' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_tree_planting/',
	 '17' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_warranty_1yr/',
	 '18' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_warranty_2yr/',
	 '19' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_warranty_12mo/',
	 '51' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contractor_plant_tree/',
	 '52' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_nusery_inspection/',
	 '53' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_nursery_requirement/',
	 '54' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contract_item_summary/',
	 '55' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contract_item_summary/',
	 '56' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contract_summary/',
	 '57' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contract_summary/',
	 '101': 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contract_item_detail/'}


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
	start = time.time()
	#report id

	rid, year, con_num, asgn_num, item_num = -1,-1,-1,-1,-1,
	if 'rid' in request.GET:
		rid = request.GET['rid']

	if 'p_year' in request.GET:
		year = request.GET['p_year']

	if 'con_num' in request.GET:
		con_num = request.GET['con_num']

	if 'asgn_num' in request.GET:
		asgn_num = request.GET['asgn_num']

	if 'item_num' in request.GET:
		item_num = request.GET['item_num']
	


	#if the id is in the dictionary
	if rid in d:
		url = d[rid]
		#add year
		if rid != '8':
			if year != -1:
				url += str(year)
			if item_num != -1:
				url += str(item_num)
		else:
			url += '{}/{}'.format(year, item_num)


		response = requests.get(url)

		it = json.loads(response.content)
		content = json.loads(response.content)
		
		pageNum = 1
		while "next" in it:
			response = requests.get(url + '?page=' + str(pageNum))
			it = json.loads(response.content)
			content["items"].extend(it["items"])
			pageNum += 1


		# TODO: Convert config into json
		file = HttpResponse(rgen.ReportGenerator.formExcel(content, rid, year, con_num, asgn_num), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
		file['Content-Disposition'] = 'attachment; filename=' + file_name[rid] + '.xlsx'


		end = time.time()
		#print('Time Elapsed: ' + str(end - start))

		return file 
	else:
		return HttpResponse('No API in Dictionary')
