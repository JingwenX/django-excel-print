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
d = {'3' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contract_detail/', #species summary
	 '4' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_top_performer/', #top performers
	 '6' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_costing_bid/',
	 '7' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_costing_bid/',
	 '17' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_warranty_1yr/',
	 '18' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_warranty_2yr/',
	 '19' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_warranty_12mo/',
	 '51' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contractor_plant_tree/',
	 '52' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_nusery_inspection/',
	 '53' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_nursery_requirement/',
	 '54' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contract_item_summary/',
	 '55' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contract_item_summary/',
	 '56' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contract_summary/',
	 '57' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_contract_summary/'}

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
	rid = request.GET['rid']
	year = request.GET['p_year']
	#if the id is in the dictionary
	if rid in d:
		#add year
		url = d[rid] + str(year)
		response = requests.get(url)

		it = json.loads(response.content)
		content = json.loads(response.content)

		pageNum = 1
		while "next" in it:
			response = requests.get(d[rid] + str(year) + '?page=' + str(pageNum))
			it = json.loads(response.content)
			content["items"].extend(it["items"])
			pageNum += 1

		file = HttpResponse(rgen.ReportGenerator.formExcel(content, rid), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
		file['Content-Disposition'] = 'attachment; filename=test.xlsx'


		end = time.time()
		#print('Time Elapsed: ' + str(end - start))

		return file 
	else:
		return HttpResponse('No API in Dictionary')
