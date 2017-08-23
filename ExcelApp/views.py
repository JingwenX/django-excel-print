# -*- coding: utf-8 -*-
from django.shortcuts import render
from django.template import loader
from django.urls import reverse
from django.http import HttpResponse
from django.http import JsonResponse
import requests, io, json, time
from . import rgen

#Dictionary mapping report id to restful API
d = {'2' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_pop_test/',
	 '6' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_costing_bid/2014',
	 '7' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_costing_bid/2014',
	 '17' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_warranty_1yr/',
	 '18' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_warranty_2yr/',
	 '19' : 'http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/bsmart_data/stp_ws/stp_warranty_12mo/'}

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

	#if the id is in the dictionary
	if rid in d:
		
		response = requests.get(d[rid])

		it = json.loads(response.content)
		content = json.loads(response.content)

		pageNum = 1
		while "next" in it:
			response = requests.get(d[rid]+ '?page=' + str(pageNum))
			it = json.loads(response.content)
			content["items"].extend(it["items"])
			pageNum += 1

		file = HttpResponse(rgen.ReportGenerator.formExcel(content, rid), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
		file['Content-Disposition'] = 'attachment; filename=test.xlsx'

		end = time.time()
		print('Time Elapsed: ' + str(end - start))

		return file 
	else:
		return HttpResponse('No API in Dictionary')
