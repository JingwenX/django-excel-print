# -*- coding: utf-8 -*-
from django.shortcuts import render
from django.template import loader
from django.urls import reverse
from django.http import HttpResponse
from django.http import JsonResponse
import requests, io, json, time
from . import rgen
from . import stp_config

def index(request):
	"""
	Index page for gui - only used in dev
	"""
	return render(request, 'ExcelApp/index.html')

def details(request):
	"""
	Main page for gui - only used in dev
	"""
	return render(request, 'ExcelApp/main.html')

def getReport(request):
	"""
	Gets the report from a specific report class and returns a downloadable file
	"""

	#parameters needed for different REST API's
	params = {
		'rid':-1,
		'year':-1,
		'con_num':-1,
		'assign_num':-1,
		'item_num':-1,
		'wtype': -1,
		'payno': -1,
		'snap': 0, #default is 0 for snapshots (for now)
		'issue_date': -1,
	}

	#loop over the parameters and set them if they appear in the api url
	for p in params:
		if p in request.GET:
			params[p] = request.GET[p]


	#get the request session and load data
	s = requests.Session()
	if not isinstance(rgen.ReportGenerator.get_url(params), dict):
		response = s.get(rgen.ReportGenerator.get_url(params))

		#set the iterator and the content
		it = json.loads(response.content)
		content = json.loads(response.content)
		
		#while a next page exists, parse the api
		pageNum = 1
		while "next" in it:
			response = s.get(rgen.ReportGenerator.get_url(params) + '?page=' + str(pageNum))
			it = json.loads(response.content)
			content["items"].extend(it["items"])
			pageNum += 1

	else:
		#if the url is a list
		content = {}
		for part in rgen.ReportGenerator.get_url(params):
			response = s.get(rgen.ReportGenerator.get_url(params)[part])
			it = json.loads(response.content)
			#content = {"part1":{"items":[]}, "part2":{"items":[]}, "part3":{"items":[]}}
			
			content[part] = {}
			content[part]["items"] = []
			content[part]["items"].extend(it["items"])

			pageNum = 1
			while "next" in it:
				response = s.get(rgen.ReportGenerator.get_url(params)[part] + '?page=' + str(pageNum))
				it = json.loads(response.content)
				content[part]["items"].extend(it["items"])
				pageNum += 1
	
	#set the file object to be returned as a download
	file = HttpResponse(rgen.ReportGenerator.formExcel(content, params), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
	if params["rid"] == '70':
		file['Content-Disposition'] = 'attachment; filename=' + rgen.r_dict[params["rid"]][1] + ' No.' + params['issue_date'] + '.xlsx'
	else:
		file['Content-Disposition'] = 'attachment; filename=' + rgen.r_dict[params["rid"]][1]  + '.xlsx'
	s.close()
	return file 

