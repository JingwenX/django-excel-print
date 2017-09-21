# -*- coding: utf-8 -*-
from django.shortcuts import render
from django.template import loader
from django.urls import reverse
from django.http import HttpResponse
from django.http import JsonResponse
import requests, io, json, time
from . import rgen
from . import tpe_config

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
		'asPDF': 0
	}

	for p in params:
		if p in request.GET:
			params[p] = request.GET[p]


	s = requests.Session()
	#print(rgen.ReportGenerator.get_url(params))
	if not isinstance(rgen.ReportGenerator.get_url(params), dict):
		response = s.get(rgen.ReportGenerator.get_url(params))

		it = json.loads(response.content)
		content = json.loads(response.content)
		
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
	# TODO: Convert config into json

	if(params['asPDF'] == '1'):
		file = HttpResponse(rgen.ReportGenerator.formExcel(content, params), content_type='application/pdf')
		file['Content-Disposition'] = 'attachment; filename=' + rgen.r_dict[params["rid"]][1] + '.pdf' #'.xlsx'
	else:
		file = HttpResponse(rgen.ReportGenerator.formExcel(content, params), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
		file['Content-Disposition'] = 'attachment; filename=' + rgen.r_dict[params["rid"]][1] + '.xlsx' #'.xlsx'

	s.close()
	return file 

