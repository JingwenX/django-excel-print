# -*- coding: utf-8 -*-
import xlsxwriter
import json
from . import reports

#calls the specific report function and returns the excel as a download
class ReportGenerator(object):

    def __init__(self, response):
        self.rid = rid
        self.response = response

    def formExcel(res, rid):
        down = reports.reports.d[rid](res)
        return down



