import xlsxwriter
import json
from . import reports

#calls the specific report function and returns the excel as a download
class report(object):

    def __init__(self, response):
        self.rid = rid
        self.response = response

    def formExcel(res, rid):
        reports.reports.d[rid](res)



