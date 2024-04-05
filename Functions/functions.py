from fpdf import FPDF
from openpyxl import load_workbook as wb
from openpyxl import Workbook as wrk
from openpyxl.styles import Font, Alignment, Border, Side


class functions:
    times = []
    emt_days = {
        "Sat": {},
        "Sun": {},
        "Mon": {},
        "Tue": {},
        "Wed": {},
        "Thu": {}
    }

    days = {
        "Sat": [],
        "Sun": [],
        "Mon": [],
        "Tue": [],
        "Wed": [],
        "Thu": []
    }

    dummy = {}

    # Extracts color of each semester information
    def extractor(loc):
        pass

    # This generates individual section routine
    def routine_separateor(loc, bach):
        pass

    # This generates teacher routine with teacher initial
    def teacher(loc, ti):
        pass

    # This genarates empty slot info
    def empty_slot(loc):
        pass

    # Creates routine of blank slots as pdf file
    def blank_pdf(dic, times, path):
        pass

    # Creates routine as pdf file
    def routine_pdf( dic, times, sec, path):
        pass
            

    # Creates routine as excel file
    def routine_excel(dic, times, sec, path):
        pass
