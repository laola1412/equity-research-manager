# create a report using reportlab

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.lib.pagesizes import letter
import os
import csv
from datetime import date
from decorators import timeit

def create_report(save_name, name, close, date, target_value, rating, description):
    styles = getSampleStyleSheet()
    report = SimpleDocTemplate(f"reports/{date}_{save_name.lower()}.pdf")
    report.leftMargin, report.rightMargin = 30, 30
    report.topMargin, report.bottomMargin = 30, 30
    
    report_title = Paragraph(f"{name} Report", styles["h1"])
    report_info = Paragraph(f"Created on {date}. Stock closed at {close}", styles["BodyText"])
    report_rating = Paragraph(f"Rating: {rating} with a target value of {target_value}", styles["BodyText"])
    description = Paragraph(description, styles["BodyText"])
    empty_line = Spacer(1,20)
    
    report.build([report_title, report_info, report_rating, empty_line, description])