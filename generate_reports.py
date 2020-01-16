from docx import Document
from openpyxl import Workbook
from pony.orm import db_session

from models import Event, Participant


@db_session
def generate_master_report():
    events = Event.select().order_by(Event.name)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = 'Report'
    for event in events:
        worksheet.append([event.name, len(event.registrations)])
    workbook.save("Master.Report.xlsx")


@db_session
def generate_judge_sheet():
    events = Event.select().order_by(Event.name)
    events_report = Document()
    for event in events:
        events_report.add_heading(event.name, 0)
        table = events_report.add_table(rows=1, cols=2)
        header_cells = table.rows[0].cells
        if event.max_groups > 0:
            header_cells[0].text = "School"
            header_cells[1].text = "Participants"
            for registration in event.registrations:
                row_cells = table.add_row().cells
                school = list(registration.participants)[0].school
                row_cells[0].text = school.name
                participants = registration.participants.order_by(Participant.name)
                row_cells[1].text = '\n'.join(p.name for p in participants)
        else:
            header_cells[0].text = "Participant"
            header_cells[1].text = "School"
            for registration in event.registrations:
                for participant in registration.participants.order_by(Participant.name):
                    row_cells = table.add_row().cells
                    row_cells[0].text = participant.name
                    school = list(registration.participants)[0].school
                    row_cells[1].text = school.name
        events_report.add_page_break()
    events_report.save("Judge.Report.docx")
