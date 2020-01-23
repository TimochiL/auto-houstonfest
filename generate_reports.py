from docx import Document
from openpyxl import Workbook
from pony.orm import db_session

from boomer_utils import serialize_yes_or_no, adjust_column_width
from models import Event, Participant, School


@db_session
def generate_master_report():
    events = Event.select().order_by(Event.name)
    schools = School.select().order_by(School.name)
    workbook = Workbook()
    event_worksheet = workbook.active
    event_worksheet.title = 'Events'
    for event in events:
        event_worksheet.append([event.name, len(event.registrations)])
    adjust_column_width(event_worksheet)

    school_worksheet = workbook.create_sheet('Schools')
    school_worksheet.append([
        "School",
        "Regular Registrations",
        "Late Registrations",
        "Registration Fees",
        "Total Enrolled",
        "Rookie Teacher",
        "Rookie School",
        "Attending State"
    ])
    for school in schools:
        school_worksheet.append([
            school.name,
            school.regular_registrations,
            school.late_registrations,
            school.regular_registrations * 10 + school.late_registrations * 12,
            school.total_enrolled,
            serialize_yes_or_no(school.rookie_teacher),
            serialize_yes_or_no(school.rookie_school),
            serialize_yes_or_no(school.attending_state)
        ])
    adjust_column_width(school_worksheet)

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
