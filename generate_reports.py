from operator import attrgetter

from docx import Document
from openpyxl import Workbook
from pony.orm import db_session

from boomer_utils import serialize_yes_or_no, adjust_column_width
from models import Event, Participant, School

MASTER_REPORT = "Master.Report.xlsx"
JUDGE_REPORT = "Judge.Report.docx"


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

    workbook.save(MASTER_REPORT)
    print("Saved", MASTER_REPORT)


@db_session
def generate_judge_report():
    events = Event.select().order_by(Event.name)

    events_report = Document()
    for event in events:
        print("Writing", event.name)
        events_report.add_heading(event.name, 0)
        table = events_report.add_table(rows=1, cols=2)
        if event.max_groups > 0:
            create_group_table(event, table)
        else:
            create_individual_table(event, table)
        table.style = 'Table Grid'
        events_report.add_page_break()

    events_report.save(JUDGE_REPORT)
    print("Saved", JUDGE_REPORT)


def create_group_table(event, table):
    header_cells = table.rows[0].cells
    header_cells[0].text = "School"
    header_cells[1].text = "Participants"
    registrations = event.registrations.order_by(lambda r: r.school.name)
    for registration in registrations:
        row_cells = table.add_row().cells
        school = list(registration.participants)[0].school
        row_cells[0].text = school.name
        participants = registration.participants.order_by(Participant.name)
        row_cells[1].text = '\n'.join(p.name for p in participants)


def create_individual_table(event, table):
    header_cells = table.rows[0].cells
    header_cells[0].text = "Participant"
    header_cells[1].text = "School"
    participants = sorted(event.registrations.participants, key=attrgetter('name'))
    for participant in participants:
        row_cells = table.add_row().cells
        row_cells[0].text = participant.name
        row_cells[1].text = participant.school.name
