from openpyxl import Workbook
from pony.orm import db_session

from boomer_utils import serialize_yes_or_no, adjust_column_width
from models import Event, School

MASTER_REPORT = "output/Master.Report.xlsx"
JUDGE_REPORT = "output/Judge.Report.docx"


@db_session
def generate_master_report():
    print("Generating master report")

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


@db_session
def generate_event_sheets():
    events = Event.select().order_by(Event.name)
    for event in events:
        generate_event_sheet(event)


def generate_event_sheet(event):
    print("Generating", event.name, "judge sheet")
    workbook = Workbook()
    worksheet = workbook.active

    worksheet.append([
        "Participant(s)",
        "School"
    ])
    for registration in event.registrations:
        worksheet.append([
            '\n'.join(p.name for p in registration.participants),
            registration.school.name
        ])
    adjust_column_width(worksheet)

    event_sheet = F"output/Event.{event.name}.xlsx"
    workbook.save(event_sheet)
