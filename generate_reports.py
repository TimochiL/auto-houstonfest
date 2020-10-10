from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from pony.orm import db_session

from boomer_utils import serialize_yes_or_no, adjust_cell_sizes, adjust_cell_sizes_for_judge_feedback
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
    adjust_cell_sizes(event_worksheet)

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
    adjust_cell_sizes(school_worksheet)

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
        "School",
        "Score on Crit. 1, Judge 1",
        "Score on Crit. 2, Judge 1",
        "Score on Crit. 3, Judge 1",
        "Score on Crit. 4, Judge 1",
        "Score on Crit. 5, Judge 1",
        "Total, Judge 1",
        "Comments, Judge 1",
        "Score on Crit. 1, Judge 2",
        "Score on Crit. 2, Judge 2",
        "Score on Crit. 3, Judge 3",
        "Score on Crit. 4, Judge 4",
        "Score on Crit. 5, Judge 5",
        "Total, Judge 2",
        "Comments, Judge 2",
        "Score on Crit. 1, Judge 3",
        "Score on Crit. 2, Judge 3",
        "Score on Crit. 3, Judge 3",
        "Score on Crit. 4, Judge 3",
        "Score on Crit. 5, Judge 3",
        "Total, Judge 3",
        "Comments, Judge 3",
        "One-time deduction, if any (enter as a positive number)",
        "Total Score"
    ])

    # Only allow point values from zero to twenty
    point_validator = DataValidation(type='whole', operator='between', formula1=0, formula2=20)
    worksheet.add_data_validation(point_validator)

    for row, registration in enumerate(event.registrations, 2):
        worksheet.append([
            '\n'.join(p.name for p in registration.participants),
            registration.school.name,
            '', '', '', '', '',
            f"=SUM(C{row}:G{row})",
            '', '', '', '', '', '',
            f"=SUM(J{row}:N{row})",
            '', '', '', '', '', '',
            f"=SUM(Q{row}:U{row})",
            '', '',
            f"=H{row} + O{row} + V{row} - X{row}"
        ])
        point_validator.add(f"C{row}:G{row}")
        point_validator.add(f"J{row}:N{row}")
        point_validator.add(f"Q{row}:U{row}")
    adjust_cell_sizes_for_judge_feedback(worksheet)

    event_sheet = F"output/Event.{event.name}.xlsx"
    workbook.save(event_sheet)
