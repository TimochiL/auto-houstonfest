from docx import Document
from pony.orm import db_session, select

from models import Registration, Event, Participant


@db_session
def generate_names(hfest_workbook):
    registrations = select(r for r in Registration).order_by(Registration.event)
    names_generated = hfest_workbook.create_sheet("Names (Generated)")
    # print("Event | Student | School")
    names_generated.cell(1, 1, "Event")
    names_generated.cell(1, 2, "Student")
    names_generated.cell(1, 3, "School")
    current_row = 2
    for registration in registrations:
        for participant in registration.participants:
            # print(registration.event.name, '|', participant.name, '|', participant.school.name)
            names_generated.cell(current_row, 1, registration.event.name)
            names_generated.cell(current_row, 2, participant.name)
            names_generated.cell(current_row, 3, participant.school.name)
            current_row += 1
    # table = Table(displayName="NamesTable", ref=)  # or use auto filters?
    hfest_workbook.save(filename="Hfest.Registr.19.Generated.xlsx")


@db_session
def generate_judge_sheet():
    events = select(e for e in Event).order_by(Event.name)
    events_report = Document()
    for event in events:
        events_report.add_heading(event.name, 0)
        table = events_report.add_table(rows=1, cols=2)
        header_cells = table.rows[0].cells
        if event.is_group:
            header_cells[0].text = "School"
            header_cells[1].text = "Participants"
            for registration in event.registrations:
                row_cells = table.add_row().cells
                school = list((p for p in registration.participants))[0].school
                row_cells[0].text = school.name
                participants = (p.name for p in registration.participants)
                row_cells[1].text = '\n'.join(participants)
        else:
            header_cells[0].text = "Participant"
            header_cells[1].text = "School"
            for registration in event.registrations:
                for participant in registration.participants:
                    row_cells = table.add_row().cells
                    row_cells[0].text = participant.name
                    school = list((p for p in registration.participants))[0].school
                    row_cells[1].text = school.name
        events_report.add_page_break()

    events_report.save("Judge.Report.docx")
