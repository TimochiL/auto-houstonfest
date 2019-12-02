from openpyxl.worksheet.worksheet import Worksheet
from pony.orm import db_session, select

from models import Registration


@db_session
def generate_names(hfest_workbook):
    registrations = select(r for r in Registration).order_by(Registration.event.name)
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
    hfest_workbook.save(filename="Hfest.Registr.19_Generated.xlsx")

