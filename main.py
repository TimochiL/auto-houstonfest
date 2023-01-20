import glob
import re
from pathlib import Path

import openpyxl
from pony.orm import db_session

from boomer_utils import parse_yes_or_no
from generate_reports import generate_master_report, generate_event_sheet, generate_judge_report
from models import School, Participant, Event, Registration

EVENT_START_ROW = 25  # Where school information ends and event listings start


@db_session
def main():
    registration_files = glob.glob("Reg.*.xlsx")
    if not registration_files:  # Exit if no registration files found
        print("NO REGISTRATION FILES FOUND")
        return
    enable_event_sheets = input("GENERATE EVENT SHEETS? (YES/NO) ")
    print(F"FOUND {len(registration_files)} REGISTRATION FILE(S)")
    event_count = import_events(registration_files[0])  # Import event listings from the first workbook
    print(F"IMPORTED {event_count} EVENTS")

    schools = []
    for workbook_file in registration_files:
        workbook = openpyxl.load_workbook(workbook_file)
        worksheet = workbook['Original']
        school_name = worksheet.cell(4, 2).value
        print("Processing", school_name, "registration")

        school = School(
            name=school_name,
            regular_registrations=worksheet.cell(16, 2).value or 0,  # Default to zero if not specified
            late_registrations=worksheet.cell(17, 2).value or 0,
            total_enrolled=worksheet.cell(19, 2).value,
            rookie_teacher=parse_yes_or_no(worksheet.cell(20, 2).value or 'no'),  # Default to no if not specified
            rookie_school=parse_yes_or_no(worksheet.cell(21, 2).value or 'no'),
            attending_state=parse_yes_or_no(worksheet.cell(22, 2).value or 'no')
        )
        schools.append(school)

        event_row = EVENT_START_ROW
        for event in Event.select():
            for _ in range(max(event.max_groups, 1)):
                participant_row = 0
                participants = []
                while participant_row < event.max_participants:
                    participant_name = worksheet.cell(event_row + participant_row, 2).value
                    if participant_name is not None and participant_name.strip():  # If a non-whitespace entry exists
                        participant = find_or_create_participant(participant_name, school)
                        participants.append(participant)
                    participant_row += 1
                if len(participants) > 0:  # Only create a registration if there are more than zero participants
                    if event.max_groups == 0:  # Individual events have a max_groups of 0
                        for participant in participants:  # Register each participant separately for individual events
                            Registration(event=event, school=school, participants=participant)
                    else:  # Register participants together for group events
                        Registration(event=event, school=school, participants=participants)
                event_row += participant_row

    Path('output').mkdir(exist_ok=True)
    events = Event.select().order_by(Event.name)
    schools = School.select().order_by(School.name)
    generate_master_report(events, schools)
    generate_judge_report(events)
    if parse_yes_or_no(enable_event_sheets):
        Path('output/events').mkdir(exist_ok=True)
        for event in events:
            generate_event_sheet(event)

    print()
    print("SCRIPT WRITTEN BY DAMIAN LALL, CHS '21")
    print("TASK COMPLETE, PRESS ENTER TO EXIT", end='')
    input()


def import_events(workbook_file) -> int:
    workbook = openpyxl.load_workbook(workbook_file)
    worksheet = workbook['Original']
    participant_count = 0
    previous_event = None

    for row in worksheet.iter_rows(EVENT_START_ROW):
        row_event = row[0].value
        if previous_event is not None and row_event != previous_event:
            create_event(previous_event, participant_count)
            participant_count = 0
        participant_count += 1
        previous_event = row_event
    create_event(previous_event, participant_count)

    return Event.select().count()


def create_event(event_name, participant_count):
    is_group = "Group" in event_name
    event_name = re.sub(R"[(\[].*?[)\]]", "", event_name).strip()  # Remove anything between parenthesis
    event = Event.get(name=event_name)
    if event is not None:
        event.max_groups += 1
    else:
        Event(name=event_name, max_participants=participant_count, max_groups=is_group)


def find_or_create_participant(name, school) -> Participant:
    participant = Participant.get(name=name, school=school)
    if participant is None:
        return Participant(name=name, school=school)
    return participant


if __name__ == '__main__':
    main()
