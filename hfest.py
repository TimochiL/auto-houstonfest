import csv
from typing import List

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from pony.orm import db_session

from generate_reports import generate_names, generate_judge_sheet
from models import School, Participant, Event, Registration


@db_session
def main():
    events = import_events()
    hfest_workbook = openpyxl.load_workbook("Hfest.Registr.19.xlsx")  # TODO: Take in as arg
    hfest_registration: Worksheet = hfest_workbook['Registration']
    school_count = hfest_registration.max_column
    schools = list()
    school_column = 3
    while school_column < school_count:
        school = School(name=hfest_registration.cell(1, school_column).value)
        schools.append(school)

        current_row = 32
        for event in events:
            # print("curr row is", current_row)
            # print("curr event is", event.name)
            current_row += 1  # TODO: Validate if enter 1/number of entrants is consistent
            current_participant = 0
            participants = list()
            while current_participant < event.max_participants:
                participant_name = hfest_registration.cell(current_row + current_participant, school_column).value
                # print("got", participant_name, "at", current_row + current_participant, ",", school_column)
                if participant_name is not None:
                    participant = find_or_create_participant(participant_name, school)
                    participants.append(participant)
                    # print("added", participant)
                current_participant += 1
                # print("inc part to", current_participant)
            if len(participants) > 0:
                Registration(event=event, participants=participants)
            current_row += current_participant
        school_column += 1

    # for r in select(r for r in Registration if r.event.name == "Skit 1"):
    #     print(r)
    generate_names(hfest_workbook)
    generate_judge_sheet()


def import_events() -> List[Event]:
    events_reader = csv.reader(open("events.csv"))
    events = list()
    for row in list(events_reader)[1:]:
        if row[2] == 'TRUE':
            events.append(Event(name=row[0], max_participants=int(row[1]), is_group=True))
        else:
            events.append(Event(name=row[0], max_participants=int(row[1]), is_group=False))
    return events


def find_or_create_participant(name, school) -> Participant:
    participant = Participant.get(name=name, school=school)
    if participant is None:
        return Participant(name=name, school=school)
    else:
        return participant


if __name__ == '__main__':
    main()
