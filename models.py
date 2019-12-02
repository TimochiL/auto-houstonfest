from pony.orm import *

db = Database()
db.bind(provider='sqlite', filename=':memory:', create_db=True)


class School(db.Entity):
    id = PrimaryKey(int, auto=True)
    participants = Set('Participant')
    name = Required(str)


class Participant(db.Entity):
    id = PrimaryKey(int, auto=True)
    name = Required(str)
    school = Required(School)
    registrations = Set('Registration')


class Event(db.Entity):
    id = PrimaryKey(int, auto=True)
    registrations = Set('Registration')
    name = Required(str)
    max_participants = Required(int)
    is_group = Required(bool)


class Registration(db.Entity):
    id = PrimaryKey(int, auto=True)
    event = Required(Event)
    participants = Set(Participant)


db.generate_mapping(create_tables=True)
