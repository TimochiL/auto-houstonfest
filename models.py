from pony.orm import Database, PrimaryKey, Required, Set

db = Database()
db.bind(provider='sqlite', filename=':memory:', create_db=True)


class School(db.Entity):
    id = PrimaryKey(int, auto=True)
    name = Required(str)
    participants = Set('Participant')
    registrations = Set('Registration')
    regular_registrations = Required(int)
    late_registrations = Required(int)
    total_enrolled = Required(int)
    rookie_teacher = Required(bool)
    rookie_school = Required(bool)
    attending_state = Required(bool)

    def __repr__(self):
        return F"<School(id={self.id}, name='{self.name}', participants={self.participants}, registrations={self.registrations})>"


class Participant(db.Entity):
    id = PrimaryKey(int, auto=True)
    name = Required(str)
    school = Required(School)
    registrations = Set('Registration')

    def __repr__(self):
        return F"<Participant(id={self.id}, name='{self.name}', school='{self.school.name}', registrations={self.registrations})>"


class Event(db.Entity):
    id = PrimaryKey(int, auto=True)
    name = Required(str)
    max_participants = Required(int)
    max_groups = Required(int)  # 0 = Individual
    registrations = Set('Registration')

    def __repr__(self):
        return F"<Event(id={self.id}, name='{self.name}', max_participants={self.max_participants}, max_groups={self.max_groups}, registrations={self.registrations})>"


class Registration(db.Entity):
    id = PrimaryKey(int, auto=True)
    event = Required(Event)
    school = Required(School)
    participants = Set(Participant)

    def __repr__(self):
        return F"<Registration(id={self.id}, event='{self.event.name}', school='{self.school.name}', participants={[p.name for p in self.participants]})>"


db.generate_mapping(create_tables=True)
