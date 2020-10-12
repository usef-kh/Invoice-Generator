import sqlite3


class Sheet:

    def __init__(self, sheet='', db='Database.db'):
        self.conn = sqlite3.connect(db)
        self.cursor = self.conn.cursor()

        self.cursor.execute("select * from " + sheet)
        self.columns = [description[0] for description in self.cursor.description]
        self.sheet = sheet

    def get_table(self):
        return list(self.cursor.execute("select * from " + self.sheet))

    def insert(self, record):
        assert len(record) == len(self.columns)

        command = "insert into " + self.sheet + " values("
        placements = {}
        for i, info in enumerate(record):
            command += ':a' + str(i) + ','
            placements['a' + str(i)] = info

        command = command[:-1] + ")"

        self.cursor.execute(command, placements)
        self.conn.commit()

    def delete(self, record):
        assert len(record) == len(self.columns)

        command = "delete from " + self.sheet + " where "
        for col, val in zip(self.columns[:-1], record[:-1]):
            command += col + "=\"" + str(val) + "\" AND "
        command += self.columns[-1] + "=\"" + str(record[-1]) + "\""

        self.cursor.execute(command)
        self.conn.commit()


if __name__ == '__main__':
    sheet = Sheet('morespace_spaces')

    for record in list(sheet.get_table()):
        print(record)
