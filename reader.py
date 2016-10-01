#!/bin/python3

import re
import os.path
import dateutil.parser
import uuid
import sqlite3
import xlrd
import tkinter
from tkinter import filedialog

# Globals
# PARTCP_DB = 'participations.db'
# EVENTS_DB = 'events.db'
# CLUBS_DB = 'clubs.db'
PARTCP_DB = EVENTS_DB = CLUBS_DB = 'master.db'


def xlread(filename):
    """
    Read roll numbers from excel file
    """
    wbook = xlrd.open_workbook(filename)
    sheet = wbook.sheet_by_index(0)
    rollnos = list()
    regp = re.compile(r'\w{2}\d{2}\w\d{3}', re.MULTILINE)
    for rowindex in range(sheet.nrows):
        st = ' '.join(map(str, sheet.row_values(rowindex)))
        rollmatch = regp.search(st)
        if (rollmatch is not None):
            firstmatch = rollmatch.group(0).lower().strip()
            rollnos.append(firstmatch)
    return rollnos


def init_partcpdb(filename):
    """
    Set up the student database if it doesn't exist
    """
    _command = 'CREATE TABLE IF NOT EXISTS partcp ( indx integer PRIMARY KEY, ' \
        'rollno text NOT NULL, eventindx integer NOT NULL);'
    dbconn = sqlite3.connect(filename)
    cursor = dbconn.cursor()
    cursor.execute(_command)
    dbconn.commit()
    dbconn.close()


def init_eventsdb(filename):
    """
    Set up the event database if it doesn't exist
    """
    _command = 'CREATE TABLE IF NOT EXISTS events ( indx integer PRIMARY KEY, ' \
        'name text NOT NULL, date text NOT NULL, ' \
        'club integer NOT NULL, uuid text NOT NULL);'
    dbconn = sqlite3.connect(filename)
    cursor = dbconn.cursor()
    cursor.execute(_command)
    dbconn.commit()
    dbconn.close()


def init_clubsdb(filename):
    """
    Set up the event database if it doesn't exist
    """
    _command = 'CREATE TABLE IF NOT EXISTS clubs ( indx integer PRIMARY KEY, ' \
        'name text NOT NULL);'
    dbconn = sqlite3.connect(filename)
    cursor = dbconn.cursor()
    cursor.execute(_command)
    dbconn.commit()
    dbconn.close()


def insert_event(_name, _date, _uuid, _club, _dbname):
    """
    Insert event details into database
    """
    dbconn = sqlite3.connect(_dbname)
    cursor = dbconn.cursor()
    cursor.execute('SELECT COALESCE(MAX(indx),0) FROM events')
    eventindx = int(cursor.fetchone()[0]) + 1
    data = (_name, str(_date), _club, str(_uuid))
    cursor.execute(
        'INSERT INTO events (name,date,club,uuid) VALUES (?,?,?,?)', data)
    dbconn.commit()
    dbconn.close()
    return eventindx


def insert_participants(_participants, _eventindx, _dbname):
    """
    Insert all rollnumbers who participated in event
    """
    dbconn = sqlite3.connect(_dbname)
    cursor = dbconn.cursor()
    data = list(map(lambda x: (x, _eventindx), _participants))
    cursor.executemany(
        'INSERT INTO partcp (rollno,eventindx) VALUES (?,?)', data)
    dbconn.commit()
    dbconn.close()


def insert_club(_name, _dbname):
    """
    Insert new club into database
    """
    dbconn = sqlite3.connect(_dbname)
    cursor = dbconn.cursor()
    cursor.execute('SELECT COALESCE(MAX(indx),0) FROM clubs')
    clubindx = int(cursor.fetchone()[0]) + 1
    data = (_name,)
    cursor.execute('INSERT INTO clubs (name) VALUES (?)', data)
    dbconn.commit()
    dbconn.close()
    return clubindx


def list_clubs(_dbname):
    """
    List all clubs
    """
    dbconn = sqlite3.connect(_dbname)
    cursor = dbconn.cursor()
    cursor.execute('SELECT * from clubs')
    data = cursor.fetchall()
    dbconn.commit()
    dbconn.close()
    return data


def main():

    # Init DB
    init_partcpdb(PARTCP_DB)
    init_eventsdb(EVENTS_DB)
    init_clubsdb(CLUBS_DB)

    # Import file
    root_window = tkinter.Tk()
    root_window.withdraw()
    data_file_path = filedialog.askopenfilename()
    print(data_file_path)
    assert(len(data_file_path) > 0)

    # Extract data
    rollnos = xlread(data_file_path)
    assert(len(rollnos) > 0)

    # Get details from user
    i_evname = input('+ Event Name : ')
    i_evdate = input('+ Event Date : ')
    i_evdatep = dateutil.parser.parse(i_evdate, dayfirst=True)
    i_uuid = uuid.uuid4()

    # Set up club data
    print('+ Clubs +')
    clubs = list_clubs(CLUBS_DB)
    for club in clubs:
        print(' '.join(map(str, club)))
    clubindex = input('+ Club Number (-1 for new club) : ')
    if int(clubindex) == -1:
        newclub = input('+ Club Name : ')
        clubindex = insert_club(newclub, CLUBS_DB)

    # Add event data
    eventindx = insert_event(i_evname, str(i_evdatep),
                             i_uuid, clubindex, EVENTS_DB)
    insert_participants(rollnos, eventindx, PARTCP_DB)


if __name__ == '__main__':
    main()
