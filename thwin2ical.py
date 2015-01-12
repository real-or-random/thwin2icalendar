#!/usr/bin/env python

"""
THWin2iCal - Converts CSV output from THWin to an iCal calendar

Written in 2015 by Tim Ruffing <tim@timruffing.de>

To the extent possible under law, the author(s) have dedicated all copyright and related and neighboring rights to this software to the public domain worldwide. This software is distributed without any warranty.

You should have received a copy of the CC0 Public Domain Dedication along with this software. If not, see <http://creativecommons.org/publicdomain/zero/1.0/>.
"""

import csv
from icalendar import Calendar, Event, vDatetime
from datetime import datetime, timedelta
import hashlib
from tkinter.filedialog import askopenfilename, asksaveasfilename

# TODO constants
START = 1
END = 2
CATEGORIES = 3
LOCATION = 4
TYPE = 5
CLOTHES = 6
SUMMARY = 7
RESPONSIBLE = 8
PARTICIPANTS = 9

# TODO config file?
UID_SUFFIX = '@thw-igb.de'
DEFAULT_INFILE = 'dienstplan.csv'
DEFAULT_OUTFILE = 'dienstplan.ics'

counter = 0

def main(outfile=DEFAULT_OUTFILE):
    infile = infile_picker()
    reader = csv.reader(open(infile, 'r'), delimiter=';')
    cal = create_calendar(reader)

    # save file
    f = open(outfile, 'wb')
    f.write(cal.to_ical())
    f.close()

def infile_picker():
    infile_dialog_opt = {}
    infile_dialog_opt['defaultextension'] = '.csv'
    infile_dialog_opt['filetypes'] = [('CSV-Dateien', '.csv'), ('Alle Dateien', '.*')]
    infile_dialog_opt['initialdir'] = 'G:\\Thwin Export' # FIXME check dir name
    infile_dialog_opt['title'] = 'CSV-Datei (aus THWin) auswählen'
    return askopenfilename(**infile_dialog_opt)

def create_calendar(csv_reader):
    # skip first row
    # TODO verify first row
    next(csv_reader)

    cal = Calendar()
    cal.add('prodid', '-//THW//THWIn2iCal//')
    cal.add('version', '2.0')
    for row in csv_reader:
        cal.add_component(create_event(row))

    return cal

def create_event(row):
    global counter
    event = Event()
    event.add('summary', get_summary(row))
    event.add('dtstart', get_dtstart(row))
    event.add('dtend', get_dtend(row))
    event.add('description', get_description(row))
    event.add('location', get_location(row))
    cats = get_categories(row)
    if len(cats) > 0:
        event.set_inline('categories', get_categories(row))
    event.add('uid', get_uid(row))
    counter += 1
    return event

def parse_date(s):
    # TODO error handling
    return datetime.strptime(s, '%d.%m.%Y, %H:%M:%S')

def get_dtstart(row):
    return parse_date(row[START])

def get_dtend(row):
    return parse_date(row[END])

# Categories can be added at the beginning of the summary,
# using square brackets
# e.g. "[FUG Saarpfalzkreis][Übung] Flächenlage"
def get_categories(row):
    tags = []
    summary = row[SUMMARY]
    while len(summary) > 0 and summary[0] == '[':
        tag, summary  = summary[1::].split(']', 1)
        # there might be leading space
        row[SUMMARY].lstrip()
        tags.append(tag)
    return tags

def get_description(row):
    # TODO check if summary can be multiline in THWin
    desc  = row[TYPE] + ' (' + row[CLOTHES] + ')'     + '\n\n'
    desc +=                                            '\n'
    desc += 'Verantwortliche:\n' + row[RESPONSIBLE]  + '\n\n'
    desc += 'Teilnehmer:\n'      + row[PARTICIPANTS]
    return desc

def get_summary(row):
    return row[SUMMARY]

def get_location(row):
    return row[LOCATION]

def get_uid(row):
    # "persistent" is not really persistent, it is at most a good heuristic.
    # We cannot do really better as THWin does not provide us an unique persistent input.
    # TODO check if PARTICIPANTS is a good idea... maybe this is often changed because we fiddle with the statistics
    persistent = row[START] + row[PARTICIPANTS]
    return sha1(persistent) + str(counter) + UID_SUFFIX

def sha1(s):
    return hashlib.sha1(s.encode('utf-8')).hexdigest()

if __name__ == "__main__":
    main()
