#!/usr/bin/env python

"""
THWin2iCalendar - Converts CSV output from THWin to an iCalendar

Written in 2015 by Tim Ruffing <tim@timruffing.de>

To the extent possible under law, the author(s) have dedicated all copyright and related and neighboring rights to this software to the public domain worldwide. This software is distributed without any warranty.

You should have received a copy of the CC0 Public Domain Dedication along with this software. If not, see <http://creativecommons.org/publicdomain/zero/1.0/>.
"""

# Exportieren aus THWin:
# 1. Modul "Verwaltung" -> "Dienst" -> "Dienste"
# 2. Menü "Datei" -> "Drucken"
# 3. "Dienst- / Ausbildungsplan" -> CSV-Datei

import csv
from icalendar import Calendar, Event, vDatetime
from datetime import datetime, timedelta
import hashlib
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename

START = 1
END = 2
LOCATION = 4
TYPE = 5
CLOTHES = 6
SUMMARY_TOPIC = 7
RESPONSIBLE = 8
PARTICIPANTS = 9

# TODO config file?
UID_SUFFIX = '@thw-igb.de'
DEFAULT_INFILE = 'dienstplan.csv'
DEFAULT_OUTFILE = 'dienstplan.ics'

counter = 0

def main():
    root = Tk()
    root.withdraw()

    infile = infile_picker()
    if infile == "":
        sys.exit(1)

    outfile = outfile_picker()
    if outfile == "":
        sys.exit(1)

    # TODO error handling
    reader = csv.reader(open(infile, 'r', encoding='iso-8859-15'), delimiter=';')
    cal = create_calendar(reader)

    # save file
    f = open(outfile, 'wb')
    f.write(cal.to_ical())
    f.close()

def infile_picker():
    dialog_opt = {}
    dialog_opt['defaultextension'] = '.csv'
    dialog_opt['filetypes'] = [('CSV-Format', '.csv'), ('Alle Dateien', '.*')]
    dialog_opt['initialdir'] = 'G:\\THWInExport'
    dialog_opt['title'] = 'CSV-Datei (aus THWin) öffnen'
    return askopenfilename(**dialog_opt)

def outfile_picker():
    # TODO implement initial filename
    dialog_opt = {}
    dialog_opt['defaultextension'] = '.ics'
    dialog_opt['filetypes'] = [('iCalendar', '.ics'), ('Alle Dateien', '.*')]
    dialog_opt['initialdir'] = 'G:\\THWinExport'
    dialog_opt['title'] = 'iCalendar-Datei speichern'
    return asksaveasfilename(**dialog_opt)

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
    event.add('dtstart', get_dtstart(row))
    event.add('dtend', get_dtend(row))
    summary, desc = get_summary_and_description(row)
    event.add('summary', summary)
    event.add('description', desc)
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
    summary = row[SUMMARY_TOPIC]
    while len(summary) > 0 and summary[0] == '[':
        tag, summary  = summary[1::].split(']', 1)
        # there might be leading space
        summary.lstrip()
        tags.append(sanitize(tag))
    tags.append(get_type(row))
    return tags

def get_summary_and_description(row):
    typ = get_type(row)
    desc  = typ + ' (' + row[CLOTHES] + ')' + '\n\n'
    if get_training(row):
        desc += 'Themen:\n'
    desc += row[SUMMARY_TOPIC].strip() + '\n\n'

    if len(row[RESPONSIBLE]) > 0:
        desc += 'Leitende:\n' + row[RESPONSIBLE] + '\n\n'

    if len(row[PARTICIPANTS]) > 0:
        desc += 'Teilnehmer:\n' + row[PARTICIPANTS]

    lines = row[SUMMARY_TOPIC].strip().splitlines()

    summary = ''
    if not is_general(row):
        summary += typ
        if len(lines) > 0:
            summary += ': '

    if len(lines) > 0:
        summary += lines[0]
    if len(lines) > 1:
        summary += " (...)"

    return (sanitize(summary), desc)

def get_type(row):
    training = get_training(row)
    if training:
        return training
    else:
        return sanitize(row[TYPE])


def get_training(row):
    if row[TYPE][:3] == 'S -':
        return sanitize(row[TYPE][3:])
    else:
        return None

def is_general(row):
    return row[TYPE] in ['Dienst allgemein', 'Jugendarbeit']

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

def sanitize(s):
    return s.strip().replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')

if __name__ == "__main__":
    main()
