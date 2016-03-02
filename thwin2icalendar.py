#!/usr/bin/env python

"""
THWin2iCalendar - Converts CSV output from THWin to an iCalendar

Written in 2015 by Tim Ruffing <tim@timruffing.de>

To the extent possible under law, the author(s) have dedicated all copyright and related and neighboring rights to this software to the public domain worldwide. This software is distributed without any warranty.

You should have received a copy of the CC0 Public Domain Dedication along with this software. If not, see <http://creativecommons.org/publicdomain/zero/1.0/>.
"""

import csv
import codecs
from icalendar import Calendar, Event, vDatetime
from datetime import datetime
from pytz import timezone
import hashlib
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from tkinter.messagebox import showerror
import os.path
import sys
import re

START = 'Beginn'
END = 'Ende'
LOCATION = 'Ort'
TYPE = 'Dienstart'
CLOTHES = 'Bekleidung'
SUMMARY_TOPIC = 'Thema'
RESPONSIBLE = 'Leitende'
PARTICIPANTS = 'Teilnehmer'

FIELDNAMES = [START, END, LOCATION, TYPE, CLOTHES, SUMMARY_TOPIC, RESPONSIBLE, PARTICIPANTS]

TZ_DEF = b'''BEGIN:VTIMEZONE\r\n\
TZID:Europe/Berlin\r\n\
BEGIN:DAYLIGHT\r\n\
TZOFFSETFROM:+0100\r\n\
TZOFFSETTO:+0200\r\n\
TZNAME:MESZ\r\n\
DTSTART:19700329T020000\r\n\
RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=-1SU;BYMONTH=3\r\n\
END:DAYLIGHT\r\n\
BEGIN:STANDARD\r\n\
TZOFFSETFROM:+0200\r\n\
TZOFFSETTO:+0100\r\n\
TZNAME:MEZ\r\n\
DTSTART:19701025T030000\r\n\
RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=-1SU;BYMONTH=10\r\n\
END:STANDARD\r\n\
END:VTIMEZONE'''

def main():
    root = Tk()
    root.withdraw()

    infilename = infile_picker()
    if len(infilename) == 0:
        sys.exit(1)

    infile = None
    # The THWin export is ISO-8859-15 encoded. To avoid hassle with manually
    # edited files we support UTF-8 as well.
    # Try UTF-8 first.
    try:
        infile = codecs.open(infilename, 'r', encoding='utf-8', errors='strict')
        indata = infile.readlines()
    except UnicodeDecodeError:
        infile.close()
        infile = codecs.open(infilename, 'r', encoding='iso-8859-15')
        indata = infile.readlines()

    reader = csv.DictReader(indata, delimiter=';')
    # verify first row
    try:
        next(reader)
    except StopIteration:
        fatal_error('Die ausgewählte Eingabedatei ist keine gültige CSV-Datei.')

    if not all((f in reader.fieldnames) for f in FIELDNAMES):
        fatal_error('Die ausgewählte Eingabedatei enthält keinen THWin-Dienstplan.')

    mtime = os.path.getmtime(infilename)
    dtstamp = datetime.fromtimestamp(mtime)
    cal = create_calendar(reader, dtstamp)

    infile.close()

    outfilename = outfile_picker(infilename)
    if len(outfilename) == 0:
        sys.exit(1)

    # save file
    f = open(outfilename, 'wb')
    outlines = cal.to_ical().splitlines()
    # output to file and add hackish VTIMEZONE definition
    # (The icalendar module does not support exporting of pytz objects.
    # An alternative way would be to create an icalendar Timezone
    # object manually.)
    f.write(b'\r\n'.join(outlines[:3]) + b'\r\n')
    f.write(TZ_DEF + b'\r\n')
    f.write(b'\r\n'.join(outlines[3:]))
    f.close()

def infile_picker():
    dialog_opt = {}
    dialog_opt['defaultextension'] = '.csv'
    dialog_opt['filetypes'] = [('CSV-Format', '.csv'), ('Alle Dateien', '.*')]
    dialog_opt['initialdir'] = 'G:\\THWInExport'
    dialog_opt['title'] = 'CSV-Datei (aus THWin) öffnen'
    return askopenfilename(**dialog_opt)

def outfile_picker(infile):
    dialog_opt = {}
    dialog_opt['defaultextension'] = '.ics'
    dialog_opt['filetypes'] = [('iCalendar', '.ics'), ('Alle Dateien', '.*')]
    dialog_opt['initialdir'] = 'G:\\THWinExport'
    dialog_opt['initialfile'] = os.path.splitext(os.path.basename(infile))[0] + '.ics'
    dialog_opt['title'] = 'iCalendar-Datei speichern'
    return asksaveasfilename(**dialog_opt)

def create_calendar(csv_reader, dtstamp):
    cal = Calendar()
    cal.add('prodid', '-//THW//THWIn2iCal//')
    cal.add('version', '2.0')
    for row in csv_reader:
        event = create_event(row)
        event.add('dtstamp', dtstamp)
        cal.add_component(event)

    return cal

def create_event(row):
    event = Event()
    event.add('dtstart', get_dtstart(row))
    event.add('dtend', get_dtend(row))
    event.add('uid', get_uid(row))
    summary, desc, categories = get_summary_description_categories(row)
    event.add('summary', summary)
    event.add('description', desc)
    event.add('categories', categories)
    event.add('location', get_location(row))
    return event

def parse_date(s):
    d = datetime.strptime(s, '%d.%m.%Y, %H:%M:%S')
    return d.replace(tzinfo = timezone('Europe/Berlin'))

def get_dtstart(row):
    return parse_date(row[START])

def get_dtend(row):
    return parse_date(row[END])

# Categories can be added at the beginning of the description,
# using square brackets
# e.g. "[Besprechung] OA-Sitzung"
def get_tags(indesc):
    tags = []
    while len(indesc) > 0 and indesc[0] == '[':
        tag, indesc  = indesc[1::].split(']', 1)
        # there might be leading space
        indesc.lstrip()
        tags.append(sanitize(tag))
    return tags

def get_summary_description_categories(row):
    typ = get_type(row)
    desc = typ + '\n\n'

    lines = row[SUMMARY_TOPIC].strip().splitlines()

    topics = []
    i = 0
    if get_training(row):
        for i in range(len(lines)):
            # check for curriculum numbers
            if not re.match('\(([0-9]|\.)*\)  ', lines[i]):
                break
    topics = lines[:i]
    indesc = lines[i:]
    # remove empty line between topics and description
    if len(indesc[0]) == 0:
        indesc.pop(0)

    desc += 'Beschreibung:\n  '
    desc += '\n  '.join(indesc)
    desc += '\n\n'

    if len(topics) > 0:
        desc += 'Themen:\n'
        desc += format_list(topics) + '\n\n'

    if len(row[CLOTHES]) > 0:
        desc += 'Bekleidung:\n  ' + row[CLOTHES].strip() + '\n\n'

    if len(row[RESPONSIBLE]) > 0:
        responsible = sanitize_persons(row[RESPONSIBLE]).splitlines()
        desc += 'Leitende:\n' + format_list(responsible) + '\n\n'

    if len(row[PARTICIPANTS]) > 0:
        participants = sanitize_persons(row[PARTICIPANTS]).splitlines()
        desc += 'Teilnehmer:\n' + format_list(participants)

    categories = [typ]

    if len(indesc) > 0 and len(indesc[0]) > 0:
        summary = indesc[0]
        categories += get_tags(indesc[0])
    else:
        summary = typ

    return (summary, desc, categories)

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
    return row[TYPE] in ['Dienst allgemein', 'Jugendarbeit', 'sonstige technische Hilfeleistung']

def get_location(row):
    return row[LOCATION]

def get_uid(row):
    # "persistent" is not really persistent or unique, it is at most a good heuristic.
    # We cannot do really better as THWin does not provide us an unique persistent input.
    persistent = row[START] + row[END] + row[TYPE] + row[CLOTHES] + row[LOCATION] + row[SUMMARY_TOPIC]
    return digest(persistent)

def digest(s):
    return hashlib.sha256(s.encode('utf-8')).hexdigest()

def sanitize(s):
    return s.strip().replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')

def sanitize_persons(s):
    # remove unnecessary linebreak after name
    s = s.replace('\r\n(', ' (').replace('\n(', ' (').replace('\r(', ' (')
    return s

def format_list(ss):
    prefix = '  *' + chr(160) # NO-BREAK SPACE
    return '\n'.join([prefix + sanitize(s) for s in ss])

def error(msg):
    showerror("Fehler", msg)

def fatal_error(msg):
    error(msg)
    sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        error(e)
        raise
