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

START = 'Beginn'
END = 'Ende'
LOCATION = 'Ort'
TYPE = 'Dienstart'
CLOTHES = 'Bekleidung'
SUMMARY_TOPIC = 'Thema'
RESPONSIBLE = 'Leitende'
PARTICIPANTS = 'Teilnehmer'

FIELDNAMES = [START, END, LOCATION, TYPE, CLOTHES, SUMMARY_TOPIC, RESPONSIBLE, PARTICIPANTS]

UID_SUFFIX = 'thwin2icalendar'

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
        error('Die ausgewählte Eingabedatei ist keine gültige CSV-Datei.')

    if not all((f in reader.fieldnames) for f in FIELDNAMES):
        error('Die ausgewählte Eingabedatei enthält keinen THWin-Dienstplan.')

    cal = create_calendar(reader)

    infile.close()

    outfilename = outfile_picker(infilename)
    if len(outfilename) == 0:
        sys.exit(1)

    # save file
    f = open(outfilename, 'wb')
    f.write(cal.to_ical())
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

def create_calendar(csv_reader):
    cal = Calendar()
    cal.add('prodid', '-//THW//THWIn2iCal//')
    cal.add('version', '2.0')
    for row in csv_reader:
        cal.add_component(create_event(row))

    return cal

def create_event(row):
    event = Event()
    event.add('dtstart', get_dtstart(row))
    event.add('dtend', get_dtend(row))
    summary, desc = get_summary_and_description(row)
    event.add('summary', summary)
    event.add('description', desc)
    event.add('location', get_location(row))
    event.add('categories', get_categories(row))
    return event

def parse_date(s):
    d = datetime.strptime(s, '%d.%m.%Y, %H:%M:%S')
    return d.replace(tzinfo = timezone('Europe/Berlin'))

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
    desc = typ + '\n\n'

    if get_training(row):
        desc += 'Themen:\n'
        desc += format_list(row[SUMMARY_TOPIC].strip())
    else:
        desc += 'Beschreibung:\n'
        desc += row[SUMMARY_TOPIC].strip()
    desc += '\n\n'

    if len(row[CLOTHES]) > 0:
        desc += 'Bekleidung:\n' + row[CLOTHES].strip() + '\n\n'

    if len(row[RESPONSIBLE]) > 0:
        desc += 'Leitende:\n' + format_list_persons(row[RESPONSIBLE]) + '\n\n'

    if len(row[PARTICIPANTS]) > 0:
        desc += 'Teilnehmer:\n' + format_list_persons(row[PARTICIPANTS])

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
    return row[TYPE] in ['Dienst allgemein', 'Jugendarbeit', 'sonstige technische Hilfeleistung']

def get_location(row):
    return row[LOCATION]

def get_uid(row):
    # "persistent" is not really persistent or unique, it is at most a good heuristic.
    # We cannot do really better as THWin does not provide us an unique persistent input.
    persistent = row[START] + row[END] + row[TYPE] + row[CLOTHES] + row[LOCATION] + row[SUMMARY_TOPIC]
    return digest(persistent) + UID_SUFFIX

def digest(s):
    return hashlib.sha256(s.encode('utf-8')).hexdigest()

def sanitize(s):
    return s.strip().replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')

def format_list(s):
    prefix = '  *' + chr(160) # NO-BREAK SPACE
    return  prefix + s.replace('\n', '\n' + prefix).replace('\r', '\r' + prefix)

def format_list_persons(s):
    # remove unnecessary linebreak after name
    s = s.replace('\r\n(', ' (').replace('\n(', ' (').replace('\r(', ' (')
    return format_list(s)

def error(msg):
    showerror("Fehler", msg)
    sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        error(e)
