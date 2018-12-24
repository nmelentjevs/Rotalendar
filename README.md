# Rotalendar

Rotalendar is a simple script initially created for Hippodrome Casino London staff which parses xls file and 
automatically creates events in Google Calendar with your work shift times.

## Installation

Script uses xlrd and datautil.parser modules to read and parse xls file.
Tkinter as an input gui.
And google authentication modules which can be found and downloaded at https://developers.google.com/calendar/auth

## Usage

Download the xlrotaparse.py and eventadder.py

First, you will need to allow google authenticator to access your Calendar and create events.
After than you will have your credentials.json file which should be copied to the same folder where eventadder.py is.
Then execute python eventadder.py, choose excel file to parse, input your name and surname -> done.

