# G-Suite meeting notes automation

Google App Script for automating the creation of meeting minutes.
This will create two new objects:
* A google document with the meeting minutes
* A calendar event in a minutes calendar with the minutes document attached, so you can locate them quickly from the calendar

This tries to replicate the integration between Outlook and OneNote.

This script is heavily based on https://github.com/daubejb/meeting-notes-for-google-calendar.

## Installation

1.- Create a new Google App Script from Drive
2.- Copy the contents of the meeting_minutes.gs file in the editor.
3.- Save the script
4.- Run it to test the script

To schedule it, follow the instructions in https://developers.google.com/apps-script/guides/triggers/installable#managing_triggers_manually 
