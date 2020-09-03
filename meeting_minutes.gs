// Based on https://github.com/daubejb/meeting-notes-for-google-calendar/blob/master/meeting_notes.gs
function defineOptions() {
  
  var options = {};
  
  // define a custom style for all data labels
  var labelStyle ={};
  labelStyle[DocumentApp.Attribute.BOLD] = true;
  labelStyle[DocumentApp.Attribute.FONT_SIZE] = 12;
  labelStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Nunito';
  options.labelStyle = labelStyle;
  
  var titleStyle = {};
  titleStyle[DocumentApp.Attribute.BOLD] = true;
  titleStyle[DocumentApp.Attribute.FONT_SIZE] = 16;
  titleStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Nunito';
  options.titleStyle = titleStyle;
  
  var textStyle = {};
  textStyle[DocumentApp.Attribute.BOLD] = false;
  textStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  textStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Nunito';
  options.textStyle = textStyle;
  
  // timezone, by default the one defined in the default calendar
  options.timezone = CalendarApp.getDefaultCalendar().getTimeZone();
  
  // date formats
  var directoryStructureDateFormat = "yyyy-MM-dd";
  var readDateFormat = "dd-MM-yyyy";
  options.readDateTimeFormat = "dd-MM-yyyy HH:mm";
  
  // create a formatted version of today's date
  var today = new Date();
  options.formattedToday = Utilities.formatDate(today, options.timezone, directoryStructureDateFormat);
  options.formattedTodayHeader = Utilities.formatDate(today, options.timezone, readDateFormat);
  
  // number of hours from now to check for meetings
  options.hours = 2;
  
  // create variable for now
  options.now = new Date();
  
  // create variable for number of hours from now in milliseconds
  options.periodFromNow = new Date(options.now.getTime() + (options.hours * 60 * 60 * 1000));

  // name of the root folder
  options.rootFolder = 'Meeting Notes';
  
  // name of the minutes calendar
  options.calendarName = 'Meeting minutes';
  
  return options;

}

function createMeetingNotesNextTimePeriod() {
  //<-------------------------------------VARIABLES------------------------------------------>
  var opts = defineOptions();
  
  // creates the meeting folder in GDrive if it doesn't exists and return the reference to it
  var meetingNotesFolder = createMeetingsFolderIfDoesNotExists(opts.rootFolder);
  
  //<-------------------CREATE A MEETING NOTES CALENDAR IF MISSING-------------------------->
  
  var minutesCalendarId = createMinutesCalendarIfDoesNotExists(opts.calendarName);
  
  
  //<--------------------GET EVENTS/ATTRIBUTES FOR TIME PERIOD FROM NOW--------------------->
  
  
  // retrieve all calendar events for time period
  var events = CalendarApp.getDefaultCalendar().getEvents(opts.now, opts.periodFromNow);  
  Logger.log('Number of events in the next ' + opts.hours + ' hours: ' + events.length);
  
  // create a folder for today's notes only if folder does not exist and if events are found
  if (events.length > 0) {
    var todaysNotesFolderId = getOrCreateDayMinutesFolder(meetingNotesFolder, opts.formattedToday);
    
    // loop through each event an get meeting attributes, 
    for (var i=0;i<events.length;i++) {
      var meeting = {};
      meeting.title = events[i].getTitle();
      meeting.description = events[i].getDescription();
      meeting.eventStart = events[i].getStartTime();
      meeting.eventEnd = events[i].getEndTime();
      meeting.location = events[i].getLocation();
      meeting.owner = events[i].getCreators();
      meeting.guestList = events[i].getGuestList();
      
      var params = {};
      params.formattedToday = opts.formattedToday;
      params.formattedTodayHeader = opts.formattedTodayHeader;
      params.readDateTimeFormat = opts.readDateTimeFormat;
      params.timezone = opts.timezone;
      params.folderId = todaysNotesFolderId;
      params.styles = {};
      params.styles.title = opts.titleStyle;
      params.styles.label = opts.labelStyle;
      params.styles.text = opts.textStyle;
      
      // create a google doc with the meeting name as the title
      var fileExists = DriveApp.getFoldersByName(opts.formattedToday).next().getFilesByName(meeting.title).hasNext();
      
      // check to see if file already exists, if does skip if loop
      if ((!fileExists) && (meeting.guestList.length >= 1)) {
        Logger.log('Minutes file for meeting ' + meeting.title + ' does not exist. Creating ...');
        var document = createMeetingMinutesDocument(meeting, params);
        
        Logger.log('Creating calendar event in the meeting minutes calendar...');
        createCalendarEventWithAttachment(minutesCalendarId, meeting, document);
      } else {
        Logger.log('Minutes file for meeting ' + meeting.title + ' already exists. Skipping it.');
      }
    }
  }
}

/*
 *
 * HELPER FUNCTIONS
 *
 */

/*
 * Function that creates the root minutes folder in your Drive account with the name specified in the parameter. 
 * Please note that, if this folder has already been created in your account, all the minutes will be stored there.
 *
 */
function createMeetingsFolderIfDoesNotExists(folderName) {
  // check to see if a meeting notes folder exists
  var meetingNotesFolderExists = DriveApp.getFoldersByName(folderName).hasNext();
  
  // create the main meeting notes folder if it does note exist
  if (!meetingNotesFolderExists) {
    Logger.log(folderName + ' folder does not exist. Creating it...');
    DriveApp.createFolder(folderName);
    Logger.log(folderName + ' folder created');
  } else {
    Logger.log(folderName + ' folder exists.');
  }
  
  // locate folder named Meeting Notes and instatiate variable
  return DriveApp.getFoldersByName(folderName).next();
}

/*
 * Function that creates the minutes calendar in your google account with the name specified in the parameter. 
 * Please note that, if this calendar has already been created in your account, all the minute events will be stored there.
 *
 */
function createMinutesCalendarIfDoesNotExists(calendarName) {
  var calendarExists = CalendarApp.getCalendarsByName(calendarName).length > 0;
  
  if (!calendarExists) {
    Logger.log(calendarName + ' calendar does not exist. Creating it...');
    // create the new calendar
    var minutesCalendar = CalendarApp.createCalendar(calendarName);
    Logger.log(calendarName + ' calendar created.');
  } else {
    Logger.log(calendarName + ' calendar already exists.');
    var minutesCalendar = CalendarApp.getCalendarsByName(calendarName)[0];
  }
  
  return minutesCalendar.getId();
}

function getOrCreateDayMinutesFolder(rootFolder, date) {
  var dateFolderExists = rootFolder.getFoldersByName(date).hasNext();
  
  // create the folder if it does not exist
  if (!dateFolderExists) {
    Logger.log(date + ' folder does not exist. Creating it...');
    rootFolder.createFolder(date);
    Logger.log(date + 'folder created');
  } else {
    Logger.log(date + ' folder already exists.');
  }            
  return rootFolder.getFoldersByName(date).next().getId();
}

function createMeetingMinutesDocument(meeting, params) {

  var documentResource = {
    title: meeting.title,
    mimeType: MimeType.GOOGLE_DOCS,
    parents: [{id: params.folderId}]
  }
  
  var docJson = Drive.Files.insert(documentResource);
  var documentId = docJson.id;
  
  // Create a calendar event with the referenced 
  
  var doc = DocumentApp.openById(documentId);
  
  var body = doc.getBody();
  
  // create title header
  var titleParagraph = body.appendParagraph(meeting.title + ' [' + params.formattedTodayHeader + ']');
  titleParagraph.setHeading(DocumentApp.ParagraphHeading.TITLE);
  titleParagraph.setAttributes(params.styles.title);
  
  var descriptionParagraph = body.appendParagraph('Description:\n' + meeting.description);
  descriptionParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  descriptionParagraph.setAttributes(params.styles.label);
  
  var formattedStartDate = Utilities.formatDate(meeting.eventStart, params.timezone, params.readDateTimeFormat);
  var startDateParagraph = body.appendParagraph('Start: ' + formattedStartDate);
  startDateParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  startDateParagraph.setAttributes(params.styles.label);
  
  var formattedEndDate = Utilities.formatDate(meeting.eventEnd, params.timezone, params.readDateTimeFormat);
  var endDateParagraph = body.appendParagraph('End: ' + formattedEndDate);
  endDateParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  endDateParagraph.setAttributes(params.styles.label);
  
  var locationParagraph = body.appendParagraph('Location: ' + meeting.location);
  locationParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  locationParagraph.setAttributes(params.styles.label);
  
  var ownerParagraph = body.appendParagraph('Organizer: ' + meeting.owner);
  ownerParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  ownerParagraph.setAttributes(params.styles.label);
  
  var guestsParagraph = body.appendParagraph('Guest List:');
  guestsParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  guestsParagraph.setAttributes(params.styles.label);
  
  // loop through event guests and get their emails and status
  for (var j=0 ; j < meeting.guestList.length ; j++) {
    var guestParagraph = body.appendParagraph(meeting.guestList[j].getEmail() + ': ' + meeting.guestList[j].getGuestStatus());
    guestParagraph.setAttributes(params.styles.text);
  }
  
  body.appendHorizontalRule();
  
  var discussionsParagraph = body.appendParagraph('Discussions:');
  discussionsParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  discussionsParagraph.setAttributes(params.styles.label);
  
  var discussionsText = body.appendParagraph('...');
  discussionsText.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  discussionsText.setAttributes(params.styles.text);
  
  var actionsParagraph = body.appendParagraph('Action Points:');
  actionsParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  actionsParagraph.setAttributes(params.styles.label);
  
  var actionsText = body.appendParagraph('...');
  actionsText.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  actionsText.setAttributes(params.styles.text);
  
  // Get the minutes file reference from Google Drive
  var minutesFileId = DriveApp.getFoldersByName(params.formattedToday).next().getFilesByName(meeting.title).next().getId();
  var minutesDriveFile = Drive.Files.get(minutesFileId);
  
  Logger.log('The file reference in google drive is: ' + minutesFileId);
  return minutesDriveFile;
}

function createCalendarEventWithAttachment(calendarId, meeting, attachment) {
  var eventResource = {
    summary: meeting.title,
    start: { dateTime: meeting.eventStart.toISOString() },
    end: { dateTime: meeting.eventEnd.toISOString() },
    attachments: [{
      fileUrl: attachment.alternateLink,
      mimeType: attachment.mimeType,
      title: attachment.title
    }]
  };
  
  var minutesEvent = Calendar.Events.insert(eventResource, calendarId, {'supportsAttachments': true});  
  Logger.log('Calendar event created with id: ' + minutesEvent.id);
}
