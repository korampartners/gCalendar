function updateCalendarOps() {
  var CALENDAR_ID = "2e3cdd25e138ea93ff2acb1138330b2d9fd8ff29a94422700cc87d542e0c0092@group.calendar.google.com";

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  var eventCal = CalendarApp.getCalendarById(CALENDAR_ID);

  // Calculate dynamic date range: 3 months before today to 1 year from today
  var today = new Date();
  var startDate = new Date(today);
  startDate.setMonth(today.getMonth() - 3); // Go back 3 months from today
  
  var endDate = new Date(today);
  endDate.setFullYear(today.getFullYear() + 1); // Go forward 1 year from today
  
  // Delete existing events within the calculated date range
  var events = eventCal.getEvents(startDate, endDate);
  for (var i = 0; i < events.length; i++) {
    events[i].deleteEvent();
  }

  // Add new events
  for (var i = 1; i < data.length; i++) {
    var year = data[i][0];
    var week = data[i][1];
    var date = data[i][2];
    var event = data[i][3];
    var important = data[i][4];
    var actionItems = data[i][5];
    var comment = data[i][6];

    if (date && event) {
      var eventDate = parseDate(year, date);
      
      // Only add events if date is equal to or newer than startDate
      if (eventDate && eventDate >= startDate) {
        var description = "";
        Logger.log(`Adding ${event} on ${eventDate}`);
        if (actionItems) {
          description += "Action Items: " + actionItems + "\n";
          Logger.log(`Action Items Found : ${actionItems} for ${event} on ${eventDate}`); 
        }
        if (comment) { 
          description += "Comment: " + comment;
          Logger.log(`Comment Found : ${comment} for ${event} on ${eventDate}`);
        }
        var newEvent = eventCal.createAllDayEvent(event, eventDate, {
          description: description
        });

        if (important && (important.toLowerCase() === 'yes' || important.toLowerCase() === 'y')) {
          newEvent.addPopupReminder(24 * 60); // 24-hour reminder
        }

        Logger.log(`Added Event : ${event} on ${eventDate}`);
      } else if (eventDate) {
        Logger.log(`Skipped event as date is before startDate: ${event} on ${eventDate}`);
      } else {
        Logger.log(`Skipped event due to invalid date: ${event} (${date})`);
      }
    }
  }
}

function parseDate(yearString, dateString) {
  if (!dateString) return null;
  
  // Convert to string and ensure it's a new string object
  dateString = new String(dateString).toString();
  
  // Parse Korean format using regular expression
  try {
    // Extract numbers from the Korean date format
    var numbers = dateString.split(/[월일]/).filter(function(n) {
      return n.trim() !== '';
    });
    
    if (numbers.length === 2) {
      var month = parseInt(numbers[0], 10);
      var day = parseInt(numbers[1], 10);
      
      if (!isNaN(month) && !isNaN(day)) {
        return new Date(parseInt(yearString), month - 1, day);
      }
    }

    // If Korean parsing fails, try other formats
    var formats = [
      'yyyy년 M월 d일',
      'yyyy년 MM월 dd일',
      'yyyy-MM-dd',
      'M/d',
      'MM/dd'
    ];

    for (var i = 0; i < formats.length; i++) {
      try {
        var parsedDate = Utilities.parseDate(dateString, 'Asia/Seoul', formats[i]);
        if (parsedDate) {
          parsedDate.setFullYear(parseInt(yearString));
          return parsedDate;
        }
      } catch (e) {
        // Continue to next format if parsing fails
      }
    }
  } catch (e) {
    Logger.log('Error parsing date: ' + e);
  }

  return null;
}
