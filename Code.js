function exportScheduleToCalendar() {
    var sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule");
    if (!sheet) {
        throw new Error("Sheet named 'Schedule' not found.");
    }
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    var calendarName = "Class Schedule";
    var calendars = CalendarApp.getCalendarsByName(calendarName);
    var calendar = calendars.length
        ? calendars[0]
        : CalendarApp.createCalendar(calendarName);

    var dayMap = {
        M: CalendarApp.Weekday.MONDAY,
        T: CalendarApp.Weekday.TUESDAY,
        W: CalendarApp.Weekday.WEDNESDAY,
        Th: CalendarApp.Weekday.THURSDAY,
        F: CalendarApp.Weekday.FRIDAY,
    };

    function parseTime(timeStr) {
        var date = new Date("1970-01-01T" + timeStr.replace(/(AM|PM)/, " $1"));
        if (isNaN(date.getTime())) {
            var parts = timeStr.match(/(\d+):(\d+)(AM|PM)/);
            if (parts) {
                var h = parseInt(parts[1], 10);
                var m = parseInt(parts[2], 10);
                if (parts[3] === "PM" && h !== 12) h += 12;
                if (parts[3] === "AM" && h === 12) h = 0;
                date = new Date(1970, 0, 1, h, m);
            }
        }
        return date;
    }

    var headers = data[0];
    var col = {};
    headers.forEach(function (h, i) {
        col[h] = i;
    });

    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var courseCode = row[col["COURSE CODE"]];
        var component = row[col["COMPONENT"]];
        var section = row[col["SECTION"]];
        var classNum = row[col["CLASS NUMBER"]];
        var repeatedDays = row[col["REPEATED DAYS"]];
        var startTimeStr = row[col["START TIME"]];
        var endTimeStr = row[col["END TIME"]];
        var room = row[col["ROOM"]];
        var instructor = row[col["INSTRUCTOR"]];
        var startDateStr = row[col["START DATE"]];
        var endDateStr = row[col["END DATE"]];

        var title = courseCode + " - " + component;
        var location = room;
        var description =
            "Instructor: " +
            instructor +
            "\n\nSection: " +
            section +
            "\nClass #: " +
            classNum;

        var startDate = new Date(startDateStr);
        var endDate = new Date(endDateStr);
        var startTime = parseTime(startTimeStr);
        var endTime = parseTime(endTimeStr);

        // Parse days like MWF, TTh, etc.
        var days = [];
        if (typeof repeatedDays === "string") {
            var dayMatches = repeatedDays.match(/Th|M|T|W|F/g);
            if (dayMatches) {
                days = dayMatches.filter(function (d) {
                    return dayMap[d];
                });
            }
        }

        for (
            var d = new Date(startDate);
            d <= endDate;
            d.setDate(d.getDate() + 1)
        ) {
            var weekday = d.getDay();
            var dayLetter = null;
            if (weekday === 1 && days.indexOf("M") !== -1) dayLetter = "M";
            if (weekday === 2 && days.indexOf("T") !== -1) dayLetter = "T";
            if (weekday === 3 && days.indexOf("W") !== -1) dayLetter = "W";
            if (weekday === 4 && days.indexOf("Th") !== -1) dayLetter = "Th";
            if (weekday === 5 && days.indexOf("F") !== -1) dayLetter = "F";
            if (dayLetter) {
                var eventStart = new Date(
                    d.getFullYear(),
                    d.getMonth(),
                    d.getDate(),
                    startTime.getHours(),
                    startTime.getMinutes()
                );
                var eventEnd = new Date(
                    d.getFullYear(),
                    d.getMonth(),
                    d.getDate(),
                    endTime.getHours(),
                    endTime.getMinutes()
                );
                calendar.createEvent(title, eventStart, eventEnd, {
                    location: location,
                    description: description,
                });
                Utilities.sleep(500);
            }
        }
    }
}
