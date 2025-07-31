function exportScheduleToCalendar() {
    const TZ = "America/Toronto"; // Eastern Time with DST

    const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schedule");
    if (!sheet) throw new Error("Sheet named 'Schedule' not found.");

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const calendarName = "Class Schedule 12312312";
    const calendar =
        CalendarApp.getCalendarsByName(calendarName)[0] ??
        CalendarApp.createCalendar(calendarName);

    calendar.setTimeZone(TZ);

    const dayMap = {
        M: CalendarApp.Weekday.MONDAY,
        T: CalendarApp.Weekday.TUESDAY,
        W: CalendarApp.Weekday.WEDNESDAY,
        Th: CalendarApp.Weekday.THURSDAY,
        F: CalendarApp.Weekday.FRIDAY,
    };

    function parseTimeToDate(baseDate, timeStr) {
        const m = /(\d+):(\d+)(AM|PM)/.exec(timeStr);
        if (!m) return new Date(baseDate);
        let h = +m[1],
            min = +m[2];
        if (m[3] === "PM" && h !== 12) h += 12;
        if (m[3] === "AM" && h === 12) h = 0;
        const d = new Date(baseDate);
        d.setHours(h, min, 0, 0);
        return d;
    }

    const headers = data[0];
    const col = {};
    headers.forEach((h, i) => (col[h] = i));

    for (let r = 1; r < data.length; r++) {
        const row = data[r];
        const courseCode = row[col["COURSE CODE"]];
        const component = row[col["COMPONENT"]];
        const section = row[col["SECTION"]];
        const classNum = row[col["CLASS NUMBER"]];
        const repeatedDays = row[col["REPEATED DAYS"]];
        const startTimeStr = row[col["START TIME"]];
        const endTimeStr = row[col["END TIME"]];
        const room = row[col["ROOM"]];
        const instructor = row[col["INSTRUCTOR"]];
        const startDate = new Date(row[col["START DATE"]]);
        const endDate = new Date(row[col["END DATE"]]);

        const title = `${courseCode} - ${component}`;
        const description = `Instructor: ${instructor}\n\nSection: ${section}\nClass #: ${classNum}`;

        const days =
            typeof repeatedDays === "string"
                ? (repeatedDays.match(/Th|M|T|W|F/g) || []).filter(
                      (d) => dayMap[d]
                  )
                : [];

        if (startDate.getTime() === endDate.getTime() || days.length === 0) {
            calendar.createEvent(
                title,
                parseTimeToDate(startDate, startTimeStr),
                parseTimeToDate(startDate, endTimeStr),
                { location: room, description, timeZone: TZ }
            );
            Utilities.sleep(10); // keep under quota
            continue;
        }

        const weekdayEnums = days.map((d) => dayMap[d]);

        const recurrence = CalendarApp.newRecurrence()
            .addWeeklyRule()
            .onlyOnWeekdays(weekdayEnums)
            // make the rule inclusive of the last meeting
            .until(parseTimeToDate(endDate, endTimeStr))
            .setTimeZone(TZ);

        // first date that matches one of the repeat days
        const jsWeekNums = { M: 1, T: 2, W: 3, Th: 4, F: 5 };
        const validNums = days.map((d) => jsWeekNums[d]);
        const firstDate = new Date(startDate);
        while (
            !validNums.includes(firstDate.getDay()) &&
            firstDate <= endDate
        ) {
            firstDate.setDate(firstDate.getDate() + 1);
        }

        calendar.createEventSeries(
            title,
            parseTimeToDate(firstDate, startTimeStr),
            parseTimeToDate(firstDate, endTimeStr),
            recurrence,
            { location: room, description }
        );
        Utilities.sleep(10);
    }
}
