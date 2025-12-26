function getMonthString(month) {
    const monthNames = [
        'January', 'February', 'March', 'April',
        'May', 'June', 'July', 'August',
        'September', 'October', 'November', 'December'
    ];

    return monthNames[month]
}

function isHoliday(date) {
    const calendarId = 'en.indonesian#holiday@group.v.calendar.google.com'; // Indonesian Holidays Calendar ID
    const calendar = CalendarApp.getCalendarById(calendarId);

    if (!calendar) {
        Logger.log('Calendar not found. Check the calendar ID.');
        return false;
    }

    // Ensure the date is properly formatted
    const normalizedDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const events = calendar.getEventsForDay(normalizedDate);

    return events.length > 0
}


function isWeekend(date) {
    const day = date.getDay();
    return (day === 0 || day === 6)
}

function isSameDay(date1, date2) {
    return date1.toDateString() === date2.toDateString();
}


function getDateNow(date = null) {
    return Utilities.formatDate(date || new Date(), "Asia/Jakarta", "MM/d/yyyy H:mm:ss");
}

const getDateAWeekAgo = () => {
    const date = new Date();
    date.setDate(date.getDate() - 7);
    return date.toLocaleString(); // or format as needed
};

function getDayDiff(date) {
    const todayDate = new Date();
    const dayDiff = (todayDate - date) / (1000 * 60 * 60 * 24);
    return dayDiff;
}

function getMinuteDiff(date) {
    const todayDate = new Date();

    // Convert milliseconds to minutes
    const minuteDiff = (todayDate - new Date(date)) / (1000 * 60);
    return minuteDiff;
}

function isDateExpired(date) {
    date = new Date(date);
    const dayDiff = getDayDiff(date);

    if (dayDiff < EXPIRED_DAY_LIMIT) return false;

    // Calculate extended day limit considering holidays and weekends
    let extendedLimit = EXPIRED_DAY_LIMIT;
    for (let i = 0; i < dayDiff; i++) {
        const checkDate = new Date(date);
        checkDate.setDate(date.getDate() + i);
        if (isHoliday(checkDate) || isWeekend(checkDate)) {
            extendedLimit++;
        }
    }
    // Check if the date is expired considering the extended limit
    return dayDiff > extendedLimit;
}

function parseMDYHMS(str) {
    const [datePart, timePart] = str.split(' ');
    if (!datePart || !timePart) {
        throw new Error(`Invalid date format: "${str}"`);
    }

    const [month, day, year] = datePart.split('/').map(Number);
    const [hour, minute, second] = timePart.split(':').map(Number);

    const d = new Date(year, month - 1, day, hour, minute, second);
    if (isNaN(d.getTime())) {
        throw new Error(`Parsed to invalid Date: "${str}"`);
    }
    return d;
}

function parseHms(str) {
    if (typeof str !== 'string') return Infinity;
    const parts = str.split(':').map(p => parseInt(p, 10));
    if (parts.some(isNaN)) return Infinity;

    let seconds = 0;
    if (parts.length === 3) {
      // [jam, menit, detik]
      seconds = parts[0] * 3600 + parts[1] * 60 + parts[2];
    } else if (parts.length === 2) {
      // [menit, detik]
      seconds = parts[0] * 60 + parts[1];
    } else {
      return Infinity;
    }
    return seconds;
}

/**
 * Add an arbitrary amount of time to a Date.
 * All fields default to 0.
 */
function addTime(origDate, options) {
    // 1) Validate & clone
    if (!(origDate instanceof Date) || isNaN(origDate.getTime())) {
        throw new TypeError("addTime: first argument must be a valid Date");
    }
    var d = new Date(origDate.getTime());

    // 2) Pull out each piece (default to 0)
    options = options || {};
    var days = Number(options.days) || 0;
    var hours = Number(options.hours) || 0;
    var minutes = Number(options.minutes) || 0;
    var seconds = Number(options.seconds) || 0;

    // 3) Apply them using Date setters
    if (days) d.setDate(d.getDate() + days);
    if (hours) d.setHours(d.getHours() + hours);
    if (minutes) d.setMinutes(d.getMinutes() + minutes);
    if (seconds) d.setSeconds(d.getSeconds() + seconds);

    return d;
}