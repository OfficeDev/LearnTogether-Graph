import { getMyUpcomingMeetings } from '../graph/events.js';

export async function loadMeetings() {
    const myMeetings = await getMyUpcomingMeetings();
    document.querySelector('#events mgt-agenda').events = myMeetings;
    document.querySelector('#events .loading').style = 'display: none';

    if (myMeetings.length === 0) {
        document.querySelector('#events .noContent').style = 'display: block';
        return;
    }
}

function dayFromDateTime (dateTimeString) {

    let date = new Date(dateTimeString);
    date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
    let monthNames = [
        'January',
        'February',
        'March',
        'April',
        'May',
        'June',
        'July',
        'August',
        'September',
        'October',
        'November',
        'December'
    ];

    let monthIndex = date.getMonth();
    let day = date.getDate();
    let year = date.getFullYear();

    return monthNames[monthIndex] + ' ' + day + ' ' + year;
}

function timeRangeFromEvent (event) {
    if (event.isAllDay) {
        return 'ALL DAY';
    }

    let prettyPrintTimeFromDateTime = date => {
        date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
        let hours = date.getHours();
        let minutes = date.getMinutes();
        let ampm = hours >= 12 ? 'PM' : 'AM';
        hours = hours % 12;
        hours = hours ? hours : 12;
        let minutesStr = minutes < 10 ? '0' + minutes : minutes;
        return hours + ':' + minutesStr + ' ' + ampm;
    };

    let start = prettyPrintTimeFromDateTime(new Date(event.start.dateTime));
    let end = prettyPrintTimeFromDateTime(new Date(event.end.dateTime));

    return start + ' - ' + end;
}

