import { getMyUpcomingMeetings } from '../graph/events.js';
import { getSelectedUserId, getUser } from '../graph/user.js';

export async function loadMeetings() {

    document.querySelector('#events .loading').style = 'display: block';
    document.querySelector('#events .noContent').style = 'display: none';
    document.querySelector('#events mgt-agenda').events = [];

    const selectedUserId = getSelectedUserId();
    if (!selectedUserId) {
        document.querySelector('#events h2').innerHTML = 'Your upcoming meetings next week';
    } else {
        let selectedUser = await getUser(selectedUserId);
        document.querySelector('#events h2').innerHTML = `Your upcoming meetings next week with ${selectedUser.displayName}`;
    }

    const myMeetings = await getMyUpcomingMeetings(selectedUserId);
    document.querySelector('#events mgt-agenda').events = myMeetings;
    document.querySelector('#events .loading').style = 'display: none';

    if (myMeetings.length === 0) {
        document.querySelector('#events .noContent').style = 'display: block';
        return;
    }
}

// Function to format the meeting time - needs to be global for access from
// the MGT template in index.html
window.timeRangeFromEvent = function (event) {

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

