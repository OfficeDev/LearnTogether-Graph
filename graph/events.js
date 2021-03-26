import { getUserPhoto, getUser } from './user.js';
import graphClient from './graphClient.js';

//get calendar events for upcoming week
export async function getMyUpcomingMeetings(selectedUserId) {

  const dateNow = new Date();
  const dateNextWeek = new Date();
  dateNextWeek.setDate(dateNextWeek.getDate() + 7);
  const query = `startDateTime=${dateNow.toISOString()}&endDateTime=${dateNextWeek.toISOString()}`;
  var meetings = [];
  const response = await graphClient
    .api(`/me/calendar/calendarView`)
    .query(query)
    .orderby(`start/DateTime`)
    .get();
  meetings = response.value;
  //if filter is applied, select meeting you have with the selected Colleague.
  if (selectedUserId) {
    const selectedUser = await getUser(selectedUserId);
    meetings = meetings.filter(meeting =>
      meeting.attendees.some(attendee => attendee.emailAddress.address === selectedUser.mail));
  }
  //photos of attendees
  var photoRequests = [];
  meetings.forEach(meeting => {
    // get attendees' photos
    meeting.attendees.forEach(
      attendee => photoRequests.push(getUserPhoto(attendee.emailAddress.address)));
  });
  const attendeePhotos = await Promise.allSettled(photoRequests);
  var count = 0;
  meetings.forEach(meeting => {
    meeting.attendees.forEach((attendee) => {
      if (count < attendeePhotos.length) {
        attendee.personImage = attendeePhotos[count].status === 'fulfilled' ? URL.createObjectURL(attendeePhotos[count].value) : null;
        count += 1;
      }
    });
  });
  return meetings;
}
