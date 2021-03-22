//#region Set up Graph provider

// Create an authentication provider
const authProvider = {
  getAccessToken: async () => {
    // Call getToken in auth.js
    return await getToken();
  }
};
// Initialize the Graph client
const graphClient = MicrosoftGraph.Client.initWithMiddleware({ authProvider });
//Get user info from Graph
async function getUser() {
  return await graphClient
    .api('/me')
    .select('id,displayName,jobTitle')
    .get();
}
//#endregion

//#region Colleagues

async function getMyColleagues() {
  // get my manager
  var manager;
  try {
    manager = await graphClient
      .api('/me/manager')
      .select('id')
      .get();
  }
  catch {
    throw 'Manager not found';
  }

  // get my colleagues
  const colleagues = await graphClient
    .api(`/users/${manager.id}/directReports`)
    .select('id,displayName,jobTitle,department,city,state,country')
    .get();

  // exclude the current user, since this is not supported in Graph, we need to
  // do it locally
  const accountName = sessionStorage.getItem('msalAccount');
  const account = msalClient.getAccountByUsername(accountName);
  const currentUserId = account.homeAccountId.substr(0, account.homeAccountId.indexOf('.'));
  colleagues.value = colleagues.value.filter(c => c.id !== currentUserId);

  // add a new property that combines job title and department and a placeholder
  // for timezone
  const timezonePromises = await Promise.allSettled(
    colleagues.value.map(
      colleague => getTimezoneInfo(colleague.city, colleague.state, colleague.country)));
  colleagues.value.forEach((colleague, i) => {
    colleague.jobTitleAndDepartment = `${colleague.jobTitle || ''} (${colleague.department || ''})`;
    colleague.localTime = [colleague.city, colleague.state, colleague.country].join(', ');

    const timezonePromise = timezonePromises[i];
    if (timezonePromise.status === 'fulfilled') {
      const localTime = new Date(timezonePromise.value.convertedTime.localTime);
      colleague.localTime = `${toShortTimeString(localTime)} (${timezonePromise.value.abbreviation}; ${colleague.localTime})`;
    }
  })

  // get colleagues' photos
  const colleaguesPhotosRequests = colleagues.value.map(
    colleague => getUserPhoto(colleague.id));
  const colleaguesPhotos = await Promise.allSettled(colleaguesPhotosRequests);
  colleagues.value.forEach((colleague, i) => {
    if (colleaguesPhotos[i].status === 'fulfilled') {
      colleague.personImage = URL.createObjectURL(colleaguesPhotos[i].value);
    }
  });

  return colleagues;
}

async function getUserPhoto(userId) {
  return graphClient
    .api(`/users/${userId}/photo/$value`)
    .get();
}

async function getTimezoneInfo(city, state, country) {
  const query = [city, state, country].join(', ');
  const result = await fetch(`https://dev.virtualearth.net/REST/v1/TimeZone/?query=${query}&key=${constants.bingMapsApiKey}`);
  if (result.ok) {
    const json = await result.json();
    return json.resourceSets[0].resources[0].timeZoneAtLocation[0].timeZone[0];
  }
  else {
    throw result.error;
  }
}

function toShortTimeString(date) {
  const timeString = date.toLocaleTimeString();
  const match = timeString.match(/(\d+\:\d+)\:\d+(.*)/);
  if (!match) {
    return timeString;
  }

  return `${match[1]}${match[2]}`;
}
//#endregion

//#region  Email
async function getMyUnreadEmails() {
  const selectedUserId = getSelectedUserId();

  let query = graphClient
    .api('/me/messages')
    .select('subject,bodyPreview,sender')
    .top(5);
  let filter = 'isRead eq false';

  if (selectedUserId) {
    const selectedUsersEmail = await getEmailForUser(selectedUserId);
    filter += ` and sender/emailAddress/address eq '${selectedUsersEmail}'`;
  }

  return await query
    .filter(filter)
    .get();
}

async function getEmailForUser(userId) {
  const user = await graphClient
    .api(`/users/${userId}`)
    .select('mail')
    .get();
  return user.mail;
}
//#endregion

//#region Meetings

//get calendar events for upcoming week
async function getMyUpcomingMeetings() {
  const selectedUserId = getSelectedUserId();
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
    const selectedUsersEmail = await getEmailForUser(selectedUserId);
    meetings = meetings.filter(meeting =>
      meeting.attendees.some(attendee => attendee.emailAddress.address === selectedUsersEmail));
  }
  //photos of attendees
  var photoRequests = [];
  meetings.forEach(meeting => {
    // get attendees' photos
    meeting.attendees.forEach(
      attendee => photoRequests.push(getUserPhoto(attendee.emailAddress.address)));
  });

  const attendeePhotos = await Promise.allSettled(photoRequests);
  meetings.forEach(meeting => {
    meeting.attendees.forEach((attendee, i) => {
      if (attendeePhotos[i].status === 'fulfilled') {
        attendee.personImage = URL.createObjectURL(attendeePhotos[i].value);
      }
    });
  });

  return meetings;
}
//#endregion

//#region Files
async function getTrendingFiles() {
  result = [];
  const selectedUserId = getSelectedUserId();
  const userQueryPart = selectedUserId ? `/users/${selectedUserId}` : '/me';

  const trendingIds = await graphClient
    .api(`${userQueryPart}/insights/trending`)
    .select('id')
    .filter("resourceReference/type eq 'microsoft.graph.driveItem'")
    .top(5)
    .get();

  if (trendingIds.value.length > 0) {
    let i = 1;
    const batchRequests = trendingIds.value.map(t => ({
      id: (i++).toString(),
      request: new Request(`${userQueryPart}/insights/trending/${t.id}/resource`,
        { method: "GET" })
    }));

    const batchContent = await (new MicrosoftGraph.BatchRequestContent(batchRequests)).getContent();
    const batchResponse = await graphClient
      .api('/$batch')
      .post(batchContent);
    const batchResponseContent = new MicrosoftGraph.BatchResponseContent(batchResponse);
    for (let j = 1; j < i; j++) {
      let response = await batchResponseContent.getResponseById(j.toString());
      if (response.ok) {
        result.push(await response.json());
      }
    }
  }
  return result;
}
//#endregion

//#region  Profile
async function getProfile() {
  const selectedUserId = getSelectedUserId();
  const userQueryPart = selectedUserId ? `/users/${selectedUserId}` : '/me';

  const profile = await graphClient
    .api(`${userQueryPart}`)
    .select('displayName,jobTitle,department,mail,aboutMe,city,state,country')
    .get();

  return profile;
}
//#endregion

