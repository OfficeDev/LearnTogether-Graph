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
    .select('id,displayName')
    .get();
}

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
    .get();

  // exclude the current user, since this is not supported in Graph, we need to
  // do it locally
  const accountName = sessionStorage.getItem('msalAccount');
  const account = msalClient.getAccountByUsername(accountName);
  const currentUserId = account.homeAccountId.substr(0, account.homeAccountId.indexOf('.'));
  colleagues.value = colleagues.value.filter(c => c.id !== currentUserId);

  // get colleagues' photos
  const colleaguesPhotosRequests = colleagues.value.map(
    colleague => graphClient
      .api(`/users/${colleague.id}/photo/$value`)
      .get());
  const colleaguesPhotos = await Promise.allSettled(colleaguesPhotosRequests);
  colleagues.value.forEach((colleague, i) => {
    if (colleaguesPhotos[i].status === 'fulfilled') {
      colleague.personImage = URL.createObjectURL(colleaguesPhotos[i].value);
    }
  });

  return colleagues;
}

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

//get calendar events for upcoming week
async function getMyUpcomingMeetings() {
  const dateNow = new Date();
  const dateNextWeek = new Date();
  dateNextWeek.setDate(dateNextWeek.getDate() + 7);
  return await graphClient
    .api(`/me/calendar/calendarView`)
    .query(`startDateTime=${dateNow.toISOString()}&endDateTime=${dateNextWeek.toISOString()}`)
    .orderby(`start/DateTime`)
    .get();
}

async function getTrendingFiles() {
  const selectedUserId = getSelectedUserId();
  const userQueryPart = selectedUserId ? `/users/${selectedUserId}` : '/me';

  const trendingIds = await graphClient
    .api(`${userQueryPart}/insights/trending`)
    .select('id')
    .top(5)
    .get();
  const trendingResponses = await Promise.allSettled(trendingIds.value.map(t =>
    graphClient
      .api(`${userQueryPart}/insights/trending/${t.id}/resource`)
      .get()));
  return trendingResponses
    .filter(res => res.status === 'fulfilled')
    .map(res => res.value);
}

async function getEmailForUser(userId) {
  const user = await graphClient
    .api(`/users/${userId}`)
    .select('mail')
    .get();
  return user.mail;
}