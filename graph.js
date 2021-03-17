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
  ensureScope('user.read');
  return await graphClient
    .api('/me')
    .select('id,displayName')
    .get();
}

async function getMyColleagues() {
  ensureScope('User.Read.All');
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