import graphClient from './graphClient.js';
import { getAccount } from '../auth.js';
import { getUserPhoto } from './user.js';

export async function getMyColleagues() {
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
  const account = getAccount();
  const currentUserId = account.homeAccountId.substr(0, account.homeAccountId.indexOf('.'));
  colleagues.value = colleagues.value.filter(c => c.id !== currentUserId);

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
