import { getMapUrl } from '../ui/profile.js';
import graphClient from './graphClient.js';
import { getAccount } from '../auth.js';

//#region Colleagues

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

  // plot everyone on a map
  const places = colleagues.value.map(
    colleague => ({ city: colleague.city, state: colleague.state, country: colleague.country })
  );
  const mapUrl = await getMapUrl(places);

  // exclude the current user, since this is not supported in Graph, we need to
  // do it locally
  const account = getAccount();
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

  return ({ myColleagues: colleagues, mapUrl: mapUrl });
}

//Get user info from Graph
export async function getUser() {
  return await graphClient
    .api('/me')
    .select('id,displayName,jobTitle')
    .get();
}

export async function getUserPhoto(userId) {
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
