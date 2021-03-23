import graphClient from './graphClient.js';

export function getSelectedUserId() {
  if (location.hash.length < 2) {
    return undefined;
  }

  return location.hash.substr(1);
}

export function setSelectedUserId(id) {
  selectedUserId = id;
}

export async function getProfile() {
  const selectedUserId = getSelectedUserId();
  const userQueryPart = selectedUserId ? `/users/${selectedUserId}` : '/me';

  const profile = await graphClient
    .api(`${userQueryPart}`)
    .select('displayName,jobTitle,department,mail,aboutMe,city,state,country')
    .get();

  return profile;
}