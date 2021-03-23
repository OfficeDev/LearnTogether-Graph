import graphClient from './graphClient.js';

export function getSelectedUserId() {
  if (location.hash.length < 2) {
    return undefined;
  }

  return location.hash.substr(1);
}

export function setSelectedUserId(id) {
  location.hash = `#${id}`;
}

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

export async function getEmailForUser(userId) {
  const user = await graphClient
    .api(`/users/${userId}`)
    .select('mail')
    .get();
  return user.mail;
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