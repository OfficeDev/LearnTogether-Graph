import { getSelectedUserId } from '../ui/colleagues.js';
import graphClient from './graphClient.js';

export async function getProfile() {
    const selectedUserId = getSelectedUserId();
    const userQueryPart = selectedUserId ? `/users/${selectedUserId}` : '/me';
  
    const profile = await graphClient
      .api(`${userQueryPart}`)
      .select('displayName,jobTitle,department,mail,aboutMe,city,state,country')
      .get();
  
    return profile;
  }