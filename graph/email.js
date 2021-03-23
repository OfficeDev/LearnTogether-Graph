import { getSelectedUserId } from './user.js';
import graphClient from './graphClient.js';

export async function getMyUnreadEmails() {
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

export async function getEmailForUser(userId) {
  const user = await graphClient
    .api(`/users/${userId}`)
    .select('mail')
    .get();
  return user.mail;
}
