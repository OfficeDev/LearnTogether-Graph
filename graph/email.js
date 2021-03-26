import { getUser } from './user.js';
import graphClient from './graphClient.js';

export async function getMyUnreadEmails(userId) {

  let query = graphClient
    .api('/me/messages')
    .select('subject,bodyPreview,sender')
    .top(5);
  let filter = 'isRead eq false';

  if (userId) {
    const selectedUser = await getUser(userId);
    filter += ` and sender/emailAddress/address eq '${selectedUser.mail}'`;
  }

  return await query
    .filter(filter)
    .get();
}

