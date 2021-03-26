import { getMyUnreadEmails } from '../graph/email.js';
import { getSelectedUserId, getUser } from '../graph/user.js';

export async function loadUnreadEmails() {

  document.querySelector('#emails .loading').style = 'display: block';
  document.querySelector('#emails .noContent').style = 'display: none';
  document.querySelector('#emails ul').innerHTML = '';

  const selectedUserId = getSelectedUserId();

  const myUnreadEmails = await getMyUnreadEmails(selectedUserId);
  document.querySelector('#emails .loading').style = 'display: none';

  const emailsList = document.querySelector('#emails ul');
  emailsList.innerHTML = '';

  let selectedUser;
  if (!selectedUserId) {
    document.querySelector('#emails h2').innerHTML = 'Your unread emails';
  } else {
    selectedUser = await getUser(selectedUserId);
    document.querySelector('#emails h2').innerHTML = `Your unread emails from ${selectedUser.displayName}`;
  }
  if (myUnreadEmails.value.length === 0) {
    document.querySelector('#emails .noContent').style = 'display: block';
    return;
  }

  myUnreadEmails.value.forEach(email => {
    const emailLi = document.createElement('li');
    emailLi.className = 'ms-depth-8';

    const emailLine1 = document.createElement('div');
    emailLine1.className = 'ms-fontWeight-bold';
    emailLine1.innerHTML = email.sender.emailAddress.name;

    const emailLine2 = document.createElement('div');
    emailLine2.className = 'ms-fontWeight-bold ms-color-sharedCyanBlue20'
    emailLine2.innerHTML = email.subject;

    const emailLine3 = document.createElement('div');
    emailLine3.className = 'mt8';
    emailLine3.innerHTML = email.bodyPreview;

    emailLi.append(emailLine1, emailLine2, emailLine3);
    emailsList.append(emailLi);
  });
}