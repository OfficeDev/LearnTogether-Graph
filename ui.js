async function displayUI(auto) {
  if (auto) {
    await silentSignIn();
  }
  else {
    await signIn();
  }

  // Display info from user profile
  const user = await getUser();
  var userName = document.getElementById('userName');
  userName.innerText = user.displayName;

  // Hide login button and initial UI
  var signInButton = document.getElementById('signin');
  signInButton.style = "display: none";
  var content = document.getElementById('content');
  content.style = "display: block";

  await Promise.all([
    loadColleagues(),
    loadUnreadEmails(),
    loadMeetings()
  ]);
}

function selectPerson(personId) {
  location.hash = `#${personId}`;
}

async function loadColleagues() {
  const myColleagues = await getMyColleagues();
  const colleaguesComponent = document.getElementById('myColleagues');
  colleaguesComponent.people = myColleagues.value;
}

async function loadUnreadEmails() {
  const myUnreadEmails = await getMyUnreadEmails();
  const emailsLoading = document.querySelector('#emails .loading');
  emailsLoading.style = 'display: none';

  const emailsList = document.querySelector('#emails ul');
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
//load meetings to MGT agenda component
async function loadMeetings() {
  const myMeetings = await getMyUpcomingMeetings();
  const meetingsComponent = document.getElementById('myMeetings');
  meetingsComponent.events = myMeetings.value;
}