async function displayUI() {    
  await signIn();

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
    loadUnreadEmails()
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
  const emailsList = document.querySelector('#emails dl');
  myUnreadEmails.value.forEach(email => {
    const emailDt = document.createElement('dt');
    emailDt.innerHTML = `${email.subject} (${email.sender.emailAddress.name})`;
    const emailDd = document.createElement('dd');
    emailDd.innerHTML = email.bodyPreview;
    emailsList.append(emailDt, emailDd);
  });
}