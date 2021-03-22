//#region Set up UI

async function displayUI(auto) {
  if (auto) {
    const loggedIn = await silentSignIn();
    if (!loggedIn) {
      return;
    }
  }
  else {
    await signIn();
  }

  // Display info from user profile
  const user = await getUser();

  try {
    const userPhoto = await getUserPhoto(user.id);
    user.personImage = URL.createObjectURL(userPhoto);
  }
  catch { }

  const myAvatar = document.querySelector('.card mgt-person');
  myAvatar.personDetails = user;

  var userName = document.getElementById('userName');
  userName.innerText = user.displayName;
  //Job Title
  var userJobComponent = document.getElementById('myJob');
  userJobComponent.innerHTML = user.jobTitle;
  // Hide login button and initial UI
  var signInButton = document.getElementById('signin');
  signInButton.style = "display: none";
  var content = document.getElementById('content');
  content.style = "display: block";

  await Promise.all([
    loadColleagues(),
    loadData()
  ]);
}

async function loadData() {
  await Promise.all([
    loadUnreadEmails(),
    loadMeetings(),
    loadTrendingFiles(),
    loadProfile()
  ]);
}

//#endregion

//#region Colleagues

function selectPerson(personElement, personId) {
  const selectedUserId = getSelectedUserId();
  if (personElement && selectedUserId === personId) {
    personId = '';
    personElement = undefined;
    document.querySelector('#emails h2').innerHTML = 'Your unread emails';
    document.querySelector('#trending h2').innerHTML = 'Trending files';
    document.querySelector('#events h2').innerHTML = 'Your upcoming meetings next week';
  }

  location.hash = `#${personId}`;

  // unselect all users
  document
    .querySelectorAll('#colleagues li.selected')
    .forEach(elem => elem.className = elem.className.replace('selected', ''));
  // remove profile
  const profile = document.getElementById('profile');
  if (profile) {
    profile.parentNode.removeChild(profile);
  }

  if (!personElement) {
    personElement = document.querySelector(`#colleagues li[data-personid="${personId}"]`);
  }

  const allColleaguesMapElement = document.querySelector('#allColleaguesMap');
  if (allColleaguesMapElement) {
    allColleaguesMapElement.style = personElement ? "display:none" : "display:inline";
  }
  if (personElement) {
    personElement.className += 'selected';
    const personName = personElement.dataset['personname'];
    document.querySelector('#emails h2').innerHTML = `Your unread emails from ${personName}`;
    document.querySelector('#trending h2').innerHTML = `Files trending around ${personName}`;
    document.querySelector('#events h2').innerHTML = `Your upcoming meetings next week with ${personName}`;

    const profileElement = document.createElement('div');
    profileElement.setAttribute('id', 'profile');
    profileElement.className = 'ms-motion-slideDownIn';
    profileElement.innerHTML = '<div></div>';
    personElement.parentNode.insertBefore(profileElement, personElement.nextSibling);
  }

  document.querySelector('#emails .loading').style = 'display: block';
  document.querySelector('#emails .noContent').style = 'display: none';
  document.querySelector('#emails ul').innerHTML = '';
  document.querySelector('#trending .loading').style = 'display: block';
  document.querySelector('#trending .noContent').style = 'display: none';
  document.querySelector('#trending ul').innerHTML = '';
  document.querySelector('#events .loading').style = 'display: block';
  document.querySelector('#events .noContent').style = 'display: none';
  document.querySelector('#events mgt-agenda').events = [];

  loadData();
}

async function loadColleagues() {
  const { myColleagues, mapUrl } = await getMyColleagues();
  document.querySelector('#colleagues .loading').style = 'display: none';

  const colleaguesList = document.querySelector('#colleagues ul');
  myColleagues.value.forEach(person => {
    const colleagueLi = document.createElement('li');
    colleagueLi.addEventListener('click', () => selectPerson(colleagueLi, person.id));
    colleagueLi.setAttribute('data-personid', person.id);
    colleagueLi.setAttribute('data-personname', person.displayName);

    const mgtPerson = document.createElement('mgt-person');
    mgtPerson.personDetails = person;
    mgtPerson.line2Property = 'jobTitleAndDepartment';
    mgtPerson.line3Property = 'localTime';
    mgtPerson.view = mgt.ViewType.threelines;

    colleagueLi.append(mgtPerson);
    colleaguesList.append(colleagueLi);
  });

  const selectedUserId = getSelectedUserId();
  if (selectedUserId) {
    // we need to put marking selected user on a different rendering loop to
    // wait for the DOM to refresh
    setTimeout(() => {
      selectPerson(undefined, selectedUserId);
    }, 1);
  }

  const mapLi = document.createElement('li');
  mapLi.setAttribute("id","allColleaguesMap");
  mapLi.style = selectedUserId ? "display:none" : "display:inline";
  mapImage = document.createElement('img');
  mapImage.setAttribute("src", mapUrl);
  mapImage.setAttribute("class", "map");
  mapLi.append(mapImage);
  colleaguesList.append(mapLi);

}

function getSelectedUserId() {
  if (location.hash.length < 2) {
    return undefined;
  }

  return location.hash.substr(1);
}
//#endregion

//#region Email

async function loadUnreadEmails() {
  const myUnreadEmails = await getMyUnreadEmails();
  document.querySelector('#emails .loading').style = 'display: none';

  const emailsList = document.querySelector('#emails ul');
  emailsList.innerHTML = '';

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

//#endregion

//#region Meetings
//load meetings to MGT agenda component
async function loadMeetings() {
  const myMeetings = await getMyUpcomingMeetings();
  document.querySelector('#events mgt-agenda').events = myMeetings;
  document.querySelector('#events .loading').style = 'display: none';

  if (myMeetings.length === 0) {
    document.querySelector('#events .noContent').style = 'display: block';
    return;
  }
}
//#endregion

//#region Files
async function loadTrendingFiles() {
  const trendingFiles = await getTrendingFiles();
  document.querySelector('#trending .loading').style = 'display: none';

  const trendingList = document.querySelector('#trending ul');
  trendingList.innerHTML = '';

  if (trendingFiles.length === 0) {

    document.querySelector('#trending .noContent').style = 'display: block';

  } else {

    trendingFiles.forEach(file => {

      const fileLi = document.createElement('li');
      fileLi.className = "ms-depth-8";
  
      const fileElement = document.createElement('mgt-file');
      fileElement.driveItem = file;
      fileLi.append(fileElement);
      fileLi.addEventListener('click', () => { window.open(file.webUrl,'_blank'); });
      trendingList.append(fileLi);
    });
    
  }
}
//#endregion

//#region Profile

async function loadProfile() {
  const profile = await getProfile();

  // Fill in the data
  const aboutMe = profile.aboutMe ? profile.aboutMe : "";
  const mapUrl = await getMapUrl([{city: profile.city, state: profile.state, country: profile.country}]);
  const profileDetail = document.querySelector('#profile div');
  html = `
      <table>
        <tr><td colspan="2">${aboutMe}</td></tr>
        <tr><td>Mail</td><td>${profile.mail}</td></tr>
        <tr><td colspan="2"><img class="map" src=${mapUrl} /></td></tr>
      </table>
    `;
  profileDetail.innerHTML = html;
}

//#endregion