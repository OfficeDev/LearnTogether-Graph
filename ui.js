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
  const userPhoto=await getUserPhoto(user.id);
  var userName = document.getElementById('userName');
  userName.innerText = user.displayName;
  //profile photo
  var userPhotoComponent = document.getElementById('myProfileImage');
  userPhotoComponent.src = URL.createObjectURL(userPhoto);
  //Job Title
  var userJobComponent= document.getElementById('myJob');
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
    loadTrendingFiles()
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
  }

  location.hash = `#${personId}`;

  document
    .querySelectorAll('#colleagues li.selected')
    .forEach(elem => elem.className = elem.className.replace('selected', ''));

  if (!personElement) {
    personElement = document.querySelector(`#colleagues li[data-personid="${personId}"]`);
  }

  if (personElement) {
    personElement.className += 'selected';
    const personName = personElement.dataset['personname'];
    document.querySelector('#emails h2').innerHTML = `Your unread emails from ${personName}`;
    document.querySelector('#trending h2').innerHTML = `Files trending around ${personName}`;
    document.querySelector('#events h2').innerHTML = `You upcoming meetings next week with ${personName}`;
  }

  document.querySelector('#emails .loading').style = 'display: block';
  document.querySelector('#trending .loading').style = 'display: block';
  document.querySelector('#emails ul').innerHTML = '';
  document.querySelector('#trending ul').innerHTML = '';


  loadData();
}

async function loadColleagues() {
  const myColleagues = await getMyColleagues();
  const colleaguesLoading = document.querySelector('#colleagues .loading');
  colleaguesLoading.style = 'display: none';

  const colleaguesList = document.querySelector('#colleagues ul');
  myColleagues.value.forEach(person => {
    const colleagueLi = document.createElement('li');
    colleagueLi.addEventListener('click', () => selectPerson(colleagueLi, person.id));
    colleagueLi.setAttribute('data-personid', person.id);
    colleagueLi.setAttribute('data-personname', person.displayName);

    const mgtPerson = document.createElement('mgt-person');
    mgtPerson.personDetails = person;
    mgtPerson.line2Property = 'jobTitle';
    mgtPerson.view = mgt.ViewType.twolines;

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
  const emailsLoading = document.querySelector('#emails .loading');
  emailsLoading.style = 'display: none';

  const emailsList = document.querySelector('#emails ul');
  emailsList.innerHTML = '';

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
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}
//#region Meetings
//load meetings to MGT agenda component
async function loadMeetings() {
  const myMeetings = await getMyUpcomingMeetings();
  const meetingsComponent = document.getElementById('myMeetings');
  const noMeetingMessage=document.getElementById('noMeetingMsg');
  
  if(myMeetings.length>0){ 
    noMeetingMessage.style = 'display: none';
    meetingsComponent.events = myMeetings;
  }else{
  noMeetingMessage.style = 'display: block';
  meetingsComponent.events=[];
 
  }  
}
//#endregion

//#region Files
async function loadTrendingFiles() {
  const trendingFiles = await getTrendingFiles();
  const trendingLoading = document.querySelector('#trending .loading');
  trendingLoading.style = 'display: none';

  const trendingList = document.querySelector('#trending ul');
  trendingList.innerHTML = '';

  let html = "";

  trendingFiles.forEach(file => {

    const name = file.name;
    const extension = file.name.split('.').slice(-1)[0].toLowerCase();
    iconUrl = ['docx','xlsx','pptx','vsdx','msg','mpp'].includes(extension) ?
              `https://static2.sharepointonline.com/files/fabric/assets/item-types/48_1.5x/${extension}.svg` :
              "data:image/gif;base64,R0lGODlhAQABAIAAAP///wAAACH5BAEAAAAALAAAAAABAAEAAAICRAEAOw=="; // single blank pixel
    const modified = new Date (file.lastModifiedDateTime).toLocaleDateString();
    const size = (file.size / 1000000).toFixed(2) + "MB";
    const webUrl = file.webUrl;

    html += `
      <li class="ms-depth-8">
        <div class="file" onclick="window.open('${webUrl}','_blank');">
          <div class="icon"><img src="${iconUrl}"></div>
          <div class="detail">
            <div class="line1">${name}</div>
            <div class="line2">Modified ${modified}</div>
            <div class="line3">Size: ${size}</div>
          </div>
        </div>
      </li>`;
  });
  trendingList.innerHTML = html;
}
//#endregion

