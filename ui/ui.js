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
  