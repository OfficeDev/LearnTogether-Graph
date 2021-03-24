import { loadUnreadEmails } from './ui/email.js';
import { loadColleagues } from './ui/colleagues.js';
import { loadMeetings } from './ui/events.js';
import { loadTrendingFiles } from './ui/files.js';

import { getUser, getUserPhoto } from './graph/user.js';
import { signIn, silentSignIn } from './auth.js';

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

export async function loadData() {
    await Promise.all([
        loadUnreadEmails(),
        loadMeetings(),
        loadTrendingFiles(),
    ]);
}

// Expose for login button AND call immediately to attempt auto-login
window.displayUI = displayUI;
displayUI();