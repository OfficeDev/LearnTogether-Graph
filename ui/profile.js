import { getProfile } from '../graph/user.js';
import { getMapUrl } from '../bingMaps.js';

export async function loadProfile() {
    const profile = await getProfile();
  
    // Fill in the data
    const aboutMe = profile.aboutMe ? profile.aboutMe : "";
    const mapUrl = await getMapUrl([{city: profile.city, state: profile.state, country: profile.country}]);
    const profileDetail = document.querySelector('#profile div');
    const html = `
        <table>
          <tr><td colspan="2">${aboutMe}</td></tr>
          <tr><td>Mail</td><td>${profile.mail}</td></tr>
          <tr><td colspan="2"><img class="map" src=${mapUrl} /></td></tr>
        </table>
      `;
    profileDetail.innerHTML = html;
  }
  
