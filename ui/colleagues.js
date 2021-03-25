import { getMyColleagues } from '../graph/colleagues.js';

  export async function loadColleagues() {
    const myColleagues = await getMyColleagues();
    document.querySelector('#colleagues .loading').style = 'display: none';
  
    const colleaguesList = document.querySelector('#colleagues ul');
    myColleagues.value.forEach(person => {
      const colleagueLi = document.createElement('li');
      colleagueLi.innerHTML = `${person.displayName}`;
      colleaguesList.append(colleagueLi);
    });
  }
  
  