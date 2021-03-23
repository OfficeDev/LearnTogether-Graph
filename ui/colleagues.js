import { getMyColleagues } from '../graph/colleagues.js';

  export async function loadColleagues() {
    const myColleagues = await getMyColleagues();
    document.querySelector('#colleagues .loading').style = 'display: none';
  
    const colleaguesList = document.querySelector('#colleagues ul');
    myColleagues.value.forEach(person => {
      const colleagueLi = document.createElement('li');
      colleagueLi.addEventListener('click', () => selectPerson(colleagueLi, person.id));
      colleagueLi.setAttribute('data-personid', person.id);
      colleagueLi.setAttribute('data-personname', person.displayName);
  
      const mgtPerson = document.createElement('mgt-person');
      mgtPerson.personDetails = person;
      mgtPerson.view = mgt.ViewType.threelines;
  
      colleagueLi.append(mgtPerson);
      colleaguesList.append(colleagueLi);
    });
  }
  
  