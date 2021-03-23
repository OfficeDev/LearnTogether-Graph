import { getMyColleagues } from '../graph/colleagues.js';
import { loadData } from '../index.js';
import { getSelectedUserId, setSelectedUserId } from '../graph/user.js';

export function selectPerson(personElement, personId) {
    const selectedUserId = getSelectedUserId();
  
    setSelectedUserId(personId);
  
    // unselect all users
    document
      .querySelectorAll('#colleagues li.selected')
      .forEach(elem => elem.className = elem.className.replace('selected', ''));
  
    if (!personElement) {
      personElement = document.querySelector(`#colleagues li[data-personid="${personId}"]`);
    }
  
    document.querySelector('#events .loading').style = 'display: block';
    document.querySelector('#events .noContent').style = 'display: none';
    document.querySelector('#events mgt-agenda').events = [];
  
    loadData();
  }
  
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
  
    const selectedUserId = getSelectedUserId();
    if (selectedUserId) {
      // we need to put marking selected user on a different rendering loop to
      // wait for the DOM to refresh
      setTimeout(() => {
        selectPerson(undefined, selectedUserId);
      }, 1);
    }  
  }
  
  