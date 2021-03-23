import { getMyColleagues } from '../graph/colleagues.js';
import { loadData } from '../index.js';
import { getSelectedUserId, setSelectedUserId } from '../graph/user.js';

export function selectPerson(personElement, personId) {
    const selectedUserId = getSelectedUserId();
    if (personElement && selectedUserId === personId) {
      personId = '';
      personElement = undefined;
      document.querySelector('#emails h2').innerHTML = 'Your unread emails';
      document.querySelector('#trending h2').innerHTML = 'Trending files';
      document.querySelector('#events h2').innerHTML = 'Your upcoming meetings next week';
    }
  
    setSelectedUserId(personId);
  
    // unselect all users
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
      document.querySelector('#events h2').innerHTML = `Your upcoming meetings next week with ${personName}`;  
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
  
  