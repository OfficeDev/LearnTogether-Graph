import { getTrendingFiles } from '../graph/files.js';
import { getSelectedUserId, getUser } from '../graph/user.js';

export async function loadTrendingFiles() {
  const selectedUserId = getSelectedUserId();
  if (!selectedUserId) {
    document.querySelector('#trending h2').innerHTML = 'Trending files';
  } else {
    const selectedUser = await getUser(selectedUserId);
    document.querySelector('#trending h2').innerHTML = `Files trending around ${selectedUser.displayName}`;
  }

  document.querySelector('#trending .loading').style = 'display: block';

  const trendingFiles = await getTrendingFiles(selectedUserId);

  document.querySelector('#trending .loading').style = 'display: none';
  document.querySelector('#trending .noContent').style = 'display: none';
  document.querySelector('#trending ul').innerHTML = '';

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
      fileLi.addEventListener('click', () => { window.open(file.webUrl, '_blank'); });
      trendingList.append(fileLi);
    });

  }
}
