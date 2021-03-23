import { getTrendingFiles } from '../graph/files.js';

export async function loadTrendingFiles() {
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
  