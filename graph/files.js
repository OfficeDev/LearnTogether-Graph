import { getSelectedUserId } from './user.js';
import graphClient from './graphClient.js';

export async function getTrendingFiles() {
    const result = [];
    const selectedUserId = getSelectedUserId();
    const userQueryPart = selectedUserId ? `/users/${selectedUserId}` : '/me';
  
    const trendingIds = await graphClient
      .api(`${userQueryPart}/insights/trending`)
      .select('resourceReference')
      .filter("resourceReference/type eq 'microsoft.graph.driveItem'")
      .top(5)
      .get();
  
    if (trendingIds.value.length > 0) {
      let i = 1;
      const batchRequests = trendingIds.value.map(t => ({
        id: (i++).toString(),
        request: new Request(t.resourceReference.id,
          { method: "GET" })
      }));
  
      const batchContent = await (new MicrosoftGraph.BatchRequestContent(batchRequests)).getContent();
      const batchResponse = await graphClient
        .api('/$batch')
        .post(batchContent);
      const batchResponseContent = new MicrosoftGraph.BatchResponseContent(batchResponse);
      for (let j = 1; j < i; j++) {
        let response = await batchResponseContent.getResponseById(j.toString());
        if (response.ok) {
          result.push(await response.json());
        }
      }
    }
    return result;
  }