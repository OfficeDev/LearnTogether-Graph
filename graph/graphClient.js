import { getToken } from '../auth.js';

// Create an authentication provider
function AuthProvider() {
  nextMiddleware: undefined;

  return {
    getAccessToken: async () => {
      // Call getToken in auth.js
      return await getToken();
    },
    setNext: (next) => {
      this.nextMiddleware = next;
    }
  }
};

const middleware = MicrosoftGraph.MiddlewareFactory.getDefaultMiddlewareChain(new AuthProvider());

function CacheMiddleware() {
  nextMiddleware: undefined;

  return {
    execute: async (context) => {
      const requestKey = btoa(context.request);
      const response = window.localStorage.getItem(requestKey);
      if (response) {
        const resp = JSON.parse(response);
        context.response = new Response(resp.body, resp);
      }
      else {
        await this.nextMiddleware.execute(context);
        const resp = context.response.clone();
        console.log(resp.headers.get());
        const response = {
          url: resp.url,
          status:  resp.status,
          statusText:  resp.statusText,
          headers: resp.headers,
          body: await resp.json()
        };
        window.localStorage.setItem(requestKey, JSON.stringify(response));
      }
    },
    setNext: (next) => {
      this.nextMiddleware = next;
    }
  }
};

middleware.unshift(new CacheMiddleware());

// Graph client singleton
const graphClient = MicrosoftGraph.Client.initWithMiddleware({ middleware });

export default graphClient;
