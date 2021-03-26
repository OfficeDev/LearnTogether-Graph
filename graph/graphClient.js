import { getToken } from '../auth.js';
import { CacheMiddleware } from './CacheMiddleware.js';

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
middleware.unshift(new CacheMiddleware());

// Graph client singleton
const graphClient = MicrosoftGraph.Client.initWithMiddleware({ middleware });

export default graphClient;
