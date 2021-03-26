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
middleware.unshift(new CacheMiddleware([
  // email messages might change so let's refresh them once in a while
  {
    path: '/messages',
    expirationInMinutes: 5
  },
  // calendarView queries contain current date and time with seconds so no
  // point in caching them
  {
    path: '/calendarView',
    expirationInMinutes: 0
  }
]));

// Graph client singleton
const graphClient = MicrosoftGraph.Client.initWithMiddleware({ middleware });

export default graphClient;
