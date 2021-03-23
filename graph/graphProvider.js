// Create an authentication provider
const authProvider = {
    getAccessToken: async () => {
      // Call getToken in auth.js
      return await getToken();
    }
  };
  // Initialize the Graph client
  const graphClient = MicrosoftGraph.Client.initWithMiddleware({ authProvider });
