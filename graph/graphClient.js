// Create an authentication provider
const authProvider = {
    getAccessToken: async () => {
      // Call getToken in auth.js
      return await getToken();
    }
  };

  // Graph client singleton
const graphClient = MicrosoftGraph.Client.initWithMiddleware({ authProvider });

export default graphClient;
