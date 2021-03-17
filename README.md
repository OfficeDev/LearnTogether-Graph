# graph-learn-together

## Minimal Path to Awesome

1. Create AAD web app
    Setting|Value
    -------|-----
    Name|My app
    Platform|Single-page application
    Redirect URIs|http://localhost
    Supported account types|Accounts in this organizational directory only (Single tenant)
    API permissions|Microsoft Graph User.Read (delegated)
1. Based on the `.env_sample` file, create a new file named `.env.js`
1. In the `.env.js` file update `clientId` and `authority` with the IDs from your AAD app
1. Run `npm start` to start the app