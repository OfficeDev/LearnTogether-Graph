# Learn Together Graph - Application Sample

## Minimal Path to Awesome

1. Ensure [NodeJS](https://nodejs.org) is installed on your computer. Note that NodeJS is only used to run a small HTTP server; there is no build step required and the application code runs entirely in the user's web browser.

1. Register an application in Azure Active Directory and save your clientId and tenant Id from the registration.

    Setting|Value
    -------|-----
    Name|My app
    Platform|Single-page application
    Redirect URIs|http://localhost
    Supported account types|Accounts in this organizational directory only (Single tenant)
    API permissions|Microsoft Graph User.Read (delegated)

1. Obtain a [Bing Maps API key](https://docs.microsoft.com/en-us/bingmaps/getting-started/bing-maps-dev-center-help/getting-a-bing-maps-key?WT.mc_id=m365-16105-cxa)

1. Based on the `.env_sample` file, create a new file named `.env.js`

1. In the `.env.js` file update the following values:

   Take the value from | Paste it into your .env.js as
   --------------------|------------------------------
   App registration application (client) ID | clientId
   App registration tenant (directory) ID | authority
   Bing maps key | bingMapsApiKey

1. Run `npm start` to start the app

## Branches

The main branch includes all program features as well as Graph call caching.
Branches are also included that correspond to the "scenes" in the Learn Together - Graph presentation:

 - 01-StartWithColleaguesAndPictures - shows a list of the user's colleagues and their pictures
 - 02-AddEvents - adds the next week's events for the current user
 - 03-AddPersonNavigation - adds simple navigation by clicking on colleagues; filters events to show those which include the selected colleague
 - 04-AddFilteredEmail - adds email for the current user, or email from a colleague if one is selected
 - 05-AddTrendingFiles - adds display of trending files for the current user, or for a colleague if one is selected
 - 06-Final - adds time zones and maps courtesy of the Bing Maps service

## Interesting concepts

This application illustrates the following concepts (some are missing in certain branches; main branch has all of them):

   Concept | Shown in file(s)
   --------|-----------------
   ES6 Modules | all .js files
   Get colleagues from Microsoft Graph | graph/colleagues.js
   Get email from Microsoft Graph, optionally filter by sender | graph/email.js
   Get calendar events from Microsoft Graph across a date range | graph/events.js
   Get trending files from Microsoft Graph | graph/files.js
   Batch requests in Microsoft Graph | graph/files.js
   Set up Graph client with MSAL | auth.js, graph/graphClient.js
   Cache Graph results in browser storage | graph/graphClient.js, graph/CacheMiddleware.js
   Get timezone and map based on user's city | bingMaps.js

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).

Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.