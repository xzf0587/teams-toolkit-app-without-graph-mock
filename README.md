# Getting Started with Hello World Tab with Backend Sample (Azure)

This is a project based on Hello World Tab with Backend Sample for Teams Toolkit.
It is used to develop and test graph api request without real endpoint but leveraging the M365 proxy to mock response.
Try it in Windows system. For MacOX/linux, the m365proxy config should be updated as there is no system level proxy.

## Running Steps 
1. Use VSCode to open this project.
1. Install Teams Toolkit extension in VSCode extension if it is not installed.
<!-- 2. Create `.localConfigs` file in project folder and add an enter `REACT_APP_HOOK_GRAPH=true`. Use the flag to enable hooking method about graph api and teamsfx auth for react. -->
<!-- 3. Create another `.localConfigs` file in `api` folder and add an enter `HOOK_GRAPH=true`. Use the flag to enable hooking method about graph api and teamsfx auth for backend. -->
3. Click `F5` to launch the local debug.(Select Debug (Edge) option. It will be the default option)
   <br>
   It requires a M365 account to login and create a teams app and an aad app.
1. Install the Teams app in popup Edge window.
1. Run `npm run m365proxy` in command under project folder to launch the m365 proxy.
1. Try to test access to graph api from frontend and backend.

