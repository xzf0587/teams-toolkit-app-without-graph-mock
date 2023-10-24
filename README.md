# Getting Started with Hello World Tab with Backend Sample (Azure)

This is a project updated from Hello World Tab with Backend Sample for Teams Toolkit.
It is modified to develop and test graph api request without real endpoint but leveraging the M365 proxy to mock response.

## Setup 
1. Download the M365 Proxy project.
1. Use VSCode to open this project.
1. Install Teams Toolkit extension in VSCode extension.
1. Click `F5` to launch the local debug. (Select Debug (Edge) option. It will be default option)
   <br>
   It requires a M365 account to login and create a teams app and an aad app.
1. Install the Teams app in popup Edge window.
1. Launch M365 proxy in M365 folder by command `./m365proxy.exe  -f 0 --mocks-file ./responses.sample.json --watch-process-names node msedge`.
1. Try to test access to graph api from frontend and backend.

