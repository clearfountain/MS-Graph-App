[![Hack Together: Microsoft Graph and .NET](https://img.shields.io/badge/Microsoft%20-Hack--Together-orange?style=for-the-badge&logo=microsoft)](https://github.com/microsoft/hack-together)

# MS-Graph-App

An app for aggregating files from OneDrive for all users in an organization, the files can then be sent as email attachments to users within and outside the organization. 

This app was built with the **MS-Graph .NET SDK** hosted on an Azure tenant and utilizes **app-only authentication**.

## How To Setup And Run
- Clone this repo.
- Run the following command in the project root folder > dotnet user-secrets set settings:clientSecret yft8Q~Npc5zsJaPC_j_Kbn_dMFMVc.GjZQCL~biY
- Start in Visual Studio
- Select user 3 when prompted (only user with non-empty drive as at the time of writing)
