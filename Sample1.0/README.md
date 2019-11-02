# Microsoft Graph Search API Sample for SPFx

## Table of contents

* [Introduction](#introduction)
* [Prerequisites](#prerequisites)
* [Getting started with the sample](#getting-started-with-the-sample)
* [Build and run the sample](#build-and-run-the-sample)
* [Code of note](#code-of-note)
* [Questions and comments](#questions-and-comments)
* [Contributing](#contributing)
* [Additional resources](#additional-resources)

## Introduction

This sample demonstrates how to make calls to the Microsoft Graph Search API using SPFx web parts.

## Prerequisites

This sample requires the following:  

  * [Visual Studio Code](TODO) 
  * [Microsoft work or school account](https://www.outlook.com) 

## Getting started with the sample

 1. Download or clone this repo.
 2. Open Visual Studio code to the **/Sample1.0/spfx** directory
 3. Run the following commands to setup your development environment

 ```Javascript
npm install -g yo gulp
npm install -g @microsoft/generator-sharepoint
 ```
 
### Create an Azure AD Application

Follow the steps in [Configuring Azure](./ConfigureAzure.md).

## Test the Web Part (Local)

1.  In the debug console, type the following:

```javascript
cd spfx/sample3.0

gulp trust-dev-cert

gulp serve
```

2. A window will open to the local workbench, click the **+** sign.

3.  Select the Graph Search API web part

![Add the webpart to the workbench](./media/01_Workbench.png 'Add the web part')

4.  Click the edit icon

![Edit the web part](./media/02_EditWebpart.png 'Enter edit mode')

5.  Select **MSGraphClient**

![Set the auth mode](./media/03_SetAuthMode.png 'Set the auth mode')

5.  Type a search term, then click **Search**

## Deploy and Assign Permissions (SharePoint Online)

1.  Switch to Visual Studio Code and the terminal window

1.  Press **Ctrl-C** to stop debugging, press **Y** and **ENTER**

1.  In the debug console, type the following:

```javascript
gulp package-solution
```

2. Open a new browser window to your (SharePoint Online Administration site)[https://YOURTENANT-admin.sharepoint.com]

3.  Click **Active Sites**, ensure that you have an app catalog template site created

>NOTE:  If you do not have one, you will need to create one and wait for 20-30 minutes for it to completely provision.  Failure to do so will require you to deploy your web part several times.

4.  Open the App Catalog site

![Open your active sites](./media/04_ActiveSites.png 'Open active sites')

5.  In the App Catalog site, select **Apps for SharePoint**

6.  Click **Upload**

![Open Apps for SharePoint](./media/05_AppsForSharepoint.png 'Select apps for sharepoint and upload the web part')

7.  Browse to the **Sample1.0/spfx/sample-3/sharepoint/solution** directory and select the **sample-3.sppkg** file

8.  Click **Open**, then select **OK**

9.  If prompted, click **Deploy** in the trust dialog

![Trust and deploy the web part](./media/06_TrustDialog.png 'Trust the web part')

10.  Switch back to the SharePoint Online admin center, click **API Management**

>NOTE:  API Management will not display until your App Catalog has been created and the backend has converged (~30 minutes).

![Approve the permissions](./media/07_ApiApproval.png 'Approve the permissions')

11.  Approve all the permissions that your application has requested.

##  Add web part to a page

1.  Open a SharePoint site, then click **Site Contents**

1.  Click **New->App**

![Add an app](./media/08_AddApp.png 'Add an app')

1.  Select **From your organization**, then select the **sample-3** web part app

![Add the sample web part](./media/09_AddSample.png 'Add an app')

1.  The web part should now be installed on your site:

![Web part installed](./media/10_AppInstalled.png 'Web part is installed')

1.  In quick launch, select the **Pages** document library

2.  Click **+New->Web Part Page** in the menu

![Add a web part](./media/08_CreateWPP.png 'Add a web part')

3.  For the name, type **GraphSearch**

4.  Select the **Site Pages** library, then click **Create**

4.  In the **Header** web part zone, click **Add a Web Part**

5. Select the **Other** category, and then select the **GraphSearchAPI** web part, then click **Add**

![Add the web part](./media/11_AddWebPart.png 'Add the web part')

6.  In the ribbon, click **Edit page**

7.  In the web part drop down, select **Edit web part**, then click **Configure**

8.  Select **MSGraphClient** and then close the property window

9.  Click **OK** to save the web part properties

10.  Type a search term, then click **Search**, you will see the Graph Search API results display in the area below the web part.

![Perform a search and review the results](./media/12_SearchResults.png 'Perform a search')

## Code of note

- The **package-solution.json** file contains the permissions that are needed for your web part.
- The **GraphSearchApi.tsx** file is the main code and html for the web part

## Questions and comments

We'd love to get your feedback about this sample! 
Please send us your questions and suggestions in the [Issues](https://github.com/microsoftgraph/aspnet-connect-rest-sample/issues) section of this repository.

Your feedback is important to us. Connect with us on [Stack Overflow](https://stackoverflow.com/questions/tagged/microsoftgraph).
Tag your questions with [MicrosoftGraph].

## Contributing ##

If you'd like to contribute to this sample, see [CONTRIBUTING.md](CONTRIBUTING.md).

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). 
For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Additional resources

- [Microsoft Graph Security API Documentaion](https://aka.ms/graphsecuritydocs)
- [Other Microsoft Graph Connect samples](https://github.com/MicrosoftGraph?utf8=%E2%9C%93&query=-Connect)
- [Microsoft Graph overview](https://graph.microsoft.io)
- [Consume the Microsoft Graph in the SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/use-aad-tutorial)

## Copyright
Copyright (c) 2019 Microsoft. All rights reserved.
