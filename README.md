# WebViewer - Sharepoint integration sample

This sample serves as a guide to integrate your hosted WebViewer app with a Sharepoint document library. It showcases how you can leverage a control block to send file metadata to your WebViewer environment, which then opens the Sharepoint file using an authenticated request.
## Prerequisites

- (Optional but recommended) [Node Version Manager](http://npm.github.io/installation-setup-docs/installing/using-a-node-version-manager.html)
- Set up on Windows environment highly recommended (if you are not using a Windows machine, we recommend using [PnP PowerShell](https://pnp.github.io/powershell/index.html))

## For step-by-step help on setting up a SharePoint development environment, see one of the following:

* [Sharepoint Extension](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/building-simple-cmdset-with-dialog-api)\
* [Set up your Microsoft 365 Tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)\
* [Set up your SharePoint Framework Development Environment](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)\
* [Create your team site in Sharepoint](https://support.microsoft.com/en-us/office/create-a-team-site-in-sharepoint-ef10c1e7-15f3-42a3-98aa-b5972711777d)

## Project Description

![](https://pdftron.s3.amazonaws.com/custom/test/jack/sharepoint_readme_pics/Screen+Shot+2022-03-14+at+1.39.40+PM.png)
In the github repo there are two individual projects.

   <ol>
      <li>client</li>
      <li>sharepoint-extension</li>
   </ol>
   * <strong>Client</strong> is the project which will host PDFTron WebViewer - this is a representation of your cloud-hosted WebViewer instance or your local development environment. <br>
   * <strong>Sharepoint-extension</strong> is Sharepoint Extension which is called            
   
   [Edit Control Block](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/guidance/migrate-from-ecb-to-spfx-extensions). This extension allows us to add a button to any List View in Sharepoint, such as your document libraries. We add `Open in PDFTron` as a menu option to open and process any document in WebViewer.
   ![](https://pdftron.s3.amazonaws.com/custom/test/jack/sharepoint_readme_pics/Screen+Shot+2022-03-14+at+2.02.38+PM.png)

### Prepare your Sharepoint Tenant

1. Clone GitHub repo and then open the project in your development environment.

2. Open _sharepoint-extension/ExportToDocCommandSet.ts_ and set the WebViewer url in the file - it defaults to your local WebViewer dev environment (`const pdftronUrl = 'http://localhost:3000'`). Change this URL to your hosting environment if needed.

3. Secondly, you need to set **pageUrl** in _config/serve.json_, which should point at your document library's Sharepoint URL. After you are done, your `serve.json` file should look similar to this:
```
   {
      "$schema": "https://developer.microsoft.com/json-schemas/core-build/serve.schema.json",
      "port": 4321,
      "https": true,
      "serveConfigurations": {
         "default": {
            "pageUrl": "https://jingjackhou.sharepoint.com/sites/jingjackhou/Shared%20Documents/Forms/AllItems.aspx",
            "customActions": {
               "bf232d1d-279c-465e-a6e4-359cb4957377": {
                  "location": "ClientSideExtension.ListViewCommandSet.CommandBar",
                  "properties": {
                     "sampleTextOne": "One item is selected in the list",
                     "sampleTextTwo": "This command is always visible."
                  }
               }
            }
         },
         "helloWorld": {
            "pageUrl": "https://sppnp.sharepoint.com/sites/Group/Lists/Orders/AllItems.aspx",
            "customActions": {
               "bf232d1d-279c-465e-a6e4-359cb4957377": {
                  "location": "ClientSideExtension.ListViewCommandSet.CommandBar",
                  "properties": {
                     "sampleTextOne": "One item is selected in the list",
                     "sampleTextTwo": "This command is always visible."
                  }
               }
            }
         }
      }
   }
```

4. <strong>Important</strong>: For the Sharepoint Extension to work, we recommend using a node version higher than v14.16.x. We recommend <strong>v14.16.0</strong>. In the root folder, run `npm install` to set up dependencies control block extension.

5. Finally, run `gulp serve` to add your control block extension to your Sharepoint environment. When the codes compiles without errors, it serves the resulting manifest and should open your Sharepoint environment specified in Step 3.

6. Accept the loading of debug manifests by selection <strong>Load debug scripts</strong> when prompted.
   ![](https://docs.microsoft.com/en-us/sharepoint/dev/images/ext-com-accept-debug-scripts.png)

7. When clicking on the three dots next to a file in document library, you should see the `Open in PDFTron` Button.[](https://pdftron.s3.amazonaws.com/custom/test/jack/sharepoint_readme_pics/Screen+Shot+2022-03-14+at+2.02.38+PM.png)

### Integrate your WebViewer client
To Set up the **client** side, you proceed with these steps:

1. Get started with Sharepoint Online Management Shell (Windows) or PnP PowerShell (Open Source, or non-Windows OS) and connect to your Sharepoint tenant environment:

   * Sharepoint Online Management Shell `Connect-SPOService -Url https://{your-sharepoint-url}.sharepoint.com`
   * [Connect-PnPOnline](https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html) `-Url https://{your-sharepoint-url}.sharepoint.com`

2. We need to disable Custom App Authentication in your Sharepoint environment to make sure the Sharepoint REST API is available to WebViewer:

   * `set-spotenant -DisableCustomAppAuthentication $false`
   * [Set-PnPTenant](https://pnp.github.io/powershell/cmdlets/Set-PnPTenant.html) `-DisableCustomAppAuthentication $false` 

3. To get your access token, you need to [Register SharePoint Add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/register-sharepoint-add-ins). Follow these steps to generate authentication for your WebViewer app:

- Go to `{username}.sharepoint.com/sites/sitename/_layouts/15/AppRegNew.aspx`
  ![](https://pdftron.s3.amazonaws.com/custom/test/jack/sharepoint_readme_pics/Screen+Shot+2022-03-10+at+10.36.18+AM.png)
- generate **client_id**, **client_secret** and other information and click **Create** button
  ![](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/media/apponly/sharepointapponly1.png)
- remember **client_id** and **client_secret**

4. Also, you need to [Grant access using SharePoint App-Only](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs) or you can check out this [youtube video](https://www.youtube.com/watch?v=YMliU4vB_YM&t=631s) to learn how to get your **client_id**, **client_secret**, and **tenant_info**. In your integration it should look similar to this:

    - Go to `{username}.sharepoint.com/sites/sitename/_layouts/15/appinv.aspx`
    - Look up the **client_id**
      ![](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/media/apponly/sharepointapponly2.png)
    - Write Permission Request XML:
      ```
           <AppPermissionRequests AllowAppOnlyPolicy="true">
                <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="FullControl" />
                <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web/list" Right="Write"/>
           </AppPermissionRequests>
      ```
    - click **Create** button, and click **Trust** in the modal if applicable.

5. Now, let's open our **client** project, and create an .env file in your root folder and set each of the following variables:

```
REACT_APP_CLIENT_ID=<client_id>@<tenant_id>
REACT_APP_CLIENT_SECRET= <client_secret>
REACT_APP_RESOURCE= 00000003-0000-0ff1-ce00-000000000000/<username>.sharepoint.com@<tenant_id>
REACT_APP_GRANT_TYPE= client_credentials
REACT_APP_TENANT_ID= <tenant_id>
REACT_APP_ABSOLUTE_URL= <url you can get from step 6>
```

* For your **absolute_url**, in your **sharepoint-extension** project, you can add `console.log(this.context.pageContext.web)` in the `onInit` function to find your absolute_url in ExportToDocCommandSet.ts file
* To find your **tenant_id**, you can checkout this [link](https://piyushksingh.com/2017/03/06/get-office-365-tenant-id/)

6. run `npm install`

7. Next we must copy the static assets required for WebViewer to run. The files are located in `node_modules/@pdftron/webviewer/public` and must be moved into a location that will be served and publicly accessible. In React, it will be `public` folder.

Run the following script from the `/client` folder:

`node tools/copy-webviewer-files.js`

8. Run `npm run start` to start your WebViewer server. You can now use the button we added in the first part, which will open a new tab within your WebViewer environment which shows the selected file.
