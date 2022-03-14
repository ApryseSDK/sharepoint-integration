# WebViewer - Sharepoint integration sample

## Prerequisites

- (Optional but recommended) [Node Version Manager](http://npm.github.io/installation-setup-docs/installing/using-a-node-version-manager.html)
- Set up on Windows environment highly recommended

## For step-by-step help on setting up a SharePoint development environment, see one of the following:

[Sharepoint Extension](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/building-simple-cmdset-with-dialog-api)\
[Set up your Microsoft 365 Tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)\
[Set up your SharePoint Framework Development Environment](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)\
[Create your team site in Sharepoint](https://support.microsoft.com/en-us/office/create-a-team-site-in-sharepoint-ef10c1e7-15f3-42a3-98aa-b5972711777d)

## Project Description

![](https://pdftron.s3.amazonaws.com/custom/test/jack/sharepoint_readme_pics/Screen+Shot+2022-03-14+at+1.39.40+PM.png)
In the github repo there are two individual projects.

   <ol>
      <li>client</li>
      <li>sharepoint-extension</li>
   </ol>
   <strong>Client</strong> is the project which will host PDFTron Webview and Operating PDF Editing. <br>
   <strong>Sharepoint-extension</strong> is Sharepoint Extension which is called            
   
   [Edit Control Block](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/guidance/migrate-from-ecb-to-spfx-extensions) We add PDFTron Button in Sharepoint Document page so that we can open PDFTron Webviewer externally from sharepoint.
   ![](https://pdftron.s3.amazonaws.com/custom/test/jack/sharepoint_readme_pics/Screen+Shot+2022-03-14+at+2.02.38+PM.png)

### Initial Setup

1. Download the git repo and then open the project **sharepoint-extension**.

2. Firstly, you need to set up the webviewer url in _ExportToDocCommandSet.ts_ file
   `const pdftronUrl = 'http://localhost:3000'`, if your webviewer is host on <strong>http://localhost:3000</strong> you can set the pdftronUrl, otherwise you could other location where you set the PDFTron Webview.

3. Secondly, you need to set **pageUrl** in _config/serve.json_, which is your doucment folder in sharepoint. It's just for demo. Open <strong>./config/serve.json</strong> file. Update the <strong>pageUrl</strong> to match a URL of the list where you want to test the solution. After edits your serve.json should look somewhat like:

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

4. <strong>Important</strong>: For the Sharepoint Extension to work, it's better you could use a node version higher than v14.16.x. We recommend <strong>v14.16.0</strong>. In the root fold, you could run `npm install`.

5. Finally, run `gulp serve` to get the app running. When the codes compiles without errors, it serves the resulting manifest from <strong>https://localhost:4321</strong>. <br />

This will also start your default browser within the URL defined in <strong>./config/serve.json</strong> file. Notice that at least in Windows, you can control which browser window is used by activating the preferred one before executing this command.

6. Accept the loading of debug manifests by selection <strong>Load debug scripts</strong> when prompted.
   ![](https://docs.microsoft.com/en-us/sharepoint/dev/images/ext-com-accept-debug-scripts.png)

7. Get started with Sharepoint Online Management Shell and connect to your account

   `Connect-SPOService -Url https://contoso-admin.sharepoint.com`

8. Ensure you disable your custom App Authentication so that you can use sharepoint rest api

   `set-spotenant -DisableCustomAppAuthentication $false`

9. To get your access token, you need to [Register SharePoint Add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/register-sharepoint-add-ins). Following these steps:

   - Go to `{username}.sharepoint.com/sites/sitename/_layouts/15/AppRegNew.aspx`
     ![](https://pdftron.s3.amazonaws.com/custom/test/jack/sharepoint_readme_pics/Screen+Shot+2022-03-10+at+10.36.18+AM.png)
   - generate **client_id**, **client_secret** and other information and click **Create** button
     ![](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/media/apponly/sharepointapponly1.png)
   - remember **client_id** and **client_secret**

10. Also you need to [Granting access using SharePoint App-Only](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs) or you can check out this [youtube channel](https://www.youtube.com/watch?v=YMliU4vB_YM&t=631s) to get your **client_id**, **client_secret**, and **tenant_info**.

    - Go to `{username}.sharepoint.com/sites/sitename/_layouts/15/appinv.aspx`
    - Look up the **clent_id**
      ![](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/media/apponly/sharepointapponly2.png)
    - Write Permission Request XML:
      ```
           <AppPermissionRequests AllowAppOnlyPolicy="true">
                <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="FullControl" />
                <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web/list" Right="Write"/>
           </AppPermissionRequests>
      ```
    - click **Create** button, and click **Trust** if there is a modal shows up

11. To get your **tenant_id**, you can checkout this [link](https://piyushksingh.com/2017/03/06/get-office-365-tenant-id/)

12. There are lots of ways to get your **absolute_url**, in your **sharepoint-extension** project, you can `console.log(this.context.pageContext.web)` to find your absolute_url in ExportToDocCommandSet.ts file.

13. After get all information we want, we can easily set each of your projects up.

14. Firstly, let's open **client** project, create .env in your root folder and set each following variables:

    - REACT_APP_CLIENT_ID: `<client_id>@<tenant_id>`
    - REACT_APP_CLIENT_SECRET: `client_secret`
    - REACT_APP_RESOURCE: `00000003-0000-0ff1-ce00-000000000000/<username>.sharepoint.com@<tenant_id>`
    - REACT_APP_GRANT_TYPE: `client_credentials`
    - REACT_APP_TENANT_ID: `tenant_id`
    - REACT_APP_ABSOLUTE_URL: `<url you can get from step 6>`

15. run `npm instal`

16. Next we must copy the static assets required for WebViewer to run. The files are located in `node_modules/@pdftron/webviewer/public` and must be moved into a location that will be served and publicly accessible. In React, it will be `public` folder.

Inside of a [GitHub project](https://github.com/PDFTron/sharepoint-integration/tree/main/client), we automate the copying of static resources by executing [copy-webviewer-files.js]()

17. run `npm run`

18. If you have any other question, you can contact me directly to my email: **jhou@pdftron.com**
