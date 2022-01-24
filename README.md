# WebViewer - Sharepoint integration sample

## Prerequisites

- (Optional but recommended) [Node Version Manager](http://npm.github.io/installation-setup-docs/installing/using-a-node-version-manager.html)
- Set up on Windows environment highly recommended

## For step-by-step help on setting up a SharePoint development environment, see one of the following:

[Sharepoint Extension](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/building-simple-cmdset-with-dialog-api)\
[Set up your Microsoft 365 Tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)\
[Set up your SharePoint Framework Development Environment](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)\
[Create your team site in Sharepoint](https://support.microsoft.com/en-us/office/create-a-team-site-in-sharepoint-ef10c1e7-15f3-42a3-98aa-b5972711777d)

### Initial Setup

1. Get started with Sharepoint Online Management Shell and connect to your account

   `Connect-SPOService -Url https://contoso-admin.sharepoint.com`

2. Ensure you disable your custom App Authentication so that you can use sharepoint rest api

   `set-spotenant -DisableCustomAppAuthentication $false`

3. To get your access token, you need to [Register SharePoint Add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/register-sharepoint-add-ins). Following these steps:

   - Go to `{username}.sharepoint.com/sites/sitename/_layouts/15/AppRegNew.aspx`
   - generate **client_id**, **client_secret** and other information and click **Create** button
   - remember **client_id** and **client_secret**

4. Also you need to [Granting access using SharePoint App-Only](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs) or you can check out this [youtube channel](https://www.youtube.com/watch?v=YMliU4vB_YM&t=631s) to get your **client_id**, **client_secret**, and **tenant_info**.

   - Go to `{username}.sharepoint.com/sites/sitename/_layouts/15/appinv.aspx`
   - Look up the **clent_id**
   - Write Permission Request XML:
     ```
          <AppPermissionRequests AllowAppOnlyPolicy="true">
               <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="FullControl" />
               <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web/list" Right="Write"/>
          </AppPermissionRequests>
     ```
   - click **Create** button, and click **Trust** if there is a modal shows up

5. To get your **tenant_id**, you can checkout this [link](https://piyushksingh.com/2017/03/06/get-office-365-tenant-id/)

6. There are lots of ways to get your **absolute_url**, in your **sharepoint-extension** project, you can `console.log(this.context.pageContext.web)` to find your absolute_url in ExportToDocCommandSet.ts file.

7. After get all information we want, we can easily set each of your projects up.

8. Firstly, let's open **client** project, create .env in your root folder and set each following variables:

   - REACT_APP_CLIENT_ID: `<client_id>@<tenant_id>`
   - REACT_APP_CLIENT_SECRET: `client_secret`
   - REACT_APP_RESOURCE: `00000003-0000-0ff1-ce00-000000000000/<username>.sharepoint.com@<tenant_id>`
   - REACT_APP_GRANT_TYPE: `client_credentials`
   - REACT_APP_TENANT_ID: `tenant_id`
   - REACT_APP_ABSOLUTE_URL: `<url you can get from step 6>`

9. npm instal and npm start

10. Now You have your webviewer sample run, you can set your sharepoint extension up. Open the project **sharepoint-extension**.

11. Firstly, you need to set up the webviewer url in _ExportToDocCommandSet.ts_ file
    `const pdftronUrl = 'http://localhost:3000'`

12. Secondly, you need to set **pageUrl** in _config/serve.json_, which is your doucment folder in sharepoint. It's just for demo.

13. Finally, run `gulp serve` to get the app running

14. If you have any other question, you can contact me directly to my email: **jhou@pdftron.com**
