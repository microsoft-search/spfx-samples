# Frequently Asked Questions and Troubleshooting

## Why do I need the token service endpoint?

The Office API token that is retrieved is only for the user.  It does not request any remote application permissions and simply represents the person that logged in.  It must be exchanged for a specific access token that has the necessary permissions to execute Graph Search API calls.

Reference - [v2 OAuth2 Implict Grant Flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-implicit-grant-flow)

## IdentityAPI is not supported

This means your Office client is not updated to a version that supports the inital identity token retreival.  You may also be using an older Office API version:

- [How to switch from Semi-Annual Channel to Monthly Channel for the Office 365 suite](https://docs.microsoft.com/en-us/office365/troubleshoot/administration/switch-channel-for-office-365)   
- Be sure you are using the "beta" endpoint (https://appsforoffice.microsoft.com/lib/beta/hosted/Office.js) for the Office API - [Understanding the Javascript Api for Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)

## Outlook Addin Doesn't Authenticate

Older Office 365 tenants may not have a default setting enabled to allow for modern authentication to the Exchange Online endpoints. This will effect your Outlook based AddIns and will prevent you from getting the identity token from the IdentityAPI.  This will then of course prevent you from getting the Graph Search API endpoint.

Reference - - [Enable your tenant for Modern Autentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)

## Not getting an Identity Token at startup

You can enable debugging to see what error message is being returned from the IdentityAPI.  This is your first clue as to what might be occuring.  

Reference - [Troubleshoot error messages for single sign-on (SSO)](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/troubleshoot-sso-in-office-add-ins)

Also note that when using Fiddler, the inital identity token exchange may fail with a generic error.  Disable Fiddler during the inital startup, then re-enable it after you have the identity token.