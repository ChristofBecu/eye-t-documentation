# Net Core console turorial

[.NET core tutorial](https://docs.microsoft.com/en-us/graph/tutorials/dotnet-core)

## dependencies

``` python
dotnet add package Microsoft.Extensions.Configuration.UserSecrets --version 3.1.9
dotnet add package Microsoft.Identity.Client --version 4.20.0
dotnet add package Microsoft.Graph --version 3.17.0
```

## Register app

[Azure AD admin center](https://aad.portal.azure.com/)

Azure Active Directory > App Registration

- Naam
- Supported account types : Accounts in een organisatiemap (alle Azure AD-mappen - meerdere tenants) en persoonlijke Microsoft-accounts (bijvoorbeeld Skype, Xbox)
- Omleidings url : openbare client/systeemeigen (mobiel en desktop) value: `https://login.microsoftonline.com/common/oauth2/nativeclient`
- select Registreer
- Kopieer Toepassings-id (client id)
- select Authenticatie > default client type > threat app as public client : Yes
- Save

## Add Azure AD authentication

`dotnet user-secrets init`

`dotnet user-secrets set appId "YOUR_APP_ID_HERE"` Toepassings-id (client-id)

`dotnet user-secrets set scopes "User.Read;Calendars.Read"`

## Implement sign-in - sign in & display access token

[code](https://docs.microsoft.com/en-us/graph/tutorials/dotnet-core?tutorial-step=3)

## Get calendar data

[User details & Calendar events (code)](https://docs.microsoft.com/en-us/graph/tutorials/dotnet-core?tutorial-step=4)