# ASP.NET Core MVC Graph Tutorial

[ASP.NET Core MVC Tutorial](https://docs.microsoft.com/en-us/graph/tutorials/aspnet-core)

---

- [ASP.NET Core MVC Graph Tutorial](#ASPNET-Core-MVC-Graph-Tutorial)
  - [Create](#Create)
  - [dependencies](#dependencies)
  - [Design](#Design)
    - [alert extensions](#alert-extensions)
      - [./Alerts/WithAlertResult.cs](#AlertsWithAlertResultcs)
      - [./Alerts/AlertExtensions.cs](#AlertsAlertExtensionscs)
    - [user data extensions](#user-data-extensions)
      - [./Graph/GraphClaimsPrincipalExtensions.cs](#GraphGraphClaimsPrincipalExtensionscs)
    - [views](#views)
      - [./Views/Shared/_LoginPartial.cshtml](#ViewsSharedLoginPartialcshtml)
      - [./Views/Shared/_AlertPartial.cshtml](#ViewsSharedAlertPartialcshtml)
      - [./Views/Shared/_Layout.cshtml](#ViewsSharedLayoutcshtml)
      - [./wwwroot/css/site.css](#wwwrootcsssitecss)
      - [./Views/Home/index.cshtml](#ViewsHomeindexcshtml)
      - [./wwwroot/img/no-profile-photo.png](#wwwrootimgno-profile-photopng)
  - [Register the app](#Register-the-app)
    - [Azure Active Directory/App registrations](#Azure-Active-DirectoryApp-registrations)
    - [Manage/Authentication](#ManageAuthentication)
    - [Certificates & secrets](#Certificates--secrets)
  - [Azure AD authentication](#Azure-AD-authentication)
    - [./appsettings.json](#appsettingsjson)
    - [add secrets in cli](#add-secrets-in-cli)
    - [./Graph/GraphConstants.cs](#GraphGraphConstantscs)
    - [./Startup.cs](#Startupcs)
      - [usings](#usings)
      - [ConfigureServices](#ConfigureServices)
      - [Configure](#Configure)
    - [./Controllers/HomeController.cs](#ControllersHomeControllercs)
  - [Get User Details](#Get-User-Details)
    - [./Graph/GraphClaimsPrincipalExtensions.cs replace content](#GraphGraphClaimsPrincipalExtensionscs-replace-content)
    - [change ./Startup.cs](#change-Startupcs)
      - [replace .AddMicrosoftIdentityWebApp(Configuration)](#replace-AddMicrosoftIdentityWebAppConfiguration)
      - [Add function call AddMicrosoftGraph](#Add-function-call-AddMicrosoftGraph)
    - [Change ./Controllers/HomeController.cs](#Change-ControllersHomeControllercs)
      - [replace index()](#replace-index)
      - [remove all references to `ITokenAcquisition`](#remove-all-references-to-ITokenAcquisition)
    - [Debug](#Debug)
      - [Authentication error: value cannot be null. (Parameter 'value')](#Authentication-error-value-cannot-be-null-Parameter-value)
      - [Authentication error](#Authentication-error)
  - [Calendar view](#Calendar-view)
    - [Get calendar events from outlook](#Get-calendar-events-from-outlook)
      - [./Controllers/CalendarController.cs](#ControllersCalendarControllercs)
    - [Debug](#Debug-1)
      - [Timezone](#Timezone)
      - [Photo not found (2nd attempt)](#Photo-not-found-2nd-attempt)
  - [Display results](#Display-results)
    - [Viewmodels](#Viewmodels)
      - [./Models/CalendarViewEvent.cs](#ModelsCalendarViewEventcs)
      - [./Models/DailyViewModel.cs](#ModelsDailyViewModelcs)
      - [./Models/CalendarViewModel.cs](#ModelsCalendarViewModelcs)
    - [Views](#Views)
      - [./Views/Calendar/_DailyEventsPartial.cshtml](#ViewsCalendarDailyEventsPartialcshtml)
      - [./View/Calendar/Index.cshtml](#ViewCalendarIndexcshtml)
    - [Update calendar controller](#Update-calendar-controller)
      - [replace index method](#replace-index-method)
    - [Debug calendar views](#Debug-calendar-views)
  - [Create a new event](#Create-a-new-event)
    - [Create models](#Create-models)
      - [./Models/NewEvent.cs](#ModelsNewEventcs)
    - [Create views](#Create-views)
      - [./Views/Calendar/Index.cs](#ViewsCalendarIndexcs)
    - [Add controller actions](#Add-controller-actions)
      - [./Controllers/CalendarController.cs](#ControllersCalendarControllercs-1)
        - [render the new event form](#render-the-new-event-form)
        - [receive new event from form & save through MS Graph to users calendar](#receive-new-event-from-form--save-through-MS-Graph-to-users-calendar)

---

## Create

`dotnet new mvc -o GraphTutorial`

## dependencies

```python
dotnet add package Microsoft.Identity.Web --version 0.3.0-preview
dotnet add package Microsoft.Identity.Web.UI --version 0.3.0-preview
dotnet add package Microsoft.Graph --version 3.12.0
```

---

## Design

### alert extensions

extension methods for the IActionResult type returned by controller views. This extension will enable passing temporary error or success messages to the view.

#### ./Alerts/WithAlertResult.cs

```c#
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Microsoft.Extensions.DependencyInjection;
using System.Threading.Tasks;

namespace GraphTutorial
{
    // WithAlertResult adds temporary error/info/success
    // messages to the result of a controller action.
    // This data is read and displayed by the _AlertPartial view
    public class WithAlertResult : IActionResult
    {
        public IActionResult Result { get; }
        public string Type { get; }
        public string Message { get; }
        public string DebugInfo { get; }

        public WithAlertResult(IActionResult result,
                                    string type,
                                    string message,
                                    string debugInfo)
        {
            Result = result;
            Type = type;
            Message = message;
            DebugInfo = debugInfo;
        }

        public async Task ExecuteResultAsync(ActionContext context)
        {
            var factory = context.HttpContext.RequestServices
            .GetService<ITempDataDictionaryFactory>();

            var tempData = factory.GetTempData(context.HttpContext);

            tempData["_alertType"] = Type;
            tempData["_alertMessage"] = Message;
            tempData["_alertDebugInfo"] = DebugInfo;

            await Result.ExecuteResultAsync(context);
        }
    }
}
```

#### ./Alerts/AlertExtensions.cs

```c#
using Microsoft.AspNetCore.Mvc;

namespace GraphTutorial
{
    public static class AlertExtensions
    {
        public static IActionResult WithError(this IActionResult result,
                                            string message,
                                            string debugInfo = null)
        {
            return Alert(result, "danger", message, debugInfo);
        }

        public static IActionResult WithSuccess(this IActionResult result,
                                            string message,
                                            string debugInfo = null)
        {
            return Alert(result, "success", message, debugInfo);
        }

        public static IActionResult WithInfo(this IActionResult result,
                                            string message,
                                            string debugInfo = null)
        {
            return Alert(result, "info", message, debugInfo);
        }

        private static IActionResult Alert(IActionResult result,
                                        string type,
                                        string message,
                                        string debugInfo)
        {
            return new WithAlertResult(result, type, message, debugInfo);
        }
    }
}
```

### user data extensions

extension methods for the ClaimsPrincipal object generated by the Microsoft Identity platform. This will allow you to extend the existing user identity with data from Microsoft Graph

#### ./Graph/GraphClaimsPrincipalExtensions.cs

```c#
using System.Security.Claims;

namespace GraphTutorial
{
    public static class GraphClaimTypes {
        public const string DisplayName ="graph_name";
        public const string Email = "graph_email";
        public const string Photo = "graph_photo";
        public const string TimeZone = "graph_timezone";
        public const string DateTimeFormat = "graph_datetimeformat";
    }

    // Helper methods to access Graph user data stored in
    // the claims principal
    public static class GraphClaimsPrincipalExtensions
    {
        public static string GetUserGraphDisplayName(this ClaimsPrincipal claimsPrincipal)
        {
            return "Adele Vance";
        }

        public static string GetUserGraphEmail(this ClaimsPrincipal claimsPrincipal)
        {
            return "adelev@contoso.com";
        }

        public static string GetUserGraphPhoto(this ClaimsPrincipal claimsPrincipal)
        {
            return "/img/no-profile-photo.png";
        }
    }
}
```

### views

#### ./Views/Shared/_LoginPartial.cshtml

```js
@using GraphTutorial

<ul class="nav navbar-nav">
@if (User.Identity.IsAuthenticated)
{
    <li class="nav-item dropdown">
        <a class="nav-link dropdown-toggle" data-toggle="dropdown" href="#" role="button">
            <img src="@User.GetUserGraphPhoto()" class="nav-profile-photo rounded-circle align-self-center mr-2">
        </a>
        <div class="dropdown-menu dropdown-menu-right">
            <h5 class="dropdown-item-text mb-0">@User.GetUserGraphDisplayName()</h5>
            <p class="dropdown-item-text text-muted mb-0">@User.GetUserGraphEmail()</p>
            <div class="dropdown-divider"></div>
            <a class="dropdown-item" asp-area="MicrosoftIdentity" asp-controller="Account" asp-action="SignOut">Sign out</a>
        </div>
    </li>
}
else
{
    <li class="nav-item">
        <a class="nav-link" asp-area="MicrosoftIdentity" asp-controller="Account" asp-action="SignIn">Sign in</a>
    </li>
}
</ul>
```

#### ./Views/Shared/_AlertPartial.cshtml

```js
@{
    var type = $"{TempData["_alertType"]}";
    var message = $"{TempData["_alertMessage"]}";
    var debugInfo = $"{TempData["_alertDebugInfo"]}";
}

@if (!string.IsNullOrEmpty(type))
{
    <div class="alert alert-@type" role="alert">
        @if (string.IsNullOrEmpty(debugInfo))
        {
            <p class="mb-0">@message</p>
        }
        else
        {
            <p class="mb-3">@message</p>
            <pre class="alert-pre border bg-light p-2"><code>@debugInfo</code></pre>
        }
    </div>
}
```

#### ./Views/Shared/_Layout.cshtml

```html
@{
    string controller = $"{ViewContext.RouteData.Values["controller"]}";
}

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>@ViewData["Title"] - GraphTutorial</title>
    <link rel="stylesheet" href="~/lib/bootstrap/dist/css/bootstrap.min.css" />
    <link rel="stylesheet" href="~/css/site.css" />
</head>
<body>
    <header>
        <nav class="navbar navbar-expand-sm navbar-toggleable-sm navbar-dark bg-dark border-bottom box-shadow mb-3">
            <div class="container">
                <a class="navbar-brand" asp-area="" asp-controller="Home" asp-action="Index">GraphTutorial</a>
                <button class="navbar-toggler" type="button" data-toggle="collapse" data-target=".navbar-collapse"
                    aria-controls="navbarSupportedContent" aria-expanded="false" aria-label="Toggle navigation">
                    <span class="navbar-toggler-icon"></span>
                </button>
                <div class="navbar-collapse collapse mr-auto">
                    <ul class="navbar-nav flex-grow-1">
                        <li class="@(controller == "Home" ? "nav-item active" : "nav-item")">
                            <a class="nav-link" asp-area="" asp-controller="Home" asp-action="Index">Home</a>
                        </li>
                        @if (User.Identity.IsAuthenticated)
                        {
                        <li class="@(controller == "Calendar" ? "nav-item active" : "nav-item")">
                            <a class="nav-link" asp-area="" asp-controller="Calendar" asp-action="Index">Calendar</a>
                        </li>
                        }
                    </ul>
                    <partial name="_LoginPartial"/>
                </div>
            </div>
        </nav>
    </header>
    <div class="container">
        <main role="main" class="pb-3">
            <partial name="_AlertPartial"/>
            @RenderBody()
        </main>
    </div>

    <footer class="border-top footer text-muted">
        <div class="container">
            © 2020 - GraphTutorial - <a asp-area="" asp-controller="Home" asp-action="Privacy">Privacy</a>
        </div>
    </footer>
    <script src="~/lib/jquery/dist/jquery.min.js"></script>
    <script src="~/lib/bootstrap/dist/js/bootstrap.bundle.min.js"></script>
    <script src="~/js/site.js" asp-append-version="true"></script>
    @RenderSection("Scripts", required: false)
</body>
</html>
```

#### ./wwwroot/css/site.css

add to bottom

```css
.nav-profile-photo {
  width: 32px;
}

.alert-pre {
  word-wrap: break-word;
  word-break: break-all;
  white-space: pre-wrap;
}

.calendar-view-date-cell {
  width: 150px;
}

.calendar-view-date {
  width: 40px;
  font-size: 36px;
  line-height: 36px;
  margin-right: 10px;
}

.calendar-view-month {
  font-size: 0.75em;
}

.calendar-view-timespan {
  width: 200px;
}

.calendar-view-subject {
  font-size: 1.25em;
}

.calendar-view-organizer {
  font-size: .75em;
}
```

#### ./Views/Home/index.cshtml

```html
@{
    ViewData["Title"] = "Home Page";
}

@using GraphTutorial

<div class="jumbotron">
    <h1>ASP.NET Core Graph Tutorial</h1>
    <p class="lead">This sample app shows how to use the Microsoft Graph API to access a user's data from ASP.NET Core</p>
    @if (User.Identity.IsAuthenticated)
    {
        <h4>Welcome @User.GetUserGraphDisplayName()!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
    }
    else
    {
        <a class="btn btn-primary btn-large" asp-area="MicrosoftIdentity" asp-controller="Account" asp-action="SignIn">Click here to sign in</a>
    }
</div>
```

#### ./wwwroot/img/no-profile-photo.png

[github](https://github.com/microsoftgraph/msgraph-training-aspnet-core/blob/master/demo/GraphTutorial/wwwroot/img/no-profile-photo.png)

---

## Register the app

### Azure Active Directory/App registrations

- set Name
- Set Supported account types to Accounts in any organizational directory and personal Microsoft accounts.
- Under Redirect URI, set the first drop-down to Web and set the value to `https://localhost:5001/`

### Manage/Authentication

- Under Redirect URIs add a URI with the value `https://localhost:5001/signin-oidc`
- Set the Logout URL to `https://localhost:5001/signout-oidc`
- Locate the Implicit grant section and enable ID tokens. Select Save.

### Certificates & secrets

- value : Forever
- Expires: Never

---

## Azure AD authentication

***required to obtain the necessary OAuth access token to call the Microsoft Graph API***

***To avoid storing the application ID and secret in source, you will use the .NET Secret Manager to store these values. The Secret Manager is for development purposes only, production apps should use a trusted secret manager for storing secrets.***

### ./appsettings.json

```json
"AzureAd": {
    "Instance": "https://login.microsoftonline.com/",
    "TenantId": "common",
    "CallbackPath": "/signin-oidc"
  },
```

### add secrets in cli

```python
dotnet user-secrets init
dotnet user-secrets set "AzureAd:ClientId" "YOUR_APP_ID"
dotnet user-secrets set "AzureAd:ClientSecret" "YOUR_APP_SECRET"
```

### ./Graph/GraphConstants.cs

adding the Microsoft Identity platform services to the application

```c#
namespace GraphTutorial
{
    public static class GraphConstants
    {
        // Defines the permission scopes used by the app
        public readonly static string[] Scopes =
        {
            "User.Read",
            "MailboxSettings.Read",
            "Calendars.ReadWrite"
        };
    }
}
```

### ./Startup.cs

#### usings

```c#
using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc.Authorization;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.UI;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.Graph;
using System.Net;
using System.Net.Http.Headers;
```

#### ConfigureServices

```c#
public void ConfigureServices(IServiceCollection services)
{
    services
        // Use OpenId authentication
        .AddAuthentication(OpenIdConnectDefaults.AuthenticationScheme)
        // Specify this is a web app and needs auth code flow
        .AddMicrosoftIdentityWebApp(Configuration)
        // Add ability to call web API (Graph)
        // and get access tokens
        .EnableTokenAcquisitionToCallDownstreamApi(options => {
            Configuration.Bind("AzureAd", options);
        }, GraphConstants.Scopes)
        // Use in-memory token cache
        // See https://github.com/AzureAD/microsoft-identity-web/wiki/token-cache-serialization
        .AddInMemoryTokenCaches();

    // Require authentication
    services.AddControllersWithViews(options =>
    {
        var policy = new AuthorizationPolicyBuilder()
            .RequireAuthenticatedUser()
            .Build();
        options.Filters.Add(new AuthorizeFilter(policy));
    })
    // Add the Microsoft Identity UI pages for signin/out
    .AddMicrosoftIdentityUI();
}
```

#### Configure

```c#
app.UseAuthentication();
app.UseAuthorization();
```

### ./Controllers/HomeController.cs

```c#
using GraphTutorial.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Web;
using System.Diagnostics;
using System.Threading.Tasks;

namespace GraphTutorial.Controllers
{
    public class HomeController : Controller
    {
        ITokenAcquisition _tokenAcquisition;
        private readonly ILogger<HomeController> _logger;

        // Get the ITokenAcquisition interface via
        // dependency injection
        public HomeController(
            ITokenAcquisition tokenAcquisition,
            ILogger<HomeController> logger)
        {
            _tokenAcquisition = tokenAcquisition;
            _logger = logger;
        }

        public async Task<IActionResult> Index()
        {
            // TEMPORARY
            // Get the token and display it
            try
            {
                string token = await _tokenAcquisition
                    .GetAccessTokenForUserAsync(GraphConstants.Scopes);
                return View().WithInfo("Token acquired", token);
            }
            catch (MicrosoftIdentityWebChallengeUserException)
            {
                return Challenge();
            }
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        [AllowAnonymous]
        public IActionResult ErrorWithMessage(string message, string debug)
        {
            return View("Index").WithError(message, debug);
        }
    }
}
```

## Get User Details

### ./Graph/GraphClaimsPrincipalExtensions.cs replace content

```c#
using Microsoft.Graph;
using System;
using System.IO;
using System.Security.Claims;

namespace GraphTutorial
{
    public static class GraphClaimTypes {
        public const string DisplayName ="graph_name";
        public const string Email = "graph_email";
        public const string Photo = "graph_photo";
        public const string TimeZone = "graph_timezone";
        public const string TimeFormat = "graph_timeformat";
    }

    // Helper methods to access Graph user data stored in
    // the claims principal
    public static class GraphClaimsPrincipalExtensions
    {
        public static string GetUserGraphDisplayName(this ClaimsPrincipal claimsPrincipal)
        {
            return claimsPrincipal.FindFirstValue(GraphClaimTypes.DisplayName);
        }

        public static string GetUserGraphEmail(this ClaimsPrincipal claimsPrincipal)
        {
            return claimsPrincipal.FindFirstValue(GraphClaimTypes.Email);
        }

        public static string GetUserGraphPhoto(this ClaimsPrincipal claimsPrincipal)
        {
            return claimsPrincipal.FindFirstValue(GraphClaimTypes.Photo);
        }

        public static string GetUserGraphTimeZone(this ClaimsPrincipal claimsPrincipal)
        {
            return claimsPrincipal.FindFirstValue(GraphClaimTypes.TimeZone);
        }

        public static string GetUserGraphTimeFormat(this ClaimsPrincipal claimsPrincipal)
        {
            return claimsPrincipal.FindFirstValue(GraphClaimTypes.TimeFormat);
        }

        public static void AddUserGraphInfo(this ClaimsPrincipal claimsPrincipal, User user)
        {
            var identity = claimsPrincipal.Identity as ClaimsIdentity;

            identity.AddClaim(
                new Claim(GraphClaimTypes.DisplayName, user.DisplayName));
            identity.AddClaim(
                new Claim(GraphClaimTypes.Email,
                    user.Mail ?? user.UserPrincipalName));
            identity.AddClaim(
                new Claim(GraphClaimTypes.TimeZone,
                    user.MailboxSettings.TimeZone));
            identity.AddClaim(
                new Claim(GraphClaimTypes.TimeFormat, user.MailboxSettings.TimeFormat));
        }

        public static void AddUserGraphPhoto(this ClaimsPrincipal claimsPrincipal, Stream photoStream)
        {
            var identity = claimsPrincipal.Identity as ClaimsIdentity;

            if (photoStream == null)
            {
                // Add the default profile photo
                identity.AddClaim(
                    new Claim(GraphClaimTypes.Photo, "/img/no-profile-photo.png"));
                return;
            }

            // Copy the photo stream to a memory stream
            // to get the bytes out of it
            var memoryStream = new MemoryStream();
            photoStream.CopyTo(memoryStream);
            var photoBytes = memoryStream.ToArray();

            // Generate a date URI for the photo
            var photoUrl = $"data:image/png;base64,{Convert.ToBase64String(photoBytes)}";

            identity.AddClaim(
                new Claim(GraphClaimTypes.Photo, photoUrl));
        }
    }
}
```


### change ./Startup.cs

#### replace .AddMicrosoftIdentityWebApp(Configuration)

- It adds an event handler for the OnAuthorizationCodeReceived event. This handler uses the default handler to exchange the authorization code for a token and initialize the Microsoft.Identity.Web classes.
- It uses the ITokenAcquisition interface to get an access token.
- It calls Microsoft Graph to get the user's profile and photo.
- It adds the Graph information to the user's identity.

```c#
// Specify this is a web app and needs auth code flow
.AddMicrosoftIdentityWebApp(options => {
    Configuration.Bind("AzureAd", options);

    options.Prompt = "select_account";

    var authCodeHandler = options.Events.OnAuthorizationCodeReceived;

    options.Events.OnAuthorizationCodeReceived += async context => {
        // Invoke the original handler first
        // This allows the Microsoft.Identity.Web library to
        // add the user to its token cache
        await authCodeHandler(context);

        var tokenAcquisition = context.HttpContext.RequestServices
            .GetRequiredService<ITokenAcquisition>();

        var graphClient = new GraphServiceClient(
            new DelegateAuthenticationProvider(async (request) => {
                var token = await tokenAcquisition
                    .GetAccessTokenForUserAsync(GraphConstants.Scopes, user:context.Principal);
                request.Headers.Authorization =
                    new AuthenticationHeaderValue("Bearer", token);
            })
        );

        // Get user information from Graph
        var user = await graphClient.Me.Request()
            .Select(u => new {
                u.DisplayName,
                u.Mail,
                u.UserPrincipalName,
                u.MailboxSettings
            })
            .GetAsync();

        context.Principal.AddUserGraphInfo(user);

        // Get the user's photo
        // If the user doesn't have a photo, this throws
        try
        {
            var photo = await graphClient.Me
                .Photos["48x48"]
                .Content
                .Request()
                .GetAsync();

            context.Principal.AddUserGraphPhoto(photo);
        }
        catch (ServiceException ex)
        {
            if (ex.IsMatch("ErrorItemNotFound"))
            {
                context.Principal.AddUserGraphPhoto(null);
            }
            else
            {
                throw ex;
            }
        }
    };

    options.Events.OnAuthenticationFailed = context => {
        var error = WebUtility.UrlEncode(context.Exception.Message);
        context.Response
            .Redirect($"/Home/ErrorWithMessage?message=Authentication+error&debug={error}");
        context.HandleResponse();

        return Task.FromResult(0);
    };

    options.Events.OnRemoteFailure = context => {
        if (context.Failure is OpenIdConnectProtocolException)
        {
            var error = WebUtility.UrlEncode(context.Failure.Message);
            context.Response
                .Redirect($"/Home/ErrorWithMessage?message=Sign+in+error&debug={error}");
            context.HandleResponse();
        }

        return Task.FromResult(0);
    };
})

```

#### Add function call AddMicrosoftGraph

- This will make an authenticated GraphServiceClient available to controllers via dependency injection.
- after `EnableTokenAcquisitionToCallDownstreamApi`
- before `AddInMemoryTokenCaches`

```c#
// Add a GraphServiceClient via dependency injection
.AddMicrosoftGraph(options => {
    options.Scopes = string.Join(' ', GraphConstants.Scopes);
})
```

### Change ./Controllers/HomeController.cs

#### replace index()

```c#
public IActionResult Index()
{
    return View();
}
```

#### remove all references to `ITokenAcquisition`

### Debug

#### Authentication error: value cannot be null. (Parameter 'value')

```c#
identity.AddClaim(
                new Claim(GraphClaimTypes.TimeZone,
                user.MailboxSettings.TimeZone));
```

***user.MailboxSettings.TimeZone is null, cannot be added to claims***

#### Authentication error

```python
Code: ConsumerPhotoIsNotSupported
Message: The photo wasn't found.
Inner error:
	AdditionalData:
	date: 2020-10-15T07:10:57
	request-id: 0b48dbf6-0a41-4282-afa0-34e56bba2217
	client-request-id: 0b48dbf6-0a41-4282-afa0-34e56bba2217
ClientRequestId: 0b48dbf6-0a41-4282-afa0-34e56bba2217
```

```c#
try
{
    var photo = await graphClient.Me
       .Photos["48x48"]
       .Content
       .Request()
       .GetAsync();

    context.Principal.AddUserGraphPhoto(photo);
}
```

throws exception if the user has no photo set
can be fixed by adding variable initializer

```c# 
System.IO.Stream photo = null;
```

---

## Calendar view

### Get calendar events from outlook

#### ./Controllers/CalendarController.cs

```c#
using GraphTutorial.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Web;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GraphTutorial.Controllers
{
    public class CalendarController : Controller
    {
        private readonly GraphServiceClient _graphClient;
        private readonly ILogger<HomeController> _logger;

        public CalendarController(
            GraphServiceClient graphClient,
            ILogger<HomeController> logger)
        {
            _graphClient = graphClient;
            _logger = logger;
        }
    }
}
```

+ add following methods to controller

```c#
private async Task<IList<Event>> GetUserWeekCalendar(DateTime startOfWeek)
{
    // Configure a calendar view for the current week
    var endOfWeek = startOfWeek.AddDays(7);

    var viewOptions = new List<QueryOption>
    {
        new QueryOption("startDateTime", startOfWeek.ToString("o")),
        new QueryOption("endDateTime", endOfWeek.ToString("o"))
    };

    var events = await _graphClient.Me
        .CalendarView
        .Request(viewOptions)
        // Send user time zone in request so date/time in
        // response will be in preferred time zone
        .Header("Prefer", $"outlook.timezone=\"{User.GetUserGraphTimeZone()}\"")
        // Get max 50 per request
        .Top(50)
        // Only return fields app will use
        .Select(e => new
        {
            e.Subject,
            e.Organizer,
            e.Start,
            e.End
        })
        // Order results chronologically
        .OrderBy("start/dateTime")
        .GetAsync();

    IList<Event> allEvents;
    // Handle case where there are more than 50
    if (events.NextPageRequest != null)
    {
        allEvents = new List<Event>();
        // Create a page iterator to iterate over subsequent pages
        // of results. Build a list from the results
        var pageIterator = PageIterator<Event>.CreatePageIterator(
            _graphClient, events,
            (e) => {
                allEvents.Add(e);
                return true;
            }
        );
        await pageIterator.IterateAsync();
    }
    else
    {
        // If only one page, just use the result
        allEvents = events.CurrentPage;
    }

    return allEvents;
}

private static DateTime GetUtcStartOfWeekInTimeZone(DateTime today, TimeZoneInfo timeZone)
{
    // Assumes Sunday as first day of week
    int diff = System.DayOfWeek.Sunday - today.DayOfWeek;

    // create date as unspecified kind
    var unspecifiedStart = DateTime.SpecifyKind(today.AddDays(diff), DateTimeKind.Unspecified);

    // convert to UTC
    return TimeZoneInfo.ConvertTimeToUtc(unspecifiedStart, timeZone);
}
```

***consider what the code in GetUserWeekCalendar does.***

- It uses the user's time zone to get UTC start and end date/time values for the week.
- It queries the user's calendar view to get all events that fall between the start and end date/times. Using a calendar view instead of listing events expands recurring events, returning any occurrences that happen in the specified time window.
- It uses the Prefer: outlook.timezone header to get results back in the user's timezone.
- It uses Select to limit the fields that come back to just those used by the app.
- It uses OrderBy to sort the results chronologically.
- It uses a PageIterator to page through the events collection. This handles the case where the user has more events on their calendar than the requested page size.

+ add following method to controller

```c#
// Minimum permission scope needed for this view
[AuthorizeForScopes(Scopes = new[] { "Calendars.Read" })]
public async Task<IActionResult> Index()
{
    try
    {
        var userTimeZone = TimeZoneInfo.FindSystemTimeZoneById(
            User.GetUserGraphTimeZone());
        var startOfWeek = CalendarController.GetUtcStartOfWeekInTimeZone(
            DateTime.Today, userTimeZone);

        var events = await GetUserWeekCalendar(startOfWeek);

        // Return a JSON dump of events
        return new ContentResult {
            Content = _graphClient.HttpProvider.Serializer.SerializeObject(events),
            ContentType = "application/json"
        };
    }
    catch (ServiceException ex)
    {
        if (ex.InnerException is MicrosoftIdentityWebChallengeUserException)
        {
            throw ex;
        }

        return new ContentResult {
            Content = $"Error getting calendar view: {ex.Message}",
            ContentType = "text/plain"
        };
    }
}
```
---

### Debug

#### Timezone

```c#
try
{
    var userTimeZone = TimeZoneInfo.FindSystemTimeZoneById(
        User.GetUserGraphTimeZone());
        ...
```

***ArgumentNullException: Value cannot be null. (Parameter 'id')***

installeer [TimeZoneConverter 3.3.0](https://www.nuget.org/packages/TimeZoneConverter/)

Code aanpassen:

```c#
public static string GetUserGraphTimeZone(this ClaimsPrincipal claimsPrincipal)
{
    var localtz_string = "Europe/Brussels";
    var localtz = TZConvert.IanaToWindows(localtz_string);

    //string tzclaim = GraphClaimTypes.TimeZone;

    //string tz = claimsPrincipal.FindFirstValue(tzclaim);

    return localtz;
}
```

#### Photo not found (2nd attempt)

In plaats van oplossing van hierboven (System.IO.Stream photo = null;)

```c#
if (ex.StatusCode.ToString() == ("NotFound"))
{
    context.Principal.AddUserGraphPhoto(null);
}
else
```

---

## Display results

### Viewmodels

#### ./Models/CalendarViewEvent.cs

```c#
using Microsoft.Graph;
using System;

namespace GraphTutorial.Models
{
    public class CalendarViewEvent
    {
        public string Subject { get; private set; }
        public string Organizer { get; private set; }
        public DateTime Start { get; private set; }
        public DateTime End { get; private set; }

        public CalendarViewEvent(Event graphEvent)
        {
            Subject = graphEvent.Subject;
            Organizer = graphEvent.Organizer.EmailAddress.Name;
            Start = DateTime.Parse(graphEvent.Start.DateTime);
            End = DateTime.Parse(graphEvent.End.DateTime);
        }
    }
}
```

#### ./Models/DailyViewModel.cs

```c#
using System;
using System.Collections.Generic;

namespace GraphTutorial.Models
{
    public class DailyViewModel
    {
        // Day the view is for
        public DateTime Day { get; private set; }
        // Events on this day
        public IEnumerable<CalendarViewEvent> Events { get; private set; }

        public DailyViewModel(DateTime day, IEnumerable<CalendarViewEvent> events)
        {
            Day = day;
            Events = events;
        }
    }
}
```

#### ./Models/CalendarViewModel.cs

```c#
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GraphTutorial.Models
{
    public class CalendarViewModel
    {
        private DateTime _startOfWeek;
        private List<CalendarViewEvent> _events;

        public CalendarViewModel()
        {
            _startOfWeek = DateTime.MinValue;
            _events = new List<CalendarViewEvent>();
        }

        public CalendarViewModel(DateTime startOfWeek, IEnumerable<Event> events)
        {
            _startOfWeek = startOfWeek;
            _events = new List<CalendarViewEvent>();

            if (events != null)
            {
              foreach (var item in events)
              {
                  _events.Add(new CalendarViewEvent(item));
              }
            }
        }

        // Get the start - end dates of the week
        public string TimeSpan()
        {
            return $"{_startOfWeek.ToString("MMMM d, yyyy")} - {_startOfWeek.AddDays(6).ToString("MMMM d, yyyy")}";
        }

        // Property accessors to pass to the daily view partial
        // These properties get all events on the specific day
        public DailyViewModel Sunday
        {
            get
            {
                return new DailyViewModel(
                  _startOfWeek,
                  GetEventsForDay(System.DayOfWeek.Sunday));
            }
        }

        public DailyViewModel Monday
        {
            get
            {
                return new DailyViewModel(
                  _startOfWeek.AddDays(1),
                  GetEventsForDay(System.DayOfWeek.Monday));
            }
        }

        public DailyViewModel Tuesday
        {
            get
            {
                return new DailyViewModel(
                  _startOfWeek.AddDays(2),
                  GetEventsForDay(System.DayOfWeek.Tuesday));
            }
        }

        public DailyViewModel Wednesday
        {
            get
            {
                return new DailyViewModel(
                  _startOfWeek.AddDays(3),
                  GetEventsForDay(System.DayOfWeek.Wednesday));
            }
        }

        public DailyViewModel Thursday
        {
            get
            {
                return new DailyViewModel(
                  _startOfWeek.AddDays(4),
                  GetEventsForDay(System.DayOfWeek.Thursday));
            }
        }

        public DailyViewModel Friday
        {
            get
            {
                return new DailyViewModel(
                  _startOfWeek.AddDays(5),
                  GetEventsForDay(System.DayOfWeek.Friday));
            }
        }

        public DailyViewModel Saturday
        {
            get
            {
                return new DailyViewModel(
                  _startOfWeek.AddDays(6),
                  GetEventsForDay(System.DayOfWeek.Saturday));
            }
        }

        private IEnumerable<CalendarViewEvent> GetEventsForDay(System.DayOfWeek day)
        {
            return _events.Where(e => e.Start.DayOfWeek.Equals(day));
        }
    }
}
```

### Views

#### ./Views/Calendar/_DailyEventsPartial.cshtml

```c#
@model DailyViewModel

@{
    bool dateCellAdded = false;
    var timeFormat = User.GetUserGraphTimeFormat();
    var rowClass = Model.Day.Date.Equals(DateTime.Today.Date) ? "table-warning" : "";
}

@if (Model.Events.Count() <= 0)
{
  // Render an empty row for the day
  <tr>
    <td class="calendar-view-date-cell">
      <div class="calendar-view-date float-left text-right">@Model.Day.Day</div>
      <div class="calendar-view-day">@Model.Day.ToString("dddd")</div>
      <div class="calendar-view-month text-muted">@Model.Day.ToString("MMMM, yyyy")</div>
    </td>
    <td></td>
    <td></td>
  </tr>
}

@foreach(var item in Model.Events)
{
  <tr class="@rowClass">
    @if (!dateCellAdded)
    {
      // Only add the day cell once
      dateCellAdded = true;
      <td class="calendar-view-date-cell" rowspan="@Model.Events.Count()">
        <div class="calendar-view-date float-left text-right">@Model.Day.Day</div>
        <div class="calendar-view-day">@Model.Day.ToString("dddd")</div>
        <div class="calendar-view-month text-muted">@Model.Day.ToString("MMMM, yyyy")</div>
      </td>
    }
    <td class="calendar-view-timespan">
      <div>@item.Start.ToString(timeFormat) - @item.End.ToString(timeFormat)</div>
    </td>
    <td>
      <div class="calendar-view-subject">@item.Subject</div>
      <div class="calendar-view-organizer">@item.Organizer</div>
    </td>
  </tr>
}
```

#### ./View/Calendar/Index.cshtml

```c#
@model CalendarViewModel

@{
    ViewData["Title"] = "Calendar";
}

<div class="mb-3">
  <h1 class="mb-3">@Model.TimeSpan()</h1>
  <a class="btn btn-light btn-sm" asp-controller="Calendar" asp-action="New">New event</a>
</div>
<div class="calendar-week">
  <div class="table-responsive">
    <table class="table table-sm">
      <thead>
        <tr>
          <th>Date</th>
          <th>Time</th>
          <th>Event</th>
        </tr>
      </thead>
      <tbody>
        <partial name="_DailyEventsPartial" for="Sunday" />
        <partial name="_DailyEventsPartial" for="Monday" />
        <partial name="_DailyEventsPartial" for="Tuesday" />
        <partial name="_DailyEventsPartial" for="Wednesday" />
        <partial name="_DailyEventsPartial" for="Thursday" />
        <partial name="_DailyEventsPartial" for="Friday" />
        <partial name="_DailyEventsPartial" for="Saturday" />
      </tbody>
    </table>
  </div>
</div>
```

### Update calendar controller

#### replace index method

```c#
// Minimum permission scope needed for this view
[AuthorizeForScopes(Scopes = new[] { "Calendars.Read" })]
public async Task<IActionResult> Index()
{
    try
    {
        var userTimeZone = TimeZoneInfo.FindSystemTimeZoneById(
            User.GetUserGraphTimeZone());
        var startOfWeek = CalendarController.GetUtcStartOfWeekInTimeZone(
            DateTime.Today, userTimeZone);

        var events = await GetUserWeekCalendar(startOfWeek);

        var model = new CalendarViewModel(startOfWeek, events);

        return View(model);
    }
    catch (ServiceException ex)
    {
        if (ex.InnerException is MicrosoftIdentityWebChallengeUserException)
        {
            throw ex;
        }

        return View(new CalendarViewModel())
            .WithError("Error getting calendar view", ex.Message);
    }
}
```

### Debug calendar views

***appointments are one day off***

e.g. an appointment for 15/10/2020 is shown on day 14/10/2020

TODO : fix it

In de calendar controller in methode GetUtcStartOfWeekInTimeZone wordt zondag als eerste dag van de week gezet, door deze op maandag te zetten worden de events op de juiste datum geplaatst

Door daarna de volgore van de week aan te passen in de calendar/index view, start de week in het resultaat op maandag

---

## Create a new event

### Create models

#### ./Models/NewEvent.cs

```c#
using System;
using System.ComponentModel.DataAnnotations;

namespace GraphTutorial.Models
{
    public class NewEvent
    {
        [Required]
        public string Subject { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        [DataType(DataType.MultilineText)]
        public string Body { get; set; }
        [RegularExpression(@"((\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*)*([;])*)*",
          ErrorMessage="Please enter one or more email addresses separated by a semi-colon (;)")]
        public string Attendees { get; set; }
    }
}
```

### Create views

#### ./Views/Calendar/Index.cs

```html
@model NewEvent

@{
    ViewData["Title"] = "New event";
}

<form asp-action="New">
  <div asp-validation-summary="ModelOnly" class="text-danger"></div>
  <div class="form-group">
    <label asp-for="Subject" class="control-label"></label>
    <input asp-for="Subject" class="form-control" />
    <span asp-validation-for="Subject" class="text-danger"></span>
  </div>
  <div class="form-group">
    <label asp-for="Attendees" class="control-label"></label>
    <input asp-for="Attendees" class="form-control" />
    <span asp-validation-for="Attendees" class="text-danger"></span>
  </div>
  <div class="form-row">
    <div class="col">
      <div class="form-group">
        <label asp-for="Start" class="control-label"></label>
        <input asp-for="Start" class="form-control" />
        <span asp-validation-for="Start" class="text-danger"></span>
      </div>
    </div>
    <div class="col">
      <div class="form-group">
        <label asp-for="End" class="control-label"></label>
        <input asp-for="End" class="form-control" />
        <span asp-validation-for="End" class="text-danger"></span>
      </div>
    </div>
  </div>
  <div class="form-group">
    <label asp-for="Body" class="control-label"></label>
    <textarea asp-for="Body" class="form-control"></textarea>
    <span asp-validation-for="Body" class="text-danger"></span>
  </div>
  <div class="form-group">
    <input type="submit" value="Save" class="btn btn-primary" />
  </div>
</form>

@section Scripts {
    @{await Html.RenderPartialAsync("_ValidationScriptsPartial");}
}
```

### Add controller actions

#### ./Controllers/CalendarController.cs

##### render the new event form

```c#
// Minimum permission scope needed for this view
[AuthorizeForScopes(Scopes = new[] { "Calendars.ReadWrite" })]
public IActionResult New()
{
    return View();
}
```

##### receive new event from form & save through MS Graph to users calendar

```c#
[HttpPost]
[ValidateAntiForgeryToken]
[AuthorizeForScopes(Scopes = new[] { "Calendars.ReadWrite" })]
public async Task<IActionResult> New([Bind("Subject,Attendees,Start,End,Body")] NewEvent newEvent)
{
    var timeZone = User.GetUserGraphTimeZone();

    // Create a Graph event with the required fields
    var graphEvent = new Event
    {
        Subject = newEvent.Subject,
        Start = new DateTimeTimeZone
        {
            DateTime = newEvent.Start.ToString("o"),
            // Use the user's time zone
            TimeZone = timeZone
        },
        End = new DateTimeTimeZone
        {
            DateTime = newEvent.End.ToString("o"),
            // Use the user's time zone
            TimeZone = timeZone
        }
    };

    // Add body if present
    if (!string.IsNullOrEmpty(newEvent.Body))
    {
        graphEvent.Body = new ItemBody
        {
            ContentType = BodyType.Text,
            Content = newEvent.Body
        };
    }

    // Add attendees if present
    if (!string.IsNullOrEmpty(newEvent.Attendees))
    {
        var attendees =
            newEvent.Attendees.Split(';', StringSplitOptions.RemoveEmptyEntries);

        if (attendees.Length > 0)
        {
            var attendeeList = new List<Attendee>();
            foreach (var attendee in attendees)
            {
                attendeeList.Add(new Attendee{
                    EmailAddress = new EmailAddress
                    {
                        Address = attendee
                    },
                    Type = AttendeeType.Required
                });
            }
        }
    }

    try
    {
        // Add the event
        await _graphClient.Me.Events
            .Request()
            .AddAsync(graphEvent);

        // Redirect to the calendar view with a success message
        return RedirectToAction("Index").WithSuccess("Event created");
    }
    catch (ServiceException ex)
    {
        // Redirect to the calendar view with an error message
        return RedirectToAction("Index")
            .WithError("Error creating event", ex.Error.Message);
    }
}
```