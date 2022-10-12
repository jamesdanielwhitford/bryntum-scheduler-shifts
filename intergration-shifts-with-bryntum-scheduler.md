# How to connect and sync Bryntum Scheduler to a Microsoft Teams 

[Bryntum Scheduler](https://www.bryntum.com/products/scheduler/) is a fully customizable, responsive, high-performance JavaScript component that's built using ES6+ and Sass. It can easily be used with React, Vue, or Angular. In this tutorial, we'll connect and sync a Bryntum Scheduler to a Microsoft Teams. We'll do the following:

- Create a JavaScript app that a user can log in to using their Microsoft 365 Developer Program account. 
- Use Microsoft Graph to get the user's Microsoft Teams Shifts
- Display the Microsoft Teams Shifts in a Bryntum Scheduler. 
- Sync event changes in the Bryntum Scheduler with the user's Microsoft Teams. 



![Bryntum - Microsoft Outlook events sync](images/cover-image.png)

## Getting started

Clone the [starter GitHub repository](https://github.com/ritza-co/bryntum-outlook-calendar-app-starter). The starter repository uses [Vite](https://vitejs.dev/), which is a development server and JavaScript bundler. You'll need Node.js version 14.18+ for Vite to work.

Now install the Vite dev dependency by running the following command:

```bash
npm install
```

Run the local dev server using `npm run dev` and you’ll see a blank page. The dev server is configured to run on `http://localhost:8080/` in the `vite.config.js` file. This will be needed for Microsoft 365 authentication later.

Let’s create our Bryntum Calendar now.

## Creating a calendar using Bryntum

We'll install the Bryntum Calendar component using npm. Follow the guide to installing the Bryntum Calendar component [here](https://www.bryntum.com/docs/calendar/guide/Calendar/quick-start/javascript-npm).

The `style.css` file contains some basic styling for the calendar. We set the `<HTML>` and `<body>` elements to have a height of 100vh so that the Bryntum Calendar will take up the full height of the screen.

Let’s import the Bryntum Calendar component and give it some basic configuration. In the `main.js` file add the following lines:

```js
import { Calendar } from "@bryntum/calendar";
import "@bryntum/calendar/calendar.stockholm.css";

const startDate = new Date();
const endDate = new Date();
endDate.setHours(endDate.getHours() + 1);
const startDateStr = startDate.toISOString().substring(0, 19);
const endDateStr = endDate.toISOString().substring(0, 19);

const calendar = new Calendar({
  appendTo: "calendar",

  resources: [
    {
      id: 1,
      name: "Default Calendar",
      eventColor: "green",
    },
  ],
  events: [
    {
      id: 1,
      name: "Meeting",
      startDate: startDateStr,
      endDate: endDateStr,
      resourceId: 1,
    },
  ],
});
```



We imported the Bryntum Calendar and the CSS for the Stockholm theme, which is one of five available themes. You can also create custom themes. You can read more about styling the calendar [here](https://bryntum.com/docs/calendar/guide/Calendar/customization/styling). We created a new Bryntum Calendar instance and passed a configuration object into it. We added the calendar to the DOM as a child of the `<div>` element with an `id` of `"calendar"`. 

We passed in data inline to populate the Calendar Resources and events stores for simplicity. You can learn more about working with data in the [Bryntum docs](https://www.bryntum.com/docs/calendar/guide/Calendar/data/project_data). We have a single resource, the `"Default Calendar"`. Within the calendar, there's one example `"Meeting"` event. If you run your dev server now, you'll see the event in the Bryntum Calendar:

![Bryntum Calendar with example event](images/bryntum-calendar-initial.png)

Now let's learn how to retrieve a list of calendar events from a user’s Microsoft Outlook Calendar using Microsoft Graph.

## Getting access to Microsoft Graph

We're going to register a Microsoft 365 application by creating an application registration in [Azure Active Directory](https://azure.microsoft.com/en-us/products/active-directory/) (Azure AD), which is an authentication service. We'll do this so that a user can sign in to our app using their Microsoft 365 account. This will allow our app access to the data the user gives the app permission to access. A user will sign in using OAuth, which will send an access token to our app that will be stored in session storage. We'll then use the token to make authorized requests for Microsoft 365 Outlook Calendar data using Microsoft Graph. Microsoft Graph is a single endpoint REST API that enables you to access data from Microsoft 365 applications. 

To use Microsoft Graph you'll need a [Microsoft account](https://account.microsoft.com/account?lang=en-hk) and you'll need to join the [Microsoft 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program) with that Microsoft account. When joining the Microsoft 365 Developer Program, you'll be asked what areas of Microsoft 365 development you’re interested in, select the Microsoft Graph option. Choose the closest data center region, create your admin username and password, then click "Continue". Next, select “Instant Sandbox” and click “Next”. 

![Setting up your Microsoft 365 developer sandbox](images/microsoft-account-365-sandbox.png)



Now that you have successfully joined the developer program, you can get your administrator’s email address in the dashboard window. We'll use it to create an application with Microsoft Azure.

![Setting up your Microsoft 365 developer sandbox](images/microsoft-account-admin-email.png)



## Creating an Azure AD app to connect to Microsoft 365

Let's register a Microsoft 365 application by creating an application registration in the [Azure Active Directory admin portal](https://aad.portal.azure.com/). Sign in using the admin email address from your Microsoft 365 Developer Program account. Now follow these steps to create an Azure Active Directory application:

1. In the menu, select "Azure Active Directory".

![Adzure AD setup - step 1](images/azure-ad-app-1.png)



2.  Select "App registrations".

![Adzure AD setup - step 2](images/azure-ad-app-2.png)



3.  Click "New registration" to create a new app registration.

![Adzure AD setup - step 3](images/azure-ad-app-3.png)



4. Give your app a name, select the "Single tenant" option, select "Single page application" for the redirect URI, and enter http://localhost:8080 for the redirect URL. Then click the "Register" button.

![Adzure AD setup - step 4](images/azure-ad-app-4.png)



After registering your application, take note of the Application (client) ID and the Directory (tenant) ID, you'll need these to set up authentication for your web app later.

![Adzure AD setup - step 5](images/azure-ad-app-5.png)



Now we can create a JavaScript web app that can get user data using the Microsoft Graph API. The next step is to set up authentication within our web app.



## Setting up Microsoft 365 authentication in the JavaScript app

To get data using the Microsoft Graph REST API, our app needs to prove that we're the owners of the app that we just created in Azure AD. Your application will get an access token from Azure AD and include it in each request to Microsoft Graph. After this is set up, users will be able to sign in to your app using their Microsoft 365 account. This means that you won’t have to implement authentication in your app or maintain users' credentials.

![Auth flow diagram](images/auth-flow.png)

First we'll create the variables and functions we need for authentication and retrieving calendar events from Microsoft Outlook Calendar. Then we'll add the Microsoft Authentication Library and Microsoft Graph SDK, which we'll need for authentication and using Microsoft Graph. 

Create a file called `auth.js` in your project’s root directory and add the following code:

```js
const msalConfig = {
  auth: {
    clientId: "<your-client-ID-here>",
    // comment out if you use a multi-tenant AAD app
    authority: "https://login.microsoftonline.com/<your-directory-ID-here>",
    redirectUri: "http://localhost:8080",
  },
}; 
```



In the `msalConfig` variable, replace the value for `clientID` with the client ID that came with your Azure AD application and replace the `authority` value with your directory ID.

The following code will check permissions, create a Microsoft Authentication Library client, log a user in, and get the authentication token. Add it to the bottom of the file.

```js
const msalRequest = { scopes: [] };
function ensureScope(scope) {
  if (
    !msalRequest.scopes.some((s) => s.toLowerCase() === scope.toLowerCase())
  ) {
    msalRequest.scopes.push(scope);
  }
}

// Initialize MSAL client
const msalClient = new msal.PublicClientApplication(msalConfig);

// Log the user in
async function signIn() {
  const authResult = await msalClient.loginPopup(msalRequest);
  sessionStorage.setItem("msalAccount", authResult.account.username);
}

async function getToken() {
  let account = sessionStorage.getItem("msalAccount");
  if (!account) {
    throw new Error(
      "User info cleared from session. Please sign out and sign in again."
    );
  }
  try {
    // First, attempt to get the token silently
    const silentRequest = {
      scopes: msalRequest.scopes,
      account: msalClient.getAccountByUsername(account),
    };

    const silentResult = await msalClient.acquireTokenSilent(silentRequest);
    return silentResult.accessToken;
  } catch (silentError) {
    // If silent requests fails with InteractionRequiredAuthError,
    // attempt to get the token interactively
    if (silentError instanceof msal.InteractionRequiredAuthError) {
      const interactiveResult = await msalClient.acquireTokenPopup(msalRequest);
      return interactiveResult.accessToken;
    } else {
      throw silentError;
    }
  }
}
```



The `msalRequest` variable stores the current Microsoft Authentication Library request. It initially contains an empty array of scopes. The list of permissions granted to your app is part of the access token. These are the scopes of the [OAuth](https://oauth.net/) standard. When your app requests an access token from the Azure Active Directory, it needs to include a list of scopes. Each operation in Microsoft Graph has its own list of scopes. The list of the permissions required for each operation is available in the [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference).



## Using Microsoft Graph to access a user's Outlook Calendar events for the next seven days

Create a file called `graph.js` in the project’s root directory and add the following code:

```js
const userTimeZone = "Africa/Johannesburg";

const authProvider = {
  getAccessToken: async () => {
    // Call getToken in auth.js
    return await getToken();
  },
};
// Initialize the Graph client
const graphClient = MicrosoftGraph.Client.initWithMiddleware({ authProvider });

async function getEvents() {
  ensureScope("Calendars.read");
  const dateNow = new Date();
  const dateNextWeek = new Date();
  dateNextWeek.setDate(dateNextWeek.getDate() + 7);
  const query = `startDateTime=${dateNow.toISOString()}&endDateTime=${dateNextWeek.toISOString()}`;
  return await graphClient
    .api("/me/calendarView")
    .query(query)
    .header("Prefer", `outlook.timezone="${userTimeZone}"`)
    .select("subject,start,end")
    .orderby(`Start/DateTime`)
    .get();
}
```



We get the access token using the `getToken` method in the `auth.js` file. We then use the Microsoft Graph SDK (which we'll add later) to create a Microsoft Graph client that will handle Microsoft Graph API requests.

The `getEvents` function retrieves the user's calendar events for the next seven days. We use the `ensureScope` function to specify the permissions needed to access the calendar event data. We then call the `"/me/calendarview"` endpoint using the `graphClient` API  method to get the data from Microsoft Graph. We create a `query` string and use the `query` method to get calendar events for the next seven days. The `header` method allows us to set our preferred time zone. Outlook Calendar event dates are stored using UTC. We need to set the time zone for the returned Outlook Calendar event start and end dates. This is done so that the correct event times are displayed in our Bryntum Calendar. Replace the value for the `userTimeZone` with the time zone value for your region. You can find the available strings [here](https://learn.microsoft.com/en-us/graph/api/resources/datetimetimezone?view=graph-rest-1.0).

We use the `select` method to select the properties in the results that our app will use. We only request the data that we need, which improves our app's performance.

The `orderby` method specifies how to sort the result items. We order the calendar data by the `Start/DateTime` field in ascending order. Ascending order is the default if the keywords `asc` or `desc` are not specified.

Now let's add the user's Outlook Calendar events to our Bryntum Calendar. 



## Adding the Microsoft Outlook Calendar events to Bryntum Calendar

Let's add the Microsoft 365 sign in link and import the Microsoft Authentication Library and Microsoft Graph SDK. In the `index.html` file, replace the child elements of the `<body>` HTML element with the following elements:

```html
    <main id="main-container" role="main" class="container">
      <div id="content" style="display: none">
        <div id="calendar"></div>
      </div>
      <a id="signin" href="#">
        <img
          src="./images/ms-symbollockup_signin_light.png"
          alt="Sign in with Microsoft"
        />
      </a>
    </main>
    <script
      src="https://alcdn.msauth.net/browser/2.1.0/js/msal-browser.min.js"
      integrity="sha384-EmYPwkfj+VVmL1brMS1h6jUztl4QMS8Qq8xlZNgIT/luzg7MAzDVrRa2JxbNmk/e"
      crossorigin="anonymous"
    ></script>
    <script src="https://cdn.jsdelivr.net/npm/@microsoft/microsoft-graph-client/lib/graph-js-sdk.js"></script>
    <script src="auth.js"></script>
    <script src="graph.js"></script>
    <script type="module" src="main.js"></script>
```



Initially, our app will display the sign-in link only. When a user signs in, the Bryntum Calendar will be displayed. 

In the `main.js` file, add the following line to store the "Sign in with Microsoft" link element object in a variable:

```js
const signInButton = document.getElementById("signin");
```



Now add the following function at the bottom of the file:

```js
async function displayUI() {
  await signIn();

  // Hide sign in link and initial UI
  signInButton.style = "display: none";
  var content = document.getElementById("content");
  content.style = "display: block";

  // Display calendar after sign in
  var events = await getEvents();
  var calendarEvents = [];
  var eventId = 1;
  var resourceID = 1;
  events.value.forEach((event) => {
    calendarEvents.push({
      id: eventId,
      name: event.subject,
      startDate: event.start.dateTime,
      endDate: event.end.dateTime,
      resourceId: resourceID,
    });
    eventId++;
  });
  calendar.events = calendarEvents;
}

signInButton.addEventListener("click", displayUI);

export { displayUI };
```



The `displayUI` function calls the `signIn` function in `auth.js` to sign the user in. Once the user is signed in, the sign-in link is hidden and the Bryntum Calendar is displayed. We use the `getEvents` function in the `graph.js` file to get the calendar events for the next seven days. We then use the retrieved Outlook Calendar events to create calendar events for the Bryntum Calendar and add them to the `calendar.events` store.  

Now sign in to [Microsoft Outlook](https://www.microsoft.com/en-za/microsoft-365/outlook/email-and-calendar-software-microsoft-outlook) using the admin email address from your Microsoft 365 Developer Program account and create some events for the following week:

![Microsoft Outlook events](images/outlook-events.png)



Run your dev server using `npm run dev` and you'll see the sign-in link:

![Sign in link](images/sign-in-link.png)



Sign in with the same admin email address that you used to log in to Microsoft Outlook:

![Sign in pop-up](images/sign-in-popup.png)



You'll now see your Outlook Calendar events in your Bryntum Calendar:

![Bryntum Calendar events](images/bryntum-events.png)



Next, we'll sync the calendars by implementing CRUD functionality in our Bryntum Calendar. Updates to the Bryntum Calendar events will update the events in the Microsoft Outlook Calendar.



## Implementing CRUD

Now that we have connected our calendar to the Graph API, we'll implement CRUD functionality by taking advantage of Microsoft Graph's `post`, `get`, `patch`, and `delete` methods, passing a query string where relevant.

### Create events

In the `graph.js` file, add the following lines:

```js
async function createEvent(name, startDate, endDate) {
  ensureScope("Calendars.ReadWrite");
  const event = {
    subject: `${name}`,
    body: {
      contentType: "HTML",
      content: "This is a test event",
    },
    start: {
      dateTime: `${startDate.toISOString()}`,
      timeZone: userTimeZone,
    },
    end: {
      dateTime: `${endDate.toISOString()}`,
      timeZone: userTimeZone,
    },
  };
  return await graphClient.api("/me/events").post(event);
}
```



Here we create a function that will create an Outlook event with a name, start date, and end date collected from the Bryntum Calendar. The function is passed the appropriate scope and the new event data is defined. The timezone needs to be specified otherwise Microsoft Graph will assume the timezone is GMT.



### Get all events

In the `graph.js` file, add the following function:

```js
async function getAllEvents() {
  ensureScope("Calendars.ReadWrite");
  return await graphClient
    .api("/me/events")
    .select(
      "id,subject,body,bodyPreview,organizer,attendees,start,end,location"
    )
    .get();
}
```



This function is similar to the `getNextWeeksEvents` function that we created earlier. The function is passed the appropriate scope, and we pass in an appropriate query string that filters the data to return only the event data appropriate to us. However, we will get all events from the user's calendar in this case.

We need all the events from the user's calendar, as we will search through these events later on.



### Update events

In the `graph.js` file, add the following function:

```js
async function updateEvent(id, name, startDate, endDate) {
  ensureScope("Calendars.ReadWrite");
  const event = {
    subject: `${name}`,
    body: {
      contentType: "HTML",
      content: "This is a test event",
    },
    start: {
      dateTime: `${startDate.toISOString()}`,
      timeZone: userTimeZone,
    },
    end: {
      dateTime: `${endDate.toISOString()}`,
      timeZone: userTimeZone,
    },
  };
  return await graphClient.api(`/me/events/${id}`).patch(event);
}
```



The `updateEvent` function will identify the appropriate Outlook event by `id`, and then it will use the new name, start date, and end date from the Bryntum Calendar to update the event. The function is passed the appropriate scope, and the new event data is defined.



### Delete events

In the `graph.js` file, add the following function:

```js
async function deleteEvent(id) {
  ensureScope("Calendars.ReadWrite");
  return await graphClient.api(`/me/events/${id}`).delete();
}
```



The `deleteEvent` function will identify the appropriate Outlook event by `id`, and delete the event.



### Listening for event data changes in the Bryntum Calendar

Next, we'll set the listeners for our Bryntum Calendar so that it will know when the user updates the calendar events.

Replace the definition of `calendar` with the following code:

```js
const calendar = new Calendar({
  appendTo: "calendar",

  listeners: {
    dataChange: function (event) {
      updateMicrosoft(event);
    },
  },

  resources: [
    {
      id: 1,
      name: "Default Calendar",
      eventColor: "green",
    },
  ],
});
```



Here we set a listener on our Bryntum Calendar to listen for any changes to the Bryntum Calendar's data store. This will fire an event called `"update"` whenever a calendar event is created or updated, and an event called `"remove"` whenever an event is deleted.

The event that's retrieved from the `dataChange` listener will also carry event data about the specific calendar event that has been altered. We'll use the event data to identify which event is being altered and what's being changed.



In the `main.js` file, remove the following variables as we don't need them anymore:

```js
const startDate = new Date();
const endDate = new Date();
endDate.setHours(endDate.getHours() + 1);
const startDateStr = startDate.toISOString().substring(0, 19);
const endDateStr = endDate.toISOString().substring(0, 19);
```



Next we'll create a function called `updateMicrosoft` that will update the Outlook Calendar when the appropriate `"update"` or `"delete"` event is fired.

Add the following code below the definition of `calendar` in the `main.js` file:

```js
async function updateMicrosoft(event) {
  if (event.action == "update") {
    const microEvents = await getAllEvents();
    // check if event exists in microsoft, if it does, update it, if not, create it
    var eventExists = false;

    for (var i = 0; i < microEvents.value.length; i++) {
      // event exists in both microsoft and bryntum with the same name
      if (microEvents.value[i].subject == event.record.name) {
        eventExists = true;
        updateEvent(
          microEvents.value[i].id,
          event.record.name,
          event.record.startDate,
          event.record.endDate
        );
        return;
      } else if ("name" in event.changes) {
        if (event.changes.name.oldValue == microEvents.value[i].subject) {
          eventExists = true;
          updateEvent(
            microEvents.value[i].id,
            event.record.name,
            event.record.startDate,
            event.record.endDate
          );
          return;
        }
      } else if ("resourceId" in event.changes) {
        eventExists = true;
      }
    }
    // event does not exist in microsoft, create it
    if (!eventExists) {
      if (event.record.name != undefined) {
        createEvent(
          event.record.name,
          event.record.startDate,
          event.record.endDate
        );
      }
    }
  }
  // event is deleted
  else if (event.action == "remove") {
    const microEvents = await getAllEvents();
    var eventName = event.records[0].data.name;
    for (var i = 0; i < microEvents.value.length; i++) {
      if (microEvents.value[i].subject == eventName) {
        deleteEvent(microEvents.value[i].id);
        return;
      }
    }
  }
}
```



Here we create a function that is called on all changes to the data store of the Bryntum Calendar. The function then calls one of the Microsoft Graph CRUD functions that we defined.

On `"update"` we call `getAllEvents` to get a list of all the user's Outlook Calendar events. We then loop through this list and compare the Bryntum Calendar event data from the `"update"` event to the Microsoft events.

```js
async function updateMicrosoft(event) {
  if (event.action == "update") {
    const microEvents = await getAllEvents();
    // check if event exists in microsoft, if it does, update it, if not, create it
    var eventExists = false;

    for (var i = 0; i < microEvents.value.length; i++) {
```



If there are events that match by name, then that Microsoft event needs to be updated.

```js
      // event exists in both microsoft and bryntum with the same name
      if (microEvents.value[i].subject == event.record.name) {
        eventExists = true;
        updateEvent(
          microEvents.value[i].id,
          event.record.name,
          event.record.startDate,
          event.record.endDate
        );
        return;
```



But if the name has been changed, there won't be a match. We need to check if any of the Microsoft event names match the name found in the "oldData" of the `update` data. This `oldName` data represents the name before the change made by `update`.

If a Microsoft event name matches the name in the "oldData", then that Microsoft event needs to be updated to match the name.

```js
      } else if ("name" in event.changes) {
        if (event.changes.name.oldValue == microEvents.value[i].subject) {
          eventExists = true;
          updateEvent(
            microEvents.value[i].id,
            event.record.name,
            event.record.startDate,
            event.record.endDate
          );
          return;
```



Next, we check the `event.changes` for a resource ID. If this resource ID exists, then the change has already been made and this `update` event should be ignored.

```js
      } else if ("resourceId" in event.changes) {
        eventExists = true;
      }
```



If the `update` passes any of these tests, then `eventExists` will be set to `true`. If not, then it will be set to `false` and the `createEvent` function will be triggered.

If no Microsoft event matches it, it's because a new event was created in the Bryntum Calendar, so we create a new Outlook event using this new data.

```js
    // event does not exist in microsoft, create it
    if (!eventExists) {
      if (event.record.name != undefined) {
        createEvent(
          event.record.name,
          event.record.startDate,
          event.record.endDate
        );
      }
    }
```



Finally, if the `dataChange` event is a `"remove"` event, then we delete the matching Outlook event using the `deleteEvent` function.

```js
  // event is deleted
  else if (event.action == "remove") {
    const microEvents = await getAllEvents();
    var eventName = event.records[0].data.name;
    for (var i = 0; i < microEvents.value.length; i++) {
      if (microEvents.value[i].subject == eventName) {
        deleteEvent(microEvents.value[i].id);
        return;
      }
    }
  }
```



Now try to create, update, delete, and edit an event in the Bryntum Calendar. You'll see the changes reflected in the Outlook Calendar.

![Updating Outlook calendar using CRUD](images/crud.gif)



## Next steps

This tutorial gives you a starting point for creating a Bryntum Calendar using vanilla JavaScript and syncing it with Microsoft Outlook. There are many ways that you can improve the Bryntum Calendar. For example, you can add features such as [per-resource calendar views](https://bryntum.com/examples/calendar/resourceview/). Take a look at our [Calendar demos page](https://bryntum.com/examples/calendar/) to see demos of the available features.
