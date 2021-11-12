// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

var graph = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");

module.exports = {
  getUserDetails: async function (msalClient, userId) {
    const client = getAuthenticatedClient(msalClient, userId);

    const user = await client
      .api("/me")
      .select("displayName,mail,mailboxSettings,userPrincipalName")
      .get();
    return user;
  },

  // <GetCalendarViewSnippet>
  getCalendarView: async function (msalClient, userId, start, end, timeZone) {
    const client = getAuthenticatedClient(msalClient, userId);

    const events = await client
      .api("/me/calendarview")
      // Add Prefer header to get back times in user's timezone
      // .header("Prefer", `outlook.timezone="${timeZone}"`)
      // Add the begin and end of the calendar window
      .query({ startDateTime: start, endDateTime: end })
      // Get just the properties used by the app
      .select("subject,organizer,start,end")
      // Order by start time
      .orderby("start/dateTime")
      // Get at most 50 results
      .top(50)
      .get();

    return events;
  },
  // </GetCalendarViewSnippet>

  // <CreateEventSnippet>
  createEvent: async function (msalClient, userId, formData, timeZone) {
    const client = getAuthenticatedClient(msalClient, userId);

    // Build a Graph event
    const newEvent = {
      subject: formData.subject,
      start: {
        dateTime: formData.start,
        timeZone: timeZone,
      },
      end: {
        dateTime: formData.end,
        timeZone: timeZone,
      },
      body: {
        contentType: "text",
        content: formData.body,
      },
    };

    // Add attendees if present
    if (formData.attendees) {
      newEvent.attendees = [];
      formData.attendees.forEach((attendee) => {
        newEvent.attendees.push({
          type: "required",
          emailAddress: {
            address: attendee,
          },
        });
      });
    }
    console.log(JSON.stringify(newEvent));
    // POST /me/events
    await client.api("/me/events").post(newEvent);
  },
  createEventApi: async function (token, data) {
    const client = getAuthenticatedClient(null, null, token);
    // POST /me/events
    try {
      return await client.api("/me/events").post(data);
    } catch (err) {
      console.log("here");
      console.log(err);
      return {
        error: err,
      };
    }
  },
  // </CreateEventSnippet>
};

function getAuthenticatedClient(msalClient, userId, token = null) {
  // if (!msalClient || !userId) {
  //   throw new Error(
  //     `Invalid MSAL state. Client: ${
  //       msalClient ? "present" : "missing"
  //     }, User ID: ${userId ? "present" : "missing"}`
  //   );
  // }
  // Initialize Graph client
  const client = graph.Client.init({
    // Implement an auth provider that gets a token
    // from the app's MSAL instance
    authProvider: async (done) => {
      try {
        if (!token) {
          // Get the user's account
          const account = await msalClient
            .getTokenCache()
            .getAccountByHomeId(userId);

          if (account) {
            // Attempt to get the token silently
            // This method uses the token cache and
            // refreshes expired tokens as needed
            const response = await msalClient.acquireTokenSilent({
              scopes: process.env.OAUTH_SCOPES.split(","),
              redirectUri: process.env.OAUTH_REDIRECT_URI,
              account: account,
            });
            // First param to callback is the error,
            // Set to null in success case
            done(null, response.accessToken);
          }
        } else {
          done(null, token);
        }
      } catch (err) {
        console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
        done(err, null);
      }
    },
  });

  return client;
}
