var graph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const fs = require('fs');

module.exports = {
    getUserDetails: async function (msalClient, userId) {
        const client = getAuthenticatedClient(msalClient, userId);

        const user = await client
            .api('/me')
            .select('displayName,mail,mailboxSettings,userPrincipalName')
            .get();
        return user;
    },
    getUserAvatar: async function (msalClient, loginUserId, userIds) {
        async function streamToString(readableStream) {
            return new Promise((resolve, reject) => {
                const chunks = [];
                readableStream.on("data", (data) => {
                    chunks.push(data);
                });
                readableStream.on("end", () => {
                    resolve(Buffer.concat(chunks));
                });
                readableStream.on("error", reject);
            });
        }

        const client = getAuthenticatedClient(msalClient, loginUserId);

        let avatars = {};
        for (const userId of userIds) {
            const blob = await client
                .api(`/users/${userId}/photos('120x120')/$value`)
                .get();

            const data = await streamToString(blob.stream());
            const blobString = Buffer.from(data).toString('base64');
            avatars[userId] = `data:image/jpeg;base64,${blobString}`;
        }
        return avatars;
    },

    getCalendarView: async function (msalClient, userId, start, end, timeZone) {
        const client = getAuthenticatedClient(msalClient, userId);

        const events = await client
            .api('/me/calendarview')
            // Add Prefer header to get back times in user's timezone
            .header("Prefer", `outlook.timezone="${timeZone}"`)
            // Add the begin and end of the calendar window
            .query({ startDateTime: start, endDateTime: end })
            // Get just the properties used by the app
            .select('subject,organizer,start,end')
            // Order by start time
            .orderby('start/dateTime')
            // Get at most 50 results
            .top(50)
            .get();

        return events;
    },
};

function getAuthenticatedClient(msalClient, userId) {
    if (!msalClient || !userId) {
        throw new Error(
            `Invalid MSAL state. Client: ${msalClient ? 'present' : 'missing'}, User ID: ${userId ? 'present' : 'missing'}`);
    }

    // Initialize Graph client
    const client = graph.Client.init({
        // Implement an auth provider that gets a token
        // from the app's MSAL instance
        authProvider: async (done) => {
            try {
                // Get the user's account
                const account = await msalClient
                    .getTokenCache()
                    .getAccountByHomeId(userId);

                if (account) {
                    // Attempt to get the token silently
                    // This method uses the token cache and
                    // refreshes expired tokens as needed
                    const response = await msalClient.acquireTokenSilent({
                        scopes: process.env.OAUTH_SCOPES.split(','),
                        redirectUri: process.env.OAUTH_REDIRECT_URI,
                        account: account
                    });

                    // First param to callback is the error,
                    // Set to null in success case
                    done(null, response.accessToken);
                }
            } catch (err) {
                console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
                done(err, null);
            }
        }
    });

    return client;
}