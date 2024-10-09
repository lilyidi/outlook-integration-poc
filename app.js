const express = require('express');
const session = require('express-session');
const axios = require('axios');
const bodyParser = require('body-parser');
const { Client } = require('@microsoft/microsoft-graph-client');
const fs = require('fs');
const path = require('path');
require('isomorphic-fetch');
require('dotenv').config();
const { ConfidentialClientApplication } = require('@azure/msal-node');

const app = express();
app.use(bodyParser.json());
const port = 3000;

// Microsoft App Credentials
const config = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: 'https://login.microsoftonline.com/common',
    clientSecret: process.env.CLIENT_SECRET,
  },
  system: {
    loggerOptions: {
      loggerCallback(loglevel, message) {
        console.log(message);
      },
      piiLoggingEnabled: false,
      logLevel: "Info",
    },
  },
};


function getAuthenticatedClient(accessToken) {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken); // Pass the access token
    },
  });
}


const redirectHost = process.env.REDIRECT_HOST

const msalClient = new ConfidentialClientApplication(config);

const createSubscription = async (accessToken, homeAccountId, username) => {
  // const accessToken = 'YOUR_ACCESS_TOKEN';  // Get the access token using OAuth flow
  const subscriptionData = {
    "changeType": "created,updated",
    "notificationUrl": `${redirectHost}/webhook`,
    "resource": "me/messages",
    "expirationDateTime": new Date(Date.now() + 30000).toISOString(),  // 5 minutes from now
    "clientState": `${username}`
  };

  try {
    const response = await axios.post('https://graph.microsoft.com/v1.0/subscriptions', subscriptionData, {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });

    console.log('Subscription created:', response.data);
  } catch (error) {
    console.error('Error creating subscription:', error.response.data);
  }
};

// Session setup
app.use(session({
  secret: 'secret-key',
  resave: false,
  saveUninitialized: true,
}));

app.set('view engine', 'ejs');

//TODO need to validate this
app.get('/refresh-token', (req, res) => {
  if (!req.session.refreshToken) {
    return res.redirect('/login');
  }

  const refreshToken = req.session.refreshToken;
  const tokenRequest = {
    refreshToken: refreshToken,
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    redirectUri: `${redirectHost}/auth/callback`,
    scopes: ["user.read", "mail.readwrite", "mail.send", "mail.read", "offlince_access"],
  };

  const tokenEndpoint = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;

  const params = new URLSearchParams();
  params.append('client_id', tokenRequest.clientId);
  params.append('client_secret', tokenRequest.clientSecret);  // Required for backend apps
  params.append('grant_type', 'refresh_token');
  params.append('refresh_token', tokenRequest.refreshToken);  // The refresh token
  params.append('redirect_uri', tokenRequest.redirectUri);
  params.append('scope', tokenRequest.scopes.join(" "));  // Space-separated list of scopes

  axios.post(tokenEndpoint, params, {
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    }
  }).then((response) => {
    // Save the new access token in session (or wherever needed)
    req.session.accessToken = response.data.access_token;
    console.log('Access token refreshed:', response.data.access_token);
  }).catch((error) => {
    console.error('Error refreshing access token:', error.response.data);
    return res.redirect('/login');
  });
  const authCodeUrlParameters = {
    scopes: ["user.read", "mail.readwrite", "mail.send", "mail.read"],
    redirectUri: `${redirectHost}/auth/callback`,
  };

  msalClient.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
    res.redirect(response);
  }).catch((error) => console.log(JSON.stringify(error)));
});

// Login Route
app.get('/login', (req, res) => {
  const authCodeUrlParameters = {
    scopes: ["user.read", "mail.readwrite", "mail.send", "mail.read"],
    redirectUri: `${redirectHost}/auth/callback`,
  };

  msalClient.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
    res.redirect(response);
  }).catch((error) => console.log(JSON.stringify(error)));
});

// Auth callback route
app.get('/auth/callback', (req, res) => {
  const tokenRequest = {
    code: req.query.code,  // Authorization code received from /authorize
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    redirectUri: `${redirectHost}/auth/callback`,  // Must match the one registered in Azure
    scopes: ["user.read", "mail.readwrite", "offline_access"],  // Requested scopes
  };

  const tokenEndpoint = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;

  const params = new URLSearchParams();
  params.append('client_id', tokenRequest.clientId);
  params.append('client_secret', tokenRequest.clientSecret);  // Required for backend apps
  params.append('grant_type', 'authorization_code');
  params.append('code', tokenRequest.code);  // The authorization code
  params.append('redirect_uri', tokenRequest.redirectUri);
  params.append('scope', tokenRequest.scopes.join(" "));  // Space-separated list of scopes

  axios.post(tokenEndpoint, params, {
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    }
  }).then((response) => {
    // Save the tokens in session (or wherever needed)
    req.session.accessToken = response.data.access_token;
    req.session.refreshToken = response.data.refresh_token;
    res.redirect('/emails');
  }).catch((error) => {
    console.log('Error during token exchange:', error.response.data);
    res.status(500).send(error.response.data);
  });
});
app.post('/webhook', (req, res) => {
  if (req.query.validationToken) {
    res.send(req.query.validationToken);  // Respond with the validation token
  } else {
    const notification = req.body
    if (notification) {
      console.log(`New email received or updated: ${JSON.stringify(notification)}`);
    }
    res.status(202).send();  // Acknowledge the notification
  }
});

// Get emails
app.get('/emails', async (req, res) => {
  if (!req.session.accessToken) {
    return res.redirect('/login');
  }
  const showGrouped = req.query.showGrouped || 'false';
  const client = getAuthenticatedClient(req.session.accessToken);
  try {
    let messages = [];
    const queryOptions = {
      '$select': 'id,subject,body,bodyPreview,conversationId,conversationIndex,internetMessageHeaders,receivedDateTime,from,toRecipients',
      '$expand': 'attachments',
      '$top': 25 // Limit the number of messages to fetch
    };
    let nextPageUrl = `/me/mailFolders/inbox/messages?${new URLSearchParams(queryOptions).toString()}`;
    while (nextPageUrl) {
      const response = await client.api(nextPageUrl).get();
      messages = messages.concat(response.value);
      nextPageUrl = response['@odata.nextLink'];
    }

    if (showGrouped == 'true') {
      messages = messages.sort((a, b) => new Date(a.receivedDateTime) - new Date(b.receivedDateTime));
      let preProcessDict = {}
      for (let msg of messages) {
        if (msg.internetMessageHeaders && msg.internetMessageHeaders.length > 0) {
          const messageIdHeader = msg.internetMessageHeaders.find(header => header.name === "Message-ID");
          if (messageIdHeader) {
            preProcessDict[messageIdHeader.value] = {}
            preProcessDict[messageIdHeader.value].msg = msg;
            const referencesHeader = msg.internetMessageHeaders.find(header => header.name === "References");
            if (referencesHeader) {
              preProcessDict[messageIdHeader.value].references = referencesHeader.value;
            }
          }
        }
      }

      let threadedDict = {}
      Object.entries(preProcessDict).forEach(([key, value]) => {
        if (!value.references) {
          value.threadId = key;
        }
        else {
          const references = value.references.split(' ');
          let foundThreadId = null;
          for (const ref of references) {
            if (preProcessDict[ref] && preProcessDict[ref].threadId) {
              foundThreadId = preProcessDict[ref].threadId;
              break;
            }
          }
          if (foundThreadId) {
            value.threadId = foundThreadId;
          } else {
            value.threadId = key; // If no threadId found, use the current message ID as threadId
          }
        }
      });

      let groupedMessages = {};
      Object.entries(preProcessDict).forEach(([key, value]) => {
        const threadId = value.threadId;
        if (!groupedMessages[threadId]) {
          groupedMessages[threadId] = [];
        }
        groupedMessages[threadId].push(value.msg);
      });
      // console.log(`pre-process-dict`);
      // const filePath1 = path.join(__dirname, `email-${req.session.username}.json`);
      // fs.writeFileSync(filePath1, JSON.stringify(groupedMessages));
      res.render('grouped-email', { groupedMessages: groupedMessages });
      return;
    }
    res.render('emails', { emails: messages });
  } catch (error) {
    res.status(500).send(error);
    console.error(error);
  }
});


// Send email
app.post('/send-email', (req, res) => {
  if (!req.session.accessToken) {
    return res.redirect('/login');
  }
  const mail = {
    message: {
      subject: `Sample Email Sent : ${new Date().toLocaleString()}`,
      body: {
        contentType: 'Text',
        content: 'hello, this is a test email outlook-integration-poc'
      },
      toRecipients: [
        {
          emailAddress: {
            address: 'periyv-triage-test-aaaan4kkn7epczwe7a6n4tblxu@regalvoice.slack.com'
          }
        }
      ]
    },
    saveToSentItems: "true"
  };
  const client = getAuthenticatedClient(req.session.accessToken)
  client.api('/me/sendMail').post(mail).then(response => {
    res.redirect('/emails');
  }).catch(error => {
    console.log(error);
    res.status(500).send(error);
  });
});

app.post('/list-subs', async (req, res) => {
  if (!req.session.accessToken) {
    return res.redirect('/login');
  }
  const client = getAuthenticatedClient(req.session.accessToken);
  try {
    let messages = [];
    let nextPageUrl = `/subscriptions`;
    while (nextPageUrl) {
      const response = await client.api(nextPageUrl).get();
      messages = messages.concat(response.value);
      nextPageUrl = response['@odata.nextLink'];
    }
    console.log(`subscriptions are ${JSON.stringify(messages)}`);
    res.render('subscriptions', { subscriptions: messages });
  } catch (error) {
    res.status(500).send(error);
    console.error(error);
  }
});

app.post('/extend-subs', async (req, res) => {
  if (!req.session.accessToken) {
    return res.redirect('/login');
  }
  const client = getAuthenticatedClient(req.session.accessToken);
  try {
    let messages = [];
    let nextPageUrl = `/subscriptions`;
    while (nextPageUrl) {
      const response = await client.api(nextPageUrl).get();
      messages = messages.concat(response.value);
      nextPageUrl = response['@odata.nextLink'];
    }
    console.log(`subscriptions are ${JSON.stringify(messages)}`);

    for (let message of messages) {
      const newExpirationDateTime = new Date(Date.now() + 3600 * 1000).toISOString();
      try {
        const updatedSubscription = await client.api(`/subscriptions/${message.id}`)
          .patch({
            expirationDateTime: newExpirationDateTime
          });

        console.log('Subscription renewed successfully:', updatedSubscription);
      } catch (error) {
        console.error('Error renewing subscription:', error);
      }
    }

    res.render('emails', { emails: messages });
  } catch (error) {
    res.status(500).send(error);
    console.error(error);
  }
});

// Start server
app.listen(port, () => {
  if (!process.env.CLIENT_ID || !process.env.CLIENT_SECRET || !process.env.REDIRECT_HOST) {
    console.error('One or more environment variables are not set.');
    process.exit(1);
  }
  console.log(`App listening at http://localhost:${port}`);
});
