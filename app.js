const express = require('express');
const session = require('express-session');
const axios = require('axios');
const bodyParser = require('body-parser');
const { Client } = require('@microsoft/microsoft-graph-client');
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
    code: req.query.code,
    scopes: ["user.read", "mail.readwrite"],
    redirectUri:`${redirectHost}/auth/callback`,
  };

  

  msalClient.acquireTokenByCode(tokenRequest).then((response) => {
    if (response.account) {
      console.log('User ID:', response.account.homeAccountId);
    }
    console.log(`Response is ${JSON.stringify(response)}`);
    req.session.accessToken = response.accessToken;
    createSubscription(response.accessToken, response.account.homeAccountId, response.account.username);
    res.redirect('/emails');
  }).catch((error) => {
    console.log(error);
    res.status(500).send(error);
  });
});

app.post('/webhook', (req, res) => {
    if (req.query.validationToken) {
      res.send(req.query.validationToken);  // Respond with the validation token
    } else {
      const notification = req.body
      if (notification) {
        console.log(`New email received or updated: ${JSON.stringify(notification)}` );
      }
      res.status(202).send();  // Acknowledge the notification
    }
});

// Get emails
app.get('/emails', async (req, res) => {
  if (!req.session.accessToken) {
    return res.redirect('/login');
  }
  const client = getAuthenticatedClient(req.session.accessToken);
  try {
    let messages = [];
    let nextPageUrl = `/me/messages`;
    while (nextPageUrl) {
      const response = await client.api(nextPageUrl).get();
      messages = messages.concat(response.value);
      nextPageUrl = response['@odata.nextLink'];
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
