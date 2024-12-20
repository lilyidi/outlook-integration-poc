const express = require('express');
const multer = require("multer");
const session = require('express-session');
const axios = require('axios');
const bodyParser = require('body-parser');
const { Client } = require('@microsoft/microsoft-graph-client');
const fs = require('fs');
const path = require('path');
require('isomorphic-fetch');
require('dotenv').config();
const { ConfidentialClientApplication } = require('@azure/msal-node');
const {getAccessTokenForRefreshToken, msalConfig, redirect_Host, getAuthenticatedClient} = require('./auth.utils.js');

const app = express();
app.use(bodyParser.json());
const port = 3000;

// Configure multer for file uploads
const upload = multer({ dest: './uploads' }); // Temp directory
let uploadedFile = null; // Temporary variable to store uploaded file info

const config = msalConfig;
const redirectHost = redirect_Host
const msalClient = new ConfidentialClientApplication(config);

const createSubscription = async (req) => {
  const accessToken = req.session.accessToken;
  const userInfoResponse = await axios.get('https://graph.microsoft.com/v1.0/me', {
    headers: {
        'Authorization': `Bearer ${accessToken}`,
    }
  });
  const emailAddress = userInfoResponse.data.mail || userInfoResponse.data.userPrincipalName; // Use mail if available, otherwise userPrincipalName
  const domainname = `${process.env.DOMAIN_NAME}`;
  const eventhubnamespace = `${process.env.EVENTHUB_NAMESPACE}`;
  const eventhubname = `${process.env.EVENTHUB_NAME}`;
  try {
    const client = getAuthenticatedClient(accessToken);
    const subscription = {
      changeType: "created,updated",
      notificationUrl: `EventHub:https://${eventhubnamespace}.servicebus.windows.net/eventhubname/${eventhubname}?tenantId=${domainname}`,
      lifecycleNotificationUrl: `EventHub:https://${eventhubnamespace}.servicebus.windows.net/eventhubname/${eventhubname}?tenantId=${domainname}`,
      resource: "me/messages",
      expirationDateTime: new Date(Date.now() + 4230 * 60 * 1000).toISOString(), // 3 days from now.
      clientState: `${emailAddress}`
    };
    const response = await client.api('/subscriptions')
      .post(subscription);
    console.log('Subscription created:', response);
  } catch (error) {
    console.error('Error creating subscription:', error.response.data);
  }
};

app.post('/renew-subscription', async (req, res) => {
  const events = req.body.value;
    if (events === undefined){
      console.log("Received empty body");
      res.send(req.query.validationToken);  // Respond with the validation token
    } else {
      console.log("received event", events);
      for (let ev of events) {
        if (ev.lifecycleEvent ==="reauthorizationRequired") {
          console.log("Received reauthorizationRequired event for subscription:", ev.subscriptionId);
          const response = axios.patch(
              `https://graph.microsoft.com/v1.0/subscriptions/${ev.subscriptionId}`,
              {
                  expirationDateTime: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString()
              },
              {
                  headers: {
                      Authorization: `Bearer ${req.session.accessToken}`,
                      'Content-Type': 'application/json'
                  }
              }
            )
          console.log("Reauthorization required, extending subscription expiration time by 1 day", response.data);
        }
      }
    }
});

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
    redirectUri: `${redirectHost}/outlook-integration/auth/callback`,
  };

  msalClient.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
    res.redirect(response);
  }).catch((error) => console.log(JSON.stringify(error)));
});

// Auth callback route
app.get('/outlook-integration/auth/callback', async(req, res) => {
  const tokenRequest = {
    code: req.query.code,  // Authorization code received from /authorize
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    redirectUri: `${redirectHost}/outlook-integration/auth/callback`,  // Must match the one registered in Azure
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
  }).then(async (response) => {
    // Save the tokens in session (or wherever needed)
    req.session.accessToken = response.data.access_token;
    req.session.refreshToken = response.data.refresh_token;
    console.log(`refreshToken is ${req.session.refreshToken}`);
    // await createSubscription(req);
    res.redirect('/emails'); 
  }).catch((error) => {
    console.log('Error during token exchange:', error);
    res.status(500).send(error);
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
  if(req.session.refreshToken) {
    req.session.accessToken = await getAccessTokenForRefreshToken(req.session.refreshToken);
      if(!req.session.accessToken) {
        return res.redirect('/login');
      }
  }
  else {
    return res.redirect('/login');
  }
  const showGrouped = req.query.showGrouped || 'false';
  const numberOfDaysSince = req.query.numberOfDaysSince || 100;
  const client = getAuthenticatedClient(req.session.accessToken);
  try {
    let messages = [];
    const queryOptions = {
      '$select': 'id,subject,body,bodyPreview,conversationId,conversationIndex,internetMessageHeaders,receivedDateTime,from,toRecipients',
      '$expand': 'attachments',
      '$filter' : `receivedDateTime ge ${new Date(new Date().setDate(new Date().getDate() - numberOfDaysSince)).toISOString()}`,
      '$top': 5 // Limit the number of messages to fetch
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
      res.render('grouped-email', { groupedMessages: groupedMessages });
      return;
    }
    res.render('emails', { emails: messages });
  } catch (error) {
    res.status(500).send(error);
    console.error(error);
  }
});

app.post('/upload', upload.single('file'), async (req, res) => { 
  if (!req.file) {
    return res.status(400).send("No file uploaded.");
  }

  // Store the uploaded file info
  uploadedFile = req.file;
  res.status(200).send(`File uploaded: ${req.file.originalname}`);
});

// Send email
app.post('/send', async (req, res) => {
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
            address: 'lily@regalvoice.com'
          }
        }
      ],
      attachments: [],
    }
  };

  const client = getAuthenticatedClient(req.session.accessToken);

  const draftResponse = await client.api("/me/messages/AQMkADAwATY0MDABLTQ3OTQtZWJkMy0wMAItMDAKAEYAAAMEx9FjT3P0SbDNJcxKIVGyBwAYNk85euZ7RKKsocScIeKlAAACAQwAAAAYNk85euZ7RKKsocScIeKlAAAAFstGmQAAAA==/createReply")
  .post({});
  const messageId = draftResponse.id;

  if (uploadedFile) {
    const filePath = uploadedFile.path;
    const fileName = uploadedFile.originalname;
    if (!fs.existsSync(filePath)) {
      return res.status(400).send("File not found.");
    }
    // Read the file for attachment
    const fileContent = fs.readFileSync(filePath);
    const fileSize = fileContent.length;

    if (uploadedFile.size > 3 * 1024 * 1024) {
        const uploadSession = await client.api(`/me/messages/${messageId}/attachments/createUploadSession`).post({
            AttachmentItem: {
                attachmentType: "file",
                name: fileName,
                size: fileSize,
                contentType: uploadedFile.mimeType,
            },
        });
        const uploadUrl = uploadSession.uploadUrl;
        let bytesUploaded = 0;
        const readStream = fs.createReadStream(uploadedFile.path);

        for await (const chunk of readStream) {
          console.log(`Uploading ${chunk.length} bytes...`);
          const start = bytesUploaded;
          const end = Math.min(start + chunk.length-1, fileSize-1);

          try {
            const response = await axios.put(uploadUrl, chunk, {
                headers: {
                    "Content-Range": `bytes ${start}-${end}/${fileSize}`,
                    "Content-Type": "application/octet-stream",
                    "Content-Length": chunk.length,
                },
            });
            console.log(response.data);
          } catch (error) {
            console.log(error.response.data);
          }

          bytesUploaded += chunk.length;
        }
        console.log("Upload complete.");
    } else {
      console.log("adding file to email");
      await client.api(`/me/messages/${encodeURIComponent(messageId)}`).update(mail.message);
      await client.api(`/me/messages/${messageId}/attachments`).post([{
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: fileName,
        contentBytes: fileContent.toString("base64"),
        contentType: uploadedFile.mimeType,
      }]);
    }
    console.log('Sending email');
    await client.api(`/me/messages/${messageId}/send`).post({})
    .then(response => {
      // Delete file after sending
      fs.unlinkSync(filePath);
      res.redirect('/emails');
    }).catch(error => {
      console.log(error);
      res.status(500).send(error);
    });
  }
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
      const newExpirationDateTime = new Date(Date.now() + 4230 * 60 * 1000).toISOString();
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
