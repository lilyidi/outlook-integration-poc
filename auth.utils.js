// auth.service.js

const axios = require('axios'); // Ensure axios is imported
const { Client } = require('@microsoft/microsoft-graph-client');



const msalConfig = {
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

const redirect_Host = process.env.REDIRECT_HOST

function getAuthenticatedClient(accessToken) {
    return Client.init({
        authProvider: (done) => {
            done(null, accessToken); // Pass the access token
        },
    });
}

const getAccessTokenForRefreshToken = async (refreshToken) => {
    console.log(`Getting a new access-token using the refresh-token`);
    const tokenRequest = {
        refreshToken: refreshToken,
        clientId: process.env.CLIENT_ID,
        clientSecret: process.env.CLIENT_SECRET,
        redirectUri: `${redirect_Host}/auth/callback`,
        scopes: ["user.read", "mail.readwrite", "mail.send", "mail.read", "offline_access"],
    };

    const tokenEndpoint = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('grant_type', 'refresh_token');
    params.append('refresh_token', refreshToken);
    params.append('client_id', tokenRequest.clientId);
    params.append('client_secret', tokenRequest.clientSecret);
    params.append('scope', tokenRequest.scopes.join(" "));

    try {
        const response = await axios.post(tokenEndpoint, params, {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
            }
        });
        // Update the session with the new access token
        return response.data.access_token;
    } catch (error) {
        console.log('Error during token refresh:', error.response.data);
        return null;
    }
};

module.exports = { getAccessTokenForRefreshToken, msalConfig, redirect_Host, getAuthenticatedClient };

