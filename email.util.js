const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { getAccessTokenForRefreshToken, getAuthenticatedClient } = require('./auth.utils');


/**
 * Retrieves an email by its ID from Microsoft Graph API using MSAL.
 * @param {string} emailId - The ID of the email to retrieve.
 * @param {string} refreshToken - The refresh token to acquire a new access token.
 * @returns {Promise<Object|null>} - The email object or null if not found.
 */
const getEmailById = async (emailId, refreshToken) => {
    try {
        const accessToken = await getAccessTokenForRefreshToken(refreshToken)
        if(accessToken) {
            const client = getAuthenticatedClient(accessToken);
            const emailEndpoint = `https://graph.microsoft.com/v1.0/me/messages/${emailId}`;
            const response = await client.api(emailEndpoint).get();
            return response;
        }
        else{
            console.log(`Couldn't obtain the access token for refresh-token ${refreshToken}`);
        }
    } catch (error) {
        console.error('Error retrieving email:', error.response ? error.response.data : error.message);
    }
    return null;
};

module.exports = { getEmailById };
