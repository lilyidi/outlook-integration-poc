const { EventHubConsumerClient, earliestEventPosition } = require("@azure/event-hubs");
require('dotenv').config();
const axios = require('axios');
const {getEmailById} = require('./email.util');

const consumerGroup = process.env.CONSUMER_GROUP;
const connectionString = process.env.EVENTHUB_CONNECTION_STRING;
const eventHubName = process.env.EVENTHUB_NAME;
const refreshToken = process.env.SAMPLE_REFRESH_TOKEN;
const accessToken = process.env.ACCESS_TOKEN;

async function main() {
    const client = new EventHubConsumerClient(consumerGroup, connectionString, eventHubName);
    console.log(`Subscribing to events... ${eventHubName}, ${consumerGroup}, ${connectionString}`);
    // Subscribe to events
    const subscription = client.subscribe({
        processEvents: async (events, context) => {
            if (events.length === 0) {
                console.log(`No events received.`);
                return;
            }
            for (const event of events) {
                const eventBody = event.body;
                const eventValue = eventBody?.value;
                if (eventValue) {
                    for (const evtVal of eventValue) {
                        if(evtVal?.resourceData["@odata.type"]  === "#Microsoft.Graph.Message") {
                            if (evtVal?.clientState === "ltest2038@hotmail.com") {
                                console.log(`Email Data is ${JSON.stringify(evtVal)}`);
                                const emailData = await getEmailById(evtVal?.resourceData?.id, refreshToken);
                                console.log(`Email Data is ${JSON.stringify(emailData.attachments)}`);
                            }
                        } else if (evtVal?.resourceData["@odata.type"]  === "#microsoft.graph.subscription") {
                            if (evtVal?.lifecycleEvent ==="reauthorizationRequired") {
                                // console.log("Reauthorization required, extending subscription expiration time by 1 day");
                                const response = await axios.patch(
                                    `https://graph.microsoft.com/v1.0/subscriptions/${evtVal.subscriptionId}`,
                                    {
                                        expirationDateTime: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString()
                                    },
                                    {
                                        headers: {
                                            Authorization: `Bearer ${accessToken}`,
                                            'Content-Type': 'application/json'
                                        }
                                    }
                                )
                                // console.log("Finished extending subscription expiration time by 1 day");
                                // console.log(response.data);
                            }
                        } else if(evtVal?.subscriptionId !== 'NA') {
                            // const emailData = await getEmailById(evtVal?.resourceData?.id, refreshToken);
                            // console.log(`Email Data is ${JSON.stringify(emailData)}`);
                        }
                    }
                }
            }
        },
        processError: async (err, context) => {
            // console.error(`Error: ${err}`);
        }
    }
        , { startPosition: earliestEventPosition }
    );
}

main().catch((err) => {
    console.error("Error running sample:", err);
});