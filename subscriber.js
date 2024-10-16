const { EventHubConsumerClient, earliestEventPosition } = require("@azure/event-hubs");
require('dotenv').config();
const {getEmailById} = require('./email.util');

const consumerGroup = process.env.CONSUMER_GROUP;
const connectionString = process.env.EVENTHUB_CONNECTION_STRING;
const eventHubName = process.env.EVENTHUB_NAME;
const refreshToken = process.env.SAMPLE_REFRESH_TOKEN;

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
                console.log(`Received event: ${JSON.stringify(event.body)}`);
                const eventBody = event.body;
                const eventValue = eventBody?.value;
                if (eventValue) {
                    for (const evtVal of eventValue) {
                        if(evtVal?.subscriptionId !== 'NA') {
                            const emailData = await getEmailById(evtVal?.resourceData?.id, refreshToken);
                            console.log(`Email Data is ${JSON.stringify(emailData)}`);
                        }
                    }
                }
            }
        },
        processError: async (err, context) => {
            console.error(`Error: ${err}`);
        }
    }
        , { startPosition: earliestEventPosition }
    );
}

main().catch((err) => {
    console.error("Error running sample:", err);
});
