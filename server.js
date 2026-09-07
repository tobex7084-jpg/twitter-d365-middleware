const express = require('express');
const crypto = require('crypto');

const app = express();

app.use(
    express.json({
        limit: '2mb'
    })
);

/*
 * Required Render environment variables.
 */
const REQUIRED_ENVIRONMENT_VARIABLES = [
    'X_CONSUMER_SECRET',
    'D365_TENANT_ID',
    'D365_CLIENT_ID',
    'D365_CLIENT_SECRET',
    'D365_ORGANIZATION_ID',
    'D365_CHANNEL_ID',
    'D365_MESSAGING_BASE_URL',
    'D365_TOKEN_SCOPE'
];

/*
 * Check that the required environment variables exist.
 * Secret values are not displayed in the logs.
 */
function validateEnvironmentVariables() {
    const missingVariables =
        REQUIRED_ENVIRONMENT_VARIABLES.filter(
            variableName => !process.env[variableName]
        );

    if (missingVariables.length > 0) {
        console.error(
            'Missing environment variables:',
            missingVariables.join(', ')
        );

        return false;
    }

    console.log(
        'Required environment variables are configured.'
    );

    return true;
}

/*
 * Remove trailing slashes from a URL.
 */
function removeTrailingSlash(value) {
    return String(value || '').replace(/\/+$/, '');
}

/*
 * Retrieve the first available value from one of the
 * supplied object paths.
 */
function readFirstValue(object, paths) {
    for (const path of paths) {
        const value = path
            .split('.')
            .reduce(
                (currentValue, propertyName) => {
                    if (
                        currentValue &&
                        Object.prototype.hasOwnProperty.call(
                            currentValue,
                            propertyName
                        )
                    ) {
                        return currentValue[propertyName];
                    }

                    return undefined;
                },
                object
            );

        if (
            value !== undefined &&
            value !== null &&
            String(value).trim() !== ''
        ) {
            return value;
        }
    }

    return undefined;
}

/*
 * Extract useful information from an X Direct Message event.
 *
 * Multiple object paths are supported because the payload can
 * vary depending on the X API event format.
 */
function extractXDirectMessage(body) {
    const eventType = readFirstValue(body, [
        'event_type',
        'data.event_type',
        'type'
    ]);

    const messageText = readFirstValue(body, [
        'message_data.text',
        'data.message_data.text',
        'data.text',
        'text',
        'direct_message_events.0.message_create.message_data.text'
    ]);

    const senderId = readFirstValue(body, [
        'sender_id',
        'message_data.sender_id',
        'data.sender_id',
        'data.sender.id',
        'data.author_id',
        'direct_message_events.0.message_create.sender_id'
    ]);

    const senderUsername = readFirstValue(body, [
        'sender_username',
        'message_data.sender_username',
        'data.sender_username',
        'data.sender.username'
    ]);

    const senderName = readFirstValue(body, [
        'sender_name',
        'message_data.sender_name',
        'data.sender_name',
        'data.sender.name'
    ]);

    const messageId = readFirstValue(body, [
        'id',
        'message_id',
        'message_data.id',
        'data.id',
        'data.message_data.id',
        'direct_message_events.0.id'
    ]);

    return {
        eventType:
            eventType !== undefined
                ? String(eventType)
                : undefined,

        messageText:
            typeof messageText === 'string'
                ? messageText.trim()
                : messageText,

        senderId:
            senderId !== undefined
                ? String(senderId)
                : undefined,

        senderUsername:
            senderUsername !== undefined
                ? String(senderUsername)
                : undefined,

        senderName:
            senderName !== undefined
                ? String(senderName)
                : undefined,

        messageId:
            messageId !== undefined
                ? String(messageId)
                : undefined
    };
}

/*
 * Request an application access token from Microsoft Entra ID.
 */
async function getD365AccessToken() {
    const tenantId = process.env.D365_TENANT_ID;

    const tokenUrl =
        'https://login.microsoftonline.com/' +
        encodeURIComponent(tenantId) +
        '/oauth2/v2.0/token';

    const tokenRequestBody = new URLSearchParams({
        client_id: process.env.D365_CLIENT_ID,
        client_secret: process.env.D365_CLIENT_SECRET,
        grant_type: 'client_credentials',
        scope: process.env.D365_TOKEN_SCOPE
    });

    console.log(
        'Requesting Microsoft Entra access token.'
    );

    const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: {
            'Content-Type':
                'application/x-www-form-urlencoded'
        },
        body: tokenRequestBody.toString()
    });

    const responseText = await response.text();

    let responseBody;

    try {
        responseBody = responseText
            ? JSON.parse(responseText)
            : {};
    } catch {
        responseBody = {
            rawResponse: responseText
        };
    }

    if (!response.ok) {
        throw new Error(
            'Microsoft Entra token request failed. ' +
            `Status: ${response.status}. ` +
            `Response: ${JSON.stringify(responseBody)}`
        );
    }

    if (!responseBody.access_token) {
        throw new Error(
            'Microsoft Entra did not return an access token.'
        );
    }

    console.log(
        'Microsoft Entra access token obtained successfully.'
    );

    return responseBody.access_token;
}

/*
 * Create the headers required by the Dynamics 365
 * Messaging API.
 */
function getD365RequestHeaders(accessToken) {
    return {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'channel-id': process.env.D365_CHANNEL_ID,
        'organization-id':
            process.env.D365_ORGANIZATION_ID
    };
}

/*
 * Send a request to the Dynamics 365 Messaging API.
 */
async function callD365MessagingApi(
    path,
    method,
    body
) {
    const accessToken =
        await getD365AccessToken();

    const baseUrl = removeTrailingSlash(
        process.env.D365_MESSAGING_BASE_URL
    );

    const requestUrl = `${baseUrl}${path}`;

    console.log(
        `Calling Dynamics Messaging API: ${method} ${path}`
    );

    const response = await fetch(requestUrl, {
        method,
        headers:
            getD365RequestHeaders(accessToken),
        body:
            body !== undefined
                ? JSON.stringify(body)
                : undefined
    });

    const responseText = await response.text();

    let responseBody;

    try {
        responseBody = responseText
            ? JSON.parse(responseText)
            : {};
    } catch {
        responseBody = {
            rawResponse: responseText
        };
    }

    if (!response.ok) {
        throw new Error(
            'Dynamics Messaging API request failed. ' +
            `Status: ${response.status}. ` +
            `Response: ${JSON.stringify(responseBody)}`
        );
    }

    return responseBody;
}

/*
 * Send another customer message to an existing
 * Dynamics 365 conversation.
 */
async function sendCustomerMessageToD365(
    conversationId,
    xMessage,
    customerId,
    preferredName
) {
    const activity = {
        type: 'message',

        id:
            xMessage.messageId ||
            crypto.randomUUID(),

        channelId:
            process.env.D365_CHANNEL_ID,

        from: {
            id: customerId,
            name: preferredName
        },

        conversation: {
            id: conversationId
        },

        text: xMessage.messageText
    };

    await callD365MessagingApi(
        '/api/v1.0/consumer/conversation/' +
            encodeURIComponent(conversationId),
        'POST',
        activity
    );

    console.log(
        'Customer message sent to existing Dynamics conversation:',
        conversationId
    );
}

/*
 * Create a new Dynamics 365 Omnichannel conversation or
 * resume an active conversation for the same X customer.
 */
async function createOrResumeD365Conversation(
    xMessage
) {
    let stableCustomerId =
        xMessage.senderId ||
        xMessage.senderUsername;

    /*
     * Generate a diagnostic fallback customer ID when X
     * does not provide a sender ID or username.
     */
    if (!stableCustomerId) {
        stableCustomerId =
            'x-user-' +
            crypto
                .createHash('sha256')
                .update(
                    String(
                        xMessage.messageId ||
                        xMessage.messageText
                    )
                )
                .digest('hex')
                .substring(0, 32);
    }

    const cleanUsername =
        xMessage.senderUsername
            ? xMessage.senderUsername.replace(
                /^@/,
                ''
            )
            : undefined;

    const preferredName =
        xMessage.senderName ||
        (
            cleanUsername
                ? `@${cleanUsername}`
                : `X User ${stableCustomerId}`
        );

    const conversationRequestId =
        xMessage.messageId ||
        crypto.randomUUID();

    const payload = {
        customercontext: {
            customerid: stableCustomerId,
            preferredname: preferredName
        },

        conversationcontext: {
            source: {
                isDisplayable: true,
                value: 'X Direct Message'
            },

            xUserId: {
                isDisplayable: false,
                value: stableCustomerId
            },

            xUsername: {
                isDisplayable: true,
                value: cleanUsername || ''
            }
        },

        conversationrequestid:
            conversationRequestId,

        startmessage: {
            message: xMessage.messageText,
            displayname: preferredName
        },

        skipdeflectionbot: true
    };

    console.log(
        'Creating or resuming Dynamics conversation for:',
        stableCustomerId
    );

    const result = await callD365MessagingApi(
        '/api/v1.0/consumer/conversation/create',
        'POST',
        payload
    );

    if (!result.conversationId) {
        throw new Error(
            'Dynamics did not return a conversationId.'
        );
    }

    console.log(
        'Dynamics conversation response:',
        JSON.stringify(
            {
                conversationId:
                    result.conversationId,
                isNew: result.isNew,
                messageId: result.messageId
            },
            null,
            2
        )
    );

    /*
     * For a new conversation, startmessage already contains
     * the first customer message.
     *
     * For an existing conversation, send the latest DM as a
     * standard conversation message.
     */
    if (result.isNew === false) {
        await sendCustomerMessageToD365(
            result.conversationId,
            xMessage,
            stableCustomerId,
            preferredName
        );
    }

    return result;
}

/*
 * Main health-check endpoint.
 */
app.get('/', (req, res) => {
    const environmentConfigured =
        validateEnvironmentVariables();

    res.status(200).json({
        service: 'Twitter-D365 Middleware',
        status: 'running',
        environmentConfigured
    });
});

/*
 * X webhook CRC verification.
 */
app.get('/webhook', (req, res) => {
    console.log('GET /webhook RECEIVED');

    const crcToken = req.query.crc_token;

    const consumerSecret =
        process.env.X_CONSUMER_SECRET;

    /*
     * Normal browser health check when there is no CRC token.
     */
    if (!crcToken) {
        return res
            .status(200)
            .send('Webhook is running');
    }

    if (!consumerSecret) {
        console.error(
            'X_CONSUMER_SECRET is not configured.'
        );

        return res.status(500).json({
            error:
                'Webhook secret is not configured'
        });
    }

    const hmac = crypto
        .createHmac(
            'sha256',
            consumerSecret
        )
        .update(crcToken)
        .digest('base64');

    console.log(
        'X CRC validation successful.'
    );

    return res.status(200).json({
        response_token: `sha256=${hmac}`
    });
});

/*
 * Receive X webhook events.
 */
app.post('/webhook', (req, res) => {
    console.log(
        '===================================='
    );

    console.log(
        'POST /webhook RECEIVED'
    );

    console.log(
        'X webhook body:',
        JSON.stringify(req.body, null, 2)
    );

    console.log(
        '===================================='
    );

    /*
     * Acknowledge the X webhook immediately.
     */
    res.status(200).send('OK');

    /*
     * Process the event after acknowledging X.
     */
    setImmediate(async () => {
        try {
            const xMessage =
                extractXDirectMessage(req.body);

            console.log(
                'Extracted X event:',
                JSON.stringify(
                    xMessage,
                    null,
                    2
                )
            );

            /*
             * Ignore events that are not received DMs.
             */
            if (
                xMessage.eventType &&
                xMessage.eventType !==
                    'dm.received'
            ) {
                console.log(
                    'Ignoring unsupported X event type:',
                    xMessage.eventType
                );

                return;
            }

            if (!xMessage.messageText) {
                console.log(
                    'The X event did not contain a text Direct Message.'
                );

                return;
            }

            const result =
                await createOrResumeD365Conversation(
                    xMessage
                );

            console.log(
                'X DM processed successfully.'
            );

            console.log(
                'Dynamics conversation ID:',
                result.conversationId
            );
        } catch (error) {
            console.error(
                'Failed to send X DM to Dynamics 365:',
                error.message
            );
        }
    });
});

/*
 * Receive agent activities and conversation events from
 * Dynamics 365.
 *
 * Your configured custom channel endpoint is:
 * https://twitter-d365-middleware.onrender.com/webhook
 *
 * Dynamics appends:
 * /v3/conversations/{conversationId}/activities
 */
app.post(
    '/webhook/v3/conversations/:conversationId/activities',
    (req, res) => {
        console.log(
            '===================================='
        );

        console.log(
            'DYNAMICS ACTIVITY RECEIVED'
        );

        console.log(
            'Conversation ID:',
            req.params.conversationId
        );

        console.log(
            'Activity body:',
            JSON.stringify(req.body, null, 2)
        );

        console.log(
            '===================================='
        );

        /*
         * Acknowledge the Dynamics activity.
         *
         * The next phase will send agent replies back to X.
         */
        res.status(200).json({
            success: true
        });
    }
);

/*
 * Compatibility and endpoint health checks.
 */
app.get('/api/messages', (req, res) => {
    res.status(200).json({
        endpoint: '/api/messages',
        status: 'online'
    });
});

app.post('/api/messages', (req, res) => {
    console.log(
        'POST /api/messages RECEIVED'
    );

    console.log(
        JSON.stringify(req.body, null, 2)
    );

    res.status(200).json({
        success: true
    });
});

/*
 * Express error handler.
 */
app.use((error, req, res, next) => {
    console.error(
        'Unhandled middleware error:',
        error
    );

    if (!res.headersSent) {
        res.status(500).json({
            error: 'Internal Server Error'
        });
    }
});

/*
 * Start the Render web service.
 */
const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {
    console.log(
        `Server running on port ${PORT}`
    );

    validateEnvironmentVariables();
});
