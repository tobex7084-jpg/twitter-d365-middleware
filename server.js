'use strict';

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
 * Cached Microsoft Entra access token.
 */
let cachedAccessToken = null;
let cachedAccessTokenExpiresAt = 0;

/*
 * Temporary in-memory duplicate event protection.
 *
 * This resets whenever Render restarts.
 */
const processedXEvents = new Map();
const EVENT_CACHE_DURATION_MS = 60 * 60 * 1000;

/*
 * Validate the required Render environment variables.
 *
 * This function logs only missing variable names.
 * It never prints secret values.
 */
function validateEnvironmentVariables() {
    const missingVariables =
        REQUIRED_ENVIRONMENT_VARIABLES.filter(
            variableName =>
                !process.env[variableName] ||
                String(process.env[variableName]).trim() === ''
        );

    if (missingVariables.length > 0) {
        console.error(
            'Missing environment variables:',
            missingVariables.join(', ')
        );

        return false;
    }

    console.log(
        'All required environment variables are configured.'
    );

    return true;
}

/*
 * Remove trailing slash characters from a URL.
 */
function removeTrailingSlash(value) {
    return String(value || '')
        .trim()
        .replace(/\/+$/, '');
}

/*
 * Get an X profile from the users object supplied
 * in the webhook payload.
 */
function getXUserProfile(body, userId) {
    if (
        !body ||
        !body.data ||
        !body.data.payload ||
        !body.data.payload.users ||
        !userId
    ) {
        return null;
    }

    const userContainer =
        body.data.payload.users[String(userId)];

    if (!userContainer) {
        return null;
    }

    return userContainer.data || userContainer;
}

/*
 * Create a safe profile object for logging.
 *
 * Only X profile values already supplied in the webhook
 * event are included.
 */
function buildExtractedXProfile(
    userId,
    userProfile
) {
    if (!userId && !userProfile) {
        return null;
    }

    return {
        id:
            userId
                ? String(userId)
                : null,

        username:
            userProfile &&
            userProfile.username
                ? String(userProfile.username)
                : null,

        name:
            userProfile &&
            userProfile.name
                ? String(userProfile.name)
                : null,

        description:
            userProfile &&
            userProfile.description
                ? String(userProfile.description)
                : null,

        profileImageUrl:
            userProfile &&
            userProfile.profile_image_url
                ? String(userProfile.profile_image_url)
                : null,

        profileBannerUrl:
            userProfile &&
            userProfile.profile_banner_url
                ? String(userProfile.profile_banner_url)
                : null,

        accountCreatedAt:
            userProfile &&
            userProfile.created_at
                ? String(userProfile.created_at)
                : null,

        protected:
            userProfile &&
            typeof userProfile.protected === 'boolean'
                ? userProfile.protected
                : null,

        verified:
            userProfile &&
            typeof userProfile.verified === 'boolean'
                ? userProfile.verified
                : null,

        verifiedType:
            userProfile &&
            userProfile.verified_type
                ? String(userProfile.verified_type)
                : null,

        isIdentityVerified:
            userProfile &&
            typeof userProfile.is_identity_verified === 'boolean'
                ? userProfile.is_identity_verified
                : null,

        publicMetrics:
            userProfile &&
            userProfile.public_metrics
                ? userProfile.public_metrics
                : null
    };
}

/*
 * Extract the incoming X Direct Message event.
 *
 * This matches the webhook payload already confirmed
 * in the Render logs.
 */
function extractXDirectMessage(body) {
    const data =
        body && body.data
            ? body.data
            : {};

    const payload =
        data && data.payload
            ? data.payload
            : {};

    const directMessageEvents =
        Array.isArray(
            payload.direct_message_events
        )
            ? payload.direct_message_events
            : [];

    const directMessageEvent =
        directMessageEvents.length > 0
            ? directMessageEvents[0]
            : null;

    const messageCreate =
        directMessageEvent &&
        directMessageEvent.message_create
            ? directMessageEvent.message_create
            : null;

    const messageData =
        messageCreate &&
        messageCreate.message_data
            ? messageCreate.message_data
            : null;

    const target =
        messageCreate &&
        messageCreate.target
            ? messageCreate.target
            : null;

    const senderId =
        messageCreate &&
        messageCreate.sender_id
            ? String(messageCreate.sender_id)
            : null;

    const recipientId =
        target &&
        target.recipient_id
            ? String(target.recipient_id)
            : null;

    const senderProfile =
        getXUserProfile(
            body,
            senderId
        );

    const recipientProfile =
        getXUserProfile(
            body,
            recipientId
        );

    return {
        eventUuid:
            data.event_uuid
                ? String(data.event_uuid)
                : null,

        eventType:
            data.event_type
                ? String(data.event_type)
                : null,

        subscribedUserId:
            data.filter &&
            data.filter.user_id
                ? String(data.filter.user_id)
                : null,

        message: {
            id:
                directMessageEvent &&
                directMessageEvent.id
                    ? String(
                        directMessageEvent.id
                    )
                    : null,

            type:
                directMessageEvent &&
                directMessageEvent.type
                    ? String(
                        directMessageEvent.type
                    )
                    : null,

            createdTimestamp:
                directMessageEvent &&
                directMessageEvent.created_timestamp
                    ? String(
                        directMessageEvent.created_timestamp
                    )
                    : null,

            text:
                messageData &&
                typeof messageData.text === 'string'
                    ? messageData.text.trim()
                    : null,

            entities:
                messageData &&
                messageData.entities
                    ? messageData.entities
                    : null
        },

        sender:
            buildExtractedXProfile(
                senderId,
                senderProfile
            ),

        recipient:
            buildExtractedXProfile(
                recipientId,
                recipientProfile
            )
    };
}

/*
 * Remove expired duplicate-event records.
 */
function cleanProcessedEventCache() {
    const now = Date.now();

    for (
        const [eventId, processedAt]
        of processedXEvents.entries()
    ) {
        if (
            now - processedAt >
            EVENT_CACHE_DURATION_MS
        ) {
            processedXEvents.delete(eventId);
        }
    }
}

/*
 * Determine whether an X event has already been processed.
 */
function isDuplicateXEvent(eventId) {
    if (!eventId) {
        return false;
    }

    cleanProcessedEventCache();

    if (processedXEvents.has(eventId)) {
        return true;
    }

    processedXEvents.set(
        eventId,
        Date.now()
    );

    return false;
}

/*
 * Obtain a Microsoft Entra access token through
 * the OAuth 2.0 client-credentials flow.
 */
async function getD365AccessToken() {
    const currentTime = Date.now();

    if (
        cachedAccessToken &&
        cachedAccessTokenExpiresAt >
            currentTime + 60000
    ) {
        console.log(
            'Using cached Microsoft Entra access token.'
        );

        return cachedAccessToken;
    }

    const tenantId =
        process.env.D365_TENANT_ID;

    const tokenUrl =
        'https://login.microsoftonline.com/' +
        encodeURIComponent(tenantId) +
        '/oauth2/v2.0/token';

    const tokenRequestBody =
        new URLSearchParams({
            client_id:
                process.env.D365_CLIENT_ID,

            client_secret:
                process.env.D365_CLIENT_SECRET,

            grant_type:
                'client_credentials',

            scope:
                process.env.D365_TOKEN_SCOPE
        });

    console.log(
        'Requesting Microsoft Entra access token.'
    );

    let response;

    try {
        response = await fetch(
            tokenUrl,
            {
                method: 'POST',

                headers: {
                    'Content-Type':
                        'application/x-www-form-urlencoded'
                },

                body:
                    tokenRequestBody.toString()
            }
        );
    } catch (error) {
        console.error(
            'Microsoft Entra network request failed.'
        );

        console.error(
            'Error message:',
            error.message
        );

        console.error(
            'Error cause:',
            error.cause ||
                'No underlying error cause supplied'
        );

        console.error(
            'Error stack:',
            error.stack
        );

        throw error;
    }

    const responseText =
        await response.text();

    let responseBody;

    try {
        responseBody =
            responseText
                ? JSON.parse(responseText)
                : {};
    } catch {
        responseBody = {
            rawResponse:
                responseText
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

    const expiresInSeconds =
        Number(responseBody.expires_in) ||
        3600;

    cachedAccessToken =
        responseBody.access_token;

    cachedAccessTokenExpiresAt =
        Date.now() +
        expiresInSeconds * 1000;

    console.log(
        'Microsoft Entra access token obtained successfully.'
    );

    return cachedAccessToken;
}

/*
 * Create the Dynamics Messaging API headers.
 */
function getD365RequestHeaders(accessToken) {
    return {
        Authorization:
            `Bearer ${accessToken}`,

        'Content-Type':
            'application/json',

        Accept:
            'application/json',

        'channel-id':
            process.env.D365_CHANNEL_ID,

        'organization-id':
            process.env.D365_ORGANIZATION_ID
    };
}

/*
 * Read the API response as JSON where possible.
 */
async function readResponseBody(response) {
    const responseText =
        await response.text();

    if (!responseText) {
        return {};
    }

    try {
        return JSON.parse(responseText);
    } catch {
        return {
            rawResponse:
                responseText
        };
    }
}

/*
 * Call the Dynamics 365 Messaging API.
 */
async function callD365MessagingApi(
    path,
    method,
    requestBody
) {
    const accessToken =
        await getD365AccessToken();

    const baseUrl =
        removeTrailingSlash(
            process.env.D365_MESSAGING_BASE_URL
        );

    const requestUrl =
        `${baseUrl}${path}`;

    console.log(
        'Dynamics request URL:',
        requestUrl
    );

    console.log(
        'Dynamics request method:',
        method
    );

    let response;

    try {
        response = await fetch(
            requestUrl,
            {
                method,

                headers:
                    getD365RequestHeaders(
                        accessToken
                    ),

                body:
                    requestBody !== undefined
                        ? JSON.stringify(
                            requestBody
                        )
                        : undefined
            }
        );
    } catch (error) {
        console.error(
            'Network connection to Dynamics failed.'
        );

        console.error(
            'Request URL:',
            requestUrl
        );

        console.error(
            'Error message:',
            error.message
        );

        console.error(
            'Error cause:',
            error.cause ||
                'No underlying error cause supplied'
        );

        console.error(
            'Error stack:',
            error.stack
        );

        throw error;
    }

    const responseBody =
        await readResponseBody(
            response
        );

    console.log(
        'Dynamics HTTP status:',
        response.status
    );

    console.log(
        'Dynamics HTTP status text:',
        response.statusText
    );

    if (!response.ok) {
        throw new Error(
            'Dynamics Messaging API rejected the request. ' +
            `Status: ${response.status}. ` +
            `Response: ${JSON.stringify(responseBody)}`
        );
    }

    return responseBody;
}

/*
 * Create a stable customer ID from the X sender.
 */
function getStableCustomerId(xEvent) {
    if (
        xEvent.sender &&
        xEvent.sender.id
    ) {
        return (
            `x-${xEvent.sender.id}`
        );
    }

    if (
        xEvent.sender &&
        xEvent.sender.username
    ) {
        return (
            `x-${xEvent.sender.username}`
                .toLowerCase()
                .replace(/^x-@/, 'x-')
        );
    }

    return (
        'x-user-' +
        crypto
            .createHash('sha256')
            .update(
                String(
                    xEvent.message.id ||
                    xEvent.eventUuid ||
                    xEvent.message.text
                )
            )
            .digest('hex')
            .substring(0, 32)
    );
}

/*
 * Create the display name shown to the Dynamics agent.
 */
function getPreferredName(xEvent) {
    if (
        xEvent.sender &&
        xEvent.sender.name
    ) {
        return xEvent.sender.name;
    }

    if (
        xEvent.sender &&
        xEvent.sender.username
    ) {
        return (
            '@' +
            xEvent.sender.username.replace(
                /^@/,
                ''
            )
        );
    }

    if (
        xEvent.sender &&
        xEvent.sender.id
    ) {
        return (
            `X User ${xEvent.sender.id}`
        );
    }

    return 'X Customer';
}

/*
 * Send a customer message into an existing
 * Dynamics conversation.
 */
async function sendMessageToExistingConversation(
    conversationId,
    xEvent,
    customerId,
    preferredName
) {
    const messageActivity = {
        type:
            'message',

        id:
            xEvent.message.id ||
            crypto.randomUUID(),

        channelId:
            process.env.D365_CHANNEL_ID,

        from: {
            id:
                customerId,

            name:
                preferredName
        },

        conversation: {
            id:
                conversationId
        },

        text:
            xEvent.message.text
    };

    await callD365MessagingApi(
        '/api/v1.0/consumer/conversation/' +
        encodeURIComponent(
            conversationId
        ),
        'POST',
        messageActivity
    );

    console.log(
        'Customer message sent to existing Dynamics conversation:',
        conversationId
    );
}

/*
 * Create or resume a Dynamics 365 conversation.
 */
async function createOrResumeD365Conversation(
    xEvent
) {
    const customerId =
        getStableCustomerId(
            xEvent
        );

    const preferredName =
        getPreferredName(
            xEvent
        );

    const conversationRequestId =
        xEvent.message.id ||
        xEvent.eventUuid ||
        crypto.randomUUID();

    const payload = {
        customercontext: {
            customerid:
                customerId,

            firstname:
                xEvent.sender &&
                xEvent.sender.name
                    ? xEvent.sender.name
                    : undefined,

            preferredname:
                preferredName
        },

        conversationcontext: {
            source: {
                isDisplayable:
                    true,

                value:
                    'X Direct Message'
            },

            xSenderId: {
                isDisplayable:
                    false,

                value:
                    xEvent.sender &&
                    xEvent.sender.id
                        ? xEvent.sender.id
                        : ''
            },

            xSenderUsername: {
                isDisplayable:
                    true,

                value:
                    xEvent.sender &&
                    xEvent.sender.username
                        ? xEvent.sender.username
                        : ''
            },

            xSenderName: {
                isDisplayable:
                    true,

                value:
                    xEvent.sender &&
                    xEvent.sender.name
                        ? xEvent.sender.name
                        : ''
            },

            xRecipientId: {
                isDisplayable:
                    false,

                value:
                    xEvent.recipient &&
                    xEvent.recipient.id
                        ? xEvent.recipient.id
                        : ''
            },

            xRecipientUsername: {
                isDisplayable:
                    false,

                value:
                    xEvent.recipient &&
                    xEvent.recipient.username
                        ? xEvent.recipient.username
                        : ''
            },

            xMessageId: {
                isDisplayable:
                    false,

                value:
                    xEvent.message.id ||
                    ''
            }
        },

        conversationrequestid:
            conversationRequestId,

        startmessage: {
            message:
                xEvent.message.text,

            displayname:
                preferredName
        },

        /*
         * Route directly to live-agent handling.
         */
        skipdeflectionbot:
            true
    };

    console.log(
        'Dynamics customer ID:',
        customerId
    );

    console.log(
        'Dynamics customer display name:',
        preferredName
    );

    console.log(
        'Creating or resuming Dynamics conversation.'
    );

    const result =
        await callD365MessagingApi(
            '/api/v1.0/consumer/conversation/create',
            'POST',
            payload
        );

    if (!result.conversationId) {
        throw new Error(
            'Dynamics response did not contain a conversationId.'
        );
    }

    console.log(
        'Dynamics conversation response:',
        JSON.stringify(
            {
                conversationId:
                    result.conversationId,

                isNew:
                    result.isNew,

                messageId:
                    result.messageId
            },
            null,
            2
        )
    );

    /*
     * The first message is delivered through startmessage
     * when the conversation is new.
     */
    if (result.isNew === false) {
        await sendMessageToExistingConversation(
            result.conversationId,
            xEvent,
            customerId,
            preferredName
        );
    }

    return result;
}

/*
 * Root health-check endpoint.
 */
app.get(
    '/',
    (req, res) => {
        res.status(200).json({
            service:
                'Twitter-D365 Middleware',

            status:
                'running',

            mode:
                'X to Dynamics',

            environmentConfigured:
                validateEnvironmentVariables()
        });
    }
);

/*
 * X webhook CRC verification.
 */
app.get(
    '/webhook',
    (req, res) => {
        console.log(
            'GET /webhook RECEIVED'
        );

        const crcToken =
            req.query.crc_token;

        const consumerSecret =
            process.env.X_CONSUMER_SECRET;

        /*
         * Normal browser health check.
         */
        if (!crcToken) {
            return res
                .status(200)
                .send(
                    'X-to-Dynamics webhook is running'
                );
        }

        if (!consumerSecret) {
            console.error(
                'X_CONSUMER_SECRET is not configured.'
            );

            return res
                .status(500)
                .json({
                    error:
                        'X_CONSUMER_SECRET is not configured'
                });
        }

        const hmac =
            crypto
                .createHmac(
                    'sha256',
                    consumerSecret
                )
                .update(crcToken)
                .digest('base64');

        console.log(
            'X CRC validation successful.'
        );

        return res
            .status(200)
            .json({
                response_token:
                    `sha256=${hmac}`
            });
    }
);

/*
 * Receive X Direct Message webhook events.
 */
app.post(
    '/webhook',
    (req, res) => {
        console.log(
            '===================================='
        );

        console.log(
            'POST /webhook RECEIVED'
        );

        /*
         * Acknowledge the X webhook immediately.
         */
        res.status(200).send('OK');

        setImmediate(
            async () => {
                try {
                    const xEvent =
                        extractXDirectMessage(
                            req.body
                        );

                    console.log(
                        'FULL EXTRACTED X DM DETAILS'
                    );

                    console.log(
                        JSON.stringify(
                            xEvent,
                            null,
                            2
                        )
                    );

                    if (
                        xEvent.eventType !==
                        'dm.received'
                    ) {
                        console.log(
                            'Event ignored because it is not dm.received:',
                            xEvent.eventType
                        );

                        return;
                    }

                    if (
                        !xEvent.message ||
                        !xEvent.message.text
                    ) {
                        console.log(
                            'DM event received without message text.'
                        );

                        return;
                    }

                    const duplicateKey =
                        xEvent.message.id ||
                        xEvent.eventUuid;

                    if (
                        isDuplicateXEvent(
                            duplicateKey
                        )
                    ) {
                        console.log(
                            'Duplicate X event ignored:',
                            duplicateKey
                        );

                        return;
                    }

                    console.log(
                        'SENDER DETAILS'
                    );

                    console.log(
                        JSON.stringify(
                            xEvent.sender,
                            null,
                            2
                        )
                    );

                    console.log(
                        'RECIPIENT DETAILS'
                    );

                    console.log(
                        JSON.stringify(
                            xEvent.recipient,
                            null,
                            2
                        )
                    );

                    console.log(
                        'MESSAGE DETAILS'
                    );

                    console.log(
                        JSON.stringify(
                            xEvent.message,
                            null,
                            2
                        )
                    );

                    const result =
                        await createOrResumeD365Conversation(
                            xEvent
                        );

                    console.log(
                        'X DM sent to Dynamics successfully.'
                    );

                    console.log(
                        'Dynamics conversation ID:',
                        result.conversationId
                    );
                } catch (error) {
                    console.error(
                        'Failed to send X DM to Dynamics 365.'
                    );

                    console.error(
                        'Error message:',
                        error.message
                    );

                    console.error(
                        'Error cause:',
                        error.cause ||
                            'No underlying error cause supplied'
                    );

                    console.error(
                        'Error stack:',
                        error.stack
                    );
                } finally {
                    console.log(
                        '===================================='
                    );
                }
            }
        );
    }
);

/*
 * Receive agent messages and conversation events from Dynamics.
 *
 * Dynamics appends this route to the configured webhook URL.
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
            'Dynamics conversation ID:',
            req.params.conversationId
        );

        console.log(
            'Dynamics activity:',
            JSON.stringify(
                req.body,
                null,
                2
            )
        );

        console.log(
            '===================================='
        );

        /*
         * Agent-to-X message delivery is not implemented yet.
         */
        res.status(200).json({
            success:
                true
        });
    }
);

/*
 * Diagnostic GET endpoint.
 */
app.get(
    '/api/messages',
    (req, res) => {
        res.status(200).json({
            endpoint:
                '/api/messages',

            status:
                'online',

            purpose:
                'Diagnostic endpoint only'
        });
    }
);

/*
 * Diagnostic POST endpoint.
 */
app.post(
    '/api/messages',
    (req, res) => {
        console.log(
            'POST /api/messages RECEIVED'
        );

        console.log(
            JSON.stringify(
                req.body,
                null,
                2
            )
        );

        res.status(200).json({
            success:
                true
        });
    }
);

/*
 * Route-not-found handler.
 */
app.use(
    (req, res) => {
        res.status(404).json({
            error:
                'Route not found',

            method:
                req.method,

            path:
                req.path
        });
    }
);

/*
 * Express error handler.
 */
app.use(
    (error, req, res, next) => {
        console.error(
            'Unhandled Express error:',
            error
        );

        if (!res.headersSent) {
            res.status(500).json({
                error:
                    'Internal Server Error'
            });
        }
    }
);

/*
 * Start Render service.
 */
const PORT =
    process.env.PORT || 10000;

app.listen(
    PORT,
    () => {
        console.log(
            `Server running on port ${PORT}`
        );

        validateEnvironmentVariables();
    }
);
