'use strict';

const express = require('express');
const crypto = require('crypto');

const app = express();

/*
 * Node.js 18 or later is required because this application
 * uses the built-in fetch function.
 */
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
 * Validate environment variables without displaying
 * any secret values.
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
 * Remove trailing slashes from a URL.
 */
function removeTrailingSlash(value) {
    return String(value || '').replace(/\/+$/, '');
}

/*
 * Retrieve the first available value from several
 * possible object paths.
 */
function readFirstValue(object, paths) {
    for (const path of paths) {
        const value = path
            .split('.')
            .reduce(
                (currentValue, propertyName) => {
                    if (
                        currentValue !== undefined &&
                        currentValue !== null &&
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
 * Extract an X Direct Message webhook event.
 *
 * Multiple possible paths are checked because X event
 * payloads can differ across API versions.
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
        'text'
    ]);

    const senderId = readFirstValue(body, [
        'sender_id',
        'message_data.sender_id',
        'data.sender_id',
        'data.sender.id',
        'data.author_id',
        'sender.id'
    ]);

    const senderUsername = readFirstValue(body, [
        'sender_username',
        'message_data.sender_username',
        'data.sender_username',
        'data.sender.username',
        'sender.username'
    ]);

    const senderName = readFirstValue(body, [
        'sender_name',
        'message_data.sender_name',
        'data.sender_name',
        'data.sender.name',
        'sender.name'
    ]);

    const messageId = readFirstValue(body, [
        'message_id',
        'message_data.id',
        'data.message_data.id',
        'data.id',
        'id'
    ]);

    const xConversationId = readFirstValue(body, [
        'dm_conversation_id',
        'conversation_id',
        'message_data.dm_conversation_id',
        'data.dm_conversation_id',
        'data.conversation_id',
        'data.conversation.id'
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
                : undefined,

        xConversationId:
            xConversationId !== undefined
                ? String(xConversationId)
                : undefined
    };
}

/*
 * Request a Microsoft Entra access token using the
 * OAuth 2.0 client-credentials flow.
 */
async function getD365AccessToken() {
    const tokenUrl =
        `https://login.microsoftonline.com/` +
        `${encodeURIComponent(process.env.D365_TENANT_ID)}` +
        `/oauth2/v2.0/token`;

    const tokenRequestBody = new URLSearchParams({
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
            `Microsoft Entra token request failed. ` +
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
 * Build the HTTP headers required by the
 * Dynamics 365 Messaging API.
 */
function getD365RequestHeaders(accessToken) {
    return {
        Authorization:
            `Bearer ${accessToken}`,

        'Content-Type':
            'application/json',

        'channel-id':
            process
