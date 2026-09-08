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
 * Validate required environment variables without exposing
 * secret values in Render logs.
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
 * Return the first non-empty value found at one of the
 * supplied object paths.
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
 * Look up an X user record inside data.payload.users.
 */
function getXUserFromPayload(body, senderId) {
    if (!senderId) {
        return undefined;
    }

    return (
        body &&
        body.data &&
        body.data.payload &&
        body.data.payload.users &&
        body.data.payload.users[senderId] &&
        body.data.payload.users[senderId].data
    );
}

/*
 * Extract the Direct Message information from an X event.
 *
 * These paths include the exact X payload structure that
 * appeared in your Render logs.
 */
function extractXDirectMessage(body) {
    const eventType = readFirstValue(body, [
        'data.event_type',
        'event_type',
        'type'
    ]);

    const messageText = readFirstValue(body, [
        'data.payload.direct_message_events.0.message_create.message_data.text',
        'payload.direct_message_events.0.message_create.message_data.text',
        'direct_message_events.0.message_create.message_data.text',
        'data.message_data.text',
        'message_data.text',
        'data.text',
        'text'
    ]);

    const senderId = readFirstValue(body, [
        'data.payload.direct_message_events.0.message_create.sender_id',
        'payload.direct_message_events.0.message_create.sender_id',
        'direct_message_events.0.message_create.sender_id',
        'data.sender_id',
        'sender_id'
    ]);

    const recipientId = readFirstValue(body, [
        'data.payload.direct_message_events.0.message_create.target.recipient_id',
        'payload.direct_message_events.0.message_create.target.recipient_id',
        'direct_message_events.0.message_create.target.recipient_id'
    ]);

    const messageId = readFirstValue(body, [
        'data.payload.direct_message_events.0.id',
        'payload.direct_message_events.0.id',
        'direct_message_events.0.id',
        'data.message_data.id',
        'message_data.id',
        'data.id',
        'message_id',
        'id'
    ]);

    const senderIdString =
        senderId !== undefined
            ? String(senderId)
            : undefined;

    const xUser =
        getXUserFromPayload(
            body,
            senderIdString
        );

    const senderUsername =
        xUser && xUser.username
            ? String(xUser.username)
            : undefined;

    const senderName =
        xUser && xUser.name
            ? String(xUser.name)
            : undefined;

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
            senderIdString,

        senderUsername,

        senderName,

        recipientId:
            recipientId !== undefined
                ? String(recipientId)
                : undefined,

        messageId:
            messageId !== undefined
                ? String(messageId)
                : undefined
