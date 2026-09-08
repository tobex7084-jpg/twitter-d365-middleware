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
 * Verify required environment variables without exposing values.
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
 * Retrieve the first non-empty value from a list of object paths.
 *
 * Numeric path components, such as direct_message_events.0,
 * also work with JavaScript arrays.
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
 * Extract the required fields from an incoming X Direct Message.
 *
 * The data.payload paths match the payload structure observed
 * in the Render logs.
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
        'direct_message_events.0.message_create.message_data.text',
        'payload.direct_message_events.0.message_create.message_data.text',
        'data.payload.direct_message_events.0.message_create.message_data.text'
    ]);

    const senderId = readFirstValue(body, [
        'sender_id',
        'message_data.sender_id',
        'data.sender_id',
        'data.sender.id',
        'data.author_id',
        'direct_message_events.0.message_create.sender_id',
        'payload.direct_message_events.0.message_create.sender_id',
        'data.payload.direct_message_events.0.message_create.sender_id'
    ]);

    const senderUsername = readFirstValue(body, [
        'sender_username',
        'message_data.sender_username',
        'data.sender_username',
        'data.sender.username',
        'payload.users.0.screen_name',
        'data.payload.users.0.screen_name'
    ]);

    const senderName = readFirstValue(body, [
        'sender_name',
        'message_data.sender_name',
        'data.sender_name',
        'data.sender.name',
        'payload.users.0.name',
        'data.payload.users.0.name'
    ]);

    const messageId = readFirstValue(body, [
        'id',
        'message_id',
        'message_data.id',
        'data.id',
        'data.message_data.id',
        'direct_message_events.0.id',
        'payload.direct_message_events.0.id',
        'data.payload.direct_message_events.0.id'
    ]);

    const xConversationId = readFirstValue(body, [
        'dm_conversation_id',
        'conversation_id',
        'message_data.dm_conversation_id',
        'data.dm_conversation_id',
        'data.conversation_id',
        'payload.direct_message_events.0.dm_conversation_id',
        'data.payload.direct_message_events.0.dm_conversation_id'
    ]);

    return {
        eventType:
           
