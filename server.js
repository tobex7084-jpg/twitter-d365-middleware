const express = require('express');
const crypto = require('crypto');

const app = express();

// Parse incoming JSON webhook events
app.use(express.json());

/*
 * Health-check endpoint
 * Open: https://twitter-d365-middleware.onrender.com
 */
app.get('/', (req, res) => {
    res.status(200).send('Twitter-D365 Middleware is running');
});

/*
 * X webhook CRC verification endpoint
 * X sends a GET request containing crc_token.
 */
app.get('/webhook', (req, res) => {
    const crcToken = req.query.crc_token;
    const consumerSecret = process.env.X_CONSUMER_SECRET;

    if (!crcToken) {
        return res.status(200).send('Webhook is running');
    }

    if (!consumerSecret) {
        console.error('X_CONSUMER_SECRET is not configured in Render.');
        return res.status(500).json({
            error: 'Webhook secret is not configured'
        });
    }

    const hmac = crypto
        .createHmac('sha256', consumerSecret)
        .update(crcToken)
        .digest('base64');

    const responseToken = `sha256=${hmac}`;

    console.log('X CRC verification request received.');

    return res.status(200).json({
        response_token: responseToken
    });
});

/*
 * Receives webhook events from X
 */
app.post('/webhook', async (req, res) => {
    try {
        console.log(
            'Received X webhook event:',
            JSON.stringify(req.body, null, 2)
        );

        // Acknowledge receipt immediately.
        res.status(200).send('OK');

        // Later, add the logic for sending incoming X messages
        // to the Dynamics 365 Omnichannel custom channel here.
    } catch (error) {
        console.error('Webhook processing error:', error);

        if (!res.headersSent) {
            res.status(500).send('Internal Server Error');
        }
    }
});

/*
 * Start the Render web service.
 */
const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
