const express = require('express');
const crypto = require('crypto');

const app = express();

app.use(express.json());

/*
 * Health Check
 */
app.get('/', (req, res) => {
    console.log('GET / called');
    res.status(200).send('Twitter-D365 Middleware is running');
});

/*
 * X CRC Verification
 */
app.get('/webhook', (req, res) => {

    console.log('====================================');
    console.log('GET /webhook RECEIVED');
    console.log('Query:', req.query);
    console.log('====================================');

    const crcToken = req.query.crc_token;
    const consumerSecret = process.env.X_CONSUMER_SECRET;

    if (!crcToken) {
        return res.status(200).send('Webhook is running');
    }

    if (!consumerSecret) {
        console.error('X_CONSUMER_SECRET not configured');

        return res.status(500).json({
            error: 'Webhook secret not configured'
        });
    }

    const hmac = crypto
        .createHmac('sha256', consumerSecret)
        .update(crcToken)
        .digest('base64');

    const responseToken = `sha256=${hmac}`;

    console.log('CRC Validation Successful');

    return res.status(200).json({
        response_token: responseToken
    });
});

/*
 * X Event Receiver
 */
app.post('/webhook', async (req, res) => {

    console.log('====================================');
    console.log('POST /webhook RECEIVED');
    console.log('HEADERS');
    console.log(JSON.stringify(req.headers, null, 2));

    console.log('BODY');
    console.log(JSON.stringify(req.body, null, 2));
    console.log('====================================');

    try {

        /*
         * Detect DM Events
         */
        if (
            req.body &&
            req.body.data
        ) {
            console.log('Potential X event detected');
        }

        res.status(200).send('OK');

    } catch (error) {

        console.error('Webhook Processing Error');
        console.error(error);

        if (!res.headersSent) {
            res.status(500).send('Internal Server Error');
        }
    }
});

/*
 * Endpoint used by Dynamics 365
 */
app.post('/api/messages', (req, res) => {

    console.log('====================================');
    console.log('POST /api/messages RECEIVED');
    console.log(JSON.stringify(req.body, null, 2));
    console.log('====================================');

    res.status(200).json({
        success: true
    });
});

/*
 * Test Endpoint
 */
app.get('/api/messages', (req, res) => {
    res.status(200).send('Dynamics Endpoint Online');
});

const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
