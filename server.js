const express = require('express');
const app = express();

app.use(express.json());

app.get('/', (req, res) => {
    res.send('Twitter-D365 Middleware is running');
});

app.get('/webhook', (req, res) => {
    res.status(200).send('Webhook is running');
});

app.post('/webhook', async (req, res) => {
    console.log('Received webhook:', JSON.stringify(req.body, null, 2));

    res.status(200).send('OK');
});

const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
