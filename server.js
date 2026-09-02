const express = require('express');
const app = express();

app.use(express.json());

app.get('/', (req, res) => {
    res.send('Twitter-D365 Middleware is running');
});

app.post('/webhook', async (req, res) => {
    console.log('Received webhook:', JSON.stringify(req.body, null, 2));

    // Later we will send this message to Dynamics 365 Omnichannel

    res.status(200).send('OK');
});

const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
