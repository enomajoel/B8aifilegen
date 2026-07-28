const express = require('express');
const cors = require('cors');
const filegenRoutes = require('./b8ai-filegen-routes');
const hostRoutes = require('./b8ai-host-route');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(filegenRoutes);
app.use(hostRoutes);

app.get('/', (req, res) => res.send('B8AI Filegen Service is running'));
app.get('/healthz', (req, res) => res.send('ok'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Filegen service running on port ${PORT}`));
