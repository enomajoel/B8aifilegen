const express = require('express');
const cors = require('cors');
const filegenRoutes = require('./b8ai-filegen-routes');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(filegenRoutes);

app.get('/healthz', (req, res) => res.send('ok'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Filegen service running on port ${PORT}`));
