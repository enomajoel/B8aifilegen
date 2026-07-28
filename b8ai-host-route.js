// ============================================================
// B8AI Director — /host route (B8Host integration)
//
// IMPORTANT SECURITY NOTE — read before deploying:
// This route uses your Supabase SERVICE ROLE key, which has full
// access to your 'sites' storage bucket, bypassing Row Level
// Security. That key must ONLY ever live here, as a Render
// environment variable — never in b8ai_v12.html, never in any
// file a browser downloads. That's the whole point of routing
// through a backend: the public chat app calls this endpoint,
// and only this server-side code ever touches the real key.
//
// Install on your Render backend:
//   npm install express-fileupload @supabase/supabase-js
//
// Render environment variables to set:
//   SUPABASE_URL              = https://<your-project>.supabase.co
//   SUPABASE_SERVICE_ROLE_KEY = <service_role key, NOT the anon key>
//   B8HOST_ACCOUNT_NAME       = b8ai   (the shared namespace all
//                                       chat-uploaded sites go under)
// ============================================================

const express = require('express');
const fileUpload = require('express-fileupload');
const { createClient } = require('@supabase/supabase-js');

const router = express.Router();

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

const B8HOST_ACCOUNT = process.env.B8HOST_ACCOUNT_NAME || 'b8ai';

router.use(fileUpload());

function slugifySiteName(name) {
  const slug = (name || '')
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, '')
    .trim()
    .replace(/\s+/g, '-')
    .slice(0, 40);
  return slug || `site-${Date.now()}`;
}

router.post('/v1/host', async (req, res) => {
  try {
    if (!req.files || !req.files.files) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    // express-fileupload gives a single object if there's only one file,
    // or an array if there are several — normalize to an array either way.
    let files = req.files.files;
    if (!Array.isArray(files)) files = [files];

    const siteName = slugifySiteName(req.body?.siteName || files[0]?.name?.split('.')[0]);
    const basePath = `${B8HOST_ACCOUNT}/${siteName}`;

    // Auto-add an index.html if none was included, so anything uploaded
    // is actually viewable — a bare image or text file gets wrapped in a
    // minimal page rather than 404ing when visited.
    const hasIndex = files.some(f => f.name === 'index.html');
    if (!hasIndex) {
      const firstFile = files[0];
      const isImage = /\.(jpe?g|png|gif|webp|svg)$/i.test(firstFile.name);
      const wrapperHtml = isImage
        ? `<!DOCTYPE html><html><body style="margin:0;background:#111;display:flex;align-items:center;justify-content:center;min-height:100vh"><img src="${firstFile.name}" style="max-width:100%;max-height:100vh"></body></html>`
        : `<!DOCTYPE html><html><body style="font-family:sans-serif;background:#111;color:#eee;padding:40px"><h2>${firstFile.name}</h2><iframe src="${firstFile.name}" style="width:100%;height:80vh;border:1px solid #333"></iframe></body></html>`;
      files.push({ name: 'index.html', data: Buffer.from(wrapperHtml, 'utf-8'), mimetype: 'text/html' });
    }

    for (const file of files) {
      const filePath = `${basePath}/${file.name}`;
      const { error } = await supabase.storage
        .from('sites')
        .upload(filePath, file.data, { upsert: true, contentType: file.mimetype });
      if (error) throw error;
    }

    const siteUrl = `https://b8host.b8build48.eu.cc/${B8HOST_ACCOUNT}/${siteName}/`;

    // Record it the same way b8host2.html's own dashboard does, so it
    // shows up there too, not just as a bare storage upload.
    await supabase.from('sites').upsert({
      username: B8HOST_ACCOUNT,
      site_name: siteName,
      description: 'Hosted via B8AI Director chat',
      url: siteUrl,
      file_count: files.length,
      created_at: new Date().toISOString()
    });

    res.json({ url: siteUrl, fileCount: files.length });
  } catch (err) {
    console.error('host error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ------------------------------------------------------------
// Re-edit an already-hosted file: overwrites one file in an
// existing site with new content, keeping everything else as-is.
// ------------------------------------------------------------
router.post('/v1/host/edit', async (req, res) => {
  try {
    const { siteName, fileName, content } = req.body;
    if (!siteName || !fileName || content === undefined) {
      return res.status(400).json({ error: 'siteName, fileName, and content are required' });
    }

    const filePath = `${B8HOST_ACCOUNT}/${slugifySiteName(siteName)}/${fileName}`;
    const { error } = await supabase.storage
      .from('sites')
      .upload(filePath, Buffer.from(content, 'utf-8'), { upsert: true, contentType: 'text/html' });
    if (error) throw error;

    res.json({ ok: true });
  } catch (err) {
    console.error('host edit error:', err);
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;
