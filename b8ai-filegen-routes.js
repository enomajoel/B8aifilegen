// ============================================================
// B8AI Director — Advanced File Generation Module
// Drop into your existing Render Node/Express backend.
// Generates every practical project file type and bundles
// them into a single ZIP for download.
//
// Install dependencies on your Render backend:
//   npm install archiver pdfkit docx pptxgenjs exceljs
// ============================================================

const express = require('express');
const archiver = require('archiver');
const PDFDocument = require('pdfkit');
const { Document, Packer, Paragraph, HeadingLevel, TextRun } = require('docx');
const PptxGenJS = require('pptxgenjs');
const ExcelJS = require('exceljs');
const { PassThrough } = require('stream');

const router = express.Router();

// ------------------------------------------------------------
// Helper: turn any generator into an in-memory Buffer
// ------------------------------------------------------------
function streamToBuffer(stream) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    stream.on('data', (c) => chunks.push(c));
    stream.on('end', () => resolve(Buffer.concat(chunks)));
    stream.on('error', reject);
  });
}

// ------------------------------------------------------------
// 1) PDF — Storyboard / shot-by-shot breakdown
// ------------------------------------------------------------
async function generatePDF(project) {
  const doc = new PDFDocument({ margin: 50 });
  const stream = new PassThrough();
  doc.pipe(stream);

  doc.fontSize(20).text(project.title || 'B8AI Director — Storyboard', { align: 'center' });
  doc.moveDown();

  (project.scenes || []).forEach((scene, i) => {
    doc.fontSize(14).fillColor('black').text(`Scene ${i + 1}: ${scene.title || ''}`, { underline: true });
    doc.fontSize(11).fillColor('gray').text(scene.description || '');
    if (scene.effects?.length) {
      doc.fontSize(10).fillColor('#555').text(`Effects: ${scene.effects.join(', ')}`);
    }
    doc.moveDown();
  });

  doc.end();
  return streamToBuffer(stream);
}

// ------------------------------------------------------------
// 2) DOCX — Shot list document
// ------------------------------------------------------------
async function generateDOCX(project) {
  const children = [
    new Paragraph({ text: project.title || 'B8AI Director — Shot List', heading: HeadingLevel.TITLE }),
  ];

  (project.scenes || []).forEach((scene, i) => {
    children.push(new Paragraph({ text: `Scene ${i + 1}: ${scene.title || ''}`, heading: HeadingLevel.HEADING_2 }));
    children.push(new Paragraph({ children: [new TextRun(scene.description || '')] }));
    if (scene.effects?.length) {
      children.push(new Paragraph({ children: [new TextRun({ text: `Effects: ${scene.effects.join(', ')}`, italics: true })] }));
    }
  });

  const doc = new Document({ sections: [{ children }] });
  return Packer.toBuffer(doc);
}

// ------------------------------------------------------------
// 3) PPTX — Slide-per-scene presentation
// ------------------------------------------------------------
async function generatePPTX(project) {
  const pptx = new PptxGenJS();

  const titleSlide = pptx.addSlide();
  titleSlide.addText(project.title || 'B8AI Director Project', { x: 0.5, y: 2, fontSize: 32, bold: true });

  (project.scenes || []).forEach((scene, i) => {
    const slide = pptx.addSlide();
    slide.addText(`Scene ${i + 1}: ${scene.title || ''}`, { x: 0.5, y: 0.3, fontSize: 22, bold: true });
    slide.addText(scene.description || '', { x: 0.5, y: 1.2, fontSize: 14, w: 9 });
    if (scene.effects?.length) {
      slide.addText(`Effects: ${scene.effects.join(', ')}`, { x: 0.5, y: 4.5, fontSize: 11, italic: true, color: '888888' });
    }
  });

  return pptx.write('nodebuffer');
}

// ------------------------------------------------------------
// 4) XLSX — Schedule / budget / shot tracker
// ------------------------------------------------------------
async function generateXLSX(project) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Shot Tracker');

  sheet.columns = [
    { header: 'Scene #', key: 'num', width: 10 },
    { header: 'Title', key: 'title', width: 30 },
    { header: 'Description', key: 'desc', width: 50 },
    { header: 'Effects', key: 'effects', width: 40 },
    { header: 'Status', key: 'status', width: 15 },
  ];
  sheet.getRow(1).font = { bold: true };

  (project.scenes || []).forEach((scene, i) => {
    sheet.addRow({
      num: i + 1,
      title: scene.title || '',
      desc: scene.description || '',
      effects: (scene.effects || []).join(', '),
      status: 'Pending',
    });
  });

  return workbook.xlsx.writeBuffer();
}

// ------------------------------------------------------------
// 5) SRT — Subtitle file (if project has narration/dialogue)
// ------------------------------------------------------------
function generateSRT(project) {
  let srt = '';
  let index = 1;
  let time = 0;
  (project.scenes || []).forEach((scene) => {
    if (!scene.narration) return;
    const start = formatSRTTime(time);
    const duration = scene.duration || 4;
    time += duration;
    const end = formatSRTTime(time);
    srt += `${index}\n${start} --> ${end}\n${scene.narration}\n\n`;
    index++;
  });
  return Buffer.from(srt, 'utf-8');
}

function formatSRTTime(seconds) {
  const h = String(Math.floor(seconds / 3600)).padStart(2, '0');
  const m = String(Math.floor((seconds % 3600) / 60)).padStart(2, '0');
  const s = String(Math.floor(seconds % 60)).padStart(2, '0');
  const ms = String(Math.floor((seconds % 1) * 1000)).padStart(3, '0');
  return `${h}:${m}:${s},${ms}`;
}

// ------------------------------------------------------------
// 6) Plain-text / Markdown / JSON / CSV — always included
// ------------------------------------------------------------
function generateScriptTXT(project) {
  let txt = `${project.title || 'B8AI Project'}\n${'='.repeat(40)}\n\n`;
  (project.scenes || []).forEach((scene, i) => {
    txt += `SCENE ${i + 1}: ${scene.title || ''}\n${scene.description || ''}\n\n`;
  });
  return Buffer.from(txt, 'utf-8');
}

function generateReadmeMD(project) {
  let md = `# ${project.title || 'B8AI Project'}\n\n`;
  md += `Generated by B8AI Director on ${new Date().toISOString()}\n\n`;
  md += `## Scenes (${(project.scenes || []).length})\n\n`;
  (project.scenes || []).forEach((scene, i) => {
    md += `### Scene ${i + 1}: ${scene.title || ''}\n${scene.description || ''}\n\n`;
  });
  return Buffer.from(md, 'utf-8');
}

function generateProjectJSON(project) {
  return Buffer.from(JSON.stringify(project, null, 2), 'utf-8');
}

function generateSceneCSV(project) {
  let csv = 'Scene,Title,Description,Effects,Duration\n';
  (project.scenes || []).forEach((scene, i) => {
    const row = [
      i + 1,
      `"${(scene.title || '').replace(/"/g, '""')}"`,
      `"${(scene.description || '').replace(/"/g, '""')}"`,
      `"${(scene.effects || []).join(', ')}"`,
      scene.duration || '',
    ];
    csv += row.join(',') + '\n';
  });
  return Buffer.from(csv, 'utf-8');
}

// ------------------------------------------------------------
// Main route: POST /v1/generate-files
// Body: { project: { title, scenes: [...] }, formats: [...] }
// Returns: a ZIP stream containing every requested/available file
// ------------------------------------------------------------
router.post('/v1/generate-files', async (req, res) => {
  try {
    const { project, formats } = req.body;
    if (!project) return res.status(400).json({ error: 'project is required' });

    const wanted = formats && formats.length
      ? formats
      : ['pdf', 'docx', 'pptx', 'xlsx', 'srt', 'txt', 'md', 'json', 'csv'];

    const generators = {
      pdf:  { name: 'storyboard.pdf',    fn: generatePDF },
      docx: { name: 'shotlist.docx',     fn: generateDOCX },
      pptx: { name: 'presentation.pptx', fn: generatePPTX },
      xlsx: { name: 'tracker.xlsx',      fn: generateXLSX },
      srt:  { name: 'subtitles.srt',     fn: async (p) => generateSRT(p) },
      txt:  { name: 'script.txt',        fn: async (p) => generateScriptTXT(p) },
      md:   { name: 'README.md',         fn: async (p) => generateReadmeMD(p) },
      json: { name: 'project.json',      fn: async (p) => generateProjectJSON(p) },
      csv:  { name: 'scenes.csv',        fn: async (p) => generateSceneCSV(p) },
    };

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="${(project.title || 'b8ai-project').replace(/\s+/g, '_')}.zip"`);

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.pipe(res);

    for (const key of wanted) {
      const gen = generators[key];
      if (!gen) continue;
      try {
        const buffer = await gen.fn(project);
        archive.append(buffer, { name: gen.name });
      } catch (err) {
        console.error(`Failed generating ${key}:`, err.message);
        archive.append(Buffer.from(`Failed to generate ${key}: ${err.message}`), { name: `${key}-ERROR.txt` });
      }
    }

    await archive.finalize();
  } catch (err) {
    console.error('generate-files error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// ------------------------------------------------------------
// Optional: generate ONE file type only (no zip), used by the
// frontend "download single file" button in the file box UI.
// GET /v1/generate-file/:format
// ------------------------------------------------------------
router.post('/v1/generate-file/:format', async (req, res) => {
  const { format } = req.params;
  const { project } = req.body;
  if (!project) return res.status(400).json({ error: 'project is required' });

  const single = {
    pdf:  { name: 'storyboard.pdf',    mime: 'application/pdf', fn: generatePDF },
    docx: { name: 'shotlist.docx',     mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', fn: generateDOCX },
    pptx: { name: 'presentation.pptx', mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', fn: generatePPTX },
    xlsx: { name: 'tracker.xlsx',      mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fn: generateXLSX },
    srt:  { name: 'subtitles.srt',     mime: 'text/plain', fn: async (p) => generateSRT(p) },
    txt:  { name: 'script.txt',        mime: 'text/plain', fn: async (p) => generateScriptTXT(p) },
    md:   { name: 'README.md',         mime: 'text/markdown', fn: async (p) => generateReadmeMD(p) },
    json: { name: 'project.json',      mime: 'application/json', fn: async (p) => generateProjectJSON(p) },
    csv:  { name: 'scenes.csv',        mime: 'text/csv', fn: async (p) => generateSceneCSV(p) },
  }[format];

  if (!single) return res.status(400).json({ error: `Unknown format: ${format}` });

  try {
    const buffer = await single.fn(project);
    res.setHeader('Content-Type', single.mime);
    res.setHeader('Content-Disposition', `attachment; filename="${single.name}"`);
    res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ------------------------------------------------------------
// Repackage: takes the current (possibly user-edited) files
// array from the frontend and re-zips them.
// - Text files (content included) go in as-is, edits included
// - Binary files (pdf/docx/pptx/xlsx) are re-fetched from their
//   original URL, since binary content isn't edited in-browser
// POST /v1/repackage
// Body: { files: [{ name, format, content?, url? }] }
// ------------------------------------------------------------
router.post('/v1/repackage', async (req, res) => {
  try {
    const { files } = req.body;
    if (!files || !Array.isArray(files) || files.length === 0) {
      return res.status(400).json({ error: 'files[] is required' });
    }

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="b8ai_project_repackaged.zip"');

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.pipe(res);

    for (const file of files) {
      try {
        if (file.content !== undefined && file.content !== null) {
          // Edited text file — use the (possibly edited) content directly
          archive.append(Buffer.from(file.content, 'utf-8'), { name: file.name });
        } else if (file.url) {
          // Binary file — re-fetch original bytes from where it was generated
          const response = await fetch(file.url);
          if (!response.ok) throw new Error(`Fetch failed: ${response.status}`);
          const buffer = Buffer.from(await response.arrayBuffer());
          archive.append(buffer, { name: file.name });
        } else {
          archive.append(Buffer.from(''), { name: file.name });
        }
      } catch (err) {
        console.error(`Failed to repackage ${file.name}:`, err.message);
        archive.append(Buffer.from(`Failed to include ${file.name}: ${err.message}`), { name: `${file.name}-ERROR.txt` });
      }
    }

    await archive.finalize();
  } catch (err) {
    console.error('repackage error:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

module.exports = router;
