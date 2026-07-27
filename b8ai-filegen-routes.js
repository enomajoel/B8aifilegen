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
// Lightweight Markdown parsing shared by the PDF and DOCX
// renderers. Handles headings (#, ##, ###), **bold** inline
// spans, "- " bullet lists, and "---" horizontal rules.
// Deliberately simple — this isn't a full CommonMark parser,
// just enough to make AI-written prose look properly formatted
// instead of dumping raw markdown symbols onto the page.
// ------------------------------------------------------------
function parseMarkdownLines(markdown) {
  const lines = (markdown || '').split('\n');
  const blocks = [];
  let paragraphBuffer = [];

  const flushParagraph = () => {
    if (paragraphBuffer.length) {
      blocks.push({ type: 'paragraph', text: paragraphBuffer.join(' ').trim() });
      paragraphBuffer = [];
    }
  };

  for (const rawLine of lines) {
    const line = rawLine.trim();

    if (!line) { flushParagraph(); continue; }
    if (/^-{3,}$/.test(line)) { flushParagraph(); blocks.push({ type: 'rule' }); continue; }
    if (/^###\s+/.test(line)) { flushParagraph(); blocks.push({ type: 'h3', text: line.replace(/^###\s+/, '') }); continue; }
    if (/^##\s+/.test(line)) { flushParagraph(); blocks.push({ type: 'h2', text: line.replace(/^##\s+/, '') }); continue; }
    if (/^#\s+/.test(line)) { flushParagraph(); blocks.push({ type: 'h1', text: line.replace(/^#\s+/, '') }); continue; }
    if (/^[-*]\s+/.test(line)) { flushParagraph(); blocks.push({ type: 'bullet', text: line.replace(/^[-*]\s+/, '') }); continue; }

    paragraphBuffer.push(line);
  }
  flushParagraph();
  return blocks;
}

// Splits a line into { text, bold } segments on **bold** markers.
function parseInlineSegments(text) {
  const segments = [];
  const parts = (text || '').split(/(\*\*[^*]+\*\*)/g);
  for (const part of parts) {
    if (!part) continue;
    if (part.startsWith('**') && part.endsWith('**')) {
      segments.push({ text: part.slice(2, -2), bold: true });
    } else {
      segments.push({ text: part, bold: false });
    }
  }
  return segments.length ? segments : [{ text: '', bold: false }];
}

// ------------------------------------------------------------
// PDF markdown renderer — real headings, real bold, real bullets,
// thin rule lines instead of literal "---".
// ------------------------------------------------------------
function renderMarkdownToPDF(doc, markdown) {
  const blocks = parseMarkdownLines(markdown);

  for (const block of blocks) {
    if (block.type === 'rule') {
      doc.moveDown(0.3);
      const y = doc.y;
      doc.strokeColor('#cccccc').moveTo(doc.page.margins.left, y)
        .lineTo(doc.page.width - doc.page.margins.right, y).stroke();
      doc.moveDown(0.5);
      continue;
    }

    if (block.type === 'h1') { doc.moveDown(0.4); doc.fontSize(18).fillColor('#111').font('Helvetica-Bold').text(block.text); doc.moveDown(0.2); continue; }
    if (block.type === 'h2') { doc.moveDown(0.4); doc.fontSize(15).fillColor('#111').font('Helvetica-Bold').text(block.text); doc.moveDown(0.2); continue; }
    if (block.type === 'h3') { doc.moveDown(0.3); doc.fontSize(13).fillColor('#222').font('Helvetica-Bold').text(block.text); doc.moveDown(0.15); continue; }

    if (block.type === 'bullet') {
      doc.fontSize(11).fillColor('#222');
      const segments = parseInlineSegments(block.text);
      doc.text('•  ', { continued: true, indent: 10 });
      segments.forEach((seg, i) => {
        doc.font(seg.bold ? 'Helvetica-Bold' : 'Helvetica');
        doc.text(seg.text, { continued: i < segments.length - 1 });
      });
      doc.moveDown(0.15);
      continue;
    }

    // paragraph
    doc.fontSize(11).fillColor('#222');
    const segments = parseInlineSegments(block.text);
    segments.forEach((seg, i) => {
      doc.font(seg.bold ? 'Helvetica-Bold' : 'Helvetica');
      doc.text(seg.text, { continued: i < segments.length - 1 });
    });
    doc.moveDown(0.4);
  }
}

// ------------------------------------------------------------
// DOCX markdown renderer — same parsing, output as real Word
// headings / bold runs / bullet paragraphs.
// ------------------------------------------------------------
function renderMarkdownToDocxChildren(markdown) {
  const blocks = parseMarkdownLines(markdown);
  const children = [];

  for (const block of blocks) {
    if (block.type === 'rule') {
      children.push(new Paragraph({ text: '', border: { bottom: { color: 'CCCCCC', space: 1, value: 'single', size: 6 } } }));
      continue;
    }
    if (block.type === 'h1') { children.push(new Paragraph({ text: block.text, heading: HeadingLevel.HEADING_1 })); continue; }
    if (block.type === 'h2') { children.push(new Paragraph({ text: block.text, heading: HeadingLevel.HEADING_2 })); continue; }
    if (block.type === 'h3') { children.push(new Paragraph({ text: block.text, heading: HeadingLevel.HEADING_3 })); continue; }

    const runs = parseInlineSegments(block.text).map(seg => new TextRun({ text: seg.text, bold: seg.bold }));

    if (block.type === 'bullet') {
      children.push(new Paragraph({ children: runs, bullet: { level: 0 } }));
      continue;
    }
    children.push(new Paragraph({ children: runs }));
  }
  return children;
}

// ------------------------------------------------------------
// 1) PDF — either a written document (project.content, markdown)
//    or a scene-by-scene storyboard (project.scenes, video mode)
// ------------------------------------------------------------
async function generatePDF(project) {
  const doc = new PDFDocument({ margin: 50 });
  const stream = new PassThrough();
  doc.pipe(stream);

  doc.fontSize(20).font('Helvetica-Bold').fillColor('#111').text(project.title || 'B8AI Director', { align: 'center' });
  doc.moveDown();

  if (project.content) {
    renderMarkdownToPDF(doc, project.content);
  } else {
    (project.scenes || []).forEach((scene, i) => {
      doc.fontSize(14).fillColor('#111').font('Helvetica-Bold').text(`Scene ${i + 1}: ${scene.title || ''}`, { underline: true });
      doc.fontSize(11).fillColor('#444').font('Helvetica').text(scene.description || '');
      if (scene.effects?.length) {
        doc.fontSize(10).fillColor('#777').text(`Effects: ${scene.effects.join(', ')}`);
      }
      doc.moveDown();
    });
  }

  doc.end();
  return streamToBuffer(stream);
}

// ------------------------------------------------------------
// 2) DOCX — same split: written document vs. shot list
// ------------------------------------------------------------
async function generateDOCX(project) {
  const children = [
    new Paragraph({ text: project.title || 'B8AI Director', heading: HeadingLevel.TITLE }),
  ];

  if (project.content) {
    children.push(...renderMarkdownToDocxChildren(project.content));
  } else {
    (project.scenes || []).forEach((scene, i) => {
      children.push(new Paragraph({ text: `Scene ${i + 1}: ${scene.title || ''}`, heading: HeadingLevel.HEADING_2 }));
      children.push(new Paragraph({ children: [new TextRun(scene.description || '')] }));
      if (scene.effects?.length) {
        children.push(new Paragraph({ children: [new TextRun({ text: `Effects: ${scene.effects.join(', ')}`, italics: true })] }));
      }
    });
  }

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

  if (project.content) {
    // Written content — one slide per top-level heading section instead
    // of a meaningless "scene per paragraph" breakdown.
    const blocks = parseMarkdownLines(project.content);
    let current = { heading: null, lines: [] };
    const sections = [];
    for (const b of blocks) {
      if (b.type === 'h1' || b.type === 'h2' || b.type === 'h3') {
        if (current.heading || current.lines.length) sections.push(current);
        current = { heading: b.text, lines: [] };
      } else if (b.type === 'paragraph' || b.type === 'bullet') {
        current.lines.push(b.text.replace(/\*\*/g, ''));
      }
    }
    if (current.heading || current.lines.length) sections.push(current);
    if (!sections.length) sections.push({ heading: null, lines: [project.content] });

    sections.forEach((section) => {
      const slide = pptx.addSlide();
      if (section.heading) slide.addText(section.heading, { x: 0.5, y: 0.3, fontSize: 22, bold: true });
      slide.addText(section.lines.join('\n\n'), { x: 0.5, y: section.heading ? 1.2 : 0.5, fontSize: 13, w: 9 });
    });
  } else {
    (project.scenes || []).forEach((scene, i) => {
      const slide = pptx.addSlide();
      slide.addText(`Scene ${i + 1}: ${scene.title || ''}`, { x: 0.5, y: 0.3, fontSize: 22, bold: true });
      slide.addText(scene.description || '', { x: 0.5, y: 1.2, fontSize: 14, w: 9 });
      if (scene.effects?.length) {
        slide.addText(`Effects: ${scene.effects.join(', ')}`, { x: 0.5, y: 4.5, fontSize: 11, italic: true, color: '888888' });
      }
    });
  }

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
  if (project.content) {
    // Strip markdown symbols for a clean plain-text read
    const plain = project.content
      .replace(/^#{1,3}\s+/gm, '')
      .replace(/\*\*(.+?)\*\*/g, '$1')
      .replace(/^-{3,}$/gm, '—'.repeat(20))
      .replace(/^[-*]\s+/gm, '• ');
    return Buffer.from(`${project.title || 'B8AI Project'}\n${'='.repeat(40)}\n\n${plain}`, 'utf-8');
  }
  let txt = `${project.title || 'B8AI Project'}\n${'='.repeat(40)}\n\n`;
  (project.scenes || []).forEach((scene, i) => {
    txt += `SCENE ${i + 1}: ${scene.title || ''}\n${scene.description || ''}\n\n`;
  });
  return Buffer.from(txt, 'utf-8');
}

function generateReadmeMD(project) {
  if (project.content) {
    return Buffer.from(`# ${project.title || 'B8AI Project'}\n\n${project.content}`, 'utf-8');
  }
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
  if (project.content) {
    // Not a scene-based project — CSV/spreadsheet format doesn't apply
    // to prose. Return a short note instead of fabricating fake rows.
    return Buffer.from('Note,This project is written content, not scene data — a spreadsheet format doesn\'t apply. See the PDF/DOCX/MD instead.\n', 'utf-8');
  }
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
      : project.content
        ? ['pdf', 'docx', 'md', 'txt']              // written content — no fake video-only formats
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
