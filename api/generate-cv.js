const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, LevelFormat, TabStopType, TabStopPosition } = require('docx');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }
  if (req.method !== 'POST') { res.status(405).json({ error: 'Method not allowed' }); return; }

  try {
    const { cv, profile } = req.body;
    if (!cv) { res.status(400).json({ error: 'No CV rewrite data provided' }); return; }

    const lineColor = 'CCCCCC';
    const mutedColor = '888888';
    const darkColor = '1a1a1a';

    const divider = () => new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: lineColor, space: 1 } },
      spacing: { before: 160, after: 160 }
    });

    const sectionHead = (text) => new Paragraph({
      children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 20, font: 'Arial', color: darkColor, characterSpacing: 80 })],
      spacing: { before: 200, after: 100 }
    });

    const competencies = [
      'Training Program Design', 'Performance Coaching',
      'Team Leadership & Development', 'Process & SOP Development',
      'Stakeholder Management', 'Contact Center Operations',
      'Revenue Operations', 'L&D Strategy'
    ];

    const compRows = [];
    for (let i = 0; i < competencies.length; i += 2) {
      compRows.push(new Paragraph({
        tabStops: [{ type: TabStopType.LEFT, position: 4500 }],
        children: [
          new TextRun({ text: `\u2022  ${competencies[i] || ''}`, size: 20, font: 'Arial' }),
          ...(competencies[i+1] ? [new TextRun({ text: `\t\u2022  ${competencies[i+1]}`, size: 20, font: 'Arial' })] : [])
        ],
        spacing: { after: 60 }
      }));
    }

    const doc = new Document({
      numbering: {
        config: [{
          reference: 'bullets',
          levels: [{
            level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 360, hanging: 240 } } }
          }]
        }]
      },
      styles: {
        default: { document: { run: { font: 'Arial', size: 22, color: '1a1a1a' } } }
      },
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
          }
        },
        children: [
          // NAME
          new Paragraph({
            children: [new TextRun({ text: profile?.name || 'Your Name', bold: true, size: 48, font: 'Arial', color: darkColor })],
            spacing: { after: 60 }
          }),

          // HEADLINE
          new Paragraph({
            children: [new TextRun({ text: cv.target_headline || 'Professional Headline', size: 26, font: 'Arial', color: '444444' })],
            spacing: { after: 80 }
          }),

          // Contact placeholder
          new Paragraph({
            children: [new TextRun({ text: 'Mumbai, India  |  your@email.com  |  linkedin.com/in/yourprofile  |  +91 XXXXX XXXXX', size: 18, color: mutedColor, font: 'Arial' })],
            spacing: { after: 40 }
          }),

          new Paragraph({
            children: [new TextRun({ text: '[Update contact details above before sending]', size: 16, color: 'CC0000', font: 'Arial', italics: true })],
            spacing: { after: 160 }
          }),

          divider(),
          sectionHead('Professional Summary'),
          new Paragraph({
            children: [new TextRun({ text: cv.summary_paragraph || '', size: 22, font: 'Arial' })],
            alignment: AlignmentType.JUSTIFIED,
            spacing: { after: 200 }
          }),

          divider(),
          sectionHead('Core Competencies'),
          ...compRows,
          new Paragraph({ children: [new TextRun('')], spacing: { after: 80 } }),

          divider(),
          sectionHead('Key Achievements'),
          ...(cv.top_bullets || []).map(b =>
            new Paragraph({
              numbering: { reference: 'bullets', level: 0 },
              children: [new TextRun({ text: b, size: 22, font: 'Arial' })],
              spacing: { after: 100 }
            })
          ),

          divider(),
          sectionHead('Professional Experience'),
          new Paragraph({
            children: [new TextRun({
              text: 'INSTRUCTION: Paste your existing experience section here, under the achievements above. Keep your company names, titles, and dates. The achievements section above already captures your strongest proof points.',
              size: 18, color: mutedColor, font: 'Arial', italics: true
            })],
            spacing: { after: 160 }
          }),

          divider(),
          sectionHead('Education & Certifications'),
          new Paragraph({
            children: [new TextRun({
              text: 'INSTRUCTION: Add your education and any relevant certifications here.',
              size: 18, color: mutedColor, font: 'Arial', italics: true
            })],
            spacing: { after: 160 }
          }),

          divider(),
          new Paragraph({
            children: [new TextRun({
              text: 'This document is ATS-compliant: single column, standard fonts (Arial), no images or tables. Safe to upload directly to Naukri, LinkedIn, or Workday.',
              size: 16, color: mutedColor, font: 'Arial', italics: true
            })],
            spacing: { before: 80 }
          }),
        ]
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    const safeName = (profile?.name || 'Professional').replace(/[^a-zA-Z0-9]/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="HireOS_CV_${safeName}.docx"`);
    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || 'CV generation failed' });
  }
};
