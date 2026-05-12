const mammoth = require('mammoth');
const cheerio = require('cheerio');
const fs = require('fs');
const PizZip = require('pizzip');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  VerticalAlign,
  convertInchesToTwip,
  PageBreak,
  UnderlineType,
  ShadingType,
  Header,
  Footer,
  PageNumber,
  LineRuleType,
  NumberFormat,
  SectionType,
} = require('docx');
const DocxMerger = require('docx-merger');
const path = require('path');

function cmToTwip(cm) {
  return Math.round(cm * 567);
}

function ptToHalfPoint(pt) {
  return pt * 2;
}

async function processWord(inputPath, outputPath, options = {}) {
  const {
    coverType = '2026',
    fontFamily = 'Times New Roman',
    fontSize = 11,
    lineSpacing = 1.5,
    marginTop = 3,
    marginBottom = 3,
    marginLeft = 4,
    marginRight = 3,
    removeImages = true,
    fixTables = true,
    fixHeadings = true,
  } = options;

  const report = {
    headingsFixed: 0,
    tablesFixed: 0,
    imagesRemoved: 0,
    emptyLinesRemoved: 0,
    formattingFixed: 0,
    totalParagraphs: 0,
  };

  const docBuffer = fs.readFileSync(inputPath);

  const result = await mammoth.convertToHtml(docBuffer, {
    styleMap: [
      "p[style-name='Heading 1'] => h1:fresh",
      "p[style-name='Heading 2'] => h2:fresh",
      "p[style-name='Heading 3'] => h3:fresh",
      "p[style-name='Heading 4'] => h4:fresh",
      "p[style-name='Title'] => h1:fresh",
      "p[style-name='Subtitle'] => h2:fresh",
    ],
    includeDefaultStyleMap: true,
  });

  const html = result.value;
  const $ = cheerio.load(html, { decodeEntities: false });

  // ============================================================
  // PHASE 1: CLEAN THE HTML
  // ============================================================

  if (removeImages) {
    $('img').each((i, el) => {
      report.imagesRemoved++;
      $(el).remove();
    });
  }

  $('p').each((i, el) => {
    const text = $(el).text().trim();
    if (text === '' && $(el).children().length === 0) {
      report.emptyLinesRemoved++;
      $(el).remove();
    }
  });

  $('body *').each((i, el) => {
    $(el).contents().each((j, node) => {
      if (node.type === 'text') {
        let text = $(node).text();
        text = text.replace(/[\u200B\u200C\u200D\uFEFF]/g, '');
        text = text.replace(/  +/g, ' ');
        text = text.replace(/\u00A0/g, ' ');
        $(node).replaceWith(text);
      }
    });
  });

  // ============================================================
  // PHASE 2: PARSE STRUCTURE INTO BLOCKS
  // ============================================================

  const blocks = [];

  // ---- SMART HEADING DETECTION ----
  // Detect headings by pattern — covers Indonesian formal doc styles
  function looksLikeHeading(text) {
    if (!text || text.trim().length === 0) return false;
    const t = text.trim();
    // Too long to be a heading (more than ~120 chars is likely a paragraph)
    if (t.length > 120) return false;

    const patterns = [
      // BAB I, BAB II, CHAPTER 1
      /^(BAB|BAGIAN|CHAPTER)\s+[IVXLCDM\d]+/i,
      // Letter-prefixed: A. Pendahuluan, B. Tujuan, etc.
      /^[A-Z]\.\s+\S/,
      // Roman numeral prefix: I. Format, II. Tabel, III. Hasil
      /^[IVXLCDM]+\.\s+\S/,
      // Numbered: 1. Judul, 2. Metode (short text, likely heading)
      /^\d+\.\s+[A-Z]/,
      // Sub-numbered: 1.1 Latar Belakang, 2.3.1 Data
      /^\d+\.\d+[\.\s]/,
      // ALL CAPS words (likely section title), at least 5 chars
      /^[A-Z][A-Z\s]{4,}$/,
      // Known heading keywords
      /^(PENDAHULUAN|TINJAUAN PUSTAKA|METODOLOGI|METODE PENELITIAN|HASIL DAN PEMBAHASAN|PEMBAHASAN|KESIMPULAN|SARAN|DAFTAR PUSTAKA|DAFTAR ISI|KATA PENGANTAR|ABSTRAK|ABSTRACT|LAMPIRAN|PENUTUP|LANDASAN TEORI|VISI|MISI|TUJUAN|SASARAN|STRATEGI|CAPAIAN|NARASI|INDIKATOR|PROGRAM|KEBIJAKAN|EVALUASI|REKOMENDASI|RINGKASAN)/i,
      // Tabel X / Gambar X caption-heading
      /^(Tabel|Gambar|Table|Figure)\s+\d+/i,
    ];
    return patterns.some((p) => p.test(t));
  }

  // ---- DETECT SHORT LABEL / METADATA LINES ----
  // Lines like "STIE Pancasetia Banjarmasin", "Program Studi Manajemen",
  // "Nama: ...", "NIM: ..." etc. — short, not a real paragraph, no first-line indent
  function looksLikeLabel(text) {
    if (!text || text.trim().length === 0) return false;
    const t = text.trim();
    // Must be short (not a real paragraph sentence)
    if (t.length > 100) return false;
    // If it ends with a period/comma followed by more text, it's a sentence
    // Labels typically don't end with period
    const endsLikeSentence = /[.!?]$/.test(t);
    // If it's very short and doesn't end with period — label
    if (t.length <= 60 && !endsLikeSentence) return true;
    // Colon-based labels: "Nama : Fikri", "NIM: 123"
    if (/^[A-Za-z\s]+\s*:\s*.+/.test(t) && t.length < 80) return true;
    // Known patterns: "Program Studi ...", "Fakultas ...", "Universitas ..."
    if (/^(Program Studi|Fakultas|Jurusan|Universitas|Institut|Sekolah Tinggi|STIE|STMIK|STKIP|Politeknik|Akademi)/i.test(t)) return true;
    // "Tahun Akademik", "Semester", date-like labels
    if (/^(Tahun|Semester|Periode|Tanggal|Hari|Waktu|Tempat|Lokasi|Ruang)/i.test(t) && t.length < 60) return true;
    return false;
  }

  // Determine heading level
  function getHeadingLevel(text) {
    const t = text.trim();
    // Level 1: BAB, ALL CAPS, main keywords
    if (/^(BAB|CHAPTER)\s+[IVXLCDM\d]+/i.test(t)) return 1;
    if (/^[A-Z][A-Z\s]{4,}$/.test(t)) return 1;
    if (/^(PENDAHULUAN|TINJAUAN PUSTAKA|METODOLOGI|KESIMPULAN|SARAN|DAFTAR PUSTAKA|DAFTAR ISI|KATA PENGANTAR|ABSTRAK|ABSTRACT|LAMPIRAN|PENUTUP)/i.test(t)) return 1;

    // Level 2: Letter prefix (A. B. C.), Roman (I. II.), numbered (1. 2.), sub-num (1.1)
    if (/^[A-Z]\.\s+/.test(t)) return 2;
    if (/^[IVXLCDM]+\.\s+/.test(t)) return 2;
    if (/^\d+\.\s+[A-Z]/.test(t)) return 2;
    if (/^\d+\.\d+\s/.test(t) && !/^\d+\.\d+\.\d+/.test(t)) return 2;

    // Level 3: sub-sub (1.1.1), Tabel/Gambar
    if (/^\d+\.\d+\.\d+/.test(t)) return 3;
    if (/^(Tabel|Gambar|Table|Figure)\s+\d+/i.test(t)) return 3;

    return 2;
  }

  // Should this heading be centered? Only main doc titles & BAB
  function shouldCenter(text) {
    const t = text.trim();
    return (
      /^(BAB|CHAPTER)\s+[IVXLCDM\d]+/i.test(t) ||
      /^[A-Z][A-Z\s]{4,}$/.test(t) // ALL CAPS titles
    );
  }

  // ---- EXTRACT TEXT RUNS (with inline formatting) ----
  function extractTextRuns(element, $ref) {
    const runs = [];

    function extractInlineRuns(node, $ref, styles = {}) {
      const inlineRuns = [];
      if (node.type === 'text') {
        const text = $ref(node).text();
        if (text) {
          inlineRuns.push({
            text,
            bold: styles.bold || false,
            italic: styles.italic || false,
            underline: styles.underline || false,
            superScript: styles.superScript || false,
            subScript: styles.subScript || false,
          });
        }
      } else if (node.type === 'tag') {
        const tagName = node.tagName.toLowerCase();
        const children = $ref(node).contents();
        const newStyles = { ...styles };

        if (tagName === 'strong' || tagName === 'b') newStyles.bold = true;
        if (tagName === 'em' || tagName === 'i') newStyles.italic = true;
        if (tagName === 'u') newStyles.underline = true;
        if (tagName === 'sup') newStyles.superScript = true;
        if (tagName === 'sub') newStyles.subScript = true;

        if (tagName === 'br') {
          inlineRuns.push({ text: '\n', bold: false, italic: false, underline: false });
        } else if (tagName === 'img') {
          // skip
        } else {
          children.each((i, child) => {
            inlineRuns.push(...extractInlineRuns(child, $ref, newStyles));
          });
        }
      }
      return inlineRuns;
    }

    $ref(element)
      .contents()
      .each((i, node) => {
        runs.push(...extractInlineRuns(node, $ref, {}));
      });

    return runs;
  }

  // ---- SMART BOLD DETECTION ----
  // If ALL runs in a paragraph are bold, it's likely a formatting error
  // from copy-paste. Strip the bold unless it's a heading.
  function smartFixBold(runs, isHeading) {
    if (isHeading) {
      // Headings should always be bold
      return runs.map((r) => ({ ...r, bold: true }));
    }

    // Check what percentage of text is bold
    let totalChars = 0;
    let boldChars = 0;
    for (const r of runs) {
      const len = (r.text || '').trim().length;
      totalChars += len;
      if (r.bold) boldChars += len;
    }

    if (totalChars === 0) return runs;

    const boldRatio = boldChars / totalChars;

    // If more than 80% of the paragraph is bold, it's probably a copy-paste
    // artifact — strip all bold from body text
    if (boldRatio > 0.8) {
      return runs.map((r) => ({ ...r, bold: false }));
    }

    // Otherwise keep the selective bold as-is (intentional emphasis)
    return runs;
  }

  // Process each top-level element
  $('body')
    .children()
    .each((i, el) => {
      const tagName = el.tagName ? el.tagName.toLowerCase() : '';
      const text = $(el).text().trim();

      if (!text && tagName !== 'table' && tagName !== 'hr') return;

      // Skip decorative lines
      if (tagName === 'hr') {
        report.imagesRemoved++;
        return;
      }

      // Handle headings from original doc
      if (['h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(tagName)) {
        const level = parseInt(tagName.charAt(1));
        blocks.push({
          type: 'heading',
          level,
          text,
          runs: extractTextRuns(el, $),
        });
        report.headingsFixed++;
        return;
      }

      // Handle tables
      if (tagName === 'table') {
        const rows = [];
        $(el)
          .find('tr')
          .each((ri, tr) => {
            const cells = [];
            $(tr)
              .find('td, th')
              .each((ci, td) => {
                const isHeader = td.tagName.toLowerCase() === 'th';
                cells.push({
                  text: $(td).text().trim(),
                  runs: extractTextRuns(td, $),
                  isHeader,
                });
              });
            if (cells.length > 0) {
              rows.push(cells);
            }
          });
        if (rows.length > 0) {
          blocks.push({ type: 'table', rows });
          report.tablesFixed++;
        }
        return;
      }

      // Handle lists
      if (tagName === 'ul' || tagName === 'ol') {
        const isOrdered = tagName === 'ol';
        $(el)
          .children('li')
          .each((li_i, li) => {
            blocks.push({
              type: 'list-item',
              ordered: isOrdered,
              index: li_i,
              text: $(li).text().trim(),
              runs: extractTextRuns(li, $),
            });
          });
        return;
      }

      // Handle regular paragraphs
      if (tagName === 'p' || tagName === 'div' || tagName === 'span') {
        // Smart heading detection
        if (fixHeadings && looksLikeHeading(text)) {
          const level = getHeadingLevel(text);
          blocks.push({
            type: 'heading',
            level,
            text,
            runs: extractTextRuns(el, $),
          });
          report.headingsFixed++;
          return;
        }

        // Broken table detection
        if (fixTables && text.includes('\t') && text.split('\t').length >= 3) {
          blocks.push({
            type: 'table-candidate',
            text,
            runs: extractTextRuns(el, $),
          });
          return;
        }

        blocks.push({
          type: 'paragraph',
          text,
          runs: extractTextRuns(el, $),
        });
        report.totalParagraphs++;
        return;
      }
    });

  // ============================================================
  // PHASE 3: MERGE TABLE CANDIDATES
  // ============================================================

  const mergedBlocks = [];
  let tableCandidateBuffer = [];

  function flushTableCandidates() {
    if (tableCandidateBuffer.length >= 2) {
      const rows = tableCandidateBuffer.map((b) => {
        return b.text.split('\t').map((cell) => ({
          text: cell.trim(),
          runs: [{ text: cell.trim(), bold: false, italic: false, underline: false }],
          isHeader: false,
        }));
      });
      if (rows.length > 0) {
        rows[0].forEach((cell) => (cell.isHeader = true));
      }
      mergedBlocks.push({ type: 'table', rows });
      report.tablesFixed++;
    } else {
      tableCandidateBuffer.forEach((b) => {
        mergedBlocks.push({ type: 'paragraph', text: b.text, runs: b.runs });
      });
    }
    tableCandidateBuffer = [];
  }

  for (const block of blocks) {
    if (block.type === 'table-candidate') {
      tableCandidateBuffer.push(block);
    } else {
      if (tableCandidateBuffer.length > 0) flushTableCandidates();
      mergedBlocks.push(block);
    }
  }
  if (tableCandidateBuffer.length > 0) flushTableCandidates();

  // ============================================================
  // PHASE 4: GENERATE CLEAN DOCX
  // ============================================================

  const lineSpacingValue = Math.round(lineSpacing * 240);
  const frontMatterElements = [];
  const mainBodyElements = [];
  const fontSizeHalf = ptToHalfPoint(fontSize);

  // ============================================================
  // BUILD MANUAL TOC WITH PAGE NUMBER ESTIMATION
  // ============================================================

  // Collect all headings from mergedBlocks (in order)
  const tocHeadings = mergedBlocks
    .filter(b => b.type === 'heading')
    .map(b => ({ level: b.level, text: b.text }));

  // Estimate page numbers for TOC:
  // Body starts at page 1 after the cover+TOC section.
  // Heuristic: count paragraphs/tables between headings, divide by ~4 per page.
  const PARAGRAPHS_PER_PAGE = 4;
  const tocPageNums = [];
  let accumBlocks = 0;
  for (const block of mergedBlocks) {
    if (block.type === 'heading') {
      tocPageNums.push(Math.max(1, Math.ceil(accumBlocks / PARAGRAPHS_PER_PAGE)));
    } else if (block.type === 'paragraph' || block.type === 'list-item') {
      accumBlocks += 1;
    } else if (block.type === 'table') {
      // Tables occupy more vertical space
      accumBlocks += Math.ceil((block.rows?.length || 3) / 2) + 1;
    }
  }

  // Add DAFTAR ISI title
  frontMatterElements.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: "DAFTAR ISI",
          bold: true,
          font: fontFamily,
          size: fontSizeHalf,
          color: '000000',
        }),
      ],
      spacing: { before: 0, after: 360 },
    })
  );

  // Halaman header row
  frontMatterElements.push(
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [
        new TextRun({
          text: "Halaman",
          bold: true,
          italics: true,
          font: fontFamily,
          size: fontSizeHalf,
          color: '000000',
        }),
      ],
      spacing: { before: 0, after: 120 },
    })
  );

  // Build each TOC row
  // Tab stop position at right margin (in twips). Page width minus margins = ~13.5cm body ≈ 7650 twips
  const TOC_TAB_STOP = cmToTwip(13.5);

  tocHeadings.forEach((heading, idx) => {
    const pg = tocPageNums[idx] || 1;
    const level = heading.level;

    // Indentation per level
    const indentCm = level === 1 ? 0 : level === 2 ? 0.5 : level === 3 ? 1.0 : 1.5;
    const indentTwip = cmToTwip(indentCm);

    const isBold = level === 1;

    frontMatterElements.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        tabStops: [
          {
            type: "right",
            position: TOC_TAB_STOP,
            leader: "dot",
          },
        ],
        spacing: {
          before: level === 1 ? 160 : 60,
          after: level === 1 ? 80 : 40,
          line: Math.round(1.15 * 240),
          lineRule: LineRuleType.AUTO,
        },
        indent: {
          left: indentTwip,
          firstLine: 0,
        },
        children: [
          new TextRun({
            text: heading.text,
            bold: isBold,
            font: fontFamily,
            size: fontSizeHalf,
            color: '000000',
          }),
          new TextRun({
            text: "\t" + pg,
            bold: isBold,
            font: fontFamily,
            size: fontSizeHalf,
            color: '000000',
          }),
        ],
      })
    );
  });

  // Add page break after TOC
  frontMatterElements.push(
    new Paragraph({
      children: [new PageBreak()],
    })
  );

  let isMainBody = false;

  const tableBorderStyle = {
    style: BorderStyle.SINGLE,
    size: 1,
    color: '000000',
  };

  for (const block of mergedBlocks) {
    if (block.type === 'heading') {
      if (!isMainBody) {
        const t = block.text.trim().toUpperCase();
        // If it's a known front matter heading, keep isMainBody = false
        const isFrontMatterHeading = /^(KATA PENGANTAR|DAFTAR (ISI|TABEL|GAMBAR|LAMPIRAN)|ABSTRAK|ABSTRACT)/.test(t);
        if (!isFrontMatterHeading) {
          isMainBody = true;
        }
      }

      // ---- HEADING: no indent, flush left (or centered for titles) ----
      const centered = shouldCenter(block.text);
      const alignment = centered ? AlignmentType.CENTER : AlignmentType.LEFT;

      // All headings same font size as body, just BOLD
      const runs = smartFixBold(block.runs, true); // force bold for headings

      let spaceBefore = 240;
      let spaceAfter = 120;
      let headingLvl = HeadingLevel.HEADING_2;
      if (block.level === 1) {
        spaceBefore = 360;
        spaceAfter = 240;
        headingLvl = HeadingLevel.HEADING_1;
      } else if (block.level === 3) {
        headingLvl = HeadingLevel.HEADING_3;
      } else if (block.level >= 4) {
        headingLvl = HeadingLevel.HEADING_4;
      }

      const targetArray = isMainBody ? mainBodyElements : frontMatterElements;

      targetArray.push(
        new Paragraph({
          heading: headingLvl,
          alignment,
          spacing: {
            before: spaceBefore,
            after: spaceAfter,
            line: lineSpacingValue,
            lineRule: LineRuleType.AUTO,
          },
          // NO indent whatsoever on headings
          indent: {
            left: 0,
            right: 0,
            firstLine: 0,
            hanging: 0,
          },
          children: [
            new TextRun({
              text: block.text,
              bold: true,
              font: fontFamily,
              size: fontSizeHalf,
              color: '000000',
            }),
          ],
        })
      );
      report.headingsFixed++;
    } else if (block.type === 'paragraph') {
      // ---- PARAGRAPH: smart bold handling ----
      let fixedRuns = smartFixBold(block.runs, false);

      const children = fixedRuns
        .filter((r) => r.text && r.text.trim())
        .map(
          (r) =>
            new TextRun({
              text: r.text,
              bold: r.bold || false,
              italics: r.italic || false,
              underline: r.underline ? { type: UnderlineType.SINGLE } : undefined,
              superScript: r.superScript || false,
              subScript: r.subScript || false,
              font: fontFamily,
              size: fontSizeHalf,
              color: '000000',
            })
        );

      if (children.length === 0) continue;

      // Detect if this is a short label/metadata line — no indent, left-aligned
      const isLabel = looksLikeLabel(block.text);

      const targetArray = isMainBody ? mainBodyElements : frontMatterElements;

      targetArray.push(
        new Paragraph({
          alignment: isLabel ? AlignmentType.LEFT : AlignmentType.JUSTIFIED,
          spacing: {
            after: 0,
            line: lineSpacingValue,
            lineRule: LineRuleType.AUTO,
          },
          indent: isLabel
            ? { left: 0, firstLine: 0 }
            : { firstLine: cmToTwip(1.27), left: 0 },
          children,
        })
      );
      report.formattingFixed++;
    } else if (block.type === 'list-item') {
      const prefix = block.ordered ? `${block.index + 1}. ` : '• ';
      let fixedRuns = smartFixBold(block.runs, false);

      const children = [
        new TextRun({
          text: prefix,
          font: fontFamily,
          size: fontSizeHalf,
          color: '000000',
        }),
        ...fixedRuns
          .filter((r) => r.text && r.text.trim())
          .map(
            (r) =>
              new TextRun({
                text: r.text,
                bold: r.bold || false,
                italics: r.italic || false,
                underline: r.underline ? { type: UnderlineType.SINGLE } : undefined,
                font: fontFamily,
                size: fontSizeHalf,
                color: '000000',
              })
          ),
      ];

      const targetArray = isMainBody ? mainBodyElements : frontMatterElements;

      targetArray.push(
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: {
            after: 0,
            line: lineSpacingValue,
            lineRule: LineRuleType.AUTO,
          },
          indent: {
            left: cmToTwip(1.27),
            hanging: cmToTwip(0.63),
          },
          children,
        })
      );
    } else if (block.type === 'table') {
      const maxCols = Math.max(...block.rows.map((r) => r.length));

      // Pad all rows to maxCols
      for (const row of block.rows) {
        while (row.length < maxCols) {
          row.push({
            text: '',
            runs: [{ text: '', bold: false, italic: false, underline: false }],
            isHeader: false,
          });
        }
      }

      // ---- COLUMN ANALYSIS ----
      // Analyze each column to determine its type and ideal width
      const headerRow = block.rows[0] || [];
      const colInfo = [];

      for (let ci = 0; ci < maxCols; ci++) {
        const headerText = (headerRow[ci] && headerRow[ci].text || '').trim().toLowerCase();
        const allCellTexts = block.rows.map((r) => (r[ci] && r[ci].text || '').trim());
        const dataCellTexts = allCellTexts.slice(1); // exclude header

        // Detect column type
        const isNoColumn = /^(no\.?|nomor|#)$/i.test(headerText);
        const isAllNumbers = dataCellTexts.every((t) => !t || /^\d+[\.,]?\d*$/.test(t));
        const isShortColumn = allCellTexts.every((t) => t.length <= 8);
        const isYearColumn = /^(tahun|waktu|wakt|periode|thn)$/i.test(headerText) ||
          dataCellTexts.every((t) => !t || /^\d{4}(\/\d{2,4})?$/.test(t));

        // Max content length across all rows (for width calc)
        const maxLen = Math.max(...allCellTexts.map((t) => t.length), 1);
        // Average content length (for better distribution)
        const avgLen = allCellTexts.reduce((s, t) => s + t.length, 0) / Math.max(allCellTexts.length, 1);

        // Determine alignment for data rows
        let dataAlignment = AlignmentType.LEFT; // default
        if (isNoColumn || isAllNumbers || isYearColumn) {
          dataAlignment = AlignmentType.CENTER;
        } else if (isShortColumn) {
          dataAlignment = AlignmentType.CENTER;
        }

        // Determine base width weight
        let widthWeight;
        if (isNoColumn) {
          widthWeight = 3; // very narrow
        } else if (isYearColumn || (isAllNumbers && maxLen <= 10)) {
          widthWeight = 5; // narrow
        } else if (isShortColumn) {
          widthWeight = 6; // fairly narrow
        } else {
          // Use average length for proportional sizing, with a sqrt curve
          // to prevent very long columns from dominating
          widthWeight = Math.max(5, Math.round(Math.sqrt(avgLen) * 3));
        }

        colInfo.push({
          isNoColumn,
          isAllNumbers,
          isShortColumn,
          isYearColumn,
          dataAlignment,
          widthWeight,
          maxLen,
        });
      }

      // Calculate final column widths from weights
      const totalWeight = colInfo.reduce((s, c) => s + c.widthWeight, 0);
      const colWidths = colInfo.map((c) => {
        const pct = Math.round((c.widthWeight / totalWeight) * 100);
        // Enforce min 4%, no max cap (let proportions work)
        return Math.max(4, pct);
      });

      // Normalize to exactly 100%
      const widthSum = colWidths.reduce((s, w) => s + w, 0);
      if (widthSum !== 100) {
        // Distribute remainder to the widest column
        const maxIdx = colWidths.indexOf(Math.max(...colWidths));
        colWidths[maxIdx] += 100 - widthSum;
      }

      // ---- BUILD TABLE ROWS ----
      const tableRows = block.rows.map((row, rowIndex) => {
        return new TableRow({
          tableHeader: rowIndex === 0,
          children: row.map(
            (cell, colIndex) => {
              const info = colInfo[colIndex] || {};
              // Header row always centered; data rows use detected alignment
              const alignment = rowIndex === 0
                ? AlignmentType.CENTER
                : (info.dataAlignment || AlignmentType.LEFT);

              return new TableCell({
                borders: {
                  top: tableBorderStyle,
                  bottom: tableBorderStyle,
                  left: tableBorderStyle,
                  right: tableBorderStyle,
                },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment,
                    spacing: {
                      before: 40,
                      after: 40,
                      line: 276,
                      lineRule: LineRuleType.AUTO,
                    },
                    indent: { left: 0, firstLine: 0 },
                    children: [
                      new TextRun({
                        text: cell.text || '',
                        bold: cell.isHeader || rowIndex === 0,
                        font: fontFamily,
                        size: ptToHalfPoint(10), // Font size 10 for tables
                        color: '000000',
                      }),
                    ],
                  }),
                ],
                width: {
                  size: colWidths[colIndex] || Math.floor(100 / maxCols),
                  type: WidthType.PERCENTAGE,
                },
              });
            }
          ),
        });
      });

      const targetArray = isMainBody ? mainBodyElements : frontMatterElements;

      targetArray.push(
        new Table({
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
          rows: tableRows,
        })
      );

      targetArray.push(
        new Paragraph({
          spacing: { before: 120, after: 120 },
          children: [],
        })
      );
    }
  }

  // Create document
  const doc = new Document({
    features: {
      updateFields: true,
    },
    styles: {
      default: {
        document: {
          run: {
            font: fontFamily,
            size: fontSizeHalf,
            color: '000000',
          },
          paragraph: {
            spacing: {
              line: lineSpacingValue,
              lineRule: LineRuleType.AUTO,
              after: 0,
            },
          },
        },
        heading1: {
          run: {
            font: fontFamily,
            size: fontSizeHalf,
            bold: true,
            color: '000000',
          },
          paragraph: {
            spacing: { before: 360, after: 240, line: lineSpacingValue, lineRule: LineRuleType.AUTO },
            alignment: AlignmentType.LEFT,
          },
        },
        heading2: {
          run: {
            font: fontFamily,
            size: fontSizeHalf,
            bold: true,
            color: '000000',
          },
          paragraph: {
            spacing: { before: 240, after: 120, line: lineSpacingValue, lineRule: LineRuleType.AUTO },
            alignment: AlignmentType.LEFT,
          },
        },
        heading3: {
          run: {
            font: fontFamily,
            size: fontSizeHalf,
            bold: true,
            color: '000000',
          },
          paragraph: {
            spacing: { before: 200, after: 80, line: lineSpacingValue, lineRule: LineRuleType.AUTO },
            alignment: AlignmentType.LEFT,
          },
        },
      },
    },
    sections: [
      // ── SECTION 1: Front Matter (Daftar Isi, roman ii, iii…) ──
      {
        properties: {
          type: SectionType.NEXT_PAGE,
          page: {
            pageNumbers: {
              start: 2,
              formatType: NumberFormat.LOWER_ROMAN,
            },
            margin: {
              top: cmToTwip(marginTop),
              bottom: cmToTwip(marginBottom),
              left: cmToTwip(marginLeft),
              right: cmToTwip(marginRight),
            },
            size: {
              width: cmToTwip(21),
              height: cmToTwip(29.7),
            },
          },
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    children: [PageNumber.CURRENT],
                    font: fontFamily,
                    size: ptToHalfPoint(10),
                    color: '000000',
                  }),
                ],
              }),
            ],
          }),
        },
        children: frontMatterElements,
      },
      // ── SECTION 2: Body (halaman 1, 2, 3…) — restart ke angka arab ──
      {
        properties: {
          type: SectionType.NEXT_PAGE,
          titlePage: true,           // ← wajib agar halaman PERTAMA juga pakai footer
          page: {
            pageNumbers: {
              start: 1,
              formatType: NumberFormat.DECIMAL,
            },
            margin: {
              top: cmToTwip(marginTop),
              bottom: cmToTwip(marginBottom),
              left: cmToTwip(marginLeft),
              right: cmToTwip(marginRight),
            },
            size: {
              width: cmToTwip(21),
              height: cmToTwip(29.7),
            },
          },
        },
        footers: {
          // Footer untuk halaman 2, 3, 4, dst.
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    children: [PageNumber.CURRENT],
                    font: fontFamily,
                    size: ptToHalfPoint(10),
                    color: '000000',
                  }),
                ],
              }),
            ],
          }),
          // Footer khusus halaman PERTAMA (halaman 1) — sama isinya
          first: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    children: [PageNumber.CURRENT],
                    font: fontFamily,
                    size: ptToHalfPoint(10),
                    color: '000000',
                  }),
                ],
              }),
            ],
          }),
        },
        children: mainBodyElements.length > 0 ? mainBodyElements : [new Paragraph("")],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);

  // ── Gabungkan dengan Cover lalu inject ulang footer ──
  let finalBuffer = buffer;
  const coverFileName = coverType === 'lama' ? 'COVER LAMA.docx' : 'COVER 2026.docx';
  const coverPath = path.join(__dirname, '..', 'public', coverFileName);

  if (fs.existsSync(coverPath)) {
    try {
      const coverBuffer = fs.readFileSync(coverPath);

      // 1. Merge: cover dulu, lalu dokumen kita
      const merged = await new Promise((resolve, reject) => {
        const merger = new DocxMerger({}, [coverBuffer, buffer]);
        merger.save('nodebuffer', (data) => resolve(data));
      });

      // 2. Ekstrak file dari buffer-buffer kita
      const ourZip = new PizZip(buffer);
      const coverZip = new PizZip(coverBuffer);
      const mergedZip = new PizZip(merged);

      const footerFiles = Object.keys(ourZip.files).filter(
        (f) => f.startsWith('word/footer') && f.endsWith('.xml')
      );

      // Salin semua footer file dari buffer kita ke merged zip
      for (const fName of footerFiles) {
        mergedZip.file(fName, ourZip.file(fName).asText());
      }

      // Ambil XML
      let relsXml = mergedZip.file('word/_rels/document.xml.rels').asText();
      let docXml  = mergedZip.file('word/document.xml').asText();
      let ctXml   = mergedZip.file('[Content_Types].xml').asText();
      const ourDocXml = ourZip.file('word/document.xml').asText();
      const coverDocXml = coverZip.file('word/document.xml').asText();

      // Tambahkan footer rels dengan Custom ID agar tidak bentrok dengan gambar Cover
      const ourRelsXml = ourZip.file('word/_rels/document.xml.rels').asText();
      const footerRels = [...ourRelsXml.matchAll(/<Relationship[^>]+Id="([^"]+)"[^>]+Type="[^"]*\/footer"[^>]*Target="([^"]+)"[^>]*\/>/g)];
      
      const footerIdMap = {};
      for (let i = 0; i < footerRels.length; i++) {
        const oldId = footerRels[i][1];
        const target = footerRels[i][2];
        const newId = `myCustomFooterId_${i}`;
        footerIdMap[oldId] = newId;

        const newRel = `<Relationship Id="${newId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="${target}"/>`;
        relsXml = relsXml.replace('</Relationships>', `${newRel}</Relationships>`);
      }
      mergedZip.file('word/_rels/document.xml.rels', relsXml);

      // Tambahkan content type untuk footer
      for (const fName of footerFiles) {
        const partName = '/' + fName;
        if (!ctXml.includes(partName)) {
          ctXml = ctXml.replace('</Types>', `<Override PartName="${partName}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`);
        }
      }
      mergedZip.file('[Content_Types].xml', ctXml);

      // ==========================================
      // FIX SECTION BREAKS DIRUSAK OLEH DOCX-MERGER
      // ==========================================
      
      // Ambil sectPr Cover
      const coverSectPrMatch = coverDocXml.match(/<w:sectPr[^>]*>.*?<\/w:sectPr>/);
      let coverSectPr = coverSectPrMatch ? coverSectPrMatch[0] : '<w:sectPr><w:type w:val="nextPage"/></w:sectPr>';
      coverSectPr = coverSectPr.replace(/<w:(footer|header)Reference[^/]*\/>/g, ''); // no footer for cover
      if (!coverSectPr.includes('<w:type')) coverSectPr = coverSectPr.replace('</w:sectPr>', '<w:type w:val="nextPage"/></w:sectPr>');

      // Ambil semua sectPr dari dokumen asli kita
      const ourSectPrs = ourDocXml.match(/<w:sectPr[^>]*>.*?<\/w:sectPr>/g) || [];
      
      // Update r:id footer di sectPr kita agar pakai Custom ID
      for (let i = 0; i < ourSectPrs.length; i++) {
        for (const [oldId, newId] of Object.entries(footerIdMap)) {
          ourSectPrs[i] = ourSectPrs[i].replace(new RegExp(`r:id="${oldId}"`, 'g'), `r:id="${newId}"`);
        }
      }

      const ourInnerSectPr = ourSectPrs[0] || '';
      const ourDocLevelSectPr = ourSectPrs[ourSectPrs.length - 1] || '';

      // 1. Ganti page break docx-merger dengan proper section break cover
      const docxMergerPageBreak = /<w:p>\s*<w:r>\s*<w:br w:type="page"\/>\s*<\/w:r>\s*<\/w:p>/;
      docXml = docXml.replace(docxMergerPageBreak, `<w:p><w:pPr>${coverSectPr}</w:pPr></w:p>`);

      // 2. Ganti inner sectPr docx-merger dengan inner sectPr asli kita (jika ada)
      // Docx-merger mungkin mengubah inner sectPr, kita cari satu-satunya inner sectPr dan replace
      docXml = docXml.replace(/<w:sectPr(?![\s\S]*<w:sectPr)[\s\S]*?<\/w:sectPr>/, ourInnerSectPr);

      // 3. Ganti doc-level sectPr dengan doc-level sectPr asli kita
      if (ourDocLevelSectPr) {
        docXml = docXml.replace(/(<w:sectPr(?![\s\S]*<w:sectPr)[\s\S]*?)(\s*<\/w:body>)/, () => `${ourDocLevelSectPr}</w:body>`);
      }

      mergedZip.file('word/document.xml', docXml);
      finalBuffer = mergedZip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    } catch (e) {
      console.error('Cover merge/footer inject failed:', e.message);
      // Fallback: output tanpa cover
    }
  }

  fs.writeFileSync(outputPath, finalBuffer);

  return { report };
}

module.exports = { processWord };
