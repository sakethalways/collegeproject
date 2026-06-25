require("dotenv").config();
const express = require("express");
const axios = require("axios");
const cheerio = require("cheerio");
const cors = require("cors");
const { GoogleGenerativeAI } = require("@google/generative-ai");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, Table, TableRow, TableCell,
  WidthType, ShadingType,
} = require("docx");

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json());
app.use(express.static("public"));

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

// ─── Keywords to skip in both scraping output and Gemini findings ───
const SKIP_PATTERN = /copyright|©|\bfooter\b|contact us|phone|email|address|tel:|fax|linkedin|facebook|twitter|instagram|youtube|whatsapp|disclaimer|all rights reserved|privacy policy|terms of use|social media|navigation|menu item/i;

// ─── Strip markdown syntax from any string ───
function stripMarkdown(text) {
  if (!text) return "";
  return text
    .replace(/\*\*(.+?)\*\*/gs, "$1")   // **bold**
    .replace(/\*(.+?)\*/gs, "$1")        // *italic*
    .replace(/^#{1,6}\s+/gm, "")         // ## headings
    .replace(/`{1,3}([^`]*)`{1,3}/g, "$1") // `code`
    .replace(/_{1,2}(.+?)_{1,2}/gs, "$1")  // __underline__
    .replace(/~~(.+?)~~/gs, "$1")          // ~~strikethrough~~
    .replace(/^\s*[-*+]\s+/gm, "• ")       // - list → bullet
    .trim();
}

// ─── Post-process Gemini analysis: filter noise, clean markdown ───
function cleanAnalysis(analysis) {
  analysis.outdated = (analysis.outdated || []).filter(item =>
    !SKIP_PATTERN.test(item.section || "") &&
    !SKIP_PATTERN.test(item.originalText || "") &&
    !SKIP_PATTERN.test(item.problem || "")
  );

  analysis.missing = (analysis.missing || []).filter(item =>
    !SKIP_PATTERN.test(item.topic || "") &&
    !SKIP_PATTERN.test(item.whyItMatters || "")
  );

  analysis.accurate = (analysis.accurate || []).filter(item =>
    !SKIP_PATTERN.test(item.section || "") &&
    !SKIP_PATTERN.test(item.content || "")
  );

  analysis.updatedFullContent = stripMarkdown(analysis.updatedFullContent || "");
  analysis.summary = stripMarkdown(analysis.summary || "");
  analysis.overallStatus = stripMarkdown(analysis.overallStatus || "");

  return analysis;
}

// ─── Scrape ───
async function scrapeWebpage(url) {
  const response = await axios.get(url, {
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
    },
    timeout: 15000,
  });

  const $ = cheerio.load(response.data);
  $("script, style, noscript, nav, footer, header, iframe, img, svg, .footer, .nav, .menu, .copyright, .contact, #footer, #nav, #header").remove();

  const title = $("title").text().trim();
  const headings = [];
  $("h1, h2, h3, h4").each((_, el) => {
    const text = $(el).text().trim();
    if (text) headings.push({ tag: el.tagName.toUpperCase(), text });
  });

  const paragraphs = [];
  $("p, li, td, th").each((_, el) => {
    const text = $(el).text().trim().replace(/\s+/g, " ");
    if (text.length > 40 && !SKIP_PATTERN.test(text)) paragraphs.push(text);
  });

  const uniqueParagraphs = [...new Set(paragraphs)].slice(0, 50);

  const rawText = [
    `Title: ${title}`,
    "",
    "=== HEADINGS ===",
    ...headings.map(h => `${h.tag}: ${h.text}`),
    "",
    "=== CONTENT ===",
    ...uniqueParagraphs,
  ].join("\n");

  return { title, headings, content: rawText, url };
}

// ─── Extract JSON from Gemini response ───
function extractJSON(text) {
  // 1. Direct parse
  try { return JSON.parse(text); } catch (_) {}

  // 2. Strip markdown fences
  const fenceStripped = text.replace(/^```(?:json)?\s*/im, "").replace(/\s*```\s*$/im, "").trim();
  try { return JSON.parse(fenceStripped); } catch (_) {}

  // 3. Find the outermost balanced { ... } block (handles preamble text)
  let depth = 0, start = -1, end = -1;
  for (let i = 0; i < text.length; i++) {
    if (text[i] === "{") {
      if (depth === 0) start = i;
      depth++;
    } else if (text[i] === "}") {
      depth--;
      if (depth === 0 && start !== -1) { end = i; break; }
    }
  }
  if (start !== -1 && end !== -1) {
    try { return JSON.parse(text.slice(start, end + 1)); } catch (_) {}
  }

  throw new Error("Could not extract valid JSON from Gemini response");
}

// ─── Gemini analysis ───
async function analyzeWithGemini(scrapedData) {
  const model = genAI.getGenerativeModel({
    model: "gemini-2.5-flash",
    generationConfig: { responseMimeType: "application/json" },
  });

  const prompt = `You are an expert content auditor, fact-checker, and research analyst. Your mission is to deeply audit a webpage's substantive content using your current knowledge base.

STRICT EXCLUSIONS — do NOT analyse or mention any of the following, ever:
- Copyright notices, footer text, "All Rights Reserved", year in copyright
- Contact information: phone numbers, email addresses, postal addresses, fax
- Social media links or handles
- Navigation menus, breadcrumbs, sidebar links
- Legal disclaimers, privacy policies, terms of use
- Advertisements or promotional banners

FOCUS ONLY ON substantive content: academic programs, curriculum, course details, technologies, tools, frameworks, industry statistics, rankings, research, career outcomes, faculty credentials, lab facilities, certifications, and educational facts.

For the scraped content below:
1. FACT-CHECK every specific claim, date, statistic, version, ranking, or technology mentioned
2. FLAG anything outdated or incorrect — quote the exact phrase and give the exact correction
3. IDENTIFY important gaps — what substantive topics should this page cover but doesn't?
4. ADD new developments — recent tools, trends, or data the page is missing
5. CONFIRM what is still accurate and current

=== SCRAPED PAGE CONTENT ===
${scrapedData.content}
===========================

Return ONLY a JSON object (no markdown, no explanation):
{
  "topic": "concise topic of this page",
  "summary": "2-3 sentence overview of what this page covers",
  "estimatedPageAge": "your best estimate of when this content was last updated",
  "freshnessScore": <integer 0-100>,
  "outdated": [
    {
      "section": "heading or topic area",
      "originalText": "exact quote or close paraphrase from the page",
      "problem": "specific reason this is outdated or wrong",
      "correction": "the accurate, current information",
      "severity": "high|medium|low"
    }
  ],
  "missing": [
    {
      "topic": "name of the missing topic",
      "whyItMatters": "why this is important for this page",
      "newContent": "the actual up-to-date information to add",
      "priority": "high|medium|low"
    }
  ],
  "accurate": [
    {
      "section": "topic or section name",
      "content": "what the page says that is still correct"
    }
  ],
  "updatedFullContent": "A complete rewritten version of the page content with all corrections and additions. Write as clean professional prose. No markdown symbols like ** or ## — use plain text only."
}`;

  const result = await model.generateContent(prompt);
  const raw = result.response.text().trim();
  console.log("Gemini raw (first 500):\n", raw.slice(0, 500));

  // Gemini 2.5 sometimes emits thinking/preamble before the JSON object.
  // Find the outermost { ... } by matching braces, not just lastIndexOf.
  const parsed = extractJSON(raw);
  return parsed;
}

// ─── Build Word document ───
async function buildWordDoc(type, data) {
  const { title, url, originalContent, analysis } = data;
  const INTERNAL_NOTE = "Only for internal use of IARE";
  const children = [];

  const h1 = (text) => new Paragraph({
    text,
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 300, after: 120 },
  });

  const h2 = (text) => new Paragraph({
    text,
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
  });

  const body = (text, opts = {}) => new Paragraph({
    children: [new TextRun({ text: String(text), ...opts, size: 22, font: "Calibri" })],
    spacing: { after: 120 },
  });

  const boldLabel = (label, value) => new Paragraph({
    children: [
      new TextRun({ text: label + ": ", bold: true, size: 22, font: "Calibri" }),
      new TextRun({ text: String(value || ""), size: 22, font: "Calibri" }),
    ],
    spacing: { after: 100 },
  });

  const divider = () => new Paragraph({
    border: { bottom: { color: "CCCCCC", space: 1, style: BorderStyle.SINGLE, size: 6 } },
    spacing: { before: 200, after: 200 },
  });

  const footer = () => new Paragraph({
    children: [new TextRun({ text: INTERNAL_NOTE, italics: true, color: "888888", size: 18, font: "Calibri" })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 400 },
    border: { top: { color: "CCCCCC", space: 1, style: BorderStyle.SINGLE, size: 6 } },
  });

  if (type === "original") {
    children.push(h1("Original Page Content"));
    children.push(boldLabel("Source URL", url));
    children.push(boldLabel("Page Title", title));
    children.push(divider());
    // Split by lines and add as paragraphs
    const lines = (originalContent || "").split("\n");
    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed) { children.push(new Paragraph({ spacing: { after: 60 } })); continue; }
      if (trimmed.startsWith("=== ") && trimmed.endsWith(" ===")) {
        children.push(h2(trimmed.replace(/===/g, "").trim()));
      } else if (/^H[1-4]: /.test(trimmed)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: trimmed.replace(/^H[1-4]: /, ""), bold: true, size: 22, font: "Calibri" })],
          spacing: { after: 80 },
        }));
      } else {
        children.push(body(trimmed));
      }
    }
    children.push(footer());

  } else {
    // Updated / full audit report
    children.push(h1("AI Content Audit Report"));
    children.push(boldLabel("Page", title));
    children.push(boldLabel("URL", url));
    children.push(boldLabel("Topic", analysis.topic || ""));
    children.push(boldLabel("Estimated Content Age", analysis.estimatedPageAge || ""));
    children.push(boldLabel("Freshness Score", `${analysis.freshnessScore ?? "N/A"} / 100`));
    children.push(divider());

    children.push(h2("Summary"));
    children.push(body(analysis.summary || ""));
    children.push(divider());

    // Outdated
    const outdated = analysis.outdated || [];
    children.push(h2(`Outdated Information (${outdated.length} items)`));
    if (outdated.length === 0) {
      children.push(body("No outdated information found."));
    } else {
      for (let i = 0; i < outdated.length; i++) {
        const item = outdated[i];
        children.push(new Paragraph({
          children: [new TextRun({ text: `${i + 1}. ${item.section || ""}  [${(item.severity || "").toUpperCase()}]`, bold: true, size: 22, font: "Calibri" })],
          spacing: { before: 160, after: 60 },
        }));
        children.push(new Paragraph({
          children: [
            new TextRun({ text: "Original: ", bold: true, size: 21, font: "Calibri", color: "888888" }),
            new TextRun({ text: `"${item.originalText || ""}"`, italics: true, size: 21, font: "Calibri", color: "888888" }),
          ],
          spacing: { after: 60 },
        }));
        children.push(new Paragraph({
          children: [
            new TextRun({ text: "Issue: ", bold: true, size: 21, font: "Calibri", color: "CC0000" }),
            new TextRun({ text: item.problem || "", size: 21, font: "Calibri", color: "CC0000" }),
          ],
          spacing: { after: 60 },
        }));
        children.push(new Paragraph({
          children: [
            new TextRun({ text: "Correction: ", bold: true, size: 21, font: "Calibri", color: "1A7A40" }),
            new TextRun({ text: item.correction || "", size: 21, font: "Calibri", color: "1A7A40" }),
          ],
          spacing: { after: 120 },
        }));
      }
    }
    children.push(divider());

    // Missing
    const missing = analysis.missing || [];
    children.push(h2(`Missing Information (${missing.length} topics)`));
    if (missing.length === 0) {
      children.push(body("No missing information identified."));
    } else {
      for (let i = 0; i < missing.length; i++) {
        const item = missing[i];
        children.push(new Paragraph({
          children: [new TextRun({ text: `${i + 1}. ${item.topic || ""}  [${(item.priority || "").toUpperCase()} PRIORITY]`, bold: true, size: 22, font: "Calibri" })],
          spacing: { before: 160, after: 60 },
        }));
        children.push(new Paragraph({
          children: [
            new TextRun({ text: "Why it matters: ", bold: true, size: 21, font: "Calibri" }),
            new TextRun({ text: item.whyItMatters || "", size: 21, font: "Calibri" }),
          ],
          spacing: { after: 60 },
        }));
        children.push(new Paragraph({
          children: [
            new TextRun({ text: "Recommended addition: ", bold: true, size: 21, font: "Calibri", color: "1A4FA0" }),
            new TextRun({ text: item.newContent || "", size: 21, font: "Calibri", color: "1A4FA0" }),
          ],
          spacing: { after: 120 },
        }));
      }
    }
    children.push(divider());

    // Accurate
    const accurate = analysis.accurate || [];
    children.push(h2(`Still Accurate (${accurate.length} items)`));
    if (accurate.length === 0) {
      children.push(body("No confirmed accurate items."));
    } else {
      for (const item of accurate) {
        children.push(new Paragraph({
          children: [
            new TextRun({ text: `${item.section || ""}: `, bold: true, size: 21, font: "Calibri" }),
            new TextRun({ text: item.content || "", size: 21, font: "Calibri" }),
          ],
          spacing: { after: 80 },
        }));
      }
    }
    children.push(divider());

    // Updated content
    children.push(h2("AI-Updated Full Content"));
    const updatedLines = (analysis.updatedFullContent || "").split("\n");
    for (const line of updatedLines) {
      const trimmed = line.trim();
      if (!trimmed) { children.push(new Paragraph({ spacing: { after: 60 } })); continue; }
      children.push(body(trimmed));
    }

    children.push(footer());
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: "Calibri", size: 22 },
        },
      },
    },
    sections: [{ children }],
  });

  return Packer.toBuffer(doc);
}

// ─── Routes ───
app.post("/api/analyze", async (req, res) => {
  const { url } = req.body;
  if (!url) return res.status(400).json({ error: "URL is required" });

  try {
    console.log(`\nScraping: ${url}`);
    const scraped = await scrapeWebpage(url);
    console.log(`Scraped ${scraped.content.length} chars. Analyzing...`);

    const rawAnalysis = await analyzeWithGemini(scraped);
    const analysis = cleanAnalysis(rawAnalysis);

    res.json({
      success: true,
      url,
      title: scraped.title,
      originalContent: scraped.content,
      analysis,
    });
  } catch (err) {
    console.error("Error:", err.message);
    res.status(500).json({ error: err.message || "Failed to analyze the page" });
  }
});

app.post("/api/download", async (req, res) => {
  const { type, title, url, originalContent, analysis } = req.body;
  if (!type || !url) return res.status(400).json({ error: "Missing required fields" });

  try {
    const buffer = await buildWordDoc(type, { title, url, originalContent, analysis });
    const filename = type === "original" ? "original-content.docx" : "ai-updated-report.docx";
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (err) {
    console.error("Download error:", err.message);
    res.status(500).json({ error: "Failed to generate document: " + err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
