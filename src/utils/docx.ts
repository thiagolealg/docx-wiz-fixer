import mammoth from "mammoth";
// @ts-ignore - no types for this module


export interface ParsedDocx {
  html: string;
  paragraphs: string[];
}

export interface NumberingIssue {
  index: number; // paragraph index
  found: number;
  expected: number;
  text: string;
}

export interface CompareResult {
  missingInTarget: string[]; // paragraphs present in reference but missing in target
  extraInTarget: string[]; // paragraphs present in target but not in reference
}

export interface ParagraphSearchResult {
  index: number;
  paragraph: string;
}

export async function fileToArrayBuffer(file: File): Promise<ArrayBuffer> {
  return await file.arrayBuffer();
}

export async function parseDocxFile(file: File): Promise<ParsedDocx> {
  const arrayBuffer = await fileToArrayBuffer(file);
  const { value: html } = await mammoth.convertToHtml({ arrayBuffer }, {
    styleMap: [
      "p[style-name='Normal'] => p:fresh",
      "h1 => p:fresh",
      "h2 => p:fresh",
      "h3 => p:fresh",
      "h4 => p:fresh",
      "h5 => p:fresh",
      "h6 => p:fresh",
      "list:unordered => ul",
      "list:ordered => ol",
    ],
    includeDefaultStyleMap: true,
  });

  const paragraphs = extractParagraphsFromHtml(html);
  return { html, paragraphs };
}

export function extractParagraphsFromHtml(html: string): string[] {
  const container = document.createElement("div");
  container.innerHTML = html;
  const paras: string[] = [];
  container.querySelectorAll("p, li").forEach((el) => {
    const text = (el.textContent || "").replace(/\s+/g, " ").trim();
    if (text) paras.push(text);
  });
  return paras;
}

export function normalizeParagraphs(paragraphs: string[]): string[] {
  // Trim, collapse inner spaces, and ensure a single blank line between paragraphs when exporting
  return paragraphs
    .map((p) => p.replace(/\s+/g, " ").trim())
    .filter((p) => p.length > 0);
}

export function buildHtmlFromParagraphs(paragraphs: string[]): string {
  const body = paragraphs
    .map((p) => `<p style="margin:0 0 1em 0; line-height:1.5;">${escapeHtml(p)}</p>`)
    .join("\n");

  // One empty paragraph (visual enter) after each paragraph is achieved via margin plus an extra break block at the end
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Documento Normalizado</title>
  <meta name="generator" content="Lovable DOCX Tool" />
  <style>
    body { font-family: Arial, sans-serif; }
  </style>
</head>
<body>
<article>
${body}
</article>
</body>
</html>`;
}

export function findParagraphByItemNumber(paragraphs: string[], itemNumber: string): ParagraphSearchResult {
  // Normalize the search term by removing extra spaces and converting to lowercase
  const searchTerm = itemNumber.trim().toLowerCase();
  
  for (let i = 0; i < paragraphs.length; i++) {
    const paragraph = paragraphs[i];
    
    // Check if paragraph starts with the item number followed by space, period, or )
    const regex = new RegExp(`^\\s*${escapeRegex(searchTerm)}[\\s\\.)]+`, 'i');
    if (regex.test(paragraph)) {
      return { index: i, paragraph };
    }
    
    // Also check for exact match at the beginning
    if (paragraph.toLowerCase().startsWith(searchTerm.toLowerCase())) {
      return { index: i, paragraph };
    }
  }
  
  return { index: -1, paragraph: "" };
}

function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function checkNumbering(paragraphs: string[]): NumberingIssue[] {
  const issues: NumberingIssue[] = [];
  let expected = 1;
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    const m = p.match(/^\s*(\d+)[\.)]\s+/);
    if (m) {
      const found = parseInt(m[1], 10);
      if (found !== expected) {
        issues.push({ index: i, found, expected, text: p });
        expected = found + 1; // resync after reporting
      } else {
        expected += 1;
      }
    }
  }
  return issues;
}

export function compareParagraphSets(reference: string[], target: string[]): CompareResult {
  const norm = (arr: string[]) => arr.map((p) => p.replace(/\s+/g, " ").trim().toLowerCase());
  const a = norm(reference);
  const b = norm(target);
  const setB = new Set(b);
  const setA = new Set(a);

  const missingInTarget: string[] = [];
  reference.forEach((p, idx) => {
    if (!setB.has(a[idx])) missingInTarget.push(p);
  });
  const extraInTarget: string[] = [];
  target.forEach((p, idx) => {
    if (!setA.has(b[idx])) extraInTarget.push(p);
  });
  return { missingInTarget, extraInTarget };
}

export async function htmlToDocxBlob(html: string): Promise<Blob> {
  const HtmlDocx = await loadHtmlDocx();
  return HtmlDocx.asBlob(html);
}

async function loadHtmlDocx(): Promise<any> {
  if (typeof window !== "undefined" && (window as any).HTMLDocx) {
    return (window as any).HTMLDocx;
  }
  await new Promise<void>((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/html-docx-js@0.3.1/dist/html-docx.min.js";
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("Failed to load html-docx-js"));
    document.head.appendChild(script);
  });
  return (window as any).HTMLDocx;
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
