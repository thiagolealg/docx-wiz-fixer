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

  // Helper to normalize whitespace
  const norm = (s: string) => s.replace(/\s+/g, " ").trim();

  // Extract text of an LI without its nested lists
  const liMainText = (li: Element) => {
    const clone = li.cloneNode(true) as Element;
    clone.querySelectorAll("ol, ul").forEach((n) => n.remove());
    return norm(clone.textContent || "");
  };

  // Process nested ordered lists to generate literal numbering like 6.1.2.1
  const processOl = (ol: Element, stack: number[]) => {
    const startAttr = (ol as HTMLOListElement).getAttribute("start");
    let start = startAttr ? parseInt(startAttr, 10) : 1;
    let counter = start - 1;

    Array.from(ol.children).forEach((li) => {
      if (li.tagName.toLowerCase() !== "li") return;
      const valueAttr = (li as HTMLLIElement).getAttribute("value");
      const current = valueAttr ? parseInt(valueAttr, 10) : (counter + 1);
      counter = current;
      const numbering = [...stack, current].join(".");

      const text = liMainText(li);
      if (text) paras.push(`${numbering} ${text}`);

      // Handle nested lists for deeper levels (e.g., 6.1.2.1)
      (li as HTMLLIElement).querySelectorAll(":scope > ol").forEach((childOl) => {
        processOl(childOl, [...stack, current]);
      });

      // Also include unordered sublists as plain paragraphs (no numbering)
      (li as HTMLLIElement).querySelectorAll(":scope > ul > li").forEach((subLi) => {
        const bulletText = norm((subLi as HTMLLIElement).textContent || "");
        if (bulletText) paras.push(bulletText);
      });
    });
  };

  const walk = (root: Element) => {
    Array.from(root.children).forEach((el) => {
      const tag = el.tagName.toLowerCase();
      if (tag === "p") {
        const text = norm(el.textContent || "");
        if (text) paras.push(text);
        return;
      }
      if (tag === "ol") {
        processOl(el, []);
        return; // processOl already handles its nested lists
      }
      if (tag === "ul") {
        // Unordered lists: keep items as plain paragraphs
        el.querySelectorAll(":scope > li").forEach((li) => {
          const text = liMainText(li);
          if (text) paras.push(text);
          // process nested ordered lists inside this li
          (li as HTMLLIElement).querySelectorAll(":scope > ol").forEach((childOl) => {
            processOl(childOl, []);
          });
          // recurse into nested unordered lists
          (li as HTMLLIElement).querySelectorAll(":scope > ul").forEach((childUl) => {
            walk(childUl);
          });
        });
        return;
      }
      // other containers: continue walking
      walk(el);
    });
  };

  walk(container);

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
  const target = itemNumber.trim();
  for (let i = 0; i < paragraphs.length; i++) {
    const paragraph = paragraphs[i];
    if (
      paragraph.startsWith(target + " ") ||
      paragraph.startsWith(target + ".") ||
      paragraph.startsWith(target + ")") ||
      paragraph === target
    ) {
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

export function renumberHierarchical(paragraphs: string[]): string[] {
  const out: string[] = [];
  let current: number[] = [];
  const re = /^\s*(\d+(?:\.\d+)*)([.)])?\s+(.*)$/;
  for (let i = 0; i < paragraphs.length; i++) {
    const original = paragraphs[i];
    const m = original.match(re);
    if (!m) {
      out.push(original);
      continue;
    }
    const [, numberingStr, delim = "", restRaw] = m;
    const rest = restRaw.replace(/\s+/g, " ").trim();
    const found = numberingStr.split(".").map((n) => parseInt(n, 10));
    let expected: number[] = [];

    if (current.length === 0) {
      // First numbered item: respect existing numbering as the starting point
      expected = found.slice();
    } else {
      const dFound = found.length;
      const dCurr = current.length;
      if (dFound === dCurr) {
        // same level: increment last
        expected = current.slice(0, dCurr - 1).concat([current[dCurr - 1] + 1]);
      } else if (dFound === dCurr + 1) {
        // one level deeper: start at 1
        expected = current.concat([1]);
      } else if (dFound < dCurr) {
        // going up levels: increment at the new last level
        expected = current.slice(0, dFound - 1).concat([current[dFound - 1] + 1]);
      } else {
        // jumped multiple levels deeper: fill missing levels with 1
        const toAdd = dFound - dCurr;
        expected = current.slice();
        for (let k = 0; k < toAdd; k++) expected.push(1);
      }
    }

    const expectedStr = expected.join(".");
    const replaced = `${expectedStr}${delim ? delim : ""} ${rest}`;
    out.push(replaced);
    current = expected;
  }
  return out;
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
