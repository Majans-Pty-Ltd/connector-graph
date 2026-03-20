/**
 * Content extraction utilities for email attachments.
 * Converts binary attachment data into readable text based on content type.
 */

import { PDFParse } from "pdf-parse";
import * as mammoth from "mammoth";
import * as XLSX from "xlsx";

// ── Result types ──

export interface ExtractionResult {
  extractedText: string;
  format: "text" | "json" | "image" | "unsupported";
  metadata?: Record<string, unknown>;
}

// ── Content type detection ──

const TEXT_EXTENSIONS = new Set([
  ".txt", ".md", ".json", ".xml", ".log", ".yaml", ".yml", ".csv",
]);

const IMAGE_CONTENT_TYPES = new Set([
  "image/png", "image/jpeg", "image/jpg", "image/gif", "image/webp", "image/svg+xml",
]);

const IMAGE_EXTENSIONS = new Set([
  ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg",
]);

function getExtension(filename: string): string {
  const dot = filename.lastIndexOf(".");
  return dot >= 0 ? filename.slice(dot).toLowerCase() : "";
}

// ── HTML extraction ──

function extractHtml(html: string): string {
  let text = html;

  // Remove style and script blocks entirely
  text = text.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "");
  text = text.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "");

  // Convert structural tags to newlines
  text = text.replace(/<br\s*\/?>/gi, "\n");
  text = text.replace(/<\/?(p|div|h[1-6]|li|tr|blockquote|pre|hr)[^>]*>/gi, "\n");
  text = text.replace(/<\/?(ul|ol|table|thead|tbody|tfoot)[^>]*>/gi, "\n");
  text = text.replace(/<td[^>]*>/gi, "\t");

  // Strip all remaining HTML tags
  text = text.replace(/<[^>]+>/g, "");

  // Decode common HTML entities
  text = text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&apos;/g, "'")
    .replace(/&nbsp;/g, " ")
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(parseInt(code, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)));

  // Collapse whitespace: multiple blank lines -> single, trim lines
  text = text
    .split("\n")
    .map((line) => line.trim())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  return text;
}

// ── PDF extraction ──

async function extractPdf(buffer: Buffer): Promise<ExtractionResult> {
  const parser = new PDFParse({ data: new Uint8Array(buffer) });
  try {
    const textResult = await parser.getText();
    const info = await parser.getInfo().catch(() => null);
    return {
      extractedText: textResult.text,
      format: "text",
      metadata: {
        pages: textResult.total,
        ...(info?.info ? { title: info.info.Title, author: info.info.Author } : {}),
      },
    };
  } finally {
    await parser.destroy().catch(() => {});
  }
}

// ── Word .docx extraction ──

async function extractDocx(buffer: Buffer): Promise<ExtractionResult> {
  const result = await mammoth.extractRawText({ buffer });
  return {
    extractedText: result.value,
    format: "text",
  };
}

// ── Excel .xlsx extraction ──

function extractXlsx(buffer: Buffer): ExtractionResult {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const sheets: string[] = workbook.SheetNames;
  let totalRows = 0;
  const allData: Record<string, unknown[]> = {};

  for (const sheetName of sheets) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet);
    allData[sheetName] = rows;
    totalRows += rows.length;
  }

  return {
    extractedText: JSON.stringify(allData, null, 2),
    format: "json",
    metadata: {
      sheets,
      sheetCount: sheets.length,
      rowCount: totalRows,
    },
  };
}

// ── CSV extraction ──

function extractCsv(text: string): ExtractionResult {
  const lines = text.split(/\r?\n/).filter((line) => line.trim() !== "");
  if (lines.length === 0) {
    return { extractedText: "[]", format: "json", metadata: { rowCount: 0 } };
  }

  const headers = parseCsvLine(lines[0]);
  const rows: Record<string, string>[] = [];

  for (let i = 1; i < lines.length; i++) {
    const values = parseCsvLine(lines[i]);
    const row: Record<string, string> = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = values[j] ?? "";
    }
    rows.push(row);
  }

  return {
    extractedText: JSON.stringify(rows, null, 2),
    format: "json",
    metadata: {
      columns: headers,
      rowCount: rows.length,
    },
  };
}

function parseCsvLine(line: string): string[] {
  const result: string[] = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];

    if (inQuotes) {
      if (char === '"') {
        // Check for escaped quote (double quote)
        if (i + 1 < line.length && line[i + 1] === '"') {
          current += '"';
          i++; // skip next quote
        } else {
          inQuotes = false;
        }
      } else {
        current += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === ",") {
        result.push(current.trim());
        current = "";
      } else {
        current += char;
      }
    }
  }

  result.push(current.trim());
  return result;
}

// ── EML extraction ──

function extractEml(text: string): ExtractionResult {
  // Split headers from body (double newline separates them)
  const headerBodySplit = text.indexOf("\n\n");
  const headerSection = headerBodySplit >= 0 ? text.slice(0, headerBodySplit) : text;
  const bodySection = headerBodySplit >= 0 ? text.slice(headerBodySplit + 2) : "";

  // Unfold headers (continuation lines start with whitespace)
  const unfoldedHeaders = headerSection.replace(/\r?\n[ \t]+/g, " ");

  // Extract key headers
  const fromMatch = unfoldedHeaders.match(/^From:\s*(.+)$/mi);
  const toMatch = unfoldedHeaders.match(/^To:\s*(.+)$/mi);
  const subjectMatch = unfoldedHeaders.match(/^Subject:\s*(.+)$/mi);
  const dateMatch = unfoldedHeaders.match(/^Date:\s*(.+)$/mi);
  const ccMatch = unfoldedHeaders.match(/^Cc:\s*(.+)$/mi);
  const contentTypeMatch = unfoldedHeaders.match(/^Content-Type:\s*(.+)$/mi);

  const headers: Record<string, string> = {};
  if (fromMatch) headers["From"] = fromMatch[1].trim();
  if (toMatch) headers["To"] = toMatch[1].trim();
  if (subjectMatch) headers["Subject"] = subjectMatch[1].trim();
  if (dateMatch) headers["Date"] = dateMatch[1].trim();
  if (ccMatch) headers["Cc"] = ccMatch[1].trim();

  // Try to extract body text
  let body = bodySection;
  const contentType = contentTypeMatch?.[1] ?? "";

  // If it's multipart, try to find the text/plain part
  const boundaryMatch = contentType.match(/boundary="?([^";\s]+)"?/i);
  if (boundaryMatch) {
    const boundary = boundaryMatch[1];
    const parts = body.split(`--${boundary}`);
    // Look for text/plain part
    for (const part of parts) {
      if (part.match(/Content-Type:\s*text\/plain/i)) {
        const partBodyStart = part.indexOf("\n\n");
        if (partBodyStart >= 0) {
          body = part.slice(partBodyStart + 2).trim();
          break;
        }
      }
    }
    // If no text/plain found, try text/html
    if (body === bodySection) {
      for (const part of parts) {
        if (part.match(/Content-Type:\s*text\/html/i)) {
          const partBodyStart = part.indexOf("\n\n");
          if (partBodyStart >= 0) {
            body = extractHtml(part.slice(partBodyStart + 2));
            break;
          }
        }
      }
    }
  } else if (contentType.includes("text/html")) {
    body = extractHtml(body);
  }

  // Clean up body — remove trailing boundary markers
  body = body.replace(/--[^\n]+--\s*$/g, "").trim();

  const headerLines = Object.entries(headers)
    .map(([k, v]) => `${k}: ${v}`)
    .join("\n");

  return {
    extractedText: `${headerLines}\n\n---\n\n${body}`,
    format: "text",
    metadata: headers,
  };
}

// ── Image handling ──

function extractImage(contentBytes: string, contentType: string): ExtractionResult {
  const dataUrl = `data:${contentType};base64,${contentBytes}`;
  return {
    extractedText: dataUrl,
    format: "image",
  };
}

// ── Main extraction entry point ──

export async function extractContent(
  contentBytes: string,
  contentType: string,
  filename: string
): Promise<ExtractionResult> {
  const ext = getExtension(filename);
  const ct = contentType.toLowerCase();

  try {
    // HTML
    if (ct.includes("text/html") || ext === ".html" || ext === ".htm") {
      const text = Buffer.from(contentBytes, "base64").toString("utf-8");
      return { extractedText: extractHtml(text), format: "text" };
    }

    // PDF
    if (ct === "application/pdf" || ext === ".pdf") {
      const buffer = Buffer.from(contentBytes, "base64");
      return await extractPdf(buffer);
    }

    // Word .docx
    if (
      ct === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
      ext === ".docx"
    ) {
      const buffer = Buffer.from(contentBytes, "base64");
      return await extractDocx(buffer);
    }

    // Excel .xlsx / .xls
    if (
      ct === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      ct === "application/vnd.ms-excel" ||
      ext === ".xlsx" ||
      ext === ".xls"
    ) {
      const buffer = Buffer.from(contentBytes, "base64");
      return extractXlsx(buffer);
    }

    // CSV
    if (ct === "text/csv" || ct === "application/csv" || ext === ".csv") {
      const text = Buffer.from(contentBytes, "base64").toString("utf-8");
      return extractCsv(text);
    }

    // EML
    if (ct === "message/rfc822" || ext === ".eml") {
      const text = Buffer.from(contentBytes, "base64").toString("utf-8");
      return extractEml(text);
    }

    // Images
    if (IMAGE_CONTENT_TYPES.has(ct) || IMAGE_EXTENSIONS.has(ext)) {
      return extractImage(contentBytes, ct || `image/${ext.slice(1)}`);
    }

    // Plain text types (by content type or extension)
    if (
      ct.startsWith("text/") ||
      ct.includes("json") ||
      ct.includes("xml") ||
      ct.includes("yaml") ||
      TEXT_EXTENSIONS.has(ext)
    ) {
      const text = Buffer.from(contentBytes, "base64").toString("utf-8");
      return { extractedText: text, format: "text" };
    }

    // Unsupported
    return {
      extractedText: `Unsupported content type: ${contentType} (${filename}). Use graph_get_attachment to get the raw base64 content.`,
      format: "unsupported",
    };
  } catch (err) {
    return {
      extractedText: `Extraction failed for ${filename} (${contentType}): ${(err as Error).message}. Use graph_get_attachment to get the raw content.`,
      format: "unsupported",
    };
  }
}
