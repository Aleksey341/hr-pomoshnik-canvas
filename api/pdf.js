import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import fontkit from "@pdf-lib/fontkit";
import { PDFDocument, rgb } from "pdf-lib";
import { plain } from "./pptx.js";

const FONTS = path.join(path.dirname(fileURLToPath(import.meta.url)), "fonts");
const CHECKS = [
  ["c:it", "ИТ"],
  ["c:b2b", "B2B"],
  ["c:wink", "B2C Wink"],
  ["c:core", "B2C Core"],
  ["c:b20", "B20"],
  ["c:solar", "Солар"],
  ["c:turbo", "Турбо"],
];

function teamRows(data) {
  return [1, 2, 3, 4].map(function(index) {
    const cols = ["n", "r", "e", "w"].map(function(part) {
      return plain(data["t" + index + part]).join(" ");
    });
    if (cols.every(function(col) { return !col; })) return "";
    return cols.join("  ·  ");
  }).filter(Boolean);
}

export function sections(data) {
  const checks = CHECKS.map(function(item) {
    const mark = data[item[0]] === "1" ? "[x]" : "[ ]";
    return mark + "  " + item[1];
  }).join("    ");
  return [
    ["Формула", plain(data.formula)],
    ["1. Сегменты потребителей", [
      "1.1 Конечные пользователи",
      ...plain(data.users),
      "1.2 К какому сегменту бизнеса внутри Ростелеком относится ваша идея:",
      ...plain(data["segment-note"]),
      checks,
    ]],
    ["2. Проблема", plain(data.problem)],
    ["3. Решение", [
      "3.1 Краткое описание",
      ...plain(data.solution),
      "3.2 Предполагаемые конкурентные преимущества",
      ...plain(data.advantage),
    ]],
    ["4. Объем рынка", plain(data.market)],
    ["5. Как продукт будет зарабатывать?", [
      ...plain(data.money),
      "5.1 Пилот",
      ...plain(data.pilot),
      "5.2 Годовая лицензия",
      ...plain(data.check),
    ]],
    ["6. Команда", [
      ...teamRows(data),
      ...plain(data["team-notes"]),
    ]],
  ];
}

function wrap(text, font, size, width) {
  const words = String(text).split(/\s+/).filter(Boolean);
  const lines = [];
  let line = "";
  words.forEach(function(word) {
    const next = line ? line + " " + word : word;
    if (font.widthOfTextAtSize(next, size) > width && line) {
      lines.push(line);
      line = word;
    } else {
      line = next;
    }
  });
  if (line) lines.push(line);
  return lines.length ? lines : [""];
}

export async function buildPdf(data) {
  const pdf = await PDFDocument.create();
  pdf.registerFontkit(fontkit);
  const regular = await pdf.embedFont(await readFile(path.join(FONTS, "Montserrat-Regular.ttf")), { subset: true });
  const bold = await pdf.embedFont(await readFile(path.join(FONTS, "Montserrat-Bold.ttf")), { subset: true });
  const pageWidth = 595.28;
  const pageHeight = 841.89;
  const margin = 42;
  const width = pageWidth - margin * 2;
  let page = pdf.addPage([pageWidth, pageHeight]);
  let y = pageHeight - margin;

  function nextPage() {
    page = pdf.addPage([pageWidth, pageHeight]);
    y = pageHeight - margin;
  }

  function drawLine(text, font, size, color) {
    const height = size + 5;
    if (y - height < margin) nextPage();
    page.drawText(text, { x: margin, y: y - size, size: size, font: font, color: color });
    y -= height;
  }

  drawLine("Хакатон Canvas. Название проекта: HR Помощник", bold, 14, rgb(0.08, 0.08, 0.17));
  y -= 6;
  sections(data || {}).forEach(function(section) {
    y -= 8;
    drawLine(section[0], bold, 12, rgb(0.47, 0, 1));
    section[1].forEach(function(paragraph) {
      wrap(paragraph, regular, 10, width).forEach(function(line) {
        drawLine(line, regular, 10, rgb(0.08, 0.08, 0.17));
      });
      y -= 2;
    });
  });
  return Buffer.from(await pdf.save());
}

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      res.setHeader("Allow", "POST");
      res.status(405).end();
      return;
    }
    const body = req.body;
    if (!body || typeof body !== "object" || Array.isArray(body)) {
      res.status(400).json({ error: "Нужен канвас" });
      return;
    }
    const file = await buildPdf(body);
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", "attachment; filename*=UTF-8''HR%20Pomoshnik%20Canvas.pdf");
    res.setHeader("Cache-Control", "no-store");
    res.status(200).send(file);
  } catch (error) {
    res.status(500).json({ error: "Не удалось собрать PDF" });
  }
}
