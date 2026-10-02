import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const TEMPLATE = path.join(path.dirname(fileURLToPath(import.meta.url)), "canvas-template.pptx");
const CHECKS = [
  ["c:it", "ИТ"],
  ["c:b2b", "B2B"],
  ["c:wink", "B2C Wink"],
  ["c:core", "B2C Core"],
  ["c:b20", "B20"],
  ["c:solar", "Солар"],
  ["c:turbo", "Турбо"],
];

function decode(value) {
  return value
    .replace(/&nbsp;/gi, " ")
    .replace(/&#160;/g, " ")
    .replace(/&#(\d+);/g, function(_, n) { return String.fromCharCode(Number(n)); })
    .replace(/&quot;/g, "\"")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

export function plain(value) {
  if (!value) return [];
  const html = String(value)
    .replace(/<li\b[^>]*>/gi, "\n• ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|li|tr|h1|h2|h3)>/gi, "\n")
    .replace(/<[^>]+>/g, "");
  return decode(html)
    .split("\n")
    .map(function(line) { return line.replace(/\s+/g, " ").trim(); })
    .filter(Boolean);
}

function esc(value) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function paragraph(text, sizePt, bold, linePt) {
  const size = Math.round(sizePt * 100);
  const line = Math.round(linePt * 100);
  const weight = bold ? " b=\"1\"" : "";
  return "<a:p><a:pPr><a:spcBef><a:spcPts val=\"0\"/></a:spcBef><a:spcAft><a:spcPts val=\"0\"/></a:spcAft><a:lnSpc><a:spcPts val=\"" + line + "\"/></a:lnSpc></a:pPr><a:r><a:rPr lang=\"ru-RU\" sz=\"" + size + "\"" + weight + " dirty=\"0\"><a:solidFill><a:srgbClr val=\"14142B\"/></a:solidFill><a:latin typeface=\"Montserrat\"/><a:ea typeface=\"Montserrat\"/><a:cs typeface=\"Montserrat\"/></a:rPr><a:t>" + esc(text) + "</a:t></a:r></a:p>";
}

function txBody(blocks, linePt) {
  const body = blocks.map(function(block) {
    return paragraph(block[0], block[1], block[2], linePt);
  }).join("");
  return "<p:txBody><a:bodyPr wrap=\"square\" lIns=\"40000\" tIns=\"20000\" rIns=\"40000\" bIns=\"20000\" rtlCol=\"0\" anchor=\"t\"><a:spAutoFit/></a:bodyPr><a:lstStyle/>" + body + "</p:txBody>";
}

function shapeBounds(xml, id) {
  const at = xml.indexOf("id=\"" + id + "\"");
  if (at < 0) throw new Error("Нет фигуры " + id);
  const start = xml.lastIndexOf("<p:sp>", at);
  const end = xml.indexOf("</p:sp>", at);
  if (start < 0 || end < 0) throw new Error("Нет границ фигуры " + id);
  return { start: start, end: end + "</p:sp>".length };
}

function fillShape(xml, id, blocks, linePt, height) {
  const bounds = shapeBounds(xml, id);
  let shape = xml.slice(bounds.start, bounds.end);
  if (height) shape = shape.replace(/(<a:ext cx="\d+" cy=")\d+(")/, "$1" + height + "$2");
  if (!/<p:txBody>[\s\S]*?<\/p:txBody>/.test(shape)) throw new Error("Нет текста у фигуры " + id);
  shape = shape.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, txBody(blocks, linePt));
  return xml.slice(0, bounds.start) + shape + xml.slice(bounds.end);
}

function lines(value, size, bold) {
  return plain(value).map(function(line) { return [line, size, bold]; });
}

function teamLines(data) {
  const blocks = [
    ["6. Команда. Кто сейчас есть в команде?", 10, true],
  ];
  [1, 2, 3, 4].forEach(function(index) {
    const cols = ["n", "r", "e", "w"].map(function(part) {
      return plain(data["t" + index + part]).join(" ");
    });
    if (cols.every(function(col) { return !col; })) return;
    blocks.push([cols.join("  ·  "), 8, false]);
  });
  lines(data["team-notes"], 7, false).forEach(function(line) { blocks.push(line); });
  return blocks;
}

export function slideXml(xml, data) {
  const checks = CHECKS.map(function(item) {
    const mark = data[item[0]] === "1" ? "[x]" : "[ ]";
    return [mark + "  " + item[1], 8, false];
  });
  const jobs = [
    ["163", [["Хакатон Canvas. Название проекта: HR Помощник", 14, true]], 16],
    ["154", [
      ...lines(data.formula, 7, true),
      ["1. Сегменты потребителей", 10, true],
      ["1.1 Конечные пользователи", 8, true],
      ...lines(data.users, 7, false),
      ["1.2 К какому сегменту бизнеса внутри Ростелеком относится ваша идея:", 8, true],
      ...lines(data["segment-note"], 7, false),
      ...checks,
    ], 8],
    ["155", [
      ["2. Проблема", 10, true],
      ...lines(data.problem, 7, false),
    ], 8, 3986575],
    ["157", [
      ["3. Решение", 10, true],
      ["3.1 Краткое описание", 8, true],
      ...lines(data.solution, 6.5, false),
      ["3.2 Предполагаемые конкурентные преимущества", 8, true],
      ...lines(data.advantage, 6.5, false),
    ], 6.5, 3986575],
    ["158", [
      ["4. Объем рынка", 10, true],
      ...lines(data.market, 7, false),
    ], 8],
    ["160", [
      ["5. Как продукт будет зарабатывать?", 10, true],
      ...lines(data.money, 7, false),
      ["5.1 Пилот", 8, true],
      ...lines(data.pilot, 7, false),
      ["5.2 Годовая лицензия", 8, true],
      ...lines(data.check, 7, false),
    ], 8, 3964600],
    ["159", teamLines(data), 8],
  ];
  const filled = jobs.map(function(job) {
    return { id: job[0], blocks: job[1], line: job[2], height: job[3], ...shapeBounds(xml, job[0]) };
  }).sort(function(a, b) { return b.start - a.start; });
  filled.forEach(function(job) {
    xml = fillShape(xml, job.id, job.blocks, job.line, job.height);
  });
  return xml;
}

export async function buildPptx(data) {
  const zip = await JSZip.loadAsync(await readFile(TEMPLATE));
  const slide = zip.file("ppt/slides/slide1.xml");
  if (!slide) throw new Error("В шаблоне нет слайда");
  zip.file("ppt/slides/slide1.xml", slideXml(await slide.async("string"), data || {}));
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
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
    const file = await buildPptx(body);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", "attachment; filename*=UTF-8''HR%20Pomoshnik%20Canvas.pptx");
    res.setHeader("Cache-Control", "no-store");
    res.status(200).send(file);
  } catch (error) {
    res.status(500).json({ error: "Не удалось собрать презентацию" });
  }
}
