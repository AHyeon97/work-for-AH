// 활동지 · 해설지 공통 라이브러리 (A4 세로 2쪽)
// 천안·아산 5~6차시 활동지 원본의 좌표·색·글꼴 체계를 따름.
const pptxgen = require("pptxgenjs");

const W = 8.27, H = 11.69;
const M = 0.55, CW = 7.17;            // 좌여백 / 본문 폭
const COL = [0.55, 4.21], COLW = 3.5; // 2단 답란
const LBLX = [0.69, 4.35], LBLW = 3.22;

const ORANGE = "F0813C", ORANGE_L = "F7AE72", ORANGE_P = "F5CDA9", CREAM = "FDF3EB";
const INK = "2B2F3A", WHITE = "FFFFFF", GREY = "7B8190";
const DECO = ["E8615C", "F0813C", "F2C230", "6FC08A", "5B8DEF", "B58BE0", "F7AE72", "6FC08A", "E8615C"];

const FONT = "학교안심 지우개 R";
const FONT_T = "학교안심 알림장 OTF R";

function newSheet(title) {
  const p = new pptxgen();
  p.defineLayout({ name: "A4P", width: W, height: H });
  p.layout = "A4P";
  p.author = "우리충남 탐험놀이 개발팀";
  p.title = title;
  return p;
}

/* ---------- 페이지 뼈대 ---------- */
function page(pres, opts) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.22, fill: { color: ORANGE }, line: { width: 0 } });
  s.addShape(pres.ShapeType.rect, { x: 0, y: H - 0.22, w: W, h: 0.22, fill: { color: ORANGE }, line: { width: 0 } });
  DECO.forEach((c, i) => {
    s.addShape(pres.ShapeType.rect, {
      x: 0.12 + i * 0.893, y: H - 0.40, w: 0.79, h: 0.15,
      fill: { color: c }, line: { width: 0 },
    });
  });
  if (opts && opts.foot) {
    s.addText(opts.foot, {
      x: M, y: H - 0.66, w: CW, h: 0.20, align: "right", valign: "middle",
      fontFace: FONT, fontSize: 8, color: GREY, margin: 0, isTextBox: true,
    });
  }
  return s;
}

// 1쪽 머리글: 권역·회차 / 제목 박스 / 이름란
function header(pres, s, d) {
  s.addText(d.head, {
    x: M, y: 0.50, w: CW, h: 0.24, fontFace: FONT, fontSize: 10, color: GREY,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.39, y: 0.96, w: 0.26, h: 0.40, rectRadius: 0.08, fill: { color: ORANGE_P }, line: { width: 0 },
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 6.57, y: 0.96, w: 0.26, h: 0.40, rectRadius: 0.08, fill: { color: ORANGE_P }, line: { width: 0 },
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 0.80, w: 6.12, h: 0.72, rectRadius: 0.12, fill: { color: CREAM }, line: { color: ORANGE, width: 1.25 },
  });
  s.addShape(pres.ShapeType.rect, {
    x: 0.90, y: 1.33, w: 5.42, h: 0.09, fill: { color: ORANGE_P }, line: { width: 0 },
  });
  s.addText(d.title, {
    x: 0.90, y: 0.80, w: 5.42, h: 0.72, fontFace: FONT_T, fontSize: 20, bold: true, color: INK,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 6.89, y: 0.74, w: 0.83, h: 0.84, rectRadius: 0.16, fill: { color: ORANGE }, line: { width: 0 },
  });
  s.addText(d.badge, {
    x: 6.89, y: 0.74, w: 0.83, h: 0.84, fontFace: FONT, fontSize: 11, bold: true, color: WHITE,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("(        ) 학년     (        ) 반     (        ) 번호     이름 (                    )", {
    x: M, y: 1.67, w: CW, h: 0.26, fontFace: FONT, fontSize: 10, color: INK,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
}

/* ---------- 문항 요소 ---------- */
function itemHead(pres, s, no, title, y) {
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y, w: 0.46, h: 0.34, rectRadius: 0.08, fill: { color: ORANGE }, line: { width: 0 },
  });
  s.addText(no, {
    x: M, y, w: 0.46, h: 0.34, fontFace: FONT, fontSize: 11, bold: true, color: WHITE,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText(title, {
    x: 1.19, y, w: 6.53, h: 0.34, fontFace: FONT_T, fontSize: 13, bold: true, color: INK,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  return y + 0.47;
}

function guideBox(pres, s, text, y, h) {
  h = h || 0.36;
  s.addShape(pres.ShapeType.rect, { x: M, y, w: CW, h, fill: { color: CREAM }, line: { width: 0 } });
  s.addShape(pres.ShapeType.rect, { x: M, y, w: 0.07, h, fill: { color: ORANGE }, line: { width: 0 } });
  s.addText(text, {
    x: 0.79, y, w: 6.71, h, fontFace: FONT, fontSize: 10, color: INK, lineSpacing: 14,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  return y + h + 0.23;
}

// 라벨이 위에 붙은 답 상자 (칸 수 1 또는 2), 안에 필기선
function topBoxes(pres, s, labels, y, h, lines, fills) {
  h = h || 1.40;
  const n = labels.length, gap = 0.16;
  const bw = (CW - gap * (n - 1)) / n;
  labels.forEach((lb, i) => {
    const x = M + i * (bw + gap), w = bw;
    const lx = x + 0.14, lw = w - 0.28;
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w, h, rectRadius: 0.08,
      fill: { color: (fills && fills[i]) || WHITE }, line: { color: ORANGE_P, width: 1.25 },
    });
    s.addShape(pres.ShapeType.rect, { x: lx, y: y + 0.10, w: lw, h: 0.30, fill: { color: ORANGE_P }, line: { width: 0 } });
    s.addText(lb, {
      x: lx, y: y + 0.10, w: lw, h: 0.30, fontFace: FONT, fontSize: 10, bold: true, color: INK,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
    const nl = lines === undefined ? 2 : lines;
    const top = y + 0.46, avail = h - 0.58;   // 라벨 아래 ~ 상자 바닥
    for (let k = 0; k < nl; k++) {
      s.addShape(pres.ShapeType.line, {
        x: lx + 0.06, y: top + (avail * (k + 1)) / (nl + 1), w: lw - 0.12, h: 0,
        line: { color: ORANGE_P, width: 0.75, dashType: "sysDash" },
      });
    }
  });
  return y + h + 0.24;
}

// 윗단 띠 + 아래 라벨 카드 (2열 n행)
function labelCards(pres, s, labels, y, h, cols) {
  h = h || 1.37;
  const nc = cols || 2, gap = 0.16;
  const bw = (CW - gap * (nc - 1)) / nc;
  labels.forEach((lb, i) => {
    const r = Math.floor(i / nc), c = i % nc;
    const x = M + c * (bw + gap), yy = y + r * (h + gap);
    s.addShape(pres.ShapeType.roundRect, {
      x, y: yy, w: bw, h, rectRadius: 0.08, fill: { color: WHITE }, line: { color: ORANGE_P, width: 1.25 },
    });
    s.addShape(pres.ShapeType.rect, { x, y: yy, w: bw, h: 0.11, fill: { color: ORANGE }, line: { width: 0 } });
    s.addShape(pres.ShapeType.rect, {
      x: x + 0.14, y: yy + h - 0.42, w: bw - 0.28, h: 0.32, fill: { color: CREAM }, line: { width: 0 },
    });
    s.addText(lb, {
      x: x + 0.14, y: yy + h - 0.42, w: bw - 0.28, h: 0.32, fontFace: FONT,
      fontSize: nc >= 4 ? 9 : 10, bold: true, color: INK,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  const rows = Math.ceil(labels.length / nc);
  return y + rows * h + (rows - 1) * gap + 0.24;
}

// 가로 흐름 칩 (생산→제조→판매→소비 같은 띠)
function chips(pres, s, items, y) {
  const n = items.length, gap = 0.16;
  const w = (CW - gap * (n - 1)) / n, h = 0.66;
  items.forEach((it, i) => {
    const x = M + i * (w + gap);
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w, h, rectRadius: 0.08, fill: { color: CREAM }, line: { color: ORANGE_P, width: 1 },
    });
    s.addShape(pres.ShapeType.rect, { x, y: y + h - 0.10, w, h: 0.10, fill: { color: ORANGE_L }, line: { width: 0 } });
    s.addText(it, {
      x: x + 0.08, y, w: w - 0.16, h: h - 0.07, fontFace: FONT, fontSize: 9, color: INK, lineSpacing: 12,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  return y + h + 0.24;
}

// 필기선 n줄
function writeLines(pres, s, y, n, step) {
  step = step || 0.42;
  for (let k = 0; k < n; k++) {
    s.addShape(pres.ShapeType.line, {
      x: 0.70, y: y + step * (k + 1), w: 6.87, h: 0,
      line: { color: ORANGE_P, width: 0.75, dashType: "sysDash" },
    });
  }
  return y + step * n + 0.26;
}

// 체크 줄 (짝 점검 / 확인!)
function checkRow(pres, s, label, items, y) {
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y, w: 1.42, h: 0.40, rectRadius: 0.08, fill: { color: ORANGE }, line: { width: 0 },
  });
  s.addText(label, {
    x: M, y, w: 1.42, h: 0.40, fontFace: FONT, fontSize: 10, bold: true, color: WHITE,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  const area = CW - 1.58, each = area / items.length;
  items.forEach((it, i) => {
    const x = 2.13 + i * each;
    s.addShape(pres.ShapeType.rect, {
      x, y: y + 0.12, w: 0.14, h: 0.14, fill: { color: WHITE }, line: { color: INK, width: 0.75 },
    });
    s.addText(it, {
      x: x + 0.21, y: y + 0.05, w: each - 0.28, h: 0.26, fontFace: FONT, fontSize: 9, color: INK,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  return y + 0.40 + 0.24;
}

// 다짐 상자
function pledge(pres, s, lead, text, y) {
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y, w: CW, h: 0.94, rectRadius: 0.10, fill: { color: CREAM }, line: { color: ORANGE, width: 1.25 },
  });
  s.addShape(pres.ShapeType.rect, { x: M, y, w: CW, h: 0.10, fill: { color: ORANGE }, line: { width: 0 } });
  s.addText(lead, {
    x: 0.79, y: y + 0.18, w: 6.69, h: 0.26, fontFace: FONT, fontSize: 9, color: GREY,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText(text, {
    x: 0.79, y: y + 0.45, w: 6.69, h: 0.37, fontFace: FONT_T, fontSize: 12, bold: true, color: ORANGE,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  return y + 0.94 + 0.24;
}

/* ---------- 해설지 요소 ---------- */
function keyBox(pres, s, pairs, y) {
  // pairs: [[라벨, 예시답안], ...]  1~2열 자동
  const two = pairs.length > 1;
  let maxH = 0;
  pairs.forEach((pr, i) => {
    const x = two ? COL[i % 2] : M, w = two ? COLW : CW;
    const r = two ? Math.floor(i / 2) : i;
    const lines = Math.ceil(pr[1].length / (two ? 26 : 54));
    const h = Math.max(0.78, 0.40 + lines * 0.20);
    const yy = y + r * (h + 0.14);
    maxH = Math.max(maxH, (r + 1) * (h + 0.14));
    s.addShape(pres.ShapeType.roundRect, {
      x, y: yy, w, h, rectRadius: 0.08, fill: { color: "F4FBF6" }, line: { color: "6FC08A", width: 1 },
    });
    s.addText(pr[0], {
      x: x + 0.14, y: yy + 0.06, w: w - 0.28, h: 0.24, fontFace: FONT, fontSize: 9, bold: true, color: "2E7D4F",
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(pr[1], {
      x: x + 0.14, y: yy + 0.30, w: w - 0.28, h: h - 0.38, fontFace: FONT, fontSize: 9.5, color: INK, lineSpacing: 13,
      align: "left", valign: "top", margin: 0, isTextBox: true,
    });
  });
  return y + maxH + 0.10;
}

function tipLine(pres, s, tag, text, y, color) {
  const c = color || ORANGE;
  const est = Math.ceil((tag.length * 2 + text.length) / 60);
  const h = Math.max(0.26, est * 0.20 + 0.08);
  s.addText([
    { text: tag + "  ", options: { fontSize: 9, bold: true, color: c } },
    { text: text, options: { fontSize: 9, color: INK } },
  ], {
    x: 0.70, y, w: 6.87, h, fontFace: FONT, lineSpacing: 12.5,
    align: "left", valign: "top", margin: 0, isTextBox: true,
  });
  return y + h + 0.06;
}

function warnBox(pres, s, text, y) {
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y, w: CW, h: 0.46, rectRadius: 0.08, fill: { color: "FDECEA" }, line: { color: "E8615C", width: 1 },
  });
  s.addText([
    { text: "배부 주의  ", options: { fontSize: 9, bold: true, color: "C0392B" } },
    { text: text, options: { fontSize: 9, color: INK } },
  ], {
    x: 0.72, y, w: 6.83, h: 0.46, fontFace: FONT, lineSpacing: 12,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  return y + 0.46 + 0.22;
}

module.exports = {
  W, H, M, CW, COL, COLW, ORANGE, ORANGE_L, ORANGE_P, CREAM, INK, WHITE, GREY, FONT, FONT_T,
  newSheet, page, header, itemHead, guideBox, topBoxes, labelCards, chips,
  writeLines, checkRow, pledge, keyBox, tipLine, warnBox,
};
