// 2026 앵커 늘봄 『우리충남 탐험놀이』 — 특산물 · 생활과 공동체
// 29~30차시(3회차) / 31~32차시(4회차) 수업PPT
// 구성은 천안·아산 권역 확정본(5~6차시·7~8차시 수업PPT)의 슬라이드 체계를 따름.
const pptxgen = require("pptxgenjs");

const W = 13.333, H = 7.5, M = 0.75;
const FOREST = "2C5F2D", FOREST_DK = "1C3D1D", MOSS = "97BC62";
const TINT = "EDF3E7", TINT2 = "F7FAF4", ACCENT = "D2622C";
const INK = "1F2A20", MUTED = "6E7C70", WHITE = "FFFFFF", PALE = "C9D8C2";
const FONT = "맑은 고딕";

function newDeck(title) {
  const p = new pptxgen();
  p.defineLayout({ name: "K16x9", width: W, height: H });
  p.layout = "K16x9";
  p.author = "우리충남 탐험놀이 개발팀";
  p.title = title;
  return p;
}

/* ---------------- 공통 요소 ---------------- */

function darkBase(pres, s) {
  s.background = { color: FOREST_DK };
  s.addShape(pres.ShapeType.ellipse, {
    x: W - 3.1, y: -1.5, w: 4.6, h: 4.6,
    fill: { color: FOREST }, line: { color: FOREST, width: 0 },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: -1.4, y: H - 2.4, w: 3.4, h: 3.4,
    fill: { color: FOREST }, line: { color: FOREST, width: 0 },
  });
}

// 우측 하단 활동 번호 표식 (확정본의 '1 / 01' 표기)
function footMark(pres, s, actNo, secNo, opts) {
  if (!actNo) return;
  const fy = (opts && opts.y) || H - 0.88;
  const sc = (opts && opts.secColor) || MUTED;
  s.addShape(pres.ShapeType.ellipse, {
    x: W - M - 0.86, y: fy, w: 0.34, h: 0.34,
    fill: { color: MOSS }, line: { color: MOSS, width: 0 },
  });
  s.addText(String(actNo), {
    x: W - M - 0.86, y: fy, w: 0.34, h: 0.34, align: "center", valign: "middle",
    fontFace: FONT, fontSize: 11, bold: true, color: FOREST_DK, margin: 0, isTextBox: true,
  });
  s.addText(secNo, {
    x: W - M - 0.44, y: fy, w: 0.44, h: 0.34, align: "right", valign: "middle",
    fontFace: FONT, fontSize: 11, bold: true, color: sc, margin: 0, isTextBox: true,
  });
}

function head(pres, s, chip, right) {
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 0.42, w: 2.26, h: 0.38, rectRadius: 0.19,
    fill: { color: TINT }, line: { color: MOSS, width: 1 },
  });
  s.addText(chip, {
    x: M, y: 0.42, w: 2.26, h: 0.38, align: "center", valign: "middle",
    fontFace: FONT, fontSize: 12, bold: true, color: FOREST, margin: 0, isTextBox: true,
  });
  if (right) {
    s.addText(right, {
      x: W - M - 4.6, y: 0.42, w: 4.6, h: 0.38, align: "right", valign: "middle",
      fontFace: FONT, fontSize: 12, color: MUTED, margin: 0, isTextBox: true,
    });
  }
}

function title(s, txt, sub) {
  s.addText(txt, {
    x: M, y: 0.98, w: W - M * 2, h: 0.72,
    fontFace: FONT, fontSize: 33, bold: true, color: INK,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  if (sub) {
    s.addText(sub, {
      x: M, y: 1.70, w: W - M * 2, h: 0.38,
      fontFace: FONT, fontSize: 14, color: MUTED,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  }
}

function badge(pres, s, n, x, y, d, bg, fg, size) {
  s.addShape(pres.ShapeType.ellipse, {
    x, y, w: d, h: d, fill: { color: bg }, line: { color: bg, width: 0 },
  });
  s.addText(String(n), {
    x, y, w: d, h: d, align: "center", valign: "middle",
    fontFace: FONT, fontSize: size, bold: true, color: fg, margin: 0, isTextBox: true,
  });
}

function question(pres, s, txt, y) {
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y, w: W - M * 2, h: 0.92, rectRadius: 0.12,
    fill: { color: FOREST }, line: { color: FOREST, width: 0 },
  });
  s.addText([
    { text: "Q.   ", options: { fontSize: 15, color: MOSS, bold: true } },
    { text: txt, options: { fontSize: 19, color: WHITE, bold: true } },
  ], {
    x: M + 0.32, y, w: W - M * 2 - 0.64, h: 0.92, align: "left", valign: "middle",
    fontFace: FONT, margin: 0, isTextBox: true,
  });
}

function photoSlot(pres, s, label, x, y, w, h, kind, img, fit) {
  if (img) {
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w, h, rectRadius: 0.08,
      fill: { color: TINT2 }, line: { color: MOSS, width: 0 },
    });
    s.addImage({ path: img, x, y, w, h, sizing: { type: fit || "cover", w, h } });
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w, h, rectRadius: 0.08,
      fill: { type: "none" }, line: { color: MOSS, width: 1 },
    });
    return;
  }
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: TINT2 }, line: { color: MOSS, width: 1.5, dashType: "dash" },
  });
  s.addText(kind || "사진 삽입", {
    x, y: y + h / 2 - 0.42, w, h: 0.3, align: "center", valign: "middle",
    fontFace: FONT, fontSize: 11, color: MUTED, margin: 0, isTextBox: true,
  });
  s.addText(label, {
    x: x + 0.1, y: y + h / 2 - 0.10, w: w - 0.2, h: 0.5, align: "center", valign: "middle",
    fontFace: FONT, fontSize: 15, bold: true, color: FOREST, margin: 0, isTextBox: true,
  });
}

// 동기유발 사진 줄 — 1장은 원본 비율, 여러 장은 같은 크기로 맞춰 배치
function imageRow(pres, s, imgs) {
  const gap = 0.24, maxW = W - M * 2, y0 = 1.95, maxH = 2.80, n = imgs.length;
  const boxes = [];
  if (n === 1) {
    const ar = imgs[0].ar || 1.5;
    const w = Math.min(maxW, maxH * ar);
    boxes.push({ x: (W - w) / 2, y: y0, w, h: maxH, path: imgs[0].path, fit: "contain" });
  } else {
    let bw = (maxW - gap * (n - 1)) / n;
    const h = Math.min(maxH, bw / 1.35);
    bw = Math.min(bw, h * 2.0);
    const x0 = (W - (bw * n + gap * (n - 1))) / 2;
    for (let i = 0; i < n; i++) {
      boxes.push({ x: x0 + i * (bw + gap), y: y0, w: bw, h, path: imgs[i].path, fit: "cover" });
    }
  }
  boxes.forEach((b) => {
    s.addImage({ path: b.path, x: b.x, y: b.y, w: b.w, h: b.h, sizing: { type: b.fit, w: b.w, h: b.h } });
    s.addShape(pres.ShapeType.roundRect, {
      x: b.x, y: b.y, w: b.w, h: b.h, rectRadius: 0.08,
      fill: { type: "none" }, line: { color: MOSS, width: 1 },
    });
  });
}

/* ---------------- 슬라이드 종류 ---------------- */

function coverSlide(pres, d) {
  const s = pres.addSlide();
  darkBase(pres, s);
  s.addText(d.region, {
    x: M, y: 1.26, w: 9.6, h: 0.40, fontFace: FONT, fontSize: 17, color: MOSS,
    align: "left", valign: "middle", margin: 0, charSpacing: 2, isTextBox: true,
  });
  s.addText("우리충남 탐험놀이", {
    x: M, y: 1.74, w: 10.5, h: 1.12, fontFace: FONT, fontSize: 54, bold: true, color: WHITE,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 3.16, w: 2.52, h: 0.46, rectRadius: 0.23,
    fill: { color: MOSS }, line: { color: MOSS, width: 0 },
  });
  s.addText(d.round, {
    x: M, y: 3.16, w: 2.52, h: 0.46, align: "center", valign: "middle",
    fontFace: FONT, fontSize: 14, bold: true, color: FOREST_DK, margin: 0, isTextBox: true,
  });
  s.addText(d.subtitle, {
    x: M + 2.80, y: 3.16, w: 9.2, h: 0.46, fontFace: FONT, fontSize: 21, color: WHITE,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("대상 : 3~4학년   ·   80분(2차시)", {
    x: M, y: 4.04, w: 8.0, h: 0.36, fontFace: FONT, fontSize: 14, color: MOSS,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("강사 :                              보조강사 :", {
    x: M, y: 6.28, w: 8.0, h: 0.40, fontFace: FONT, fontSize: 14, color: PALE,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(d.notes);
  return s;
}

function tocSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(pres, s, d.chip, d.round);
  title(s, "목차");
  const rowH = 0.70, gap = 0.16;
  d.items.forEach((it, i) => {
    const y = 2.20 + i * (rowH + gap);
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y, w: W - M * 2, h: rowH, rectRadius: 0.1,
      fill: { color: i % 2 === 0 ? TINT : TINT2 }, line: { color: MOSS, width: 1 },
    });
    badge(pres, s, it.no, M + 0.26, y + 0.13, 0.44, FOREST, WHITE, 13);
    s.addText(it.name, {
      x: M + 0.90, y, w: 8.4, h: rowH,
      fontFace: FONT, fontSize: 17, bold: true, color: INK,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(it.cha, {
      x: W - M - 2.4, y, w: 2.1, h: rowH,
      fontFace: FONT, fontSize: 13, color: MUTED,
      align: "right", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addText("참고자료 : " + d.ref, {
    x: M, y: 6.56, w: W - M * 2, h: 0.40,
    fontFace: FONT, fontSize: 11, color: MUTED,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(d.notes);
  return s;
}

function goalSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(pres, s, d.chip, "학습목표");
  title(s, "학습목표");
  d.goals.forEach((g, i) => {
    const y = 2.34 + i * 1.52;
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y, w: W - M * 2, h: 1.26, rectRadius: 0.1,
      fill: { color: i === 0 ? TINT : TINT2 }, line: { color: MOSS, width: 1 },
    });
    badge(pres, s, i + 1, M + 0.36, y + 0.37, 0.52, i === 0 ? FOREST : MOSS,
      i === 0 ? WHITE : FOREST_DK, 15);
    s.addText(g, {
      x: M + 1.10, y, w: W - M * 2 - 1.5, h: 1.26,
      fontFace: FONT, fontSize: 19, color: INK, lineSpacing: 28,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.62, w: W - M * 2, h: 0.62, rectRadius: 0.1,
    fill: { color: FOREST }, line: { color: FOREST, width: 0 },
  });
  s.addText("관련 교과 :  " + d.subjects, {
    x: M + 0.34, y: 5.62, w: W - M * 2 - 0.68, h: 0.62,
    fontFace: FONT, fontSize: 15, bold: true, color: WHITE,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(d.notes);
  return s;
}

// 동기유발 / 생각해 보아요 — 큰 문장 + 지시문
function promptSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: d.bg || WHITE };
  head(pres, s, d.chip, d.right || "동기유발");
  title(s, d.title);
  if (!d.photo && !d.imgs) {
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y: 2.20, w: W - M * 2, h: 1.70, rectRadius: 0.14,
      fill: { color: WHITE }, line: { color: FOREST, width: 2 },
    });
    s.addText(d.big, {
      x: M + 0.5, y: 2.20, w: W - M * 2 - 1.0, h: 1.70,
      fontFace: FONT, fontSize: d.bigSize || 28, bold: true, color: FOREST, lineSpacing: 40,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
  }
  const stepTop = d.imgs ? 5.01 : 4.16;
  (d.steps || []).forEach((t, i) => {
    const y = stepTop + i * 0.62;
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y, w: W - M * 2, h: 0.52, rectRadius: 0.1,
      fill: { color: i === 0 ? TINT : TINT2 }, line: { color: MOSS, width: 1 },
    });
    badge(pres, s, i + 1, M + 0.22, y + 0.10, 0.32, MOSS, FOREST_DK, 11);
    s.addText(t, {
      x: M + 0.68, y, w: W - M * 2 - 1.0, h: 0.52,
      fontFace: FONT, fontSize: 15, color: INK,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  if (d.foot) {
    const fy = Math.max(6.10, stepTop + (d.steps || []).length * 0.62 + 0.10);
    s.addText(d.foot, {
      x: M, y: fy, w: W - M * 2, h: 0.40,
      fontFace: FONT, fontSize: 12, color: MUTED,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  }
  if (d.imgs) imageRow(pres, s, d.imgs);
  else if (d.photo) photoSlot(pres, s, d.photo, M, 2.20, W - M * 2, 1.70, d.photoKind);
  footMark(pres, s, d.act, d.sec);
  s.addNotes(d.notes);
  return s;
}

// 용어 풀이
function termSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: TINT2 };
  head(pres, s, d.chip, "용어 알아보기");
  title(s, "‘" + d.term + "’ 이란 무엇일까요?");
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 2.26, w: W - M * 2, h: 1.46, rectRadius: 0.14,
    fill: { color: FOREST }, line: { color: FOREST, width: 0 },
  });
  s.addText(d.desc, {
    x: M + 0.5, y: 2.26, w: W - M * 2 - 1.0, h: 1.46,
    fontFace: FONT, fontSize: 23, bold: true, color: WHITE, lineSpacing: 33,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("예를 들어,", {
    x: M, y: 3.94, w: 3.0, h: 0.36,
    fontFace: FONT, fontSize: 14, bold: true, color: ACCENT,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  const cw = (W - M * 2 - 0.42 * 2) / 3;
  d.examples.forEach((e, i) => {
    const x = M + i * (cw + 0.42);
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 4.38, w: cw, h: 1.30, rectRadius: 0.12,
      fill: { color: WHITE }, line: { color: MOSS, width: 1 },
    });
    s.addText(e, {
      x: x + 0.2, y: 4.38, w: cw - 0.4, h: 1.30,
      fontFace: FONT, fontSize: 20, bold: true, color: INK,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  if (d.foot) {
    s.addText(d.foot, {
      x: M, y: 5.90, w: W - M * 2, h: 0.40,
      fontFace: FONT, fontSize: 12, color: MUTED,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  }
  s.addNotes(d.notes);
  return s;
}

// 활동 간지
function dividerSlide(pres, d) {
  const s = pres.addSlide();
  darkBase(pres, s);
  s.addText(d.no, {
    x: M, y: 1.60, w: 3.0, h: 1.9, fontFace: FONT, fontSize: 96, bold: true, color: MOSS,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("활동 " + d.act, {
    x: M, y: 3.48, w: 8.0, h: 0.40, fontFace: FONT, fontSize: 16, color: MOSS,
    align: "left", valign: "middle", margin: 0, charSpacing: 2, isTextBox: true,
  });
  s.addText(d.title, {
    x: M, y: 3.92, w: 10.6, h: 0.96, fontFace: FONT, fontSize: 40, bold: true, color: WHITE,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.14, w: 2.50, h: 0.42, rectRadius: 0.21,
    fill: { color: MOSS }, line: { color: MOSS, width: 0 },
  });
  s.addText(d.meta, {
    x: M, y: 5.14, w: 2.50, h: 0.42, align: "center", valign: "middle",
    fontFace: FONT, fontSize: 13, bold: true, color: FOREST_DK, margin: 0, isTextBox: true,
  });
  if (d.lead) {
    s.addText(d.lead, {
      x: M + 2.78, y: 5.14, w: 8.6, h: 0.42, fontFace: FONT, fontSize: 15, color: PALE,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  }
  s.addNotes(d.notes);
  return s;
}

// '이렇게 해요' ①②③④ + Q.
function howtoSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(pres, s, d.chip, d.right);
  title(s, d.title, "이렇게 해요");
  const circled = ["①", "②", "③", "④", "⑤"];
  d.steps.forEach((t, i) => {
    const y = 2.28 + i * 0.74;
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y, w: W - M * 2, h: 0.62, rectRadius: 0.1,
      fill: { color: i % 2 === 0 ? TINT : TINT2 }, line: { color: MOSS, width: 1 },
    });
    s.addText(circled[i], {
      x: M + 0.26, y, w: 0.5, h: 0.62,
      fontFace: FONT, fontSize: 18, bold: true, color: FOREST,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(t, {
      x: M + 0.84, y, w: W - M * 2 - 1.1, h: 0.62,
      fontFace: FONT, fontSize: 16, color: INK,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  const qy = 2.28 + d.steps.length * 0.74 + 0.24;
  if (d.q) question(pres, s, d.q, Math.min(qy, 6.1));
  footMark(pres, s, d.act, d.sec);
  s.addNotes(d.notes);
  return s;
}

// 카드 가로 배열 (제목 + 본문 + 선택 사진칸)
function cardsSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: d.bg || WHITE };
  head(pres, s, d.chip, d.right);
  title(s, d.title, d.sub);
  const n = d.cards.length;
  const gap = n >= 4 ? 0.32 : 0.42;
  const cw = (W - M * 2 - gap * (n - 1)) / n;
  const top = 2.38, ch = d.photo ? 3.52 : 2.46;
  d.cards.forEach((c, i) => {
    const x = M + i * (cw + gap);
    if (d.photo) {
      const ph = c.img ? 1.80 : 1.92;
      const by = c.img ? 2.24 : 2.06;
      photoSlot(pres, s, c.head, x, top, cw, ph, d.photoKind, c.img, d.imgFit);
      if (c.img) {
        s.addText(c.head, {
          x: x + 0.22, y: top + ph + 0.06, w: cw - 0.44, h: 0.34,
          fontFace: FONT, fontSize: 14, bold: true, color: FOREST,
          align: "left", valign: "middle", margin: 0, isTextBox: true,
        });
      }
      s.addShape(pres.ShapeType.roundRect, {
        x, y: top + by, w: cw, h: ch - by, rectRadius: 0.1,
        fill: { color: TINT2 }, line: { color: MOSS, width: 1 },
      });
      s.addText(c.body, {
        x: x + 0.22, y: top + by + 0.12, w: cw - 0.44, h: ch - by - 0.24,
        fontFace: FONT, fontSize: 13, color: INK, lineSpacing: 18,
        align: "left", valign: "top", margin: 0, isTextBox: true,
      });
    } else {
      s.addShape(pres.ShapeType.roundRect, {
        x, y: top, w: cw, h: ch, rectRadius: 0.12,
        fill: { color: i === 0 ? TINT : TINT2 }, line: { color: MOSS, width: 1 },
      });
      if (c.tag) {
        s.addText(c.tag, {
          x: x + 0.28, y: top + 0.20, w: cw - 0.56, h: 0.30,
          fontFace: FONT, fontSize: 12, bold: true, color: ACCENT,
          align: "left", valign: "middle", margin: 0, isTextBox: true,
        });
      }
      s.addText(c.head, {
        x: x + 0.28, y: top + (c.tag ? 0.56 : 0.28), w: cw - 0.56, h: 0.72,
        fontFace: FONT, fontSize: 19, bold: true, color: INK, lineSpacing: 25,
        align: "left", valign: "top", margin: 0, isTextBox: true,
      });
      s.addText(c.body, {
        x: x + 0.28, y: top + (c.tag ? 1.34 : 1.06), w: cw - 0.56, h: ch - (c.tag ? 1.52 : 1.24),
        fontFace: FONT, fontSize: 13, color: MUTED, lineSpacing: 19,
        align: "left", valign: "top", margin: 0, isTextBox: true,
      });
    }
  });
  const by = top + ch + 0.26;
  if (d.band) {
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y: by, w: W - M * 2, h: 0.82, rectRadius: 0.1,
      fill: { color: FOREST }, line: { color: FOREST, width: 0 },
    });
    s.addText(d.band, {
      x: M + 0.34, y: by, w: W - M * 2 - 1.34, h: 0.82,
      fontFace: FONT, fontSize: 16, bold: true, color: WHITE,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  } else if (d.q) {
    question(pres, s, d.q, by);
  }
  if (d.foot) {
    const fy = (d.band || d.q) ? Math.min(by + 0.94, 6.46) : Math.min(by, 6.46);
    s.addText(d.foot, {
      x: M, y: fy, w: W - M * 2 - 1.5, h: 0.36,
      fontFace: FONT, fontSize: 12, color: MUTED,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  }
  footMark(pres, s, d.act, d.sec, d.band ? { y: by + 0.24, secColor: PALE } : null);
  s.addNotes(d.notes);
  return s;
}

// 표 슬라이드
function tableSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: d.bg || WHITE };
  head(pres, s, d.chip, d.right);
  title(s, d.title, d.sub);
  const cols = d.headers.length;
  const cw = (W - M * 2) / cols;
  const hy = 2.42;
  // 표 + 핵심 발문이 활동 번호 표식과 겹치지 않도록 행 높이를 자동으로 줄인다
  const limit = d.q ? 5.92 : 6.38;
  const avail = (limit - (hy + 0.62)) / d.rows.length;
  const rh = Math.min(d.rowH || 0.86, avail);
  d.headers.forEach((h, i) => {
    const x = M + i * cw;
    s.addShape(pres.ShapeType.rect, {
      x, y: hy, w: cw, h: 0.62,
      fill: { color: FOREST }, line: { color: WHITE, width: 1 },
    });
    s.addText(h, {
      x: x + 0.1, y: hy, w: cw - 0.2, h: 0.62,
      fontFace: FONT, fontSize: 16, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  d.rows.forEach((row, r) => {
    const y = hy + 0.62 + r * rh;
    row.forEach((cell, i) => {
      const x = M + i * cw;
      s.addShape(pres.ShapeType.rect, {
        x, y, w: cw, h: rh,
        fill: { color: r % 2 === 0 ? TINT2 : WHITE }, line: { color: MOSS, width: 1 },
      });
      const firstCol = !d.plain && i === 0;
      s.addText(cell, {
        x: x + 0.14, y, w: cw - 0.28, h: rh,
        fontFace: FONT, fontSize: d.plain ? 17 : (firstCol ? 15 : 14),
        bold: firstCol, color: firstCol ? FOREST : INK, lineSpacing: d.plain ? 26 : 19,
        align: d.plain || firstCol ? "center" : "left",
        valign: "middle", margin: 0, isTextBox: true,
      });
    });
  });
  const by = hy + 0.62 + d.rows.length * rh + 0.26;
  if (d.q && by < 6.2) question(pres, s, d.q, by);
  else if (d.foot) {
    s.addText(d.foot, {
      x: M, y: Math.min(by, 6.5), w: W - M * 2, h: 0.40,
      fontFace: FONT, fontSize: 12, color: MUTED,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  }
  footMark(pres, s, d.act, d.sec);
  s.addNotes(d.notes);
  return s;
}

// '원하는 것은 무엇이든!' 아이디어 슬라이드
function ideaSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: TINT2 };
  head(pres, s, d.chip, d.right);
  title(s, d.title, d.sub);
  const cw = (W - M * 2 - 0.42 * 2) / 3;
  d.examples.forEach((e, i) => {
    const x = M + i * (cw + 0.42);
    const last = i === 2;
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 2.44, w: cw, h: 1.90, rectRadius: 0.14,
      fill: { color: last ? MOSS : WHITE }, line: { color: last ? MOSS : MOSS, width: last ? 0 : 1 },
    });
    s.addText(e.head, {
      x: x + 0.24, y: 2.62, w: cw - 0.48, h: 0.94,
      fontFace: FONT, fontSize: last ? 20 : 21, bold: true,
      color: last ? FOREST_DK : INK, lineSpacing: 27,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(e.body, {
      x: x + 0.24, y: 3.64, w: cw - 0.48, h: 0.62,
      fontFace: FONT, fontSize: 13, color: last ? "2C5F2D" : MUTED, lineSpacing: 18,
      align: "center", valign: "top", margin: 0, isTextBox: true,
    });
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 4.66, w: W - M * 2, h: 1.10, rectRadius: 0.12,
    fill: { color: FOREST }, line: { color: FOREST, width: 0 },
  });
  s.addText(d.band, {
    x: M + 0.4, y: 4.66, w: W - M * 2 - 0.8, h: 1.10,
    fontFace: FONT, fontSize: 21, bold: true, color: WHITE, lineSpacing: 29,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  if (d.foot) {
    s.addText(d.foot, {
      x: M, y: 5.96, w: W - M * 2, h: 0.40,
      fontFace: FONT, fontSize: 12, color: MUTED,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  }
  footMark(pres, s, d.act, d.sec);
  s.addNotes(d.notes);
  return s;
}

// 3단 규칙 슬라이드
function rulesSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(pres, s, d.chip, d.right);
  title(s, d.title, d.sub);
  const cw = (W - M * 2 - 0.42 * 2) / 3;
  d.cols.forEach((c, i) => {
    const x = M + i * (cw + 0.42);
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 2.40, w: cw, h: 2.86, rectRadius: 0.12,
      fill: { color: WHITE }, line: { color: FOREST, width: 1.5 },
    });
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 2.40, w: cw, h: 0.66, rectRadius: 0.12,
      fill: { color: FOREST }, line: { color: FOREST, width: 0 },
    });
    s.addText(c.head, {
      x: x + 0.2, y: 2.40, w: cw - 0.4, h: 0.66,
      fontFace: FONT, fontSize: 18, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(c.body, {
      x: x + 0.28, y: 3.22, w: cw - 0.56, h: 1.90,
      fontFace: FONT, fontSize: 15, color: INK, lineSpacing: 24,
      align: "left", valign: "top", margin: 0, isTextBox: true,
    });
  });
  if (d.band) {
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y: 5.52, w: W - M * 2, h: 0.86, rectRadius: 0.1,
      fill: { color: TINT }, line: { color: TINT, width: 0 },
    });
    s.addText(d.band, {
      x: M + 0.34, y: 5.52, w: W - M * 2 - 0.68, h: 0.86,
      fontFace: FONT, fontSize: 16, bold: true, color: FOREST,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  }
  footMark(pres, s, d.act, d.sec);
  s.addNotes(d.notes);
  return s;
}

// 학습정리 / 차시예고
function listSlide(pres, d) {
  const s = pres.addSlide();
  s.background = { color: d.bg || TINT2 };
  head(pres, s, d.chip, d.right);
  title(s, d.title, d.sub);
  d.items.forEach((t, i) => {
    const y = 2.56 + i * 1.46;
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y, w: W - M * 2, h: 1.20, rectRadius: 0.12,
      fill: { color: WHITE }, line: { color: FOREST, width: i === 0 ? 2 : 1 },
    });
    badge(pres, s, i + 1, M + 0.36, y + 0.34, 0.52, i === 0 ? FOREST : MOSS,
      i === 0 ? WHITE : FOREST_DK, 15);
    s.addText(t, {
      x: M + 1.10, y, w: W - M * 2 - 1.5, h: 1.20,
      fontFace: FONT, fontSize: 20, bold: true, color: INK, lineSpacing: 28,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  if (d.foot) {
    s.addText(d.foot, {
      x: M, y: 5.84, w: W - M * 2, h: 0.40,
      fontFace: FONT, fontSize: 13, color: MUTED,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  }
  s.addNotes(d.notes);
  return s;
}

function thanksSlide(pres, notes, region) {
  const s = pres.addSlide();
  darkBase(pres, s);
  s.addText("수고했어요!", {
    x: M, y: 2.90, w: 10.5, h: 1.30, fontFace: FONT, fontSize: 48, bold: true, color: WHITE,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("2026 앵커 늘봄 프로그램  ·  우리충남 탐험놀이  ·  " + (region || "공주 · 부여"), {
    x: M, y: 4.30, w: 11.0, h: 0.40, fontFace: FONT, fontSize: 14, color: MOSS,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes(notes);
  return s;
}

module.exports = {
  newDeck, coverSlide, tocSlide, goalSlide, promptSlide, termSlide,
  dividerSlide, howtoSlide, cardsSlide, tableSlide, ideaSlide,
  rulesSlide, listSlide, thanksSlide,
};
