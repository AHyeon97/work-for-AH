// 2026 앵커 늘봄 프로그램 『우리고장 탐험놀이』 7~8차시 수업용 PPT
// 출처: 2026 앵커늘봄_우리고장 탐험놀이_유혜리, 유아현 (시트2 '1~8차시 수업흐름' 확정본)
const pptxgen = require("pptxgenjs");

const W = 13.333, H = 7.5;
const M = 0.75;                 // 좌우 여백
const FOREST = "2C5F2D";
const FOREST_DK = "1C3D1D";
const MOSS = "97BC62";
const TINT = "EDF3E7";
const TINT2 = "F7FAF4";
const ACCENT = "D2622C";
const INK = "1F2A20";
const MUTED = "6E7C70";
const WHITE = "FFFFFF";
const FONT = "맑은 고딕";

const pres = new pptxgen();
pres.defineLayout({ name: "K16x9", width: W, height: H });
pres.layout = "K16x9";
pres.author = "우리고장 탐험놀이 개발팀";
pres.title = "우리고장 탐험놀이 7~8차시 수업용 프레젠테이션";

/* ---------- 공통 헬퍼 ---------- */

// 상단 좌측 차시 배지 + 우측 단계/시간 표기
function header(s, chaTxt, rightTxt) {
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 0.42, w: 1.62, h: 0.38, rectRadius: 0.19,
    fill: { color: TINT }, line: { color: MOSS, width: 1 },
  });
  s.addText(chaTxt, {
    x: M, y: 0.42, w: 1.62, h: 0.38, align: "center", valign: "middle",
    fontFace: FONT, fontSize: 12, bold: true, color: FOREST, margin: 0, isTextBox: true,
  });
  if (rightTxt) {
    s.addText(rightTxt, {
      x: W - M - 4.0, y: 0.42, w: 4.0, h: 0.38, align: "right", valign: "middle",
      fontFace: FONT, fontSize: 12, color: MUTED, margin: 0, isTextBox: true,
    });
  }
}

// 본문 슬라이드 제목
function title(s, txt, sub) {
  s.addText(txt, {
    x: M, y: 0.98, w: W - M * 2, h: 0.72,
    fontFace: FONT, fontSize: 34, bold: true, color: INK,
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

// 번호 원형 배지
function badge(s, n, x, y, d, bg, fg, size) {
  s.addShape(pres.ShapeType.ellipse, {
    x, y, w: d, h: d, fill: { color: bg || FOREST }, line: { color: bg || FOREST, width: 0 },
  });
  s.addText(String(n), {
    x, y, w: d, h: d, align: "center", valign: "middle",
    fontFace: FONT, fontSize: size || 14, bold: true, color: fg || WHITE, margin: 0, isTextBox: true,
  });
}

// 발문 박스 (강조 인용)
function question(s, txt, x, y, w, h) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.12,
    fill: { color: FOREST }, line: { color: FOREST, width: 0 },
  });
  s.addText([
    { text: "핵심 발문   ", options: { fontSize: 12, color: MOSS, bold: true } },
    { text: txt, options: { fontSize: 19, color: WHITE, bold: true } },
  ], {
    x: x + 0.32, y, w: w - 0.64, h, align: "left", valign: "middle",
    fontFace: FONT, margin: 0, isTextBox: true,
  });
}

// 사진 자리표시
function photoSlot(s, label, x, y, w, h, kind) {
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

// 표지 / 간지 / 마무리 : 짙은 배경
function darkBase(s) {
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

/* ---------- 슬라이드 생성기 ---------- */

function coverSlide(cha, subject) {
  const s = pres.addSlide();
  darkBase(s);
  s.addText("2026 앵커 늘봄 프로그램   ·   천안 · 아산 〈호서대학교〉", {
    x: M, y: 1.30, w: 9.6, h: 0.36, fontFace: FONT, fontSize: 14, color: MOSS,
    align: "left", valign: "middle", margin: 0, charSpacing: 1, isTextBox: true,
  });
  s.addText("우리고장 탐험놀이", {
    x: M, y: 1.78, w: 10.5, h: 1.12, fontFace: FONT, fontSize: 54, bold: true, color: WHITE,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 3.14, w: 1.34, h: 0.46, rectRadius: 0.23,
    fill: { color: MOSS }, line: { color: MOSS, width: 0 },
  });
  s.addText(cha, {
    x: M, y: 3.14, w: 1.34, h: 0.46, align: "center", valign: "middle",
    fontFace: FONT, fontSize: 15, bold: true, color: FOREST_DK, margin: 0, isTextBox: true,
  });
  s.addText(subject, {
    x: M + 1.62, y: 3.14, w: 10.2, h: 0.46, fontFace: FONT, fontSize: 21, color: WHITE,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("공통 주제  생활과 공동체   ·   40분", {
    x: M, y: 4.02, w: 8.0, h: 0.36, fontFace: FONT, fontSize: 14, color: MOSS,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("강사 :                              보조강사 :", {
    x: M, y: 6.28, w: 8.0, h: 0.40, fontFace: FONT, fontSize: 14, color: "C9D8C2",
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  return s;
}

function tocSlide(chaTxt, items, closing) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  header(s, chaTxt, "40분 수업");
  title(s, "목차");
  const cw = (W - M * 2 - 0.42) / 2, ch = 1.32;
  items.forEach((it, i) => {
    const x = M + (i % 2) * (cw + 0.42);
    const y = 2.18 + Math.floor(i / 2) * (ch + 0.34);
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w: cw, h: ch, rectRadius: 0.1,
      fill: { color: TINT2 }, line: { color: MOSS, width: 1 },
    });
    badge(s, it.no, x + 0.34, y + 0.34, 0.64, FOREST, WHITE, 17);
    s.addText(it.name, {
      x: x + 1.18, y: y + 0.28, w: cw - 1.5, h: 0.44,
      fontFace: FONT, fontSize: 18, bold: true, color: INK,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(it.meta, {
      x: x + 1.18, y: y + 0.72, w: cw - 1.5, h: 0.34,
      fontFace: FONT, fontSize: 13, color: MUTED,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.86, w: W - M * 2, h: 0.62, rectRadius: 0.1,
    fill: { color: TINT }, line: { color: TINT, width: 0 },
  });
  s.addText(closing, {
    x: M + 0.34, y: 5.86, w: W - M * 2 - 0.68, h: 0.62,
    fontFace: FONT, fontSize: 14, color: FOREST, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  return s;
}

function goalSlide(chaTxt, goals) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  header(s, chaTxt, "학습목표");
  title(s, "학습목표", "오늘 수업에서 우리가 할 수 있게 되는 것들이에요.");
  goals.forEach((g, i) => {
    const y = 2.70 + i * 1.36;
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y, w: W - M * 2, h: 1.06, rectRadius: 0.1,
      fill: { color: i === 0 ? TINT : TINT2 }, line: { color: MOSS, width: 1 },
    });
    badge(s, i + 1, M + 0.36, y + 0.27, 0.52, i === 0 ? FOREST : MOSS, i === 0 ? WHITE : FOREST_DK, 15);
    s.addText(g, {
      x: M + 1.08, y, w: W - M * 2 - 1.5, h: 1.06,
      fontFace: FONT, fontSize: 18, color: INK,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  return s;
}

function dividerSlide(no, stage, name, minutes, lead) {
  const s = pres.addSlide();
  darkBase(s);
  s.addText(no, {
    x: M, y: 1.66, w: 3.0, h: 1.9, fontFace: FONT, fontSize: 96, bold: true, color: MOSS,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText(stage, {
    x: M, y: 3.52, w: 8.0, h: 0.40, fontFace: FONT, fontSize: 16, color: MOSS,
    align: "left", valign: "middle", margin: 0, charSpacing: 2, isTextBox: true,
  });
  s.addText(name, {
    x: M, y: 3.96, w: 10.6, h: 0.96, fontFace: FONT, fontSize: 40, bold: true, color: WHITE,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.16, w: 1.12, h: 0.42, rectRadius: 0.21,
    fill: { color: MOSS }, line: { color: MOSS, width: 0 },
  });
  s.addText(minutes, {
    x: M, y: 5.16, w: 1.12, h: 0.42, align: "center", valign: "middle",
    fontFace: FONT, fontSize: 14, bold: true, color: FOREST_DK, margin: 0, isTextBox: true,
  });
  if (lead) {
    s.addText(lead, {
      x: M + 1.40, y: 5.16, w: 9.2, h: 0.42, fontFace: FONT, fontSize: 15, color: "C9D8C2",
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  }
  return s;
}

function baseSlide(chaTxt, rightTxt, titleTxt, subTxt, bg) {
  const s = pres.addSlide();
  s.background = { color: bg || WHITE };
  header(s, chaTxt, rightTxt);
  title(s, titleTxt, subTxt);
  return s;
}

// 가로 3~4단 흐름(화살표 연결)
function flowRow(s, steps, y, h, opts) {
  const o = opts || {};
  const x0 = o.x !== undefined ? o.x : M;
  const total = o.w !== undefined ? o.w : W - M * 2;
  const gap = 0.46;
  const cw = (total - gap * (steps.length - 1)) / steps.length;
  steps.forEach((st, i) => {
    const x = x0 + i * (cw + gap);
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w: cw, h, rectRadius: 0.1,
      fill: { color: o.fill || TINT2 }, line: { color: MOSS, width: 1 },
    });
    s.addText(st.head, {
      x: x + 0.16, y: y + 0.20, w: cw - 0.32, h: 0.42,
      fontFace: FONT, fontSize: o.headSize || 16, bold: true, color: FOREST,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
    if (st.body) {
      s.addText(st.body, {
        x: x + 0.16, y: y + 0.62, w: cw - 0.32, h: h - 0.80,
        fontFace: FONT, fontSize: o.bodySize || 13, color: INK,
        align: "center", valign: "top", margin: 0, isTextBox: true,
      });
    }
    if (i < steps.length - 1) {
      s.addShape(pres.ShapeType.rightArrow, {
        x: x + cw + 0.09, y: y + h / 2 - 0.13, w: 0.28, h: 0.26,
        fill: { color: MOSS }, line: { color: MOSS, width: 0 },
      });
    }
  });
}

/* =========================================================
   7차시 | 천안·아산의 생활 공간과 서로 돕는 사람들을 알아보아요
   ========================================================= */
const C7 = "7차시";

// 1. 표지
coverSlide("7차시", "천안 · 아산의 생활 공간과 서로 돕는 사람들을 알아보아요")
  .addNotes("4회차(7~8차시) 중 앞 차시입니다. 공통 주제는 '생활과 공동체'이며, 관광 명소가 아니라 주민의 일상 생활 공간에 초점을 둡니다.");

// 2. 목차
tocSlide(C7, [
  { no: "01", name: "내 하루 속 도움 찾기", meta: "도입 · 5분" },
  { no: "02", name: "생활 사진 탐정", meta: "활동 1 · 8분" },
  { no: "03", name: "우리 마을 도움 연결지도", meta: "활동 2 · 15분" },
  { no: "04", name: "도움이 멈춘 하루", meta: "활동 3 · 8분" },
], "정리   ·   연결 문장 완성하기 (4분)  —  완성한 연결지도는 8차시 마을회의에 그대로 사용해요.");

// 3. 학습목표
goalSlide(C7, [
  "천안·아산의 생활 공간이 어떤 곳이고 누가 이용하는지 설명할 수 있다.",
  "마을에서 도움을 주고받는 사람과 장소의 관계를 연결지도로 나타낼 수 있다.",
  "도움이 멈춘 상황에서 불편을 겪을 사람을 찾아 까닭과 함께 말할 수 있다.",
]);

// 4. 간지 01
dividerSlide("01", "배움 열기", "내 하루 속 도움 찾기", "5분", "어린이의 하루 이야기를 들으며 도움을 준 사람과 장소를 찾아요.");

// 5. 도입
{
  const s = baseSlide(C7, "도입 · 5분", "내 하루 속 도움 찾기",
    "강사가 들려주는 어린이의 하루를 따라가며, 도움을 준 사람과 장소를 찾아보아요.");
  flowRow(s, [
    { head: "버스를 타요", body: "정류장에서\n버스를 기다려요" },
    { head: "도서관에 가요", body: "읽고 싶은 책을\n찾아요" },
    { head: "공원에서 놀아요", body: "친구와 함께\n뛰어놀아요" },
    { head: "집으로 돌아와요", body: "다시 버스를 타고\n집에 가요" },
  ], 2.42, 1.86);
  question(s, "이 아이가 하루를 보내려면 누구의 도움이 필요했을까요?", M, 4.78, W - M * 2, 0.96);
  s.addText("장소 이름과 사람의 역할을 나누어 말하게 해 주세요.", {
    x: M, y: 5.98, w: W - M * 2, h: 0.36, fontFace: FONT, fontSize: 13, color: MUTED,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes("강사가 '버스를 타고 도서관에 갔다가 공원에서 놀고 돌아온 어린이' 이야기를 들려줍니다. 학생은 도움을 준 사람과 장소를 찾아 말합니다.");
}

// 6. 간지 02
dividerSlide("02", "배움 몰입", "생활 사진 탐정", "8분", "천안·아산의 도시와 농촌 생활 사진을 살펴보아요.");

// 7. 활동 1 - 사진 관찰
{
  const s = baseSlide(C7, "활동 1 · 8분", "생활 사진 탐정",
    "천안·아산의 생활 사진을 보고, 주민들이 날마다 이용하는 공간을 찾아보아요.");
  const labels = ["버스 정류장", "도서관", "공원", "마을 공동공간"];
  const pw = (W - M * 2 - 0.36 * 3) / 4;
  labels.forEach((lb, i) => {
    photoSlot(s, lb, M + i * (pw + 0.36), 2.42, pw, 2.10);
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 4.80, w: W - M * 2, h: 1.10, rectRadius: 0.1,
    fill: { color: TINT }, line: { color: TINT, width: 0 },
  });
  const qs = ["어떤 장소인가요?", "누가 이용하나요?", "왜 필요한가요?"];
  const qw = (W - M * 2 - 0.68) / 3;
  qs.forEach((q, i) => {
    const x = M + 0.34 + i * qw;
    badge(s, i + 1, x, 5.13, 0.44, FOREST, WHITE, 13);
    s.addText(q, {
      x: x + 0.58, y: 5.13, w: qw - 0.78, h: 0.44,
      fontFace: FONT, fontSize: 16, bold: true, color: FOREST,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addText("사진은 공공누리 자료 또는 직접 촬영·제작 자료로 준비해 주세요.", {
    x: M, y: 6.14, w: W - M * 2, h: 0.36, fontFace: FONT, fontSize: 12, color: MUTED,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes("[교사용] 관광 명소가 아니라 주민의 일상 공간에 초점을 맞춥니다. 도시(천안)와 농촌(아산 외곽) 사진을 섞어 제시하면 비교가 쉬워집니다. 사진은 웹 검색 이미지 대신 공공누리 자료 또는 직접 촬영·제작 자료를 사용합니다.");
}

// 8. 활동 1 - 보는 눈
{
  const s = baseSlide(C7, "활동 1 · 8분", "사진을 보는 세 가지 눈",
    "같은 사진도 이렇게 나누어 보면 훨씬 많은 것이 보여요.", TINT2);
  const cards = [
    { h: "어떤 장소인지", b: "이름과 하는 일을 말해요.\n\n예) 책을 빌리고 읽는\n도서관이에요." },
    { h: "누가 이용하는지", b: "이용하는 사람을 떠올려요.\n\n예) 학생, 어르신,\n아기와 함께 온 부모님" },
    { h: "왜 필요한지", b: "없으면 어떤 점이 불편할지\n생각해요.\n\n예) 책을 사지 않아도\n읽을 수 있어요." },
  ];
  const cw = (W - M * 2 - 0.42 * 2) / 3;
  cards.forEach((c, i) => {
    const x = M + i * (cw + 0.42);
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 2.42, w: cw, h: 2.72, rectRadius: 0.12,
      fill: { color: WHITE }, line: { color: MOSS, width: 1 },
    });
    badge(s, i + 1, x + 0.34, 2.72, 0.58, FOREST, WHITE, 16);
    s.addText(c.h, {
      x: x + 1.06, y: 2.72, w: cw - 1.38, h: 0.58,
      fontFace: FONT, fontSize: 18, bold: true, color: INK,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(c.b, {
      x: x + 0.34, y: 3.46, w: cw - 0.68, h: 1.48,
      fontFace: FONT, fontSize: 13, color: MUTED, lineSpacing: 19,
      align: "left", valign: "top", margin: 0, isTextBox: true,
    });
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.42, w: W - M * 2, h: 0.86, rectRadius: 0.1,
    fill: { color: FOREST }, line: { color: FOREST, width: 0 },
  });
  s.addText("관광 명소가 아니라, 주민이 날마다 이용하는 생활 공간을 찾아요.", {
    x: M + 0.34, y: 5.42, w: W - M * 2 - 0.68, h: 0.86,
    fontFace: FONT, fontSize: 17, bold: true, color: WHITE,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes("세 가지 질문을 한 장소씩 순서대로 적용하게 하면 학생들이 관찰과 짐작을 구분하기 쉬워집니다.");
}

// 9. 간지 03
dividerSlide("03", "배움 몰입", "우리 마을 도움 연결지도", "15분", "사람 · 장소 · 도움받는 주민을 카드와 화살표로 연결해요.");

// 10. 활동 2 - 예시
{
  const s = baseSlide(C7, "활동 2 · 15분", "이렇게 연결해요",
    "사람에서 시작해 장소를 거쳐, 도움을 받는 주민까지 이어 보아요.");
  const rows = [
    { a: "사서", b: "도서관", c: "책을 찾는 주민", d: "도움 :  필요한 책을 찾아 주고, 조용히 읽을 자리를 마련해 주어요." },
    { a: "버스 운전기사", b: "버스와 정류장", c: "이동하는 주민", d: "도움 :  걸어서 가기 먼 곳까지 안전하게 데려다주어요." },
  ];
  rows.forEach((r, i) => {
    const y = 2.42 + i * 1.72;
    flowRow(s, [{ head: r.a }, { head: r.b }, { head: r.c }],
      y, 0.74, { w: 8.4, headSize: 16, fill: i === 0 ? TINT : TINT2 });
    s.addText(r.d, {
      x: M + 8.76, y, w: W - M - (M + 8.76), h: 0.74,
      fontFace: FONT, fontSize: 13, color: MUTED,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText("이용하는 주민이 함께 할 일 :  " + (i === 0 ? "반납일을 지켜요" : "차례를 지켜 타고 내려요"), {
      x: M + 0.1, y: y + 0.80, w: 8.3, h: 0.34,
      fontFace: FONT, fontSize: 13, color: FOREST, bold: true,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  question(s, "우리 마을에서는 또 누가, 어디에서, 누구를 도와줄까요?", M, 5.86, W - M * 2, 0.92);
  s.addNotes("예시 두 가지를 함께 읽은 뒤 모둠 활동으로 넘어갑니다. 화살표 방향이 '도움을 주는 사람 → 장소 → 도움받는 주민'임을 짚어 주세요.");
}

// 11. 활동 2 - 만들기
{
  const s = baseSlide(C7, "활동 2 · 15분", "우리 모둠 도움 연결지도",
    "카드와 화살표로 두 가지 이상의 관계를 만들어 보아요.", TINT2);
  for (let i = 0; i < 2; i++) {
    const y = 2.42 + i * 1.34;
    flowRow(s, [
      { head: "사람", body: "(       )" },
      { head: "장소", body: "(       )" },
      { head: "도움받는 주민", body: "(       )" },
    ], y, 1.06, { w: 7.9, headSize: 14, bodySize: 13, fill: WHITE });
    const ax = M + 8.26, aw = W - M - (M + 8.26);
    s.addText("어떤 도움인가요?", {
      x: ax, y: y - 0.02, w: aw, h: 0.30,
      fontFace: FONT, fontSize: 13, bold: true, color: FOREST,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addShape(pres.ShapeType.roundRect, {
      x: ax, y: y + 0.32, w: aw, h: 0.72, rectRadius: 0.08,
      fill: { color: WHITE }, line: { color: MOSS, width: 1, dashType: "dash" },
    });
  }
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.22, w: W - M * 2, h: 1.16, rectRadius: 0.1,
    fill: { color: TINT }, line: { color: TINT, width: 0 },
  });
  s.addText("이것만은 꼭 확인해요", {
    x: M + 0.34, y: 5.34, w: 3.0, h: 0.34,
    fontFace: FONT, fontSize: 13, bold: true, color: ACCENT,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText([
    { text: "관계를 두 가지 이상 만들었나요?", options: { breakLine: false } },
    { text: "        어떤 도움인지 적었나요?", options: { breakLine: false } },
    { text: "        이용하는 주민이 함께 할 일도 덧붙였나요?", options: {} },
  ], {
    x: M + 0.34, y: 5.70, w: W - M * 2 - 0.68, h: 0.50,
    fontFace: FONT, fontSize: 14, color: FOREST, bold: true,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes("모둠별로 카드를 놓고 화살표를 그립니다. 관계가 하나만 나온 모둠에는 '이 마을에서 또 도움을 주는 사람은 누구일까요?'라고 되물어 주세요.");
}

// 12. 간지 04
dividerSlide("04", "배움 몰입", "도움이 멈춘 하루", "8분", "도움이 사라졌을 때 누가 어떤 불편을 겪을지 생각해요.");

// 13. 활동 3
{
  const s = baseSlide(C7, "활동 3 · 8분", "도움이 멈춘 하루",
    "가상 상황 하나를 골라, 연결지도에서 불편을 겪을 사람을 찾아보아요.");
  const cases = [
    { t: "도서관을 하루 동안\n이용할 수 없다면?", n: "상황 A" },
    { t: "공원 시설이\n고장 났다면?", n: "상황 B" },
  ];
  const cw = 3.9;
  cases.forEach((c, i) => {
    const x = M + i * (cw + 0.42);
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 2.42, w: cw, h: 2.06, rectRadius: 0.12,
      fill: { color: i === 0 ? TINT : TINT2 }, line: { color: MOSS, width: 1 },
    });
    s.addText(c.n, {
      x: x + 0.30, y: 2.62, w: cw - 0.6, h: 0.32,
      fontFace: FONT, fontSize: 12, bold: true, color: ACCENT,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(c.t, {
      x: x + 0.30, y: 3.00, w: cw - 0.6, h: 1.20,
      fontFace: FONT, fontSize: 20, bold: true, color: INK, lineSpacing: 28,
      align: "left", valign: "top", margin: 0, isTextBox: true,
    });
  });
  const sx = M + (cw + 0.42) * 2;
  s.addShape(pres.ShapeType.roundRect, {
    x: sx, y: 2.42, w: W - M - sx, h: 2.06, rectRadius: 0.12,
    fill: { color: WHITE }, line: { color: FOREST, width: 1.5 },
  });
  s.addText("생각하는 순서", {
    x: sx + 0.30, y: 2.56, w: W - M - sx - 0.6, h: 0.34,
    fontFace: FONT, fontSize: 14, bold: true, color: FOREST,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  ["상황 카드 하나를 골라요.", "연결지도에서 불편을 겪을\n사람을 찾아요.", "어떤 불편인지 설명해요."].forEach((t, i) => {
    const y = 2.98 + i * 0.52;
    badge(s, i + 1, sx + 0.30, y, 0.30, MOSS, FOREST_DK, 11);
    s.addText(t, {
      x: sx + 0.70, y: y - 0.04, w: W - M - sx - 1.0, h: 0.40,
      fontFace: FONT, fontSize: 13, color: INK, lineSpacing: 15,
      align: "left", valign: "top", margin: 0, isTextBox: true,
    });
  });
  question(s, "같은 상황에서도 사람마다 필요한 도움이 다를까요?", M, 4.80, W - M * 2, 0.96);
  s.addText("한 사람이 아니라 여러 사람을 떠올리도록 도와주세요. 어린이, 어르신, 아기와 함께 온 부모님, 일하러 가는 주민 등", {
    x: M, y: 6.00, w: W - M * 2, h: 0.36, fontFace: FONT, fontSize: 12, color: MUTED,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes("불편을 '누가' 겪는지 반드시 사람을 지목해 말하게 합니다. 여기서 나온 답이 8차시 마을회의의 출발점이 됩니다.");
}

// 14. 학습정리
{
  const s = baseSlide(C7, "정리 · 4분", "연결 문장 완성하기",
    "연결지도를 보며 문장 두 개를 완성하고 짝에게 설명해요.", TINT2);
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 2.42, w: W - M * 2, h: 1.62, rectRadius: 0.14,
    fill: { color: WHITE }, line: { color: FOREST, width: 2 },
  });
  s.addText("○○ 은/는  ○○ 을/를  도와주어요.   그래서  ○○ 할 수 있어요.", {
    x: M + 0.4, y: 2.42, w: W - M * 2 - 0.8, h: 1.62,
    fontFace: FONT, fontSize: 28, bold: true, color: FOREST,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  const ex = [
    { h: "문장 ①", b: "사서는 책을 찾는 주민을 도와주어요.\n그래서 사고 싶지 않아도 책을 읽을 수 있어요." },
    { h: "문장 ②", b: "(     )은/는 (     )을/를 도와주어요.\n그래서 (     )할 수 있어요." },
  ];
  const cw2 = (W - M * 2 - 0.42) / 2;
  ex.forEach((e, i) => {
    const x = M + i * (cw2 + 0.42);
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 4.30, w: cw2, h: 1.36, rectRadius: 0.1,
      fill: { color: i === 0 ? TINT : WHITE }, line: { color: MOSS, width: 1 },
    });
    s.addText(e.h, {
      x: x + 0.30, y: 4.46, w: cw2 - 0.6, h: 0.32,
      fontFace: FONT, fontSize: 12, bold: true, color: ACCENT,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(e.b, {
      x: x + 0.30, y: 4.82, w: cw2 - 0.6, h: 0.72,
      fontFace: FONT, fontSize: 14, color: INK, lineSpacing: 20,
      align: "left", valign: "top", margin: 0, isTextBox: true,
    });
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 6.02, w: W - M * 2, h: 0.64, rectRadius: 0.1,
    fill: { color: FOREST }, line: { color: FOREST, width: 0 },
  });
  s.addText("완성한 도움 연결지도는 다음 차시 마을회의에 그대로 활용해요.", {
    x: M + 0.34, y: 6.02, w: W - M * 2 - 0.68, h: 0.64,
    fontFace: FONT, fontSize: 15, bold: true, color: WHITE,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes("두 문장을 모두 완성하게 하고, 짝에게 설명할 때 연결지도의 어느 부분을 보고 말하는지 가리키게 합니다.");
}

/* =========================================================
   8차시 | 모두가 함께 이용할 마을 공원 배치를 제안해요
   ========================================================= */
const C8 = "8차시";

// 15. 표지
coverSlide("8차시", "모두가 함께 이용할 마을 공원 배치를 제안해요")
  .addNotes("4회차(7~8차시) 중 뒤 차시이자 천안·아산 권역의 마지막 차시입니다. 7차시 도움 연결지도를 반드시 준비해 주세요.");

// 16. 목차
tocSlide(C8, [
  { no: "01", name: "같은 장소, 다른 바람", meta: "도입 · 5분" },
  { no: "02", name: "주민의 마음 읽기", meta: "활동 1 · 8분" },
  { no: "03", name: "어린이 마을회의", meta: "활동 2 · 15분" },
  { no: "04", name: "이웃 모둠 의견 듣기", meta: "활동 3 · 8분" },
], "정리   ·   우리들의 마을 제안 (4분)  —  제안 · 도움이 되는 사람 · 그렇게 정한 이유를 기록해요.");

// 17. 학습목표
goalSlide(C8, [
  "같은 공간을 이용하는 주민들의 서로 다른 바람을 찾아 말할 수 있다.",
  "서로 다른 요구를 두 가지 이상 반영하여 마을 공원 배치를 제안할 수 있다.",
  "이웃 모둠의 의견을 듣고 우리 제안을 한 가지 고쳐 까닭과 함께 설명할 수 있다.",
]);

// 18. 간지 01
dividerSlide("01", "배움 열기", "같은 장소, 다른 바람", "5분", "지난 시간 연결지도에서 공원과 관계된 사람들을 찾아요.");

// 19. 도입
{
  const s = baseSlide(C8, "도입 · 5분", "같은 장소, 다른 바람",
    "작은 공원 하나를 세 주민이 함께 이용한다면 어떤 일이 생길까요?");
  const people = [
    { h: "놀이 공간이\n필요한 주민", b: "“마음껏 뛰어놀\n자리가 있으면 좋겠어요.”" },
    { h: "조용한 쉼터가\n필요한 주민", b: "“잠시 앉아 쉴\n조용한 자리가 필요해요.”" },
    { h: "넓은 이동길이\n필요한 주민", b: "“유아차와 휠체어도\n지날 수 있어야 해요.”" },
  ];
  const cw = (W - M * 2 - 0.42 * 2) / 3;
  people.forEach((p, i) => {
    const x = M + i * (cw + 0.42);
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 2.42, w: cw, h: 2.14, rectRadius: 0.12,
      fill: { color: i === 1 ? TINT : TINT2 }, line: { color: MOSS, width: 1 },
    });
    s.addText(p.h, {
      x: x + 0.30, y: 2.64, w: cw - 0.6, h: 0.80,
      fontFace: FONT, fontSize: 18, bold: true, color: INK, lineSpacing: 25,
      align: "left", valign: "top", margin: 0, isTextBox: true,
    });
    s.addText(p.b, {
      x: x + 0.30, y: 3.52, w: cw - 0.6, h: 0.86,
      fontFace: FONT, fontSize: 14, color: MUTED, italic: true, lineSpacing: 20,
      align: "left", valign: "top", margin: 0, isTextBox: true,
    });
  });
  question(s, "공원이 있어도 모두 똑같이 만족할까요?", M, 4.88, W - M * 2, 0.96);
  s.addText("세 사람 모두 ‘공원이 필요하다’는 점은 같지만, 필요한 모습은 서로 달라요.", {
    x: M, y: 6.08, w: W - M * 2, h: 0.36, fontFace: FONT, fontSize: 13, color: MUTED,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes("7차시에 만든 도움 연결지도를 펼쳐 공원과 관계된 사람을 먼저 찾게 한 뒤, 가상 상황을 들려줍니다.");
}

// 20. 간지 02
dividerSlide("02", "배움 몰입", "주민의 마음 읽기", "8분", "역할 카드를 나누어 읽고, 서로의 바람을 말과 귀로 주고받아요.");

// 21. 활동 1
{
  const s = baseSlide(C8, "활동 1 · 8분", "주민의 마음 읽기",
    "모둠원이 역할 카드를 하나씩 맡아, 자기 바람을 까닭과 함께 말해요.");
  const roles = [
    "놀이 공간이\n필요한 주민",
    "조용한 쉼터가\n필요한 주민",
    "넓은 이동길이\n필요한 주민",
    "공원 관리\n담당자",
  ];
  const cw = (W - M * 2 - 0.34 * 3) / 4;
  roles.forEach((r, i) => {
    const x = M + i * (cw + 0.34);
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 2.42, w: cw, h: 1.40, rectRadius: 0.12,
      fill: { color: TINT2 }, line: { color: MOSS, width: 1 },
    });
    badge(s, i + 1, x + 0.26, 2.60, 0.42, FOREST, WHITE, 13);
    s.addText(r, {
      x: x + 0.26, y: 3.08, w: cw - 0.52, h: 0.64,
      fontFace: FONT, fontSize: 15, bold: true, color: INK, lineSpacing: 20,
      align: "left", valign: "top", margin: 0, isTextBox: true,
    });
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 4.14, w: W - M * 2, h: 1.20, rectRadius: 0.14,
    fill: { color: WHITE }, line: { color: FOREST, width: 2 },
  });
  s.addText("“저는 ○○ 이 필요해요.  왜냐하면 ○○ 이기 때문이에요.”", {
    x: M + 0.4, y: 4.14, w: W - M * 2 - 0.8, h: 1.20,
    fontFace: FONT, fontSize: 26, bold: true, color: FOREST,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.66, w: W - M * 2, h: 0.82, rectRadius: 0.1,
    fill: { color: TINT }, line: { color: TINT, width: 0 },
  });
  s.addText("듣는 친구는 “너는 ○○ 이 필요하구나.” 하고 들은 내용을 되짚어 말해 주어요.", {
    x: M + 0.34, y: 5.66, w: W - M * 2 - 0.68, h: 0.82,
    fontFace: FONT, fontSize: 16, bold: true, color: FOREST,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes("카드를 읽기만 하고 넘어가지 않도록, 반드시 '까닭'까지 말하게 합니다. 되짚어 말하기는 다음 활동에서 다른 사람의 요구를 반영하기 위한 준비입니다.");
}

// 22. 간지 03
dividerSlide("03", "배움 몰입", "어린이 마을회의", "15분", "공원 도면에 카드를 놓으며 모두를 위한 배치를 만들어요.");

// 23. 활동 2
{
  const s = baseSlide(C8, "활동 2 · 15분", "어린이 마을회의",
    "공원 도면 위에 공간 카드를 놓고, 입구와 각 공간을 길로 연결해요.");
  photoSlot(s, "공원 도면  ·  모둠별 1장", M, 2.42, 5.6, 3.22, "도면 준비");
  const items = [
    { h: "놀이 공간", b: "뛰어놀 수 있는 자리" },
    { h: "쉼터", b: "앉아 쉴 수 있는 자리" },
    { h: "이동길", b: "유아차·휠체어도 지나는 길" },
  ];
  const bx = M + 5.96;
  const bw = W - M - bx;
  s.addText("놓을 카드", {
    x: bx, y: 2.42, w: bw, h: 0.34,
    fontFace: FONT, fontSize: 13, bold: true, color: ACCENT,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  items.forEach((it, i) => {
    const y = 2.84 + i * 0.62;
    s.addShape(pres.ShapeType.roundRect, {
      x: bx, y, w: bw, h: 0.52, rectRadius: 0.1,
      fill: { color: TINT2 }, line: { color: MOSS, width: 1 },
    });
    s.addText(it.h, {
      x: bx + 0.24, y, w: 1.7, h: 0.52,
      fontFace: FONT, fontSize: 15, bold: true, color: FOREST,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(it.b, {
      x: bx + 1.98, y, w: bw - 2.22, h: 0.52,
      fontFace: FONT, fontSize: 13, color: MUTED,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: bx, y: 4.92, w: bw, h: 0.80, rectRadius: 0.1,
    fill: { color: FOREST }, line: { color: FOREST, width: 0 },
  });
  s.addText("서로 다른 주민의 요구를 두 가지 이상 반영해요.", {
    x: bx + 0.28, y: 4.92, w: bw - 0.56, h: 0.80,
    fontFace: FONT, fontSize: 16, bold: true, color: WHITE,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("입구와 각 공간을 길로 연결하고, 그렇게 정한 까닭을 함께 적어요.", {
    x: bx, y: 5.92, w: bw, h: 0.52,
    fontFace: FONT, fontSize: 13, color: MUTED, lineSpacing: 17,
    align: "left", valign: "top", margin: 0, isTextBox: true,
  });
  s.addText("카드와 도면은 모둠 책상 위에서만 주고받고, 자리에서 일어나 뛰지 않아요.", {
    x: M, y: 5.86, w: 5.6, h: 0.44,
    fontFace: FONT, fontSize: 12, color: MUTED,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes("[교사용] 모둠을 돌며 '누구의 바람이 아직 빠져 있나요?', '두 의견을 함께 반영할 수 있을까요?'라고 물어 주세요. 안전: 카드와 도면은 모둠 책상 앞에서만 주고받고, 이동을 최소화합니다.");
}

// 24. 활동 2 - 발문
{
  const s = baseSlide(C8, "활동 2 · 15분", "회의를 도와주는 질문",
    "회의가 한쪽으로 기울 때, 이 두 질문이 다시 길을 열어 주어요.", TINT2);
  const qcards = [
    { n: "질문 ①", t: "누구의 바람이\n아직 빠져 있나요?", b: "네 역할 카드를 모두 펼쳐 놓고\n빠진 사람이 있는지 확인해요." },
    { n: "질문 ②", t: "두 의견을 함께\n반영할 수 있을까요?", b: "하나를 버리는 대신, 자리를 나누거나\n시간을 나누는 방법을 찾아요." },
  ];
  const cw = (W - M * 2 - 0.42) / 2;
  qcards.forEach((q, i) => {
    const x = M + i * (cw + 0.42);
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 2.42, w: cw, h: 2.86, rectRadius: 0.14,
      fill: { color: i === 0 ? FOREST : WHITE }, line: { color: FOREST, width: i === 0 ? 0 : 2 },
    });
    s.addText(q.n, {
      x: x + 0.38, y: 2.66, w: cw - 0.76, h: 0.34,
      fontFace: FONT, fontSize: 12, bold: true, color: i === 0 ? MOSS : ACCENT,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(q.t, {
      x: x + 0.38, y: 3.06, w: cw - 0.76, h: 1.10,
      fontFace: FONT, fontSize: 25, bold: true, color: i === 0 ? WHITE : INK, lineSpacing: 34,
      align: "left", valign: "top", margin: 0, isTextBox: true,
    });
    s.addText(q.b, {
      x: x + 0.38, y: 4.30, w: cw - 0.76, h: 0.78,
      fontFace: FONT, fontSize: 14, color: i === 0 ? "C9D8C2" : MUTED, lineSpacing: 20,
      align: "left", valign: "top", margin: 0, isTextBox: true,
    });
  });
  s.addText("교사용 참고  ·  모둠을 돌며 한 번씩 건네 주세요. 답을 정해 주지 않고, 모둠이 스스로 고치게 합니다.", {
    x: M, y: 5.52, w: W - M * 2, h: 0.40,
    fontFace: FONT, fontSize: 13, color: MUTED,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes("두 질문 모두 '누구를 위한 공원인가'로 돌아오게 하는 장치입니다. 시간을 나누는 해결(예: 오전은 조용한 시간)도 좋은 답으로 인정해 주세요.");
}

// 25. 간지 04
dividerSlide("04", "배움 몰입", "이웃 모둠 의견 듣기", "8분", "다른 모둠의 눈으로 우리 제안을 다시 살펴보아요.");

// 26. 활동 3
{
  const s = baseSlide(C8, "활동 3 · 8분", "이웃 모둠 의견 듣기",
    "설명하고, 듣고, 한 가지를 고치는 순서로 진행해요.");
  flowRow(s, [
    { head: "설명하기", body: "이웃 모둠에 우리 공원의\n배치와 그렇게 정한 이유를\n설명해요." },
    { head: "의견 주기", body: "듣는 모둠은 좋은 점 한 가지,\n아직 불편할 수 있는 사람\n한 가지를 이야기해요." },
    { head: "고치기", body: "의견을 받은 모둠은\n배치나 이용 약속을\n한 가지 수정해요." },
  ], 2.42, 2.16, { headSize: 20, bodySize: 14 });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 4.86, w: W - M * 2, h: 1.02, rectRadius: 0.1,
    fill: { color: TINT }, line: { color: TINT, width: 0 },
  });
  s.addText("의견을 말할 때는 배치를 고른 ‘사람’을 말해요.  예) “유아차를 미는 주민은 이 길이 좁아 지나가기 어려울 것 같아요.”", {
    x: M + 0.34, y: 4.86, w: W - M * 2 - 0.68, h: 1.02,
    fontFace: FONT, fontSize: 15, color: FOREST, bold: true, lineSpacing: 22,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("고친 내용은 지우지 말고 옆에 덧붙여 적어요. 처음 생각과 고친 생각이 함께 보이도록 남겨 두세요.", {
    x: M, y: 6.06, w: W - M * 2, h: 0.36,
    fontFace: FONT, fontSize: 13, color: MUTED,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes("모둠 전체가 이동하지 않고, 설명할 한 명만 이웃 모둠으로 가거나 도면을 돌려 보게 하면 이동이 줄어 안전합니다.");
}

// 27. 학습정리
{
  const s = baseSlide(C8, "정리 · 4분", "우리들의 마을 제안",
    "완성한 제안을 세 칸에 기록하고 발표해요.", TINT2);
  const cols = [
    { h: "우리의 제안", b: "공원을 어떻게\n배치했나요?" },
    { h: "도움이 되는 사람", b: "이 배치로 누가\n편해지나요?" },
    { h: "그렇게 정한 이유", b: "왜 이렇게\n정했나요?" },
  ];
  const cw = (W - M * 2 - 0.42 * 2) / 3;
  cols.forEach((c, i) => {
    const x = M + i * (cw + 0.42);
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 2.42, w: cw, h: 2.30, rectRadius: 0.12,
      fill: { color: WHITE }, line: { color: FOREST, width: 1.5 },
    });
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 2.42, w: cw, h: 0.62, rectRadius: 0.12,
      fill: { color: FOREST }, line: { color: FOREST, width: 0 },
    });
    s.addText(c.h, {
      x: x + 0.2, y: 2.42, w: cw - 0.4, h: 0.62,
      fontFace: FONT, fontSize: 17, bold: true, color: WHITE,
      align: "center", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(c.b, {
      x: x + 0.28, y: 3.18, w: cw - 0.56, h: 0.70,
      fontFace: FONT, fontSize: 13, color: MUTED, lineSpacing: 18,
      align: "left", valign: "top", margin: 0, isTextBox: true,
    });
  });
  s.addText("마지막으로 확인해요", {
    x: M, y: 5.02, w: 3.4, h: 0.34,
    fontFace: FONT, fontSize: 13, bold: true, color: ACCENT,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  const checks = [
    "서로 다른 두 가지 이상의 요구가 반영되었나요?",
    "어린이가 실천할 일과 어른 · 기관에 요청할 일을 구분했나요?",
  ];
  checks.forEach((c, i) => {
    const y = 5.44 + i * 0.62;
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y, w: W - M * 2, h: 0.52, rectRadius: 0.1,
      fill: { color: i === 0 ? TINT : WHITE }, line: { color: MOSS, width: 1 },
    });
    badge(s, i + 1, M + 0.22, y + 0.10, 0.32, MOSS, FOREST_DK, 11);
    s.addText(c, {
      x: M + 0.68, y, w: W - M * 2 - 1.0, h: 0.52,
      fontFace: FONT, fontSize: 15, color: INK,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
  });
  s.addNotes("'어린이가 실천할 일'과 '어른·기관에 요청할 일'을 구분하는 것이 이 차시의 마무리 지점입니다. 예: 쓰레기를 되가져가기(어린이) / 길 폭을 넓혀 달라고 요청하기(기관).");
}

// 28. 마무리
{
  const s = pres.addSlide();
  darkBase(s);
  s.addText("탐험을 마치며", {
    x: M, y: 1.10, w: 9.0, h: 0.42, fontFace: FONT, fontSize: 15, color: MOSS,
    align: "left", valign: "middle", margin: 0, charSpacing: 2, isTextBox: true,
  });
  s.addText("우리 고장을 알고,\n함께 쓰는 마을을 제안하기까지", {
    x: M, y: 1.60, w: 10.8, h: 1.54, fontFace: FONT, fontSize: 38, bold: true, color: WHITE,
    lineSpacing: 48, align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  const journey = [
    { n: "1~2차시", t: "역사와 사람들", b: "유관순 · 이순신" },
    { n: "3~4차시", t: "여행지", b: "천안삼거리공원 · 신정호" },
    { n: "5~6차시", t: "특산물", b: "호두과자 · 배" },
    { n: "7~8차시", t: "생활과 공동체", b: "도움 연결지도 · 마을 공원" },
  ];
  const cw = (W - M * 2 - 0.34 * 3) / 4;
  journey.forEach((j, i) => {
    const x = M + i * (cw + 0.34);
    const last = i === 3;
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 3.66, w: cw, h: 1.86, rectRadius: 0.12,
      fill: { color: last ? MOSS : FOREST }, line: { color: last ? MOSS : FOREST, width: 0 },
    });
    s.addText(j.n, {
      x: x + 0.26, y: 3.86, w: cw - 0.52, h: 0.32,
      fontFace: FONT, fontSize: 12, bold: true, color: last ? FOREST_DK : MOSS,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(j.t, {
      x: x + 0.26, y: 4.22, w: cw - 0.52, h: 0.46,
      fontFace: FONT, fontSize: 18, bold: true, color: last ? FOREST_DK : WHITE,
      align: "left", valign: "middle", margin: 0, isTextBox: true,
    });
    s.addText(j.b, {
      x: x + 0.26, y: 4.72, w: cw - 0.52, h: 0.62,
      fontFace: FONT, fontSize: 13, color: last ? "2C5F2D" : "C9D8C2", lineSpacing: 17,
      align: "left", valign: "top", margin: 0, isTextBox: true,
    });
  });
  s.addText("2026 앵커 늘봄 프로그램  ·  천안 · 아산 〈호서대학교〉  ·  우리고장 탐험놀이", {
    x: M, y: 6.34, w: 11.0, h: 0.36, fontFace: FONT, fontSize: 13, color: "8FA391",
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addNotes("천안·아산 권역 8차시를 마무리하는 슬라이드입니다. 네 가지 공통 주제를 되짚으며 프로그램 전체를 정리합니다.");
}

pres.writeFile({ fileName: process.argv[2] || "out.pptx" }).then((f) => console.log("saved:", f));
