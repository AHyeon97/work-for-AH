// 활동지 / 해설지 렌더러
const L = require("./sheet_lib.js");
const LIMIT = 10.84;
function chk(tag, y) {
  if (y > LIMIT) console.log(`  ⚠ 넘침 ${tag}: y=${y.toFixed(2)} (한계 ${LIMIT})`);
  return y;
}

function area(pres, s, a, y) {
  switch (a.kind) {
    case "topBoxes": return L.topBoxes(pres, s, a.labels, y, a.h, a.lines, a.fills);
    case "labelCards": return L.labelCards(pres, s, a.labels, y, a.h, a.cols);
    case "chips": return L.chips(pres, s, a.items, y);
    case "lines": return L.writeLines(pres, s, y - 0.30, a.n, a.step);
    case "check": return L.checkRow(pres, s, a.label, a.items, y);
    case "pledge": return L.pledge(pres, s, a.lead, a.text, y);
    case "gap": return y + (a.h || 0.2);
    default: return y;
  }
}

function renderItem(pres, s, it, y) {
  y = L.itemHead(pres, s, it.no, it.title, y);
  if (it.guide) y = L.guideBox(pres, s, it.guide, y, it.guideH);
  (it.areas || []).forEach((a) => { y = area(pres, s, a, y); });
  return y;
}

/* ---------------- 활동지 ---------------- */
function buildWorksheet(d, out) {
  const pres = L.newSheet(d.title + " 활동지");
  const foot = `${d.region} 우리충남 탐험놀이 · ${d.round}`;

  let s = L.page(pres, { foot });
  L.header(pres, s, { head: foot, title: d.title, badge: d.badge });
  let y = 2.13;
  d.items.slice(0, 3).forEach((it) => { y = renderItem(pres, s, it, y); });
  chk(d.file + " 활동지 1쪽", y);

  s = L.page(pres, { foot: null });
  y = 0.50;
  d.items.slice(3).forEach((it) => { y = renderItem(pres, s, it, y); });
  chk(d.file + " 활동지 2쪽", y);

  return pres.writeFile({ fileName: out });
}

/* ---------------- 해설지 ---------------- */
function buildKey(d, out) {
  const pres = L.newSheet(d.title + " 활동지 해설지");
  const foot = `${d.region} 우리충남 탐험놀이 · ${d.round} · 교사용`;

  let s = L.page(pres, { foot });
  s.addText("교사용 해설지 · " + d.shortTitle, {
    x: L.M, y: 0.46, w: L.CW, h: 0.30, fontFace: L.FONT_T, fontSize: 14, bold: true, color: L.INK,
    align: "left", valign: "middle", margin: 0, isTextBox: true,
  });
  let y = L.warnBox(pres, s, "학생 배부본에는 넣지 않습니다. 아래 예시 답안은 하나의 보기일 뿐이니, 자료에서 근거를 찾아 쓴 답이면 모두 인정합니다.", 0.86);

  const put = (it) => {
    y = L.itemHead(pres, s, it.no, it.title, y);
    if (it.key && it.key.length) y = L.keyBox(pres, s, it.key, y);
    (it.tips || []).forEach((t) => { y = L.tipLine(pres, s, t[0], t[1], y, t[2]); });
    y += 0.16;
  };
  d.items.slice(0, 3).forEach(put);
  chk(d.file + " 해설지 1쪽", y);

  s = L.page(pres, { foot: foot + " · 2쪽" });
  y = 0.50;
  d.items.slice(3).forEach(put);
  chk(d.file + " 해설지 2쪽", y);

  if (d.wrap) {
    y = Math.max(y, 0.5);
    L.tipLine(pres, s, "학습정리", d.wrap, y, "2E7D4F");
  }
  return pres.writeFile({ fileName: out });
}

module.exports = { buildWorksheet, buildKey };
