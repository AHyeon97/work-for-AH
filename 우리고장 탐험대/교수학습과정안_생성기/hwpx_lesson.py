# -*- coding: utf-8 -*-
"""교수학습과정안(HWPX) 생성기.

기존 '2. 회차별 교수학습과정안(5~8차시)' 파일을 양식으로 삼아,
표 구조와 문단 서식을 그대로 둔 채 내용만 바꾼다.

문단 서식 (원본에서 확인):
  ■ 제목   paraPr 24 / run charPr 18 (진한 파랑 #1F4E79)
  · 강사   paraPr 24 / run charPr 8
  ▷ 학생   paraPr 24 / run charPr 9("▷ ") + charPr 8(본문)
  ※ 유의   paraPr 24 / run charPr 9("※ ") + charPr 8(본문)
  “발문”   paraPr 29 / run charPr 8   (파란 상자)
"""
import copy, os, re, shutil, zipfile
from lxml import etree

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"


def q(name):
    return "{%s}%s" % (HP, name)


def para_text(p):
    return "".join(t.text or "" for t in p.iter(q("t")))


class LessonPlan:
    def __init__(self, src_path, workdir):
        self.work = workdir
        if os.path.isdir(workdir):
            shutil.rmtree(workdir)
        os.makedirs(workdir)
        with zipfile.ZipFile(src_path) as z:
            z.extractall(workdir)
            self.names = z.namelist()
        self.sec_path = os.path.join(workdir, "Contents", "section0.xml")
        self.tree = etree.parse(self.sec_path)
        self.root = self.tree.getroot()
        self.tables = list(self.root.iter(q("tbl")))
        self._collect_templates()

    # ---------- 서식 템플릿 수집 ----------
    def _collect_templates(self):
        self.tpl = {}
        for p in self.root.iter(q("p")):
            s = para_text(p).strip()
            if not s:
                continue
            runs = [r for r in p.findall(q("run")) if r.find(q("t")) is not None]
            if not runs:
                continue
            ppr = p.get("paraPrIDRef")
            cps = [r.get("charPrIDRef") for r in runs]
            key = None
            if ppr == "29":
                key = "q"
            elif ppr == "24" and cps == ["18"]:
                key = "title"
            elif ppr == "24" and cps == ["8"]:
                key = "t"
            elif ppr == "24" and cps == ["9", "8"]:
                key = "s" if s.startswith("▷") else "n"
            if key and key not in self.tpl:
                self.tpl[key] = copy.deepcopy(p)
        # ※ 템플릿이 범례 문단밖에 없을 때 ▷ 템플릿으로 대체
        if "n" not in self.tpl and "s" in self.tpl:
            self.tpl["n"] = copy.deepcopy(self.tpl["s"])
        missing = {"title", "t", "s", "n", "q"} - set(self.tpl)
        if missing:
            raise RuntimeError("문단 템플릿을 찾지 못함: %s" % missing)

    # ---------- 문단 만들기 ----------
    @staticmethod
    def _set_runs(p, parts):
        """parts: [(charPrIDRef, text), ...] 순서대로 run 의 hp:t 를 채운다."""
        runs = [r for r in p.findall(q("run")) if r.find(q("t")) is not None]
        for i, (_, text) in enumerate(parts):
            if i < len(runs):
                runs[i].find(q("t")).text = text
            else:  # 런이 모자라면 마지막 런을 복제
                nr = copy.deepcopy(runs[-1])
                nr.find(q("t")).text = text
                runs[-1].addnext(nr)
                runs.append(nr)
        for extra in runs[len(parts):]:
            extra.getparent().remove(extra)

    def make_para(self, kind, text):
        p = copy.deepcopy(self.tpl[kind])
        if kind == "title":
            self._set_runs(p, [("18", "■ " + text)])
        elif kind == "t":
            self._set_runs(p, [("8", "· " + text)])
        elif kind == "s":
            self._set_runs(p, [("9", "▷ "), ("8", text)])
        elif kind == "n":
            self._set_runs(p, [("9", "※ "), ("8", text)])
        elif kind == "q":
            self._set_runs(p, [("8", "“" + text + "”")])
        elif kind == "plain":
            p = copy.deepcopy(self.tpl["t"])
            self._set_runs(p, [("8", text)])
        else:
            raise ValueError(kind)
        return p

    # ---------- 셀 채우기 ----------
    def _cell(self, ti, ri, ci):
        return self.tables[ti].findall(q("tr"))[ri].findall(q("tc"))[ci]

    def fill_cell_plain(self, ti, ri, ci, texts):
        """셀의 기존 문단 서식을 그대로 쓰면서 줄 수만 맞춘다."""
        tc = self._cell(ti, ri, ci)
        sub = tc.find(q("subList"))
        ps = sub.findall(q("p"))
        base = copy.deepcopy(ps[0])
        for p in ps:
            sub.remove(p)
        for text in texts:
            p = copy.deepcopy(base)
            runs = [r for r in p.findall(q("run")) if r.find(q("t")) is not None]
            if not runs:
                continue
            runs[0].find(q("t")).text = text
            for extra in runs[1:]:
                extra.getparent().remove(extra)
            sub.append(p)

    def fill_cell_rich(self, ti, ri, ci, lines):
        """lines: [(kind, text), ...]  kind in title/t/s/n/q/plain"""
        tc = self._cell(ti, ri, ci)
        sub = tc.find(q("subList"))
        for p in sub.findall(q("p")):
            sub.remove(p)
        for kind, text in lines:
            sub.append(self.make_para(kind, text))

    # ---------- 표 바깥 제목 문단 ----------
    def replace_heading(self, old_startswith, new_text, nth=0):
        hit = 0
        for p in self.root.iter(q("p")):
            if p.getparent().tag == q("subList"):
                continue  # 표 안 문단은 건너뜀
            s = para_text(p).strip()
            if s.startswith(old_startswith):
                if hit == nth:
                    runs = [r for r in p.findall(q("run")) if r.find(q("t")) is not None]
                    if runs:
                        runs[0].find(q("t")).text = new_text
                        for extra in runs[1:]:
                            extra.getparent().remove(extra)
                    return True
                hit += 1
        return False

    # ---------- 저장 ----------
    def save(self, out_path):
        self.tree.write(self.sec_path, encoding="UTF-8",
                        xml_declaration=True, standalone=True)
        if os.path.exists(out_path):
            os.remove(out_path)
        # mimetype 은 첫 항목 · 무압축이어야 한다
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            mt = os.path.join(self.work, "mimetype")
            if os.path.exists(mt):
                z.write(mt, "mimetype", compress_type=zipfile.ZIP_STORED)
            for name in self.names:
                if name == "mimetype" or name.endswith("/"):
                    continue
                full = os.path.join(self.work, name)
                if os.path.exists(full):
                    z.write(full, name)
        return out_path


# ---------- 한 회차 전체를 채우는 도우미 ----------
def fill_round(lp, slot, info, rows):
    """slot: 0 = 앞 회차, 1 = 뒤 회차
    info: dict(chasi, topic, goals[2], subject, tools, scores[5])
    rows: [(stage_lines[], content_lines[], bigo[]), ...] 7개
          stage_lines 는 단계 칸의 줄 목록(예: ['사전준비','3′'])
    """
    bi, ai = slot * 2, slot * 2 + 1          # 기본정보 표, 교수학습활동 표
    lp.fill_cell_plain(bi, 0, 3, [info["chasi"]])
    lp.fill_cell_plain(bi, 1, 1, ["▪ " + info["topic"]])
    lp.fill_cell_plain(bi, 2, 1, ["▪ " + g for g in info["goals"]])
    lp.fill_cell_plain(bi, 3, 1, [info["subject"]])
    lp.fill_cell_plain(bi, 4, 1, [info["tools"]])
    for ci, v in enumerate(info["scores"]):
        lp.fill_cell_plain(bi, 7, ci, [str(v)])
    # 교수·학습활동: r1~r7. 단계 칸 위치는 행마다 다르다(병합 때문).
    stage_ci = {1: 1, 2: 0, 3: 0, 4: 1, 5: 0, 6: 0, 7: 1}
    for n, (stage, content, bigo) in enumerate(rows, start=1):
        sc = stage_ci[n]
        lp.fill_cell_plain(ai, n, sc, stage)
        lp.fill_cell_rich(ai, n, sc + 1, content)
        lp.fill_cell_plain(ai, n, sc + 2, bigo)
