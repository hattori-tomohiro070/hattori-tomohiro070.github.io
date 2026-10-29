"""website_content.docx を読み込んで data.js を生成する。

docx の構成: 「見出し名 [key]」という見出しの直後に表を1つ置く。
表の1行目は列名で、列名で値を取り出すため列の順番は入れ替えてもよい。
空行は無視し、各セクションは日付の新しい順に自動で並べ替える。

使い方: python3 scripts/build_data.py [入力docx] [出力js]
"""
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).resolve().parent.parent
SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "website_content.docx"
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "data.js"

PROFILE_ROWS = {"名前": "name", "所属・身分": "position", "研究室": "lab", "自己紹介": "bio",
                "研究キーワード": "interests", "強調表示する自分の名前": "self"}
LINK_ROWS = {"email": "email", "orcid": "orcid", "google scholar": "scholar",
             "researchmap": "researchmap", "github": "github", "cv": "cv"}
TALK_TYPES = {"口頭": "oral", "口頭発表": "oral", "oral": "oral", "ポスター": "poster", "ポスター発表": "poster",
              "poster": "poster", "招待": "invited", "招待講演": "invited", "invited": "invited"}
PRESENT = {"", "現在", "present", "至現在", "現在に至る"}


def sections(doc):
    """見出し [key] → 直後の表（行ごとの dict のリスト）"""
    out, key = {}, None
    for el in doc.element.body.iterchildren():
        tag = el.tag.rsplit("}", 1)[-1]
        if tag == "p":
            m = re.search(r"\[(\w+)\]\s*$", Paragraph(el, doc).text.strip())
            if m:
                key = m.group(1)
        elif tag == "tbl" and key:
            rows = [[c.text.strip() for c in r.cells] for r in Table(el, doc).rows]
            head = rows[0]
            out[key] = [dict(zip(head, r)) for r in rows[1:] if any(r)]
            key = None
    return out


def date(s, where):
    """'2024/4/1', '2024年4月', '2024-04' などを 'YYYY-MM-DD' / 'YYYY-MM' / 'YYYY' に正規化"""
    m = re.fullmatch(r"(\d{4})(?:\D+(\d{1,2}))?(?:\D+(\d{1,2}))?\D*", s.strip())
    if not m:
        sys.exit(f"日付の形式が読めません: 「{s}」（{where}）")
    return "-".join(f"{int(p):02d}" if i else p for i, p in enumerate(m.groups()) if p)


def sort_key(d):
    # 年のみの日付はその年の先頭（新しい側）に並べる
    return d if len(d) > 4 else d + "-99"


def authors(s):
    return [a.strip() for a in re.split(r"[,，、;；]", s) if a.strip()]


def bi(row, col):
    return {"en": row.get(f"{col} (EN)", ""), "ja": row.get(f"{col} (JA)", "")}


def build(sec):
    prof = {PROFILE_ROWS[r["項目"]]: r for r in sec["profile"] if r["項目"] in PROFILE_ROWS}
    profile = {k: {"en": prof[k]["English"], "ja": prof[k]["日本語"]} for k in ("name", "position", "lab", "bio")}
    profile["interests"] = {lang: [x.strip() for x in re.split(r"[,，、]", prof["interests"][col]) if x.strip()]
                            for lang, col in (("en", "English"), ("ja", "日本語"))}
    profile["links"] = {LINK_ROWS[r["項目"].lower()]: r["URL"] for r in sec["links"] if r["項目"].lower() in LINK_ROWS}
    self_names = [n for col in ("English", "日本語") for n in authors(prof["self"][col])]

    def career(key):
        items = [{"start": date(r["開始"], f"{key}: {r['内容 (JA)']}"),
                  "end": "" if r["終了"].lower() in PRESENT else date(r["終了"], key),
                  "title": bi(r, "内容"), "org": bi(r, "所属")} for r in sec.get(key, [])]
        return sorted(items, key=lambda x: sort_key(x["start"]), reverse=True)

    def pubs(key):
        items = []
        for r in sec.get(key, []):
            p = {"authors": authors(r["著者"]), "title": r["タイトル"], "venue": r["掲載誌"],
                 "detail": r["巻・号・ページ"], "year": int(date(r["年"], f"{key}: {r['タイトル']}")[:4])}
            doi = re.sub(r"^(https?://)?(dx\.)?doi\.org/", "", r.get("DOI", ""))
            arxiv = re.sub(r"^(https?://arxiv\.org/abs/|arxiv:)", "", r.get("arXiv", ""), flags=re.I)
            if doi: p["doi"] = doi
            if arxiv: p["arxiv"] = arxiv
            items.append(p)
        return sorted(items, key=lambda x: x["year"], reverse=True)

    def talks(key):
        items = []
        for r in sec.get(key, []):
            kind = TALK_TYPES.get(r["種別"].lower())
            if not kind:
                sys.exit(f"種別は「口頭」「ポスター」「招待講演」のいずれかにしてください: 「{r['種別']}」（{r['タイトル']}）")
            t = {"date": date(r["日付"], f"{key}: {r['タイトル']}"), "type": kind,
                 "authors": authors(r["著者"]), "title": r["タイトル"], "venue": r["会議名"]}
            note = bi(r, "備考")
            if note["en"] or note["ja"]:
                t["note"] = {"en": note["en"] or note["ja"], "ja": note["ja"] or note["en"]}
            items.append(t)
        return sorted(items, key=lambda x: sort_key(x["date"]), reverse=True)

    awards = sorted(({"date": date(r["日付"], f"awards: {r['内容 (JA)']}"), "title": bi(r, "内容")}
                     for r in sec.get("awards", [])), key=lambda x: sort_key(x["date"]), reverse=True)

    site = {"profile": profile,
            **{k: career(k) for k in ("education", "teaching", "experience")},
            **{k: pubs(k) for k in ("journals", "proceedings", "preprints")},
            **{k: talks(k) for k in ("international", "domestic")},
            "awards": awards}
    return self_names, site


def main():
    sec = sections(Document(SRC))
    for k in ("profile", "links"):
        if k not in sec:
            sys.exit(f"見出し [{k}] とその表が見つかりません")
    self_names, site = build(sec)
    OUT.write_text(
        f"// このファイルは scripts/build_data.py が {SRC.name} から自動生成する。直接編集しないこと。\n"
        f"window.SELF_NAMES = {json.dumps(self_names, ensure_ascii=False)};\n"
        f"window.SITE = {json.dumps(site, ensure_ascii=False, indent=1)};\n", encoding="utf-8")
    counts = ", ".join(f"{k}: {len(v)}" for k, v in site.items() if isinstance(v, list))
    print(f"{OUT.name} を生成しました（{counts}）")


if __name__ == "__main__":
    main()
