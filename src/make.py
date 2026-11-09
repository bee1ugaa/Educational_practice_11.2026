"""Сборка отчёта с подбором номеров страниц оглавления и разбиения таблиц.

Запуск: python3 src/make.py
Требуются: Node.js с пакетом docx, LibreOffice (soffice), poppler-utils (pdftotext).
Результат: report/Практическая_работа_1.docx и report/Практическая_работа_1.pdf
"""
import json, os, re, shutil, subprocess, sys, tempfile, pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "report"
NAME = "Практическая_работа_1"


def norm(s):
    return re.sub(r"\s+", " ", s.replace("­", "")).strip()


def build(tmp, layout):
    lay = tmp / "layout.json"
    lay.write_text(json.dumps(layout, ensure_ascii=False), encoding="utf-8")
    docx = tmp / f"{NAME}.docx"
    env = dict(os.environ, MARKERS=str(tmp / "markers.json"))
    subprocess.run(["node", str(ROOT / "src/build.js"), str(docx), str(lay)], check=True, env=env,
                   stdout=subprocess.DEVNULL)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp), str(docx)],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    pdf = tmp / f"{NAME}.pdf"
    n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True).stdout).group(1))
    pages = []
    for i in range(1, n + 1):
        t = subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), str(pdf), "-"], capture_output=True, text=True).stdout
        pages.append(t)
    markers = json.loads((tmp / "markers.json").read_text(encoding="utf-8"))
    return docx, pdf, pages, markers


def analyse(pages, markers):
    # Номера страниц для оглавления: строка, начинающаяся с заголовка, после страницы оглавления
    # первая страница с текстом «Оглавление» – дальше ищутся заголовки
    prev = next(i for i, p in enumerate(pages, 1) if 'Оглавление' in p) + 1
    result = {}
    for title in markers["toc"]:
        key = norm(title)[:30]
        for i in range(prev, len(pages) + 1):
            lines = [norm(l) for l in pages[i - 1].splitlines() if l.strip() and "...." not in l]
            if any(l.startswith(key) for l in lines):
                result[title] = i
                prev = i
                break
        else:
            sys.exit(f"Заголовок не найден: {title}")
    layout = {"pages": result}
    # Распределение строк таблиц по страницам
    flat = [norm(p) for p in pages]
    for num, cells in markers["tables"].items():
        # поиск начинается со страницы, на которой находится название таблицы
        start = next(i for i, p in enumerate(flat, 1) if f"Таблица {num} –" in p)
        row_pages = []
        for c in cells:
            key = norm(c)[:22]
            prev_p = row_pages[-1] if row_pages else start
            found = [i for i, p in enumerate(flat, 1) if key in p and i >= prev_p]
            if not found:
                sys.exit(f"Строка таблицы {num} не найдена: {key}")
            row_pages.append(found[0])
        chunks = []
        for i, p in enumerate(row_pages):
            if i and p == row_pages[i - 1]:
                chunks[-1] += 1
            else:
                chunks.append(1)
        layout[f"table{num}"] = chunks
    return layout


def main():
    tmp = pathlib.Path(tempfile.mkdtemp())
    layout = {}
    for it in range(6):
        docx, pdf, pages, markers = build(tmp, layout)
        new = analyse(pages, markers)
        print(f"проход {it + 1}: {len(pages)} стр., {new}")
        if new == layout:
            break
        layout = new
    else:
        sys.exit("Разметка не стабилизировалась")
    # Попытка объединить соседние фрагменты таблиц, если они помещаются на одной странице
    for key in [k for k in layout if k.startswith("table")]:
        i = 0
        while i < len(layout[key]) - 1:
            ch = layout[key]
            cand = dict(layout, **{key: ch[:i] + [ch[i] + ch[i + 1]] + ch[i + 2:]})
            d, p, pages, markers = build(tmp, cand)
            got = analyse(pages, markers)
            if got[key] == cand[key]:
                cand["pages"] = got["pages"]
                d, p, pages, markers = build(tmp, cand)
                if analyse(pages, markers) == cand:
                    layout, docx, pdf = cand, d, p
                    print(f"объединение {key}: {layout[key]}")
                    continue
            i += 1
    docx, pdf, pages, markers = build(tmp, layout)
    assert analyse(pages, markers) == layout, "итоговая разметка не совпала"
    print(f"итог: {len(pages)} стр., {layout}")
    OUT_DIR.mkdir(exist_ok=True)
    shutil.copy(docx, OUT_DIR / docx.name)
    shutil.copy(pdf, OUT_DIR / pdf.name)
    print("готово:", OUT_DIR)


if __name__ == "__main__":
    main()
