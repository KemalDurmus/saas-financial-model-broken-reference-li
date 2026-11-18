"""Excel (.xlsx) finansal model denetimi: saf Python, Flask'a bağımlı değil."""
import csv
import datetime
import io
import re
import zipfile
from collections import Counter, defaultdict
from urllib.parse import quote

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.workbook.defined_name import DefinedName

ERROR_CODES = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NULL!", "#NUM!")
MAX_CELLS = 400000
MAX_ISSUES = 5000
MAX_GRAPH_NODES = 30000

T_EXT = "Dış bağlantı"
T_BROKEN = "Bozuk formül"
T_ERRVAL = "Hata değeri"
T_CIRC = "Dairesel başvuru"
T_NAME = "Geçersiz adlandırılmış alan"

SEVERITY = {T_EXT: "Yüksek", T_BROKEN: "Yüksek", T_CIRC: "Yüksek", T_NAME: "Yüksek", T_ERRVAL: "Orta"}

STRING_RE = re.compile(r'"(?:[^"]|"")*"')
EXT_RE = re.compile(r"\[([^\[\]]+)\][^\[\]!]*!")
EXT_REF_RE = re.compile(
    r"(?:'[^']*\[[^\]]+\][^']*'|\[[^\]]+\][\w\.]*)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?"
)
ERR_LIT_RE = re.compile(r"#(?:REF!|DIV/0!|VALUE!|NAME\?|N/A|NULL!|NUM!)")
REF_RE = re.compile(
    r"(?<![\w.\]])(?:(?:'((?:[^']|'')+)'|([A-Za-z_][\w\.]*))!)?"
    r"\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?(?![\w(])"
)


class AuditError(Exception):
    """Kullanıcıya gösterilecek anlaşılır hata."""


def _formula_text(cell):
    v = cell.value
    if isinstance(v, str):
        return v if v.startswith("=") else None
    text = getattr(v, "text", None)
    if isinstance(text, str):
        return text
    return None


def _external_targets(wb):
    targets = []
    try:
        for link in getattr(wb, "_external_links", []) or []:
            try:
                targets.append(str(link.file_link.Target))
            except Exception:
                targets.append("")
    except Exception:
        pass
    return targets


def _ext_files(text, targets):
    """Metindeki dış çalışma kitabı adlarını döndürür."""
    found = []
    for m in EXT_RE.finditer(text):
        key = m.group(1).strip()
        if key.isdigit():
            idx = int(key) - 1
            name = targets[idx] if 0 <= idx < len(targets) and targets[idx] else "Dış bağlantı [%s]" % key
        else:
            name = key
        name = name.replace("\\", "/").rsplit("/", 1)[-1] or name
        if name not in found:
            found.append(name)
    return found


def _hidden_columns(ws):
    hidden = set()
    try:
        for key, dim in ws.column_dimensions.items():
            if getattr(dim, "hidden", False):
                try:
                    start = dim.min or column_index_from_string(key)
                    end = dim.max or start
                except Exception:
                    continue
                for i in range(start, min(end, 16384) + 1):
                    hidden.add(i)
    except Exception:
        pass
    return hidden


def _find_cycles(graph):
    index, low, onstack, stack, sccs = {}, {}, set(), [], []
    counter = 0
    for root in graph:
        if root in index:
            continue
        index[root] = low[root] = counter
        counter += 1
        stack.append(root)
        onstack.add(root)
        work = [(root, iter(graph.get(root, ())))]
        while work:
            node, it = work[-1]
            advanced = False
            for nxt in it:
                if nxt not in index:
                    index[nxt] = low[nxt] = counter
                    counter += 1
                    stack.append(nxt)
                    onstack.add(nxt)
                    work.append((nxt, iter(graph.get(nxt, ()))))
                    advanced = True
                    break
                elif nxt in onstack:
                    low[node] = min(low[node], index[nxt])
            if advanced:
                continue
            work.pop()
            if work:
                parent = work[-1][0]
                low[parent] = min(low[parent], low[node])
            if low[node] == index[node]:
                comp = []
                while True:
                    w = stack.pop()
                    onstack.discard(w)
                    comp.append(w)
                    if w == node:
                        break
                if len(comp) > 1 or node in graph.get(node, ()):
                    sccs.append(comp)
    return sccs


def _refs_of(formula, current_sheet, sheet_map, formula_sets):
    """Formülün başvurduğu formül hücrelerini (sayfa, satır, sütun) döndürür."""
    text = STRING_RE.sub('""', formula)
    text = EXT_REF_RE.sub("0", text)
    out = set()
    for m in REF_RE.finditer(text):
        quoted, plain, c1, r1, c2, r2 = m.groups()
        name = quoted.replace("''", "'") if quoted else plain
        if name is None:
            sheet = current_sheet
        else:
            sheet = sheet_map.get(name.lower())
            if sheet is None:
                continue
        try:
            ci1 = column_index_from_string(c1.upper())
            ri1 = int(r1)
            if c2:
                ci2 = column_index_from_string(c2.upper())
                ri2 = int(r2)
            else:
                ci2, ri2 = ci1, ri1
        except Exception:
            continue
        if max(ci1, ci2) > 16384 or max(ri1, ri2) > 1048576:
            continue
        lo_c, hi_c = min(ci1, ci2), max(ci1, ci2)
        lo_r, hi_r = min(ri1, ri2), max(ri1, ri2)
        fset = formula_sets.get(sheet, set())
        area = (hi_c - lo_c + 1) * (hi_r - lo_r + 1)
        if area <= len(fset):
            for r in range(lo_r, hi_r + 1):
                for c in range(lo_c, hi_c + 1):
                    if (r, c) in fset:
                        out.add((sheet, r, c))
        else:
            for (r, c) in fset:
                if lo_r <= r <= hi_r and lo_c <= c <= hi_c:
                    out.add((sheet, r, c))
    return out


def _safe_csv(value):
    s = "" if value is None else str(value)
    if s and s[0] in "=+-@\t\r":
        return "'" + s
    return s


def _build_csv(issues):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Sayfa", "Hücre", "Tür", "Önem", "Ayrıntı", "Formül", "Gizli"])
    for i in issues:
        w.writerow([_safe_csv(i["sheet"]), i["cell"], i["type"], i["severity"],
                    _safe_csv(i["detail"]), _safe_csv(i["formula"]), "Evet" if i["hidden"] else "Hayır"])
    return buf.getvalue()


def analyze_workbook(data, filename="model.xlsx"):
    """xlsx baytlarını bellekte denetler ve sonuç sözlüğü döndürür."""
    if not data:
        raise AuditError("Dosya boş görünüyor. Lütfen bir .xlsx dosyası seçin.")
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise AuditError("Bu dosya eski .xls biçiminde. Excel'de 'Farklı Kaydet' ile .xlsx olarak kaydedip yeniden yükleyin.")
    bio = io.BytesIO(data)
    if not zipfile.is_zipfile(bio):
        raise AuditError("Bu dosya bir .xlsx çalışma kitabı gibi görünmüyor (metin ya da başka bir ikili biçim olabilir).")
    warnings = []
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            names = set(zf.namelist())
    except Exception:
        raise AuditError("Dosya bozuk görünüyor ve açılamadı.")
    if "xl/workbook.xml" not in names:
        raise AuditError("Bu zip dosyası bir Excel çalışma kitabı değil (xl/workbook.xml bulunamadı).")
    if "xl/vbaProject.bin" in names:
        warnings.append("Dosyada makro (VBA) var. Makrolar bu araçta incelenmez; yalnızca hücre formülleri ve adlandırılmış alanlar taranır.")

    try:
        wb = load_workbook(io.BytesIO(data), data_only=False)
    except Exception:
        raise AuditError("Çalışma kitabı okunamadı. Dosya bozuk, parola korumalı ya da desteklenmeyen bir biçimde olabilir.")
    try:
        wb_vals = load_workbook(io.BytesIO(data), data_only=True)
    except Exception:
        wb_vals = None
        warnings.append("Kaydedilmiş hesap sonuçları okunamadı; yalnızca formül metinleri denetlendi.")

    targets = _external_targets(wb)
    sheets = [ws for ws in wb.worksheets]
    sheet_names = [ws.title for ws in sheets]
    sheet_map = {n.lower(): n for n in sheet_names}
    order = {n: i for i, n in enumerate(sheet_names)}

    issues = []
    ext_counter = Counter()
    formula_sets = defaultdict(set)
    formula_texts = {}
    formula_count = 0
    cell_count = 0
    no_cache = 0
    truncated_cells = False

    def add(sheet, row, col, typ, detail, formula="", hidden=False):
        issues.append({
            "sheet": sheet,
            "cell": "%s%d" % (get_column_letter(col), row) if row else "(ad)",
            "row": row, "col": col, "type": typ, "severity": SEVERITY[typ],
            "detail": detail, "formula": formula, "hidden": hidden,
        })

    for ws in sheets:
        sname = ws.title
        sheet_hidden = getattr(ws, "sheet_state", "visible") != "visible"
        hidden_cols = _hidden_columns(ws)
        vals_ws = None
        if wb_vals is not None:
            try:
                vals_ws = wb_vals[sname]
            except Exception:
                vals_ws = None
        try:
            rows = ws.iter_rows()
            for row in rows:
                for c in row:
                    v = c.value
                    if v is None:
                        continue
                    cell_count += 1
                    if cell_count > MAX_CELLS:
                        truncated_cells = True
                        break
                    hid = sheet_hidden or (c.column in hidden_cols)
                    ftxt = _formula_text(c)
                    if ftxt is None:
                        if c.data_type == "e" and isinstance(v, str) and v in ERROR_CODES:
                            add(sname, c.row, c.column, T_ERRVAL, "Hücre %s hata değeri içeriyor." % v, "", hid)
                        continue
                    formula_count += 1
                    formula_sets[sname].add((c.row, c.column))
                    formula_texts[(sname, c.row, c.column)] = ftxt
                    stripped = STRING_RE.sub('""', ftxt)
                    files = _ext_files(stripped, targets)
                    if files:
                        for f in files:
                            ext_counter[f] += 1
                        add(sname, c.row, c.column, T_EXT,
                            "Dış çalışma kitabına başvuruyor: %s" % ", ".join(files), ftxt, hid)
                    codes = []
                    for m in ERR_LIT_RE.finditer(stripped):
                        if m.group(0) not in codes:
                            codes.append(m.group(0))
                    if codes:
                        add(sname, c.row, c.column, T_BROKEN,
                            "Formülün içinde %s var; başvuru kopmuş ya da silinmiş." % ", ".join(codes), ftxt, hid)
                    if vals_ws is not None:
                        try:
                            cached = vals_ws.cell(row=c.row, column=c.column).value
                        except Exception:
                            cached = None
                        if cached is None:
                            no_cache += 1
                        elif isinstance(cached, str) and cached in ERROR_CODES:
                            add(sname, c.row, c.column, T_ERRVAL,
                                "Kaydedilmiş hesap sonucu %s." % cached, ftxt, hid)
                if truncated_cells:
                    break
        except Exception:
            warnings.append("'%s' sayfası tam okunamadı; bu sayfanın sonuçları eksik olabilir." % sname)
        if truncated_cells:
            warnings.append("Çalışma kitabı çok büyük: ilk %d dolu hücre tarandı, kalanı atlandı." % MAX_CELLS)
            break

    # Adlandırılmış alanlar
    names_checked = 0
    name_items = []
    try:
        for n, dn in wb.defined_names.items():
            name_items.append(("(Çalışma kitabı)", n, dn))
    except Exception:
        pass
    for ws in sheets:
        try:
            for n, dn in ws.defined_names.items():
                name_items.append((ws.title, n, dn))
        except Exception:
            pass
    for scope, n, dn in name_items:
        names_checked += 1
        text = str(getattr(dn, "attr_text", "") or "")
        if ERR_LIT_RE.search(text):
            add(scope, 0, 0, T_NAME, "'%s' adlı alan geçersiz bir başvuruya işaret ediyor: %s" % (n, text), text)
        else:
            files = _ext_files(text, targets)
            if files:
                for f in files:
                    ext_counter[f] += 1
                add(scope, 0, 0, T_EXT, "'%s' adlı alan dış çalışma kitabına işaret ediyor: %s" % (n, ", ".join(files)), text)

    # Dairesel başvurular
    circular_cells = 0
    if formula_count > MAX_GRAPH_NODES:
        warnings.append("Formül sayısı %d sınırını aştığı için dairesel başvuru taraması yapılmadı." % MAX_GRAPH_NODES)
    else:
        try:
            graph = {}
            for (sname, r, c), ftxt in formula_texts.items():
                graph[(sname, r, c)] = _refs_of(ftxt, sname, sheet_map, formula_sets)
            for comp in _find_cycles(graph):
                members = sorted(comp, key=lambda x: (order.get(x[0], 0), x[1], x[2]))
                labels = ["%s!%s%d" % (s, get_column_letter(c), r) for s, r, c in members[:6]]
                more = " …" if len(members) > 6 else ""
                for s, r, c in members:
                    circular_cells += 1
                    add(s, r, c, T_CIRC,
                        "Bu hücre %d hücrelik bir döngünün parçası: %s%s" % (len(members), " → ".join(labels), more),
                        formula_texts[(s, r, c)])
        except Exception:
            warnings.append("Dairesel başvuru taraması tamamlanamadı.")

    if cell_count == 0:
        raise AuditError("Çalışma kitabında okunabilir dolu hücre bulunamadı; denetlenecek bir şey yok.")
    if formula_count == 0 and not issues:
        warnings.append("Çalışma kitabında hiç formül yok; yalnızca değerler var.")
    if vals_ws_missing(formula_count, no_cache):
        warnings.append("Dosyada kaydedilmiş hesap sonuçları yok (dosya Excel'de hesaplanıp kaydedilmemiş olabilir); yalnızca formül metinleri denetlendi.")

    issues.sort(key=lambda i: (order.get(i["sheet"], -1), i["row"], i["col"], i["type"]))
    total_issues = len(issues)
    if total_issues > MAX_ISSUES:
        issues = issues[:MAX_ISSUES]
        warnings.append("%d sorun bulundu; liste ve CSV ilk %d sorunla sınırlandı." % (total_issues, MAX_ISSUES))

    type_counts = Counter(i["type"] for i in issues)
    problem_cells = len({(i["sheet"], i["cell"]) for i in issues})
    hidden_issues = sum(1 for i in issues if i["hidden"])
    csv_text = "﻿" + _build_csv(issues)
    return {
        "filename": filename,
        "sheet_count": len(sheets),
        "formula_count": formula_count,
        "cell_count": min(cell_count, MAX_CELLS),
        "names_checked": names_checked,
        "issues": issues,
        "total_issues": total_issues,
        "problem_cells": problem_cells,
        "hidden_issues": hidden_issues,
        "circular_cells": circular_cells,
        "type_counts": [(t, type_counts.get(t, 0)) for t in (T_EXT, T_BROKEN, T_ERRVAL, T_CIRC, T_NAME)],
        "ext_files": sorted(ext_counter.items(), key=lambda kv: (-kv[1], kv[0])),
        "warnings": warnings,
        "csv_uri": "data:text/csv;charset=utf-8," + quote(csv_text),
    }


def vals_ws_missing(formula_count, no_cache):
    return formula_count > 0 and no_cache == formula_count


def build_demo_workbook(today=None):
    """Bellekte, bilinçli hatalar içeren örnek bir SaaS modeli üretir (bugüne göreli tarihlerle)."""
    today = today or datetime.date.today()
    wb = Workbook()
    ws = wb.active
    ws.title = "Varsayimlar"
    for r in [("Varsayım", "Değer"), ("Aylık büyüme", 0.08), ("Başlangıç MRR", 12000),
              ("Kasa", 250000), ("Aylık churn", 0.03), ("Model tarihi", today)]:
        ws.append(r)

    months = []
    y, m = today.year, today.month
    for _ in range(6):
        months.append(datetime.date(y, m, 1))
        m += 1
        if m > 12:
            m, y = 1, y + 1
    g = wb.create_sheet("Gelir")
    g.append(["Kalem"] + months + ["Eski toplam"])
    g.append(["MRR", "=Varsayimlar!B3", "=B2*(1+Varsayimlar!$B$2)",
              "=C2*(1+'[Eski_Model.xlsx]Varsayimlar'!$B$2)", "=D2*(1+Varsayimlar!$B$2)",
              "=E2*(1+Varsayimlar!#REF!)", "=F2*(1+Varsayimlar!$B$2)", "=SUM(B2:G2)+#REF!"])
    g.append(["Churn kaybı"] + ["=%s2*Varsayimlar!$B$5" % col for col in "BCDEFG"])
    g.append(["Net MRR", "=B2-B3", "=C2-C3", "=D2-D3", "=E2-E3", "=F2-F3", "#DIV/0!"])
    g.column_dimensions["H"].hidden = True

    n = wb.create_sheet("Nakit")
    n.append(["Kalem"] + months[:3])
    n.append(["Kasa", "=Varsayimlar!B4", "=B2+Gelir!C4-C3", "=C2+Gelir!D4-D3"])
    n.append(["Gider", 40000, 41000, 42000])
    n.append(["Aylık net yakım", "=B3-Gelir!B4", "=C3-Gelir!C4", "=D3-Gelir!D4"])
    n.append(["Pist (ay)", "=B6+1", None, None])
    n.append(["Düzeltme", "=B5*2", None, None])
    n.append(["Pist", "=B2/B4", "=C2/C4", "=D2/D4"])

    wb.defined_names["Eski_Kasa"] = DefinedName("Eski_Kasa", attr_text="#REF!")
    wb.defined_names["Dis_Parametre"] = DefinedName("Dis_Parametre", attr_text="'[Eski_Model.xlsx]Varsayimlar'!$B$4")
    wb.defined_names["Kasa_Bakiye"] = DefinedName("Kasa_Bakiye", attr_text="Varsayimlar!$B$4")
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def analyze_demo():
    return analyze_workbook(build_demo_workbook(), "ornek_saas_modeli.xlsx")
