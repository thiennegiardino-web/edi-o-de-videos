# -*- coding: utf-8 -*-
import os, sys, shutil, subprocess, datetime, tempfile
sys.path.insert(0, os.path.dirname(__file__))
from common import *
import sheets_a as A, sheets_b as B, sheets_c as C
import openpyxl

ARQ = [(False, 'Milhas-do-zero-Planilha-do-aluno.xlsx'), (True, 'Milhas-do-zero-Exemplos-preenchidos.xlsx')]

def build(out_dir, values):
    os.makedirs(out_dir, exist_ok=True)
    cells = {}
    for ex, fn in ARQ:
        path = os.path.join(out_dir, fn)
        book = Book(path, ex, values.get(fn, {}))
        for f in (A.sheet01, A.sheet02, A.sheet03, A.sheet04, A.sheet05, A.sheet06, B.sheet07, B.sheet08, B.sheet09, B.sheet10, B.sheet11, C.sheet12, C.sheet13, C.sheet14):
            f(book)
        book.wb.close()
        cells[fn] = list(book.formula_cells)
    return cells

LO_PROFILE = os.path.join(tempfile.gettempdir(), 'lo_perfil_recalc')

def _perfil():
    """Perfil do LibreOffice que força recálculo total ao abrir arquivos .xlsx (sem confiar em valores em cache)."""
    d = os.path.join(LO_PROFILE, 'user'); os.makedirs(d, exist_ok=True)
    open(os.path.join(d, 'registrymodifications.xcu'), 'w', encoding='utf-8').write(
        '<?xml version="1.0" encoding="UTF-8"?>\n<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
        '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>\n'
        '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>\n</oor:items>\n')
    return 'file://' + LO_PROFILE

def recalc(src_dir, tmp_dir, files=None):
    os.makedirs(tmp_dir, exist_ok=True)
    prof = _perfil()
    for fn in (files or [f for _, f in ARQ]):
        subprocess.run(['soffice', '-env:UserInstallation=' + prof, '--headless', '--calc', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', tmp_dir, os.path.join(src_dir, fn)],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)

def read_values(tmp_dir, cells):
    out = {}
    for _, fn in ARQ:
        wb = openpyxl.load_workbook(os.path.join(tmp_dir, fn), data_only=True)
        d = {}
        for (sh, cell) in cells[fn]:
            v = wb[sh][cell].value
            if isinstance(v, datetime.datetime): v = (v - datetime.datetime(1899, 12, 30)).total_seconds() / 86400.0
            elif isinstance(v, datetime.date): v = excel_serial(v)
            elif isinstance(v, datetime.time): v = (v.hour * 3600 + v.minute * 60 + v.second) / 86400.0
            elif isinstance(v, datetime.timedelta): v = v.total_seconds() / 86400.0
            if v is None: v = ''
            d[(sh, cell)] = v
        out[fn] = d
    return out

if __name__ == '__main__':
    final = sys.argv[1]
    work = tempfile.mkdtemp(prefix='xl_')
    p1 = os.path.join(work, 'p1'); tmp = os.path.join(work, 'lo')
    cells = build(p1, {})
    recalc(p1, tmp)
    vals = read_values(tmp, cells)
    build(final, vals)
    errs = [(fn, k, v) for fn, d in vals.items() for k, v in d.items() if isinstance(v, str) and v.startswith(('#', 'Err:'))]
    print('fórmulas:', {fn: len(c) for fn, c in cells.items()}, '| erros nos valores calculados:', errs[:10])
