# -*- coding: utf-8 -*-
import subprocess, re, os, sys
base = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
tmp = os.path.join(base, 'previas-tmp'); out = os.path.join(base, 'previas'); os.makedirs(tmp, exist_ok=True); os.makedirs(out, exist_ok=True)
prof = 'file:///tmp/lo_perfil_recalc'
slug = {'01 Comece aqui': '01-comece-aqui', '02 Diagnóstico': '02-diagnostico', '03 Contas': '03-contas', '04 Lotes': '04-lotes', '05 Acúmulo': '05-acumulo', '06 Compras': '06-compras', '07 Comparar': '07-comparar', '08 Alterar': '08-alterar', '09 Clubes e compra': '09-clubes-e-compra', '10 Transferências': '10-transferencias', '11 Renovar': '11-renovar', '12 Plano 30 dias': '12-plano-30-dias', '13 Checklists': '13-checklists', '14 Referências': '14-referencias'}
def pdf(fn):
    subprocess.run(['soffice', '-env:UserInstallation=' + prof, '--headless', '--convert-to', 'pdf', '--outdir', tmp, os.path.join(base, 'saida', fn + '.xlsx')], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    P = os.path.join(tmp, fn + '.pdf'); n = int(re.search(r'Pages:\s+(\d+)', subprocess.check_output(['pdfinfo', P], text=True)).group(1)); first = {}
    for pg in range(1, n + 1):
        t = subprocess.check_output(['pdftotext', '-f', str(pg), '-l', str(pg), '-layout', P, '-'], text=True)
        m = re.search(r'Milhas do zero . (\d\d [^\n]+?)\s*$', t, flags=re.M)
        if m and m.group(1).strip() not in first: first[m.group(1).strip()] = pg
    return P, first
P, fe = pdf('Milhas-do-zero-Exemplos-preenchidos')
for sh, pg in fe.items(): subprocess.run(['pdftoppm', '-r', '85', '-png', '-f', str(pg), '-l', str(pg), '-singlefile', P, os.path.join(out, 'exemplos-' + slug[sh])], check=True)
P2, fa = pdf('Milhas-do-zero-Planilha-do-aluno')
for sh in ('01 Comece aqui', '02 Diagnóstico', '04 Lotes', '06 Compras'): subprocess.run(['pdftoppm', '-r', '85', '-png', '-f', str(fa[sh]), '-l', str(fa[sh]), '-singlefile', P2, os.path.join(out, 'aluno-' + slug[sh])], check=True)
print(len(os.listdir(out)), 'prévias')
