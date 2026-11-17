# -*- coding: utf-8 -*-
"""Teste de inserção de linhas dentro das tabelas (LibreOffice via UNO). Uso: uno_insercao.py <xlsx> <saida.json>"""
import sys, os, time, json, subprocess, tempfile, datetime
import uno
from com.sun.star.beans import PropertyValue

src, out = sys.argv[1], sys.argv[2]
prof = os.path.join(tempfile.gettempdir(), 'lo_perfil_uno')
os.makedirs(os.path.join(prof, 'user'), exist_ok=True)
open(os.path.join(prof, 'user', 'registrymodifications.xcu'), 'w', encoding='utf-8').write(
    '<?xml version="1.0" encoding="UTF-8"?>\n<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
    '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>\n</oor:items>\n')
port = 2002
proc = subprocess.Popen(['soffice', '-env:UserInstallation=file://' + prof, '--headless', '--invisible', '--norestore', '--nologo',
                         '--accept=socket,host=localhost,port=%d;urp;' % port], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
try:
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    ctx = None
    for _ in range(60):
        try:
            ctx = resolver.resolve('uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext' % port); break
        except Exception:
            time.sleep(1)
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    p = PropertyValue(); p.Name = 'Hidden'; p.Value = True
    doc = desktop.loadComponentFromURL('file://' + os.path.abspath(src), '_blank', 0, (p,))
    sheets = doc.Sheets
    def serial(d): return (d - datetime.date(1899, 12, 30)).days
    checks = []
    ref = datetime.date(2026, 11, 16)
    s1 = sheets.getByName('01 Comece aqui'); s3 = sheets.getByName('03 Contas'); s4 = sheets.getByName('04 Lotes'); s6 = sheets.getByName('06 Compras')
    s1.getCellRangeByName('B48').setValue(serial(ref))
    # as abas são protegidas SEM senha: para inserir linhas, o aluno desproteger (Revisão > Desproteger planilha)
    for sh in (s3, s4, s6, s1):
        sh.unprotect('')
    # formula antes
    f_antes = s3.getCellRangeByName('D7').getFormula()
    # 1) insere uma linha DENTRO da tabela de lotes (antes da linha 10) e preenche
    s4.Rows.insertByIndex(9, 1)
    # a linha nova nasce sem as fórmulas das colunas calculadas: o aluno copia as células calculadas da linha de cima
    s4.copyRange(s4.getCellRangeByName('H10').getCellAddress(), s4.getCellRangeByName('H9:I9').getRangeAddress())
    s4.getCellRangeByName('A10').setString('Programa teste')
    s4.getCellRangeByName('D10').setValue(777)
    s4.getCellRangeByName('F10').setValue(serial(ref + datetime.timedelta(days=20)))
    s3.getCellRangeByName('A7').setString('Programa teste')
    doc.calculateAll()
    f_depois = s3.getCellRangeByName('D7').getFormula()
    checks.append(['4.01', 'Inserir linha na aba 04: a soma de lotes da aba 03 inclui a nova linha (777)', 777, s3.getCellRangeByName('E7').getValue()])
    checks.append(['4.02', '… a faixa da fórmula cresceu de 156 para 157', [True, True], ['$156' in f_antes.replace(',', ';') or '156' in f_antes, '157' in f_depois]])
    checks.append(['4.03', '… copiando as colunas calculadas da linha de cima, a situação do lote novo é calculada (vence em até 30 dias)', 'Vence em até 30 dias', s4.getCellRangeByName('I10').getString()])
    checks.append(['4.04', '… e a contagem da aba 01 acompanha', 1, s1.getCellRangeByName('B54').getValue()])
    # 2) preenche a última linha de capacidade (agora 157)
    s4.getCellRangeByName('A157').setString('Programa teste'); s4.getCellRangeByName('D157').setValue(23)
    doc.calculateAll()
    checks.append(['4.05', 'Última linha de capacidade da aba 04 entra na soma (777 + 23)', 800, s3.getCellRangeByName('E7').getValue()])
    # 3) insere uma linha na aba 06 e confere que as contagens da aba 01 acompanham
    s6.Rows.insertByIndex(9, 1)
    s6.getCellRangeByName('B10').setString('Loja teste'); s6.getCellRangeByName('U10').setString('Aguardando')
    doc.calculateAll()
    checks.append(['4.06', 'Inserir linha na aba 06: a contagem de compras aguardando (aba 01) conta a linha nova', 1, s1.getCellRangeByName('B58').getValue()])
    s1_f = s1.getCellRangeByName('B58').getFormula()
    checks.append(['4.07', '… e a faixa da fórmula passou de 106 para 107', True, '107' in s1_f])
    # 4) copiar a linha de fórmulas para baixo (linha extra após a última capacidade)
    doc.close(True)
    json.dump({'checks': checks}, open(out, 'w', encoding='utf-8'), ensure_ascii=False, default=str)
finally:
    try:
        desktop.terminate()
    except Exception:
        pass
    time.sleep(1); proc.kill()
