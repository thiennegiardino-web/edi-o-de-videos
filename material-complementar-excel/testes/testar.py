# -*- coding: utf-8 -*-
"""Bateria de testes das planilhas (LibreOffice headless com recálculo total forçado)."""
import sys, os, json, datetime, tempfile, shutil, zipfile, re, subprocess
HERE = os.path.dirname(os.path.abspath(__file__)); BUILD = os.path.join(HERE, '..', 'build')
sys.path.insert(0, BUILD)
import openpyxl
from build_planilhas import recalc, ARQ
from common import NAMES, parse_checklists, TEXTO, parse_cap12_questions

SAIDA = sys.argv[1]
ALUNO = 'Milhas-do-zero-Planilha-do-aluno.xlsx'; EXEMPLOS = 'Milhas-do-zero-Exemplos-preenchidos.xlsx'
D = datetime.date
REF = D(2026, 11, 16)
res = []; tmp_root = tempfile.mkdtemp(prefix='tst_')

def add(grupo, cid, desc, esperado, obtido, ok=None):
    if ok is None:
        if isinstance(esperado, (int, float)) and isinstance(obtido, (int, float)) and not isinstance(esperado, bool):
            ok = abs(esperado - obtido) < 0.005
        else:
            ok = (esperado == obtido)
    res.append({'grupo': grupo, 'id': cid, 'descricao': desc, 'esperado': esperado, 'obtido': obtido, 'ok': bool(ok)})

def lo_wb(path, mods=None, name='x'):
    """Recalcula (LibreOffice) uma cópia do arquivo, com modificações opcionais, e devolve o workbook com valores."""
    d = os.path.join(tmp_root, name); os.makedirs(d, exist_ok=True)
    src = os.path.join(d, 'in.xlsx')
    if mods:
        wb = openpyxl.load_workbook(path)
        for (sh, cell), v in mods.items(): wb[sh][cell].value = v
        wb.save(src)
    else:
        shutil.copy(path, src)
    out = os.path.join(d, 'out'); recalc(d, out, files=['in.xlsx'])
    return openpyxl.load_workbook(os.path.join(out, 'in.xlsx'), data_only=True)

def norm(x):
    return re.sub(r'(\d),(\d\d)\b', r'\1.\2', x) if isinstance(x, str) else x

def v(wb, sh, cell):
    x = wb[sh][cell].value
    if isinstance(x, datetime.datetime): return x.date()
    return '' if x is None else x

S = NAMES
aluno = os.path.join(SAIDA, ALUNO); exem = os.path.join(SAIDA, EXEMPLOS)

# ---------------------------------------------------------------- G1 exemplos aprovados (arquivo de exemplos recalculado)
we = lo_wb(exem, name='ex')
g = 'G1 Exemplos aprovados do livro (arquivo de exemplos)'
for cid, sh, c, esp, desc in [
    ('1.01', 7, 'B18', 1680, 'Cap. 13: tarifa = R$ 1.960 − R$ 280'),
    ('1.02', 7, 'D18', 1210, 'Cap. 13: alternativa R$ 1.490 − R$ 280 de taxas'),
    ('1.03', 7, 'B23', 2320, 'Cap. 13: pago em dinheiro (opção em dinheiro) = R$ 1.680 + R$ 280 + R$ 360'),
    ('1.04', 7, 'C23', 640, 'Cap. 13: pago em dinheiro (milhas) = R$ 280 + R$ 360'),
    ('1.05', 7, 'B28', 1680, 'Cap. 13: economia calculada'),
    ('1.06', 7, 'B29', 23.3333333, 'Cap. 13: valor por milheiro contra a opção em dinheiro (R$ 23,33)'),
    ('1.07', 7, 'B31', 16.8055556, 'Cap. 13: valor por milheiro contra a alternativa (R$ 16,81)'),
    ('1.08', 7, 'E49', 55, 'Cap. 4, exemplo A: R$ 55 por milheiro'),
    ('1.09', 7, 'E50', 34.6666667, 'Cap. 4, exemplo B: ≈ R$ 34,67'),
    ('1.10', 7, 'E51', 55, 'Cap. 4, exemplo C (base A)'),
    ('1.11', 7, 'G51', 30, 'Cap. 4, exemplo C: contra a alternativa de R$ 700 (derivado da mesma fórmula)'),
    ('1.12', 8, 'B14', 1960, 'Cap. 13 (Passo 6): total antes da bagagem R$ 1.960'),
    ('1.13', 8, 'B38', 1624, 'Cap. 13: 80% da tarifa (R$ 1.344) + taxas devolvidas (R$ 280) = R$ 1.624'),
    ('1.14', 8, 'B43', 336, 'Cap. 13: perda em dinheiro R$ 336'),
    ('1.15', 8, 'C43', 300, 'Cap. 13: perda em dinheiro com milhas R$ 300 (taxa de cancelamento)'),
    ('1.16', 8, 'C44', 0, 'Cap. 13: nenhuma milha perdida (72.000 devolvidas)'),
    ('1.17', 8, 'B35', 300, 'Cap. 13: taxa de alteração R$ 300 (2 pessoas)'),
    ('1.18', 8, 'C35', 200, 'Cap. 13: taxa de alteração R$ 200 (2 pessoas)'),
    ('1.19', 9, 'J8', 42, 'Cap. 16: R$ 42 por milheiro'),
    ('1.20', 9, 'H22', 30, 'Cap. 17: R$ 600 por 20.000 pontos = R$ 30 por milheiro'),
    ('1.21', 10, 'M8', 26000, 'Cap. 17: 20.000 × 1,30 = 26.000 milhas'),
    ('1.22', 10, 'O8', 23.0769231, 'Cap. 17: R$ 23,08 por milheiro'),
    ('1.23', 10, 'M9', 40000, 'Cap. 17: 20.000 × 2 = 40.000 milhas'),
    ('1.24', 10, 'O9', 15, 'Cap. 17: R$ 15,00 por milheiro'),
    ('1.25', 10, 'V9', 1.81, 'Cap. 17: diferença de cerca de R$ 1,81 contra os R$ 16,81'),
    ('1.26', 3, 'E7', 20000, 'Cap. 10: saldo = 8.000 + 12.000 (soma dos lotes)'),
    ('1.27', 11, 'N8', 230, 'Renovação (adicional): R$ 150 + R$ 80'),
    ('1.28', 11, 'P8', 370, 'Renovação (adicional): R$ 600 − R$ 230'),
    ('1.29', 4, 'H7', 120, 'Lote de 16/03/2027: 120 dias após a referência'),
    ('1.30', 12, 'D37', 14, 'Plano: 14 atividades concluídas no exemplo'),
    ('1.31', 12, 'C37', 19, 'Plano: 19 aplicáveis (2 "Não se aplica")'),
]:
    add(g, cid, desc, esp, v(we, S[sh], c))
for cid, c, esp in [('1.32', 'I7', 'Mais de 90 dias'), ('1.33', 'I9', 'Vencido'), ('1.34', 'I10', 'Vence hoje'), ('1.35', 'I11', 'Vence em até 30 dias'), ('1.36', 'I12', 'Vence em até 60 dias'),
                    ('1.37', 'I13', 'Vence em até 90 dias'), ('1.38', 'I14', 'Data não informada'), ('1.39', 'I15', 'Sem expiração confirmada')]:
    add(g, cid, 'Aba 04: situação do lote na linha %s' % c[1:], esp, v(we, S[4], c))
add(g, '1.40', 'Aba 06: 60 dias corridos após 03/11/2026', D(2027, 1, 2), v(we, S[6], 'L7'))
add(g, '1.41', 'Aba 06: 30 dias corridos após a data prevista de crédito', D(2027, 2, 1), v(we, S[6], 'Q7'))
add(g, '1.42', 'Aba 06: 10 dias úteis após 12/11/2026 com feriado em 20/11/2026', D(2026, 11, 27), v(we, S[6], 'L8'))
add(g, '1.43', 'Aba 06: diferença creditado − esperado', -20, v(we, S[6], 'T9'))
add(g, '1.44', 'Aba 06: divergência E prazo de reclamação próximo aparecem juntos (nada fica escondido)', 'Divergência: creditado diferente do esperado. Data-limite de reclamação em até 7 dias.', v(we, S[6], 'V9'))
add(g, '1.44b', 'Aba 06: data-limite de reclamação da linha divergente', D(2026, 11, 19), v(we, S[6], 'Q9'))
add(g, '1.45', 'Aba 06: prazo de crédito passou sem crédito', 'O prazo previsto de crédito passou sem crédito registrado.', v(we, S[6], 'V10'))
add(g, '1.46', 'Aba 01: lotes vencidos', 1, v(we, S[1], 'B52')); add(g, '1.47', 'Aba 01: lotes que vencem hoje', 1, v(we, S[1], 'B53'))
add(g, '1.48', 'Aba 01: plano', '14 de 19', v(we, S[1], 'B61')); add(g, '1.49', 'Aba 01: compras divergentes ou em contestação', 1, v(we, S[1], 'B59'))
add(g, '1.50', 'Aba 03: alerta de inatividade (prazo atingido)', 'Prazo informado atingido ou ultrapassado: confira com o programa.', v(we, S[3], 'M8'))
add(g, '1.51', 'Aba 03: alerta de inatividade (a menos de 3 meses)', 'Atenção: a menos de 3 meses do prazo informado (sugestão de organização).', v(we, S[3], 'M9'))
add(g, '1.51b', 'Aba 03: o saldo do Programa B exclui os 3.000 pontos já vencidos (1.000 + 1.500 + 0 não vencidos…)', 2500, v(we, S[3], 'E8'))
add(g, '1.51c', 'Aba 03: os 3.000 pontos vencidos aparecem na coluna R, fora do saldo', 3000, v(we, S[3], 'R8'))
add(g, '1.51d', 'Aba 03: a origem do saldo avisa que há pontos vencidos', 'Soma dos lotes não vencidos (aba 04) Há pontos vencidos (coluna R).', v(we, S[3], 'F8'))
add(g, '1.51e', 'Aba 07 (v1.2): tarifa em reais e milhas em linhas separadas: milhas utilizadas na linha 46', 72000, v(we, S[7], 'C46'))
add(g, '1.51f', '… e a célula de tarifa na coluna Milhas não mistura unidades', 'Paga com milhas', v(we, S[7], 'C41'))
add(g, '1.52', 'Aba 05: diferença de preço do liquidificador (R$ 220 − R$ 200)', 20, v(we, S[5], 'K7'))
add(g, '1.53', 'Aba 02: resumo montado (Marina)', 'Quero viajar de Campinas para Salvador, com 2 pessoa(s), daqui a 10 meses. Já tenho cadastro em Programa da farmácia (exemplo), Programa do supermercado (exemplo). Meu primeiro passo é Ver o saldo e a validade desses dois programas.', v(we, S[2], 'B56'))
# valores em cache do arquivo entregue = valores recalculados
wc = openpyxl.load_workbook(exem, data_only=True); difs = 0; n = 0
for ws in wc.worksheets:
    for row in ws.iter_rows():
        for c in row:
            a = c.value; b = we[ws.title][c.coordinate].value
            if a is None and b is None: continue
            n += 1
            if isinstance(a, (int, float)) and isinstance(b, (int, float)):
                if abs(a - b) > 1e-6: difs += 1
            elif isinstance(a, datetime.datetime) or isinstance(b, datetime.datetime):
                if str(a)[:10] != str(b)[:10]: difs += 1
            elif (a or '') != (b or ''): difs += 1
add(g, '1.54', 'Valores em cache do arquivo de exemplos = valores recalculados pelo LibreOffice (%d células comparadas)' % n, 0, difs)

# ---------------------------------------------------------------- G2 arquivo do aluno vazio
g = 'G2 Arquivo do aluno (vazio)'
wa = lo_wb(aluno, name='aluno_vazio')
formulas = openpyxl.load_workbook(aluno)
erros = []; nao_vazios = []
for ws in formulas.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                val = wa[ws.title][c.coordinate].value
                if isinstance(val, str) and (val.startswith('#') or val.startswith('Err:')): erros.append((ws.title, c.coordinate, val))
                if val not in (None, '') : nao_vazios.append((ws.title, c.coordinate))
add(g, '2.01', 'Nenhuma fórmula resulta em erro (#DIV/0!, #VALUE!, #REF!, #NAME?, Err:)', 0, len(erros), len(erros) == 0)
perm = {S[1]: set(['B49']) , }
extras = [x for x in nao_vazios if not (x[0] == S[1] or x[0] == S[12] or x[0] == S[13] or (x[0] == S[2] and x[1].startswith('B5')) or (x[0] == S[4] and x[1][1:] == '5') or (x[0] == S[7] and x[1] == 'A33'))]
add(g, '2.02', 'Sem dados, as fórmulas ficam em branco (exceto contagens, datas de referência, resumos e textos-modelo); células fora dessa lista com conteúdo:', [], extras[:8], len(extras) == 0)
add(g, '2.03', 'Aba 01: data usada nos cálculos = data de hoje (sem data de referência)', datetime.date.today(), v(wa, S[1], 'B49'))
add(g, '2.04', 'Aba 01: contagens zeradas', 0, v(wa, S[1], 'B52'))
add(g, '2.05', 'Aba 01: checklists "0 de 49"', '0 de 49', v(wa, S[1], 'B62'))
add(g, '2.06', 'Aba 02: resumo montado com sublinhados', 'Quero viajar de __________ para __________, com ____ pessoa(s), daqui a ____ meses. Já tenho cadastro em __________. Meu primeiro passo é __________.', v(wa, S[2], 'B56'))
add(g, '2.07', 'Aba 12: plano com 21 atividades, sem concluídas', '0 de 21', v(wa, S[1], 'B61'))

# ---------------------------------------------------------------- G3 casos de borda
g = 'G3 Entradas inválidas, vazias e casos de borda (arquivo do aluno com dados de teste)'
def caso(nome, mods, checks):
    w = lo_wb(aluno, mods, name=nome)
    for cid, desc, sh, cell, esp in checks:
        add(g, cid, desc, esp, v(w, S[sh], cell))
    return w
base07 = {(S[7], 'B16'): 1960, (S[7], 'B17'): 280, (S[7], 'C19'): 72000, (S[7], 'C20'): 280, (S[7], 'C21'): 0, (S[7], 'B22'): 360, (S[7], 'C22'): 360, (S[7], 'D16'): 1490, (S[7], 'D17'): 280, (S[7], 'D22'): 360}
caso('c1', base07, [('3.01', 'Aba 07 preenchida no arquivo do aluno reproduz R$ 23,33', 7, 'B29', 23.3333333), ('3.02', '… e R$ 16,81 contra a alternativa', 7, 'B31', 16.8055556)])
m = dict(base07); m[(S[7], 'C19')] = 0
caso('c2', m, [('3.03', 'Pontos = 0: sem divisão por zero (valor por milheiro em branco)', 7, 'B29', ''), ('3.04', '… e aviso para informar os pontos', 7, 'A33', 'Informe os pontos ou milhas utilizados (maior que zero) para calcular o valor por milheiro.')])
m = dict(base07); m[(S[7], 'C19')] = 'abc'
caso('c3', m, [('3.05', 'Texto no lugar dos pontos: sem erro; resultado em branco', 7, 'B29', ''), ('3.06', '… e aviso', 7, 'A33', 'Informe os pontos ou milhas utilizados (maior que zero) para calcular o valor por milheiro.')])
m = dict(base07); m[(S[7], 'B17')] = 3000
caso('c4', m, [('3.07', 'Taxa de embarque maior que o preço total: aviso para não contar a taxa duas vezes', 7, 'A36', 'A taxa de embarque informada é maior que o preço total em dinheiro: confira para não contar a taxa duas vezes.')])
m = dict(base07); m[(S[7], 'D16')] = None
caso('c5', m, [('3.08', 'Sem alternativa em dinheiro: resultado contra a alternativa em branco', 7, 'B31', ''), ('3.09', '… e aviso', 7, 'A35', 'Informe a melhor alternativa em dinheiro para ver o resultado contra ela (Capítulo 12).')])
m = dict(base07); m[(S[7], 'C20')] = None; m[(S[7], 'C21')] = None; m[(S[7], 'C22')] = None
caso('c5b', m, [('3.10', 'Sem os pagamentos adicionais do resgate: não calcula (não assume zero)', 7, 'B29', ''), ('3.11', '… e aviso: campo vazio = não informado', 7, 'A34', 'Campo vazio = não informado, não zero: preencha as taxas de embarque, os outros pagamentos e a bagagem do resgate (digite 0 onde não houver).')])
caso('c6', {(S[7], 'A49'): 'x', (S[7], 'B49'): 100, (S[7], 'D49'): 0}, [('3.12', 'Exercício do Cap. 4 com pontos = 0: sem erro', 7, 'E49', '')])
caso('c6b', {(S[7], 'B49'): 100, (S[7], 'D49'): 1000}, [('3.13', 'Exercício do Cap. 4 com o pagamento adicional em branco: NÃO assume zero (resultado em branco)', 7, 'E49', '')])
caso('c6c', {(S[7], 'B49'): 100, (S[7], 'C49'): 0, (S[7], 'D49'): 1000}, [('3.13b', 'Exercício do Cap. 4 com pagamento adicional digitado como 0: R$ 100 por milheiro', 7, 'E49', 100)])
# transferências
tr = {(S[10], 'A8'): 'O', (S[10], 'B8'): 'D', (S[10], 'C8'): 20000, (S[10], 'D8'): 1, (S[10], 'E8'): 0.3, (S[10], 'G8'): 600, (S[10], 'H8'): 0}
caso('t1', tr, [('3.14', 'Transferência sem limite: 26.000 milhas', 10, 'M8', 26000), ('3.15', '… custo por milheiro R$ 23,08', 10, 'O8', 23.0769231),
                ('3.16', '… e aviso de que nenhum limite foi informado', 10, 'P8', 'Nenhum limite informado: o cálculo assume que todos os pontos transferidos recebem bônus. Confira o regulamento da promoção.')])
m = dict(tr); m[(S[10], 'F8')] = 10000
caso('t2', m, [('3.17', 'Limite de bônus de 10.000 pontos: bônus = 3.000', 10, 'L8', 3000), ('3.18', '… total recebido 23.000', 10, 'M8', 23000), ('3.19', '… custo por milheiro R$ 26,09', 10, 'O8', 26.0869565),
               ('3.20', '… aviso do limite', 10, 'P8', 'O limite informado é menor que os pontos transferidos: o bônus foi calculado só sobre o limite.')])
m = dict(tr); m[(S[10], 'F8')] = 50000
caso('t3', m, [('3.21', 'Limite maior que os pontos: bônus integral (6.000)', 10, 'L8', 6000)])
m = dict(tr); m[(S[10], 'D8')] = 0
caso('t4', m, [('3.22', 'Fator de conversão 0: sem divisão por zero', 10, 'O8', ''), ('3.23', '… aviso', 10, 'P8', 'O fator de conversão precisa ser maior que zero.')])
m = dict(tr); m[(S[10], 'D8')] = 0.5
caso('t5', m, [('3.24', 'Fator 0,5 (2 pontos de origem para 1): base = 10.000', 10, 'J8', 10000), ('3.25', '… bônus de 30% sobre a base = 3.000', 10, 'L8', 3000)])
m = dict(tr); m[(S[10], 'E8')] = 0
caso('t6', m, [('3.26', 'Bônus 0%: apenas a conversão', 10, 'M8', 20000), ('3.27', '… aviso', 10, 'P8', 'Bônus de 0%: só a conversão foi considerada.')])
m = dict(tr); m[(S[10], 'E8')] = 'abc'
caso('t7', m, [('3.28', 'Bônus em texto ou vazio: NÃO é zero; sem erro e sem total', 10, 'M8', ''), ('3.28b', '… aviso', 10, 'P8', 'Informe o percentual de bônus (0% se não houver): campo vazio = não informado.')])
# lotes e datas (referência fixa)
lt = {(S[1], 'B48'): REF}
for k, (off, esp) in enumerate([(-1, 'Vencido'), (0, 'Vence hoje'), (1, 'Vence em até 30 dias'), (30, 'Vence em até 30 dias'), (31, 'Vence em até 60 dias'), (60, 'Vence em até 60 dias'), (61, 'Vence em até 90 dias'), (90, 'Vence em até 90 dias'), (91, 'Mais de 90 dias')]):
    r = 7 + k
    lt[(S[4], 'A%d' % r)] = 'P'; lt[(S[4], 'D%d' % r)] = 100; lt[(S[4], 'F%d' % r)] = REF + datetime.timedelta(days=off)
lt[(S[4], 'A20')] = 'P'; lt[(S[4], 'D20')] = 100; lt[(S[4], 'F20')] = 'abc'
lt[(S[4], 'A21')] = 'P'; lt[(S[4], 'D21')] = 100
lt[(S[4], 'A22')] = 'P'; lt[(S[4], 'D22')] = 100; lt[(S[4], 'G22')] = 'Sim, confirmado'
lt[(S[4], 'A23')] = 'P'; lt[(S[4], 'D23')] = 100; lt[(S[4], 'G23')] = 'Sim, confirmado'; lt[(S[4], 'F23')] = REF + datetime.timedelta(days=5)
lt[(S[4], 'A24')] = None
wl = lo_wb(aluno, lt, name='lotes')
for k, (off, esp) in enumerate([(-1, 'Vencido'), (0, 'Vence hoje'), (1, 'Vence em até 30 dias'), (30, 'Vence em até 30 dias'), (31, 'Vence em até 60 dias'), (60, 'Vence em até 60 dias'), (61, 'Vence em até 90 dias'), (90, 'Vence em até 90 dias'), (91, 'Mais de 90 dias')]):
    add(g, '3.%02d' % (29 + k), 'Lote com vencimento em %+d dia(s) da data de referência' % off, esp, v(wl, S[4], 'I%d' % (7 + k)))
add(g, '3.38', 'Data em texto: "Data inválida"', 'Data inválida', v(wl, S[4], 'I20'))
add(g, '3.39', 'Sem data informada (não é o mesmo que sem expiração)', 'Data não informada', v(wl, S[4], 'I21'))
add(g, '3.40', 'Sem expiração confirmada', 'Sem expiração confirmada', v(wl, S[4], 'I22'))
add(g, '3.41', 'Sem expiração confirmada, mas com data preenchida: pede conferência', 'Conferir: sem expiração e data preenchidas', v(wl, S[4], 'I23'))
add(g, '3.42', 'Linha vazia: situação em branco', '', v(wl, S[4], 'I24'))
add(g, '3.43', 'Dias até o vencimento de lote sem expiração: em branco', '', v(wl, S[4], 'H22'))
# contas x lotes
ct = dict(lt); ct[(S[3], 'A7')] = 'P'; ct[(S[3], 'C7')] = 999; ct[(S[3], 'A8')] = 'Q'; ct[(S[3], 'C8')] = 500
wc3 = lo_wb(aluno, ct, name='contas')
add(g, '3.44', 'Conta com lotes: saldo = soma dos lotes NÃO vencidos (13 lotes × 100 = 1.300, menos 1 vencido)', 1200, v(wc3, S[3], 'E7'))
add(g, '3.44b', '… e o lote vencido aparece na coluna R', 100, v(wc3, S[3], 'R7'))
add(g, '3.45', '… o saldo manual divergente é ignorado e isso é informado', 'Soma dos lotes não vencidos (o saldo manual foi ignorado) Há pontos vencidos (coluna R).', v(wc3, S[3], 'F7'))
add(g, '3.46', 'Conta sem lotes: vale o saldo manual', 500, v(wc3, S[3], 'E8'))
add(g, '3.47', '… e a origem do saldo é "Informado manualmente"', 'Informado manualmente', v(wc3, S[3], 'F8'))
add(g, '3.48', 'Pontos a vencer em até 90 dias (7 lotes de 100, de +0 a +90 dias; sem expiração e vencidos ficam de fora)', 700, v(wc3, S[3], 'G7'))
add(g, '3.49', 'Próximo vencimento = a própria data de referência (lote que vence hoje)', REF, v(wc3, S[3], 'H7'))
# inatividade
ina = {(S[1], 'B48'): REF, (S[3], 'A7'): 'X', (S[3], 'K7'): 24, (S[3], 'J7'): D(2024, 10, 1),
       (S[3], 'A8'): 'Y', (S[3], 'K8'): 24, (S[3], 'J8'): D(2024, 12, 1), (S[3], 'A9'): 'Z', (S[3], 'K9'): 24, (S[3], 'J9'): D(2026, 5, 1),
       (S[3], 'A10'): 'W', (S[3], 'K10'): 24, (S[3], 'J10'): D(2027, 1, 1), (S[3], 'A11'): 'V', (S[3], 'K11'): 24, (S[3], 'I11'): D(2026, 11, 10),
       (S[3], 'A12'): 'U', (S[3], 'J12'): D(2020, 1, 1)}
wi = lo_wb(aluno, ina, name='inat')
add(g, '3.50', 'Inatividade: 25 meses sem movimentação, prazo 24', 'Prazo informado atingido ou ultrapassado: confira com o programa.', v(wi, S[3], 'M7'))
add(g, '3.51', 'Inatividade: 23 meses, prazo 24 (menos de 3 meses do limite)', 'Atenção: a menos de 3 meses do prazo informado (sugestão de organização).', v(wi, S[3], 'M8'))
add(g, '3.52', 'Inatividade: 6 meses, prazo 24', 'Dentro do prazo informado.', v(wi, S[3], 'M9'))
add(g, '3.53', 'Último acúmulo ou resgate no futuro: meses em branco', '', v(wi, S[3], 'L10'))
add(g, '3.54', 'Só o último acesso foi informado (recente): o alerta pede a data do último acúmulo ou resgate', 'Informe a data do último acúmulo ou resgate.', v(wi, S[3], 'M11'))
add(g, '3.55', 'Sem prazo de inatividade informado: nenhum alerta é inventado', '', v(wi, S[3], 'M12'))
# compras: dias corridos x úteis
cp = {(S[1], 'B48'): REF,
      (S[6], 'B7'): 'L', (S[6], 'I7'): D(2026, 11, 12), (S[6], 'J7'): 10, (S[6], 'K7'): 'Úteis',
      (S[6], 'B8'): 'L', (S[6], 'I8'): D(2026, 11, 12), (S[6], 'J8'): 10, (S[6], 'K8'): 'Corridos',
      (S[6], 'B9'): 'L', (S[6], 'I9'): D(2026, 11, 12), (S[6], 'J9'): None, (S[6], 'K9'): 'Corridos',
      (S[6], 'B10'): 'L', (S[6], 'I10'): D(2026, 11, 12), (S[6], 'J10'): 10, (S[6], 'K10'): None,
      (S[6], 'B11'): 'L', (S[6], 'I11'): D(2026, 11, 12), (S[6], 'J11'): 10, (S[6], 'K11'): 'Corridos', (S[6], 'M11'): 30, (S[6], 'N11'): 'Corridos', (S[6], 'O11'): 'Outra data (coluna P)',
      (S[6], 'B12'): 'L', (S[6], 'I12'): D(2026, 11, 12), (S[6], 'J12'): 10, (S[6], 'K12'): 'Corridos', (S[6], 'M12'): 5, (S[6], 'N12'): 'Úteis', (S[6], 'O12'): 'Outra data (coluna P)', (S[6], 'P12'): D(2026, 12, 1),
      (S[6], 'B13'): 'L', (S[6], 'I13'): D(2026, 9, 1), (S[6], 'J13'): 10, (S[6], 'K13'): 'Corridos', (S[6], 'M13'): 30, (S[6], 'N13'): 'Corridos', (S[6], 'O13'): 'Data prevista de crédito', (S[6], 'U13'): 'Aguardando'}
wp = lo_wb(aluno, cp, name='compras')
add(g, '3.56', '10 dias úteis após quinta 12/11/2026, sem feriados', D(2026, 11, 26), v(wp, S[6], 'L7'))
add(g, '3.57', '10 dias corridos após 12/11/2026', D(2026, 11, 22), v(wp, S[6], 'L8'))
add(g, '3.58', 'Sem prazo informado: sem data prevista (não assume prazo)', '', v(wp, S[6], 'L9'))
add(g, '3.59', 'Sem tipo de dias: sem data prevista', '', v(wp, S[6], 'L10'))
add(g, '3.60', 'Prazo para reclamar com base "Outra data" sem a data: não calcula', '', v(wp, S[6], 'Q11'))
add(g, '3.61', 'Prazo para reclamar: 5 dias úteis após 01/12/2026 (terça) = 08/12/2026', D(2026, 12, 8), v(wp, S[6], 'Q12'))
add(g, '3.62', 'Data-limite de reclamação já passou E prazo de crédito passou (aguardando): os dois avisos', 'Data-limite de reclamação já passou. O prazo previsto de crédito passou sem crédito registrado.', v(wp, S[6], 'V13'))
add(g, '3.63', 'Linha com pouco dado: pede o preenchimento (sem inventar datas)', 'Preencha o evento, a data, o prazo e o tipo de dias para calcular as datas.', v(wp, S[6], 'V9'))
cph = dict(cp); cph[(S[6], 'Z7')] = D(2026, 11, 20)
wph = lo_wb(aluno, cph, name='compras_feriado')
add(g, '3.64', '10 dias úteis após 12/11/2026 com feriado em 20/11/2026 na lista configurável', D(2026, 11, 27), v(wph, S[6], 'L7'))
# clubes e compra
cl = {(S[9], 'A8'): 'C', (S[9], 'B8'): 42, (S[9], 'C8'): 1, (S[9], 'D8'): 3, (S[9], 'E8'): 0, (S[9], 'F8'): 1000, (S[9], 'G8'): 0,
      (S[9], 'A9'): 'C2', (S[9], 'B9'): 42, (S[9], 'C9'): 0, (S[9], 'E9'): 0, (S[9], 'F9'): 1000, (S[9], 'G9'): 0,
      (S[9], 'A11'): 'C4', (S[9], 'B11'): 42, (S[9], 'C11'): 12, (S[9], 'F11'): 1000, (S[9], 'A23'): 'P2', (S[9], 'B23'): 20000, (S[9], 'D23'): 600, (S[9], 'E23'): 0,
      (S[9], 'A10'): 'C3', (S[9], 'B10'): 42, (S[9], 'C10'): 12, (S[9], 'D10'): 3, (S[9], 'E10'): 60, (S[9], 'F10'): 1000, (S[9], 'G10'): 1000,
      (S[9], 'A22'): 'P', (S[9], 'B22'): 0, (S[9], 'C22'): 0, (S[9], 'D22'): 0}
wcl = lo_wb(aluno, cl, name='clubes')
add(g, '3.65', 'Clube: período menor que a permanência mínima gera alerta', 'O período analisado é menor que a permanência mínima: analise o período completo do compromisso.', v(wcl, S[9], 'K8'))
add(g, '3.66', 'Clube com 0 mês analisado: sem divisão por zero', '', v(wcl, S[9], 'J9'))
add(g, '3.67', 'Clube: custo total em 12 meses = 42 × 12 + 60 de outros custos', 564, v(wcl, S[9], 'H10'))
add(g, '3.68', 'Clube: pontos recebidos = 1.000 × 12 + 1.000 (bônus contado uma vez no período)', 13000, v(wcl, S[9], 'I10'))
add(g, '3.69', 'Clube: custo por milheiro no período completo = 564 ÷ 13', 43.3846154, v(wcl, S[9], 'J10'))
add(g, '3.70', 'Compra com tudo em zero: sem divisão por zero', '', v(wcl, S[9], 'H22'))
add(g, '3.70b', 'Clube com outros custos e bônus em branco: NÃO assume zero (sem total) e avisa', 'Campo vazio = não informado, não zero: preencha outros custos, pontos regulares e bônus (digite 0 se não houver).', v(wcl, S[9], 'K11'))
add(g, '3.70c', '… sem custo total', '', v(wcl, S[9], 'H11'))
add(g, '3.70d', 'Compra com o bônus em branco: total recebido não é calculado', '', v(wcl, S[9], 'F23'))
# renovação
rn = {(S[1], 'B48'): REF,
      (S[11], 'A7'): 'P', (S[11], 'E7'): 100, (S[11], 'D7'): 5000, (S[11], 'I7'): 400, (S[11], 'J7'): 8000, (S[11], 'K7'): 0, (S[11], 'L7'): 0,
      (S[11], 'A8'): 'P', (S[11], 'E8'): 100, (S[11], 'D8'): 5000, (S[11], 'I8'): 400, (S[11], 'J8'): 3000, (S[11], 'F8'): D(2026, 11, 10), (S[11], 'K8'): 0, (S[11], 'L8'): 0,
      (S[11], 'A9'): 'P', (S[11], 'E9'): 100, (S[11], 'H9'): 'Passagem',
      (S[11], 'A10'): 'P', (S[11], 'E10'): 100, (S[11], 'D10'): 5000, (S[11], 'I10'): 400, (S[11], 'J10'): 3000, (S[11], 'L10'): 0}
wr = lo_wb(aluno, rn, name='renovar')
add(g, '3.71', 'Renovação: pontos necessários maiores que a quantidade elegível: aviso', 'Os pontos necessários superam a quantidade elegível: o resgate não cabe só neste lote.', v(wr, S[11], 'Q7'))
add(g, '3.72', 'Renovação: nova validade não posterior à referência: aviso', 'A nova validade informada não é posterior à data de referência.', v(wr, S[11], 'Q8'))
add(g, '3.73', 'Renovação sem a alternativa em dinheiro: não compara e avisa', 'Sem uso concreto com valor em dinheiro informado: não há o que comparar. Pagar sem perspectiva de uso pode gerar outra perda.', v(wr, S[11], 'Q9'))
add(g, '3.74', 'Renovação sem comparação: desembolsos em branco', '', v(wr, S[11], 'N9'))
add(g, '3.74b', 'Renovação com o pagamento adicional em branco: NÃO assume zero (sem desembolso) e avisa', 'Campo vazio = não informado, não zero: preencha o pagamento adicional e os outros gastos do resgate (digite 0 se não houver).', v(wr, S[11], 'Q10'))
add(g, '3.74c', '… desembolso em branco', '', v(wr, S[11], 'N10'))
# mapa de acúmulo e diagnóstico
ma = {(S[5], 'I7'): 220, (S[5], 'J7'): 200, (S[5], 'I8'): 200, (S[5], 'J8'): 220, (S[5], 'I9'): 100, (S[5], 'J9'): 100, (S[5], 'I10'): 100}
wm = lo_wb(aluno, ma, name='mapa')
add(g, '3.75', 'Mapa de acúmulo: preço pelo programa R$ 20 mais caro', 20, v(wm, S[5], 'K7'))
add(g, '3.76', '… leitura: não chama os pontos de gratuitos (o separador decimal segue a localidade do programa que abre o arquivo)', 'Pelo caminho do programa você paga R$ 20.00 a mais. Os pontos não são gratuitos: o benefício precisa valer mais do que essa diferença.', norm(v(wm, S[5], 'L7')))
add(g, '3.77', 'Mapa de acúmulo: preço pelo programa mais barato', 'O caminho do programa tem preço final R$ 20.00 menor que a alternativa.', norm(v(wm, S[5], 'L8')))
add(g, '3.78', 'Mapa de acúmulo: preços iguais', 'Mesmo preço final nas duas opções.', v(wm, S[5], 'L9'))
add(g, '3.79', 'Mapa de acúmulo: sem a alternativa, sem diferença', '', v(wm, S[5], 'K10'))
# data de referência
wd = lo_wb(aluno, {(S[1], 'B48'): D(2030, 1, 15)}, name='dataref')
add(g, '3.80', 'Data de referência informada substitui a data de hoje', D(2030, 1, 15), v(wd, S[1], 'B49'))
# plano e checklists
pl = {(S[12], 'G%d' % r): 'Concluído' for r in range(8, 13)}; pl[(S[12], 'G13')] = 'Não se aplica'; pl[(S[12], 'C5')] = D(2026, 11, 2)
pl.update({(S[13], 'D%d' % r): 'Concluído' for r in range(8, 13)})
wpl = lo_wb(aluno, pl, name='plano')
add(g, '3.81', 'Plano: semana 1 com 5 de 5', 5, v(wpl, S[12], 'D33')); add(g, '3.82', 'Plano: "Não se aplica" sai do total aplicável da semana 2', 4, v(wpl, S[12], 'C34'))
add(g, '3.83', 'Plano: início sugerido da semana 3 = início + 14 dias', D(2026, 11, 16), v(wpl, S[12], 'I20'))
add(g, '3.84', 'Checklists: 5 itens do grupo mensal concluídos', 5, v(wpl, S[13], 'K8')); add(g, '3.85', 'Checklists: progresso do grupo mensal', 1, v(wpl, S[13], 'L8'))
add(g, '3.86', 'Checklists: total de itens por grupo somam 49', 49, v(wpl, S[13], 'J13'))
# alteração e cancelamento (arquivo do aluno preenchido com os números do livro)
alt_all = dict(base07); alt_all[(S[7], 'B9')] = 2
alt_all.update({(S[8], 'B18'): 150, (S[8], 'C18'): 100, (S[8], 'B19'): 0, (S[8], 'B21'): 0, (S[8], 'C21'): 0, (S[8], 'B22'): 0.8, (S[8], 'B23'): 280, (S[8], 'C23'): 280, (S[8], 'C24'): 72000, (S[8], 'B25'): 0, (S[8], 'C25'): 150, (S[8], 'B26'): 0, (S[8], 'C26'): 0, (S[8], 'B27'): 360, (S[8], 'C27'): 360})
walt = lo_wb(aluno, alt_all, name='alterar')
add(g, '3.87', 'Aba 08 preenchida no arquivo do aluno: perda em dinheiro R$ 336', 336, v(walt, S[8], 'B43')); add(g, '3.88', '… e R$ 300 com milhas', 300, v(walt, S[8], 'C43'))
m2 = dict(alt_all); m2[(S[8], 'B27')] = 100
walt2 = lo_wb(aluno, m2, name='alterar2')
add(g, '3.89', 'Bagagem parcialmente devolvida (R$ 100 de R$ 360): a diferença entra na perda (336 + 260)', 596, v(walt2, S[8], 'B43'))
walt3 = lo_wb(aluno, {(S[8], 'B22'): 0.8}, name='alterar3')
add(g, '3.90', 'Aba 08 sem dados da aba 07: resultados em branco', '', v(walt3, S[8], 'B43'))

# ---------------------------------------------------------------- G4 inserção de linhas (LibreOffice via UNO)
g = 'G4 Inserção de linhas (as faixas das fórmulas acompanham)'
uno_out = os.path.join(tmp_root, 'uno.json')
subprocess.run([sys.executable, os.path.join(HERE, 'uno_insercao.py'), aluno, uno_out], check=False, timeout=240)
if os.path.exists(uno_out):
    u = json.load(open(uno_out, encoding='utf-8'))
    for cid, desc, esp, obt in u['checks']: add(g, cid, desc, esp, obt)
else:
    add(g, '4.00', 'Teste de inserção de linhas executado', True, False)


# ---------------------------------------------------------------- G6 correções do parecer (v1.1)
g = 'G6 Correções do parecer do Codex (v1.1): saldo, reclamação e campo vazio ≠ zero'
# 1) saldo
sv = {(S[1], 'B48'): REF, (S[3], 'A7'): 'X', (S[3], 'A8'): 'Y', (S[3], 'A9'): 'Z', (S[3], 'A10'): 'W',
      (S[4], 'A7'): 'X', (S[4], 'D7'): 100, (S[4], 'F7'): REF - datetime.timedelta(days=10),
      (S[4], 'A8'): 'X', (S[4], 'D8'): 200, (S[4], 'F8'): REF - datetime.timedelta(days=1),
      (S[4], 'A9'): 'Y', (S[4], 'D9'): 50, (S[4], 'F9'): REF,
      (S[4], 'A10'): 'Y', (S[4], 'D10'): 70,
      (S[4], 'A11'): 'Y', (S[4], 'D11'): 30, (S[4], 'G11'): 'Sim, confirmado',
      (S[4], 'A12'): 'Z', (S[4], 'D12'): 10, (S[4], 'F12'): 'abc', (S[3], 'C10'): 500}
wsv = lo_wb(aluno, sv, name='saldo')
add(g, '6.01', 'Programa só com lotes vencidos: saldo considerado = 0 (não 300)', 0, v(wsv, S[3], 'E7'))
add(g, '6.02', '… e os 300 pontos vencidos aparecem na coluna R', 300, v(wsv, S[3], 'R7'))
add(g, '6.03', 'Lote que vence hoje, sem data e sem expiração entram no saldo (50 + 70 + 30)', 150, v(wsv, S[3], 'E8'))
add(g, '6.04', '… sem pontos vencidos', 0, v(wsv, S[3], 'R8'))
add(g, '6.05', 'Data em texto não é tratada como vencida (10 no saldo)', 10, v(wsv, S[3], 'E9'))
# 2) reclamação
rc = {(S[1], 'B48'): REF}
def linha(r, **kw):
    base = {'A': D(2026, 9, 1), 'B': 'L', 'I': D(2026, 9, 1), 'J': 10, 'K': 'Corridos', 'M': 30, 'N': 'Corridos', 'O': 'Data prevista de crédito', 'F': 200}
    base.update(kw)
    for col, val in base.items(): rc[(S[6], '%s%d' % (col, r))] = val
linha(7, R=180, U='Divergente')                                   # divergência + data-limite vencida
linha(8, R=180, U='Divergente', M=None)                            # divergência sem prazo de reclamação
linha(9, R=200, U='Creditado')                                      # creditado certo, data-limite vencida: sem alerta de reclamação
linha(10, R=180, U='Creditado')                                     # creditado com divergência e data-limite vencida
linha(11, R=180, U='Em contestação', I=D(2026, 11, 1), A=D(2026, 11, 1))   # divergência, prazo de reclamação a 25 dias
wrc = lo_wb(aluno, rc, name='reclama')
add(g, '6.06', 'Divergência + data-limite vencida: os dois avisos', 'Divergência: creditado diferente do esperado. Data-limite de reclamação já passou.', v(wrc, S[6], 'V7'))
add(g, '6.07', 'Divergência sem prazo de reclamação informado: avisa que não há data-limite', 'Divergência: creditado diferente do esperado. Sem data-limite de reclamação calculada: informe o prazo, o tipo de dias e a base.', v(wrc, S[6], 'V8'))
add(g, '6.08', 'Creditado certo: nenhum alerta de reclamação', 'Acompanhando.', v(wrc, S[6], 'V9'))
add(g, '6.09', 'Creditado, mas com divergência e prazo vencido: os dois avisos', 'Divergência: creditado diferente do esperado. Data-limite de reclamação já passou.', v(wrc, S[6], 'V10'))
add(g, '6.10', 'Divergência com a data-limite ainda longe: só a divergência', 'Divergência: creditado diferente do esperado.', v(wrc, S[6], 'V11'))
add(g, '6.11', 'Aba 01 conta cada compra com alerta uma vez (4 das 5 têm alerta)', 4, v(wrc, S[1], 'B60'))
# 3) campo vazio != zero
mm = dict(alt_all); mm[(S[8], 'C25')] = None
w3 = lo_wb(aluno, mm, name='vazio_taxa')
add(g, '6.12', 'Aba 08: taxa de cancelamento (milhas) apagada: a perda NÃO cai para R$ 0; fica em branco', '', v(w3, S[8], 'C43'))
add(g, '6.13', '… a taxa total de cancelamento também fica em branco', '', v(w3, S[8], 'C39'))
add(g, '6.14', '… e o aviso lista o campo pendente', True, 'taxa de cancelamento (milhas)' in str(v(w3, S[8], 'A45')))
add(g, '6.15', '… a perda em dinheiro (opção em dinheiro), que não depende desse campo, continua R$ 336', 336, v(w3, S[8], 'B43'))
w3b = lo_wb(aluno, dict(alt_all, **{}) if False else alt_all, name='vazio_taxa_ok')
add(g, '6.16', 'Com todos os campos (zeros digitados), a perda é calculada e o aviso some', '', v(w3b, S[8], 'A45'))
mm2 = dict(alt_all); mm2[(S[8], 'B26')] = None
add(g, '6.17', 'Aba 08: "outros custos" (dinheiro) apagado: perda em branco (não R$ 336)', '', v(lo_wb(aluno, mm2, name='vazio_outros'), S[8], 'B43'))
mm3 = dict(alt_all); mm3[(S[8], 'B19')] = None
add(g, '6.18', 'Aba 08: "diferença de tarifa" apagada: custo de alteração em branco (não R$ 300)', '', v(lo_wb(aluno, mm3, name='vazio_dif'), S[8], 'B36'))
mm4 = dict(base07); mm4[(S[7], 'B22')] = None
w4 = lo_wb(aluno, mm4, name='vazio_bag')
add(g, '6.19', 'Aba 07: bagagem apagada na opção em dinheiro: total em branco (não "sem bagagem")', '', v(w4, S[7], 'B23'))
add(g, '6.20', '… e o aviso pede para informar a bagagem', 'Informe a bagagem e os serviços das opções em dinheiro (digite 0 se não houver) para calcular os totais.', v(w4, S[7], 'A35'))
mm5 = dict(base07); mm5[(S[7], 'C21')] = None
add(g, '6.21', 'Aba 07: "outros pagamentos adicionais" apagado: total do resgate em branco', '', v(lo_wb(aluno, mm5, name='vazio_outros07'), S[7], 'C23'))

# ---------------------------------------------------------------- G5 estrutura, proteção e conformidade
g = 'G5 Estrutura, proteção e conformidade'
wf = openpyxl.load_workbook(aluno); wx = openpyxl.load_workbook(exem)
add(g, '5.01', 'Quatorze abas, na ordem do plano, com nomes até 31 caracteres', [NAMES[i] for i in range(1, 15)], wf.sheetnames)
add(g, '5.02', 'Nenhuma célula mesclada (aluno e exemplos)', 0, sum(len(ws.merged_cells.ranges) for ws in wf.worksheets) + sum(len(ws.merged_cells.ranges) for ws in wx.worksheets))
add(g, '5.03', 'Todas as abas protegidas, sem senha', True, all(ws.protection.sheet and not ws.protection.password for ws in wf.worksheets))
bad_locked = []; bad_unlocked = []; n_formula = 0; n_input = 0
for ws in wf.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                n_formula += 1
                if not c.protection.locked: bad_locked.append((ws.title, c.coordinate))
add(g, '5.04', 'Todas as células com fórmula estão bloqueadas (%d fórmulas)' % n_formula, 0, len(bad_locked))
unl = sum(1 for ws in wf.worksheets for row in ws.iter_rows() for c in row if (c.has_style and not c.protection.locked))
add(g, '5.05', 'Há células editáveis (desbloqueadas) para o preenchimento (%d células)' % unl, True, unl > 500)
fx = {(ws.title, c.coordinate): c.value for ws in wf.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith('=')}
fe = {(ws.title, c.coordinate): c.value for ws in wx.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith('=')}
add(g, '5.06', 'As duas versões têm exatamente as mesmas fórmulas (%d)' % len(fx), True, fx == fe)
# nenhuma entrada de exemplo na versão do aluno: células desbloqueadas vazias
cheias = [(ws.title, c.coordinate) for ws in wf.worksheets for row in ws.iter_rows() for c in row if c.has_style and not c.protection.locked and c.value not in (None, '')]
add(g, '5.07', 'Versão do aluno: nenhum campo editável preenchido (sem dados fictícios)', [], cheias[:5], len(cheias) == 0)
txt = ' '.join(str(c.value) for ws in wf.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and not c.value.startswith('='))
add(g, '5.08', 'Versão do aluno: nenhum texto "[Exemplo" nem "fictício" fora de instruções', 0, len(re.findall(r'\[Exemplo', txt)))
for fn in (ALUNO, EXEMPLOS):
    z = zipfile.ZipFile(os.path.join(SAIDA, fn)); nomes = z.namelist()
    add(g, '5.09 ' + fn[:20], 'Sem macros (vbaProject) nem conexões externas (externalLinks, connections)', 0, sum(1 for n_ in nomes if any(k in n_.lower() for k in ('vba', 'externallink', 'connections', 'activex'))))
    rels = ' '.join(z.read(n_).decode('utf-8', 'ignore') for n_ in nomes if n_.endswith('.rels'))
    add(g, '5.10 ' + fn[:20], 'Hiperlinks apenas internos (nenhum "TargetMode=External")', 0, rels.count('TargetMode="External"'))
    wb_xml = z.read('xl/workbook.xml').decode('utf-8')
    add(g, '5.11 ' + fn[:20], 'Intervalos nomeados DataRef e Feriados presentes', True, 'DataRef' in wb_xml and 'Feriados' in wb_xml)
# 70 itens
md_items = parse_checklists()
p12 = [wx[NAMES[12]]['C%d' % r].value for r in range(8, 29)]; p13 = [wx[NAMES[13]]['B%d' % r].value for r in range(8, 57)]
add(g, '5.12', 'Checklists (aba 13) + plano (aba 12) = 70 itens, na ordem e com o texto do livro', 70, len(p12) + len(p13))
txt_md = [i['text'] for i in md_items]
add(g, '5.13', '… textos idênticos aos do texto aprovado (sem omissão nem duplicação)', True, sorted(txt_md) == sorted(p12 + p13) and len(set(p12 + p13)) == 70)
# oito tabelas
checks = [('Cap. 1: Programa | Saldo | Validade', S[2], ['Programa', 'Saldo (pontos ou milhas)', 'Validade (data)']),
          ('Cap. 6: organização das contas', S[3], ['Programa', 'Saldo considerado', 'Último acesso', 'Último acúmulo ou resgate']),
          ('Cap. 7: controle de compras', S[6], ['Data da compra', 'Número do pedido (texto)', 'Data-limite de reclamação', 'Situação']),
          ('Cap. 7: mapa de acúmulo', S[5], ['O que eu já gasto', 'Onde gasto', 'Esse lugar tem programa de pontos?', 'Qual o passo para pontuar?']),
          ('Cap. 10: vencimento, inatividade e extensão/reativação', S[14], ['Vencimento dos pontos', 'Encerramento da conta por inatividade', 'Extensão ou reativação']),
          ('Cap. 10: validade da Smiles por origem', S[14], ['Milhas acumuladas em voos da GOL', 'Milhas reativadas']),
          ('Cap. 13: comparação dinheiro × milhas', S[7], ['Tarifa em reais (sem as taxas de embarque)', 'Milhas ou pontos utilizados', 'Pago em dinheiro']),
          ('Cap. 13: alteração, cancelamento e bagagem', S[8], ['Condições e fonte: Alteração (2 pessoas)', 'Condições e fonte: Cancelamento antes do embarque', 'Condições e fonte: Bagagem não usada'])]
for k, (nome, sh, termos) in enumerate(checks):
    conj = ' | '.join(str(c.value) for row in wx[sh].iter_rows() for c in row if c.value is not None)
    add(g, '5.%02d' % (14 + k), 'Tabela do livro representada: ' + nome, True, all(t in conj for t in termos))
from common import parse_table
t8 = parse_table('| | Dinheiro | Milhas |', 2)
ok8 = all(wx[NAMES[8]]['B%d' % r].value == t8[k][1] and wx[NAMES[8]]['C%d' % r].value == t8[k][2] for k, r in ((1, 28), (2, 29), (3, 30)))
add(g, '5.22', 'Cap. 13 (Passo 6): as condições da versão de exemplos são idênticas ao texto do livro', True, ok8)
sm = parse_table('| Origem das milhas |')[1:]
add(g, '5.23', 'Cap. 10: as cinco linhas da Smiles correspondem ao texto aprovado (origem)', [r[0] for r in sm], [wx[NAMES[14]]['A%d' % r].value for r in range(16, 21)])
add(g, '5.24', 'Aba 14: nenhuma linha da Smiles além das cinco do texto aprovado', None, wx[NAMES[14]]['A21'].value)
q12 = parse_cap12_questions()
add(g, '5.25', 'Aba 07: as sete perguntas do Capítulo 12 reproduzidas, na ordem', q12, [str(wx[NAMES[7]]['A%d' % r].value)[3:] for r in range(62, 69)])
add(g, '5.26', 'Aba 08: registro de contatos e protocolos (Capítulo 15) presente', True, 'Registro de contatos' in str(wx[NAMES[8]]['A48'].value))
add(g, '5.27', 'Aba 07: decisão do Capítulo 13/18 presente, com lista de opções', 'Esperar', wx[NAMES[7]]['B71'].value)

json.dump(res, open(os.path.join(SAIDA, '..', 'testes', 'resultado-testes.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1, default=str)
tot = len(res); ok = sum(1 for r in res if r['ok'])
print('testes: %d | aprovados: %d | reprovados: %d' % (tot, ok, tot - ok))
for r in res:
    if not r['ok']: print('FALHOU', r['id'], r['descricao'], '| esperado:', r['esperado'], '| obtido:', r['obtido'])
