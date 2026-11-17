# -*- coding: utf-8 -*-
"""Abas 12 a 14."""
import datetime, re
from common import *
from sheets_b import band

D = datetime.date

# ===================================================================== 12
def sheet12(book):
    s = book.sheet(12); ws = s.ws
    for c, w in {'A': 10, 'B': 38, 'C': 82, 'D': 16, 'E': 14, 'F': 14, 'G': 16, 'H': 40, 'I': 18}.items(): ws.set_column('%s:%s' % (c, c), w)
    s.title('Plano de 30 dias', 'Capítulo 18 — Quatro semanas, na ordem do livro',
            'Ajuste ao seu tempo: o importante é seguir a ordem. Compras, assinatura de clube e resgate NÃO são obrigações do plano: marque "Não se aplica" quando não fizer sentido para você.')
    s.back_link('A4')
    s.label('A5', 'Data de início do plano (opcional)', bold=True, wrap=False); s.inp('C5', 'date', ex=D(2026, 10, 5))
    ws.write('B5', '', book.f())
    s.t('D5', 'Com a data de início, a coluna I sugere o começo de cada semana.', italic=True, font_color=GRAY)
    s.validate_date('C5')
    s.head(7, ['Semana', 'Tema da semana', 'Atividade', 'Capítulo de referência', 'Data planejada', 'Data de conclusão', 'Situação', 'Observações', 'Início sugerido da semana'], height=40, calc_cols=(8,))
    items = [i for i in parse_checklists() if i['cap'] == 18]
    assert len(items) == 21
    st = {}
    if book.ex:
        sit = ['Concluído'] * 5 + ['Concluído', 'Em andamento', 'Concluído', 'Concluído', 'Não se aplica'] + ['Concluído', 'Pendente', 'Pendente', 'Pendente', 'Pendente'] + ['Concluído', 'Concluído', 'Concluído', 'Concluído', 'Concluído', 'Não se aplica']
    for k, it in enumerate(items):
        r = 8 + k
        m = re.match(r'Semana (\d) — (.*)', it['h3']); sem = int(m.group(1)); tema = m.group(2)
        refs = ', '.join(re.findall(r'Cap[íi]tulos? (\d+)', it['raw'])) or '—'
        ws.write_number('A%d' % r, sem, book.f(bg_color=ROSE_L, border=1, border_color=BORDER, num_format='"Semana "0', bold=True, align='center'))
        ws.write('B%d' % r, tema, book.f(bg_color=ROSE_L, border=1, border_color=BORDER, text_wrap=True))
        ws.write('C%d' % r, it['text'], book.f(border=1, border_color=BORDER, text_wrap=True, bg_color='#FFFFFF'))
        ws.write('D%d' % r, refs, book.f(border=1, border_color=BORDER, align='center', bg_color='#FFFFFF'))
        pd = D(2026, 10, 5) + datetime.timedelta(days=7 * (sem - 1)) if book.ex else None
        s.inp('E%d' % r, 'date', ex=pd)
        s.inp('F%d' % r, 'date', ex=(pd + datetime.timedelta(days=3)) if (book.ex and sit[k] == 'Concluído') else None)
        s.inp('G%d' % r, 'text', ex=sit[k] if book.ex else None)
        s.inp('H%d' % r, 'long', ex='[Exemplo adicional fictício] Datas e situações inventadas, a partir do exemplo da Marina no Capítulo 18.' if (book.ex and k == 0) else ('[Exemplo adicional fictício] A Marina não pagou por clube nem por compra de pontos (Capítulo 18).' if (book.ex and k == 20) else None))
        s.calc('I%d' % r, '=IF(ISNUMBER($C$5),$C$5+7*($A{r}-1),"")'.format(r=r), 'date')
        ws.set_row(r - 1, 40 if len(it['text']) < 120 else 50)
    s.validate_list('G8:G28', items=['Pendente', 'Em andamento', 'Concluído', 'Não se aplica'])
    s.validate_date('E8:F28')
    s.cf('A8:H28', '=$G8="Concluído"', {'bg_color': GREEN_L}); s.cf('A8:H28', '=$G8="Em andamento"', {'bg_color': YELLOW_LL}); s.cf('A8:H28', '=$G8="Não se aplica"', {'bg_color': GRAY_L})
    band(s, 31, 'CDE', 'Progresso por semana (não conta atividades marcadas como "Não se aplica")', book)
    s.head(32, ['', 'Semana', 'Atividades aplicáveis', 'Concluídas', 'Progresso'], height=26)
    for w in range(1, 5):
        r = 32 + w
        s.static_calc('B%d' % r, 'Semana %d' % w)
        s.calc('C%d' % r, '=COUNTIF($A$8:$A$28,%d)-COUNTIFS($A$8:$A$28,%d,$G$8:$G$28,"Não se aplica")' % (w, w), 'int')
        s.calc('D%d' % r, '=COUNTIFS($A$8:$A$28,%d,$G$8:$G$28,"Concluído")' % w, 'int')
        s.calc('E%d' % r, '=IF($C{r}=0,"",$D{r}/$C{r})'.format(r=r), 'pct')
    s.static_calc('B37', 'Total'); s.calc('C37', '=SUM(C33:C36)', 'int', bold=True); s.calc('D37', '=SUM(D33:D36)', 'int', bold=True); s.calc('E37', '=IF($C$37=0,"",$D$37/$C$37)', 'pct', bold=True)
    s.finish(header_row=7, freeze=(7, 3), autofilter='A7:I28', area='A1:I37')
    return s

# ===================================================================== 13
GRUPOS = {11: 'Checklist mensal (Capítulo 11)', 14: 'Antes de confirmar o resgate (Capítulo 14)', 15: 'Checklist rápido para problemas (Capítulo 15)',
          17: 'Antes de pagar ou transferir (Capítulo 17)', 19: 'Checklist do iniciante (Capítulo 19)'}

def sheet13(book):
    s = book.sheet(13); ws = s.ws
    for c, w in {'A': 34, 'B': 86, 'C': 20, 'D': 14, 'E': 16, 'F': 40, 'G': 26, 'H': 3, 'I': 38, 'J': 10, 'K': 12, 'L': 12, 'M': 20}.items(): ws.set_column('%s:%s' % (c, c), w)
    s.title('Checklists', 'Capítulos 11, 14, 15, 17 e 19 — 49 itens, em cinco grupos',
            'Marque "Pendente" ou "Concluído" e registre a data da revisão. Os 21 itens do plano de 30 dias (Capítulo 18) estão na aba 12; juntos, somam os 70 itens do livro.')
    s.back_link('A4')
    s.head(7, ['Grupo', 'Item', 'Capítulo(s)', 'Situação', 'Data da revisão', 'Observações', 'Subgrupo (checklist do iniciante)'], height=36)
    items = [i for i in parse_checklists() if i['cap'] in GRUPOS]
    assert len(items) == 49
    for k, it in enumerate(items):
        r = 8 + k
        refs = sorted(set(re.findall(r'Cap(?:ítulos?|s?\.)\s*(\d+)', it['raw'])), key=int)
        cap = ', '.join(refs) if refs else str(it['cap'])
        if it['cap'] == 19 and refs: cap = ', '.join(refs)
        ws.write('A%d' % r, GRUPOS[it['cap']], book.f(bg_color=ROSE_L, border=1, border_color=BORDER, text_wrap=True, bold=True))
        ws.write('B%d' % r, it['text'], book.f(border=1, border_color=BORDER, text_wrap=True, bg_color='#FFFFFF'))
        ws.write('C%d' % r, cap, book.f(border=1, border_color=BORDER, align='center', bg_color='#FFFFFF', text_wrap=True))
        s.inp('D%d' % r, 'text', ex=None); s.inp('E%d' % r, 'date', ex=None); s.inp('F%d' % r, 'long', ex=None)
        ws.write('G%d' % r, it['bold'] if it['cap'] == 19 and it['bold'] else '', book.f(border=1, border_color=BORDER, text_wrap=True, bg_color='#FFFFFF'))
        ws.set_row(r - 1, 32 if len(it['text']) < 100 else 46)
    s.validate_list('D8:D56', items=['Pendente', 'Concluído']); s.validate_date('E8:E56')
    s.cf('A8:G56', '=$D8="Concluído"', {'bg_color': GREEN_L})
    # resumo por grupo
    ws.write('I6', 'Resumo por grupo', book.f(bold=True, font_color=COPPER, font_size=11))
    s.head(7, ['Grupo', 'Itens', 'Concluídos', 'Progresso', 'Data da última revisão do grupo'], col0=8, height=36, calc_cols=(9, 10, 11))
    for k, (cap, nome) in enumerate(GRUPOS.items()):
        r = 8 + k
        ws.write('I%d' % r, nome, book.f(bg_color=ROSE_L, border=1, border_color=BORDER, text_wrap=True, bold=True))
        s.calc('J%d' % r, '=COUNTIF($A$8:$A$56,$I%d)' % r, 'int')
        s.calc('K%d' % r, '=COUNTIFS($A$8:$A$56,$I%d,$D$8:$D$56,"Concluído")' % r, 'int')
        s.calc('L%d' % r, '=IF($J{r}=0,"",$K{r}/$J{r})'.format(r=r), 'pct')
        s.inp('M%d' % r, 'date', ex=None); ws.set_row(r - 1, 46) if r < 13 else None
    ws.write('I13', 'Total', book.f(bold=True, border=1, border_color=BORDER, bg_color=ROSE_L))
    s.calc('J13', '=SUM(J8:J12)', 'int', bold=True); s.calc('K13', '=SUM(K8:K12)', 'int', bold=True); s.calc('L13', '=IF($J$13=0,"",$K$13/$J$13)', 'pct', bold=True)
    s.validate_date('M8:M12')
    # extra: perguntas ao emissor (Capítulo 8) — não faz parte dos 70 itens
    band(s, 59, 'CDEF', 'Extra (não conta nos 70 itens do livro): perguntas ao emissor do cartão — Capítulo 8', book)
    s.head(60, ['Grupo', 'Pergunta', 'Capítulo', 'Respondida?', 'Data', 'Resposta do emissor e fonte'], height=30)
    qs = ['Ele pontua? Nem todo cartão pontua.', 'Para qual programa vão os pontos? O destino define o que você poderá fazer com eles.',
          'Qual é a taxa de pontuação: por real, por dólar ou por outro critério? Se houver conversão de moeda, qual cotação é usada?',
          'Quando os pontos são creditados? Em geral, não é no momento da compra.', 'Há anuidade? E ela é cobrada ou isenta com certo gasto?',
          'Há compras que não pontuam? Alguns tipos de gasto podem ficar de fora.', 'E pagar contas no cartão? Essa operação pode ter tarifas. Confira o custo e se ela gera pontos antes de usar essa opção.']
    for k, qq in enumerate(qs):
        r = 61 + k
        ws.write('A%d' % r, 'Perguntas ao emissor (Capítulo 8)', book.f(bg_color=ROSE_L, border=1, border_color=BORDER, text_wrap=True, bold=True))
        ws.write('B%d' % r, qq, book.f(border=1, border_color=BORDER, text_wrap=True, bg_color='#FFFFFF')); ws.write('C%d' % r, '8', book.f(border=1, border_color=BORDER, align='center', bg_color='#FFFFFF'))
        s.inp('D%d' % r, 'text'); s.inp('E%d' % r, 'date'); s.inp('F%d' % r, 'long'); ws.set_row(r - 1, 36)
    s.validate_list('D61:D67', items=['Sim', 'Não']); s.validate_date('E61:E67')
    s.t('A69', 'Use esta lista como roteiro de perguntas: confirme as respostas no site oficial do emissor e no regulamento do programa.', italic=True, font_color=GRAY)
    s.finish(header_row=7, freeze=(7, 1), autofilter='A7:G56', area='A1:M69')
    return s

# ===================================================================== 14
def sheet14(book):
    s = book.sheet(14); ws = s.ws
    for c, w in {'A': 42, 'B': 62, 'C': 58, 'D': 46, 'E': 18, 'F': 46}.items(): ws.set_column('%s:%s' % (c, c), w)
    s.title('Referências', 'Capítulo 10 e fontes do material — tabelas explicativas e exemplos datados',
            'Esta aba é de consulta: não há campos para preencher. Os exemplos da Smiles são exemplos datados da edição e NÃO preenchem os lotes automaticamente.')
    s.back_link('A4')
    band(s, 6, 'BC', 'A. Três coisas diferentes (Capítulo 10)', book)
    t = parse_table('| | O que é |')
    s.head(7, ['Conceito', 'O que é', 'Exemplo de pergunta que você deve fazer'], height=26)
    for k, row in enumerate(t[1:]):
        r = 8 + k
        ws.write('A%d' % r, row[0], book.f(bold=True, bg_color=ROSE_L, border=1, border_color=BORDER, text_wrap=True, valign='top'))
        ws.write('B%d' % r, row[1], book.f(border=1, border_color=BORDER, text_wrap=True, valign='top', bg_color='#FFFFFF')); ws.write('C%d' % r, row[2], book.f(border=1, border_color=BORDER, text_wrap=True, valign='top', bg_color='#FFFFFF'))
        ws.set_row(r - 1, 52)
    s.t('A11', 'Uma conta pode estar ativa e mesmo assim ter lotes vencendo, e um lote pode estar dentro do prazo numa conta que corre risco de encerramento por inatividade. Acompanhe os três.', italic=True, font_color=GRAY)
    band(s, 13, 'BCDEF', 'B. Validade de milhas da Smiles por origem (exemplos datados, Capítulo 10)', book)
    s.t('A14', 'Cada linha vale apenas para a origem indicada. A página lista outras origens que este material não cobre, e as milhas recebidas por transferência entre contas não constam aqui porque duas páginas oficiais da Smiles traziam prazos diferentes. Consulte a página oficial.', italic=True, font_color=GRAY, text_wrap=True)
    ws.set_row(13, 34)
    s.head(15, ['Origem das milhas', 'Categoria ou condição', 'Validade informada', 'Fonte', 'Data da consulta', 'Situação da informação'], height=30)
    sm = parse_table('| Origem das milhas |')[1:]
    assert len(sm) == 5
    cond = {0: 'Varia por categoria do cliente', 4: 'Conforme o plano'}
    val = {0: '6 anos (categorias Smiles e Prata), 8 anos (Ouro) e 20 anos (Diamante)', 4: '10 anos (planos 1.000, 2.000 e 5.000) ou milhas que não expiram (planos 7.000, 10.000 e 20.000)'}
    for k, row in enumerate(sm):
        r = 16 + k
        ws.write('A%d' % r, row[0], book.f(border=1, border_color=BORDER, text_wrap=True, valign='top', bold=True, bg_color=ROSE_L))
        ws.write('B%d' % r, cond.get(k, '—'), book.f(border=1, border_color=BORDER, text_wrap=True, valign='top', bg_color='#FFFFFF'))
        ws.write('C%d' % r, val.get(k, row[1]), book.f(border=1, border_color=BORDER, text_wrap=True, valign='top', bg_color='#FFFFFF'))
        ws.write('D%d' % r, 'Smiles, página de validade de milhas (www.smiles.com.br/portal/validade-de-milhas)', book.f(border=1, border_color=BORDER, text_wrap=True, valign='top', bg_color='#FFFFFF'))
        ws.write_datetime('E%d' % r, datetime.datetime(2026, 11, 16), book.f(border=1, border_color=BORDER, num_format='dd/mm/yyyy', align='center', valign='top', bg_color='#FFFFFF'))
        ws.write('F%d' % r, 'Exemplo histórico da edição do e-book; pendente de reconferência antes da publicação.', book.f(border=1, border_color=BORDER, text_wrap=True, valign='top', bg_color=GRAY_L))
        ws.set_row(r - 1, 48)
    band(s, 22, 'BCDEF', 'C. Fontes usadas neste material e abas relacionadas', book)
    s.head(23, ['Fonte', 'Endereço (texto, sem link automático)', 'Para que serve neste material', 'Quem consultou', 'Data da consulta', 'Abas relacionadas'], height=30)
    fontes = [('Resolução ANAC nº 400, texto compilado', 'https://www.anac.gov.br/assuntos/legislacao/legislacao-1/resolucoes/resolucoes-2016/resolucao-no-400-13-12-2016', 'Separação entre tarifa e taxas na base de multas (art. 9º, parágrafo único); regras gerais de desistência e de volta em voo doméstico (arts. 11 e 19). Capítulos 13 a 15.', 'Codex — revisão assistida por IA, e Claude', '16/11/2026', '08 Alterar'),
              ('Regulamento do programa Livelo', 'www3.livelo.com.br/regulamentos/programa-pontos-livelo', 'Exemplos reais de prazo de crédito, reclamação de pontos não creditados, validade, transferência e resgates de viagem. Capítulos 6, 7, 10, 14, 15 e 17.', 'Claude', '12/11/2026 e 16/11/2026', '03 Contas, 04 Lotes, 06 Compras, 08 Alterar, 10 Transferências'),
              ('Smiles, página de validade de milhas', 'www.smiles.com.br/portal/validade-de-milhas', 'Exemplos de validade por origem das milhas. Capítulos 10, 16 e 17.', 'Claude', '16/11/2026', '04 Lotes, 09 Clubes e compra, 14 Referências'),
              ('Smiles, juntar e usar milhas', 'www.smiles.com.br/juntar-e-usar-milhas', 'Exemplo de condições de compra de milhas e planos do Clube. Capítulo 17.', 'Claude', '12/11/2026', '09 Clubes e compra'),
              ('Esfera, página de loja parceira (campanha)', 'www.esfera.com.vc/p/multi/e000100435', 'Padrão de campanha: entrar pelo botão do parceiro e finalizar na mesma janela. Capítulo 7. Regra daquela campanha, que expira.', 'Claude', '12/11/2026', '05 Acúmulo, 06 Compras'),
              ('Livelo, página "Pontos Viram Dinheiro"', 'www.livelo.com.br/carteira/pontos-viram-dinheiro/institucional', 'Existência do serviço; as condições não foram lidas. Capítulo 7.', 'Claude', '12/11/2026', '05 Acúmulo')]
    for k, f in enumerate(fontes):
        r = 24 + k
        for c, v in zip('ABCDEF', f):
            ws.write('%s%d' % (c, r), v, book.f(border=1, border_color=BORDER, text_wrap=True, valign='top', bg_color='#FFFFFF' if c != 'A' else ROSE_L, bold=(c == 'A')))
        ws.set_row(r - 1, 62)
    s.t('A31', 'Todas as fontes devem ser reconferidas perto da publicação. Regras, prazos e validades mudam: confirme sempre no site oficial e registre a data da sua consulta.', bold=True, font_color=COPPER)
    s.finish(area='A1:F31')
    return s
