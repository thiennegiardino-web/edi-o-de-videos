# -*- coding: utf-8 -*-
"""Abas 01 a 06."""
import datetime
from common import *

D = datetime.date

# ===================================================================== 01
def sheet01(book):
    s = book.sheet(1); ws = s.ws
    ws.set_column('A:A', 34); ws.set_column('B:B', 26); ws.set_column('C:C', 100)
    s.t('A1', 'DE MALAS E MILHAS', font_name='Georgia', font_size=11, bold=True, font_color=COPPER2)
    ws.set_row(1, 30)
    s.t('A2', 'Milhas do zero — ' + ('Exemplos preenchidos' if book.ex else 'Planilha do aluno'),
        font_name='Georgia', font_size=20, bold=True, font_color=COPPER)
    s.t('A3', 'Material complementar do e-book. Use as abas junto com os capítulos indicados.', italic=True, font_size=11)
    if book.ex:
        s.t('A4', 'VERSÃO COM EXEMPLOS FICTÍCIOS: nomes, valores e datas são inventados (ou vêm dos exemplos do livro). Nada aqui é caso real nem referência de preço.',
            bold=True, font_color='#FFFFFF', bg_color=COPPER)
        for c in 'BC': ws.write_blank(c + '4', None, book.f(bg_color=COPPER))
    else:
        s.t('A4', 'VERSÃO DO ALUNO: preencha com os seus dados. Se quiser ver as abas preenchidas com exemplos, abra o arquivo "Exemplos preenchidos".',
            bold=True, font_color='#FFFFFF', bg_color=COPPER)
        for c in 'BC': ws.write_blank(c + '4', None, book.f(bg_color=COPPER))
    s.section('A6', 'Como usar, em 5 passos'); ws.write_blank('B6', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2)); ws.write_blank('C6', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))
    passos = ['1. Comece pela aba 02 Diagnóstico e siga a ordem das abas, acompanhando os capítulos do e-book (coluna "Capítulo(s)" abaixo).',
              '2. Preencha apenas as células claras, com borda. Elas aceitam texto, números e datas, e algumas têm uma lista de opções.',
              '3. As células rosadas são calculadas: não digite nelas. Elas ficam em branco ou mostram um aviso quando falta algum dado.',
              '4. Em cada regra de programa, registre a fonte e a data da consulta. A planilha não adivinha prazos nem validades.',
              '5. Uma vez por mês, volte a esta aba, confira o resumo e faça o checklist mensal (aba 13 Checklists).']
    for i, p in enumerate(passos): s.t('A%d' % (7 + i), p)
    s.section('A13', 'Legenda de cores'); ws.write_blank('B13', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2)); ws.write_blank('C13', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))
    ws.write('A14', 'Campo para preencher', book.f(bg_color=INPUT_BG, border=1, border_color=BORDER, locked=True)); s.t('C14', 'Células claras, com borda: você pode digitar.')
    ws.write('A15', 'Campo calculado (não edite)', book.calc_fmt('text')); s.t('C15', 'Células rosadas: fórmulas protegidas. Mostram resultados ou avisos.')
    ws.write('A16', 'Cabeçalho de tabela', book.f(bold=True, font_color='#FFFFFF', bg_color=COPPER, border=1, border_color='#FFFFFF')); s.t('C16', 'Cabeçalho em cobre escuro; colunas calculadas têm cabeçalho em cobre mais claro.')
    alertas = [('Vencido', RED_L, 'Lote vencido (aba 04).'), ('Vence em até 30 dias', ORANGE_L, 'Intervalos de 30, 60 e 90 dias são sugestões de organização, não prazos oficiais.'),
               ('Vence em até 60 dias', YELLOW_L, ''), ('Vence em até 90 dias', YELLOW_LL, ''),
               ('Data não informada', GRAY_L, 'O lote não tem data de vencimento preenchida.'), ('Sem expiração confirmada', GREEN_L, 'Só use quando o programa informar, na fonte oficial, que esse lote não expira.')]
    for i, (tx, col, ds) in enumerate(alertas):
        ws.write('A%d' % (17 + i), tx, book.f(bg_color=col, border=1, border_color=BORDER, bold=True)); 
        if ds: s.t('C%d' % (17 + i), ds)
    s.section('A24', 'Avisos importantes'); ws.write_blank('B24', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2)); ws.write_blank('C24', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))
    avisos = ['• Pontos de programas diferentes NÃO são equivalentes. Esta planilha não soma pontos ou milhas de programas diferentes como se fossem a mesma moeda: cada programa tem as suas regras (Capítulos 5 e 11).',
              '• Regras, prazos e validades mudam e variam por programa e pela origem dos pontos. Confira sempre no site oficial e registre a fonte e a data. Nada aqui infere validade ou prazo pelo nome do programa.',
              '• A proteção das células calculadas evita alterações acidentais. Ela NÃO protege dados confidenciais e não tem senha. Não registre senhas, códigos de verificação, CPF nem número completo de cartão.',
              '• Pontos e milhas não são dinheiro nem investimento. Esta planilha é uma ferramenta de organização e de cálculo: não é recomendação de compra, de crédito ou de investimento.',
              '• Os exemplos do livro são fictícios. Substitua pelos seus números, na data da sua busca.']
    for i, a in enumerate(avisos): s.t('A%d' % (25 + i), a)
    s.section('A31', 'Para onde ir (clique no nome da aba)'); ws.write_blank('B31', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2)); ws.write_blank('C31', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))
    s.head(32, ['Aba', 'Capítulo(s) do e-book', 'O que fazer nela'], height=22)
    nav = [(2, 'Capítulo 1', 'Responder ao diagnóstico, listar programas, saldos e validades, e montar o resumo do ponto de partida.'),
           (3, 'Capítulos 6 e 9', 'Organizar suas contas: saldo, último acesso, último acúmulo ou resgate e canal oficial de atendimento.'),
           (4, 'Capítulos 10 e 11', 'Registrar uma linha por lote, com origem, quantidade e vencimento, e acompanhar os alertas.'),
           (5, 'Capítulo 7', 'Mapear onde você já gasta e como pontuar, comparando o preço final com a alternativa.'),
           (6, 'Capítulo 7', 'Controlar compras feitas pelo caminho de um programa, com datas e prazos que você confirmou.'),
           (7, 'Capítulos 4, 12 e 13', 'Comparar dinheiro e pontos ou milhas e calcular o valor obtido por milheiro.'),
           (8, 'Capítulos 13 a 15', 'Registrar condições de alteração, cancelamento e bagagem e calcular a perda nas suas condições.'),
           (9, 'Capítulos 16 e 17', 'Calcular o custo por milheiro de clubes e de compra de pontos, no período completo.'),
           (10, 'Capítulo 17', 'Calcular milhas recebidas e custo por milheiro em transferências bonificadas, respeitando o limite.'),
           (11, 'Capítulo 10', 'Comparar desembolsos antes de pagar por extensão ou reativação.'),
           (12, 'Capítulo 18', 'Acompanhar as 21 atividades do plano de 30 dias.'),
           (13, 'Capítulos 11, 14, 15, 17 e 19', 'Marcar os 49 itens dos checklists e registrar a data de revisão.'),
           (14, 'Capítulo 10 e fontes', 'Consultar as tabelas explicativas, os exemplos datados da Smiles e a relação de fontes.')]
    for i, (n, cap, desc) in enumerate(nav):
        r = 33 + i
        ws.write_url('A%d' % r, "internal:'%s'!A1" % NAMES[n], book.f(font_color=COPPER, underline=1, bold=True, border=1, border_color=BORDER),
                     string=NAMES[n], tip='Ir para ' + NAMES[n])
        ws.write('B%d' % r, cap, book.f(border=1, border_color=BORDER, text_wrap=True)); ws.write('C%d' % r, desc, book.f(border=1, border_color=BORDER, text_wrap=True))
    s.section('A47', 'Configurações'); ws.write_blank('B47', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2)); ws.write_blank('C47', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))
    s.label('A48', 'Data de referência (opcional)', bold=True)
    s.inp('B48', 'date', ex=REF_DATE if book.ex else None)
    s.t('C48', 'Deixe em branco para usar a data de hoje. Na versão de exemplos, a data fixa de 16/11/2026 mantém os resultados reproduzíveis.', text_wrap=True)
    s.label('A49', 'Data usada nos cálculos', bold=True)
    s.calc('B49', '=IF(ISNUMBER($B$48),$B$48,TODAY())', 'date')
    s.t('C49', 'Todas as abas usam esta data para dias até o vencimento, alertas e prazos.')
    s.validate_date('B48')
    book.wb.define_name('DataRef', "='%s'!$B$49" % NAMES[1])
    s.section('A51', 'Resumo simples (contagens; nunca soma pontos de programas diferentes)'); ws.write_blank('B51', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2)); ws.write_blank('C51', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))
    L = (LOTES[0], LOTES[1]); I = rng(4, 'I', *L)
    items = [('Lotes vencidos', '=COUNTIF(%s,"Vencido")' % I, 'int', 'Aba 04. Verifique o que ainda pode ser feito, conforme as regras do programa.'),
             ('Lotes que vencem hoje', '=COUNTIF(%s,"Vence hoje")' % I, 'int', 'Aba 04.'),
             ('Lotes que vencem em até 30 dias', '=COUNTIF(%s,"Vence em até 30 dias")' % I, 'int', 'Aba 04. Sugestão de organização.'),
             ('Lotes que vencem em até 60 dias', '=COUNTIF(%s,"Vence em até 60 dias")' % I, 'int', 'Aba 04.'),
             ('Lotes que vencem em até 90 dias', '=COUNTIF(%s,"Vence em até 90 dias")' % I, 'int', 'Aba 04.'),
             ('Lotes sem data informada', '=COUNTIF(%s,"Data não informada")' % I, 'int', 'Aba 04. Procure a data no site oficial do programa.'),
             ('Compras aguardando crédito', '=COUNTIF(%s,"Aguardando")' % rng(6, 'U', *COMPRAS), 'int', 'Aba 06.'),
             ('Compras divergentes ou em contestação', '=COUNTIF(%s,"Divergente")+COUNTIF(%s,"Em contestação")' % (rng(6, 'U', *COMPRAS), rng(6, 'U', *COMPRAS)), 'int', 'Aba 06.'),
             ('Compras com alerta de prazo ou divergência', '=COUNTIF(%s,"*passou*")+COUNTIF(%s,"Divergência*")+COUNTIF(%s,"*em até 7 dias*")' % ((rng(6, 'V', *COMPRAS),) * 3), 'int', 'Aba 06.'),
             ('Atividades do plano de 30 dias concluídas', '=COUNTIF(%s,"Concluído")&" de "&(COUNTA(%s)-COUNTIF(%s,"Não se aplica"))' % (rng(12, 'G', 8, 28), rng(12, 'C', 8, 28), rng(12, 'G', 8, 28)), 'text', 'Aba 12. Não conta atividades marcadas como "Não se aplica".'),
             ('Itens de checklist concluídos', '=COUNTIF(%s,"Concluído")&" de "&COUNTA(%s)' % (rng(13, 'D', 8, 56), rng(13, 'B', 8, 56)), 'text', 'Aba 13 (49 itens). Somados às 21 atividades do plano, formam os 70 itens do livro.')]
    for i, (lb, fm, kd, ds) in enumerate(items):
        r = 52 + i
        s.label('A%d' % r, lb, bold=False, wrap=False)
        s.calc('B%d' % r, fm, kd)
        s.t('C%d' % r, ds, italic=True, font_color=GRAY)
    s.finish(landscape=True, area='A1:C%d' % (52 + len(items)))
    ws.activate()
    return s

# ===================================================================== 02
def sheet02(book):
    s = book.sheet(2); ws = s.ws
    ws.set_column('A:A', 56); ws.set_column('B:B', 22); ws.set_column('C:C', 18); ws.set_column('D:D', 18); ws.set_column('E:E', 34); ws.set_column('F:F', 3); ws.set_column('G:G', 30)
    s.title('Diagnóstico inicial', 'Capítulo 1 — Por onde você começa?',
            'Responda com os seus dados. As opções com lista aceitam só as respostas previstas; deixe em branco o que não souber.')
    s.back_link('A4')
    s.section('A6', '1. Qual é o seu objetivo de viagem? (marque todas as opções que se aplicam)'); ws.write_blank('B6', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))
    obj = ['Uma viagem nacional por ano', 'Uma viagem internacional nos próximos 2 ou 3 anos', 'Ainda não sei, só quero entender como funciona']
    for i, o in enumerate(obj):
        r = 7 + i
        s.label('A%d' % r, o, wrap=False); s.inp('B%d' % r, 'text', ex='Sim' if (book.ex and i == 0) else None)
    s.validate_list('B7:B9', items=['Sim'], msg='Escolha "Sim" para marcar. Pode marcar mais de uma opção.')
    s.section('A11', '2. Detalhes da viagem que você imagina'); ws.write_blank('B11', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))
    s.label('A12', 'Cidade de saída', wrap=False); s.inp('B12', 'text', ex='Campinas'); 
    s.label('A13', 'Destino desejado', wrap=False); s.inp('B13', 'text', ex='Salvador')
    s.label('A14', 'Quantas pessoas vão viajar', wrap=False); s.inp('B14', 'int', ex=2)
    s.label('A15', 'Prazo (daqui a quantos meses)', wrap=False); s.inp('B15', 'int', ex=10)
    s.label('A16', 'Suas datas são flexíveis?', wrap=False); s.inp('B16', 'text', ex='Sim')
    s.validate_num('B14', 1, integer=True, msg='Número de viajantes (1 ou mais).'); s.validate_num('B15', 0, integer=True)
    s.validate_list('B16', items=['Sim', 'Mais ou menos', 'Não'])
    s.section('A18', '3. Com que frequência você viaja de avião?'); ws.write_blank('B18', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))
    s.label('A19', 'Frequência', wrap=False); s.inp('B19', 'text')
    s.validate_list('B19', items=['Nunca ou quase nunca', 'Uma ou duas vezes por ano', 'Três vezes por ano ou mais'])
    ws.set_column('B:B', 34)
    s.section('A21', '4. Quais são seus gastos habituais? (anote os 3 maiores do mês)'); ws.write_blank('B21', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2)); ws.write_blank('C21', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))
    s.head(22, ['Gasto (mercado, combustível, contas, compras online…)', 'Valor médio por mês (R$, opcional)'], height=32)
    gastos = [('Mercado', None), ('Contas', None), (None, None)]
    for i in range(3):
        r = 23 + i
        s.inp('A%d' % r, 'text', ex=gastos[i][0]); s.inp('B%d' % r, 'money')
    s.validate_num('B23:B25', 0)
    s.section('A27', '5. Programas em que você já tem cadastro, com saldo e validade'); 
    for c in 'BCDE': ws.write_blank(c + '27', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))
    s.t('A28', 'Liste companhias aéreas, bancos, lojas, farmácias e supermercados. Informe a validade só se houver um único prazo; se o saldo tiver vários lotes, marque "Sim" e detalhe na aba 04.')
    s.head(29, ['Programa', 'Saldo (pontos ou milhas)', 'Validade (data)', 'Saldo com vários lotes?', 'Aviso'], height=34, calc_cols=(4,))
    r0, r1 = 30, 44
    progs = [('Programa da farmácia (exemplo)', None, None, None), ('Programa do supermercado (exemplo)', None, None, None)] if book.ex else []
    for i, r in enumerate(range(r0, r1 + 1)):
        ex = progs[i] if i < len(progs) else (None, None, None, None)
        s.inp('A%d' % r, 'text', ex=ex[0]); s.inp('B%d' % r, 'pts', ex=ex[1]); s.inp('C%d' % r, 'date', ex=ex[2]); s.inp('D%d' % r, 'text', ex=ex[3])
        s.calc('E%d' % r, '=IF($A{r}="","",IF($D{r}="Sim","Detalhe os lotes na aba 04 Lotes (não precisa de uma única validade).",IF(AND(ISNUMBER($B{r}),$C{r}=""),"Falta a validade (ou marque que há vários lotes).",IF(AND($C{r}<>"",NOT(ISNUMBER($C{r}))),"Validade inválida: use uma data.",""))))'.format(r=r), 'text')
    s.validate_list('D%d:D%d' % (r0, r1), items=['Sim', 'Não'])
    s.validate_num('B%d:B%d' % (r0, r1), 0); s.validate_date('C%d:C%d' % (r0, r1))
    ws.write_url('G29', "internal:'%s'!A1" % NAMES[4], book.f(font_color=COPPER, underline=1), string='→ Ir para a aba 04 Lotes')
    ws.write_url('G30', "internal:'%s'!A1" % NAMES[3], book.f(font_color=COPPER, underline=1), string='→ Ir para a aba 03 Contas')
    s.section('A46', '6. Você tem hoje algum cartão de crédito?'); ws.write_blank('B46', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))
    s.label('A47', 'Resposta', wrap=False); s.inp('B47', 'text'); s.validate_list('B47', items=['Sim', 'Não', 'Prefiro não usar'])
    s.section('A49', 'Seu resumo de ponto de partida'); 
    for c in 'BCDE': ws.write_blank(c + '49', None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))
    s.label('A50', 'Origem', wrap=False); s.calc('B50', '=IF($B$12="","",$B$12)')
    s.label('A51', 'Destino', wrap=False); s.calc('B51', '=IF($B$13="","",$B$13)')
    s.label('A52', 'Viajantes', wrap=False); s.calc('B52', '=IF(ISNUMBER($B$14),$B$14,"")', 'int')
    s.label('A53', 'Prazo (meses)', wrap=False); s.calc('B53', '=IF(ISNUMBER($B$15),$B$15,"")', 'int')
    s.label('A54', 'Programas em que já tenho cadastro', wrap=False)
    # auxiliar de concatenação (coluna G, linhas 60 a 74): lista de programas sem TEXTJOIN
    for k, r in enumerate(range(r0, r1 + 1)):
        gr = r + 30
        if k == 0:
            s.calc('G%d' % gr, '=IF($A%d="","",$A%d)' % (r, r), 'text')
        else:
            s.calc('G%d' % gr, '=IF($A{r}="",G{pr},IF(G{pr}="",$A{r},G{pr}&", "&$A{r}))'.format(r=r, pr=gr - 1), 'text')
    s.calc('B54', '=IF(G%d="","",G%d)' % (r1 + 30, r1 + 30), 'long')
    s.label('A55', 'Meu primeiro passo', wrap=False); s.inp('B55', 'long', ex='Ver o saldo e a validade desses dois programas' if book.ex else None)
    ws.set_row(53, 32); ws.set_row(54, 32)
    s.label('A56', 'Resumo montado', bold=True, wrap=False)
    s.calc('B56', '="Quero viajar de "&IF($B$12="","__________",$B$12)&" para "&IF($B$13="","__________",$B$13)&", com "&IF(ISNUMBER($B$14),$B$14,"____")&" pessoa(s), daqui a "&IF(ISNUMBER($B$15),$B$15,"____")&" meses. Já tenho cadastro em "&IF($B$54="","__________",$B$54)&". Meu primeiro passo é "&IF($B$55="","__________",$B$55)&"."', 'long')
    ws.set_row(55, 62)
    s.t('G59', 'Coluna G (linhas 60 a 74): auxiliar da lista de programas. Não edite.', italic=True, font_color=GRAY)
    s.t('A58', 'Guarde esse resumo: nos próximos capítulos você volta a ele para escolher seus primeiros passos.', italic=True)
    if book.ex:
        s.t('A59', 'Exemplo do livro (Capítulo 1): a Marina quer viajar de Campinas a Salvador, com 2 pessoas, daqui a 10 meses. Os dois programas listados são fictícios.', italic=True, font_color=GRAY, text_wrap=True)
    s.finish(area='A1:E59', header_row=None)
    return s

# ===================================================================== 03
def sheet03(book):
    s = book.sheet(3); ws = s.ws
    widths = {'A': 30, 'B': 24, 'C': 18, 'D': 16, 'E': 16, 'F': 24, 'G': 18, 'H': 15, 'I': 14, 'J': 16, 'K': 16, 'L': 15, 'M': 34, 'N': 34, 'O': 34, 'P': 14, 'Q': 34}
    for c, w in widths.items(): ws.set_column('%s:%s' % (c, c), w)
    s.title('Organização das contas', 'Capítulos 6 e 9 — Seu painel de controle, uma linha por programa',
            'Não registre senhas, códigos de verificação, CPF nem número completo de cartão. Último acesso NÃO é o mesmo que último acúmulo ou resgate: anote os dois separadamente.')
    s.back_link('A4')
    s.t('A5', 'Saldo: quando há lotes na aba 04 para o programa, o saldo vem da SOMA DOS LOTES (calculado). O saldo manual só vale se ainda não houver lotes; assim não existem duas entradas conflitantes.', italic=True, font_color=GRAY)
    hdr = ['Programa', 'Login (e-mail ou identificador, opcional; nunca senha)', 'Saldo manual (use só se ainda não detalhou os lotes)', 'Saldo dos lotes (aba 04)', 'Saldo considerado',
           'Origem do saldo', 'Pontos a vencer em até 90 dias (aba 04)', 'Próximo vencimento (aba 04)', 'Último acesso', 'Último acúmulo ou resgate',
           'Prazo de inatividade informado pelo programa (meses)', 'Meses desde o último acúmulo ou resgate', 'Alerta de inatividade',
           'Canal oficial de atendimento', 'Link do regulamento', 'Data da consulta', 'Observações']
    s.head(6, hdr, height=62, calc_cols=(3, 4, 5, 6, 7, 11, 12))
    r0, r1 = CONTAS
    LA = rng(4, 'A', *LOTES); LD = rng(4, 'D', *LOTES); LF = rng(4, 'F', *LOTES); LG = rng(4, 'G', *LOTES)
    ex = {
        7: ('Programa de exemplo A', 'meuemail@exemplo.com', None, None, D(2026, 11, 1), D(2026, 11, 1), 36, 'Atendimento oficial do programa (consulte no site)', 'Regulamento no site oficial do programa', D(2026, 11, 12), '[Exemplo do livro, Cap. 10] Saldo vem dos dois lotes da Marina na aba 04 (8.000 + 12.000). Último acesso e prazo de inatividade são fictícios.'),
        8: ('Programa de exemplo B', None, None, None, D(2026, 11, 2), D(2024, 10, 1), 24, None, None, None, '[Exemplo adicional fictício] Mostra o alerta de inatividade: acessar a conta não é o mesmo que acumular ou resgatar. Prazo inventado.'),
        9: ('Programa de exemplo C', None, None, None, D(2026, 10, 1), D(2023, 12, 1), 36, None, None, None, '[Exemplo adicional fictício]'),
        10: ('Programa de exemplo D', None, None, None, None, D(2026, 10, 15), None, None, None, None, '[Exemplo adicional fictício] Sem prazo de inatividade informado: nenhum alerta é calculado.'),
        11: ('Programa da farmácia (exemplo)', None, 900, None, D(2026, 10, 20), D(2026, 9, 12), None, None, None, None, '[Exemplo do livro, Cap. 6] Saldo pequeno que vencia em dois meses; aqui o saldo foi informado manualmente (sem lotes na aba 04). Valores fictícios.'),
        12: ('Programa do supermercado (exemplo)', None, 2500, None, D(2026, 11, 2), D(2026, 11, 2), None, None, None, None, '[Exemplo adicional fictício]'),
    }
    for r in range(r0, r1 + 1):
        e = ex.get(r) if book.ex else None
        e = e or (None,) * 11
        s.inp('A%d' % r, 'text', ex=e[0]); s.inp('B%d' % r, 'text', ex=e[1]); s.inp('C%d' % r, 'pts', ex=e[2])
        s.calc('D%d' % r, '=IF($A{r}="","",SUMIFS({D},{A},$A{r}))'.format(r=r, D=LD, A=LA), 'pts')
        s.calc('E%d' % r, '=IF($A{r}="","",IF(COUNTIFS({A},$A{r})>0,$D{r},IF(ISNUMBER($C{r}),$C{r},"")))'.format(r=r, A=LA), 'pts', bold=True)
        s.calc('F%d' % r, '=IF($A{r}="","",IF(COUNTIFS({A},$A{r})>0,IF(AND(ISNUMBER($C{r}),$C{r}<>$D{r}),"Soma dos lotes (o saldo manual foi ignorado)","Soma dos lotes (aba 04)"),IF(ISNUMBER($C{r}),"Informado manualmente","Sem saldo informado")))'.format(r=r, A=LA), 'text')
        s.calc('G%d' % r, '=IF($A{r}="","",SUMIFS({D},{A},$A{r},{F},">="&DataRef,{F},"<="&(DataRef+90),{G},"<>Sim, confirmado"))'.format(r=r, D=LD, A=LA, F=LF, G=LG), 'pts')
        s.calc('H%d' % r, '=IF($A{r}="","",IFERROR(AGGREGATE(15,6,{F}/(({A}=$A{r})*({F}>=DataRef)*({G}<>"Sim, confirmado")),1),""))'.format(r=r, A=LA, F=LF, G=LG), 'date')
        s.inp('I%d' % r, 'date', ex=e[4]); s.inp('J%d' % r, 'date', ex=e[5]); s.inp('K%d' % r, 'mes', ex=e[6])
        s.calc('L%d' % r, '=IF(OR($A{r}="",NOT(ISNUMBER($J{r}))),"",IF($J{r}>DataRef,"",(YEAR(DataRef)-YEAR($J{r}))*12+MONTH(DataRef)-MONTH($J{r})-IF(DAY(DataRef)<DAY($J{r}),1,0)))'.format(r=r), 'mes')
        s.calc('M%d' % r, '=IF(OR($A{r}="",NOT(ISNUMBER($K{r}))),"",IF(NOT(ISNUMBER($L{r})),"Informe a data do último acúmulo ou resgate.",IF($L{r}>=$K{r},"Prazo informado atingido ou ultrapassado: confira com o programa.",IF($L{r}>=$K{r}-3,"Atenção: a menos de 3 meses do prazo informado (sugestão de organização).","Dentro do prazo informado."))))'.format(r=r), 'text')
        s.inp('N%d' % r, 'text', ex=e[7]); s.inp('O%d' % r, 'text', ex=e[8]); s.inp('P%d' % r, 'date', ex=e[9]); s.inp('Q%d' % r, 'text', ex=e[10])
    # validações
    s.validate_num('C%d:C%d' % (r0, r1), 0); s.validate_num('K%d:K%d' % (r0, r1), 0, 600, integer=True, msg='Prazo de inatividade que o PROGRAMA informa, em meses. Em branco se não souber.')
    s.validate_date('I%d:J%d' % (r0, r1)); s.validate_date('P%d:P%d' % (r0, r1))
    ws.data_validation('A%d:A%d' % (r0, r1), {'validate': 'length', 'criteria': '<=', 'value': 80, 'ignore_blank': True, 'error_title': 'Texto longo', 'error_message': 'Use até 80 caracteres no nome do programa.'})
    ws.data_validation('B%d:B%d' % (r0, r1), {'validate': 'custom', 'value': '=ISERROR(SEARCH("senha",B%d))' % r0, 'error_type': 'stop', 'error_title': 'Não registre senhas',
                                              'error_message': 'Esta coluna é só para e-mail ou identificador. Nunca escreva senhas nem códigos aqui.', 'ignore_blank': True})
    s.cf('M%d:M%d' % (r0, r1), '=LEFT($M%d,5)="Prazo"' % r0, {'bg_color': RED_L})
    s.cf('M%d:M%d' % (r0, r1), '=LEFT($M%d,7)="Atenção"' % r0, {'bg_color': ORANGE_L})
    s.cf('F%d:F%d' % (r0, r1), '=ISNUMBER(SEARCH("ignorado",$F%d))' % r0, {'bg_color': YELLOW_L})
    ws.set_default_row(30)
    s.finish(header_row=6, freeze=(6, 1), autofilter='A6:Q%d' % r1, area='A1:Q%d' % r1)
    return s

# ===================================================================== 04
def sheet04(book):
    s = book.sheet(4); ws = s.ws
    widths = {'A': 30, 'B': 12, 'C': 32, 'D': 14, 'E': 13, 'F': 14, 'G': 16, 'H': 12, 'I': 30, 'J': 30, 'K': 22, 'L': 36, 'M': 14, 'N': 34}
    for c, w in widths.items(): ws.set_column('%s:%s' % (c, c), w)
    s.title('Lotes e validade', 'Capítulos 10 e 11 — Uma linha por lote (programa, quantidade, origem e vencimento)',
            'Não infira a validade pelo nome do programa: confira a regra no site oficial, para cada origem, e registre a fonte e a data da consulta. Os alertas de 30, 60 e 90 dias são sugestões de organização.')
    # resumo (contagens)
    labels = ['Vencidos', 'Vencem hoje', 'Até 30 dias', 'Até 60 dias', 'Até 90 dias', 'Sem data informada', 'Sem expiração confirmada']
    keys = ['Vencido', 'Vence hoje', 'Vence em até 30 dias', 'Vence em até 60 dias', 'Vence em até 90 dias', 'Data não informada', 'Sem expiração confirmada']
    cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
    I = rng(4, 'I', *LOTES)
    for c, lb, k in zip(cols, labels, keys):
        s.t('%s4' % c, lb, bold=True, font_size=9, align='center', text_wrap=True, bg_color=ROSE_L, border=1, border_color=BORDER)
    ws.set_row(3, 28)
    # contagens ficam na linha 6? precisamos da linha 6 para o cabeçalho; usamos a linha 5 para rótulos e a 4 para valores
    for c, k in zip(cols, keys):
        s.calc('%s5' % c, '=COUNTIF(%s,"%s")' % (I, k), 'int', bold=True)
    ws.write_url('I5', "internal:'%s'!A1" % NAMES[1], book.f(font_color=COPPER, underline=1), string='← Voltar à aba 01 Comece aqui')
    hdr = ['Programa', 'Identificador do lote', 'Origem dos pontos', 'Quantidade disponível', 'Data de crédito (se conhecida)', 'Data de vencimento', 'Sem expiração confirmada?',
           'Dias até o vencimento', 'Situação do lote', 'Plano de uso', 'Ação escolhida', 'Fonte da regra (endereço ou documento)', 'Data da consulta', 'Observações']
    s.head(6, hdr, height=50, calc_cols=(7, 8))
    r0, r1 = LOTES
    ex = {7: ('Programa de exemplo A', 'A1', 'Origem informada pelo programa (exemplo)', 8000, D(2026, 9, 1), D(2027, 3, 16), None, 'Usar em uma compra planejada antes do vencimento', 'Usar', 'Fonte e data a registrar', D(2026, 11, 12), '[Exemplo do livro, Cap. 10] Lote da Marina: 8.000 pontos vencendo em 4 meses. Os intervalos de lembrete (90, 60 e 30 dias) são sugestão.'),
          8: ('Programa de exemplo A', 'A2', 'Origem informada pelo programa (exemplo)', 12000, D(2026, 11, 1), D(2028, 5, 16), None, 'Ainda sem uso previsto', None, 'Fonte e data a registrar', D(2026, 11, 12), '[Exemplo do livro, Cap. 10] Lote da Marina: 12.000 pontos vencendo em 18 meses.'),
          9: ('Programa de exemplo B', 'B1', 'Origem fictícia (vencido)', 3000, D(2024, 10, 1), D(2026, 11, 1), None, 'Verificar se há extensão ou reativação', 'Avaliar reativação', None, None, '[Exemplo adicional fictício] Lote vencido.'),
          10: ('Programa de exemplo B', 'B2', 'Origem fictícia (vence hoje)', 1000, D(2024, 11, 16), D(2026, 11, 16), None, None, 'Usar', None, None, '[Exemplo adicional fictício] Vence na data de referência.'),
          11: ('Programa de exemplo B', 'B3', 'Origem fictícia (até 30 dias)', 1500, None, D(2026, 12, 10), None, None, 'Usar', None, None, '[Exemplo adicional fictício]'),
          12: ('Programa de exemplo C', 'C1', 'Origem fictícia (até 60 dias)', 2000, None, D(2027, 1, 10), None, None, 'Avaliar extensão', None, None, '[Exemplo adicional fictício]'),
          13: ('Programa de exemplo C', 'C2', 'Origem fictícia (até 90 dias)', 2500, None, D(2027, 2, 10), None, None, None, None, None, '[Exemplo adicional fictício]'),
          14: ('Programa de exemplo C', 'C3', 'Origem fictícia (sem data informada)', 4000, None, None, None, 'Descobrir a validade no site oficial', None, None, None, '[Exemplo adicional fictício] Data não informada: não é o mesmo que sem expiração.'),
          15: ('Programa de exemplo D', 'D1', 'Origem fictícia (sem expiração confirmada)', 6000, None, None, 'Sim, confirmado', None, None, 'Regulamento do programa (fictício)', D(2026, 11, 12), '[Exemplo adicional fictício] Só marque "Sim, confirmado" com a fonte oficial.')}
    for r in range(r0, r1 + 1):
        e = ex.get(r) if book.ex else None
        e = e or (None,) * 12
        s.inp('A%d' % r, 'text', ex=e[0]); s.inp('B%d' % r, 'text', ex=e[1]); s.inp('C%d' % r, 'text', ex=e[2]); s.inp('D%d' % r, 'pts', ex=e[3])
        s.inp('E%d' % r, 'date', ex=e[4]); s.inp('F%d' % r, 'date', ex=e[5]); s.inp('G%d' % r, 'text', ex=e[6])
        s.calc('H%d' % r, '=IF(OR($A{r}="",$G{r}="Sim, confirmado",NOT(ISNUMBER($F{r}))),"",INT($F{r})-INT(DataRef))'.format(r=r), 'int')
        s.calc('I%d' % r, ('=IF(AND($A{r}="",$D{r}="",$F{r}="",$G{r}=""),"",IF($G{r}="Sim, confirmado",IF($F{r}<>"","Conferir: sem expiração e data preenchidas","Sem expiração confirmada"),'
                           'IF($F{r}="","Data não informada",IF(NOT(ISNUMBER($F{r})),"Data inválida",IF($H{r}<0,"Vencido",IF($H{r}=0,"Vence hoje",IF($H{r}<=30,"Vence em até 30 dias",IF($H{r}<=60,"Vence em até 60 dias",IF($H{r}<=90,"Vence em até 90 dias","Mais de 90 dias")))))))))').format(r=r), 'text', bold=True)
        s.inp('J%d' % r, 'text', ex=e[7]); s.inp('K%d' % r, 'text', ex=e[8]); s.inp('L%d' % r, 'text', ex=e[9]); s.inp('M%d' % r, 'date', ex=e[10]); s.inp('N%d' % r, 'text', ex=e[11])
    s.validate_list('A%d:A%d' % (r0, r1), source="=%s" % rng(3, 'A', *CONTAS), strict=False, msg='Escolha um programa da aba 03 Contas ou digite um novo (cadastre-o também na aba 03).')
    s.validate_num('D%d:D%d' % (r0, r1), 0); s.validate_date('E%d:F%d' % (r0, r1)); s.validate_date('M%d:M%d' % (r0, r1))
    s.validate_list('G%d:G%d' % (r0, r1), items=['Sim, confirmado'], msg='Marque só se a fonte oficial confirmar que este lote não expira. Em branco NÃO significa que não expira.')
    s.validate_list('K%d:K%d' % (r0, r1), items=['Usar', 'Avaliar extensão', 'Avaliar reativação', 'Deixar vencer'])
    # alertas visuais
    s.cf('A%d:A%d' % (r0, r1) + ',H%d:I%d' % (r0, r1), '=$I%d="Vencido"' % r0, {'bg_color': RED_L, 'bold': True})
    s.cf('A%d:A%d' % (r0, r1) + ',H%d:I%d' % (r0, r1), '=OR($I%d="Vence hoje",$I%d="Vence em até 30 dias")' % (r0, r0), {'bg_color': ORANGE_L})
    s.cf('A%d:A%d' % (r0, r1) + ',H%d:I%d' % (r0, r1), '=$I%d="Vence em até 60 dias"' % r0, {'bg_color': YELLOW_L})
    s.cf('A%d:A%d' % (r0, r1) + ',H%d:I%d' % (r0, r1), '=$I%d="Vence em até 90 dias"' % r0, {'bg_color': YELLOW_LL})
    s.cf('A%d:A%d' % (r0, r1) + ',H%d:I%d' % (r0, r1), '=OR($I%d="Data não informada",$I%d="Data inválida",LEFT($I%d,8)="Conferir")' % (r0, r0, r0), {'bg_color': GRAY_L})
    s.cf('A%d:A%d' % (r0, r1) + ',H%d:I%d' % (r0, r1), '=$I%d="Sem expiração confirmada"' % r0, {'bg_color': GREEN_L})
    ws.set_default_row(30)
    s.finish(header_row=6, freeze=(6, 1), autofilter='A6:N%d' % r1, area='A1:N56')
    return s

# ===================================================================== 05
def sheet05(book):
    s = book.sheet(5); ws = s.ws
    widths = {'A': 28, 'B': 26, 'C': 16, 'D': 38, 'E': 16, 'F': 24, 'G': 22, 'H': 18, 'I': 18, 'J': 18, 'K': 16, 'L': 46, 'M': 30, 'N': 34}
    for c, w in widths.items(): ws.set_column('%s:%s' % (c, c), w)
    s.title('Mapa de acúmulo', 'Capítulo 7 — Como acumular com os gastos que você já tem',
            'Use só gastos que você já teria. Antes de comprar pelo caminho de um programa, compare o preço final (com frete e descontos) e confira o meio de pagamento aceito pela campanha.')
    s.back_link('A4')
    s.t('A5', 'Se o caminho do programa custa mais, os pontos não são "gratuitos": há pagamento adicional para obtê-los. Não crie gastos novos só para pontuar.', italic=True, font_color=GRAY)
    hdr = ['O que eu já gasto', 'Onde gasto', 'Esse lugar tem programa de pontos?', 'Qual o passo para pontuar?', 'Valor habitual do gasto (R$ por mês)', 'Programa de destino',
           'Meio de pagamento aceito pela campanha', 'Exige ativação ou cadastro na campanha?', 'Preço final pelo caminho do programa (R$, com frete e descontos)',
           'Melhor preço final alternativo (R$, com frete e descontos)', 'Diferença de preço (R$)', 'Leitura da diferença', 'Fonte e data da consulta', 'Observações']
    s.head(6, hdr, height=66, calc_cols=(10, 11))
    r0, r1 = ACUMULO
    ex = {7: ('Liquidificador (compra planejada)', 'Loja parceira (exemplo)', 'Sim', 'Entrar pelo botão do programa e finalizar a compra na mesma janela', None, 'Programa de exemplo', None, 'Não sei', 220, 200, None, '[Exemplo do livro, Cap. 7] Carlos: R$ 220 com frete pelo link do programa contra R$ 200 com frete em outra loja. Meio de pagamento a conferir na campanha.'),
          8: ('Compras na farmácia', 'Farmácia (exemplo)', 'Sim', 'Conferir as regras, ativar o cadastro necessário e seguir o procedimento exigido para pontuar', None, 'Programa da farmácia (exemplo)', None, 'Sim', None, None, None, '[Exemplo do livro, Cap. 3] Paula. Só informar o CPF no caixa nem sempre basta; confira no site oficial do programa.')}
    for r in range(r0, r1 + 1):
        e = ex.get(r) if book.ex else None
        e = e or (None,) * 12
        s.inp('A%d' % r, 'text', ex=e[0]); s.inp('B%d' % r, 'text', ex=e[1]); s.inp('C%d' % r, 'text', ex=e[2]); s.inp('D%d' % r, 'long', ex=e[3])
        s.inp('E%d' % r, 'money', ex=e[4]); s.inp('F%d' % r, 'text', ex=e[5]); s.inp('G%d' % r, 'text', ex=e[6]); s.inp('H%d' % r, 'text', ex=e[7])
        s.inp('I%d' % r, 'money', ex=e[8]); s.inp('J%d' % r, 'money', ex=e[9])
        s.calc('K%d' % r, '=IF(AND(ISNUMBER($I{r}),ISNUMBER($J{r})),$I{r}-$J{r},"")'.format(r=r), 'money')
        s.calc('L%d' % r, '=IF(NOT(ISNUMBER($K{r})),"",IF($K{r}>0,"Pelo caminho do programa você paga R$ "&FIXED($K{r},2)&" a mais. Os pontos não são gratuitos: o benefício precisa valer mais do que essa diferença.",IF($K{r}<0,"O caminho do programa tem preço final R$ "&FIXED(-$K{r},2)&" menor que a alternativa.","Mesmo preço final nas duas opções.")))'.format(r=r), 'long')
        s.inp('M%d' % r, 'text', ex=None); s.inp('N%d' % r, 'long', ex=e[11])
    s.validate_list('C%d:C%d' % (r0, r1), items=['Sim', 'Não', 'Não sei']); s.validate_list('H%d:H%d' % (r0, r1), items=['Sim', 'Não', 'Não sei'])
    s.validate_num('E%d:E%d' % (r0, r1), 0); s.validate_num('I%d:J%d' % (r0, r1), 0)
    ws.set_default_row(30)
    s.finish(header_row=6, freeze=(6, 1), autofilter='A6:N%d' % r1, area='A1:N%d' % r1)
    return s

# ===================================================================== 06
EVENTOS = ['Data da compra', 'Recebimento do produto', 'Comunicação do acúmulo pelo parceiro', 'Preenchimento dos requisitos de acúmulo', 'Outro (descreva em Observações)', 'Não sei']

def sheet06(book):
    s = book.sheet(6); ws = s.ws
    widths = {'A': 13, 'B': 24, 'C': 15, 'D': 24, 'E': 14, 'F': 14, 'G': 26, 'H': 30, 'I': 13, 'J': 11, 'K': 12, 'L': 15, 'M': 11, 'N': 12, 'O': 24, 'P': 13, 'Q': 15, 'R': 13, 'S': 13, 'T': 15, 'U': 17, 'V': 44, 'W': 18, 'X': 34, 'Y': 3, 'Z': 16}
    for c, w in widths.items(): ws.set_column('%s:%s' % (c, c), w)
    s.title('Controle de compras', 'Capítulo 7 — Compras feitas pelo caminho de um programa de pontos',
            'O prazo NÃO começa necessariamente na compra: escolha o evento que o regulamento ou a campanha indica e confirme o prazo na fonte. Nenhum prazo é preenchido por padrão.')
    s.back_link('A4')
    s.t('A5', 'Dias úteis: a contagem usa segunda a sexta e a lista de feriados da coluna Z (opcional); começa no dia seguinte ao evento. Cada programa pode contar de outro jeito: confirme. A data-limite só é calculada com todos os dados.', italic=True, font_color=GRAY)
    hdr = ['Data da compra', 'Loja ou parceiro', 'Número do pedido (texto)', 'Programa de destino', 'Valor da compra (R$)', 'Pontuação esperada (pontos)', 'Comprovante ou link da oferta',
           'Evento que inicia a contagem do prazo de crédito', 'Data desse evento', 'Prazo informado (número)', 'Dias corridos ou úteis', 'Data prevista de crédito',
           'Prazo para reclamar (número)', 'Dias corridos ou úteis (reclamação)', 'Prazo para reclamar contado a partir de', 'Outra data-base (se escolheu)', 'Data-limite de reclamação',
           'Pontos creditados', 'Data do crédito', 'Diferença: creditado − esperado', 'Situação', 'Alerta', 'Protocolo', 'Observações']
    s.head(6, hdr, height=66, calc_cols=(11, 16, 19, 21))
    s.head(6, ['Feriados (opcional)'], col0=25, height=66)
    s.t('Z5', 'Digite datas de feriados.', font_size=8, italic=True, font_color=GRAY)
    r0, r1 = COMPRAS
    hol = [D(2026, 11, 20)] if book.ex else []
    for i in range(40):
        if i < len(hol): s.inp('Z%d' % (7 + i), 'date', ex=hol[i])
        else: s.inp('Z%d' % (7 + i), 'date')
    s.validate_date('Z7:Z46')
    book.wb.define_name('Feriados', "='%s'!$Z$7:$Z$46" % NAMES[6])
    ex = {7: (D(2026, 11, 3), 'Loja parceira (fictícia)', '000123', 'Programa A (fictício)', 200, 400, 'Print da oferta (fictício)', 'Preenchimento dos requisitos de acúmulo', D(2026, 11, 3), 60, 'Corridos', 30, 'Corridos', 'Data prevista de crédito', None, None, None, 'Aguardando', None, '[Exemplo adicional fictício] Prazos de 60 e 30 dias corridos inventados para demonstrar o cálculo; não são regra de nenhum programa.'),
          8: (D(2026, 11, 10), 'Loja parceira (fictícia)', '0045', 'Programa B (fictício)', 150, 300, None, 'Comunicação do acúmulo pelo parceiro', D(2026, 11, 12), 10, 'Úteis', 30, 'Corridos', 'Data prevista de crédito', None, None, None, 'Aguardando', None, '[Exemplo adicional fictício] Dias úteis com o feriado de 20/11/2026 configurado na coluna Z (exemplo de configuração).'),
          9: (D(2026, 9, 20), 'Loja parceira (fictícia)', '778899', 'Programa A (fictício)', 100, 200, 'Print da oferta (fictício)', 'Data da compra', D(2026, 9, 20), 30, 'Corridos', 30, 'Corridos', 'Data prevista de crédito', None, 180, D(2026, 10, 25), 'Divergente', 'PROT-0001 (fictício)', '[Exemplo adicional fictício] Creditado menos do que o esperado.'),
          10: (D(2026, 10, 1), 'Loja parceira (fictícia)', '123456', 'Programa C (fictício)', 80, 160, None, 'Data da compra', D(2026, 10, 1), 30, 'Corridos', None, None, None, None, None, None, 'Aguardando', None, '[Exemplo adicional fictício] Prazo de crédito já passou sem crédito registrado; sem prazo de reclamação informado, a data-limite não é calculada.')}
    for r in range(r0, r1 + 1):
        e = ex.get(r) if book.ex else None
        e = e or (None,) * 20
        s.inp('A%d' % r, 'date', ex=e[0]); s.inp('B%d' % r, 'text', ex=e[1]); s.inp('C%d' % r, 'textnum', ex=e[2]); s.inp('D%d' % r, 'text', ex=e[3])
        s.inp('E%d' % r, 'money', ex=e[4]); s.inp('F%d' % r, 'pts', ex=e[5]); s.inp('G%d' % r, 'text', ex=e[6]); s.inp('H%d' % r, 'text', ex=e[7])
        s.inp('I%d' % r, 'date', ex=e[8]); s.inp('J%d' % r, 'int', ex=e[9]); s.inp('K%d' % r, 'text', ex=e[10])
        s.calc('L%d' % r, '=IF(OR(NOT(ISNUMBER($I{r})),NOT(ISNUMBER($J{r})),$K{r}=""),"",IF($K{r}="Corridos",$I{r}+$J{r},IF($K{r}="Úteis",WORKDAY($I{r},$J{r},Feriados),"")))'.format(r=r), 'date', bold=True)
        s.inp('M%d' % r, 'int', ex=e[11]); s.inp('N%d' % r, 'text', ex=e[12]); s.inp('O%d' % r, 'text', ex=e[13]); s.inp('P%d' % r, 'date', ex=e[14])
        s.calc('Q%d' % r, ('=IF(OR(NOT(ISNUMBER($M{r})),$N{r}="",$O{r}=""),"",IF($O{r}="Data prevista de crédito",IF(ISNUMBER($L{r}),IF($N{r}="Corridos",$L{r}+$M{r},IF($N{r}="Úteis",WORKDAY($L{r},$M{r},Feriados),"")),""),'
                           'IF($O{r}="Outra data (coluna P)",IF(ISNUMBER($P{r}),IF($N{r}="Corridos",$P{r}+$M{r},IF($N{r}="Úteis",WORKDAY($P{r},$M{r},Feriados),"")),""),"")))').format(r=r), 'date', bold=True)
        s.inp('R%d' % r, 'pts', ex=e[15]); s.inp('S%d' % r, 'date', ex=e[16])
        s.calc('T%d' % r, '=IF(OR(NOT(ISNUMBER($R{r})),NOT(ISNUMBER($F{r}))),"",$R{r}-$F{r})'.format(r=r), 'pts')
        s.inp('U%d' % r, 'text', ex=e[17]); 
        s.calc('V%d' % r, ('=IF(AND($B{r}="",$D{r}="",$A{r}=""),"",IF(AND(ISNUMBER($T{r}),$T{r}<>0),"Divergência: creditado diferente do esperado.",'
                           'IF(AND($U{r}<>"Creditado",ISNUMBER($Q{r}),DataRef>$Q{r}),"Data-limite de reclamação já passou.",'
                           'IF(AND($U{r}<>"Creditado",NOT(ISNUMBER($R{r})),ISNUMBER($L{r}),DataRef>$L{r}),"O prazo previsto de crédito passou sem crédito registrado.",'
                           'IF(AND($U{r}<>"Creditado",ISNUMBER($Q{r}),DataRef>=N($Q{r})-7),"Data-limite de reclamação em até 7 dias.",'
                           'IF(ISNUMBER($L{r}),"Acompanhando.","Preencha o evento, a data, o prazo e o tipo de dias para calcular as datas."))))))').format(r=r), 'text')
        s.inp('W%d' % r, 'text', ex=e[18]); s.inp('X%d' % r, 'long', ex=e[19])
    s.validate_list('H%d:H%d' % (r0, r1), items=EVENTOS, strict=False, msg='Escolha o evento que o regulamento ou a campanha indica. Se for outro, escolha "Outro" e descreva em Observações.')
    s.validate_list('K%d:K%d' % (r0, r1), items=['Corridos', 'Úteis']); s.validate_list('N%d:N%d' % (r0, r1), items=['Corridos', 'Úteis'])
    s.validate_list('O%d:O%d' % (r0, r1), items=['Data prevista de crédito', 'Outra data (coluna P)'])
    s.validate_list('U%d:U%d' % (r0, r1), items=['Aguardando', 'Creditado', 'Divergente', 'Em contestação'])
    s.validate_date('A%d:A%d' % (r0, r1)); s.validate_date('I%d:I%d' % (r0, r1)); s.validate_date('P%d:P%d' % (r0, r1)); s.validate_date('S%d:S%d' % (r0, r1))
    s.validate_num('E%d:F%d' % (r0, r1), 0); s.validate_num('R%d:R%d' % (r0, r1), 0)
    s.validate_num('J%d:J%d' % (r0, r1), 0, 3650, integer=True, msg='Prazo em número de dias. Use o que o regulamento informar.'); s.validate_num('M%d:M%d' % (r0, r1), 0, 3650, integer=True)
    s.cf('V%d:V%d' % (r0, r1), '=OR(LEFT($V%d,10)="Divergênci",ISNUMBER(SEARCH("passou",$V%d)))' % (r0, r0), {'bg_color': RED_L, 'bold': True})
    s.cf('V%d:V%d' % (r0, r1), '=ISNUMBER(SEARCH("em até 7 dias",$V%d))' % r0, {'bg_color': ORANGE_L})
    s.cf('U%d:U%d' % (r0, r1), '=$U%d="Creditado"' % r0, {'bg_color': GREEN_L}); s.cf('U%d:U%d' % (r0, r1), '=OR($U%d="Divergente",$U%d="Em contestação")' % (r0, r0), {'bg_color': ORANGE_L})
    ws.set_default_row(30)
    s.finish(header_row=6, freeze=(6, 3), autofilter='A6:X%d' % r1, area='A1:X46')
    return s
