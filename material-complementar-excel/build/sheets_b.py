# -*- coding: utf-8 -*-
"""Abas 07 a 11."""
import datetime
from common import *

D = datetime.date
def band(s, row, cols, text, book):
    s.section('A%d' % row, text)
    for c in cols: s.ws.write_blank('%s%d' % (c, row), None, book.f(bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))

# ===================================================================== 07
def sheet07(book):
    s = book.sheet(7); ws = s.ws; q7 = q(7)
    for c, w in {'A': 60, 'B': 30, 'C': 30, 'D': 30, 'E': 58, 'F': 26, 'G': 28, 'H': 44}.items(): ws.set_column('%s:%s' % (c, c), w)
    s.title('Comparar resgates', 'Capítulos 4, 12 e 13 — Dinheiro × pontos ou milhas',
            'Compare opções equivalentes: mesmos aeroportos, datas, passageiros, bagagem e condições. O resultado mostra a economia calculada naquela comparação; ele não decide sozinho.')
    s.back_link('A4')
    band(s, 5, 'BCD', 'A. Ficha da busca (Capítulo 13, Passo 1)', book)
    campos = [('Origem', 'text', 'Campinas'), ('Destino', 'text', 'Salvador'), ('Datas da viagem (ida e volta)', 'text', None), ('Número de passageiros', 'int', 2),
              ('Data da busca', 'date', None), ('Programa e companhia consultados', 'text', 'Programa e companhia fictícios'), ('Observações', 'long', 'SIMULAÇÃO COM VALORES FICTÍCIOS (Capítulo 13). Não é uma busca real.')]
    for i, (lb, kd, ex) in enumerate(campos):
        r = 6 + i; s.label('A%d' % r, lb, wrap=False); s.inp('B%d' % r, kd, ex=ex)
    ws.set_row(11, 62)
    s.validate_num('B9', 1, integer=True); s.validate_date('B10')
    band(s, 14, 'BCDE', 'B. Comparação dinheiro × pontos ou milhas (Capítulo 13, Passos 2 a 5)', book)
    s.head(15, ['Item', 'Dinheiro (opção principal)', 'Pontos ou milhas', 'Alternativa mais barata em dinheiro (opcional)', 'Como ler'], height=40, calc_cols=())
    rows = {
        16: ('Preço total em dinheiro da passagem, com as taxas de embarque incluídas', 'Cap. 4: use o preço total. Em "Alternativa", informe a melhor opção adequada em dinheiro (Cap. 12).'),
        17: ('Dessas taxas, a taxa de embarque que já está dentro do preço acima (opcional)', 'Só para separar a tarifa das taxas (Cap. 13). NÃO é somada de novo ao total.'),
        18: ('Tarifa, sem as taxas de embarque (calculada)', 'Preço total menos a taxa de embarque informada.'),
        19: ('Pontos ou milhas utilizados no resgate (total de todos os passageiros)', 'Quantidade que o programa pede pela mesma passagem.'),
        20: ('Taxas de embarque pagas em dinheiro no resgate', 'O que você ainda paga em dinheiro mesmo usando pontos (Cap. 4).'),
        21: ('Outros pagamentos adicionais do resgate (taxas de serviço, etc.)', 'Digite 0 se não houver.'),
        22: ('Bagagem e serviços', 'Compare com a mesma bagagem nas duas opções: ela entra dos dois lados e se cancela na conta.'),
        23: ('Total pago em dinheiro (calculado)', 'Dinheiro e alternativa: preço total + bagagem. Pontos: taxas + outros pagamentos + bagagem. Campo vazio = não informado, não zero: digite 0 onde não houver.'),
        24: ('Condições relevantes de alteração e cancelamento', 'Resumo. Detalhe e calcule na aba 08. Cite a fonte.'),
        25: ('Observações de cada opção', ''),
    }
    for r, (lb, note) in rows.items():
        s.label('A%d' % r, lb); s.t('E%d' % r, note, italic=True, font_color=GRAY, text_wrap=True, valign='top')
    for rr in range(16, 24): ws.set_row(rr - 1, 34)
    ws.set_row(23, 52); ws.set_row(24, 40)
    s.inp('B16', 'money', ex=1960); s.static_calc('C16', '—'); s.inp('D16', 'money', ex=1490)
    s.inp('B17', 'money', ex=280); s.static_calc('C17', '—'); s.inp('D17', 'money', ex=280)
    s.calc('B18', '=IF(AND(ISNUMBER(B16),ISNUMBER(B17)),B16-B17,"")', 'money'); s.static_calc('C18', '—'); s.calc('D18', '=IF(AND(ISNUMBER(D16),ISNUMBER(D17)),D16-D17,"")', 'money')
    s.static_calc('B19', '—'); s.inp('C19', 'pts', ex=72000); s.static_calc('D19', '—')
    s.static_calc('B20', '—'); s.inp('C20', 'money', ex=280); s.static_calc('D20', '—')
    s.static_calc('B21', '—'); s.inp('C21', 'money', ex=0); s.static_calc('D21', '—')
    s.inp('B22', 'money', ex=360); s.inp('C22', 'money', ex=360); s.inp('D22', 'money', ex=360)
    s.calc('B23', '=IF(AND(ISNUMBER(B16),ISNUMBER(B22)),B16+B22,"")', 'money', bold=True)
    s.calc('C23', '=IF(AND(ISNUMBER(C20),ISNUMBER(C21),ISNUMBER(C22)),C20+C21+C22,"")', 'money', bold=True)
    s.calc('D23', '=IF(AND(ISNUMBER(D16),ISNUMBER(D22)),D16+D22,"")', 'money', bold=True)
    s.inp('B24', 'long', ex='Condições fictícias do Passo 6 (veja a aba 08).' if book.ex else None); s.inp('C24', 'long', ex='Condições fictícias do Passo 6 (veja a aba 08).' if book.ex else None); s.inp('D24', 'long', ex='Não definidas na simulação do livro.' if book.ex else None)
    s.inp('B25', 'long', ex=None); s.inp('C25', 'long', ex=None); s.inp('D25', 'long', ex='Datas flexíveis (Passo 5): R$ 745 por pessoa, R$ 1.490 no total, com as mesmas condições de bagagem.' if book.ex else None)
    s.validate_num('B16:B17', 0); s.validate_num('D16:D17', 0); s.validate_num('C19', 0); s.validate_num('C20:C21', 0); s.validate_num('B22:D22', 0)
    band(s, 27, 'BCDE', 'C. Resultados (calculados)', book)
    res = [(28, 'Economia calculada contra a opção em dinheiro (R$)', '=IF(AND(ISNUMBER(B23),ISNUMBER(C23),N(C19)>0),B23-C23,"")', 'money', 'Total em dinheiro da opção principal menos o total pago no resgate.'),
           (29, 'Valor obtido por milheiro, contra a opção em dinheiro (R$)', '=IF(AND(ISNUMBER(B28),N(C19)>0),B28/(C19/1000),"")', 'money', 'Valor por milheiro = (preço total em dinheiro da opção comparada − total ainda pago em dinheiro no resgate) ÷ (pontos ou milhas utilizados ÷ 1.000).'),
           (30, 'Economia calculada contra a alternativa mais barata (R$)', '=IF(AND(ISNUMBER(D23),ISNUMBER(C23),N(C19)>0),D23-C23,"")', 'money', 'Mesma conta, usando a melhor alternativa adequada em dinheiro (Cap. 12 e 13, Passo 5).'),
           (31, 'Valor obtido por milheiro, contra a alternativa mais barata (R$)', '=IF(AND(ISNUMBER(B30),N(C19)>0),B30/(C19/1000),"")', 'money', 'O valor muda conforme com o quê você compara. Não é o custo de aquisição dos pontos.'),
           (32, 'Diferença entre os dois valores por milheiro (R$)', '=IF(AND(ISNUMBER(B29),ISNUMBER(B31)),B29-B31,"")', 'money', 'Mostra o efeito de comparar com uma alternativa mais barata.')]
    for r, lb, fm, kd, note in res:
        s.label('A%d' % r, lb, bold=True); s.calc('B%d' % r, fm, kd, bold=True); s.t('E%d' % r, note, italic=True, font_color=GRAY, text_wrap=True, valign='top')
        ws.set_row(r - 1, 46)
    avisos = [(33, '=IF(N($C$19)<=0,"Informe os pontos ou milhas utilizados (maior que zero) para calcular o valor por milheiro.","")'),
              (34, '=IF(AND(N($C$19)>0,OR(NOT(ISNUMBER($C$20)),NOT(ISNUMBER($C$21)),NOT(ISNUMBER($C$22)))),"Campo vazio = não informado, não zero: preencha as taxas de embarque, os outros pagamentos e a bagagem do resgate (digite 0 onde não houver).","")'),
              (35, '=IF(OR(AND(ISNUMBER($B$16),NOT(ISNUMBER($B$22))),AND(ISNUMBER($D$16),NOT(ISNUMBER($D$22)))),"Informe a bagagem e os serviços das opções em dinheiro (digite 0 se não houver) para calcular os totais.",IF(AND(ISNUMBER($B$16),NOT(ISNUMBER($D$16))),"Informe a melhor alternativa em dinheiro para ver o resultado contra ela (Capítulo 12).",""))'),
              (36, '=IF(AND(ISNUMBER($B$17),ISNUMBER($B$16),$B$17>$B$16),"A taxa de embarque informada é maior que o preço total em dinheiro: confira para não contar a taxa duas vezes.","")'),
              (37, '=IF(AND(ISNUMBER($B$29),ISNUMBER($B$31)),"Um resultado positivo indica economia calculada nesta comparação; ele não significa, sozinho, que o resgate seja a melhor escolha (custo de acumular, outros usos, validade e orçamento: Capítulo 4).","")')]
    for r, fm in avisos: s.calc('A%d' % r, fm, 'long')
    for r in (33, 34, 35, 36): ws.set_row(r - 1, 30)
    ws.set_row(36, 44)
    band(s, 39, 'BCDE', 'D. Tabela do Capítulo 13 (reproduzida): comparação dinheiro × milhas', book)
    s.head(40, ['', 'Dinheiro', 'Milhas'], height=22)
    s.label('A41', 'Tarifa em reais (sem as taxas de embarque)', wrap=False); s.calc('B41', '=IF(ISNUMBER($B$18),$B$18,"")', 'money'); s.label('C41', 'Paga com milhas', wrap=False); ws.write_formula('C46', '=IF(N($C$19)>0,$C$19,"")', book.calc_fmt('pts'), book.values.get((NAMES[7], 'C46'), '')); book.formula_cells.append((NAMES[7], 'C46'))
    s.label('A46', 'Milhas ou pontos utilizados (a tarifa é paga com eles)', wrap=False); s.label('B46', 'Não se aplica', wrap=False)
    ws.write_comment('C46', 'Pontos ou milhas utilizados. Na versão 4 do texto do livro, a tarifa em reais e a quantidade de milhas ficam em linhas separadas.')
    s.label('A42', 'Taxas de embarque', wrap=False); s.calc('B42', '=IF(ISNUMBER($B$17),$B$17,"")', 'money'); s.calc('C42', '=IF(ISNUMBER($C$20),$C$20,"")', 'money')
    s.label('A43', 'Bagagem', wrap=False); s.calc('B43', '=IF(ISNUMBER($B$22),$B$22,"")', 'money'); s.calc('C43', '=IF(ISNUMBER($C$22),$C$22,"")', 'money')
    s.label('A44', 'Pago em dinheiro', bold=True, wrap=False); s.calc('B44', '=IF(ISNUMBER($B$23),$B$23,"")', 'money', bold=True); s.calc('C44', '=IF(ISNUMBER($C$23),$C$23,"")', 'money', bold=True)
    s.label('A45', 'Composição do "Pago em dinheiro"', wrap=False)
    s.calc('B45', '=IF(NOT(ISNUMBER($B$23)),"",IF(ISNUMBER($B$18),"R$ "&FIXED($B$18,2)&" + R$ "&FIXED(N($B$17),2)&" + R$ "&FIXED(N($B$22),2),"R$ "&FIXED($B$16,2)&" + R$ "&FIXED(N($B$22),2)))', 'long')
    s.calc('C45', '=IF(NOT(ISNUMBER($C$23)),"",IF(N($C$21)>0,"R$ "&FIXED(N($C$20),2)&" + R$ "&FIXED($C$21,2)&" + R$ "&FIXED(N($C$22),2),"R$ "&FIXED(N($C$20),2)&" + R$ "&FIXED(N($C$22),2)))', 'long')
    ws.set_row(44, 44)
    s.t('E41', 'A tarifa em reais e a quantidade de milhas ficam em linhas separadas (linhas 41 e 46), para não misturar unidades.', italic=True, font_color=GRAY, text_wrap=True)
    band(s, 47, 'BCDEFGH', 'E. Exercício rápido do Capítulo 4 (várias comparações; uma por linha)', book)
    s.head(48, ['Descrição', 'Preço total em dinheiro (R$)', 'Pagamento adicional no resgate (R$; 0 se não houver)', 'Pontos ou milhas necessários', 'Valor por milheiro (R$)', 'Alternativa em dinheiro (opcional, R$)', 'Valor por milheiro contra a alternativa (R$)', 'Observações'], height=52, calc_cols=(4, 6))
    ex = {49: ('Exemplo A (Capítulo 4)', 1200, 100, 20000, None, '[Exemplo do livro] Resultado esperado: R$ 55 por milheiro.'),
          50: ('Exemplo B (Capítulo 4)', 600, 80, 15000, None, '[Exemplo do livro] Resultado esperado: ≈ R$ 34,67 por milheiro.'),
          51: ('Exemplo C (Capítulo 4): Exemplo A com alternativa de R$ 700', 1200, 100, 20000, 700, '[Exemplo do livro] O livro só diz que a alternativa de R$ 700 deve ser considerada. O valor contra a alternativa (R$ 30,00) é o resultado da mesma fórmula, calculado aqui.')}
    for r in range(49, 59):
        e = ex.get(r) if book.ex else None
        e = e or (None,) * 6
        s.inp('A%d' % r, 'text', ex=e[0]); s.inp('B%d' % r, 'money', ex=e[1]); s.inp('C%d' % r, 'money', ex=e[2]); s.inp('D%d' % r, 'pts', ex=e[3])
        s.calc('E%d' % r, '=IF(AND(ISNUMBER($B{r}),ISNUMBER($C{r}),ISNUMBER($D{r}),N($D{r})>0),($B{r}-$C{r})/($D{r}/1000),"")'.format(r=r), 'money', bold=True)
        s.inp('F%d' % r, 'money', ex=e[4])
        s.calc('G%d' % r, '=IF(AND(ISNUMBER($F{r}),ISNUMBER($C{r}),ISNUMBER($D{r}),N($D{r})>0),($F{r}-$C{r})/($D{r}/1000),"")'.format(r=r), 'money', bold=True)
        s.inp('H%d' % r, 'long', ex=e[5])
        ws.set_row(r - 1, 30 if r != 51 else 62)
    s.validate_num('B49:D58', 0); s.validate_num('F49:F58', 0)
    band(s, 60, 'BCDE', 'F. As sete perguntas do Capítulo 12: dinheiro ou milhas? (registre a sua resposta)', book)
    s.head(61, ['Pergunta', 'Minha resposta', 'Observação ou fonte'], height=24)
    qs = parse_cap12_questions(); assert len(qs) == 7
    for k, qq in enumerate(qs):
        r = 62 + k
        ws.write('A%d' % r, '%d. %s' % (k + 1, qq), book.f(bold=True, border=1, border_color=BORDER, bg_color=ROSE_L, text_wrap=True, valign='top'))
        s.inp('B%d' % r, 'long'); s.inp('C%d' % r, 'long'); ws.set_row(r - 1, 36)
    band(s, 70, 'BCDE', 'G. Decisão (Capítulo 13, Passo 7; Capítulo 18, semana 4)', book)
    s.label('A71', 'Decisão', bold=True, wrap=False); s.inp('B71', 'text', ex='Esperar' if book.ex else None)
    s.validate_list('B71', items=['Resgatar com pontos ou milhas', 'Pagar em dinheiro', 'Esperar'])
    s.label('A72', 'Motivo', bold=True, wrap=False); s.inp('B72', 'long', ex='[Exemplo do livro, Cap. 18] A Marina decidiu esperar uma promoção, sem pagar nada por clube nem por compra de pontos.' if book.ex else None); ws.set_row(71, 62)
    s.label('A73', 'Data da decisão', bold=True, wrap=False); s.inp('B73', 'date'); s.validate_date('B73')
    s.t('E71', 'Não há resposta única (Capítulo 13): o que importa é ter feito a conta com as alternativas e conhecer as condições.', italic=True, font_color=GRAY, text_wrap=True)
    s.finish(area='A1:H73')
    return s

# ===================================================================== 08
def sheet08(book):
    s = book.sheet(8); ws = s.ws; q7 = q(7)
    for c, w in {'A': 64, 'B': 34, 'C': 34, 'D': 56}.items(): ws.set_column('%s:%s' % (c, c), w)
    s.title('Alterar ou cancelar', 'Capítulos 13 a 15 — Alteração, cancelamento e bagagem',
            'Registre as condições da sua opção, com fonte e data. Esta aba calcula o custo e a perda NAS CONDIÇÕES QUE VOCÊ INFORMAR. Não conclui direitos legais do passageiro (veja o quadro do Capítulo 14).')
    s.back_link('A4')
    s.t('A5', 'As condições do exemplo (versão de exemplos) são fictícias, valem só para aquela simulação e não são regras de programas ou companhias.', italic=True, font_color=GRAY)
    band(s, 6, 'BCD', 'A. Dados vindos da aba 07 (calculados; para mudar, altere a aba 07)', book)
    s.head(7, ['Item', 'Dinheiro', 'Milhas', 'Observação'], height=22)
    a = [(8, 'Passageiros', '=IF(ISNUMBER(%s!$B$9),%s!$B$9,"")' % (q7, q7), None, 'int'),
         (9, 'Tarifa em dinheiro (sem as taxas de embarque)', '=IF(ISNUMBER(%s!$B$18),%s!$B$18,"")' % (q7, q7), None, 'money'),
         (10, 'Taxas de embarque', '=IF(ISNUMBER(%s!$B$17),%s!$B$17,"")' % (q7, q7), '=IF(ISNUMBER(%s!$C$20),%s!$C$20,"")' % (q7, q7), 'money'),
         (11, 'Outros pagamentos adicionais do resgate', None, '=IF(ISNUMBER(%s!$C$21),%s!$C$21,"")' % (q7, q7), 'money'),
         (12, 'Bagagem e serviços pagos', '=IF(ISNUMBER(%s!$B$22),%s!$B$22,"")' % (q7, q7), '=IF(ISNUMBER(%s!$C$22),%s!$C$22,"")' % (q7, q7), 'money'),
         (13, 'Pontos ou milhas utilizados', None, '=IF(N(%s!$C$19)>0,%s!$C$19,"")' % (q7, q7), 'pts'),
         (14, 'Total pago em dinheiro antes da bagagem (tarifa + taxas de embarque)', '=IF(AND(ISNUMBER($B$9),ISNUMBER($B$10)),$B$9+$B$10,"")', '=IF(AND(ISNUMBER($C$10),ISNUMBER($C$11)),$C$10+$C$11,"")', 'money'),
         (15, 'Total pago em dinheiro, com a bagagem', '=IF(AND(ISNUMBER($B$14),ISNUMBER($B$12)),$B$14+$B$12,"")', '=IF(ISNUMBER(%s!$C$23),%s!$C$23,"")' % (q7, q7), 'money')]
    for r, lb, fb, fc, kd in a:
        s.label('A%d' % r, lb, wrap=False)
        if fb: s.calc('B%d' % r, fb, kd)
        else: s.static_calc('B%d' % r, '—')
        if fc: s.calc('C%d' % r, fc, kd)
        else: s.static_calc('C%d' % r, '—' if r not in (8, 9) else '')
    s.calc('C8', '=B8', 'int'); s.static_calc('C9', '—')
    s.t('D14', 'Livro (fictício): R$ 1.960 antes da bagagem na opção em dinheiro.', italic=True, font_color=GRAY, text_wrap=True)
    s.t('D15', 'Dinheiro: tarifa + taxas + bagagem. Milhas: taxas + outros pagamentos + bagagem.', italic=True, font_color=GRAY, text_wrap=True)
    band(s, 16, 'BCD', 'B. Condições informadas por você', book)
    s.head(17, ['Item', 'Dinheiro', 'Milhas', 'Observação'], height=22)
    inp = [(18, 'Taxa de alteração por pessoa (R$)', 'money', 150, 100), (19, 'Diferença de tarifa, se houver (R$; 0 se não houver)', 'money', 0, '—'), (20, 'Diferença de pontos ou milhas, se houver (positivo = pagar mais)', 'pts', '—', None),
           (21, 'Outros custos de alteração (R$; 0 se não houver)', 'money', 0, 0), (22, 'Cancelamento: percentual da tarifa devolvido', 'pct', 0.8, '—'),
           (23, 'Cancelamento: taxas de embarque devolvidas (R$)', 'money', 280, 280), (24, 'Cancelamento: pontos ou milhas devolvidos', 'pts', '—', 72000),
           (25, 'Cancelamento: taxa de cancelamento por pessoa (R$; 0 se não houver)', 'money', 0, 150), (26, 'Cancelamento: outros custos (R$; 0 se não houver)', 'money', 0, 0),
           (27, 'Bagagem não usada: valor devolvido (R$)', 'money', 360, 360)]
    for r, lb, kd, eb, ec in inp:
        s.label('A%d' % r, lb, wrap=False)
        for col, e in (('B', eb), ('C', ec)):
            if e == '—': s.static_calc('%s%d' % (col, r), '—')
            else: s.inp('%s%d' % (col, r), kd, ex=e)
    s.t('D18', 'Campo vazio = NÃO INFORMADO (não é zero). Digite 0 quando não houver o custo.', italic=True, font_color=GRAY)
    t = parse_table('| | Dinheiro | Milhas |', 2)
    txt = {r[0]: r for r in t[1:]}
    s.label('A28', 'Condições e fonte: Alteração (2 pessoas)', bold=True); s.label('A29', 'Condições e fonte: Cancelamento antes do embarque', bold=True); s.label('A30', 'Condições e fonte: Bagagem não usada', bold=True)
    for r, key in ((28, 'Alteração (2 pessoas)'), (29, 'Cancelamento antes do embarque'), (30, 'Bagagem não usada')):
        for col, idx in (('B', 1), ('C', 2)):
            s.inp('%s%d' % (col, r), 'long', ex=txt[key][idx] if book.ex else None)
        ws.set_row(r - 1, 118)
    s.label('A31', 'Fonte e data da consulta (de cada opção)', bold=True); s.inp('B31', 'long', ex='Simulação fictícia do Capítulo 13, Passo 6.' if book.ex else None); s.inp('C31', 'long', ex='Simulação fictícia do Capítulo 13, Passo 6.' if book.ex else None)
    s.validate_num('B18:C27', 0); s.validate_num('B22', 0, 1, msg='Digite o percentual, por exemplo 80%.')
    band(s, 33, 'BCD', 'C. Resultados calculados (nas condições que você informou)', book)
    s.head(34, ['Item', 'Dinheiro', 'Milhas', 'Observação'], height=22)
    res = [(35, 'Alteração: taxa total (por pessoa × passageiros) (R$)', '=IF(AND(ISNUMBER($B$18),ISNUMBER($B$8)),$B$18*$B$8,"")', '=IF(AND(ISNUMBER($C$18),ISNUMBER($B$8)),$C$18*$B$8,"")', 'money', 'Livro (fictício): R$ 300 e R$ 200.'),
           (36, 'Alteração: custo adicional total em dinheiro (R$)', '=IF(AND(ISNUMBER($B$35),ISNUMBER($B$19),ISNUMBER($B$21)),$B$35+$B$19+$B$21,"")', '=IF(AND(ISNUMBER($C$35),ISNUMBER($C$21)),$C$35+$C$21,"")', 'money', 'Dinheiro: taxa total + diferença de tarifa + outros custos. Milhas: taxa total + outros custos.'),
           (37, 'Alteração: diferença de pontos ou milhas', None, '=IF(ISNUMBER($C$20),$C$20,"")', 'pts', 'Informada por você; a diferença em pontos não vira reais aqui.'),
           (38, 'Cancelamento: valor devolvido em dinheiro, sem a bagagem (R$)', '=IF(AND(ISNUMBER($B$9),ISNUMBER($B$22),ISNUMBER($B$23)),$B$9*$B$22+$B$23,"")', '=IF(ISNUMBER($C$23),$C$23,"")', 'money', 'Dinheiro: tarifa × percentual devolvido + taxas devolvidas. Livro (fictício): R$ 1.624 em dinheiro.'),
           (39, 'Cancelamento: taxa de cancelamento total (por pessoa × passageiros) (R$)', '=IF(AND(ISNUMBER($B$25),ISNUMBER($B$8)),$B$25*$B$8,"")', '=IF(AND(ISNUMBER($C$25),ISNUMBER($B$8)),$C$25*$B$8,"")', 'money', 'Livro (fictício): R$ 300 na opção com milhas.'),
           (40, 'Cancelamento: pontos ou milhas devolvidos', None, '=IF(ISNUMBER($C$24),$C$24,"")', 'pts', ''),
           (41, 'Cancelamento: pontos ou milhas não devolvidos', None, '=IF(AND(ISNUMBER($C$13),ISNUMBER($C$24)),$C$13-$C$24,"")', 'pts', ''),
           (42, 'Bagagem não usada: valor não devolvido (R$)', '=IF(AND(ISNUMBER($B$12),ISNUMBER($B$27)),$B$12-$B$27,"")', '=IF(AND(ISNUMBER($C$12),ISNUMBER($C$27)),$C$12-$C$27,"")', 'money', 'Zero quando a bagagem é devolvida por inteiro.'),
           (43, 'Perda calculada em dinheiro no cancelamento (R$)', '=IF(AND(ISNUMBER($B$14),ISNUMBER($B$22),ISNUMBER($B$23),ISNUMBER($B$26),ISNUMBER($B$39),ISNUMBER($B$42)),$B$14-($B$9*$B$22+$B$23)+$B$39+$B$26+$B$42,"")',
            '=IF(AND(ISNUMBER($C$14),ISNUMBER($C$23),ISNUMBER($C$26),ISNUMBER($C$39),ISNUMBER($C$42)),$C$14-$C$23+$C$39+$C$26+$C$42,"")', 'money', 'Só é calculada com TODOS os campos informados (campo vazio = não informado; digite 0 se não houver). Dinheiro: tarifa + taxas − devolvido + taxa de cancelamento + outros custos + bagagem não devolvida. Livro (fictício): R$ 336 e R$ 300.'),
           (44, 'Perda calculada em pontos ou milhas', None, '=IF(ISNUMBER($C$41),$C$41,"")', 'pts', 'Pontos devolvidos voltam para a conta e precisam ser usados dentro da validade do lote (Capítulo 10).')]
    for r, lb, fb, fc, kd, note in res:
        s.label('A%d' % r, lb, bold=(r in (43, 44)), wrap=False)
        if fb: s.calc('B%d' % r, fb, kd, bold=(r == 43))
        else: s.static_calc('B%d' % r, '—')
        s.calc('C%d' % r, fc, kd, bold=(r in (43, 44)))
        s.t('D%d' % r, note, italic=True, font_color=GRAY, text_wrap=True)
        ws.set_row(r - 1, 32)
    PEND = ('=IF(AND(NOT(ISNUMBER($B$9)),NOT(ISNUMBER($C$13))),"",IF(LEN('
            'IF(ISNUMBER($B$18),"","taxa de alteração por pessoa (dinheiro); ")&IF(ISNUMBER($B$19),"","diferença de tarifa (dinheiro); ")&IF(ISNUMBER($B$21),"","outros custos de alteração (dinheiro); ")'
            '&IF(ISNUMBER($C$18),"","taxa de alteração por pessoa (milhas); ")&IF(ISNUMBER($C$21),"","outros custos de alteração (milhas); ")'
            '&IF(ISNUMBER($B$22),"","percentual da tarifa devolvido (dinheiro); ")&IF(ISNUMBER($B$23),"","taxas devolvidas (dinheiro); ")&IF(ISNUMBER($B$25),"","taxa de cancelamento (dinheiro); ")&IF(ISNUMBER($B$26),"","outros custos de cancelamento (dinheiro); ")&IF(ISNUMBER($B$27),"","bagagem devolvida (dinheiro); ")'
            '&IF(ISNUMBER($C$23),"","taxas devolvidas (milhas); ")&IF(ISNUMBER($C$24),"","pontos devolvidos (milhas); ")&IF(ISNUMBER($C$25),"","taxa de cancelamento (milhas); ")&IF(ISNUMBER($C$26),"","outros custos de cancelamento (milhas); ")&IF(ISNUMBER($C$27),"","bagagem devolvida (milhas); "))=0,"",'
            '"Campo vazio = não informado, não zero. Sem estes campos a perda não é calculada (digite 0 se não houver): "&'
            'IF(ISNUMBER($B$18),"","taxa de alteração por pessoa (dinheiro); ")&IF(ISNUMBER($B$19),"","diferença de tarifa (dinheiro); ")&IF(ISNUMBER($B$21),"","outros custos de alteração (dinheiro); ")'
            '&IF(ISNUMBER($C$18),"","taxa de alteração por pessoa (milhas); ")&IF(ISNUMBER($C$21),"","outros custos de alteração (milhas); ")'
            '&IF(ISNUMBER($B$22),"","percentual da tarifa devolvido (dinheiro); ")&IF(ISNUMBER($B$23),"","taxas devolvidas (dinheiro); ")&IF(ISNUMBER($B$25),"","taxa de cancelamento (dinheiro); ")&IF(ISNUMBER($B$26),"","outros custos de cancelamento (dinheiro); ")&IF(ISNUMBER($B$27),"","bagagem devolvida (dinheiro); ")'
            '&IF(ISNUMBER($C$23),"","taxas devolvidas (milhas); ")&IF(ISNUMBER($C$24),"","pontos devolvidos (milhas); ")&IF(ISNUMBER($C$25),"","taxa de cancelamento (milhas); ")&IF(ISNUMBER($C$26),"","outros custos de cancelamento (milhas); ")&IF(ISNUMBER($C$27),"","bagagem devolvida (milhas); ")))')
    s.calc('A45', PEND, 'long'); ws.set_row(44, 78)
    s.t('A46', 'Atenção: o resultado vale só para as condições digitadas. Esta aba não calcula nem conclui direitos legais do passageiro: consulte o quadro do Capítulo 14, a companhia e o programa.', bold=True, font_color=COPPER)
    band(s, 48, 'BCD', 'D. Registro de contatos e protocolos (Capítulo 15: "como pedir ajuda")', book)
    s.head(49, ['Assunto (código da reserva, nome dos passageiros e datas, se for o caso)', 'Data e hora do contato', 'Canal oficial usado', 'Protocolo, nome do atendente e resultado'], height=40)
    for r in range(50, 60):
        s.inp('A%d' % r, 'long'); s.inp('B%d' % r, 'text'); s.inp('C%d' % r, 'text'); s.inp('D%d' % r, 'long'); ws.set_row(r - 1, 32)
    s.t('A61', 'Use somente canais oficiais. Nunca informe senhas ou códigos a quem entrar em contato com você. Guarde comprovantes, e-mails de confirmação e prints das condições.', italic=True, font_color=GRAY)
    s.finish(area='A1:D61')
    return s

# ===================================================================== 09
def sheet09(book):
    s = book.sheet(9); ws = s.ws
    widths = {'A': 30, 'B': 14, 'C': 12, 'D': 14, 'E': 14, 'F': 14, 'G': 16, 'H': 16, 'I': 16, 'J': 16, 'K': 40, 'L': 26, 'M': 30, 'N': 30, 'O': 18, 'P': 16, 'Q': 30}
    for c, w in widths.items(): ws.set_column('%s:%s' % (c, c), w)
    s.title('Clubes e compra de pontos', 'Capítulos 16 e 17 — Calcule o custo por milheiro no período COMPLETO',
            'Nada aqui é obrigatório: só faça se a conta fechar. Parcelar sem juros não é grátis: use o valor total pago. Um bônus inicial entra UMA vez, como total do período, e não como se fosse mensal.')
    s.back_link('A4')
    band(s, 6, 'BCDEFGHIJ', 'A. Comparar clubes (uma linha por clube)', book)
    s.head(7, ['Clube (nome e programa)', 'Mensalidade (R$)', 'Meses analisados', 'Permanência mínima (meses)', 'Outros custos no período (R$; 0 se não houver)', 'Pontos regulares por mês', 'Bônus elegíveis no período (pontos, total; 0 se não houver)',
               'Custo total no período (R$)', 'Total recebido no período (pontos)', 'Custo por milheiro (R$)', 'Alerta', 'Validade', 'Regras de cancelamento', 'Uso concreto pretendido', 'Valor obtido no uso, por milheiro (R$, opcional)', 'Diferença (valor obtido − custo) (R$)', 'Fonte e data'],
           height=66, calc_cols=(7, 8, 9, 10, 15))
    ex = {8: ('Clube de exemplo (Capítulo 16)', 42, 1, None, 0, 1000, 0, 'Informar conforme a fonte oficial', 'Informar conforme a fonte oficial', 'Nenhum resgate concreto identificado', 23.33, '[Exemplo do livro] R$ 42 por mês, 1.000 pontos por mês: R$ 42 por milheiro. O valor obtido de R$ 23,33 vem da simulação do Capítulo 13.'),
          9: ('Clube B (exemplo adicional fictício)', 42, 12, 3, 0, 1000, 1000, 'Informar conforme a fonte oficial', 'Informar conforme a fonte oficial', None, None, '[Exemplo adicional fictício] Bônus de 1.000 pontos só no primeiro mês, contado uma vez no período de 12 meses.')}
    for r in range(8, 18):
        e = ex.get(r) if book.ex else None
        e = e or (None,) * 12
        s.inp('A%d' % r, 'text', ex=e[0]); s.inp('B%d' % r, 'money', ex=e[1]); s.inp('C%d' % r, 'int', ex=e[2]); s.inp('D%d' % r, 'int', ex=e[3]); s.inp('E%d' % r, 'money', ex=e[4]); s.inp('F%d' % r, 'pts', ex=e[5]); s.inp('G%d' % r, 'pts', ex=e[6])
        s.calc('H%d' % r, '=IF(AND(ISNUMBER($B{r}),ISNUMBER($C{r}),ISNUMBER($E{r})),$B{r}*$C{r}+$E{r},"")'.format(r=r), 'money')
        s.calc('I%d' % r, '=IF(AND(ISNUMBER($C{r}),ISNUMBER($F{r}),ISNUMBER($G{r})),$F{r}*$C{r}+$G{r},"")'.format(r=r), 'pts')
        s.calc('J%d' % r, '=IF(AND(ISNUMBER($H{r}),ISNUMBER($I{r}),N($I{r})>0),$H{r}/($I{r}/1000),"")'.format(r=r), 'money', bold=True)
        s.calc('K%d' % r, '=IF($A{r}="","",IF(OR(NOT(ISNUMBER($B{r})),NOT(ISNUMBER($C{r}))),"Informe a mensalidade e os meses analisados.",IF(OR(NOT(ISNUMBER($E{r})),NOT(ISNUMBER($F{r})),NOT(ISNUMBER($G{r}))),"Campo vazio = não informado, não zero: preencha outros custos, pontos regulares e bônus (digite 0 se não houver).",IF(AND(ISNUMBER($D{r}),$C{r}<$D{r}),"O período analisado é menor que a permanência mínima: analise o período completo do compromisso.",IF(AND($G{r}>0,$C{r}<12),"Há bônus e o período é curto: confira se o bônus se repete ou vale só uma vez.","")))))'.format(r=r), 'long')
        s.inp('L%d' % r, 'long', ex=e[7]); s.inp('M%d' % r, 'long', ex=e[8]); s.inp('N%d' % r, 'long', ex=e[9]); s.inp('O%d' % r, 'money', ex=e[10])
        s.calc('P%d' % r, '=IF(AND(ISNUMBER($O{r}),ISNUMBER($J{r})),$O{r}-$J{r},"")'.format(r=r), 'money')
        s.inp('Q%d' % r, 'long', ex=e[11]); ws.set_row(r - 1, 48)
    s.validate_num('B8:B17', 0); s.validate_num('C8:D17', 0, 600, integer=True); s.validate_num('E8:G17', 0); s.validate_num('O8:O17', 0)
    ws.set_row(6, 66)
    band(s, 20, 'BCDEFGHIJ', 'B. Compra de pontos ou milhas (uma linha por compra)', book)
    s.head(21, ['Programa', 'Quantidade comprada (pontos)', 'Bônus (pontos; 0 se não houver)', 'Valor total pago (R$)', 'Taxas e outros custos (R$; 0 se não houver)', 'Total recebido (pontos)', 'Custo total (R$)', 'Custo por milheiro (R$)',
                'Prazo de crédito', 'Validade', 'Resgate pretendido', 'Valor obtido no resgate, por milheiro (R$, opcional)', 'Diferença (valor obtido − custo) (R$)', 'Fonte e data'], height=66, calc_cols=(5, 6, 7, 12))
    exb = {22: ('Programa de exemplo (adicional fictício)', 20000, 0, 600, 0, 'Informar conforme a fonte oficial', 'Informar conforme a fonte oficial', 'Nenhum resgate concreto identificado', None, '[Exemplo adicional fictício] Mesmos R$ 600 por 20.000 pontos do exemplo do Capítulo 17 (R$ 30 por milheiro), tratados aqui como uma compra.')}
    for r in range(22, 32):
        e = exb.get(r) if book.ex else None
        e = e or (None,) * 10
        s.inp('A%d' % r, 'text', ex=e[0]); s.inp('B%d' % r, 'pts', ex=e[1]); s.inp('C%d' % r, 'pts', ex=e[2]); s.inp('D%d' % r, 'money', ex=e[3]); s.inp('E%d' % r, 'money', ex=e[4])
        s.calc('F%d' % r, '=IF(AND(ISNUMBER($B{r}),ISNUMBER($C{r})),$B{r}+$C{r},"")'.format(r=r), 'pts')
        s.calc('G%d' % r, '=IF(AND(ISNUMBER($D{r}),ISNUMBER($E{r})),$D{r}+$E{r},"")'.format(r=r), 'money')
        s.calc('H%d' % r, '=IF(AND(ISNUMBER($G{r}),ISNUMBER($F{r}),N($F{r})>0),$G{r}/($F{r}/1000),"")'.format(r=r), 'money', bold=True)
        s.inp('I%d' % r, 'long', ex=e[5]); s.inp('J%d' % r, 'long', ex=e[6]); s.inp('K%d' % r, 'long', ex=e[7]); s.inp('L%d' % r, 'money', ex=e[8])
        s.calc('M%d' % r, '=IF(AND(ISNUMBER($L{r}),ISNUMBER($H{r})),$L{r}-$H{r},"")'.format(r=r), 'money'); s.inp('N%d' % r, 'long', ex=e[9]); ws.set_row(r - 1, 48)
    s.validate_num('B22:E31', 0); s.validate_num('L22:L31', 0)
    s.t('A33', 'Fórmula: custo por milheiro = custo total ÷ (pontos ou milhas recebidos ÷ 1.000). O custo total inclui tudo o que foi pago. Compare com o valor obtido no resgate (aba 07), sempre na mesma unidade: reais por milheiro.', italic=True, font_color=GRAY)
    s.finish(area='A1:Q33', freeze=(0, 1))
    return s

# ===================================================================== 10
def sheet10(book):
    s = book.sheet(10); ws = s.ws
    widths = {'A': 22, 'B': 22, 'C': 14, 'D': 16, 'E': 12, 'F': 18, 'G': 16, 'H': 16, 'I': 14, 'J': 14, 'K': 16, 'L': 14, 'M': 14, 'N': 16, 'O': 16, 'P': 44, 'Q': 22, 'R': 22, 'S': 22, 'T': 22, 'U': 18, 'V': 16, 'W': 30, 'X': 34}
    for c, w in widths.items(): ws.set_column('%s:%s' % (c, c), w)
    s.title('Transferências bonificadas', 'Capítulo 17 — Quantas milhas chegam e quanto custa cada milheiro',
            'Fator de conversão = quantas milhas ou pontos do DESTINO chegam por 1 ponto de ORIGEM, antes do bônus (1 = "1 para 1"; 0,5 = 2 pontos de origem para 1 no destino). O bônus é um percentual sobre o que se transfere.')
    s.back_link('A4')
    s.t('A5', 'Fórmula básica: milhas recebidas = pontos transferidos × conversão × (1 + bônus). Com limite de bônus, esta aba calcula: base + (pontos elegíveis ao bônus × conversão × bônus). Se não houver limite informado, o cálculo assume que todos os pontos recebem bônus: confirme no regulamento da promoção. A fórmula simples só vale quando não há limite de bônus aplicável; siga a unidade e as condições da campanha (o limite pode ser em pontos de origem ou em milhas de bônus: converta antes de digitar).', italic=True, font_color=GRAY)
    hdr = ['Programa de origem', 'Programa de destino', 'Pontos transferidos', 'Fator de conversão (milhas por ponto de origem)', 'Percentual de bônus', 'Limite de pontos de origem que recebem bônus (opcional)',
           'Custo dos pontos de origem (R$ total)', 'Taxas e gastos adicionais (R$; 0 se não houver)', 'Cadastro prévio na promoção necessário?',
           'Milhas-base', 'Pontos elegíveis ao bônus', 'Total do bônus', 'Total recebido', 'Custo total (R$)', 'Custo por milheiro no destino (R$)', 'Aviso',
           'Prazo de crédito da base', 'Prazo de crédito do bônus', 'Validade da base', 'Validade do bônus', 'Valor obtido no resgate, por milheiro (R$, opcional)', 'Diferença (valor obtido − custo) (R$)', 'Fonte e data', 'Observações']
    s.head(7, hdr, height=74, calc_cols=(9, 10, 11, 12, 13, 14, 15, 21))
    ex = {8: ('Origem (fictício)', 'Destino (fictício)', 20000, 1, 0.3, None, 600, 0, 'Não sei', 'Informar conforme a promoção', 'Informar conforme a promoção', 'Informar conforme a promoção', 'Informar conforme a promoção', 16.81, '[Exemplo do livro, Cap. 17] R$ 600 por 20.000 pontos (R$ 30 por milheiro), conversão 1 para 1, bônus de 30%: 26.000 milhas, R$ 23,08 por milheiro. O valor de R$ 16,81 vem do Capítulo 13 (alternativa em dinheiro mais barata).'),
          9: ('Origem (fictício)', 'Destino (fictício)', 20000, 1, 1.0, None, 600, 0, 'Não sei', 'Informar conforme a promoção', 'Informar conforme a promoção', 'Informar conforme a promoção', 'Informar conforme a promoção', 16.81, '[Exemplo do livro, Cap. 17] Bônus de 100%: 40.000 milhas, R$ 15,00 por milheiro. Diferença de cerca de R$ 1,81 por milheiro, antes de outros custos.')}
    for r in range(8, 23):
        e = ex.get(r) if book.ex else None
        e = e or (None,) * 15
        s.inp('A%d' % r, 'text', ex=e[0]); s.inp('B%d' % r, 'text', ex=e[1]); s.inp('C%d' % r, 'pts', ex=e[2]); s.inp('D%d' % r, 'dec', ex=e[3]); s.inp('E%d' % r, 'pct', ex=e[4]); s.inp('F%d' % r, 'pts', ex=e[5])
        s.inp('G%d' % r, 'money', ex=e[6]); s.inp('H%d' % r, 'money', ex=e[7]); s.inp('I%d' % r, 'text', ex=e[8])
        s.calc('J%d' % r, '=IF(AND(ISNUMBER($C{r}),ISNUMBER($D{r})),$C{r}*$D{r},"")'.format(r=r), 'pts')
        s.calc('K%d' % r, '=IF(ISNUMBER($C{r}),IF(ISNUMBER($F{r}),MIN($C{r},$F{r}),$C{r}),"")'.format(r=r), 'pts')
        s.calc('L%d' % r, '=IF(AND(ISNUMBER($K{r}),ISNUMBER($D{r}),ISNUMBER($E{r})),$K{r}*$D{r}*$E{r},"")'.format(r=r), 'pts')
        s.calc('M%d' % r, '=IF(AND(ISNUMBER($J{r}),ISNUMBER($L{r})),$J{r}+$L{r},"")'.format(r=r), 'pts', bold=True)
        s.calc('N%d' % r, '=IF(AND(ISNUMBER($G{r}),ISNUMBER($H{r})),$G{r}+$H{r},"")'.format(r=r), 'money')
        s.calc('O%d' % r, '=IF(AND(ISNUMBER($N{r}),ISNUMBER($M{r}),N($M{r})>0),$N{r}/($M{r}/1000),"")'.format(r=r), 'money', bold=True)
        s.calc('P%d' % r, ('=IF(AND($A{r}="",$B{r}="",$C{r}=""),"",IF(NOT(ISNUMBER($C{r})),"Informe os pontos transferidos.",IF(NOT(ISNUMBER($D{r})),"Informe o fator de conversão (1 = 1 para 1).",'
                           'IF(N($D{r})<=0,"O fator de conversão precisa ser maior que zero.",IF(NOT(ISNUMBER($E{r})),"Informe o percentual de bônus (0% se não houver): campo vazio = não informado.",IF(OR(NOT(ISNUMBER($G{r})),NOT(ISNUMBER($H{r}))),"Informe o custo dos pontos de origem e as taxas e gastos adicionais (digite 0 se não houver).",IF(AND(ISNUMBER($F{r}),$F{r}<$C{r},$E{r}>0),"O limite informado é menor que os pontos transferidos: o bônus foi calculado só sobre o limite.",'
                           'IF(AND(NOT(ISNUMBER($F{r})),$E{r}>0),"Nenhum limite informado: o cálculo assume que todos os pontos transferidos recebem bônus. Confira o regulamento da promoção.",'
                           'IF($E{r}=0,"Bônus de 0%: só a conversão foi considerada.",""))))))))) ').format(r=r), 'long')
        s.inp('Q%d' % r, 'long', ex=e[9]); s.inp('R%d' % r, 'long', ex=e[10]); s.inp('S%d' % r, 'long', ex=e[11]); s.inp('T%d' % r, 'long', ex=e[12])
        s.inp('U%d' % r, 'money', ex=e[13])
        s.calc('V%d' % r, '=IF(AND(ISNUMBER($U{r}),ISNUMBER($O{r})),$U{r}-$O{r},"")'.format(r=r), 'money')
        s.inp('W%d' % r, 'long', ex=None); s.inp('X%d' % r, 'long', ex=e[14]); ws.set_row(r - 1, 62)
    s.validate_num('C8:C22', 0); s.validate_num('G8:H22', 0); s.validate_num('F8:F22', 0); s.validate_num('U8:U22', 0)
    ws.data_validation('D8:D22', {'validate': 'decimal', 'criteria': '>', 'value': 0, 'ignore_blank': True, 'error_title': 'Fator inválido', 'error_message': 'O fator de conversão precisa ser maior que zero.', 'input_title': 'Dica', 'input_message': '1 = 1 para 1. 0,5 = 2 pontos de origem para 1 no destino.'})
    ws.data_validation('E8:E22', {'validate': 'decimal', 'criteria': 'between', 'minimum': 0, 'maximum': 10, 'ignore_blank': True, 'error_title': 'Percentual inválido', 'error_message': 'Digite o bônus como percentual (por exemplo, 30%). Máximo aceito: 1000%.'})
    s.validate_list('I8:I22', items=['Sim', 'Não', 'Não sei'])
    s.cf('P8:P22', '=OR(ISNUMBER(SEARCH("menor que",$P8)),ISNUMBER(SEARCH("Nenhum limite",$P8)))', {'bg_color': ORANGE_L})
    s.cf('P8:P22', '=OR(ISNUMBER(SEARCH("Informe",$P8)),ISNUMBER(SEARCH("precisa",$P8)))', {'bg_color': GRAY_L})
    s.t('A24', 'Custo total = custo dos pontos de origem + taxas e gastos adicionais. Custo por milheiro = custo total ÷ (milhas recebidas ÷ 1.000). A comparação com o valor obtido no resgate é "antes de outros custos". Prazos e validades da base e do bônus podem ser diferentes: registre cada um com a fonte.', italic=True, font_color=GRAY)
    s.finish(header_row=7, freeze=(7, 2), area='A1:X24')
    return s

# ===================================================================== 11
def sheet11(book):
    s = book.sheet(11); ws = s.ws
    widths = {'A': 22, 'B': 12, 'C': 16, 'D': 14, 'E': 14, 'F': 14, 'G': 20, 'H': 30, 'I': 18, 'J': 14, 'K': 16, 'L': 18, 'M': 18, 'N': 18, 'O': 18, 'P': 18, 'Q': 46, 'R': 20, 'S': 34, 'T': 30, 'U': 34}
    for c, w in widths.items(): ws.set_column('%s:%s' % (c, c), w)
    s.title('Renovar pontos: extensão ou reativação', 'Capítulo 10 — Compare desembolsos, na mesma unidade (reais), antes de pagar',
            'Não compare um custo total com um valor por milheiro. Estar com pontos vencendo não decide nada sozinho: avalie se existe um uso concreto e disponível e quanto você realmente desembolsaria nos dois caminhos.')
    s.back_link('A4')
    hdr = ['Programa', 'Lote (identificador)', 'Tipo', 'Quantidade elegível (pontos)', 'Custo total da extensão ou reativação (R$)', 'Nova validade oferecida', 'Prazo de crédito da reativação (texto)', 'Uso concreto encontrado (texto)',
           'Custo da alternativa em dinheiro, para o mesmo uso (R$)', 'Pontos necessários no resgate', 'Pagamento adicional no resgate (R$)', 'Outros gastos necessários para completar o resgate (R$)', 'Outros pontos necessários (de outro lote)',
           'Desembolso se renovar e resgatar (R$)', 'Desembolso se pagar a alternativa em dinheiro (R$)', 'Diferença: dinheiro − renovar (R$)', 'Verificações', 'Decisão', 'Justificativa', 'Fonte e data', 'Observações']
    s.head(6, hdr, height=78, calc_cols=(13, 14, 15, 16))
    ex = {7: ('Programa de exemplo', 'R1', 'Reativação', 10000, 300, None, None, 'Nenhum uso concreto encontrado', None, None, None, None, None, 'Deixar vencer', 'Sem uso à vista, pagar R$ 300 significaria arriscar perder as milhas de novo.', None, '[Exemplo do livro, Cap. 10] Rafael: 10.000 milhas vencidas; reativação por R$ 300; as milhas reativadas valeriam por 12 meses. Ele não tem viagem planejada nesse período.'),
          8: ('Programa de exemplo (adicional fictício)', 'R2', 'Extensão', 8000, 150, D(2027, 6, 30), None, 'Passagem que você já encontrou (fictícia)', 600, 8000, 80, 0, None, 'Ainda não decidi', None, None, '[Exemplo adicional fictício] Mostra a comparação de desembolsos: R$ 150 + R$ 80 contra R$ 600 em dinheiro. Não é recomendação; confirme a validade do bilhete, a disponibilidade e as regras do programa.')}
    for r in range(7, 22):
        e = ex.get(r) if book.ex else None
        e = e or (None,) * 17
        s.inp('A%d' % r, 'text', ex=e[0]); s.inp('B%d' % r, 'text', ex=e[1]); s.inp('C%d' % r, 'text', ex=e[2]); s.inp('D%d' % r, 'pts', ex=e[3]); s.inp('E%d' % r, 'money', ex=e[4]); s.inp('F%d' % r, 'date', ex=e[5]); s.inp('G%d' % r, 'text', ex=e[6])
        s.inp('H%d' % r, 'long', ex=e[7]); s.inp('I%d' % r, 'money', ex=e[8]); s.inp('J%d' % r, 'pts', ex=e[9]); s.inp('K%d' % r, 'money', ex=e[10]); s.inp('L%d' % r, 'money', ex=e[11]); s.inp('M%d' % r, 'pts', ex=e[12])
        s.calc('N%d' % r, '=IF(AND(ISNUMBER($E{r}),ISNUMBER($I{r}),ISNUMBER($K{r}),ISNUMBER($L{r}),N($J{r})>0),$E{r}+$K{r}+$L{r},"")'.format(r=r), 'money', bold=True)
        s.calc('O%d' % r, '=IF(ISNUMBER($N{r}),$I{r},"")'.format(r=r), 'money', bold=True)
        s.calc('P%d' % r, '=IF(AND(ISNUMBER($N{r}),ISNUMBER($O{r})),$O{r}-$N{r},"")'.format(r=r), 'money', bold=True)
        s.calc('Q%d' % r, ('=IF(AND($A{r}="",$E{r}=""),"",IF(NOT(ISNUMBER($I{r})),"Sem uso concreto com valor em dinheiro informado: não há o que comparar. Pagar sem perspectiva de uso pode gerar outra perda.",'
                           'IF(OR(NOT(ISNUMBER($E{r})),N($J{r})<=0),"Preencha o custo da renovação e os pontos necessários no resgate.",'
                           'IF(OR(NOT(ISNUMBER($K{r})),NOT(ISNUMBER($L{r}))),"Campo vazio = não informado, não zero: preencha o pagamento adicional e os outros gastos do resgate (digite 0 se não houver).",'
                           'IF(AND(ISNUMBER($D{r}),$J{r}>$D{r}),"Os pontos necessários superam a quantidade elegível: o resgate não cabe só neste lote.",'
                           'IF(AND(ISNUMBER($F{r}),$F{r}<=DataRef),"A nova validade informada não é posterior à data de referência.","Comparação apenas de desembolsos. Não considera o valor dos pontos, a disponibilidade nem a validade do bilhete."))))))').format(r=r), 'long')
        s.inp('R%d' % r, 'text', ex=e[13]); s.inp('S%d' % r, 'long', ex=e[14]); s.inp('T%d' % r, 'long', ex=e[15]); s.inp('U%d' % r, 'long', ex=e[16]); ws.set_row(r - 1, 62)
    s.validate_list('C7:C21', items=['Extensão', 'Reativação'])
    s.validate_list('R7:R21', items=['Estender', 'Reativar', 'Usar sem pagar renovação', 'Deixar vencer', 'Ainda não decidi'])
    s.validate_num('D7:E21', 0); s.validate_num('I7:M21', 0); s.validate_date('F7:F21')
    s.cf('Q7:Q21', '=OR(ISNUMBER(SEARCH("superam",$Q7)),ISNUMBER(SEARCH("Sem uso",$Q7)),ISNUMBER(SEARCH("não é posterior",$Q7)))', {'bg_color': ORANGE_L})
    s.t('A23', 'Como ler: "Diferença" positiva = pagar a alternativa em dinheiro desembolsa mais do que renovar e resgatar, nas condições informadas. Ela não diz que vale a pena: confirme o uso, a disponibilidade, as condições do programa e o que acontece com o bilhete (Capítulos 10, 13 e 14).', italic=True, font_color=GRAY)
    s.finish(header_row=6, freeze=(6, 1), area='A1:U23')
    return s
