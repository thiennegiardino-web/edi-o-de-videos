# -*- coding: utf-8 -*-
"""Gera os documentos do pacote (Markdown) e monta a pasta de entrega."""
import os, sys, json, hashlib, shutil, subprocess, datetime, re, platform
sys.path.insert(0, os.path.dirname(__file__))
from common import *
import openpyxl, xlsxwriter

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
SAIDA = os.path.join(BASE, 'saida'); PAC = os.path.join(BASE, 'pacote')
ALUNO = 'Milhas-do-zero-Planilha-do-aluno.xlsx'; EXEM = 'Milhas-do-zero-Exemplos-preenchidos.xlsx'
res = json.load(open(os.path.join(BASE, 'testes', 'resultado-testes.json'), encoding='utf-8'))

if os.path.exists(PAC): shutil.rmtree(PAC)
os.makedirs(PAC); os.makedirs(os.path.join(PAC, 'previas'))
shutil.copy(os.path.join(SAIDA, ALUNO), PAC); shutil.copy(os.path.join(SAIDA, EXEM), PAC)
for f in sorted(os.listdir(os.path.join(BASE, 'previas'))): shutil.copy(os.path.join(BASE, 'previas', f), os.path.join(PAC, 'previas', f))

def w(nome, txt): open(os.path.join(PAC, nome), 'w', encoding='utf-8').write(txt)
git = lambda *a: subprocess.check_output(['git'] + list(a), cwd=BASE, text=True).strip()
base_commit = '4ec2cc0426b7bbd14c7d1696ad3348a60aa5505c'
try: base_commit = git('rev-parse', 'ac182ad')
except Exception: pass
branch = git('rev-parse', '--abbrev-ref', 'HEAD')
sha = lambda p: hashlib.sha256(open(p, 'rb').read()).hexdigest()
lo = subprocess.check_output(['soffice', '--version'], text=True).strip().split('\n')[0]
items = parse_checklists()

# ------------------------------------------------------------------ LEIA-ME
tabs = [(2, 'Capítulo 1', 'Diagnóstico inicial, programas com saldo e validade, resumo do ponto de partida.'),
        (3, 'Capítulos 6 e 9', 'Uma linha por programa; saldo vem dos lotes; último acesso separado de último acúmulo ou resgate.'),
        (4, 'Capítulos 10 e 11', 'Uma linha por lote; alertas de vencimento (30, 60 e 90 dias).'),
        (5, 'Capítulo 7', 'Mapa de acúmulo e comparação de preço final.'),
        (6, 'Capítulo 7', 'Controle de compras: datas, prazos, divergência, reclamação.'),
        (7, 'Capítulos 4, 12 e 13', 'Calculadora dinheiro × pontos, exercício do Capítulo 4, perguntas do Capítulo 12 e decisão.'),
        (8, 'Capítulos 13 a 15', 'Alteração, cancelamento e bagagem; registro de contatos e protocolos.'),
        (9, 'Capítulos 16 e 17', 'Custo por milheiro de clubes e de compra de pontos.'),
        (10, 'Capítulo 17', 'Transferência bonificada com fator de conversão e limite de bônus.'),
        (11, 'Capítulo 10', 'Extensão ou reativação: comparação de desembolsos.'),
        (12, 'Capítulo 18', 'As 21 atividades do plano de 30 dias, com progresso.'),
        (13, 'Capítulos 11, 14, 15, 17 e 19', 'Os 49 itens dos checklists, com grupos e datas de revisão; extra com perguntas ao emissor.'),
        (14, 'Capítulo 10 e fontes', 'Tabelas explicativas, exemplos datados da Smiles e fontes do material.')]
leia = f"""# Material complementar em Excel — Milhas do zero (De Malas e Milhas)

**Para revisão pelo Codex.** Pacote gerado em {datetime.date.today().strftime('%d/%m/%Y')}.
**Base de conteúdo:** e-book 1, versão única v3 (6 partes, 20 capítulos), pacote de referência do commit `ac182ad` (`entrega-diagramacao/ebook-1-milhas-do-zero/01-texto-consolidado.md`). O pacote de diagramação foi preservado; este material está em `material-complementar-excel/`.

## O que há neste pacote
| Arquivo | Para que serve |
|---|---|
| `Milhas-do-zero-Planilha-do-aluno.xlsx` | Versão **limpa** para o aluno preencher. Nenhum campo editável vem preenchido. |
| `Milhas-do-zero-Exemplos-preenchidos.xlsx` | **Mesma estrutura e mesmas fórmulas**, preenchida com exemplos do livro e com exemplos adicionais, todos identificados como fictícios. |
| `Mapa-de-correspondencia.md` | Cada tabela, atividade e checklist do e-book: capítulo, trecho, aba e célula. |
| `Relatorio-de-testes.md` | Casos testados, resultados esperados e obtidos, ferramenta e pendências. |
| `Chamadas-para-o-PDF.md` | Sugestões de chamadas por capítulo, com o nome exato da aba. |
| `Achados-sobre-o-texto.md` | Erros, ambiguidades e sugestões encontrados no texto aprovado (nada foi alterado no texto). |
| `Versoes-e-checksums.md` | Versões e SHA-256. |
| `previas/` | Imagens das abas principais (renderizadas no LibreOffice). |

## Como o aluno usa (resumo)
1. Abra a aba **01 Comece aqui** e siga a ordem das abas, junto com os capítulos indicados.
2. Preencha só as células **claras com borda** (campos). As células **rosadas** são calculadas e protegidas.
3. Registre a **fonte e a data da consulta** de cada regra de programa. A planilha **não** infere validade nem prazo pelo nome do programa.
4. Uma vez por mês, volte à aba 01 (resumo) e faça o checklist mensal (aba 13).

## As 14 abas
| Aba (nome exato) | Capítulo(s) | O que faz |
|---|---|---|
""" + '\n'.join('| `%s` | %s | %s |' % (NAMES[n], c, d) for n, c, d in [(1, 'Todos', 'Instruções, legenda, navegação, avisos, configurações (data de referência) e resumo simples.')] + tabs) + f"""

## Legenda
- **Campo para preencher:** fundo marfim claro (`{INPUT_BG}`), borda cobre clara; desbloqueado.
- **Campo calculado:** fundo rosé (`{CALC_BG}`); bloqueado; resultados e avisos.
- **Cabeçalho:** cobre escuro (`{COPPER}`); cabeçalho de coluna calculada em cobre mais claro (`{COPPER2}`).
- **Alertas de lote (aba 04):** vermelho = vencido; laranja = vence hoje ou em até 30 dias; amarelo = até 60; amarelo claro = até 90; cinza = data não informada; verde = sem expiração confirmada. Os intervalos são **sugestões de organização**, não prazos oficiais. Todo alerta também aparece em **texto** (não depende só da cor).
- **Exemplo adicional fictício:** aparece escrito em Observações como "[Exemplo adicional fictício]"; "[Exemplo do livro, Cap. N]" indica números ou situações do e-book.

## Decisões de projeto que o revisor deve conhecer
- **Saldo sem entrada dupla:** na aba 03, o saldo é a **soma dos lotes** da aba 04. O campo "Saldo manual" só vale para programas sem lotes; se houver lotes e um saldo manual divergente, o manual é ignorado e isso é informado.
- **Pontos de programas diferentes nunca são somados.** Contagens (lotes, compras, atividades) aparecem no resumo; não há "total geral de milhas".
- **Datas:** a aba 01 tem uma **data de referência** (opcional). Em branco, as abas usam a data de hoje; na versão de exemplos ela é fixa em 16/11/2026 para os resultados serem reproduzíveis.
- **Dias úteis (aba 06):** usa `WORKDAY` (segunda a sexta), com lista de feriados configurável na coluna Z. A contagem começa no dia seguinte ao evento. Cada programa pode contar de outro jeito: o aluno confirma. A data-limite só é calculada com todos os dados; nenhum prazo é preenchido por padrão e o prazo da Livelo **não** é aplicado a outros programas.
- **Transferência com limite (aba 10):** a fórmula básica do livro é `pontos × conversão × (1 + bônus)`. Com limite de bônus, a aba calcula `pontos × conversão + (min(pontos, limite) × conversão × bônus)` e avisa quando não há limite informado.
- **Comparações em reais:** extensão ou reativação (aba 11) compara **desembolsos** em reais; nunca um custo total com um valor por milheiro.
- **Aba 08:** calcula custo e perda **nas condições digitadas**. Não conclui direitos legais do passageiro.

## Proteção, linhas novas e senha
- As abas estão **protegidas sem senha**. A proteção evita alterações acidentais nas fórmulas; **não protege dados confidenciais**. Não registre senhas, códigos, CPF ou número de cartão.
- Para mais linhas: há folga prevista (30 contas, 150 lotes, 100 compras, 30 mapas de acúmulo, 10 clubes, 10 compras de pontos, 15 transferências, 15 renovações). Se precisar de mais, use **Revisão > Desproteger planilha** (sem senha), **copie as células calculadas da linha de cima** para a linha nova (uma linha inserida nasce sem as fórmulas) e proteja de novo se quiser. As faixas das fórmulas de outras abas acompanham a inserção (testado).
- Os filtros funcionam com a proteção ligada.

## Compatibilidade
- **Testado em:** {lo} (modo headless, recálculo total forçado), geradas com Python {platform.python_version()}, XlsxWriter {xlsxwriter.__version__} e openpyxl {openpyxl.__version__}.
- **Não testado em:** Microsoft Excel, Excel Online, Google Sheets, Numbers. **Nenhum teste nessas ferramentas foi feito, e nada aqui afirma o contrário.**
- **Esperado:** Excel 2010 ou mais recente (usa `AGGREGATE`, de 2010, na aba 03; `WORKDAY`, `IFERROR`, `SUMIFS`, `COUNTIFS`). Não usa `XLOOKUP`, `LET`, `FILTER`, matrizes dinâmicas, macros, suplementos ou conexões externas. Os arquivos pedem recálculo ao abrir (`fullCalcOnLoad`) e trazem valores calculados em cache.
- `FIXED()` (textos com valores em reais) segue a **localidade** do programa; em português do Brasil mostra vírgula decimal.
- Fontes: Calibri e Georgia. As prévias foram renderizadas com substitutas (Carlito e DejaVu Serif).

## Impressão
Formato A4 paisagem, ajustado à largura de 1 página, cabeçalhos repetidos, rodapé com numeração. As tabelas longas (04 e 06) têm área de impressão reduzida (50 e 40 linhas); amplie em **Layout da Página > Área de Impressão** se precisar.

## Limitações conhecidas
- Sem logotipo: nenhum arquivo de logotipo foi fornecido para esta tarefa; a marca aparece como texto "De Malas e Milhas" e na paleta. Se houver o arquivo original, ele pode ser inserido pequeno no cabeçalho da aba 01.
- Os exemplos adicionais e as datas dos exemplos são inventados.
- A aba 14 reproduz **somente** as cinco linhas da Smiles do texto aprovado, como exemplo histórico da edição, **pendente de reconferência** antes da publicação. As fontes devem ser reconferidas perto da publicação, como o pacote de diagramação já registra.
- O material **não foi publicado nem hospedado**; nenhum link de download ou QR code foi criado. O Codex fará a integração ao PDF depois, com o endereço definitivo.
- O PDF **não foi alterado**.

## Reprodução
Os arquivos são gerados por `material-complementar-excel/build/build_planilhas.py` e testados por `material-complementar-excel/testes/testar.py` (ambos na branch `{branch}`).
"""
w('LEIA-ME.md', leia)

# ------------------------------------------------------------------ Mapa de correspondência
def trecho(s, n=90): return (s[:n] + '…') if len(s) > n else s
mapa = ["# Mapa de correspondência — e-book 1 × planilhas Excel\n",
        "Texto de referência: `entrega-diagramacao/ebook-1-milhas-do-zero/01-texto-consolidado.md` (versão única v3; commit `ac182ad`). Localizações valem para as **duas** planilhas (mesma estrutura). Células citadas em \"linha:coluna\" no formato do Excel.\n",
        "## 1. As oito tabelas do livro (conferência explícita)\n",
        "| # | Capítulo | Tabela (trecho do livro) | Aba | Localização | Como foi representada |", "|---|---|---|---|---|---|",
        "| 1 | 1 | \"Programa \\| Saldo \\| Validade\" (pergunta 5) | `02 Diagnóstico` | cabeçalho A29:E29; dados A30:E44 | Programa (A), Saldo (B), Validade (C), mais \"Saldo com vários lotes?\" (D) e aviso calculado (E) que remete à aba 04. Validade não é obrigatória por programa. |",
        "| 2 | 6 | \"Como organizar suas contas\" (Programa \\| Login \\| Saldo \\| Pontos a vencer e data \\| Último acesso \\| Último acúmulo ou resgate) | `03 Contas` | cabeçalho A6:Q6; dados A7:Q36 | Programa A; Login B; **Saldo** = C (manual), D (soma dos lotes), **E (considerado)**; **Pontos a vencer e data** = G (a vencer em até 90 dias) e H (próximo vencimento), calculados na aba 04; **Último acesso** I; **Último acúmulo ou resgate** J; mais prazo de inatividade (K), alerta (L, M), canal oficial (N), link do regulamento (O), data da consulta (P), observações (Q). |",
        "| 3 | 7 | \"Controle de compras\" (Data, loja e nº do pedido \\| Pontuação esperada e comprovante da oferta \\| O que inicia a contagem do prazo \\| Data prevista de crédito e prazo para reclamar \\| Situação) | `06 Compras` | cabeçalho A6:X6; dados A7:X106 | Colunas agrupadas separadas: data da compra A; loja B; nº do pedido (texto) C; pontuação esperada F; comprovante G; evento que inicia a contagem H e data I; prazo J e tipo de dias K; **data prevista de crédito L (calculada)**; prazo para reclamar M, tipo N, base O/P; **data-limite Q (calculada)**; pontos creditados R; data do crédito S; diferença T; situação U; alerta V; protocolo W; observações X. Feriados: Z7:Z46. |",
        "| 4 | 7 | \"Tarefa: o seu mapa de acúmulo\" (O que eu já gasto \\| Onde gasto \\| Esse lugar tem programa de pontos? \\| Qual o passo para pontuar?) | `05 Acúmulo` | cabeçalho A6:N6; dados A7:N36 | As quatro colunas do livro em A a D; mais valor habitual (E), programa de destino (F), meio de pagamento (G), ativação ou cadastro (H), preço pelo programa (I), melhor preço alternativo (J), diferença (K), leitura (L), fonte e data (M), observações (N). |",
        "| 5 | 10 | \"Três coisas diferentes, que não devem ser confundidas\" | `14 Referências` | A7:C10 | Texto do livro (conceito, \"O que é\", exemplo de pergunta), sem alteração. |",
        "| 6 | 10 | Validade da Smiles por origem | `14 Referências` | A15:F20 | As cinco linhas do texto aprovado, com origem, categoria ou condição, validade, fonte, data (16/11/2026) e situação (exemplo histórico, pendente de reconferência). Não alimenta os lotes. |",
        "| 7 | 13 | Passo 3: comparação dinheiro × milhas (Tarifa; Taxas de embarque; Bagagem; Pago em dinheiro) | `07 Comparar` | tabela reproduzida A40:C45; calculadora A15:E25; resultados A28:B32 | Tabela do livro com linhas e colunas iguais, ligada às entradas da calculadora. A coluna \"Milhas\" mostra, na linha \"Tarifa\", os pontos utilizados (como no livro). Composição \"R$ 1.680 + R$ 280 + R$ 360\" na linha 45. |",
        "| 8 | 13 | Passo 6: Alteração (2 pessoas), Cancelamento antes do embarque, Bagagem não usada | `08 Alterar` | A28:C30 (texto das condições); entradas B18:C27; resultados A35:C44 | Texto das condições (idêntico ao livro na versão de exemplos) e cálculo de taxa de alteração, devolução, taxa de cancelamento, bagagem e perda em dinheiro e em pontos. |",
        "\n**Resultado da conferência:** as oito tabelas estão representadas (teste automatizado 5.14 a 5.21 do `Relatorio-de-testes.md`).\n",
        "## 2. Atividades, tarefas e fórmulas\n", "| Capítulo | Trecho do livro | Aba | Localização |", "|---|---|---|---|",
        "| 1 | Perguntas 1 a 6 do diagnóstico (objetivos, detalhes da viagem, frequência, gastos, programas, cartão) | `02 Diagnóstico` | B7:B9; B12:B16; B19; A23:B25; A30:D44; B47 |",
        "| 1 | \"Seu resumo de ponto de partida\" (fill-in com sublinhados) | `02 Diagnóstico` | B50:B56 (resumo montado em B56; \"Meu primeiro passo\" em B55) |",
        "| 4 | Passo a passo e fórmula \"Valor por mil = (preço em dinheiro − o que você ainda paga em dinheiro) ÷ (milhas necessárias ÷ 1.000)\" | `07 Comparar` | B29 e B31 (calculadora); E49:E58 (exercício) |",
        "| 4 | Exemplos fictícios A, B e C | `07 Comparar` (versão de exemplos) | A49:H51 |",
        "| 4 | \"Tarefa: pegue uma passagem… e faça a conta\" | `07 Comparar` | A48:H58 (dez linhas) |",
        "| 6 | \"Para preencher já\": listar programas; anotar saldo e validade; anotar separadamente último acesso e último acúmulo ou resgate; lembrete mensal | `03 Contas`, `04 Lotes`, `13 Checklists` | 03 A7:Q36; 04 A7:N156; 13 (checklist mensal) |",
        "| 7 | \"Controle de compras\" e prazos | `06 Compras` | A7:X106 |",
        "| 7 | Exemplo do liquidificador (R$ 220 × R$ 200) | `05 Acúmulo` (versão de exemplos) | linha 7 (I7:K7) |",
        "| 8 | As sete perguntas ao emissor | `13 Checklists` (seção extra) | A59:F67 (não conta nos 70 itens) |",
        "| 10 | \"Como acompanhar na prática\": uma linha por lote, lembretes, plano de uso, ação (usar, extensão, reativação, deixar vencer) | `04 Lotes` | A7:N156 (colunas J e K) |",
        "| 10 | Exemplo da Marina (8.000 e 12.000 pontos) | `04 Lotes` (exemplos) | linhas 7 e 8 |",
        "| 10 | \"Tarefa: para cada programa… descubra a regra de validade no site oficial\" | `04 Lotes` | colunas L (fonte) e M (data) |",
        "| 10 | \"Estender ou reativar: vale a pena?\" (4 passos) e exemplo do Rafael | `11 Renovar` | A7:U21; exemplo na linha 7 |",
        "| 12 | As sete perguntas \"Dinheiro ou milhas?\" | `07 Comparar` | A61:C68 (adicional) |",
        "| 13 | Passos 1 a 5 (ficha, milhas e adicionais, bagagem, comparação, alternativa) | `07 Comparar` | A5:E45 |",
        "| 13 | Passo 6 (alteração e cancelamento) e perda calculada (R$ 336 e R$ 300) | `08 Alterar` | A6:D44 |",
        "| 13 | Passo 7 (a decisão) e Cap. 18, semana 4 (\"Decidir: resgatar, pagar em dinheiro ou esperar\") | `07 Comparar` | A70:B73 |",
        "| 15 | \"Como pedir ajuda: o que guardar e o que informar\" | `08 Alterar` | A48:D59 (registro de contatos e protocolos; adicional) |",
        "| 16 | \"Tarefa: pegue um clube… calcule o custo por milheiro\" e a fórmula | `09 Clubes e compra` (área A) | A7:Q17 |",
        "| 17 | Compra de pontos: custo por milheiro | `09 Clubes e compra` (área B) | A21:N31 |",
        "| 17 | Transferência bonificada: quatro passos e exemplos de 30% e 100% | `10 Transferências` | A7:X22; exemplos nas linhas 8 e 9 |",
        "| 18 | Plano de 30 dias (quatro semanas) | `12 Plano 30 dias` | A7:I28; progresso B32:E37 |",
        "| 19 | Checklist do iniciante | `13 Checklists` | A35:G56 |",
        "\n## 3. Os 70 itens de checklist\n",
        "Soma conferida: **49 itens na aba 13 + 21 itens na aba 12 = 70**, com texto idêntico ao do texto aprovado, sem omissão nem duplicação (teste 5.12 e 5.13).\n",
        "| # | Capítulo | Trecho do item | Aba | Célula |", "|---|---|---|---|---|"]
row13 = 8; row12 = 8; n = 0
for it in items:
    n += 1
    if it['cap'] == 18: aba, cel = NAMES[12], 'C%d' % row12; row12 += 1
    else: aba, cel = NAMES[13], 'B%d' % row13; row13 += 1
    mapa.append('| %d | %d | %s | `%s` | %s |' % (n, it['cap'], trecho(it['text']).replace('|', '\\|'), aba, cel))
mapa += ["\n## 4. Itens adicionais (além da especificação mínima)\n",
         "- `07 Comparar` seção F (as sete perguntas do Capítulo 12) e seção G (decisão do Capítulo 13, passo 7, e do Capítulo 18).",
         "- `08 Alterar` seção D (registro de contatos e protocolos, Capítulo 15).",
         "- `13 Checklists` seção extra (as sete perguntas ao emissor, Capítulo 8), fora da contagem dos 70 itens.",
         "- `03 Contas`: prazo de inatividade informado pelo programa e alerta (sugestão de organização; não infere o prazo).",
         "- `07 Comparar` seção E: valor contra alternativa de R$ 700 no Exemplo C (derivado da mesma fórmula; o livro não informa o número).",
         "\n## 5. Partes do e-book que não viraram ferramenta (texto explicativo ou lista de orientações)\n",
         "Capítulos 2, 3 e 5 (conceitos e mitos); Capítulo 8 (visão geral, salvo as perguntas ao emissor); Capítulo 9 (boas práticas de segurança, sem campos); Capítulo 11 (os 10 erros comuns, sem campos; o checklist mensal foi convertido); Capítulo 14 (passo a passo geral e exemplo da Livelo, que são texto; a lista de checagem foi convertida); Capítulo 15 (textos de orientação; o checklist rápido foi convertido); Capítulo 20 (glossário e perguntas frequentes); \"Fontes e atualização\".\n"]
w('Mapa-de-correspondencia.md', '\n'.join(mapa))

# ------------------------------------------------------------------ Relatório de testes
from collections import OrderedDict
grupos = OrderedDict()
for r in res: grupos.setdefault(r['grupo'], []).append(r)
tot = len(res); ok = sum(1 for r in res if r['ok'])
def fm(x):
    s = json.dumps(x, ensure_ascii=False, default=str) if not isinstance(x, str) else x
    s = s.replace('|', '\\|').replace('\n', ' ')
    return s if len(s) <= 120 else s[:117] + '…'
rt = [f"# Relatório de testes — material complementar em Excel\n",
      f"**Ferramenta:** {lo}, em modo headless, com **recálculo total forçado ao abrir** (perfil com `OOXMLRecalcMode=0`). Arquivos gerados com Python {platform.python_version()}, XlsxWriter {xlsxwriter.__version__} e openpyxl {openpyxl.__version__}. Modificações de teste aplicadas por openpyxl em cópias; o teste de inserção de linhas usou a API UNO do LibreOffice.\n",
      "**Não foi testado em Microsoft Excel, Excel Online, Google Sheets nem Numbers.** Nenhuma alegação de teste nessas ferramentas é feita.\n",
      f"**Resultado final:** {ok} de {tot} casos aprovados, {tot - ok} reprovados.\n",
      "## Histórico honesto das execuções\n",
      "- **Primeira execução da bateria:** 169 de 179 aprovados. Os 10 resultados reprovados eram: (1) **erro real das planilhas**: o alerta da aba 06 (`#VALUE!`) quando faltava a data-limite de reclamação; corrigido com `N()`; (2) um **link externo não intencional** criado pelo XlsxWriter na aba 14 (endereço da ANAC convertido em hiperlink); corrigido (`strings_to_urls: False`); (3) uma célula de legenda desbloqueada na aba 01; corrigida; (4) expectativas de teste erradas (soma de 700, em vez de 800; data de teste de reclamação); (5) limites do próprio teste (separador decimal do LibreOffice em inglês; a API não insere linhas em aba protegida; linha inserida nasce sem as fórmulas das colunas calculadas).",
      "- **Execução final:** todos aprovados, depois das correções acima e da inclusão das abas adicionais.\n",
      "## Observações sobre o método\n",
      "- Valores esperados dos exemplos vêm do **texto aprovado** (R$ 23,33; R$ 16,81; R$ 55; R$ 34,67; R$ 336; R$ 300; R$ 23,08; R$ 15,00; etc.).",
      "- Casos de borda usam o **arquivo do aluno** com dados de teste inseridos por script (não pela interface).",
      "- Para textos com valores em reais (`FIXED`), o separador decimal segue a localidade de quem abre o arquivo. O teste normaliza vírgula e ponto.",
      "- A comparação \"cache × recálculo\" confirma que os valores guardados no arquivo de exemplos são os que o LibreOffice recalcula.\n",
      "## Resultados por grupo\n"]
for gname, lst in grupos.items():
    rt.append('### %s (%d casos, %d aprovados)\n' % (gname, len(lst), sum(1 for r in lst if r['ok'])))
    rt.append('| ID | Caso | Esperado | Obtido | Resultado |'); rt.append('|---|---|---|---|---|')
    for r in lst: rt.append('| %s | %s | %s | %s | %s |' % (r['id'], fm(r['descricao']), fm(r['esperado']), fm(r['obtido']), 'Aprovado' if r['ok'] else 'REPROVADO'))
    rt.append('')
rt += ["## Itens que o teste automatizado NÃO cobre (pendências)\n",
       "1. **Microsoft Excel (qualquer versão) e Google Sheets:** não testados. Recomenda-se abrir os dois arquivos no Excel e conferir: recálculo, listas suspensas, formatação condicional, proteção e `AGGREGATE`/`WORKDAY`.",
       "2. **Inspeção visual interativa:** as prévias foram geradas pelo LibreOffice com fontes substitutas; o aspecto no Excel pode variar (largura de coluna e quebra de linha).",
       "3. **Teste com o aluno (usabilidade):** não realizado.",
       "4. **Reconferência das fontes** (aba 14 e textos citados): deve ser feita perto da publicação.",
       "5. **Feriados:** a lista é só para o aluno configurar; nenhum calendário de feriados é fornecido.",
       "6. **Fusos e horários:** o vencimento \"às 23h59 de Brasília\" (Livelo) não é modelado; as datas são tratadas por dia.\n"]
w('Relatorio-de-testes.md', '\n'.join(rt))

# ------------------------------------------------------------------ Chamadas para o PDF
ch = [('Capítulo 1', '4 e 5', 2, 'Abra a aba "02 Diagnóstico" na sua planilha de apoio e responda às seis perguntas com os seus dados.'),
      ('Capítulo 4', '9', 7, 'Para fazer a conta do valor por milheiro, abra a aba "07 Comparar" (seção "Exercício rápido do Capítulo 4") na sua planilha de apoio.'),
      ('Capítulo 6', '14', 3, 'Organize suas contas na aba "03 Contas" da sua planilha de apoio. Separe último acesso de último acúmulo ou resgate.'),
      ('Capítulo 7 (mapa de acúmulo)', '16', 5, 'Abra a aba "05 Acúmulo" na sua planilha de apoio e preencha o seu mapa de acúmulo.'),
      ('Capítulo 7 (controle de compras)', '16', 6, 'Controle suas compras pelo caminho de um programa na aba "06 Compras" da sua planilha de apoio.'),
      ('Capítulo 8 (perguntas ao emissor)', '19', 13, 'Anote as respostas do emissor na seção extra da aba "13 Checklists" da sua planilha de apoio.'),
      ('Capítulo 10 (lotes)', '24', 4, 'Registre uma linha por lote na aba "04 Lotes" da sua planilha de apoio.'),
      ('Capítulo 10 (extensão ou reativação)', '24', 11, 'Antes de pagar, compare os desembolsos na aba "11 Renovar" da sua planilha de apoio.'),
      ('Capítulo 10 (tabelas explicativas)', '24', 14, 'As tabelas deste capítulo também estão na aba "14 Referências" da sua planilha de apoio.'),
      ('Capítulo 11 (checklist mensal)', '28', 13, 'Faça o checklist mensal na aba "13 Checklists" da sua planilha de apoio.'),
      ('Capítulo 12', '32', 7, 'Registre as suas respostas às sete perguntas na aba "07 Comparar" (seção F) da sua planilha de apoio.'),
      ('Capítulo 13 (comparação)', '34', 7, 'Simule a sua viagem na aba "07 Comparar" da sua planilha de apoio, com os seus números e a data da sua busca.'),
      ('Capítulo 13 (alterar ou cancelar)', '34 a 37', 8, 'Registre as condições de alteração e cancelamento na aba "08 Alterar" da sua planilha de apoio.'),
      ('Capítulo 14', '38', 13, 'Use a lista "Antes de confirmar o resgate" na aba "13 Checklists" da sua planilha de apoio.'),
      ('Capítulo 15', '41', 8, 'Registre contatos e protocolos na aba "08 Alterar" (seção D) e use o checklist rápido na aba "13 Checklists".'),
      ('Capítulo 16', '45', 9, 'Calcule o custo por milheiro do clube na aba "09 Clubes e compra" (área A) da sua planilha de apoio.'),
      ('Capítulo 17 (compra de pontos)', '48', 9, 'Calcule o custo por milheiro da compra na aba "09 Clubes e compra" (área B) da sua planilha de apoio.'),
      ('Capítulo 17 (transferência)', '48 a 52', 10, 'Calcule a transferência bonificada na aba "10 Transferências" da sua planilha de apoio, respeitando o limite de bônus.'),
      ('Capítulo 18', '53', 12, 'Acompanhe o seu plano de 30 dias na aba "12 Plano 30 dias" da sua planilha de apoio.'),
      ('Capítulo 19', '55', 13, 'Marque o checklist do iniciante na aba "13 Checklists" da sua planilha de apoio.'),
      ('Abertura / final', '2 a 3', 1, 'Material complementar: [link de download a definir]. Comece pela aba "01 Comece aqui" da planilha de apoio.')]
cp = ["# Chamadas para o PDF — sugestões\n",
      "Texto sugerido para o Codex integrar ao PDF **depois**, com o endereço de download definitivo. **Nenhum link, URL de download ou QR code foi criado neste pacote**; onde aparece `[link de download a definir]`, o endereço ainda não existe. **O PDF não foi alterado.**\n",
      "As páginas indicam o **início do capítulo na versão 1 da diagramação** (`Milhas-do-zero-diagramacao-v1.pdf`, 61 páginas) e podem mudar. Os nomes das abas são os **nomes exatos** das abas dos arquivos.\n",
      "| Capítulo | Página na v1 | Aba (nome exato) | Chamada sugerida |", "|---|---|---|---|"]
for c, pg, n_, txt in ch: cp.append('| %s | %s | `%s` | %s |' % (c, pg, NAMES[n_], txt))
cp += ["\n## Observações\n", "- Em todas as chamadas, a expressão \"na sua planilha de apoio\" pressupõe o nome final do arquivo para o aluno (`Milhas-do-zero-Planilha-do-aluno.xlsx`). Se o Codex preferir citar o arquivo, use o nome exato.",
       "- Chamada geral sugerida (página de abertura): \"Baixe a planilha de apoio em [link de download a definir]. Ela traz uma aba para cada atividade do livro.\"",
       "- Chamada para os exemplos: \"A versão com exemplos fictícios mostra cada aba preenchida: use-a só para ver como funciona.\""]
w('Chamadas-para-o-PDF.md', '\n'.join(cp))
print('docs 1/2 ok')

# ------------------------------------------------------------------ Achados sobre o texto
ach = ["# Achados sobre o texto aprovado (e-book 1, versão única v3)\n",
       "Nada foi alterado no texto. Cada achado traz capítulo, trecho, motivo e solução proposta, para decisão da revisão editorial. **Nenhum erro numérico foi encontrado:** todos os exemplos numéricos do livro foram reproduzidos nas planilhas e conferem (ver `Relatorio-de-testes.md`, grupo G1).\n",
       "| # | Capítulo | Trecho | Motivo | Solução proposta |", "|---|---|---|---|---|"]
ach_rows = [
 ('13 (Passo 3)', 'Tabela: linha \"Tarifa (sem as taxas de embarque)\" com \"R$ 1.680\" em Dinheiro e \"72.000 milhas\" em Milhas', 'A linha mistura unidades: na coluna Milhas, a linha \"Tarifa\" mostra a quantidade de milhas, e não um valor em reais. A planilha reproduz como está, mas pode confundir.', 'Renomear a linha para \"Tarifa (sem taxas) / pontos ou milhas utilizados\" ou criar uma linha própria \"Pontos ou milhas utilizados\".'),
 ('4', '\"Valor por mil = (preço em dinheiro − o que você ainda paga em dinheiro) ÷ (milhas necessárias ÷ 1.000)\"', 'O rótulo da fórmula usa \"valor por mil\", e o resto do livro (e os exemplos) usa \"valor por milheiro\".', 'Padronizar o rótulo como \"Valor por milheiro\".'),
 ('7', 'Tabela \"Controle de compras\": \"Situação (aguardando, creditado ou em contestação)\"', 'A planilha oferece quatro situações (inclui \"divergente\", pedida na especificação). O livro lista três.', 'Acrescentar \"divergente\" ao texto, ou tratar divergência como parte de \"em contestação\" e retirar a opção da planilha.'),
 ('6 e 18', '\"Montar a planilha de contas: saldo, uma linha por lote (programa, quantidade, origem e vencimento), último acesso e último acúmulo ou resgate\"', 'O livro descreve uma planilha única; no Excel, contas (aba 03) e lotes (aba 04) são abas separadas, para evitar dois saldos digitados para o mesmo programa.', 'Ajustar a frase para \"montar a planilha de contas e a de lotes\" ou manter e explicar na chamada do PDF.'),
 ('6', 'Tabela \"Como organizar suas contas\": coluna \"Pontos a vencer e data\"', 'Na planilha, essa informação é calculada a partir dos lotes (aba 04). Sem lotes, a coluna fica em branco.', 'Acrescentar ao texto: \"detalhe os lotes para ver os pontos a vencer\".'),
 ('11, 14, 15, 17, 19', 'Checklists: \"Abri cada conta…\", \"Você sabe como e quando…\", \"Tenho um uso concreto\", \"Reli as regras…\"', 'A voz dos itens varia (primeira pessoa e segunda pessoa), inclusive dentro de uma mesma lista (Capítulo 14 e Capítulo 17).', 'Padronizar a voz de todos os itens (por exemplo, primeira pessoa no passado, como \"Conferi…\").'),
 ('17', '\"Milhas = X × conversão × (1 + bônus)\" e \"A promoção tem regras próprias, prazo e muitas vezes limite\"', 'A fórmula não considera o limite de bônus que o próprio texto menciona. A planilha adaptou: base + (min(pontos, limite) × conversão × bônus) e avisa quando não há limite informado.', 'Acrescentar uma frase: \"se houver limite de bônus, aplique o bônus só até o limite\".'),
 ('10', 'Alertas de 90, 60 e 30 dias e \"O vencimento ocorre às 23h59 (horário de Brasília)\" (exemplo da Livelo)', 'A planilha trata o vencimento por data: no dia do vencimento, a situação é \"Vence hoje\". Se o programa vence o lote no fim do dia, o lote ainda é usável nesse dia; outros programas podem contar diferente.', 'Acrescentar ao texto que o aluno deve confirmar a hora e a regra de vencimento no site oficial.'),
 ('18 (semana 2)', '\"Como ponto de partida, considerar concentrar o acompanhamento em poucos programas\"', 'Sem campo equivalente nas planilhas (é uma orientação, não uma tarefa de preenchimento).', 'Nenhuma ação: registrado só para completar o mapa.'),
 ('7', 'Campanha (Esfera, consultada em 12/11/2026)', 'A regra daquela campanha expira (vigência de 09 a 15/11/2026). A planilha só usa o padrão \"entrar pelo botão do parceiro e finalizar na mesma janela\".', 'Reconferir perto da publicação, como já previsto.')]
for k, (c, t_, m_, s_) in enumerate(ach_rows, 1): ach.append('| %d | %s | %s | %s | %s |' % (k, c, t_, m_, s_))
w('Achados-sobre-o-texto.md', '\n'.join(ach) + '\n')

# ------------------------------------------------------------------ Versões e checksums
vc = [f"# Versões e checksums\n",
      f"- **Data:** {datetime.date.today().strftime('%d/%m/%Y')}",
      f"- **Branch:** `{branch}`",
      f"- **Commit base do conteúdo:** `{base_commit}` (pacote de diagramação preservado). Este material entra em **novo commit** (o mais recente da branch), na pasta `material-complementar-excel/`.",
      f"- **Conteúdo usado:** e-book 1, **versão única v3** (6 partes, 20 capítulos). SHA-256 do texto: `{sha(os.path.join(BASE, '..', 'entrega-diagramacao', 'ebook-1-milhas-do-zero', '01-texto-consolidado.md'))}`.",
      f"- **Versão do material complementar:** 1.0 (planilha do aluno e exemplos preenchidos).",
      f"- **Ferramentas:** {lo}; Python {platform.python_version()}; XlsxWriter {xlsxwriter.__version__}; openpyxl {openpyxl.__version__}.",
      f"- **Resultado dos testes:** {ok} de {tot} casos aprovados (LibreOffice). **Não testado em Excel nem em Google Sheets.**\n",
      "## SHA-256 dos arquivos do pacote\n", "O `.zip` não pode listar o próprio checksum; ele é informado na mensagem de entrega. Os arquivos `.xlsx` abaixo são os entregues. O `Versoes-e-checksums.md` não lista o próprio checksum.\n",
      "| Arquivo | SHA-256 | Bytes |", "|---|---|---|"]
for dp, dn, fn in sorted(os.walk(PAC)):
    for f in sorted(fn):
        p = os.path.join(dp, f)
        if f == 'Versoes-e-checksums.md': continue
        vc.append('| `%s` | `%s` | %d |' % (os.path.relpath(p, PAC), sha(p), os.path.getsize(p)))
w('Versoes-e-checksums.md', '\n'.join(vc) + '\n')

# ------------------------------------------------------------------ ZIP
import zipfile
zp = os.path.join(BASE, 'Milhas-do-zero-Material-complementar-Excel.zip')
if os.path.exists(zp): os.remove(zp)
with zipfile.ZipFile(zp, 'w', zipfile.ZIP_DEFLATED) as z:
    for dp, dn, fn in sorted(os.walk(PAC)):
        for f in sorted(fn):
            p = os.path.join(dp, f); z.write(p, os.path.join('Milhas-do-zero-Material-Excel', os.path.relpath(p, PAC)))
print('zip:', zp, os.path.getsize(zp), 'bytes |', sha(zp))
