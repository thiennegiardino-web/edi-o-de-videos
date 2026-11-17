# Relatório de testes — material complementar em Excel

**Ferramenta:** LibreOffice 24.2.7.2 420(Build:2), em modo headless, com **recálculo total forçado ao abrir** (perfil com `OOXMLRecalcMode=0`). Arquivos gerados com Python 3.13.18, XlsxWriter 3.2.9 e openpyxl 3.1.5. Modificações de teste aplicadas por openpyxl em cópias; o teste de inserção de linhas usou a API UNO do LibreOffice.

**Não foi testado em Microsoft Excel, Excel Online, Google Sheets nem Numbers.** Nenhuma alegação de teste nessas ferramentas é feita.

**Resultado final:** 221 de 221 casos aprovados, 0 reprovados.

## Histórico honesto das execuções

- **Primeira execução da bateria:** 169 de 179 aprovados. Os 10 resultados reprovados eram: (1) **erro real das planilhas**: o alerta da aba 06 (`#VALUE!`) quando faltava a data-limite de reclamação; corrigido com `N()`; (2) um **link externo não intencional** criado pelo XlsxWriter na aba 14 (endereço da ANAC convertido em hiperlink); corrigido (`strings_to_urls: False`); (3) uma célula de legenda desbloqueada na aba 01; corrigida; (4) expectativas de teste erradas (soma de 700, em vez de 800; data de teste de reclamação); (5) limites do próprio teste (separador decimal do LibreOffice em inglês; a API não insere linhas em aba protegida; linha inserida nasce sem as fórmulas das colunas calculadas).
- **Execução final da v1.0:** 188 de 188 aprovados, depois das correções acima e da inclusão das abas adicionais.
- **v1.1 (correções do parecer do Codex):** a bateria foi ampliada para cobrir (grupo G6) o saldo sem vencidos, o alerta de reclamação combinado e o campo vazio ≠ zero. Os testes de campo vazio das abas 07 a 11 e de células com valores digitados foram ajustados ao novo comportamento (zero digitado calcula; vazio não calcula). **Execução final da v1.1: todos aprovados (resultado acima).**

## Observações sobre o método

- Valores esperados dos exemplos vêm do **texto aprovado** (R$ 23,33; R$ 16,81; R$ 55; R$ 34,67; R$ 336; R$ 300; R$ 23,08; R$ 15,00; etc.).
- Casos de borda usam o **arquivo do aluno** com dados de teste inseridos por script (não pela interface).
- Para textos com valores em reais (`FIXED`), o separador decimal segue a localidade de quem abre o arquivo. O teste normaliza vírgula e ponto.
- A comparação "cache × recálculo" confirma que os valores guardados no arquivo de exemplos são os que o LibreOffice recalcula.

## Resultados por grupo

### G1 Exemplos aprovados do livro (arquivo de exemplos) (58 casos, 58 aprovados)

| ID | Caso | Esperado | Obtido | Resultado |
|---|---|---|---|---|
| 1.01 | Cap. 13: tarifa = R$ 1.960 − R$ 280 | 1680 | 1680 | Aprovado |
| 1.02 | Cap. 13: alternativa R$ 1.490 − R$ 280 de taxas | 1210 | 1210 | Aprovado |
| 1.03 | Cap. 13: pago em dinheiro (opção em dinheiro) = R$ 1.680 + R$ 280 + R$ 360 | 2320 | 2320 | Aprovado |
| 1.04 | Cap. 13: pago em dinheiro (milhas) = R$ 280 + R$ 360 | 640 | 640 | Aprovado |
| 1.05 | Cap. 13: economia calculada | 1680 | 1680 | Aprovado |
| 1.06 | Cap. 13: valor por milheiro contra a opção em dinheiro (R$ 23,33) | 23.3333333 | 23.3333333333333 | Aprovado |
| 1.07 | Cap. 13: valor por milheiro contra a alternativa (R$ 16,81) | 16.8055556 | 16.8055555555556 | Aprovado |
| 1.08 | Cap. 4, exemplo A: R$ 55 por milheiro | 55 | 55 | Aprovado |
| 1.09 | Cap. 4, exemplo B: ≈ R$ 34,67 | 34.6666667 | 34.6666666666667 | Aprovado |
| 1.10 | Cap. 4, exemplo C (base A) | 55 | 55 | Aprovado |
| 1.11 | Cap. 4, exemplo C: contra a alternativa de R$ 700 (derivado da mesma fórmula) | 30 | 30 | Aprovado |
| 1.12 | Cap. 13 (Passo 6): total antes da bagagem R$ 1.960 | 1960 | 1960 | Aprovado |
| 1.13 | Cap. 13: 80% da tarifa (R$ 1.344) + taxas devolvidas (R$ 280) = R$ 1.624 | 1624 | 1624 | Aprovado |
| 1.14 | Cap. 13: perda em dinheiro R$ 336 | 336 | 336 | Aprovado |
| 1.15 | Cap. 13: perda em dinheiro com milhas R$ 300 (taxa de cancelamento) | 300 | 300 | Aprovado |
| 1.16 | Cap. 13: nenhuma milha perdida (72.000 devolvidas) | 0 | 0 | Aprovado |
| 1.17 | Cap. 13: taxa de alteração R$ 300 (2 pessoas) | 300 | 300 | Aprovado |
| 1.18 | Cap. 13: taxa de alteração R$ 200 (2 pessoas) | 200 | 200 | Aprovado |
| 1.19 | Cap. 16: R$ 42 por milheiro | 42 | 42 | Aprovado |
| 1.20 | Cap. 17: R$ 600 por 20.000 pontos = R$ 30 por milheiro | 30 | 30 | Aprovado |
| 1.21 | Cap. 17: 20.000 × 1,30 = 26.000 milhas | 26000 | 26000 | Aprovado |
| 1.22 | Cap. 17: R$ 23,08 por milheiro | 23.0769231 | 23.0769230769231 | Aprovado |
| 1.23 | Cap. 17: 20.000 × 2 = 40.000 milhas | 40000 | 40000 | Aprovado |
| 1.24 | Cap. 17: R$ 15,00 por milheiro | 15 | 15 | Aprovado |
| 1.25 | Cap. 17: diferença de cerca de R$ 1,81 contra os R$ 16,81 | 1.81 | 1.81 | Aprovado |
| 1.26 | Cap. 10: saldo = 8.000 + 12.000 (soma dos lotes) | 20000 | 20000 | Aprovado |
| 1.27 | Renovação (adicional): R$ 150 + R$ 80 | 230 | 230 | Aprovado |
| 1.28 | Renovação (adicional): R$ 600 − R$ 230 | 370 | 370 | Aprovado |
| 1.29 | Lote de 16/03/2027: 120 dias após a referência | 120 | 120 | Aprovado |
| 1.30 | Plano: 14 atividades concluídas no exemplo | 14 | 14 | Aprovado |
| 1.31 | Plano: 19 aplicáveis (2 "Não se aplica") | 19 | 19 | Aprovado |
| 1.32 | Aba 04: situação do lote na linha 7 | Mais de 90 dias | Mais de 90 dias | Aprovado |
| 1.33 | Aba 04: situação do lote na linha 9 | Vencido | Vencido | Aprovado |
| 1.34 | Aba 04: situação do lote na linha 10 | Vence hoje | Vence hoje | Aprovado |
| 1.35 | Aba 04: situação do lote na linha 11 | Vence em até 30 dias | Vence em até 30 dias | Aprovado |
| 1.36 | Aba 04: situação do lote na linha 12 | Vence em até 60 dias | Vence em até 60 dias | Aprovado |
| 1.37 | Aba 04: situação do lote na linha 13 | Vence em até 90 dias | Vence em até 90 dias | Aprovado |
| 1.38 | Aba 04: situação do lote na linha 14 | Data não informada | Data não informada | Aprovado |
| 1.39 | Aba 04: situação do lote na linha 15 | Sem expiração confirmada | Sem expiração confirmada | Aprovado |
| 1.40 | Aba 06: 60 dias corridos após 03/11/2026 | 2027-01-02 | 2027-01-02 | Aprovado |
| 1.41 | Aba 06: 30 dias corridos após a data prevista de crédito | 2027-02-01 | 2027-02-01 | Aprovado |
| 1.42 | Aba 06: 10 dias úteis após 12/11/2026 com feriado em 20/11/2026 | 2026-11-27 | 2026-11-27 | Aprovado |
| 1.43 | Aba 06: diferença creditado − esperado | -20 | -20 | Aprovado |
| 1.44 | Aba 06: divergência E prazo de reclamação próximo aparecem juntos (nada fica escondido) | Divergência: creditado diferente do esperado. Data-limite de reclamação em até 7 dias. | Divergência: creditado diferente do esperado. Data-limite de reclamação em até 7 dias. | Aprovado |
| 1.44b | Aba 06: data-limite de reclamação da linha divergente | 2026-11-19 | 2026-11-19 | Aprovado |
| 1.45 | Aba 06: prazo de crédito passou sem crédito | O prazo previsto de crédito passou sem crédito registrado. | O prazo previsto de crédito passou sem crédito registrado. | Aprovado |
| 1.46 | Aba 01: lotes vencidos | 1 | 1 | Aprovado |
| 1.47 | Aba 01: lotes que vencem hoje | 1 | 1 | Aprovado |
| 1.48 | Aba 01: plano | 14 de 19 | 14 de 19 | Aprovado |
| 1.49 | Aba 01: compras divergentes ou em contestação | 1 | 1 | Aprovado |
| 1.50 | Aba 03: alerta de inatividade (prazo atingido) | Prazo informado atingido ou ultrapassado: confira com o programa. | Prazo informado atingido ou ultrapassado: confira com o programa. | Aprovado |
| 1.51 | Aba 03: alerta de inatividade (a menos de 3 meses) | Atenção: a menos de 3 meses do prazo informado (sugestão de organização). | Atenção: a menos de 3 meses do prazo informado (sugestão de organização). | Aprovado |
| 1.51b | Aba 03: o saldo do Programa B exclui os 3.000 pontos já vencidos (1.000 + 1.500 + 0 não vencidos…) | 2500 | 2500 | Aprovado |
| 1.51c | Aba 03: os 3.000 pontos vencidos aparecem na coluna R, fora do saldo | 3000 | 3000 | Aprovado |
| 1.51d | Aba 03: a origem do saldo avisa que há pontos vencidos | Soma dos lotes não vencidos (aba 04) Há pontos vencidos (coluna R). | Soma dos lotes não vencidos (aba 04) Há pontos vencidos (coluna R). | Aprovado |
| 1.52 | Aba 05: diferença de preço do liquidificador (R$ 220 − R$ 200) | 20 | 20 | Aprovado |
| 1.53 | Aba 02: resumo montado (Marina) | Quero viajar de Campinas para Salvador, com 2 pessoa(s), daqui a 10 meses. Já tenho cadastro em Programa da farmácia … | Quero viajar de Campinas para Salvador, com 2 pessoa(s), daqui a 10 meses. Já tenho cadastro em Programa da farmácia … | Aprovado |
| 1.54 | Valores em cache do arquivo de exemplos = valores recalculados pelo LibreOffice (1583 células comparadas) | 0 | 0 | Aprovado |

### G2 Arquivo do aluno (vazio) (7 casos, 7 aprovados)

| ID | Caso | Esperado | Obtido | Resultado |
|---|---|---|---|---|
| 2.01 | Nenhuma fórmula resulta em erro (#DIV/0!, #VALUE!, #REF!, #NAME?, Err:) | 0 | 0 | Aprovado |
| 2.02 | Sem dados, as fórmulas ficam em branco (exceto contagens, datas de referência, resumos e textos-modelo); células fora… | [] | [] | Aprovado |
| 2.03 | Aba 01: data usada nos cálculos = data de hoje (sem data de referência) | 2026-11-17 | 2026-11-17 | Aprovado |
| 2.04 | Aba 01: contagens zeradas | 0 | 0 | Aprovado |
| 2.05 | Aba 01: checklists "0 de 49" | 0 de 49 | 0 de 49 | Aprovado |
| 2.06 | Aba 02: resumo montado com sublinhados | Quero viajar de __________ para __________, com ____ pessoa(s), daqui a ____ meses. Já tenho cadastro em __________. … | Quero viajar de __________ para __________, com ____ pessoa(s), daqui a ____ meses. Já tenho cadastro em __________. … | Aprovado |
| 2.07 | Aba 12: plano com 21 atividades, sem concluídas | 0 de 21 | 0 de 21 | Aprovado |

### G3 Entradas inválidas, vazias e casos de borda (arquivo do aluno com dados de teste) (98 casos, 98 aprovados)

| ID | Caso | Esperado | Obtido | Resultado |
|---|---|---|---|---|
| 3.01 | Aba 07 preenchida no arquivo do aluno reproduz R$ 23,33 | 23.3333333 | 23.3333333333333 | Aprovado |
| 3.02 | … e R$ 16,81 contra a alternativa | 16.8055556 | 16.8055555555556 | Aprovado |
| 3.03 | Pontos = 0: sem divisão por zero (valor por milheiro em branco) |  |  | Aprovado |
| 3.04 | … e aviso para informar os pontos | Informe os pontos ou milhas utilizados (maior que zero) para calcular o valor por milheiro. | Informe os pontos ou milhas utilizados (maior que zero) para calcular o valor por milheiro. | Aprovado |
| 3.05 | Texto no lugar dos pontos: sem erro; resultado em branco |  |  | Aprovado |
| 3.06 | … e aviso | Informe os pontos ou milhas utilizados (maior que zero) para calcular o valor por milheiro. | Informe os pontos ou milhas utilizados (maior que zero) para calcular o valor por milheiro. | Aprovado |
| 3.07 | Taxa de embarque maior que o preço total: aviso para não contar a taxa duas vezes | A taxa de embarque informada é maior que o preço total em dinheiro: confira para não contar a taxa duas vezes. | A taxa de embarque informada é maior que o preço total em dinheiro: confira para não contar a taxa duas vezes. | Aprovado |
| 3.08 | Sem alternativa em dinheiro: resultado contra a alternativa em branco |  |  | Aprovado |
| 3.09 | … e aviso | Informe a melhor alternativa em dinheiro para ver o resultado contra ela (Capítulo 12). | Informe a melhor alternativa em dinheiro para ver o resultado contra ela (Capítulo 12). | Aprovado |
| 3.10 | Sem os pagamentos adicionais do resgate: não calcula (não assume zero) |  |  | Aprovado |
| 3.11 | … e aviso: campo vazio = não informado | Campo vazio = não informado, não zero: preencha as taxas de embarque, os outros pagamentos e a bagagem do resgate (di… | Campo vazio = não informado, não zero: preencha as taxas de embarque, os outros pagamentos e a bagagem do resgate (di… | Aprovado |
| 3.12 | Exercício do Cap. 4 com pontos = 0: sem erro |  |  | Aprovado |
| 3.13 | Exercício do Cap. 4 com o pagamento adicional em branco: NÃO assume zero (resultado em branco) |  |  | Aprovado |
| 3.13b | Exercício do Cap. 4 com pagamento adicional digitado como 0: R$ 100 por milheiro | 100 | 100 | Aprovado |
| 3.14 | Transferência sem limite: 26.000 milhas | 26000 | 26000 | Aprovado |
| 3.15 | … custo por milheiro R$ 23,08 | 23.0769231 | 23.0769230769231 | Aprovado |
| 3.16 | … e aviso de que nenhum limite foi informado | Nenhum limite informado: o cálculo assume que todos os pontos transferidos recebem bônus. Confira o regulamento da pr… | Nenhum limite informado: o cálculo assume que todos os pontos transferidos recebem bônus. Confira o regulamento da pr… | Aprovado |
| 3.17 | Limite de bônus de 10.000 pontos: bônus = 3.000 | 3000 | 3000 | Aprovado |
| 3.18 | … total recebido 23.000 | 23000 | 23000 | Aprovado |
| 3.19 | … custo por milheiro R$ 26,09 | 26.0869565 | 26.0869565217391 | Aprovado |
| 3.20 | … aviso do limite | O limite informado é menor que os pontos transferidos: o bônus foi calculado só sobre o limite. | O limite informado é menor que os pontos transferidos: o bônus foi calculado só sobre o limite. | Aprovado |
| 3.21 | Limite maior que os pontos: bônus integral (6.000) | 6000 | 6000 | Aprovado |
| 3.22 | Fator de conversão 0: sem divisão por zero |  |  | Aprovado |
| 3.23 | … aviso | O fator de conversão precisa ser maior que zero. | O fator de conversão precisa ser maior que zero. | Aprovado |
| 3.24 | Fator 0,5 (2 pontos de origem para 1): base = 10.000 | 10000 | 10000 | Aprovado |
| 3.25 | … bônus de 30% sobre a base = 3.000 | 3000 | 3000 | Aprovado |
| 3.26 | Bônus 0%: apenas a conversão | 20000 | 20000 | Aprovado |
| 3.27 | … aviso | Bônus de 0%: só a conversão foi considerada. | Bônus de 0%: só a conversão foi considerada. | Aprovado |
| 3.28 | Bônus em texto ou vazio: NÃO é zero; sem erro e sem total |  |  | Aprovado |
| 3.28b | … aviso | Informe o percentual de bônus (0% se não houver): campo vazio = não informado. | Informe o percentual de bônus (0% se não houver): campo vazio = não informado. | Aprovado |
| 3.29 | Lote com vencimento em -1 dia(s) da data de referência | Vencido | Vencido | Aprovado |
| 3.30 | Lote com vencimento em +0 dia(s) da data de referência | Vence hoje | Vence hoje | Aprovado |
| 3.31 | Lote com vencimento em +1 dia(s) da data de referência | Vence em até 30 dias | Vence em até 30 dias | Aprovado |
| 3.32 | Lote com vencimento em +30 dia(s) da data de referência | Vence em até 30 dias | Vence em até 30 dias | Aprovado |
| 3.33 | Lote com vencimento em +31 dia(s) da data de referência | Vence em até 60 dias | Vence em até 60 dias | Aprovado |
| 3.34 | Lote com vencimento em +60 dia(s) da data de referência | Vence em até 60 dias | Vence em até 60 dias | Aprovado |
| 3.35 | Lote com vencimento em +61 dia(s) da data de referência | Vence em até 90 dias | Vence em até 90 dias | Aprovado |
| 3.36 | Lote com vencimento em +90 dia(s) da data de referência | Vence em até 90 dias | Vence em até 90 dias | Aprovado |
| 3.37 | Lote com vencimento em +91 dia(s) da data de referência | Mais de 90 dias | Mais de 90 dias | Aprovado |
| 3.38 | Data em texto: "Data inválida" | Data inválida | Data inválida | Aprovado |
| 3.39 | Sem data informada (não é o mesmo que sem expiração) | Data não informada | Data não informada | Aprovado |
| 3.40 | Sem expiração confirmada | Sem expiração confirmada | Sem expiração confirmada | Aprovado |
| 3.41 | Sem expiração confirmada, mas com data preenchida: pede conferência | Conferir: sem expiração e data preenchidas | Conferir: sem expiração e data preenchidas | Aprovado |
| 3.42 | Linha vazia: situação em branco |  |  | Aprovado |
| 3.43 | Dias até o vencimento de lote sem expiração: em branco |  |  | Aprovado |
| 3.44 | Conta com lotes: saldo = soma dos lotes NÃO vencidos (13 lotes × 100 = 1.300, menos 1 vencido) | 1200 | 1200 | Aprovado |
| 3.44b | … e o lote vencido aparece na coluna R | 100 | 100 | Aprovado |
| 3.45 | … o saldo manual divergente é ignorado e isso é informado | Soma dos lotes não vencidos (o saldo manual foi ignorado) Há pontos vencidos (coluna R). | Soma dos lotes não vencidos (o saldo manual foi ignorado) Há pontos vencidos (coluna R). | Aprovado |
| 3.46 | Conta sem lotes: vale o saldo manual | 500 | 500 | Aprovado |
| 3.47 | … e a origem do saldo é "Informado manualmente" | Informado manualmente | Informado manualmente | Aprovado |
| 3.48 | Pontos a vencer em até 90 dias (7 lotes de 100, de +0 a +90 dias; sem expiração e vencidos ficam de fora) | 700 | 700 | Aprovado |
| 3.49 | Próximo vencimento = a própria data de referência (lote que vence hoje) | 2026-11-16 | 2026-11-16 | Aprovado |
| 3.50 | Inatividade: 25 meses sem movimentação, prazo 24 | Prazo informado atingido ou ultrapassado: confira com o programa. | Prazo informado atingido ou ultrapassado: confira com o programa. | Aprovado |
| 3.51 | Inatividade: 23 meses, prazo 24 (menos de 3 meses do limite) | Atenção: a menos de 3 meses do prazo informado (sugestão de organização). | Atenção: a menos de 3 meses do prazo informado (sugestão de organização). | Aprovado |
| 3.52 | Inatividade: 6 meses, prazo 24 | Dentro do prazo informado. | Dentro do prazo informado. | Aprovado |
| 3.53 | Último acúmulo ou resgate no futuro: meses em branco |  |  | Aprovado |
| 3.54 | Só o último acesso foi informado (recente): o alerta pede a data do último acúmulo ou resgate | Informe a data do último acúmulo ou resgate. | Informe a data do último acúmulo ou resgate. | Aprovado |
| 3.55 | Sem prazo de inatividade informado: nenhum alerta é inventado |  |  | Aprovado |
| 3.56 | 10 dias úteis após quinta 12/11/2026, sem feriados | 2026-11-26 | 2026-11-26 | Aprovado |
| 3.57 | 10 dias corridos após 12/11/2026 | 2026-11-22 | 2026-11-22 | Aprovado |
| 3.58 | Sem prazo informado: sem data prevista (não assume prazo) |  |  | Aprovado |
| 3.59 | Sem tipo de dias: sem data prevista |  |  | Aprovado |
| 3.60 | Prazo para reclamar com base "Outra data" sem a data: não calcula |  |  | Aprovado |
| 3.61 | Prazo para reclamar: 5 dias úteis após 01/12/2026 (terça) = 08/12/2026 | 2026-12-08 | 2026-12-08 | Aprovado |
| 3.62 | Data-limite de reclamação já passou E prazo de crédito passou (aguardando): os dois avisos | Data-limite de reclamação já passou. O prazo previsto de crédito passou sem crédito registrado. | Data-limite de reclamação já passou. O prazo previsto de crédito passou sem crédito registrado. | Aprovado |
| 3.63 | Linha com pouco dado: pede o preenchimento (sem inventar datas) | Preencha o evento, a data, o prazo e o tipo de dias para calcular as datas. | Preencha o evento, a data, o prazo e o tipo de dias para calcular as datas. | Aprovado |
| 3.64 | 10 dias úteis após 12/11/2026 com feriado em 20/11/2026 na lista configurável | 2026-11-27 | 2026-11-27 | Aprovado |
| 3.65 | Clube: período menor que a permanência mínima gera alerta | O período analisado é menor que a permanência mínima: analise o período completo do compromisso. | O período analisado é menor que a permanência mínima: analise o período completo do compromisso. | Aprovado |
| 3.66 | Clube com 0 mês analisado: sem divisão por zero |  |  | Aprovado |
| 3.67 | Clube: custo total em 12 meses = 42 × 12 + 60 de outros custos | 564 | 564 | Aprovado |
| 3.68 | Clube: pontos recebidos = 1.000 × 12 + 1.000 (bônus contado uma vez no período) | 13000 | 13000 | Aprovado |
| 3.69 | Clube: custo por milheiro no período completo = 564 ÷ 13 | 43.3846154 | 43.3846153846154 | Aprovado |
| 3.70 | Compra com tudo em zero: sem divisão por zero |  |  | Aprovado |
| 3.70b | Clube com outros custos e bônus em branco: NÃO assume zero (sem total) e avisa | Campo vazio = não informado, não zero: preencha outros custos, pontos regulares e bônus (digite 0 se não houver). | Campo vazio = não informado, não zero: preencha outros custos, pontos regulares e bônus (digite 0 se não houver). | Aprovado |
| 3.70c | … sem custo total |  |  | Aprovado |
| 3.70d | Compra com o bônus em branco: total recebido não é calculado |  |  | Aprovado |
| 3.71 | Renovação: pontos necessários maiores que a quantidade elegível: aviso | Os pontos necessários superam a quantidade elegível: o resgate não cabe só neste lote. | Os pontos necessários superam a quantidade elegível: o resgate não cabe só neste lote. | Aprovado |
| 3.72 | Renovação: nova validade não posterior à referência: aviso | A nova validade informada não é posterior à data de referência. | A nova validade informada não é posterior à data de referência. | Aprovado |
| 3.73 | Renovação sem a alternativa em dinheiro: não compara e avisa | Sem uso concreto com valor em dinheiro informado: não há o que comparar. Pagar sem perspectiva de uso pode gerar outr… | Sem uso concreto com valor em dinheiro informado: não há o que comparar. Pagar sem perspectiva de uso pode gerar outr… | Aprovado |
| 3.74 | Renovação sem comparação: desembolsos em branco |  |  | Aprovado |
| 3.74b | Renovação com o pagamento adicional em branco: NÃO assume zero (sem desembolso) e avisa | Campo vazio = não informado, não zero: preencha o pagamento adicional e os outros gastos do resgate (digite 0 se não … | Campo vazio = não informado, não zero: preencha o pagamento adicional e os outros gastos do resgate (digite 0 se não … | Aprovado |
| 3.74c | … desembolso em branco |  |  | Aprovado |
| 3.75 | Mapa de acúmulo: preço pelo programa R$ 20 mais caro | 20 | 20 | Aprovado |
| 3.76 | … leitura: não chama os pontos de gratuitos (o separador decimal segue a localidade do programa que abre o arquivo) | Pelo caminho do programa você paga R$ 20.00 a mais. Os pontos não são gratuitos: o benefício precisa valer mais do qu… | Pelo caminho do programa você paga R$ 20.00 a mais. Os pontos não são gratuitos: o benefício precisa valer mais do qu… | Aprovado |
| 3.77 | Mapa de acúmulo: preço pelo programa mais barato | O caminho do programa tem preço final R$ 20.00 menor que a alternativa. | O caminho do programa tem preço final R$ 20.00 menor que a alternativa. | Aprovado |
| 3.78 | Mapa de acúmulo: preços iguais | Mesmo preço final nas duas opções. | Mesmo preço final nas duas opções. | Aprovado |
| 3.79 | Mapa de acúmulo: sem a alternativa, sem diferença |  |  | Aprovado |
| 3.80 | Data de referência informada substitui a data de hoje | 2030-01-15 | 2030-01-15 | Aprovado |
| 3.81 | Plano: semana 1 com 5 de 5 | 5 | 5 | Aprovado |
| 3.82 | Plano: "Não se aplica" sai do total aplicável da semana 2 | 4 | 4 | Aprovado |
| 3.83 | Plano: início sugerido da semana 3 = início + 14 dias | 2026-11-16 | 2026-11-16 | Aprovado |
| 3.84 | Checklists: 5 itens do grupo mensal concluídos | 5 | 5 | Aprovado |
| 3.85 | Checklists: progresso do grupo mensal | 1 | 1 | Aprovado |
| 3.86 | Checklists: total de itens por grupo somam 49 | 49 | 49 | Aprovado |
| 3.87 | Aba 08 preenchida no arquivo do aluno: perda em dinheiro R$ 336 | 336 | 336 | Aprovado |
| 3.88 | … e R$ 300 com milhas | 300 | 300 | Aprovado |
| 3.89 | Bagagem parcialmente devolvida (R$ 100 de R$ 360): a diferença entra na perda (336 + 260) | 596 | 596 | Aprovado |
| 3.90 | Aba 08 sem dados da aba 07: resultados em branco |  |  | Aprovado |

### G4 Inserção de linhas (as faixas das fórmulas acompanham) (7 casos, 7 aprovados)

| ID | Caso | Esperado | Obtido | Resultado |
|---|---|---|---|---|
| 4.01 | Inserir linha na aba 04: a soma de lotes da aba 03 inclui a nova linha (777) | 777 | 777.0 | Aprovado |
| 4.02 | … a faixa da fórmula cresceu de 156 para 157 | [true, true] | [true, true] | Aprovado |
| 4.03 | … copiando as colunas calculadas da linha de cima, a situação do lote novo é calculada (vence em até 30 dias) | Vence em até 30 dias | Vence em até 30 dias | Aprovado |
| 4.04 | … e a contagem da aba 01 acompanha | 1 | 1.0 | Aprovado |
| 4.05 | Última linha de capacidade da aba 04 entra na soma (777 + 23) | 800 | 800.0 | Aprovado |
| 4.06 | Inserir linha na aba 06: a contagem de compras aguardando (aba 01) conta a linha nova | 1 | 1.0 | Aprovado |
| 4.07 | … e a faixa da fórmula passou de 106 para 107 | true | true | Aprovado |

### G6 Correções do parecer do Codex (v1.1): saldo, reclamação e campo vazio ≠ zero (21 casos, 21 aprovados)

| ID | Caso | Esperado | Obtido | Resultado |
|---|---|---|---|---|
| 6.01 | Programa só com lotes vencidos: saldo considerado = 0 (não 300) | 0 | 0 | Aprovado |
| 6.02 | … e os 300 pontos vencidos aparecem na coluna R | 300 | 300 | Aprovado |
| 6.03 | Lote que vence hoje, sem data e sem expiração entram no saldo (50 + 70 + 30) | 150 | 150 | Aprovado |
| 6.04 | … sem pontos vencidos | 0 | 0 | Aprovado |
| 6.05 | Data em texto não é tratada como vencida (10 no saldo) | 10 | 10 | Aprovado |
| 6.06 | Divergência + data-limite vencida: os dois avisos | Divergência: creditado diferente do esperado. Data-limite de reclamação já passou. | Divergência: creditado diferente do esperado. Data-limite de reclamação já passou. | Aprovado |
| 6.07 | Divergência sem prazo de reclamação informado: avisa que não há data-limite | Divergência: creditado diferente do esperado. Sem data-limite de reclamação calculada: informe o prazo, o tipo de dia… | Divergência: creditado diferente do esperado. Sem data-limite de reclamação calculada: informe o prazo, o tipo de dia… | Aprovado |
| 6.08 | Creditado certo: nenhum alerta de reclamação | Acompanhando. | Acompanhando. | Aprovado |
| 6.09 | Creditado, mas com divergência e prazo vencido: os dois avisos | Divergência: creditado diferente do esperado. Data-limite de reclamação já passou. | Divergência: creditado diferente do esperado. Data-limite de reclamação já passou. | Aprovado |
| 6.10 | Divergência com a data-limite ainda longe: só a divergência | Divergência: creditado diferente do esperado. | Divergência: creditado diferente do esperado. | Aprovado |
| 6.11 | Aba 01 conta cada compra com alerta uma vez (4 das 5 têm alerta) | 4 | 4 | Aprovado |
| 6.12 | Aba 08: taxa de cancelamento (milhas) apagada: a perda NÃO cai para R$ 0; fica em branco |  |  | Aprovado |
| 6.13 | … a taxa total de cancelamento também fica em branco |  |  | Aprovado |
| 6.14 | … e o aviso lista o campo pendente | true | true | Aprovado |
| 6.15 | … a perda em dinheiro (opção em dinheiro), que não depende desse campo, continua R$ 336 | 336 | 336 | Aprovado |
| 6.16 | Com todos os campos (zeros digitados), a perda é calculada e o aviso some |  |  | Aprovado |
| 6.17 | Aba 08: "outros custos" (dinheiro) apagado: perda em branco (não R$ 336) |  |  | Aprovado |
| 6.18 | Aba 08: "diferença de tarifa" apagada: custo de alteração em branco (não R$ 300) |  |  | Aprovado |
| 6.19 | Aba 07: bagagem apagada na opção em dinheiro: total em branco (não "sem bagagem") |  |  | Aprovado |
| 6.20 | … e o aviso pede para informar a bagagem | Informe a bagagem e os serviços das opções em dinheiro (digite 0 se não houver) para calcular os totais. | Informe a bagagem e os serviços das opções em dinheiro (digite 0 se não houver) para calcular os totais. | Aprovado |
| 6.21 | Aba 07: "outros pagamentos adicionais" apagado: total do resgate em branco |  |  | Aprovado |

### G5 Estrutura, proteção e conformidade (30 casos, 30 aprovados)

| ID | Caso | Esperado | Obtido | Resultado |
|---|---|---|---|---|
| 5.01 | Quatorze abas, na ordem do plano, com nomes até 31 caracteres | ["01 Comece aqui", "02 Diagnóstico", "03 Contas", "04 Lotes", "05 Acúmulo", "06 Compras", "07 Comparar", "08 Alterar"… | ["01 Comece aqui", "02 Diagnóstico", "03 Contas", "04 Lotes", "05 Acúmulo", "06 Compras", "07 Comparar", "08 Alterar"… | Aprovado |
| 5.02 | Nenhuma célula mesclada (aluno e exemplos) | 0 | 0 | Aprovado |
| 5.03 | Todas as abas protegidas, sem senha | true | true | Aprovado |
| 5.04 | Todas as células com fórmula estão bloqueadas (1454 fórmulas) | 0 | 0 | Aprovado |
| 5.05 | Há células editáveis (desbloqueadas) para o preenchimento (5715 células) | true | true | Aprovado |
| 5.06 | As duas versões têm exatamente as mesmas fórmulas (1454) | true | true | Aprovado |
| 5.07 | Versão do aluno: nenhum campo editável preenchido (sem dados fictícios) | [] | [] | Aprovado |
| 5.08 | Versão do aluno: nenhum texto "[Exemplo" nem "fictício" fora de instruções | 0 | 0 | Aprovado |
| 5.09 Milhas-do-zero-Plani | Sem macros (vbaProject) nem conexões externas (externalLinks, connections) | 0 | 0 | Aprovado |
| 5.10 Milhas-do-zero-Plani | Hiperlinks apenas internos (nenhum "TargetMode=External") | 0 | 0 | Aprovado |
| 5.11 Milhas-do-zero-Plani | Intervalos nomeados DataRef e Feriados presentes | true | true | Aprovado |
| 5.09 Milhas-do-zero-Exemp | Sem macros (vbaProject) nem conexões externas (externalLinks, connections) | 0 | 0 | Aprovado |
| 5.10 Milhas-do-zero-Exemp | Hiperlinks apenas internos (nenhum "TargetMode=External") | 0 | 0 | Aprovado |
| 5.11 Milhas-do-zero-Exemp | Intervalos nomeados DataRef e Feriados presentes | true | true | Aprovado |
| 5.12 | Checklists (aba 13) + plano (aba 12) = 70 itens, na ordem e com o texto do livro | 70 | 70 | Aprovado |
| 5.13 | … textos idênticos aos do texto aprovado (sem omissão nem duplicação) | true | true | Aprovado |
| 5.14 | Tabela do livro representada: Cap. 1: Programa \| Saldo \| Validade | true | true | Aprovado |
| 5.15 | Tabela do livro representada: Cap. 6: organização das contas | true | true | Aprovado |
| 5.16 | Tabela do livro representada: Cap. 7: controle de compras | true | true | Aprovado |
| 5.17 | Tabela do livro representada: Cap. 7: mapa de acúmulo | true | true | Aprovado |
| 5.18 | Tabela do livro representada: Cap. 10: vencimento, inatividade e extensão/reativação | true | true | Aprovado |
| 5.19 | Tabela do livro representada: Cap. 10: validade da Smiles por origem | true | true | Aprovado |
| 5.20 | Tabela do livro representada: Cap. 13: comparação dinheiro × milhas | true | true | Aprovado |
| 5.21 | Tabela do livro representada: Cap. 13: alteração, cancelamento e bagagem | true | true | Aprovado |
| 5.22 | Cap. 13 (Passo 6): as condições da versão de exemplos são idênticas ao texto do livro | true | true | Aprovado |
| 5.23 | Cap. 10: as cinco linhas da Smiles correspondem ao texto aprovado (origem) | ["Milhas acumuladas em voos da GOL", "Milhas bônus de promoções de transferência de pontos do cartão de crédito", "Mi… | ["Milhas acumuladas em voos da GOL", "Milhas bônus de promoções de transferência de pontos do cartão de crédito", "Mi… | Aprovado |
| 5.24 | Aba 14: nenhuma linha da Smiles além das cinco do texto aprovado | null | null | Aprovado |
| 5.25 | Aba 07: as sete perguntas do Capítulo 12 reproduzidas, na ordem | ["Quanto valor eu obtenho?", "Qual é a melhor alternativa em dinheiro?", "Quanto custou acumular essas milhas?", "Exi… | ["Quanto valor eu obtenho?", "Qual é a melhor alternativa em dinheiro?", "Quanto custou acumular essas milhas?", "Exi… | Aprovado |
| 5.26 | Aba 08: registro de contatos e protocolos (Capítulo 15) presente | true | true | Aprovado |
| 5.27 | Aba 07: decisão do Capítulo 13/18 presente, com lista de opções | Esperar | Esperar | Aprovado |

## Itens que o teste automatizado NÃO cobre (pendências)

1. **Microsoft Excel (qualquer versão) e Google Sheets:** não testados. Recomenda-se abrir os dois arquivos no Excel e conferir: recálculo, listas suspensas, formatação condicional, proteção e `AGGREGATE`/`WORKDAY`.
2. **Inspeção visual interativa:** as prévias foram geradas pelo LibreOffice com fontes substitutas; o aspecto no Excel pode variar (largura de coluna e quebra de linha).
3. **Teste com o aluno (usabilidade):** não realizado.
4. **Reconferência das fontes** (aba 14 e textos citados): deve ser feita perto da publicação.
5. **Feriados:** a lista é só para o aluno configurar; nenhum calendário de feriados é fornecido.
6. **Fusos e horários:** o vencimento "às 23h59 de Brasília" (Livelo) não é modelado; as datas são tratadas por dia.
