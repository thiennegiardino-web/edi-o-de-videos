# Mapa de correspondência — e-book 1 × planilhas Excel

Texto de referência: `entrega-diagramacao/ebook-1-milhas-do-zero/01-texto-consolidado.md` (versão única v3; commit `ac182ad`). Localizações valem para as **duas** planilhas (mesma estrutura). Células citadas em "linha:coluna" no formato do Excel.

## 1. As oito tabelas do livro (conferência explícita)

| # | Capítulo | Tabela (trecho do livro) | Aba | Localização | Como foi representada |
|---|---|---|---|---|---|
| 1 | 1 | "Programa \| Saldo \| Validade" (pergunta 5) | `02 Diagnóstico` | cabeçalho A29:E29; dados A30:E44 | Programa (A), Saldo (B), Validade (C), mais "Saldo com vários lotes?" (D) e aviso calculado (E) que remete à aba 04. Validade não é obrigatória por programa. |
| 2 | 6 | "Como organizar suas contas" (Programa \| Login \| Saldo \| Pontos a vencer e data \| Último acesso \| Último acúmulo ou resgate) | `03 Contas` | cabeçalho A6:R6; dados A7:R36 | Programa A; Login B; **Saldo** = C (manual), D (soma dos lotes), **E (considerado, sem pontos vencidos)**; **Pontos a vencer e data** = G (a vencer em até 90 dias) e H (próximo vencimento), calculados na aba 04; **Último acesso** I; **Último acúmulo ou resgate** J; mais prazo de inatividade (K), alerta (L, M), canal oficial (N), link do regulamento (O), data da consulta (P), observações (Q); **pontos vencidos fora do saldo (R)**. |
| 3 | 7 | "Controle de compras" (Data, loja e nº do pedido \| Pontuação esperada e comprovante da oferta \| O que inicia a contagem do prazo \| Data prevista de crédito e prazo para reclamar \| Situação) | `06 Compras` | cabeçalho A6:X6; dados A7:X106 | Colunas agrupadas separadas: data da compra A; loja B; nº do pedido (texto) C; pontuação esperada F; comprovante G; evento que inicia a contagem H e data I; prazo J e tipo de dias K; **data prevista de crédito L (calculada)**; prazo para reclamar M, tipo N, base O/P; **data-limite Q (calculada)**; pontos creditados R; data do crédito S; diferença T; situação U; alerta V; protocolo W; observações X. Feriados: Z7:Z46. |
| 4 | 7 | "Tarefa: o seu mapa de acúmulo" (O que eu já gasto \| Onde gasto \| Esse lugar tem programa de pontos? \| Qual o passo para pontuar?) | `05 Acúmulo` | cabeçalho A6:N6; dados A7:N36 | As quatro colunas do livro em A a D; mais valor habitual (E), programa de destino (F), meio de pagamento (G), ativação ou cadastro (H), preço pelo programa (I), melhor preço alternativo (J), diferença (K), leitura (L), fonte e data (M), observações (N). |
| 5 | 10 | "Três coisas diferentes, que não devem ser confundidas" | `14 Referências` | A7:C10 | Texto do livro (conceito, "O que é", exemplo de pergunta), sem alteração. |
| 6 | 10 | Validade da Smiles por origem | `14 Referências` | A15:F20 | As cinco linhas do texto aprovado, com origem, categoria ou condição, validade, fonte, data (16/11/2026) e situação (exemplo histórico, pendente de reconferência). Não alimenta os lotes. |
| 7 | 13 | Passo 3: comparação dinheiro × milhas (Tarifa em reais; Milhas utilizadas; Taxas de embarque; Bagagem; Pago em dinheiro) | `07 Comparar` | tabela reproduzida A40:C46; calculadora A15:E25; resultados A28:B32 | Tabela ligada às entradas da calculadora. Na v1.2, a tarifa em reais (linha 41; coluna Milhas: "Paga com milhas") e a quantidade de milhas (linha 46) ficam em linhas separadas, como na versão 4 do texto do livro (a v4 coloca "Milhas utilizadas" logo abaixo da tarifa; a planilha a mantém abaixo da composição para não mover células). Composição "R$ 1.680 + R$ 280 + R$ 360" na linha 45. |
| 8 | 13 | Passo 6: Alteração (2 pessoas), Cancelamento antes do embarque, Bagagem não usada | `08 Alterar` | A28:C30 (texto das condições); entradas B18:C27; resultados A35:C44 | Texto das condições (idêntico ao livro na versão de exemplos) e cálculo de taxa de alteração, devolução, taxa de cancelamento, bagagem e perda em dinheiro e em pontos. |

**Resultado da conferência:** as oito tabelas estão representadas (teste automatizado 5.14 a 5.21 do `Relatorio-de-testes.md`).

## 2. Atividades, tarefas e fórmulas

| Capítulo | Trecho do livro | Aba | Localização |
|---|---|---|---|
| 1 | Perguntas 1 a 6 do diagnóstico (objetivos, detalhes da viagem, frequência, gastos, programas, cartão) | `02 Diagnóstico` | B7:B9; B12:B16; B19; A23:B25; A30:D44; B47 |
| 1 | "Seu resumo de ponto de partida" (fill-in com sublinhados) | `02 Diagnóstico` | B50:B56 (resumo montado em B56; "Meu primeiro passo" em B55) |
| 4 | Passo a passo e fórmula "Valor por mil = (preço em dinheiro − o que você ainda paga em dinheiro) ÷ (milhas necessárias ÷ 1.000)" | `07 Comparar` | B29 e B31 (calculadora); E49:E58 (exercício) |
| 4 | Exemplos fictícios A, B e C | `07 Comparar` (versão de exemplos) | A49:H51 |
| 4 | "Tarefa: pegue uma passagem… e faça a conta" | `07 Comparar` | A48:H58 (dez linhas) |
| 6 | "Para preencher já": listar programas; anotar saldo e validade; anotar separadamente último acesso e último acúmulo ou resgate; lembrete mensal | `03 Contas`, `04 Lotes`, `13 Checklists` | 03 A7:R36; 04 A7:N156; 13 (checklist mensal) |
| 7 | "Controle de compras" e prazos | `06 Compras` | A7:X106 |
| 7 | Exemplo do liquidificador (R$ 220 × R$ 200) | `05 Acúmulo` (versão de exemplos) | linha 7 (I7:K7) |
| 8 | As sete perguntas ao emissor | `13 Checklists` (seção extra) | A59:F67 (não conta nos 70 itens) |
| 10 | "Como acompanhar na prática": uma linha por lote, lembretes, plano de uso, ação (usar, extensão, reativação, deixar vencer) | `04 Lotes` | A7:N156 (colunas J e K) |
| 10 | Exemplo da Marina (8.000 e 12.000 pontos) | `04 Lotes` (exemplos) | linhas 7 e 8 |
| 10 | "Tarefa: para cada programa… descubra a regra de validade no site oficial" | `04 Lotes` | colunas L (fonte) e M (data) |
| 10 | "Estender ou reativar: vale a pena?" (4 passos) e exemplo do Rafael | `11 Renovar` | A7:U21; exemplo na linha 7 |
| 12 | As sete perguntas "Dinheiro ou milhas?" | `07 Comparar` | A61:C68 (adicional) |
| 13 | Passos 1 a 5 (ficha, milhas e adicionais, bagagem, comparação, alternativa) | `07 Comparar` | A5:E45 |
| 13 | Passo 6 (alteração e cancelamento) e perda calculada (R$ 336 e R$ 300) | `08 Alterar` | A6:D44 |
| 13 | Passo 7 (a decisão) e Cap. 18, semana 4 ("Decidir: resgatar, pagar em dinheiro ou esperar") | `07 Comparar` | A70:B73 |
| 15 | "Como pedir ajuda: o que guardar e o que informar" | `08 Alterar` | A48:D59 (registro de contatos e protocolos; adicional) |
| 16 | "Tarefa: pegue um clube… calcule o custo por milheiro" e a fórmula | `09 Clubes e compra` (área A) | A7:Q17 |
| 17 | Compra de pontos: custo por milheiro | `09 Clubes e compra` (área B) | A21:N31 |
| 17 | Transferência bonificada: quatro passos e exemplos de 30% e 100% | `10 Transferências` | A7:X22; exemplos nas linhas 8 e 9 |
| 18 | Plano de 30 dias (quatro semanas) | `12 Plano 30 dias` | A7:I28; progresso B32:E37 |
| 19 | Checklist do iniciante | `13 Checklists` | A35:G56 |

## 3. Os 70 itens de checklist

Soma conferida: **49 itens na aba 13 + 21 itens na aba 12 = 70**, com texto idêntico ao do texto aprovado, sem omissão nem duplicação (teste 5.12 e 5.13).

| # | Capítulo | Trecho do item | Aba | Célula |
|---|---|---|---|---|
| 1 | 11 | Abri cada conta da planilha e anotei saldo e vencimentos. | `13 Checklists` | B8 |
| 2 | 11 | Atualizei a coluna de último acúmulo ou resgate. | `13 Checklists` | B9 |
| 3 | 11 | Conferi se há pontos a vencer nos próximos 90 dias e defini o que fazer com eles. | `13 Checklists` | B10 |
| 4 | 11 | Revisei o controle de compras: algum pedido está perto do prazo de crédito ou de reclamaçã… | `13 Checklists` | B11 |
| 5 | 11 | Anotei qualquer mudança de regra que li no site oficial. | `13 Checklists` | B12 |
| 6 | 14 | As datas e os horários estão certos (incluindo fuso e horário de partida). | `13 Checklists` | B13 |
| 7 | 14 | Os aeroportos estão certos (algumas cidades têm mais de um). | `13 Checklists` | B14 |
| 8 | 14 | O nome de cada passageiro está conforme o documento que será usado no embarque (confira a … | `13 Checklists` | B15 |
| 9 | 14 | Você sabe o total de pontos ou milhas e tudo o que ainda será pago em dinheiro (taxas, bag… | `13 Checklists` | B16 |
| 10 | 14 | Você leu as regras de alteração e cancelamento, e se a tarifa é não reembolsável. | `13 Checklists` | B17 |
| 11 | 14 | O saldo está disponível e válido para concluir o resgate. Conferi separadamente a validade… | `13 Checklists` | B18 |
| 12 | 14 | Você sabe como e quando receberá a confirmação e o bilhete. | `13 Checklists` | B19 |
| 13 | 15 | Reli as regras de alteração e cancelamento da minha opção. | `13 Checklists` | B20 |
| 14 | 15 | Anotei os prazos importantes no calendário. | `13 Checklists` | B21 |
| 15 | 15 | Guardei comprovantes e protocolos. | `13 Checklists` | B22 |
| 16 | 15 | Contatei o canal oficial dentro do prazo. | `13 Checklists` | B23 |
| 17 | 17 | Tenho um uso concreto (viagem, data, destino). | `13 Checklists` | B24 |
| 18 | 17 | Calculei o custo por milheiro com todos os custos. | `13 Checklists` | B25 |
| 19 | 17 | Comparei com o valor que obtenho no resgate. | `13 Checklists` | B26 |
| 20 | 17 | Li as condições da promoção (prazo, limite, quem pode participar). | `13 Checklists` | B27 |
| 21 | 17 | Conferi se é necessário cadastrar-se na promoção antes da transferência. | `13 Checklists` | B28 |
| 22 | 17 | Verifiquei separadamente os prazos de crédito dos pontos e do bônus. | `13 Checklists` | B29 |
| 23 | 17 | Considerei que a passagem encontrada pode mudar de preço ou deixar de estar disponível dur… | `13 Checklists` | B30 |
| 24 | 17 | Conferi a validade das milhas que vou receber, inclusive do bônus. | `13 Checklists` | B31 |
| 25 | 17 | Estou cadastrado no programa de destino, e os dados estão corretos. | `13 Checklists` | B32 |
| 26 | 17 | Sei quanto tempo o crédito leva e vou conferir a chegada. | `13 Checklists` | B33 |
| 27 | 17 | Tenho orçamento para isso, sem juros. | `13 Checklists` | B34 |
| 28 | 18 | Preencher o diagnóstico (Capítulo 1) e o resumo do seu ponto de partida. | `12 Plano 30 dias` | C8 |
| 29 | 18 | Listar seus programas e entrar em cada conta (Capítulo 6). | `12 Plano 30 dias` | C9 |
| 30 | 18 | Montar a planilha de contas: saldo, uma linha por lote (programa, quantidade, origem e ven… | `12 Plano 30 dias` | C10 |
| 31 | 18 | Marcar no calendário os lembretes de vencimento (sugestão: 90, 60 e 30 dias antes). | `12 Plano 30 dias` | C11 |
| 32 | 18 | Revisar a segurança das contas: senhas diferentes e fortes, e verificação em duas etapas q… | `12 Plano 30 dias` | C12 |
| 33 | 18 | Preencher o mapa de acúmulo: o que você já gasta, onde, e se o lugar tem programa (Capítul… | `12 Plano 30 dias` | C13 |
| 34 | 18 | Como ponto de partida, considerar concentrar o acompanhamento em poucos programas. É uma s… | `12 Plano 30 dias` | C14 |
| 35 | 18 | Se houver uma compra já planejada, testar o procedimento pelo caminho do programa, compara… | `12 Plano 30 dias` | C15 |
| 36 | 18 | Abrir o controle de compras: data, loja, nº do pedido, pontuação esperada, evento que inic… | `12 Plano 30 dias` | C16 |
| 37 | 18 | Se você tem cartão, fazer as perguntas da lista ao emissor (Capítulo 8), sem gastar a mais… | `12 Plano 30 dias` | C17 |
| 38 | 18 | Descobrir, no site oficial, a regra de validade de cada programa e de cada origem de ponto… | `12 Plano 30 dias` | C18 |
| 39 | 18 | Anotar regras de inatividade de cada programa e quando foi o último acúmulo ou resgate. | `12 Plano 30 dias` | C19 |
| 40 | 18 | Verificar se algum pedido está perto do prazo de crédito ou de reclamação (Capítulo 7). | `12 Plano 30 dias` | C20 |
| 41 | 18 | Revisar a lista de erros comuns (Capítulo 11) e marcar os que podem acontecer com você. | `12 Plano 30 dias` | C21 |
| 42 | 18 | Fazer o primeiro checklist mensal (reserve cerca de 15 minutos, ajustando ao número de con… | `12 Plano 30 dias` | C22 |
| 43 | 18 | Escolher uma viagem real que você gostaria de fazer, com datas e passageiros. | `12 Plano 30 dias` | C23 |
| 44 | 18 | Buscar o preço em dinheiro e as opções com milhas, com as mesmas condições (Capítulo 13). | `12 Plano 30 dias` | C24 |
| 45 | 18 | Calcular o valor por milheiro com e sem a melhor alternativa em dinheiro. | `12 Plano 30 dias` | C25 |
| 46 | 18 | Ler as condições de alteração e cancelamento e os valores adicionais (taxas, bagagem). | `12 Plano 30 dias` | C26 |
| 47 | 18 | Decidir: resgatar, pagar em dinheiro ou esperar. Registrar o motivo. | `12 Plano 30 dias` | C27 |
| 48 | 18 | Só então, se fizer sentido, estudar estratégias pagas (Parte 5) e calcular o custo por mil… | `12 Plano 30 dias` | C28 |
| 49 | 19 | Tenho um objetivo de viagem definido (Cap. 1). | `13 Checklists` | B35 |
| 50 | 19 | Entendo a diferença entre pontos, milhas e valor (Caps. 2 e 4). | `13 Checklists` | B36 |
| 51 | 19 | Sei que pontos não são dinheiro, não têm cotação fixa garantida e que as regras mudam (Cap… | `13 Checklists` | B37 |
| 52 | 19 | Listei meus programas e conferi saldos e validades (Cap. 6). | `13 Checklists` | B38 |
| 53 | 19 | Fiz o meu mapa de acúmulo (Cap. 7). | `13 Checklists` | B39 |
| 54 | 19 | Comparo o preço final antes de comprar por um caminho de programa (Cap. 7). | `13 Checklists` | B40 |
| 55 | 19 | Conferi se o meio de pagamento é aceito pela campanha (Cap. 7). | `13 Checklists` | B41 |
| 56 | 19 | Guardo comprovantes e acompanho prazos de crédito e de reclamação (Cap. 7). | `13 Checklists` | B42 |
| 57 | 19 | Se uso cartão, pago a fatura inteira e não gasto a mais para pontuar (Cap. 8). | `13 Checklists` | B43 |
| 58 | 19 | Uso senhas diferentes e não compartilho senhas ou códigos (Cap. 9). | `13 Checklists` | B44 |
| 59 | 19 | Acesso apenas pelos canais oficiais (Cap. 9). | `13 Checklists` | B45 |
| 60 | 19 | Mantenho uma linha por lote, com programa, quantidade, origem e vencimento (Cap. 10). | `13 Checklists` | B46 |
| 61 | 19 | Sei a diferença entre vencimento dos pontos, encerramento por inatividade e extensão ou re… | `13 Checklists` | B47 |
| 62 | 19 | Não pago para estender ou reativar sem um uso concreto (Cap. 10). | `13 Checklists` | B48 |
| 63 | 19 | Comparei com a melhor alternativa em dinheiro (Caps. 12 e 13). | `13 Checklists` | B49 |
| 64 | 19 | Somei milhas, taxas, bagagem e serviços (Cap. 13). | `13 Checklists` | B50 |
| 65 | 19 | Li as condições de alteração e cancelamento (Caps. 13 e 14). | `13 Checklists` | B51 |
| 66 | 19 | Conferi nome, datas e aeroportos duas vezes (Cap. 14). | `13 Checklists` | B52 |
| 67 | 19 | Acompanho a emissão logo após o resgate (Cap. 14). | `13 Checklists` | B53 |
| 68 | 19 | Só considero clube, compra de pontos ou transferência bonificada com uso concreto (Caps. 1… | `13 Checklists` | B54 |
| 69 | 19 | Calculo o custo por milheiro com todos os custos (Caps. 16 e 17). | `13 Checklists` | B55 |
| 70 | 19 | Confiro a validade do bônus e as regras da promoção (Cap. 17). | `13 Checklists` | B56 |

## 4. Itens adicionais (além da especificação mínima)

- `07 Comparar` seção F (as sete perguntas do Capítulo 12) e seção G (decisão do Capítulo 13, passo 7, e do Capítulo 18).
- `08 Alterar` seção D (registro de contatos e protocolos, Capítulo 15).
- `13 Checklists` seção extra (as sete perguntas ao emissor, Capítulo 8), fora da contagem dos 70 itens.
- `03 Contas`: prazo de inatividade informado pelo programa e alerta (sugestão de organização; não infere o prazo).
- `07 Comparar` seção E: valor contra alternativa de R$ 700 no Exemplo C (derivado da mesma fórmula; o livro não informa o número).

## 5. Partes do e-book que não viraram ferramenta (texto explicativo ou lista de orientações)

Capítulos 2, 3 e 5 (conceitos e mitos); Capítulo 8 (visão geral, salvo as perguntas ao emissor); Capítulo 9 (boas práticas de segurança, sem campos); Capítulo 11 (os 10 erros comuns, sem campos; o checklist mensal foi convertido); Capítulo 14 (passo a passo geral e exemplo da Livelo, que são texto; a lista de checagem foi convertida); Capítulo 15 (textos de orientação; o checklist rápido foi convertido); Capítulo 20 (glossário e perguntas frequentes); "Fontes e atualização".
