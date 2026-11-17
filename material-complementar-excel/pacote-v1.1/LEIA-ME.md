# Material complementar em Excel — Milhas do zero (De Malas e Milhas)

**Versão 1.1** (corrige os três pontos do parecer do Codex sobre a v1.0, commit `6892d31`; ver «Mudanças da v1.1»). **Para revisão pelo Codex.** Pacote gerado em 17/11/2026.
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

## Mudanças da v1.1 (em relação à v1.0, commit `6892d31`, preservada)
1. **Saldo sem pontos vencidos (aba 03).** O saldo considerado (coluna E) passou a somar só os lotes **não vencidos**. Os pontos de lotes vencidos aparecem na nova coluna **R** ("Pontos vencidos (fora do saldo)"), e a origem do saldo (coluna F) avisa "Há pontos vencidos (coluna R)". Lote com data em branco, em texto ou com «expiração confirmada = Sim, confirmado» continua no saldo; lote que vence **hoje** também continua (vence ao fim do dia).
2. **Reclamação não é mais escondida (aba 06).** O alerta (coluna V) passou a **somar** as mensagens: divergência, data-limite de reclamação vencida ou em até 7 dias (inclusive quando a situação é «Creditado» mas há divergência), falta de data-limite calculada e prazo de crédito vencido. A aba 01 conta cada compra com alerta uma única vez.
3. **Campo vazio = «não informado», nunca zero (abas 07 a 11).** Custos, taxas, bônus e outros pagamentos em branco **não** viram R$ 0: o total/perda/valor por milheiro fica em branco e a planilha lista os campos pendentes. Quem não tem o custo deve **digitar 0**. Nos exemplos, os zeros agora estão digitados. Na aba 08, apagar a taxa de cancelamento não reduz mais a perda de R$ 300 para R$ 0: a perda fica em branco até o campo ser preenchido.

## Como o aluno usa (resumo)
1. Abra a aba **01 Comece aqui** e siga a ordem das abas, junto com os capítulos indicados.
2. Preencha só as células **claras com borda** (campos). As células **rosadas** são calculadas e protegidas.
3. **Campo vazio significa «não informado», não zero.** Se um custo, taxa ou bônus não existe, digite **0**; sem isso, a planilha não calcula o total e avisa o que falta.
4. Registre a **fonte e a data da consulta** de cada regra de programa. A planilha **não** infere validade nem prazo pelo nome do programa.
5. Uma vez por mês, volte à aba 01 (resumo) e faça o checklist mensal (aba 13).

## As 14 abas
| Aba (nome exato) | Capítulo(s) | O que faz |
|---|---|---|
| `01 Comece aqui` | Todos | Instruções, legenda, navegação, avisos, configurações (data de referência) e resumo simples. |
| `02 Diagnóstico` | Capítulo 1 | Diagnóstico inicial, programas com saldo e validade, resumo do ponto de partida. |
| `03 Contas` | Capítulos 6 e 9 | Uma linha por programa; saldo vem dos lotes não vencidos (vencidos na coluna R); último acesso separado de último acúmulo ou resgate. |
| `04 Lotes` | Capítulos 10 e 11 | Uma linha por lote; alertas de vencimento (30, 60 e 90 dias). |
| `05 Acúmulo` | Capítulo 7 | Mapa de acúmulo e comparação de preço final. |
| `06 Compras` | Capítulo 7 | Controle de compras: datas, prazos, divergência, reclamação. |
| `07 Comparar` | Capítulos 4, 12 e 13 | Calculadora dinheiro × pontos, exercício do Capítulo 4, perguntas do Capítulo 12 e decisão. |
| `08 Alterar` | Capítulos 13 a 15 | Alteração, cancelamento e bagagem; registro de contatos e protocolos. |
| `09 Clubes e compra` | Capítulos 16 e 17 | Custo por milheiro de clubes e de compra de pontos. |
| `10 Transferências` | Capítulo 17 | Transferência bonificada com fator de conversão e limite de bônus. |
| `11 Renovar` | Capítulo 10 | Extensão ou reativação: comparação de desembolsos. |
| `12 Plano 30 dias` | Capítulo 18 | As 21 atividades do plano de 30 dias, com progresso. |
| `13 Checklists` | Capítulos 11, 14, 15, 17 e 19 | Os 49 itens dos checklists, com grupos e datas de revisão; extra com perguntas ao emissor. |
| `14 Referências` | Capítulo 10 e fontes | Tabelas explicativas, exemplos datados da Smiles e fontes do material. |

## Legenda
- **Campo para preencher:** fundo marfim claro (`#FFFEFA`), borda cobre clara; desbloqueado.
- **Campo calculado:** fundo rosé (`#F4E3D9`); bloqueado; resultados e avisos.
- **Cabeçalho:** cobre escuro (`#A8562B`); cabeçalho de coluna calculada em cobre mais claro (`#C2703D`).
- **Alertas de lote (aba 04):** vermelho = vencido; laranja = vence hoje ou em até 30 dias; amarelo = até 60; amarelo claro = até 90; cinza = data não informada; verde = sem expiração confirmada. Os intervalos são **sugestões de organização**, não prazos oficiais. Todo alerta também aparece em **texto** (não depende só da cor).
- **Exemplo adicional fictício:** aparece escrito em Observações como "[Exemplo adicional fictício]"; "[Exemplo do livro, Cap. N]" indica números ou situações do e-book.

## Decisões de projeto que o revisor deve conhecer
- **Saldo sem entrada dupla:** na aba 03, o saldo é a **soma dos lotes não vencidos** da aba 04 (vencidos ficam na coluna R). O campo "Saldo manual" só vale para programas sem lotes; se houver lotes e um saldo manual divergente, o manual é ignorado e isso é informado.
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
- **Testado em:** LibreOffice 24.2.7.2 420(Build:2) (modo headless, recálculo total forçado), geradas com Python 3.13.18, XlsxWriter 3.2.9 e openpyxl 3.1.5.
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
Os arquivos são gerados por `material-complementar-excel/build/build_planilhas.py` e testados por `material-complementar-excel/testes/testar.py` (ambos na branch `claude/compassionate-dijkstra-4mf2ce`).
