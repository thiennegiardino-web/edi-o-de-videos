# Decisões do Codex sobre os dez achados (registro)

**Origem:** parecer de revisão do Codex — revisão assistida por IA, relatada pela autora (Thienne Giardino) em 17/11/2026, após conferir a v1.1.
**Estado:** a v1.1 (commit `c1238bf`, `pacote-v1.1/`, ZIP SHA-256 `cc702614eb144e4c2b98ee7c3e085173b8660c0315ff4fa448c5d2909fb23f9e`) **segue para validação no Excel** sem alteração. Este registro não altera nenhum arquivo do pacote nem o texto do livro. Alterações no texto entram em **nova versão identificada**.

| # | Achado | Decisão | Onde atua | Situação |
|---|---|---|---|---|
| 1 | Tarifa: unidades misturadas na tabela do Cap. 13 | Separar tarifa em reais da quantidade de milhas | Texto do livro | Nova versão do livro (pendente) |
| 2 | "Valor por mil" × "valor por milheiro" | Padronizar "valor por milheiro" | Texto do livro (Cap. 4) | Nova versão do livro (pendente). A planilha já usa "milheiro" nos rótulos; uma ocorrência de "valor por milheiro" na aba 07 permanece conforme. |
| 3 | Situação "Divergente" | Manter na planilha; explicar nas instruções | Aba 01 / LEIA-ME | Instrução a acrescentar (próxima versão da planilha) |
| 4 | Contas e lotes em abas separadas | Manter; explicar a divisão | Aba 01 / LEIA-ME / texto | Instrução a acrescentar (próxima versão da planilha) |
| 5 | Pontos a vencer dependem dos lotes | Informar a dependência do preenchimento dos lotes | Aba 03 / aba 01 / texto | Instrução a acrescentar (próxima versão da planilha) |
| 6 | Voz dos checklists | Ajuste opcional; **preservar agora** os textos aprovados | — | Nenhuma alteração |
| 7 | Bônus e limite | Esclarecer que a fórmula simples pressupõe ausência de limite aplicável; seguir a unidade e as condições da campanha | Texto (Cap. 17) e aba 10 | Nova versão do livro; nota na aba 10 na próxima versão da planilha |
| 8 | Vencimento às 23h59 | Avisar que a ferramenta considera datas; o horário deve ser confirmado no programa | Aba 04 / LEIA-ME / texto | Instrução a acrescentar (próxima versão da planilha) |
| 9 | "Poucos programas" | Nenhuma alteração necessária | — | Encerrado |
| 10 | Campanha Esfera (09 a 15/11/2026) | Reconferir antes da publicação; não apresentar condições da campanha encerrada como regra geral | Texto (Cap. 7) | Pendência de reconferência pré-publicação |

## Observação
Os itens 3, 4, 5, 7 (aba 10) e 8 pedem **apenas notas explicativas** na planilha (sem mudar fórmulas). Como qualquer mudança no `.xlsx` altera os checksums e invalida a validação em curso no Excel, **não foram aplicadas**; ficam para uma v1.2, a ser gerada **depois** da validação da v1.1 (para a validação incorporar também as correções que o Excel apontar).
