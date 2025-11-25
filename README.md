#POP – Macro "PintarCriterio"

1. Objetivo
Este procedimento tem como finalidade identificar e destacar, com preenchimento amarelo, combinações de valores na coluna J que somam zero, agrupadas por um critério específico (ex: FLC), respeitando regras de visibilidade e uso prévio.

2. Abrangência
Aplicável a planilhas do Excel que contenham dados a partir da linha 6, com valores numéricos na coluna J e critérios de agrupamento em outra coluna definida pelo usuário.

3. Responsável
Usuário da planilha com permissão para executar macros no Excel.

4. Materiais Necessários
Planilha com dados válidos nas colunas J (valores) e na coluna do critério (ex: FLC).
Macro "PintarCriterio" devidamente instalada e associada a um botão.

5. Procedimentos
5.1 Preparação
Abra a planilha desejada.
Certifique-se de que:
Os dados iniciam na linha 6.
A coluna J contém apenas valores numéricos.
A coluna M está livre para marcação dos grupos.
A célula N3 contém o número do último grupo (ou esteja vazia para iniciar do zero).

5.2 Execução
Clique no botão da macro PintarCriterio.
Na caixa de entrada, informe a letra da coluna que contém o critério de agrupamento (ex: F).
Responda à pergunta:

Clique em Sim para incluir trios, quartetos, quintetos e sextetos.
Clique em Não para analisar apenas pares.
Clique em Cancelar para interromper o processo.

Aguarde a execução. O Excel exibirá uma barra de progresso.
Ao final, será exibida uma mensagem com:

Tempo de execução.
Total de células pintadas.
6. Regras e Restrições
Apenas células visíveis da coluna J são analisadas.
Células já pintadas de amarelo são ignoradas.
Apenas valores numéricos são considerados.
O critério informado não pode estar vazio ou ser “-”.

7. Resultados Esperados
Células da coluna J que formam combinações que somam zero são pintadas de amarelo.
A coluna M é preenchida com o número do grupo correspondente.
A célula N3 é atualizada com o número do último grupo utilizado.

8. Observações
Recomenda-se executar a macro LimparColunaJ antes de uma nova análise para remover marcações anteriores.
A macro pode ser adaptada para outros critérios além de FLC, desde que estejam em colunas com dados válidos.

 
#POP – Macro "LimparColunaJ"
1. Objetivo
Remover o preenchimento (cor de fundo) das células da coluna J, da linha 6 até a última linha com valor, para possibilitar uma nova análise limpa.

2. Abrangência
Aplicável a planilhas do Excel com dados na coluna J a partir da linha 6.

3. Responsável
Usuário da planilha com permissão para executar macros no Excel.

4. Materiais Necessários
Planilha com dados na coluna J.
Macro "LimparColunaJ" associada a um botão.

5. Procedimentos
5.1 Preparação
Abra a planilha desejada.
Verifique se há dados na coluna J a partir da linha 6.

5.2 Execução
Clique no botão da macro LimparColunaJ.
A macro será executada automaticamente.
Ao final, será exibida a mensagem: “Limpeza Completa!”

6. Regras e Restrições
A limpeza começa sempre da linha 6.
Apenas a coluna J é afetada.
A macro não apaga valores, apenas remove a cor de fundo.

7. Resultados Esperados
Células da coluna J com preenchimento removido.
Planilha pronta para nova análise com outras macros.

 
#POP – Macro "AcharValor_v2"
1. Objetivo
Encontrar uma combinação de valores na coluna J que somem exatamente o valor informado na célula J4, destacando as células envolvidas com a cor laranja.

2. Abrangência
Aplicável a planilhas do Excel com valores numéricos na coluna J e um valor-alvo definido na célula J4.

3. Responsável
Usuário da planilha com permissão para executar macros no Excel.

4. Materiais Necessários
Valor-alvo digitado na célula J4.
Intervalo de valores visíveis e não pintados na coluna J.
Macro "AcharValor_v2" associada a um botão.

5. Procedimentos
5.1 Preparação
Digite o valor-alvo na célula J4 (ex: -23659,05).
Verifique se as células da coluna J estão visíveis e sem preenchimento.

5.2 Execução
Clique no botão da macro AcharValor_v2.
Selecione com o mouse o intervalo da coluna J a ser analisado (ex: J6:J100).
A macro será executada automaticamente.
Ao final, será exibida uma das mensagens:

“Combinação encontrada com X células!”
“Nenhuma combinação encontrada.”

6. Regras e Restrições
O valor-alvo deve estar na célula J4.
Apenas células visíveis, sem cor de fundo e numéricas são consideradas.
A cor usada para destaque é laranja (RGB 255, 165, 0).

7. Resultados Esperados
Células que somam exatamente o valor de J4 são pintadas de laranja.


#POP – Macro "PintarPares"
1. Objetivo
Identificar pares de valores opostos (ex: +100 e -100) dentro de um intervalo selecionado e destacá-los com preenchimento amarelo, registrando o número do par na coluna M.

2. Abrangência
Aplicável a planilhas do Excel com valores numéricos na coluna J ou outra coluna selecionada.

3. Responsável
Usuário da planilha com permissão para executar macros no Excel.

4. Materiais Necessários
Intervalo de valores visíveis, não pintados e numéricos.
Macro "PintarPares" associada a um botão.

5. Procedimentos
5.1 Preparação
Abra a planilha desejada.
Verifique se a célula N3 está disponível para controle de numeração.
Certifique-se de que a coluna M esteja livre para marcação dos pares.

5.2 Execução
Clique no botão da macro PintarPares.
Selecione com o mouse o intervalo da coluna a ser analisada (ex: J6:J100).
A macro será executada automaticamente.
Ao final, será exibida uma mensagem com:

Tempo de execução.
Total de pares encontrados.
Lembrete: “Não deixa de verificar!”

6. Regras e Restrições
Apenas células visíveis, não pintadas e com valores numéricos diferentes de zero são consideradas.
A célula N3 guarda o número do último par.
A cor usada para destaque é amarelo (RGB 255, 255, 0).

7. Resultados Esperados
Pares de valores opostos são pintados de amarelo.
A coluna M é preenchida com “Par Nº X”.


#POP – Macro "ChequeFuncionarios_v2 – DIVERSOS"
1. Objetivo
Identificar e destacar lançamentos compensados ou passíveis de abatimento por cheques (CHQ.CX) na coluna J, agrupando por Complemento e, se necessário, por FLC. Também calcula sobras positivas e tenta compensá-las com cheques dentro da tolerância de ±6 meses.

2. Abrangência
Aplicável a planilhas do Excel com lançamentos financeiros na coluna J e dados auxiliares nas colunas de Complemento e FLC.

3. Responsável
Usuário da planilha com permissão para executar macros no Excel.

4. Materiais Necessários
Planilha com lançamentos e cheques organizados.
Macro "ChequeFuncionarios_v2 – DIVERSOS" associada a um botão.

5. Procedimentos
5.1 Preparação
Verifique se os dados estão organizados a partir da linha 6.
Confirme que a coluna J contém os valores dos lançamentos.
Certifique-se de que as colunas de Complemento e FLC estejam preenchidas corretamente.

5.2 Execução
Clique no botão da macro ChequeFuncionarios_v2 – DIVERSOS.
A macro será executada automaticamente.
Ao final, será exibida uma mensagem com o resumo da execução.

6. Regras e Restrições
Apenas células visíveis e não utilizadas são consideradas.
A macro respeita a tolerância de ±6 meses para compensação com cheques.
Agrupamento é feito por Complemento e, se necessário, por FLC.

7. Resultados Esperados
Células da coluna J que representam compensações são pintadas.
Cheques (CHQ.CX) são utilizados para abater sobras positivas dentro da tolerância.
A coluna M pode ser usada para marcação dos grupos.
