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

POP – Macro "Listar Linkar Conteúdo" no mesmo modelo:

1. Objetivo
Listar pastas e arquivos de uma pasta selecionada pelo usuário, criando uma planilha chamada “ListaConteudo” (ou nome sanitizado/único), com dados organizados e link clicável para abrir cada item. Inclui tratamento seguro para nome de planilha, exclusão de planilha existente, formatação visual (cabeçalho, zebrado, filtros, bordas) e hiperlinks.

2. Abrangência
Aplicável a qualquer pasta do Windows selecionada via diálogo e a planilhas do Excel com macros habilitadas.
A listagem contempla subpastas do nível selecionado e arquivos diretamente dentro da pasta (não recursivo).

3. Responsável
Usuário com permissão para:
Executar macros no Excel.
Acessar a pasta desejada no sistema de arquivos.

4. Materiais Necessários
Pasta contendo arquivos e/ou subpastas.
Arquivo Excel habilitado para macros (.xlsm) com a macro “ListarLinkarConteúdo”.
Botão vinculado à macro (opcional).

5. Procedimentos
5.1 Preparação
Abra a planilha com a macro habilitada.
Garanta que o Excel está com macros habilitadas.
Tenha a pasta-alvo em mente para seleção.
5.2 Execução
Clique no botão da macro ListarLinkarConteúdo (ou execute pelo Editor VBA).
Na janela, selecione a pasta a ser listada.
A macro:

Tenta excluir com segurança a planilha “ListaConteudo” se existir.
Cria uma nova planilha e aplica o nome com higiene (remove caracteres inválidos, limita a 31, evita apóstrofos nas pontas).

Se o nome ainda falhar, cria um nome único com sufixo “(n)”.

Gera cabeçalho com colunas:
A: Nome | B: Tipo | C: Data Modificação | D: Tamanho (KB) | E: Abrir
Lista subpastas (nome, tipo “Pasta”, data, “-” no tamanho, link “Abrir Pasta”).
Lista arquivos (nome, tipo “Arquivo”, data, tamanho em KB arredondado, link “Abrir Arquivo”).
Aplica formatação (AutoFit, centralização padrão, nome à esquerda, número com milhar em D, data completa em C, linhas zebradas, filtros e bordas).

Ao final, o Excel exibirá uma mensagem de sucesso com o nome da planilha criada.

6. Regras e Restrições
Não é recursivo: lista apenas subpastas e arquivos do nível selecionado (não entra em subpastas para listar seus conteúdos).
Planilha alvo é renomeada com segurança: remoção de caracteres inválidos (: \ / ? * [ ]), trim, sem apóstrofo nas extremidades, limite de 31 caracteres.
Se já existir a planilha, a macro tenta excluir com segurança (desprotege, muda a ativa caso necessário, cria uma planilha extra se for única, suprime alertas).
Hiperlinks apontam diretamente para o Path de pasta/arquivo.
Datas exibidas como dd/mm/yyyy hh:mm e tamanho em KB sem casa decimal, com separador de milhar.
Apenas itens acessíveis (sem erro do sistema) serão listados.

7. Resultados Esperados
Criação de uma planilha formatada e filtrável com:

Pastas do nível selecionado com link “Abrir Pasta”.
Arquivos do nível selecionado com link “Abrir Arquivo”.
Datas de modificação e tamanhos (KB) dos arquivos.

Visual padronizado: cabeçalho estilizado, linhas zebradas, bordas e hiperlinks destacados.
Mensagem final: “Lista criada com sucesso na planilha 'ListaConteudo'!” (ou nome final aplicado).
 
POP – Macro "Verificar Grifos Word_Excel" no mesmo modelo:

1. Objetivo
Analisar documentos Word (.docx e/ou .doc) em uma pasta (com opção de varrer subpastas) e gerar um relatório no Excel indicando:
Se o arquivo possui texto com destaque (Highlight).
Quantidade total de ocorrências de grifos.
Status (OK ou erro ao abrir).

2. Abrangência
Aplicável a:
Windows + Microsoft Word instalado (automação via Late Binding).
Arquivo Excel habilitado para macros (.xlsm).
Varredura de arquivos .docx e .doc na pasta selecionada (com opção de incluir subpastas).

3. Responsável
Usuário com permissão para:
Executar macros no Excel.
Acessar as pastas/arquivos Word.
Ter o Microsoft Word instalado e funcional.

4. Materiais Necessários
Planilha Excel com a macro “VerificarGrifosWord_Excel”.
Microsoft Word instalado (qualquer versão compatível com automação COM).
Pasta contendo os arquivos .docx e/ou .doc a analisar.

5. Procedimentos
5.1 Preparação
Abra o arquivo .xlsm com a macro.
Certifique-se de que macros estão habilitadas.
(Opcional) Ajuste os parâmetros no topo do código:
INCLUDE_SUBFOLDERS = True para incluir subpastas.
SCAN_DOCX = True para analisar .docx.
SCAN_DOC = True para analisar .doc.

5.2 Execução
Execute a macro VerificarGrifosWord_Excel (botão ou Editor VBA).
Na janela, selecione a pasta com os documentos Word.
A macro:

Prepara/limpa a planilha “Relatorio_Grifos”.
Cria (via Late Binding) uma instância do Word (oculto, sem alertas).
Lista arquivos conforme filtros (.docx/.doc) e varre subpastas se habilitado.
Para cada arquivo:

Abre somente leitura no Word.
Percorre todas as StoryRanges e usa Find com Format = True e Highlight = True.
Conta cada ocorrência de texto grifado e marca status (“Com grifo”, “OK” ou “Erro ao abrir”).

Gera tabela com colunas:

A: Arquivo
B: Caminho
C: Status
D: Ocorrências (Highlight)

Aplica AutoFit e cria Tabela (ListObject) para facilitar filtro/uso.

Ao finalizar, exibe mensagem: quantidade de arquivos analisados e localização do relatório.

6. Regras e Restrições
Automação Word via Late Binding: não precisa adicionar referência no VBA, mas exige Word instalado.
Tipos suportados: apenas .docx e/ou .doc conforme SCAN_DOCX/SCAN_DOC.
(Não analisa .docm, .rtf, .pdf, etc.)
A análise procura formatação Highlight (marca-texto) — não verifica cores de fonte sem Highlight.
Abertura é em somente leitura; arquivos protegidos por senha, corrompidos ou sem acesso geram status de erro.
Performance: muitas subpastas/arquivos podem aumentar o tempo; ajuste INCLUDE_SUBFOLDERS conforme necessário.
A planilha “Relatorio_Grifos” é recriada/limpa a cada execução.
Word é encerrado ao fim da macro (Quit) para liberar recursos.

7. Resultados Esperados
Relatório limpo e filtrável na planilha “Relatorio_Grifos” com:

Nome do arquivo e caminho completo.
Status: “Com grifo”, “OK” (sem highlight) ou “Erro ao abrir”.
Total de ocorrências de highlight por arquivo.

Mensagem final com o total de arquivos analisados.
 
POP – Macro "Criar Tabela Com Hiperlinks" no mesmo modelo:

1. Objetivo
Gerar e manter um Sumário das planilhas do arquivo Excel, criando uma tabela com:

Nome da planilha (coluna A)
Link “Acessar” para ir direto à planilha (coluna B)
Um botão “Atualizar Sumário” no topo (A1:B1) que reexecuta a macro
Em cada planilha (exceto o Sumário), garantir apenas um link “← Voltar ao Sumário” na linha 1, com formatação padronizada.

2. Abrangência
Aplicável a qualquer arquivo Excel habilitado para macros (.xlsm) contendo múltiplas planilhas.
A macro atua somente na planilha “Sumário” (criando/atualizando) e adiciona/normaliza um link de retorno nas demais planilhas.

3. Responsável
Usuário com permissão para:
Executar macros no Excel.
Organizar e manter as planilhas do arquivo.

4. Materiais Necessários
Arquivo Excel .xlsm com a macro “CriarTabelaComHiperlinks”.
(Opcional) Botão já vinculado à macro no topo do Sumário.

5. Procedimentos
5.1 Preparação
Abra o arquivo .xlsm com macros habilitadas.
Verifique/garanta que não há bloqueios de macros pelo Excel (Centro de Confiabilidade).

5.2 Execução
Execute a macro CriarTabelaComHiperlinks (via botão ou Editor VBA).
A macro:

Cria/garante a existência da planilha “Sumário” (no final do workbook).
Limpa somente as colunas A e B dessa planilha.
Escreve o cabeçalho em A3:B3:

A3: “Nome da Planilha”
B3: “Link”
Formata cabeçalho (negrito, cinza, centralizado).

Percorre todas as planilhas (exceto Sumário):

Preenche A com o nome da planilha.
Cria Hiperlink em B apontando para '<Nome>'!A1 com texto “Acessar”.
Chama GarantirLinkVoltar para que cada planilha tenha apenas um link “← Voltar ao Sumário” na linha 1, removendo duplicados e aplicando formatação padrão.

Cria/atualiza a Tabela (ListObject) tbSumarioPlanilhas em A3:B…, com estilo “TableStyleMedium9”.
Aplica formatação (Arial 14, centralizado, bordas internas e externas espessas).
AutoFit nas colunas A:B.
(Re)cria o botão arredondado “Atualizar Sumário” em A1:B1:

Estilo azul Office, texto “Atualizar Sumário”, fonte Segoe UI 14, negrito, com sombra.
OnAction aponta para “CriarTabelaComHiperlinks” para reexecutar a macro.

Ao final: aparece a mensagem “Sumário atualizado com sucesso!”.

6. Regras e Restrições
Somente A:B da planilha Sumário são limpas e formatadas; o resto do Sumário permanece intacto.
O botão é sempre recriado com nome “btnAtualizarSumario”; se existir, é excluído previamente.
A Tabela do Sumário em A:B é removida se estiver sobreposta, para evitar duplicidade, e então recriada.
Em cada planilha (exceto Sumário):

O link “← Voltar ao Sumário” é único (se houver mais de um, os extras são removidos).
Caso não exista, é criado na primeira célula vazia da linha 1 (até coluna 50) e formatado (azul, negrito, fundo azul claro, centralizado).

Hiperlinks “Acessar” na coluna B do Sumário usam SubAddress para navegar dentro do workbook (sem Address externo).
A macro usa ScreenUpdating/EnableEvents para melhorar performance e evita flicker.

7. Resultados Esperados
Planilha Sumário criada/atualizada com:

Tabela tbSumarioPlanilhas listando todas as planilhas (nome + link “Acessar”).
Botão “Atualizar Sumário” funcional no topo (A1:B1).

Em cada planilha (exceto Sumário), um único link “← Voltar ao Sumário” na linha 1, padronizado.
Navegação rápida e consistente entre planilhas com links internos.
 
POP – Macro "Padronizar Nomes" no mesmo modelo:

1. Objetivo
Padronizar e renomear planilhas do arquivo Excel com base nos novos nomes informados na coluna C da planilha “Sumário”, aplicando:

Sanitização (remoção de caracteres inválidos e limite de 31 chars).
Garantia de nome único (sufixos “ (2)”, “ (3)”, … quando necessário).
Relatório final com itens renomeados e ignorados/com atenção.
Marcação visual em C (vermelho claro) quando o novo nome for inválido ou ocorrer falha.

2. Abrangência
Aplicável a qualquer arquivo Excel habilitado para macros (.xlsm) que possua a planilha “Sumário” já gerada pela macro “CriarTabelaComHiperlinks” e os nomes atuais em A4:A.
Os novos nomes devem ser preenchidos manualmente em C4:C.

3. Responsável
Usuário com permissão para:

Executar macros no Excel.
Editar a planilha “Sumário” (coluna C).

4. Materiais Necessários
Arquivo Excel .xlsm com as macros “CriarTabelaComHiperlinks” e “PadronizarNomes”.
Planilha “Sumário” existente com:

A3: cabeçalho “Nome da Planilha”
A4:A…: nomes atuais das abas
C3: cabeçalho (opcional “Novo Nome”)
C4:C…: novos nomes desejados (preencher antes da execução)

5. Procedimentos
5.1 Preparação
Abra o arquivo .xlsm com macros habilitadas.
Execute primeiro a macro “CriarTabelaComHiperlinks” para garantir a estrutura do Sumário.
Preencha os novos nomes em C4:C… do Sumário, linha a linha conforme cada planilha listada em A.

5.2 Execução
Rode a macro “Padronizar Nomes”.
A macro:

Valida a existência da planilha “Sumário”.
Identifica a última linha de dados em A (espera cabeçalho em A3 e dados a partir de A4).
Para cada linha (A4:C…):

Lê o nome atual (A) e o novo nome (C).
Ignora se C estiver vazio (registra no relatório).
Não renomeia a planilha “Sumário” (registra no relatório).
Verifica se a planilha indicada em A existe (senão, registra no relatório).
Sanitiza o novo nome:

Remove : \ / ? * [ ] e aspas duplas.
Trim e limite de 31 caracteres.
Se o resultado ficar vazio, pinta C (vermelho claro) e registra no relatório.

Garante nome único no arquivo (adiciona sufixo “ (n)” quando necessário).
Se o nome já é igual ao atual, apenas registra como ignorado.
Tenta renomear:

Em caso de erro, pinta C (vermelho claro) e registra a descrição do erro.
Em caso de sucesso, registra na lista de renomeados.

Ao final, exibe uma mensagem com resumo (renomeados e ignorados/atenção) e orienta:
“Agora clique em ‘Atualizar Sumário’ (botão em A1:B1) para atualizar a lista e os links.”

6. Regras e Restrições
Obrigatória a existência do Sumário gerado previamente; se não houver, a macro interrompe com alerta.
Não renomeia a planilha “Sumário” por regra fixa.
Sanitização de nomes:

Remove : \ / ? * [ ] e " (aspas duplas).
Trim e limite a 31 caracteres (padrão do Excel).

Nome único:

Se já existir igual, adiciona sufixo “ (2)”, “ (3)”, ….
Se o sufixo ultrapassar 31 chars, o base é encurtado para comportar “ (99)”.

Marcação de erro na coluna C:

Fundo vermelho claro (RGB(255,199,206)) e fonte vermelha (RGB(156,0,6)) quando inválido ou houve falha ao renomear.

Não altera links automaticamente após renomear; é necessário reexecutar “CriarTabelaComHiperlinks” pelo botão “Atualizar Sumário”.

7. Resultados Esperados
Planilhas renomeadas segundo os novos nomes informados em C, respeitando as regras de sanitização e unicidade.
Relatório final na mensagem com:

Renomeados: lista “• Nome antigo → Nome novo”.
Ignorados / Atenção: detalhes (vazio em C, Sumário, não encontrada, inválido/zerado, falha ao renomear).

Coluna C com marcação visual nos casos problemáticos.
Após clicar no botão “Atualizar Sumário”, o Sumário fica coerente com os nomes atualizados e os links corretos.
 
POP – Power Query “Volumetria de Lançamentos (CSV) – Agrupado” no mesmo modelo:

1. Objetivo
Importar múltiplos CSVs (prefixo Extracao_lancamentos_origens_) de uma pasta (inclui subpastas), normalizar chaves e consolidar por combinação de campos, gerando uma tabela final agrupada com:

QTD (contagem de linhas)
TOTAL (soma dos valores)

2. Abrangência
Aplicável a arquivos CSV com delimitador ; e UTF-8, contendo as colunas:
ANOCTZ_LCT, MESCTZ_LCT, CODORG_LCT_LCT, NUMFLC_LCT, CODCTA_CTB_LCT, INDDEB_CRD_LCT, VALLCT_LCT.
Funciona via Power Query (M) no Excel ou Power BI.

3. Responsável
Usuário com permissão para:

Acessar a pasta/UNC indicada.
Executar consultas no Power Query.

4. Materiais Necessários
CSVs com o layout esperado e nomes iniciando por Extracao_lancamentos_origens_.
Caminho acessível (ex.: \\fs-cse\...).
Excel/Power BI com Power Query.

5. Procedimentos
5.1 Preparação
Confirme que os CSVs possuem delimitador ;, UTF-8, e cabeçalho compatível.
Ajuste o caminho da pasta na etapa Folder.Files("...") se necessário.

5.2 Execução (Power Query)
Abra o Power Query:

Excel: Dados > Obter Dados > Consulta Nula (ou De Outros > Consulta Nula).
Power BI: Página Inicial > Nova Fonte > Consulta em Branco.

Abra o editor de fórmula (Exibir > Editor Avançado) e cole o código M fornecido.
Aplicar/Fechar & Carregar:

Excel: Fechar & Carregar para uma planilha (Tabela).
Power BI: Fechar & Aplicar.

O que a consulta faz:

Limpa chaves (LimpaChave): remove aspas, NBSP, tab, CR/LF, trim e upper.
Processa CSV: lê com Delimiter=";", Encoding=65001 (UTF-8) e QuoteStyle.Csv, promove cabeçalho e seleciona colunas.
Normaliza Valor:

Remove ponto de milhar, converte vírgula decimal para ponto.
Suporta sinal negativo no final (ex.: 1.234,56- → -1234.56).

Adiciona auditoria:

NumArquivo: número extraído do nome (após último _, dígitos).
NomeArquivoOriginal.
ChaveConcatenada: concat dos campos principais.

Filtra arquivos pelo prefixo e expande todas as linhas.
Agrupa por NomeArquivoOriginal + chaves, gerando QTD e TOTAL.

6. Regras e Restrições
Inclui subpastas (usa Folder.Files).
Somente arquivos com nome iniciando por Extracao_lancamentos_origens_ entram.
Cabeçalho obrigatório com os nomes de colunas exatamente como no código.
Valores devem estar em texto com vírgula decimal e ponto como milhar (padrão PT-BR); o código já converte.
NumArquivo depende do padrão do nome (sufixo numérico após último _); se não houver, fica null.
Resultado final não mantém a coluna Valor por linha (apenas QTD e TOTAL após o agrupamento).

7. Resultados Esperados
Tabela final consolidada por:
NomeArquivoOriginal, ANOCTZ_LCT, MESCTZ_LCT, CODORG_LCT_LCT, NUMFLC_LCT, CODCTA_CTB_LCT, INDDEB_CRD_LCT, com:

QTD: número de linhas por combinação.
TOTAL: soma de Valor normalizado.
