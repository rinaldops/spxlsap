SPXLSAP
=======

- Versão: 1.0.0
- Data: 30/03/2026

<br>

## 1. Visão Geral

### 1.1 Propósito do framework

O SPXLSAP Tool é um framework de automação em VBA que conecta dados corporativos e transações SAP a partir de um único ponto de controle. Ele combina três pilares: 

(i) execução de SQL padronizado sobre ListObjects e listas SharePoint via fwXLSPConn, fwXLConn, fwSPConn e fwHelpers; 

(ii) manipulação estruturada de tabelas Excel com fwXLTable; e 

(iii) automação assistida do SAP GUI através de fwSAPConn, fwGuiMainWindow, fwGuiMenu, fwGuiTree e fwGuiTableControl. 

Essa arquitetura permite criar pipelines que leem, transformam e sincronizam dados antes de disparar rotinas SAP, mantendo toda a lógica encapsulada no próprio Workbook.

### 1.2 Camadas funcionais

**Orquestração SQL multiorigem** - fwXLSPConn constrói dinamicamente o catálogo de ListObjects do ActiveWorkbook, mescla as fontes SharePoint recebidas por InitSP e decide, comando a comando, se o processamento ocorre localmente (fwXLConn) ou remotamente (fwSPConn), sempre apoiado pela camada de parsing de fwHelpers.

**Motor de tabelas Excel** - fwXLTable expõe operações de navegação (firstRow, GoToNextRow, RangeRow), carga (XLLoadArray), ajustes estruturais (AddColumnByName, EnsureHelpers) e formatação orientada a processos (QtdVisibleRows, ToDateSAP), permitindo que ListObjects funcionem como buffers transacionais.

**Integração SAP GUI** - fwSAPConn garante a abertura do SAP Logon, escolhe/constrói a conexão adequada e entrega um GuiSession pronto para uso. A camada de GUI wrappers encapsula interações complexas (menus dinâmicos, árvores técnicas, ALV grids) e padroniza o acesso a findById.

**Processos e utilitários** - Módulos como modUCOrquestracao, modUCAuxiliarCore e modRibbon mostram exemplos de casos de uso, auxiliando na compreensão do funcionamento do framework em um cenário de automação de processos.

### 1.3 Funcionalidades-chave

**SQL único para múltiplas fontes** - A pilha fwXLSPConn + fwHelpers interpreta SELECT/INSERT/UPDATE/DELETE, reescreve nomes físicos (RewriteDataSourceNames), aplica o dialeto ACE (SqlAlignToAce) e opera o fluxo Stage-Update-Sync descrito no item 4, garantindo consistência entre SharePoint e Excel.

**Catálogo dinâmico e parametrizável** - CreateDynamifwXLTableCatalog descobre todas as tabelas locais, enquanto ParseSPInitParams interpreta parâmetros declarativos em trios (nome, siteURL, listID), reduzindo o boilerplate de configuração.

**Manipulação rica de ListObjects** - Rotinas como XLLoadArray, RangeRow, QtdVisibleRows, AddColumnByName e ToDateSAP permitem popular tabelas, validar conteúdo, destacar linhas em processamento e ajustar formatos exigidos por integrações SAP sem escrever código repetitivo.

**Sincronização SharePoint dirigida por linha** - Os fluxos de orquestração do módulo modUCOrquestracao demonstram o uso de RunQuery para atualização dos registros visíveis/selecionados, registrando o resultado na coluna OBSERVAÇÕES e mantendo feedback imediato ao usuário.

**Sessões SAP resilientes** - EnsureSapLogonRunning, PickOrOpenConnection, EnsureWantedSession e PerformInteractiveLogon lidam com abertura de cliente, múltiplas janelas e recuperação após quedas, enquanto a Ribbon customizada exibe o estado da conexão em tempo real.

**Wrappers reutilizáveis de GUI** - fwGuiMainWindow, fwGuiMenu, fwGuiTree e fwGuiTableControl encapsulam padrões comuns (busca híbrida por elementos, leitura de ALV, interação com menus contextuais), simplificando a escrita de rotinas SAP específicas como a existente em btnProcessar_Click.

### 1.4 Benefícios para automação

Reduz o acoplamento entre fontes de dados ao expor um SQL único independente da origem física.

Minimiza riscos operacionais com staging local, logs (TraceOn nas classes de dados e wrappers principais) e monitoramento automático da sessão SAP.

Padroniza a experiência do desenvolvedor: um único catálogo, objetos de tabela navegáveis e wrappers GUI diminuem a curva de aprendizado.

Facilita compliance (rótulos de confidencialidade, controle de acesso via SharePoint) sem inserir essas preocupações nos scripts de negócio.

Permite evoluir o framework de forma incremental, pois novas integrações podem usar os mesmos conectores e helpers já consolidados.

<br>

<br>

---

## 2. Guia Rápido para Iniciar

### 2.1 Pré-requisitos de ambiente

**Office e SAP GUI** - Excel 2019/Office 365 com macros habilitadas e SAP GUI 7.70+ com scripting liberado (perfil Z:BC_PB001_USUARIO_EXEC_SCRIPT).

**Conectividade corporativa** - Acesso VPN, permissões de leitura/escrita às listas SharePoint e direito de executar SAPExecuteCommand nas queries BW.

**Workbook confiável** - Arquivo salvo em local listado no Trust Center; pVariaveis preenchida com URLs, GUIDs, datas, flags (chkMaster, chkAtualizaSP) e credenciais de BW.

**Dependências internas** - Add-ins SAP instalados, Power Query configurado e planilhas auxiliares (SITOP_BW, Origem/Destino etc.) atualizadas pelo último CarregaPlanilhasBW.

### 2.2 Troubleshooting essencial

**SAP scripting desabilitado** - Verifique o perfil de segurança, confirme que saplogon.exe está aberto e use initSAP/fwSAPConn.Connect (que invoca EnsureSapLogonRunning) antes de repetir a macro.

**Erros ACE/IMEX** - Salve o workbook antes de DML, assegure que colunas-chave estejam formatadas como número; o Stage-Update-Sync aplica NumberFormat = "0", mas colunas manuais podem exigir ajuste.

**Lista/URL inválida** - Revise pVariaveis e use orquestrador.InitSP Nome, URL, GUID para registrar trios corretos; mensagens “Fonte não encontrada” indicam GUID ou apelido divergente.

**Registro bloqueado** - Macros PAGESP exibem “Usuário sem permissão” ou “Registro bloqueado” em OBSERVAÇÕES. Ajuste a coluna Acessos ou o Ponto Focal na tabela/SharePoint e reexecute o lote.

**Power Query não atualiza** - Execute RefreshTablesInSequence para forçar QueryTable.Refresh, cheque credenciais armazenadas e monitore logs (TraceOn) para encontrar planilhas com falha.

<br>

<br>

---

## 3. Arquitetura e Conceitos Fundamentais

### 3.1 Componentes principais

**fwXLSPConn (orquestrador SQL)** - Recebe o catálogo consolidado, interpreta o comando e delega para o conector adequado. Também mantém o dicionário de tabelas temporárias e o ciclo Stage-Update-Sync.

**fwXLConn (motor Excel)** - Traduz o SQL para o dialeto ACE, abre a conexão ADO contra o ActiveWorkbook e retorna dados ou rowsAffected. Suporta modo leitura/escrita dinâmico (IMEX=1/0).

**fwSPConn (motor SharePoint)** - Abstrai a conexão OLEDB WSS, controla o campo chave (SPKeyField) e executa DML em listas.

**fwHelpers** - Centraliza parsing de SQL, normalização de nomes, geração de catálogos (CreateDynamifwXLTableCatalog) e relatórios de execução (FillExecReport).
fwXLTable. Expõe uma API de manipulação de ListObjects, usada tanto por macros de negócio quanto pelo fluxo de staging para carregar e navegar em lotes.

**fwSAPConn e wrappers GUI** - fwSAPConn garante sessões vivas enquanto fwGuiMainWindow, fwGuiMenu, fwGuiTree e fwGuiTableControl encapsulam interações avançadas na interface SAP.

**Módulos de apoio** - modUCOrquestracao, modUCAuxiliarCore, modRibbon e fwSapGuiApi integram a solução com a Ribbon, cuidam de exportações ALV (via rotinas UC) e expõem funções reutilizáveis.

### 3.2 Padrões operacionais

Catálogo unificado. Todo comando passa por um dicionário único que descreve onde cada entidade reside (XL ou SP), permitindo renomear fontes sem tocar nos SQLs. Esse catálogo é criado automaticamente ao inicializar a classe fwXLSPConn corretamente.

**Stage-First** - Consultas híbridas trazem dados do SharePoint para tabelas ocultas (Staging_fwXLSPConn) antes de executar JOINs, garantindo performance e consistência.

**Execução adaptativa** - O orquestrador identifica se o comando é leitura ou escrita e comuta automaticamente conexão, dialeto e locking para cada origem.
Observabilidade integrada. TraceOn propaga logs para todas as classes e os helpers registram tempos, SQL final e métricas no execReport, facilitando troubleshooting.
Manutenção de contexto SAP. As rotinas de automação consultam pVariaveis, validam a sessão (IsSessionAlive) e reaproveitam a mesma conexão SAP para comandos em lote.

### 3.3 Ecossistema de dados e SAP

O desenho modular permite misturar interações de dados (SQL, manipulação de ListObjects, atualizações SharePoint) com automações SAP no mesmo fluxo. Uma rotina pode extrair dados via RunQuery, pintá-los com fwXLTable para revisão humana, disparar transações SAP com fwGuiMainWindow.TCode e, ao final, sincronizar status no SharePoint — tudo sem sair do workbook.

<br>

<br>

---

## 4. Mecanismo de Staging do Framework

### 4.1 Visão geral

O SPXLSAP opera em camadas para decidir quando rodar consultas diretamente na fonte e quando criar cópias temporárias em Excel. O orquestrador fwXLSPConn avalia cada SQL, identifica as origens envolvidas e escolhe entre rotas Direct-XL, Direct-SP ou Stage-First. O objetivo é garantir desempenho, preservar a integridade das listas SharePoint e permitir consultas federadas que misturam fontes heterogêneas.

### 4.1.1 Contexto Stage-Update-Sync
Utilizado quando comandos DML precisam combinar dados de múltiplas origens (por exemplo, UPDATE em uma lista SharePoint baseado em filtros de uma tabela Excel). O orquestrador fwXLSPConn cuida automaticamente dos estágios necessários para preservar consistência entre Excel e SharePoint.

### 4.1.2 Etapas do pipeline
- Staging: identifica as fontes envolvidas, executa SELECT em cada lista SharePoint e materializa cópias locais (XL_PAGESPHSP, XL_PAGESPSP) em planilhas ocultas.
- Simulação: executa um SELECT rápido sobre as tabelas de staging para descobrir exatamente quais IDs serão afetados e guarda essa lista em memória.
- Execução local: processa o comando DML original (por exemplo, UPDATE ... JOIN ...) apenas nas tabelas temporárias, evitando round-trips com o SharePoint.
- Coleta: faz um SELECT * FROM [TabelaStage] WHERE [ID] IN (...) para capturar os registros alterados.
- Sincronização: envia o conjunto resultante para SPUpdateRows, que aplica as mudanças na lista SharePoint em lotes seguros.

### 4.1.3 Analogia operacional
É como um chef: traz ingredientes para a bancada (staging), separa o que será usado (simulação), prepara os pratos (execução), monta os pedidos (coleta) e entrega aos clientes (sincronização). O desenvolvedor apenas dispara o comando SQL; o framework cuida do restante.

### 4.2 Operações com SELECT
Premissa: tanto o driver WSS (SharePoint) quanto o ACE (Excel) suportam SELECT com self-join e joins convencionais. O quadro abaixo mostra como o framework explora isso.

### 4.2.1 Cenário 1 — Direct-XL (apenas tabelas Excel)
Resumo rápido: consultas restritas a ListObjects residentes no workbook são encaminhadas diretamente para fwXLConn, sem qualquer staging. Ideais para dashboards e análises locais.

**Passo a passo**
Análise e roteamento. O orquestrador identifica que não há listas SharePoint na consulta (spTablesInQuery.Count = 0) e escolhe a estratégia Direct-XL.
Preparação e tradução. O SQL é repassado para fwXLConn, que executa RewriteDataSourceNames para converter nomes lógicos (ex.: [Clientes]) em nomes físicos ([XL_Clientes]).
Execução direta. fwXLConn abre o ACE em modo leitura (IMEX=1) e envia o SQL completo ao driver, que resolve qualquer combinação de self-join ou múltiplos joins.
Retorno dos dados. O Recordset é convertido em array e devolvido ao chamador; nenhuma tabela de staging é criada.

### 4.2.2 Cenário 2 — Direct-SP (uma única lista SharePoint)
Resumo rápido: quando a consulta usa apenas uma lista SharePoint, inclusive em self-joins, fwSPConn trata tudo remotamente. É a rota mais rápida para inventários e auditorias direto no servidor.

**Passo a passo**
Análise e roteamento. spListsInQuery.Count <= 1 e hasXLTable = False, então o orquestrador ativa o Direct-SP.
Execução remota. fwSPConn abre uma conexão WSS em modo leitura (IMEX=1) e executa o SQL original, aproveitando o suporte do driver a self-join.
Retorno dos dados. O resultado é convertido para array e retornado; não há staging.

### 4.2.3 Cenário 3 — Consultas federadas (combinação XL/SP)
Resumo rápido: sempre que o SQL mistura múltiplas listas SharePoint ou combina SP com Excel, entra em ação o pipeline Stage-First. Ele baixa apenas as colunas necessárias, cria tabelas temporárias e executa o JOIN localmente.

**Passo a passo**
Análise e roteamento. Se spTablesInQuery.Count > 1 ou existe mistura SP + XL, o orquestrador ativa a estratégia de staging.
Stage (download). Para cada lista SP, BuildStagingSqlFor gera um SELECT otimizado (apenas colunas usadas + ID). Os dados são materializados em tabelas ocultas no Excel.
Execução local. O SQL original é traduzido (todos os nomes viram [XL_...]) e executado pela fwXLConn, que une dados locais e staged com alto desempenho.
Retorno e limpeza. O resultado é entregue ao usuário e CleanupTempTables remove as tabelas temporárias.

### 4.2.4 Diferenças principais
- Direct-XL: todo o trabalho fica com o ACE/OLEDB (Excel).
- Direct-SP: todo o trabalho fica com o WSS/OLEDB (SharePoint).
- Stage-First: envolve WSS para download e ACE para processar o JOIN local.

### 4.3 Operações com UPDATE
Premissas: drivers ACE e WSS suportam self-join em UPDATE. Quando necessário, o framework reaproveita a infraestrutura de SELECT para planejar atualizações seguras.

### 4.3.1 Cenário 1A — UPDATE Direct-XL (com ou sem self-join)
Resumo rápido: usado quando todas as tabelas do UPDATE são Excel. Permite atualizar grandes volumes em lote com o ACE trabalhando diretamente no workbook.
Análise inteligente. ExecuteUpdate detecta JOIN, confirma que todas as tabelas são do tipo XL e usa o caminho DirectExcelUpdatePath sem staging.
Tradução. O SQL completo é repassado para fwXLConn, que traduz os nomes lógicos e abre a conexão em modo escrita (IMEX=0).
Execução. O driver ACE processa o UPDATE JOIN, resolve o filtro e aplica os valores da cláusula SET diretamente.
- Retorno. O número de linhas afetadas é capturado e devolvido ao orquestrador.

### 4.3.2 Cenário 1B - UPDATE em SharePoint (estratégia iterativa)
Resumo rápido: quando o UPDATE atinge lista SharePoint, o caminho atual prioriza robustez via estratégia iterativa (Consultar-Iterar-Executar), mantendo a operação controlada no servidor SP.
Roteamento. ExecuteUpdate identifica alvo SP e encaminha para fwSPConn.SPUpdateRowsByQuery.
Preparação. A rotina deriva e executa uma pré-consulta para identificar os IDs impactados.
Execução. Para cada ID, aplica UPDATE simples no SharePoint com controle de progresso.
- Retorno. O número de linhas afetadas retorna ao usuário.

### 4.3.3 Cenário 2 — UPDATE com self-join em SharePoint (Consultar-Iterar-Executar)
Resumo rápido: quando há maior risco no servidor SP, a estratégia SPUpdateRowsByQuery consulta previamente os IDs e executa atualizações individuais com barra de progresso.
Delegação. ExecuteUpdate envia o SQL para SPUpdateRowsByQuery, que o quebra em tablesClause, setClause e whereClause.
Pré-consulta de IDs. Um SELECT derivado recupera todos os IDs impactados via SPRunQuery.
- Loop iterativo. Para cada ID, monta-se um UPDATE simples (sem alias), garantindo máxima compatibilidade.
Finalização. O total de linhas atualizadas é retornado à RunQuery, que atualiza o status visual.

### 4.3.4 Cenário 3 — UPDATE com JOIN e Staging (SP-XL ou SP-SP)
Resumo rápido: combina Stage-Update-Sync. Baixa listas SP, alinha formatos, executa o UPDATE localmente e sincroniza apenas os registros afetados.
- Stage. BuildStagingSqlFor baixa somente as colunas necessárias de cada lista SP e cria tabelas ocultas.
Formatação. O framework aplica NumberFormat = "0" nas colunas de junção para evitar erros de tipo no ACE.
- Update local. O UPDATE JOIN é traduzido para nomes físicos [XL_] e executado pelo ACE. Se a tabela alvo for SP, um SELECT adicional captura os IDs alterados.
- Sync. Caso o alvo seja SharePoint, SPUpdateRows envia os dados atualizados em lotes seguros.

### 4.3.5 Diferenças principais
- Direct-XL: um único comando enviado ao driver ACE. Em SP, a implementação privilegia o modo iterativo para maior estabilidade.
Self-join SP iterativo: divide em SELECT + múltiplos UPDATE simples, com barra de progresso.
Stage-Update-Sync: mistura processamento local rápido com sincronização controlada no SharePoint.

### 4.4 Operações com INSERT
Premissa: o parser identifica automaticamente se o comando usa VALUES ou SELECT e delega para a estratégia correta.

### 4.4.1 Cenário 1 — INSERT ... VALUES
Resumo rápido: ideal para cadastros pontuais. O comando completo é enviado para o conector da origem alvo.
Análise. ParseInsertSqlDetails encontra VALUES e identifica a tabela destino.
Execução direta.
Destino XL: fwXLConn abre conexão em modo escrita (IMEX=0) e executa o INSERT via ACE.
Destino SP: fwSPConn abre o WSS e executa o comando diretamente.
- Retorno. Quantidade de linhas inseridas retorna ao usuário.

### 4.4.2 Cenário 2 — INSERT ... SELECT
Resumo rápido: permite copiar dados entre fontes heterogêneas. O SELECT reutiliza todo o pipeline descrito em 3.2 antes de carregar o destino.
Separação. O parser divide o comando nas partes INSERT (destino e colunas) e SELECT (origem dos dados).
Execução da subconsulta. A subconsulta é tratada por ExecuteSelect, podendo usar Direct-XL, Direct-SP ou staging.
- Carga no destino.
Destino SP: fwSPConn.SPInsertRows insere linha a linha com barra de progresso.
Destino XL: fwXLTable.XLLoadArray adiciona as linhas em lote.
- Retorno. O total de registros inseridos é informado.

### 4.5 Operações com DELETE
Premissa: ParseDeleteSqlDetails identifica alvo e filtros; a escolha da estratégia evita limitações dos drivers com exclusões em massa.

### 4.5.1 Cenário 1 — DELETE em tabelas Excel
Resumo rápido: devido às restrições do ACE, o framework usa a estratégia "Consultar e Deletar" dentro do Excel, preservando filtros do usuário.
Delegação. O orquestrador chama fwXLTable.DeleteRowsByQuery.
Preparação. O estado dos filtros é salvo e, se necessário, cria-se uma coluna ID temporária.
- Consulta de IDs. O DELETE vira SELECT [ID] ..., executado via fwXLConn em modo leitura.
Exclusão programática. A tabela é percorrida de baixo para cima e linhas com IDs encontrados são removidas com ListRow.Delete (barra de progresso incluída).
- Limpeza. Filtros e colunas temporárias são restaurados.

### 4.5.2 Cenário 2 — DELETE em listas SharePoint
Resumo rápido: usa a estratégia "Consultar-Iterar-Executar" para manter o SharePoint estável e informar progresso.
Delegação. O SQL é enviado para fwSPConn.SPDeleteRowsByQuery.
- Consulta. O comando é convertido em SELECT [ID] ... e executado via SPRunQuery em modo leitura.
Iteração. Com a lista de IDs, abre-se uma conexão de escrita e executam-se DELETE simples por ID, atualizando a barra de progresso.
- Retorno. Total de itens excluídos é devolvido ao usuário.

<br>

<br>

---

## 5. Configuração e Instanciamento

### 5.1 Descoberta automática de ListObjects
O catálogo padrão do SPXLSAP é inteiramente dinâmico. A cada chamada de fwXLSPConn.Init, o helper CreateDynamifwXLTableCatalog(ActiveWorkbook) enumera todos os ListObjects existentes no workbook, aplica vbTextCompare para evitar conflitos de caixa e registra cada tabela com o mesmo nome visível na planilha.
Nenhum dicionário precisa ser montado manualmente; basta nomear os ListObjects de forma consistente e eles estarão disponíveis nos comandos SQL.
Alterações feitas nas planilhas (adição, renomeação ou exclusão de tabelas) são detectadas automaticamente sempre que o orquestrador é inicializado.
Caso seja necessário criar aliases, renomeie o próprio ListObject ou utilize fwXLTable para expor colunas vistas pelo usuário antes de orquestrar consultas.

### 5.2 Inicialização declarativa das listas SharePoint
O orquestrador expõe um ParamArray para registrar as listas do SharePoint em trios NomeLógico, URL e GUID da lista Sharepoint, sem a criação de dicionários auxiliares. A assinatura é:
vb orquestrador.InitSP Nome1, URL1, GUID1, Nome2, URL2, GUID2 ...

**Exemplo prático:**

`vb Dim orch As New fwXLSPConn
orch.InitSP _ "FuncionariosSP", "https://empresa.sharepoint.com/sites/ti/Lists/FuncionariosSP", "181bdab9-68db-416a-b9c3-1ffea141ac67", _ "OrdensManutencao", "https://empresa.sharepoint.com/sites/ti/Lists/OrdensManutencao", "64f4041a-3ee2-4aa6-9d4e-a0d4e1dc3d46" `

**Regras de uso:**
Cada trio representa o nome lógico, a URL e o GUID interno da lista SharePoint que será referenciada no SQL.
Informe quantos trios forem necessários; o orquestrador valida se a quantidade de argumentos é múltipla de 3 e ignora valores vazios.
TraceOn continua disponível para inspecionar o catálogo final e conferir se todas as listas foram registradas corretamente.
InitSP é a API pública recomendada para novos desenvolvimentos com assinatura declarativa em trios; Init permanece interno à classe.

### 5.3 Boas práticas de inicialização
Centralize variáveis de ambiente em pVariaveis e valide-as antes de chamar Init/InitSP.
Reutilize instâncias de fwXLTable para navegar pelas linhas a processar nos fluxos UC (por exemplo, extração e processamento).
Mantenha TraceOn apenas durante diagnóstico; em produção use os retornos execReport e OBSERVAÇÕES nas tabelas.
Antes de automações SAP, invoque initSAP (ou fwSAPConn.Connect) e mantenha a sessão disponível para os wrappers GUI.

<br>

<br>

---

## 6. Premissas e Sintaxe SQL

### 6.1 Convenções gerais
Nomes únicos. Cada fonte registrada no catálogo deve ter nome exclusivo para evitar sobreposições entre tabelas Excel e listas SharePoint.
Uso de colchetes. Delimite sempre tabelas e campos com [ ], garantindo compatibilidade com o reescritor de nomes.
Sem colchetes nos títulos físicos. ListObjects e listas não devem conter [ ou ] nos cabeçalhos originais.
Cor curingas. Utilize * em cláusulas LIKE (ACE não aceita %).
Separação de responsabilidades. Evite lógica de negócio em SQL; mantenha-a nos módulos VBA e use os comandos apenas para projeções e filtros.

### 6.2 Regras específicas por origem
Excel (ACE OLEDB). Converta textos numéricos com CDbl() quando necessário, salve o workbook antes de operações DML e prefira XLLoadArray para popular ListObjects.
SharePoint. Garanta que SiteURL e ListID estejam corretos e que o usuário possua permissões de escrita. Use as colunas internas (ID, Title, etc.) respeitando SPKeyField.
Tabelas de staging. O orquestrador cria planilhas ocultas (Staging_fwXLSPConn) para trabalhar offline; não as delete manualmente enquanto o processo estiver em execução.

### 6.3 Recursos de validação e apoio
fwHelpers oferece parsing e alinhamento automático (SqlAlignToAce, RewriteDataSourceNames, IsSelectSql), evitando divergências de dialeto. Em caso de falhas, consulte o trace do orquestrador e o relatório preenchido por FillExecReport, que indica SQL final, tempo e quantidade de linhas afetadas.

<br>

<br>

---

## 7. Comandos Suportados e Exemplos Práticos

### 7.1 Modelos de dados de referência

**Lista SharePoint `[FuncionariosSP]`**

| ID | Nome | Cargo | ID_Departamento | Salário |
|---|---|---|---|---|
| 1 | Ana Silva | Analista | 10 | 5000 |
| 2 | Bruno Costa | Gerente | 10 | 8000 |
| 3 | Carlos Dias | Analista | 20 | 5500 |
| 4 | Diana Souza | Estagiária | 20 | 2000 |
| 5 | Eva Mendes | Diretora | 30 | 15000 |

**Tabela Excel `[Departamentos]`**

| ID_Dep | NomeDepartamento | Localizacao |
|---|---|---|
| 10 | TI | Prédio A |
| 20 | RH | Prédio B |
| 30 | Diretoria | Prédio A |

### 7.2 SELECT
Os SELECTs aceitam JOIN, WHERE, GROUP BY, HAVING, ORDER BY e podem ser executados diretamente no SharePoint/Excel ou via staging automático. Use Result("Data") + XLLoadArray para carregar o retorno em um ListObject.

**Exemplo 1 – Consulta simples (Direct-XL)**

SELECT [NomeDepartamento], [Localizacao]
FROM [Departamentos]
WHERE [Localizacao] = 'Prédio A'
- Resultado: retorna os departamentos de TI e Diretoria.

**Exemplo 2 – Agregação (Direct-SP)**

SELECT [ID_Departamento], AVG([Salario]) AS [MediaSalarial]
FROM [FuncionariosSP]
GROUP BY [ID_Departamento]
HAVING AVG([Salario]) > 6000
- Resultado: exibe a média salarial dos departamentos 10 e 30 diretamente no SharePoint.

**Exemplo 3 – JOIN híbrido (Stage-First)**

SELECT f.[Nome], f.[Cargo], d.[NomeDepartamento]
FROM [FuncionariosSP] AS f
INNER JOIN [Departamentos] AS d ON f.[ID_Departamento] = d.[ID_Dep]
WHERE d.[Localizacao] = 'Prédio A'
ORDER BY f.[Nome]
- Resultado: o orquestrador baixa os dados de [FuncionariosSP], cria tabelas temporárias no Excel e processa o JOIN localmente, retornando Ana Silva, Bruno Costa e Eva Mendes.

### 7.3 INSERT
Suporta INSERT ... VALUES e INSERT ... SELECT, inclusive com JOIN entre fontes. Utilize rowsAffected para validar o processamento e, quando necessário, use fwXLTable.RangeRow("OBSERVAÇÕES") para registrar feedback ao usuário.

**Exemplo 4 – Inserção direta em Excel**

INSERT INTO [Departamentos] ([ID_Dep], [NomeDepartamento], [Localizacao])
VALUES (40, 'Marketing', 'Prédio B')
- Resultado: cria o departamento de Marketing no ListObject.

**Exemplo 5 – Inserção com SELECT e JOIN**

INSERT INTO [RelatorioCusto] ([Departamento], [CustoTotal])
SELECT d.[NomeDepartamento], SUM(f.[Salario])
FROM [FuncionariosSP] AS f
INNER JOIN [Departamentos] AS d ON f.[ID_Departamento] = d.[ID_Dep]
GROUP BY d.[NomeDepartamento]
- Resultado: popula a tabela [RelatorioCusto] no Excel com o custo total por departamento, permitindo uso posterior por macros SAP.

### 7.4 DELETE
Aceita qualquer cláusula WHERE válida e respeita o estágio atual (Excel ou SharePoint).

**Exemplo 6 – Exclusão em SharePoint**

DELETE FROM [FuncionariosSP]
WHERE [Cargo] = 'Estagiária'
- Resultado: remove Diana Souza da lista SharePoint e retorna rowsAffected = 1.

**Exemplo 7 – Exclusão em Excel**

DELETE FROM [Departamentos]
WHERE [ID_Dep] = 40
- Resultado: elimina o departamento de Marketing previamente inserido.

### 7.5 UPDATE
Atualiza registros com base em WHERE. Quando o comando envolve múltiplas fontes, o framework aplica o pipeline Stage-Update-Sync para garantir consistência.

**Exemplo 8 – Atualização em Excel**

UPDATE [Departamentos]
SET [Localizacao] = 'Sede'
WHERE [ID_Dep] = 30
- Resultado: define a localização da Diretoria como “Sede”.

**Exemplo 9 – Atualização em SharePoint**

UPDATE [FuncionariosSP]
SET [Salario] = [Salario] * 1.10
WHERE [ID_Departamento] = 10
- Resultado: aumenta em 10% o salário dos funcionários do departamento de TI.

<br>

<br>

---

## 8. Elementos Coadjuvantes
Os itens a seguir não fazem parte do framework, porém foram implementados e introduzidos na solução para que fosse possível executar um caso de uso com necessidades recorrentes que demandam recursos como os que o framework é capaz de fornecer.
Tratamento de filas SharePoint <-> Excel. Fluxos UC carregam listas SharePoint para ListObjects configurados; os analistas tratam os registros e o processamento sincroniza as alterações linha a linha.
Preparação de dados para SAP. O framework aplica filtros, normaliza formatos (por exemplo, ToDateSAP) e escreve os dados já validados que serão consumidos por scripts SAP, evitando retrabalho na GUI.
Automação de relatórios SAP. Rotinas UC combinadas à exportação ALV automatizam extrações, criam abas dedicadas e rotulam arquivos exportados conforme políticas da companhia.
Execução assistida de transações. btnProcessar_Click e PERFIL mostram como chamar transações (TCode), alimentar campos e navegar em controles guiados pelos wrappers, mantendo o mesmo código-fonte independentemente da tela ativa.

### 8.1 Guia de uso resiliente dos wrappers GUI (métodos booleanos)
Para rotinas SAP com telas dinâmicas, priorize os métodos booleanos. Eles retornam `True/False` e evitam exceções desnecessárias no fluxo principal.

**Padrão recomendado**

- Use métodos booleanos quando a existência do elemento depende de contexto (menu, popup, coluna visível, nó de árvore).
- Trate `False` com mensagem curta de status/log e siga para o fallback da rotina.
- Reuse wrapper quando repetir a mesma ação em loop, evitando criação repetida do mesmo objeto.

**Menus (classe `fwGuiMenu`, via factory `GuiMenu(...)`)**

```vb

If Not w.GuiMenu("Linhas", "Visualizar as linhas ocultadas").SelectMenu() Then
    Application.StatusBar = "Menu não disponível no contexto atual."
End If

If w.GuiMenu("Objetos eliminados", "Exibir").ExistsMenu() Then
    Call w.GuiMenu("Objetos eliminados", "Exibir").SelectMenu()
End If

```

```vb

Dim m As fwGuiMenu
Set m = w.GuiMenu("Opções", "Configurações")
If Not m.SelectMenu() Then
    Application.StatusBar = "Falha ao selecionar Opções > Configurações."
End If

```

<br>

**Árvore (`fwGuiTree`)**

```vb

If Not w.GuiTree("shell").SelectFirstNodeByItemText("STATUS", "REL", "&Hierarchy", False) Then
    Application.StatusBar = "Nó não encontrado na árvore."
End If

```

```vb

If Not w.GuiTree("shell").SelectNodeByTechKey("000000000000123456") Then
    Application.StatusBar = "TECH_KEY não localizada."
End If

```

<br>

**Tabela/ALV (`fwGuiTableControl`)**

```vb

Dim outText As String
If w.GuiTable("tblSAPLCOVGTC_CONTROL").GetCellTextByColumnText("AUFNR", "1001421729", "VORNR", outText, True) Then
    Debug.Print outText
Else
    Application.StatusBar = "Linha/coluna não encontrada na grade."
End If

```

```vb

If Not w.GuiTable("tblSAPLCOVGTC_CONTROL").SetCellTextByColumnText("AUFNR", "1001421729", "STEUS", "PS01", True) Then
    Application.StatusBar = "Não foi possível atualizar a célula alvo."
End If

```

<br>

**Observação de eficiência**
- `w.GuiMenu("MenuPai","OpcaoFilha")` pode ser chamado inline, mas em chamadas repetidas no mesmo bloco prefira armazenar a instância (`Set m = ...`) e reutilizar.

**Quando usar `GuiTree/GuiTable(...)` vs `oGuiTree/oGuiTable`**
- `GuiTree("idOuNome")` e `GuiTable("idOuNome")` são a opção padrão recomendada para código de processo: já devolvem o wrapper apontando para um controle específico.
- `oGuiTree` e `oGuiTable` são wrappers compartilhados da janela; use quando quiser manter uma única instância e configurar `ID` ou `Name` manualmente antes das chamadas.

```vb

' Recomendado (direto e explícito)
If Not w.GuiTree("shell").SelectNodeByTechKey("000000000000123456") Then
    Application.StatusBar = "Nó não localizado."
End If

```

```vb

' Avançado (wrapper compartilhado com bind manual)
With w.oGuiTree
    .Name = "shell"
    Call .SelectNodeByTechKey("000000000000123456")
End With

```

<br>

<br>

---

## 9. Conclusão
O SPXLSAP Tool unifica acesso a dados e automação SAP em uma única base VBA. O catálogo dinâmico, o pipeline Stage-Update-Sync, os conectores especializados e os wrappers de GUI permitem que desenvolvedores foquem em regras de negócio sem se preocupar com infraestrutura de conectividade. Com suporte completo a operações CRUD, rastreamento integrado e componentes prontos para manipular ListObjects e SAP GUI, o framework sustenta soluções de automação robustas, auditáveis e alinhadas às políticas corporativas.

<br>

---

## Anexo A. Inventário de Arquivos

| Arquivo | Descrição |
|---|---|
| `src/fwXLSPConn.cls` | Orquestrador central; combina o catálogo dinâmico de ListObjects com as listas SharePoint declaradas e executa o pipeline Stage-Update-Sync. |
| `src/fwXLConn.cls` | Motor OLEDB para Excel; traduz SQL para o dialeto ACE, controla conexões ADO e retorna dados ou linhas afetadas. |
| `src/fwSPConn.cls` | Conector SharePoint via ACE/WSS; mantém SiteURL/ListID, campo chave e executa DML diretamente nas listas. |
| `src/fwXLTable.cls` | Wrapper de ListObjects; oferece navegação, carga de arrays, manipulação estrutural e utilitários como RangeRow e QtdVisibleRows. |
| `src/fwHelpers.cls` | Biblioteca de suporte; faz parsing de SQL, reescreve nomes físicos, gera catálogos dinâmicos e monta relatórios de execução. |
| `src/fwSAPConn.cls` | Abre e mantém sessões SAP GUI; garante scripting habilitado, escolhe conexões e retorna GuiSession pronto. |
| `src/fwGuiMainWindow.cls` | Encapsula janelas SAP (wnd[n]); executa TCodes, envia VKeys e expõe buscas generalistas por controles. |
| `src/fwGuiMenu.cls` | Wrapper de menus SAP; resolve itens por texto, navega em hierarquias e executa comandos contextuais. |
| `src/fwGuiTableControl.cls` | Abstrai controles do tipo tabela/ALV; provê rolagem, seleção e leitura de células para automações. |
| `src/fwGuiTree.cls` | Manipula árvores SAP GUI; localiza nós, expande estruturas e interage com registros hierárquicos. |
| `src/fwSapGuiApi.bas` | Declarações e helpers da API SAP GUI Scripting, facilitando chamadas late binding. |
| `README.md` | Este documento consolida visão geral, arquitetura, comandos SQL e fluxos do framework. |
