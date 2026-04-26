# SPXLSAP

O **SPXLSAP** é um framework VBA para construir automações corporativas que precisam conversar com **Excel**, **SharePoint** e **SAP GUI** sem transformar cada novo projeto em uma coleção de macros soltas, difíceis de entender e mais difíceis ainda de manter.

A proposta é bem prática: deixar o desenvolvedor escrever a regra de negócio com menos atrito. Em vez de recomeçar do zero sempre que precisar ler uma tabela do Excel, consultar uma lista SharePoint, atualizar registros em lote, navegar por uma tela do SAP, localizar uma tabela ALV ou controlar uma sessão SAP aberta, o framework oferece classes prontas para esses trabalhos comuns. Assim, o código do projeto fica mais focado no processo que está sendo automatizado e menos na infraestrutura repetitiva que todo mundo acaba reescrevendo.

Esse ponto é importante principalmente para quem está começando. O SPXLSAP não tenta esconder que automação corporativa envolve assuntos técnicos: ADODB, ACE/OLEDB, SharePoint, ListObjects, SAP GUI Scripting, sessões SAP, IDs técnicos de tela, comandos SQL, tratamento de filtros e sincronização de dados. Tudo isso continua existindo. A diferença é que o framework organiza esses temas em camadas mais previsíveis, com nomes claros, rotas padronizadas e exemplos que ajudam o desenvolvedor a aprender enquanto constrói.

## Para que o framework existe

Grande parte das automações feitas em Excel nasce de uma necessidade simples: alguém tem uma massa de dados em uma planilha, precisa consultar ou atualizar informações em algum sistema corporativo e quer reduzir trabalho manual. O problema aparece quando a solução cresce. A macro que antes só clicava em uma tela do SAP passa a depender de filtros, tabelas auxiliares, listas SharePoint, consultas SQL, validações, mensagens de retorno, logs, permissões e várias exceções de processo.

Sem uma base comum, cada automação resolve esses problemas de um jeito diferente. Um projeto abre a conexão SAP de uma forma, outro percorre tabela do Excel de outra, outro monta SQL por concatenação, outro acessa SharePoint diretamente dentro da macro principal. Funciona por algum tempo, mas a manutenção fica pesada. Pequenas mudanças viram risco, o reaproveitamento é baixo e quem chega depois precisa descobrir tudo no escuro.

O SPXLSAP entra justamente nesse espaço. Ele oferece uma arquitetura de apoio para que novas soluções sejam construídas com mais rapidez, mas sem abrir mão de robustez. O desenvolvedor continua tendo liberdade para desenhar o fluxo do seu processo, mas passa a contar com conectores, wrappers e helpers que já tratam vários detalhes desagradáveis do caminho.

Na prática, o framework ajuda você a:

- trabalhar com tabelas estruturadas do Excel como se fossem bases de dados locais;
- executar consultas e comandos SQL sobre fontes Excel e SharePoint;
- sincronizar dados entre Excel e listas SharePoint com menos código manual;
- automatizar telas SAP usando objetos mais amigáveis do que `session.findById` espalhado por todo lado;
- manter a sessão SAP viva e reaproveitável durante o processamento;
- criar rotinas que registram resultado, progresso e mensagens de erro de forma mais padronizada;
- separar melhor o que é infraestrutura do framework e o que é regra de negócio da sua solução.

## Como pensar no SPXLSAP

Um jeito simples de entender o SPXLSAP é imaginar três grandes blocos trabalhando juntos.

O primeiro bloco é o de **dados**. Ele cuida das fontes que alimentam ou recebem informações: tabelas do Excel, listas SharePoint e consultas feitas por SQL. Nesse bloco entram principalmente `fwXLSPConn`, `fwXLConn`, `fwSPConn`, `fwXLTable` e `fwHelpers`.

O segundo bloco é o de **SAP GUI**. Ele cuida da conexão com o SAP, da janela ativa, dos campos, menus, árvores, tabelas e elementos que aparecem na tela. Nesse bloco entram `fwSAPConn`, `fwGuiMainWindow`, `fwGuiMenu`, `fwGuiTree`, `fwGuiTableControl` e `fwSapGuiApi`.

O terceiro bloco é o de **orquestração do processo**. Esse não é uma classe única do framework, porque depende de cada automação. É o código que você escreve para dizer: primeiro busque dados, depois filtre as linhas visíveis, depois abra a transação SAP, depois preencha os campos, depois registre o resultado no Excel e, se for o caso, sincronize no SharePoint. O framework não toma a regra de negócio da sua mão; ele só entrega ferramentas melhores para você escrever essa regra.

Essa separação ajuda muito no aprendizado. Um iniciante pode começar usando `fwXLTable` para navegar por uma tabela do Excel. Depois pode usar `fwSAPConn` para conectar no SAP. Em seguida pode experimentar `fwGuiMainWindow.TCode` para abrir uma transação. Mais adiante, pode usar `fwXLSPConn.RunQuery` para misturar Excel e SharePoint com SQL. Não é necessário dominar tudo no primeiro dia.

## Arquitetura em camadas

O SPXLSAP foi organizado em camadas para evitar que uma macro de negócio tenha que conhecer todos os detalhes técnicos ao mesmo tempo. Cada classe tem uma responsabilidade mais clara, e isso facilita tanto o uso quanto a manutenção.

### Camada de orquestração SQL

A classe [fwXLSPConn](src/fwXLSPConn.cls) é o ponto central quando a automação precisa executar SQL envolvendo Excel e SharePoint. Ela monta um catálogo das fontes disponíveis, entende quais tabelas estão no workbook, registra as listas SharePoint informadas pelo desenvolvedor e decide qual caminho usar para cada comando.

Se a consulta usa apenas uma tabela local do Excel, o framework pode enviar o trabalho para `fwXLConn`. Se usa apenas uma lista SharePoint, pode enviar para `fwSPConn`. Se mistura fontes, o framework pode criar uma área temporária de staging no Excel, trazer os dados necessários e executar o restante localmente.

Para quem está usando, a ideia é simples: você chama `RunQuery` com um SQL. Por trás, o framework faz a leitura do comando, identifica as fontes envolvidas, ajusta nomes físicos, escolhe o conector adequado e devolve um resultado padronizado.

```vb
Dim db As New fwXLSPConn
Dim result As Object

db.InitSP _
    "FuncionariosSP", "https://empresa.sharepoint.com/sites/rh", "00000000-0000-0000-0000-000000000000"

Set result = db.RunQuery( _
    "SELECT [Nome], [Cargo] FROM [FuncionariosSP] WHERE [Cargo] = 'Analista'")
```

O objetivo não é transformar VBA em um banco de dados completo. O objetivo é dar uma linguagem comum para consultas, cargas e sincronizações que aparecem o tempo todo em automações de escritório.

### Camada Excel

A classe [fwXLConn](src/fwXLConn.cls) é o motor de conexão com o Excel via ACE/OLEDB. Ela permite executar SQL contra o workbook, tratando detalhes como dialeto ACE, modo de leitura ou escrita, nomes físicos de tabelas e retorno de dados.

Já a classe [fwXLTable](src/fwXLTable.cls) trabalha em um nível mais próximo do usuário. Ela encapsula `ListObject`, que é a tabela estruturada do Excel. Em vez de ficar procurando coluna por índice, calculando linha atual, tentando respeitar filtros e repetindo blocos de código para carregar arrays, você usa uma API própria para navegar e alterar a tabela.

Com `fwXLTable`, uma tabela do Excel vira um objeto de processo. Você pode iniciar uma tabela pelo nome, obter a linha atual, ler e escrever valores por nome de coluna, contar linhas visíveis, adicionar colunas, carregar um resultado de consulta, limpar dados, aplicar filtros e ajustar formatos usados em integrações SAP.

```vb
Dim tb As New fwXLTable

tb.Init "TabelaPedidos", "ID"
tb.firstRow.Select

Do
    tb.SetCurrentRowValue "OBSERVACOES", "Linha em processamento"
Loop While tb.ExistsNextRow(True)
```

Esse estilo é mais legível para quem ainda está aprendendo VBA, porque o código passa a falar a língua do processo: tabela, linha atual, coluna, valor, filtro, resultado. Ao mesmo tempo, continua sendo útil para desenvolvedores experientes, porque reduz duplicação e concentra os detalhes de manipulação do Excel em um lugar só.

### Camada SharePoint

A classe [fwSPConn](src/fwSPConn.cls) cuida da conexão com listas SharePoint via ADODB/OLEDB. Ela conhece a URL do site, o identificador da lista, o campo chave usado nas atualizações e as rotas de leitura e escrita.

O mais importante aqui é que o desenvolvedor não precisa espalhar detalhes de conexão SharePoint pela macro de negócio. As listas são registradas de forma declarativa em `fwXLSPConn.InitSP`, sempre em trios: nome lógico, URL e GUID da lista.

```vb
Dim db As New fwXLSPConn

db.InitSP _
    "Solicitacoes", "https://empresa.sharepoint.com/sites/operacao", "11111111-1111-1111-1111-111111111111", _
    "Historico", "https://empresa.sharepoint.com/sites/operacao", "22222222-2222-2222-2222-222222222222"
```

Depois disso, o SQL passa a usar os nomes lógicos. Se a URL mudar, ou se a lista for substituída em ambiente de homologação, a regra do processo não precisa ser reescrita. Ajusta-se a configuração, não o fluxo inteiro.

### Camada SAP GUI

A classe [fwSAPConn](src/fwSAPConn.cls) abre ou reaproveita uma sessão SAP GUI. Ela concentra a lógica de conexão que normalmente ficaria repetida em várias macros: localizar o SAP Logon, escolher a conexão, pegar a sessão correta e devolver um objeto pronto para uso.

Depois da conexão, [fwGuiMainWindow](src/fwGuiMainWindow.cls) representa a janela principal do SAP. Ela oferece métodos para abrir transações, enviar comandos, buscar objetos, ler a barra de status e criar wrappers mais específicos para menus, árvores e tabelas.

```vb
Dim sap As New fwSAPConn
Dim w As New fwGuiMainWindow

If sap.Connect("PRD", 100, 0) Then
    w.Init sap.Session
    w.TCode "SE16N"
End If
```

SAP GUI Scripting costuma gerar códigos muito dependentes de `findById`. Isso funciona, mas fica frágil quando a tela muda, quando existe popup, quando o mesmo controle aparece em outro container ou quando o desenvolvedor precisa adaptar a macro para outra transação. Os wrappers `fwGui*` não eliminam a necessidade de conhecer a tela SAP, mas deixam a automação mais organizada.

### Wrappers de elementos SAP

O [fwGuiMenu](src/fwGuiMenu.cls) encapsula menus. Em vez de depender apenas de uma cadeia rígida de IDs, você pode trabalhar com objetos de menu e tratar o caso em que aquele item não está disponível no contexto atual.

O [fwGuiTree](src/fwGuiTree.cls) encapsula árvores SAP. Ele ajuda a localizar nós, expandir estruturas, selecionar itens, ler textos, pressionar botões de árvore e navegar por hierarquias.

O [fwGuiTableControl](src/fwGuiTableControl.cls) encapsula tabelas e grades do SAP. Ele ajuda a localizar colunas, ler células, verificar se uma célula existe, rolar tabela e trabalhar com controles que seriam bem trabalhosos se fossem tratados apenas por índice.

Esses wrappers foram pensados para um código mais defensivo. Em tela SAP, nem sempre o elemento está visível, habilitado ou presente. Por isso, sempre que a classe oferecer um método que retorna `True` ou `False`, vale preferir esse formato quando a ausência do elemento for um cenário esperado.

```vb
If Not w.GuiMenu("Sistema", "Status").SelectMenu() Then
    Application.StatusBar = "Menu não disponível nesta tela."
End If
```

## O catálogo de dados

Um dos conceitos mais importantes do SPXLSAP é o catálogo. O catálogo é a lista de fontes que o framework conhece naquele momento: tabelas do Excel e listas SharePoint registradas.

Quando você inicializa o orquestrador, o framework consegue descobrir os `ListObjects` do workbook ativo. Isso significa que uma tabela chamada `Pedidos` no Excel pode ser referenciada como `[Pedidos]` em um SQL. Para listas SharePoint, você registra manualmente os nomes lógicos com `InitSP`.

Esse desenho cria uma vantagem grande: o SQL do processo fica mais limpo. Ele fala de `[Pedidos]`, `[Solicitacoes]`, `[Historico]`, `[Usuarios]`, e não de detalhes físicos de conexão. O framework se encarrega de traduzir esses nomes para a origem correta.

Para evitar confusão, vale seguir uma regra simples: use nomes claros e únicos. Se existe uma tabela Excel chamada `Solicitacoes` e uma lista SharePoint também chamada `Solicitacoes`, o leitor humano vai se perder, mesmo que o código consiga resolver parte do problema. Um bom nome de fonte já é metade da documentação do processo.

## SQL como linguagem comum

O SPXLSAP usa SQL como linguagem comum para consultar, inserir, atualizar e excluir dados. Isso é poderoso porque muitos processos corporativos se resumem a selecionar registros, aplicar filtros, cruzar tabelas, atualizar status e gerar um resultado.

Os comandos suportados incluem:

- `SELECT`, para consultar dados;
- `INSERT`, para incluir registros;
- `UPDATE`, para alterar registros;
- `DELETE`, para excluir registros.

O SQL pode envolver tabelas Excel, listas SharePoint ou uma combinação das duas. Quando a combinação exige cuidado, o framework usa staging para trazer dados temporariamente ao Excel e processar a operação de forma mais controlada.

Algumas convenções ajudam a manter tudo previsível:

- use colchetes em nomes de tabelas e campos, como `[Nome da Coluna]`;
- evite caracteres especiais desnecessários nos nomes físicos das tabelas;
- mantenha nomes de colunas estáveis, principalmente quando eles são usados em SQL ou SAP;
- use `*` como curinga em `LIKE`, seguindo o comportamento esperado pelo ACE/OLEDB;
- salve o workbook antes de operações de escrita que dependam do driver ACE;
- prefira deixar regra de negócio complexa no VBA e usar SQL para seleção, junção, filtro e atualização de dados.

Um exemplo simples, usando apenas Excel:

```sql
SELECT [ID], [Nome], [Status]
FROM [Pedidos]
WHERE [Status] = 'Pendente'
ORDER BY [Nome]
```

Um exemplo misturando SharePoint e Excel:

```sql
SELECT s.[ID], s.[Solicitante], p.[Prioridade]
FROM [SolicitacoesSP] AS s
INNER JOIN [Parametros] AS p ON s.[Tipo] = p.[Tipo]
WHERE p.[Ativo] = True
```

Para quem tem pouca experiência, esse modelo ensina uma ideia essencial: antes de automatizar uma tela SAP, organize bem os dados que vão entrar nela. Muitas falhas de automação não nascem no clique errado, mas em dados mal preparados.

## O mecanismo de staging

Staging é uma palavra técnica para uma ideia simples: antes de executar uma operação mais complexa, o framework cria uma área temporária de trabalho.

Imagine que você quer atualizar uma lista SharePoint com base em uma tabela do Excel. O SharePoint está em uma origem, o Excel em outra, e os drivers nem sempre conseguem resolver todos os tipos de `JOIN` diretamente entre essas fontes. Em vez de forçar uma operação frágil, o SPXLSAP pode trazer os dados necessários para tabelas temporárias no Excel, executar a parte pesada localmente e depois sincronizar de volta apenas os registros afetados.

Esse fluxo aparece principalmente em cenários híbridos:

- consulta que cruza Excel com SharePoint;
- consulta que envolve mais de uma lista SharePoint;
- `UPDATE` em SharePoint baseado em filtros ou joins com outras fontes;
- `INSERT ... SELECT` copiando dados entre origens diferentes.

De forma resumida, o pipeline funciona assim:

1. o framework lê o SQL e identifica as fontes envolvidas;
2. as listas SharePoint necessárias são consultadas;
3. os dados são materializados em tabelas temporárias no Excel;
4. o comando é executado localmente quando isso for mais seguro ou mais compatível;
5. se houver alteração em SharePoint, os registros afetados são sincronizados de volta;
6. as tabelas temporárias são limpas ao final do processo.

Para o desenvolvedor, a vantagem é escrever um comando de alto nível. Para o processo, a vantagem é reduzir round-trips desnecessários, controlar melhor os registros afetados e contornar limitações comuns dos drivers.

## Como cada comando é tratado

O desenvolvedor não precisa decorar todas as rotas internas para usar o SPXLSAP, mas entender a lógica geral ajuda bastante na hora de diagnosticar uma automação. O framework tenta escolher o caminho mais simples que resolve o problema, e só usa uma estratégia mais elaborada quando a consulta mistura origens ou quando o driver tem alguma limitação prática.

### SELECT

Quando o `SELECT` usa apenas tabelas do Excel, o caminho natural é o **Direct-XL**. O SQL é ajustado para o dialeto ACE/OLEDB e executado diretamente contra o workbook. Essa rota é boa para relatórios locais, filtros, cruzamentos entre `ListObjects` e preparação de dados antes de uma etapa SAP.

Quando o `SELECT` usa apenas uma lista SharePoint, o caminho natural é o **Direct-SP**. Nesse caso, a consulta é enviada ao conector SharePoint e o resultado volta para o VBA como uma estrutura que pode ser carregada no Excel ou usada pela macro.

Quando o `SELECT` mistura Excel e SharePoint, ou envolve mais de uma lista SharePoint, entra a rota **Stage-First**. O framework traz os dados necessários para tabelas temporárias no Excel, executa o cruzamento localmente e devolve o resultado final. Para quem escreve a macro, continua parecendo um único SQL; para o framework, foi uma pequena operação federada.

### INSERT

O `INSERT` pode ser usado de duas formas. Com `VALUES`, ele representa uma inclusão direta: uma linha nova vai para uma tabela Excel ou para uma lista SharePoint. Com `INSERT ... SELECT`, ele vira uma carga baseada em consulta, o que permite copiar dados entre origens diferentes.

Esse segundo caso é especialmente útil em automações que preparam uma fila de trabalho. Por exemplo: consultar registros pendentes no SharePoint, cruzar com uma tabela de parâmetros no Excel e carregar o resultado em uma tabela local que será percorrida por uma rotina SAP.

### UPDATE

O `UPDATE` é tratado com mais cuidado porque altera dados existentes. Quando a atualização envolve apenas Excel, o framework tende a usar o caminho direto via ACE/OLEDB ou as rotas auxiliares de `fwXLTable`, conforme a necessidade do comando.

Quando o alvo é SharePoint, a prioridade é controle. O framework pode consultar primeiro os registros afetados, aplicar atualizações de forma iterativa ou usar o fluxo Stage-Update-Sync quando a atualização depende de dados combinados entre origens. Essa abordagem evita mandar uma alteração grande e cega para a lista, além de facilitar o retorno de quantidade de linhas afetadas e mensagens de acompanhamento.

### DELETE

O `DELETE` também é uma operação sensível. No Excel, a exclusão precisa respeitar filtros, linhas visíveis e limitações do driver. Por isso, o framework pode transformar a exclusão em uma etapa de consulta dos IDs e depois remover as linhas correspondentes com controle programático.

No SharePoint, a lógica é parecida: primeiro identifica-se o conjunto de registros que atende ao filtro; depois a exclusão é aplicada de forma controlada. Esse padrão é mais previsível para listas corporativas e combina melhor com processos que precisam registrar o que foi removido, ignorado ou bloqueado.

Essa visão por comando mostra uma característica importante do SPXLSAP: ele não é só um atalho para executar SQL. Ele é uma camada de decisão que tenta escolher uma rota compatível com a origem, o tipo de operação e o risco envolvido.

## Padrões de desenvolvimento

O SPXLSAP fica mais fácil de usar quando algumas convenções são respeitadas desde o início do projeto. Elas não existem para engessar o desenvolvedor. Elas existem para que outro colega, ou você mesmo daqui a alguns meses, consiga entender o fluxo sem precisar desmontar a solução inteira.

### Use tabelas estruturadas como contrato

Sempre que possível, use `ListObject` no Excel em vez de intervalos soltos. Uma tabela estruturada tem nome, colunas, filtros e um corpo de dados reconhecível pelo framework. Isso torna o código mais resistente a inserção de linhas, mudança de tamanho da massa de dados e ajustes visuais na planilha.

Uma tabela chamada `tblOrdens` com colunas `ID`, `ORDEM`, `STATUS` e `OBSERVACOES` é muito mais clara do que um intervalo `A1:K5000` manipulado por coordenadas.

### Centralize parâmetros

URLs de SharePoint, GUIDs de listas, nome de conexão SAP, mandante, número de sessão, variantes, flags de execução e caminhos de arquivo devem ficar centralizados. Em muitas soluções isso será uma aba de parâmetros, uma tabela `Variaveis` ou uma estrutura equivalente.

O importante é evitar valores fixos escondidos no meio da macro. Quando o ambiente muda, a configuração muda junto. A regra de negócio deve permanecer o mais estável possível.

### Escreva rotinas pequenas de processo

Mesmo usando framework, uma automação pode ficar confusa se uma única macro fizer tudo. Prefira dividir o fluxo em etapas com nomes claros: carregar dados, validar linhas, conectar SAP, processar item, registrar retorno, sincronizar status.

Essa divisão ajuda iniciantes a acompanhar o raciocínio e ajuda desenvolvedores experientes a testar e substituir partes do processo sem mexer em tudo.

### Registre feedback para o usuário

Automação corporativa geralmente roda sobre dados reais. O usuário precisa saber o que aconteceu. Uma coluna como `OBSERVACOES`, `STATUS_PROCESSAMENTO` ou `MENSAGEM` ajuda a registrar linha a linha se o item foi processado, ignorado, bloqueado, atualizado ou se falhou.

### Prefira wrappers a chamadas soltas de SAP GUI

`session.findById("wnd[0]/usr/...")` é inevitável em alguns momentos, mas não precisa dominar todo o código. Quando uma ação se repete, coloque-a atrás de um wrapper ou método mais claro. Isso reduz duplicação, facilita fallback e deixa a rotina mais legível.

Em vez de uma sequência enorme de `findById`, procure expressar a intenção:

```vb
w.TCode "IW33"
w.SetIfChangeable "ctxtCAUFVD-AUFNR", ordem
w.Enter
```

Mesmo que por baixo ainda exista SAP GUI Scripting, o código de negócio fica mais próximo daquilo que o processo realmente faz.

### Trate filtros e linhas visíveis com cuidado

Em Excel, usuário filtra tabela o tempo todo. Uma rotina que processa todas as linhas quando o usuário esperava processar apenas as linhas visíveis pode causar erro operacional sério. Por isso, `fwXLTable` oferece recursos para contar e navegar por linhas visíveis.

Antes de criar uma rotina de lote, decida explicitamente: ela processa tudo ou apenas o que está visível? Depois deixe isso claro no nome, na mensagem e no comportamento.

## Fluxo recomendado para uma nova solução

Um caminho seguro para criar uma automação com SPXLSAP é começar pequeno e ir adicionando camadas.

Primeiro, defina a tabela principal do processo no Excel. Dê um nome claro ao `ListObject`, organize as colunas e inclua uma coluna de retorno, como `OBSERVACOES`. Essa tabela será o ponto de encontro entre o usuário, os dados e a macro.

Depois, crie uma rotina simples usando `fwXLTable` para percorrer a tabela e escrever uma mensagem em cada linha. Isso valida se a base da automação está entendendo filtros, linhas e colunas.

Em seguida, adicione a conexão SAP com `fwSAPConn` e teste apenas a abertura de uma transação. Não tente automatizar tudo de uma vez. Primeiro garanta que a sessão correta está sendo usada.

Depois, encapsule os passos SAP principais em uma rotina pequena. Abra a transação, preencha um campo, pressione Enter, leia a barra de status. Quando isso estiver estável, inclua a navegação por tabela, menu ou árvore, usando os wrappers `fwGui*` quando fizer sentido.

Se a solução também precisar de SharePoint, registre as listas com `InitSP` e comece por um `SELECT`. Depois teste `INSERT`, `UPDATE` ou `DELETE` em ambiente controlado. Operações de escrita devem sempre ser validadas com poucos registros antes de rodar em massa.

Por fim, conecte as partes: os dados entram pelo Excel ou SharePoint, o SAP executa o trabalho, o retorno volta para a tabela e a sincronização atualiza a fonte corporativa.

## Exemplos de uso

### Consultar uma tabela Excel

```vb
Dim db As New fwXLSPConn
Dim result As Object

db.InitLocal

Set result = db.RunQuery( _
    "SELECT [ID], [Nome], [Status] FROM [Pedidos] WHERE [Status] = 'Pendente'")
```

Esse exemplo usa apenas o workbook. É uma boa porta de entrada para aprender o orquestrador sem envolver SharePoint ou SAP.

### Carregar resultado em uma tabela

```vb
Dim destino As New fwXLTable

destino.Init "Resultado"
destino.LoadResult result
```

Aqui, a tabela `Resultado` recebe os dados retornados por uma consulta. Isso evita código manual para redimensionar range, escrever cabeçalho e preencher linhas.

### Atualizar SharePoint a partir de SQL

```vb
Dim db As New fwXLSPConn
Dim report As Variant

db.InitSP _
    "SolicitacoesSP", "https://empresa.sharepoint.com/sites/operacao", "11111111-1111-1111-1111-111111111111"

Call db.RunQuery( _
    "UPDATE [SolicitacoesSP] SET [Status] = 'Tratado' WHERE [ID] = 123", _
    report)
```

O `execReport` pode ser usado para inspecionar informações de execução, como SQL final, tempo e quantidade de linhas afetadas, conforme a rota usada pelo framework.

### Abrir uma transação SAP

```vb
Dim sap As New fwSAPConn
Dim w As New fwGuiMainWindow

If sap.Connect("PRD", 100, 0) Then
    w.Init sap.Session
    w.TCode "IW33"
End If
```

Esse é o primeiro teste recomendado antes de automatizar qualquer tela. Se a conexão e a transação estiverem corretas, o restante do fluxo fica mais fácil de isolar.

### Ler informação de uma tabela SAP

```vb
Dim tbl As fwGuiTableControl

Set tbl = w.GuiTable("tblSAPLCOVGTC_CONTROL")

If tbl.CellExists(0, "AUFNR") Then
    Debug.Print tbl.CellText(0, "AUFNR")
Else
    Application.StatusBar = "Célula não encontrada na tabela SAP."
End If
```

O exemplo é propositalmente simples. Em telas reais, muitas vezes será melhor localizar colunas por nome técnico ou por texto de cabeçalho, dependendo do controle SAP disponível.

## Componentes do framework

| Arquivo | Papel no framework |
|---|---|
| [src/fwXLSPConn.cls](src/fwXLSPConn.cls) | Orquestrador SQL multiorigem. Decide se uma operação vai para Excel, SharePoint ou staging. |
| [src/fwXLConn.cls](src/fwXLConn.cls) | Conector Excel via ACE/OLEDB. Executa SQL contra tabelas do workbook. |
| [src/fwSPConn.cls](src/fwSPConn.cls) | Conector SharePoint via ADODB/OLEDB. Executa leituras e escritas em listas. |
| [src/fwXLTable.cls](src/fwXLTable.cls) | Wrapper de `ListObject`. Navega, carrega, filtra, atualiza e manipula tabelas estruturadas do Excel. |
| [src/fwHelpers.cls](src/fwHelpers.cls) | Biblioteca de apoio. Faz parsing de SQL, normalização de nomes, relatórios e funções compartilhadas. |
| [src/fwSAPConn.cls](src/fwSAPConn.cls) | Gerencia conexão com SAP GUI e devolve uma sessão pronta para automação. |
| [src/fwGuiMainWindow.cls](src/fwGuiMainWindow.cls) | Representa a janela principal SAP. Abre transações, localiza controles e cria wrappers específicos. |
| [src/fwGuiMenu.cls](src/fwGuiMenu.cls) | Encapsula menus SAP e facilita seleção defensiva de itens. |
| [src/fwGuiTree.cls](src/fwGuiTree.cls) | Encapsula árvores SAP, nós, caminhos, seleções e ações contextuais. |
| [src/fwGuiTableControl.cls](src/fwGuiTableControl.cls) | Encapsula tabelas e grades SAP, com apoio a leitura, seleção, colunas e rolagem. |
| [src/fwSapGuiApi.bas](src/fwSapGuiApi.bas) | Módulo com apoio para SAP GUI Scripting. |
| [VERSION](VERSION) | Versão atual publicada do framework. |
| [CHANGELOG.md](CHANGELOG.md) | Histórico de mudanças relevantes. |

## Requisitos de ambiente

Para usar o SPXLSAP em uma solução real, normalmente você precisará de:

- Microsoft Excel com suporte a VBA e macros habilitadas;
- permissão para executar macros no arquivo da solução;
- SAP GUI instalado, quando houver automação SAP;
- SAP GUI Scripting habilitado no ambiente e permitido ao usuário;
- acesso às conexões SAP usadas pela automação;
- acesso às listas SharePoint usadas pelo processo;
- permissão de leitura ou escrita nas fontes corporativas envolvidas;
- referências ou late binding compatíveis com os objetos usados no projeto.

O framework ajuda a organizar a automação, mas não substitui permissão de sistema. Se o usuário não pode alterar uma lista SharePoint manualmente, a macro também não deve conseguir. Se o ambiente bloqueia SAP GUI Scripting, a automação SAP precisa ser liberada antes.

## Cuidados importantes

Operações de escrita merecem respeito. `UPDATE`, `INSERT` e `DELETE` podem alterar dados reais em Excel ou SharePoint. Antes de rodar uma rotina em produção, teste com poucos registros, em ambiente de homologação ou com uma lista de teste. Use filtros de tabela para reduzir o lote inicial e acompanhe as mensagens linha a linha.

Automação SAP também exige cuidado. Uma tela errada, uma variante incorreta ou uma sessão diferente da esperada pode levar o processo para outro caminho. Sempre valide conexão, mandante, sessão e transação antes de executar uma rotina em lote.

Outro ponto importante é a manutenção dos IDs de tela. SAP GUI Scripting depende da estrutura da interface. Se a transação mudar, se o usuário estiver com layout diferente ou se surgir um popup inesperado, o código precisa tratar esse contexto. Os wrappers ajudam, mas não fazem milagre: uma boa automação continua precisando prever os desvios mais comuns.

## Adaptabilidade

O SPXLSAP não foi feito para um único processo. Ele foi feito para ser uma base de construção. O framework pode sustentar soluções, como:

- carga e validação de dados antes de execução no SAP;
- consulta de informações SAP e montagem de relatórios em Excel;
- atualização controlada de listas SharePoint por usuários de negócio;
- rotinas de saneamento de dados antes de upload em sistemas corporativos;
- conciliações entre Excel, SharePoint e resultados extraídos do SAP;
- assistentes de processamento com status por linha;
- automações que reaproveitam uma mesma sessão SAP para vários itens de uma fila.

A adaptabilidade vem do desenho em camadas. Se uma solução não usa SharePoint, ela pode usar apenas Excel e SAP. Se não usa SAP, pode usar apenas SQL entre Excel e SharePoint. Se precisa apenas de uma tabela Excel mais organizada, `fwXLTable` já resolve uma parte do problema. O framework não obriga todo projeto a usar tudo.

## Para quem está aprendendo

Se você é iniciante em VBA, não tente decorar todas as classes de uma vez. Comece pelo problema mais concreto.

Se sua dificuldade é manipular tabela do Excel, abra [fwXLTable.cls](src/fwXLTable.cls) e procure exemplos de `Init`, `RangeRow`, `GetCurrentRowValue`, `SetCurrentRowValue`, `XLLoadArray` e filtros.

Se sua dificuldade é conexão SAP, olhe [fwSAPConn.cls](src/fwSAPConn.cls) e [fwGuiMainWindow.cls](src/fwGuiMainWindow.cls). Teste primeiro abrir uma transação. Depois leia a barra de status. Depois preencha um campo.

Se sua dificuldade é SharePoint, comece com `InitSP` e um `SELECT` simples. Só depois avance para escrita.

## Para quem já desenvolve em VBA

Se você já tem experiência, o ganho principal está em padronização e manutenção. O SPXLSAP reduz código repetido de conexão, navegação, carga de tabela e sincronização. Também cria uma linguagem comum entre projetos: tabelas estruturadas, catálogo, `RunQuery`, wrappers SAP, mensagens por linha e parâmetros centralizados.

Isso facilita revisão de código, passagem de conhecimento e evolução incremental. Quando uma melhoria é feita em um conector ou wrapper do framework, outras soluções podem se beneficiar sem que cada automação precise ser reescrita do zero.

## Como contribuir ou evoluir

Ao evoluir o framework, mantenha a mesma preocupação que motivou sua criação: facilitar a vida de quem vai construir em cima dele. Uma nova função deve ter nome claro, responsabilidade bem definida e comportamento previsível. Se a função resolve um detalhe técnico difícil, tente esconder a complexidade atrás de uma API simples.

Também vale preservar a separação entre framework e solução. Código genérico, reaproveitável e independente de processo pertence ao SPXLSAP. Código que conhece uma transação específica, uma lista específica, uma aba de negócio ou uma regra operacional deve ficar no projeto consumidor.

Antes de publicar mudanças, confira:

- se os nomes continuam consistentes com o restante do framework;
- se a mudança não quebra chamadas existentes sem necessidade;
- se a documentação ou exemplos precisam ser atualizados;
- se os arquivos exportados de VBA continuam em encoding compatível com o ambiente;
- se o [CHANGELOG.md](CHANGELOG.md) deve registrar a alteração.

## Fechamento

O SPXLSAP é uma base para criar automações mais organizadas em um ambiente onde Excel, SharePoint e SAP precisam trabalhar juntos. Ele não elimina a necessidade de conhecer o processo, testar com cuidado e entender os sistemas envolvidos. Mas reduz bastante o código repetitivo, cria um caminho mais claro para quem está aprendendo e oferece pontos de extensão para quem precisa construir soluções mais robustas.

Use o framework como uma caixa de ferramentas: comece com a peça que resolve o problema de hoje, entenda como ela se encaixa nas demais e evolua sua solução por etapas. Esse é o caminho mais leve para sair de uma macro isolada e chegar a uma automação corporativa mais confiável.
