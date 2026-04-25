VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fwXLSPConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' Nome do elemento: fwXLSPConn (Classe)
' Resumo:
'   Implementa o orquestrador de consultas SQL/SharePoint, cuidando do staging
'   em Excel, sincronização com listas SP e validação de esquema.
'******************************************************************************
Option Explicit

Private Const SYNC_ID_CHUNK_SIZE As Long = 400
Private Const DELTASYNCSTATESHEETNAME As String = "_fwDeltaSyncState"

Private pCatalog As Object
Private pHelpers As fwHelpers
Private pTraceOn As Boolean
Private pTempTables As Object
Private pSchemaCache As Object
Private pDeltaSyncConfig As Object
Private pEnableCellChangeHighlight As Boolean
Private pCellChangeHighlightColor As Long
Private pClearPreviousHighlightOnUpdate As Boolean

'******************************************************************************
' Nome do elemento: Class_Initialize (Sub)
' Resumo: Executa a rotina Class_Initialize.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub Class_Initialize()
    Set pHelpers = New fwHelpers
    Set pTempTables = CreateObject("Scripting.Dictionary")
    pTempTables.CompareMode = vbTextCompare
    Set pSchemaCache = CreateObject("Scripting.Dictionary")
    pSchemaCache.CompareMode = vbTextCompare
    Set pDeltaSyncConfig = CreateObject("Scripting.Dictionary")
    pDeltaSyncConfig.CompareMode = vbTextCompare
    pEnableCellChangeHighlight = False
    pCellChangeHighlightColor = RGB(255, 242, 204)
    pClearPreviousHighlightOnUpdate = False
End Sub

'******************************************************************************
' Nome do elemento: Class_Terminate (Sub)
' Resumo: Executa a rotina Class_Terminate.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub Class_Terminate()
    On Error Resume Next
    Call CleanupTempTables
    Set pSchemaCache = Nothing
    Set pDeltaSyncConfig = Nothing
    Set pCatalog = Nothing
    Set pHelpers = Nothing
    Set pTempTables = Nothing
End Sub

'******************************************************************************
' Nome do elemento: Init (Sub)
' Resumo: Executa a rotina Init.
' @param manualCatalog - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub Init(ByVal manualCatalog As Object)
    On Error GoTo ErrHandler

    ' Etapa 1: Gerar dinamicamente o catálogo de todas as tabela XL
    Dim xlCatalog As Object
    Set xlCatalog = pHelpers.CreateDynamifwXLTableCatalog(ActiveWorkbook)

    ' Etapa 2: Fundir o catálogo manual (agora vindo da InitSP)
    Dim key As Variant
    If Not pHelpers.IsVarEmpty(manualCatalog) Then
        For Each key In manualCatalog.keys
            If xlCatalog.Exists(key) Then
                Trace "Aviso: A fonte de dados '" & key & "' do SharePoint está sobrescrevendo uma tabela Excel com o mesmo nome no catálogo."
            End If
            Set xlCatalog(key) = manualCatalog(key)
        Next key
    End If

    ' Etapa 3: Define o catálogo final e completo
    Set pCatalog = xlCatalog
    If pSchemaCache Is Nothing Then
        Set pSchemaCache = CreateObject("Scripting.Dictionary")
        pSchemaCache.CompareMode = vbTextCompare
    End If
    pSchemaCache.RemoveAll
    Call LoadDeltaSyncConfigFromCatalog

    Trace "Init: Orquestrador inicializado com " & pCatalog.Count & " fonte(s) de dados."
    Exit Sub

ErrHandler:
    MsgBox "Erro em fwXLSPConn.Init: " & Err.Description
End Sub

'******************************************************************************
' Nome do elemento: InitSP (Sub)
' Resumo: Executa a rotina InitSP.
' @param spLists - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub InitSP(ParamArray spLists() As Variant)
    ' ParamArray captura todos os argumentos em um array

    ' Usa o novo helper para criar um catálogo SP a partir dos parâmetros
    Dim spCatalog As Object
    Set spCatalog = pHelpers.ParseSPInitParams(spLists)

    ' Chama a rotina de inicialização principal com o catálogo gerado
    Call Init(spCatalog)
End Sub

'******************************************************************************
' Nome do elemento: ConfigureDeltaSync (Sub)
' Resumo:
'   Habilita ou desabilita sincronização incremental para uma lista SP.
' @param logicalName - Nome lógico da lista no catálogo.
' @param modifiedField - Campo de data/hora usado como watermark.
' @param stateKey - Chave opcional para persistência do checkpoint.
' @param enable - Liga/desliga o Delta Sync para a lista.
'******************************************************************************
Public Sub ConfigureDeltaSync( _
    ByVal logicalName As String, _
    Optional ByVal modifiedField As String = "Modified", _
    Optional ByVal stateKey As String = "", _
    Optional ByVal enable As Boolean = True _
)
    On Error GoTo SafeExit

    If pDeltaSyncConfig Is Nothing Then
        Set pDeltaSyncConfig = CreateObject("Scripting.Dictionary")
        pDeltaSyncConfig.CompareMode = vbTextCompare
    End If

    If Len(Trim$(logicalName)) = 0 Then Exit Sub
    If Len(Trim$(modifiedField)) = 0 Then modifiedField = "Modified"
    If Len(Trim$(stateKey)) = 0 Then stateKey = logicalName

    Dim cfg As Object
    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.Add "Enabled", CBool(enable)
    cfg.Add "ModifiedField", Trim$(modifiedField)
    cfg.Add "StateKey", Trim$(stateKey)

    If pDeltaSyncConfig.Exists(logicalName) Then
        Set pDeltaSyncConfig(logicalName) = cfg
    Else
        pDeltaSyncConfig.Add logicalName, cfg
    End If

    Trace "ConfigureDeltaSync: list='" & logicalName & "', enabled=" & CStr(enable) & ", field='" & modifiedField & "', key='" & stateKey & "'."
    Exit Sub

SafeExit:
    Err.Clear
End Sub

'******************************************************************************
' Nome do elemento: InitLocal (Sub)
' Resumo: Executa a rotina InitLocal.
' @remarks Inicializa o orquestrador apenas com tabelas locais do workbook.
'******************************************************************************
Public Sub InitLocal()
    Call Init(Nothing)
End Sub

'******************************************************************************
' Nome do elemento: TraceOn (Property Let)
' Resumo: Define o valor da propriedade TraceOn.
' @param v - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Let TraceOn(ByVal v As Boolean)
    pTraceOn = CBool(v)
    If Not pHelpers Is Nothing Then pHelpers.TraceOn = v
End Property

'******************************************************************************
' Nome do elemento: TraceOn (Property Get)
' Resumo: Obtem o valor da propriedade TraceOn.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get TraceOn() As Boolean
    TraceOn = pTraceOn
End Property

'******************************************************************************
' Nome do elemento: EnableCellChangeHighlight (Property Let)
' Resumo: Define se o framework deve marcar celulas alteradas apos UPDATE JOIN.
' @param v - Parâmetro da rotina.
'******************************************************************************
Public Property Let EnableCellChangeHighlight(ByVal v As Boolean)
    pEnableCellChangeHighlight = CBool(v)
End Property

'******************************************************************************
' Nome do elemento: EnableCellChangeHighlight (Property Get)
' Resumo: Indica se a marcação visual de alteracoes esta habilitada.
' @return Boolean - Valor retornado pelo elemento.
'******************************************************************************
Public Property Get EnableCellChangeHighlight() As Boolean
    EnableCellChangeHighlight = pEnableCellChangeHighlight
End Property

'******************************************************************************
' Nome do elemento: CellChangeHighlightColor (Property Let)
' Resumo: Define a cor utilizada para marcar celulas alteradas.
' @param v - Parâmetro da rotina.
'******************************************************************************
Public Property Let CellChangeHighlightColor(ByVal v As Long)
    pCellChangeHighlightColor = CLng(v)
End Property

'******************************************************************************
' Nome do elemento: CellChangeHighlightColor (Property Get)
' Resumo: Retorna a cor configurada para marcação visual.
' @return Long - Valor retornado pelo elemento.
'******************************************************************************
Public Property Get CellChangeHighlightColor() As Long
    CellChangeHighlightColor = pCellChangeHighlightColor
End Property

'******************************************************************************
' Nome do elemento: ClearPreviousHighlightOnUpdate (Property Let)
' Resumo:
'   Define se a marcação anterior (mesma cor) deve ser removida antes de
'   aplicar novos destaques no UPDATE ... JOIN.
' @param v - Parâmetro da rotina.
'******************************************************************************
Public Property Let ClearPreviousHighlightOnUpdate(ByVal v As Boolean)
    pClearPreviousHighlightOnUpdate = CBool(v)
End Property

'******************************************************************************
' Nome do elemento: ClearPreviousHighlightOnUpdate (Property Get)
' Resumo:
'   Indica se o framework remove marcações anteriores antes de destacar
'   alterações de uma nova execução.
' @return Boolean - Valor retornado pelo elemento.
'******************************************************************************
Public Property Get ClearPreviousHighlightOnUpdate() As Boolean
    ClearPreviousHighlightOnUpdate = pClearPreviousHighlightOnUpdate
End Property

'******************************************************************************
' Nome do elemento: Trace (Sub)
' Resumo: Executa a rotina Trace.
' @param msg - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub Trace(ByVal msg As String)
    On Error Resume Next
    If Not pTraceOn Then Exit Sub
    Debug.Print Format$(now, "yyyy-mm-dd hh:nn:ss") & " [" & TypeName(Me) & "] " & msg
End Sub

'******************************************************************************
' Nome do elemento: ValidateRunQueryPreflight (Sub)
' Resumo:
'   Executa validacoes obrigatorias antes da execucao de RunQuery.
' @param sql - Comando SQL a ser executado.
'******************************************************************************
Private Sub ValidateRunQueryPreflight(ByVal sql As String)
    If pHelpers Is Nothing Then Set pHelpers = New fwHelpers

    Dim normalizedSql As String
    normalizedSql = Trim$(sql)

    If Len(normalizedSql) = 0 Then
        Call pHelpers.AppendPersistentLog("ERROR", "fwXLSPConn", "ValidateRunQueryPreflight", _
                                          "SQL vazio ou em branco.", sql)
        Err.Raise vbObjectError + 809, "fwXLSPConn.RunQuery", "Comando SQL obrigatorio."
    End If

    If pHelpers.IsVarEmpty(pCatalog) Then
        Call pHelpers.AppendPersistentLog("ERROR", "fwXLSPConn", "ValidateRunQueryPreflight", _
                                          "Orquestrador nao inicializado.", normalizedSql)
        Err.Raise vbObjectError + 810, "fwXLSPConn.RunQuery", "Orquestrador nao inicializado. Chame InitSP primeiro."
    End If

    If Not (pHelpers.IsSelectSql(normalizedSql) Or pHelpers.IsInsertSql(normalizedSql) Or _
            pHelpers.IsUpdateSql(normalizedSql) Or pHelpers.IsDeleteSql(normalizedSql)) Then
        Call pHelpers.AppendPersistentLog("ERROR", "fwXLSPConn", "ValidateRunQueryPreflight", _
                                          "Comando SQL nao suportado.", normalizedSql)
        Err.Raise vbObjectError + 811, "fwXLSPConn.RunQuery", "Comando SQL nao suportado."
    End If

    Dim logicalName As Variant
    For Each logicalName In pCatalog.keys
        Dim catEntry As Object
        Set catEntry = pCatalog(logicalName)
        If Not catEntry Is Nothing Then
            If catEntry.Exists("Type") Then
                If StrComp(CStr(catEntry("Type")), "SP", vbTextCompare) = 0 Then
                    If (Not catEntry.Exists("Connection")) Or (Not IsObject(catEntry("Connection"))) Then
                        Call pHelpers.AppendPersistentLog("ERROR", "fwXLSPConn", "ValidateRunQueryPreflight", _
                                                          "Catalogo SP invalido para '" & CStr(logicalName) & "': Connection ausente.", normalizedSql)
                        Err.Raise vbObjectError + 816, "fwXLSPConn.RunQuery", _
                                  "Catalogo SP invalido para '" & CStr(logicalName) & "': Connection ausente."
                    End If

                    Dim conn As Object
                    Set conn = catEntry("Connection")
                    If (Not conn.Exists("SiteURL")) Or Len(Trim$(CStr(conn("SiteURL")))) = 0 Then
                        Call pHelpers.AppendPersistentLog("ERROR", "fwXLSPConn", "ValidateRunQueryPreflight", _
                                                          "Catalogo SP invalido para '" & CStr(logicalName) & "': SiteURL vazio.", normalizedSql)
                        Err.Raise vbObjectError + 817, "fwXLSPConn.RunQuery", _
                                  "Catalogo SP invalido para '" & CStr(logicalName) & "': SiteURL vazio."
                    End If

                    If (Not conn.Exists("ListID")) Or Len(Trim$(CStr(conn("ListID")))) = 0 Then
                        Call pHelpers.AppendPersistentLog("ERROR", "fwXLSPConn", "ValidateRunQueryPreflight", _
                                                          "Catalogo SP invalido para '" & CStr(logicalName) & "': ListID vazio.", normalizedSql)
                        Err.Raise vbObjectError + 818, "fwXLSPConn.RunQuery", _
                                  "Catalogo SP invalido para '" & CStr(logicalName) & "': ListID vazio."
                    End If
                End If
            End If
        End If
    Next logicalName
End Sub

'******************************************************************************
' Nome do elemento: RunQuery (Function)
' Resumo: Retorna o resultado calculado por RunQuery.
' @param sql - Parâmetro da rotina.
' @param execReport - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function RunQuery(ByVal sql As String, Optional ByRef execReport As Variant) As Object
    Dim result As Object
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevStatusBar As Variant
    Dim prevStatesCaptured As Boolean
    On Error GoTo ErrHandler

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevStatusBar = Application.StatusBar
    prevStatesCaptured = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    DoEvents
    
    If pHelpers Is Nothing Then Set pHelpers = New fwHelpers
    pHelpers.TraceOn = Me.TraceOn
    
    Call pHelpers.UpdateStatus("Iniciando consulta...")
    Call pHelpers.AppendPersistentLog("INFO", "fwXLSPConn", "RunQuery", "Iniciando execução SQL.", sql)
    
    Trace "RunQuery ENTER; SQLLen=" & Len(sql)
    Call ValidateRunQueryPreflight(sql)

    Call pHelpers.ValidateSqlTokenSpacing(sql)

    ' --- PARSING INICIAL E OBTENÇÃO DAS TABELA ---
    Dim parsedSql As Object
    Dim tablesInQuery As Collection
    Dim tblDict As Object
    
    If pHelpers.IsSelectSql(sql) Then
        Set parsedSql = pHelpers.ParseSqlQueryDetails(sql)
    ElseIf pHelpers.IsUpdateSql(sql) Then
        Set parsedSql = pHelpers.ParseUpdateJoinSqlDetails(sql)
    ElseIf pHelpers.IsDeleteSql(sql) Then
        Set parsedSql = pHelpers.ParseDeleteSqlDetails(sql)
    ElseIf pHelpers.IsInsertSql(sql) Then
         Set parsedSql = pHelpers.ParseInsertSqlDetails(sql)
         If parsedSql Is Nothing Then
              Err.Raise vbObjectError + 812, "RunQuery", "Falha ao analisar o comando INSERT."
         End If
         
         ' --- INÍCIO DA CORREÇÃO LÓGICA (INSERT...SELECT) ---
         Trace "[DEBUG-RunQuery] INSERT detectado. Construindo lista de tabelas para validação..."
         
         ' Usamos um Dicionário para garantir que aliases únicos sejam mapeados para nomes de tabela
         Dim tablesMap As Object: Set tablesMap = CreateObject("Scripting.Dictionary")
         tablesMap.CompareMode = vbTextCompare ' Ignora caso do alias
         
         ' 1. Adiciona a Tabela Alvo (TargetTable)
         ' Ela pode não ter um alias explícito na parte INSERT, mas o parser da
         ' subconsulta (SELECT) pode encontrá-la com um alias (ex: 'a' no seu caso)
         Dim targetTableName As String: targetTableName = parsedSql("TargetTable")
         Trace "[DEBUG-RunQuery]   Tabela Alvo (Destino): '" & targetTableName & "'"
         ' Adicionamos temporariamente com seu próprio nome como alias.
         ' Se a subconsulta usar um alias diferente para ela, ele será sobrescrito/adicionado.
         If Not tablesMap.Exists(targetTableName) Then tablesMap.Add targetTableName, targetTableName
         
         ' 2. Analisa a subconsulta SELECT (se houver)
         If parsedSql("InsertType") = "SELECT" Then
             Dim subParsedSql As Object
             Dim subTables As Collection ' Coleção de Dictionaries {"Name": "...", "Alias": "..."}
             
             Set subParsedSql = pHelpers.ParseSqlQueryDetails(parsedSql("SelectSubQuery"))
             
             If Not subParsedSql Is Nothing And subParsedSql.Exists("Tables") Then
                 Set subTables = subParsedSql("Tables")
                 Trace "[DEBUG-RunQuery]   Analisando " & subTables.Count & " tabela(s) da subconsulta SELECT..."
                 Dim subTbl As Object ' Dictionary {"Name": "...", "Alias": "..."}
                 For Each subTbl In subTables
                     Dim subTblName As String: subTblName = subTbl("Name")
                     Dim subTblAlias As String: subTblAlias = subTbl("Alias")
                     Trace "[DEBUG-RunQuery]     Encontrada Tabela: '" & subTblName & "', Alias: '" & subTblAlias & "'"
                     
                     ' Adiciona ou sobrescreve o alias no mapa
                     ' Isso garante que "a" seja mapeado para "GRTS"
                     If tablesMap.Exists(subTblAlias) Then
                         Trace "[DEBUG-RunQuery]     Aviso: Alias '" & subTblAlias & "' já existe. Sobrescrevendo mapeamento."
                         tablesMap(subTblAlias) = subTblName ' Sobrescreve
                     Else
                         tablesMap.Add subTblAlias, subTblName ' Adiciona novo alias
                     End If
                 Next subTbl
             End If
         End If
         
         ' 3. Constrói a coleção final 'tablesInQuery' a partir do Mapa
         Set tablesInQuery = New Collection
         Dim aliasKey As Variant
         For Each aliasKey In tablesMap.keys
             Set tblDict = CreateObject("Scripting.Dictionary")
             tblDict.Add "Name", tablesMap(aliasKey) ' Nome Real da Tabela
             tblDict.Add "Alias", aliasKey ' Alias usado no SQL
             tablesInQuery.Add tblDict
             Trace "[DEBUG-RunQuery]   >> Adicionado à validação: Alias='" & aliasKey & "' -> Nome='" & tablesMap(aliasKey) & "'"
         Next
         
         ' --- FIM DA CORREÇÃO LÓGICA ---

         GoTo SkipStandardTableExtraction ' Pula a extração padrão
    Else
         Set parsedSql = Nothing
    End If
    
    ' Extração padrão para SELECT, UPDATE, DELETE
    If Not parsedSql Is Nothing Then
        If parsedSql.Exists("Tables") Then
            Set tablesInQuery = parsedSql("Tables")
        ElseIf parsedSql.Exists("TargetTable") Then
            Set tablesInQuery = New Collection
            Set tblDict = CreateObject("Scripting.Dictionary")
            tblDict.Add "Name", parsedSql("TargetTable")
            tblDict.Add "Alias", parsedSql("TargetTable")
            tablesInQuery.Add tblDict
        End If
    ElseIf Not pHelpers.IsInsertSql(sql) Then
         Err.Raise vbObjectError + 811, "RunQuery", "Comando SQL não suportado ou falha no parsing inicial."
    End If

SkipStandardTableExtraction:
    ' --- FIM DO PARSING ---

    ' --- NOVA VALIDAÇÃO DE CATÁLOGO ---
    Trace "[DEBUG-RunQuery] Iniciando verificação de existência de tabelas no catálogo..."
    If Not tablesInQuery Is Nothing Then
        If tablesInQuery.Count > 0 Then
            Dim tbl As Object
            Dim tblNameToCheck As String
            For Each tbl In tablesInQuery
                tblNameToCheck = tbl("Name")
                Trace "[DEBUG-RunQuery] Verificando tabela: '" & tblNameToCheck & "' (Alias: '" & tbl("Alias") & "')"
                If Not pCatalog.Exists(tblNameToCheck) Then
                    Dim errMsg As String
                    errMsg = "Erro de Catálogo: A tabela/lista '" & tblNameToCheck & "' (alias: '" & tbl("Alias") & "') referenciada no SQL não foi encontrada no catálogo. Verifique a chamada InitSP ou o nome da tabela no SQL."
                    Trace errMsg
                    Err.Raise vbObjectError + 815, "fwXLSPConn.RunQuery", errMsg
                Else
                     Trace "[DEBUG-RunQuery] Tabela '" & tblNameToCheck & "' encontrada no catálogo."
                End If
            Next tbl
             Trace "[DEBUG-RunQuery] Verificação de existência de tabelas no catálogo concluída com SUCESSO."
        Else
             Trace "[DEBUG-RunQuery] Coleção 'tablesInQuery' está vazia. Pulando verificação de catálogo."
        End If
    Else
         Trace "[DEBUG-RunQuery] Coleção 'tablesInQuery' é Nothing. Pulando verificação de catálogo."
    End If
    ' --- FIM DA NOVA VALIDAÇÃO ---

    ' --- VALIDAÇÃO DE ESQUEMA UNIVERSAL ---
    If Not tablesInQuery Is Nothing Then
        If tablesInQuery.Count > 0 Then
            Call ValidateQuerySchema(sql, tablesInQuery, parsedSql)
        End If
    End If
    ' --- FIM DA VALIDAÇÃO DE ESQUEMA ---
    
    ' --- ROTEAMENTO PARA EXECUÇÃO ---
    If pHelpers.IsSelectSql(sql) Then
        Set result = ExecuteSelect(sql, execReport)
    ElseIf pHelpers.IsInsertSql(sql) Then
        Set result = ExecuteInsert(sql, execReport)
    ElseIf pHelpers.IsDeleteSql(sql) Then
        Set result = ExecuteDelete(sql, execReport)
    ElseIf pHelpers.IsUpdateSql(sql) Then
        Set result = ExecuteUpdate(sql, execReport)
    Else
        Err.Raise vbObjectError + 811, "RunQuery", "Comando SQL não reconhecido para execução."
    End If
    ' --- FIM DO ROTEAMENTO ---

Cleanup: ' Rótulo para limpeza e saída
    On Error Resume Next
    Call pHelpers.UpdateStatus("Limpando tabelas temporárias...")
    Trace "RunQuery: [ETAPA FINAL] Limpando tabelas temporárias..."
    Call CleanupTempTables

    If prevStatesCaptured Then
        Application.EnableEvents = prevEnableEvents
        Application.ScreenUpdating = prevScreenUpdating

        If IsEmpty(prevStatusBar) Then
            Application.StatusBar = False
        ElseIf VarType(prevStatusBar) = vbBoolean Then
            If CBool(prevStatusBar) = False Then
                Application.StatusBar = False
            Else
                Application.StatusBar = CStr(prevStatusBar)
            End If
        Else
            Application.StatusBar = prevStatusBar
        End If
    Else
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.StatusBar = False
    End If
    
    If result Is Nothing Then Set result = CreateObject("Scripting.Dictionary")
    If Not result.Exists("rowsAffected") Then result.Add "rowsAffected", 0
    Call pHelpers.AppendPersistentLog("INFO", "fwXLSPConn", "RunQuery", _
                                      "Execução SQL concluída. rowsAffected=" & CStr(result("rowsAffected")) & ".", sql)
    
    Set RunQuery = result
    Trace "RunQuery: Fim da execução."
    Exit Function

ErrHandler:
    Trace "RunQuery: ERRO " & Err.Number & " - " & Err.Description & " (Source: " & Err.source & ")"
    Call pHelpers.UpdateStatus("Erro durante a execução: " & Err.Description)
    Call pHelpers.AppendPersistentLog("ERROR", "fwXLSPConn", "RunQuery", Err.Description, sql)
    
    Resume Cleanup
End Function

'******************************************************************************
' Nome do elemento: BuildStagingSqlFor (Function)
' Resumo: Retorna o resultado calculado por BuildStagingSqlFor.
' @param listToStage - Parâmetro da rotina.
' @param originalSql - Parâmetro da rotina.
' @param tablesCollection - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function BuildStagingSqlFor(ByVal listToStage As String, ByVal originalSql As String, ByVal tablesCollection As Collection) As String
    On Error GoTo ErrHandler
    
    Dim finalSql As String
    Dim pushdownWhereClause As String
    Dim colsToFetch As Object: Set colsToFetch = CreateObject("Scripting.Dictionary")
    colsToFetch.CompareMode = vbTextCompare
    
    ' --- ETAPA 1: Determinar se a chave primária é OBRIGATÓRIA ---
    Dim isUpdateQuery As Boolean: isUpdateQuery = pHelpers.IsUpdateSql(originalSql)
    Dim spKeyField As String
    
    If isUpdateQuery Then
        ' Para UPDATEs, o ID é essencial para a sincronização posterior.
        ' Obtém o nome da chave primária (normalmente "ID")
        Dim spTemp As New fwSPConn
        Dim spDetailsTemp As Object: Set spDetailsTemp = pCatalog(listToStage)("Connection")
        spTemp.Init spDetailsTemp("SiteURL"), spDetailsTemp("ListID")
        spTemp.TraceOn = Me.TraceOn
        spKeyField = spTemp.spKeyField ' Geralmente retorna "ID"
        Set spTemp = Nothing ' Libera o objeto temporário
        
        colsToFetch.Add spKeyField, True ' Adiciona a chave à lista de coluna necessárias
        Trace "  [BuildStagingSqlFor] Consulta UPDATE detectada. Chave SP '" & spKeyField & "' adicionada à lista de staging."
    End If

    ' --- ETAPA 2: Extrair coluna da consulta original ---

    ' Lógica para consulta de tabela única (geralmente SELECT)
    If InStr(1, originalSql, " JOIN ", vbTextCompare) = 0 Then
        Trace "  [BuildStagingSqlFor] Consulta de tabela única. Extraindo colunas do SELECT (se aplicável)."
        If pHelpers.IsSelectSql(originalSql) Then
            Dim parsedSelect As Object: Set parsedSelect = pHelpers.ParseSqlQueryDetails(originalSql)
            If Not parsedSelect Is Nothing And parsedSelect.Exists("Columns") Then
                Dim columnsInSelect As Collection: Set columnsInSelect = parsedSelect("Columns")
                If columnsInSelect.Count > 0 Then
                    If columnsInSelect(1)("ColumnName") = "*" Then GoTo UseSelectAll ' Se for SELECT *, busca tudo

                    Dim colDict As Object
                    For Each colDict In columnsInSelect
                        ' Adiciona coluna reais, ignorando literais e a chave já adicionada
                        If InStr(colDict("SourceExpression"), "'") = 0 And InStr(colDict("SourceExpression"), "#") = 0 Then
                            If Not colsToFetch.Exists(colDict("ColumnName")) Then
                                colsToFetch.Add colDict("ColumnName"), True
                            End If
                        End If
                    Next
                End If ' End If columnsInSelect.Count > 0
            End If ' End If Not parsedSelect Is Nothing
        Else
             ' Se não for SELECT nem JOIN (ex: UPDATE simples), buscar tudo pode ser mais seguro
             GoTo UseSelectAll
        End If ' End If pHelpers.IsSelectSql(originalSql)
        
    ' Lógica para consulta com JOIN (SELECT ou UPDATE)
    Else
        Trace "  [BuildStagingSqlFor] Consulta JOIN detectada. Extraindo colunas referenciadas."
        Dim targetAlias As String
        Dim tbl As Object
        For Each tbl In tablesCollection
            If StrComp(tbl("Name"), listToStage, vbTextCompare) = 0 Then
                targetAlias = tbl("Alias")
                Exit For
            End If
        Next
        
        If Len(targetAlias) = 0 Then GoTo UseSelectAll ' Se não achou alias, busca tudo
        
        Dim allColumnsByAlias As Object
        Set allColumnsByAlias = pHelpers.ExtractAllReferencedColumns(originalSql)
        
        If allColumnsByAlias.Exists(targetAlias) Then
            Dim requiredColsDict As Object
            Set requiredColsDict = allColumnsByAlias(targetAlias)
            
            Dim colName As Variant
            For Each colName In requiredColsDict.keys
                 If Not colsToFetch.Exists(colName) Then ' Adiciona se não estiver já na lista (como o ID)
                     colsToFetch.Add colName, True
                 End If
            Next colName
        Else
            ' Se o alias não tem coluna explicitamente referenciadas (improvável, mas possível),
            ' buscar tudo é mais seguro.
             GoTo UseSelectAll
        End If ' End If allColumnsByAlias.Exists(targetAlias)
    End If ' End If InStr JOIN

    ' --- ETAPA 3: Construir o SQL Final ---
    pushdownWhereClause = BuildPushdownWhereClauseForStaging(listToStage, originalSql, tablesCollection)
    pushdownWhereClause = ApplyDeltaSyncWhereClause(listToStage, originalSql, pushdownWhereClause)

    If colsToFetch.Count > 0 Then
        finalSql = "SELECT [" & Join(colsToFetch.keys, "], [") & "] FROM [" & listToStage & "]"
        If Len(pushdownWhereClause) > 0 Then finalSql = finalSql & " " & pushdownWhereClause
        Trace "  [BuildStagingSqlFor] SQL de Staging OTIMIZADO: " & finalSql
        BuildStagingSqlFor = finalSql
        Exit Function
    End If

ErrHandler:
    Trace "  [BuildStagingSqlFor] Erro ou condição inesperada. Usando SELECT * como fallback."
UseSelectAll:
    If Len(pushdownWhereClause) = 0 Then
        pushdownWhereClause = BuildPushdownWhereClauseForStaging(listToStage, originalSql, tablesCollection)
        pushdownWhereClause = ApplyDeltaSyncWhereClause(listToStage, originalSql, pushdownWhereClause)
    End If
    finalSql = "SELECT * FROM [" & listToStage & "]"
    If Len(pushdownWhereClause) > 0 Then finalSql = finalSql & " " & pushdownWhereClause
    Trace "  [BuildStagingSqlFor] SQL de Staging (Fallback): " & finalSql
    BuildStagingSqlFor = finalSql
End Function

'******************************************************************************
' Nome do elemento: LoadDeltaSyncConfigFromCatalog (Sub)
' Resumo:
'   Carrega configurações opt-in de Delta Sync a partir das entradas do
'   catálogo SP inicializado no InitSP.
'******************************************************************************
Private Sub LoadDeltaSyncConfigFromCatalog()
    On Error GoTo SafeExit

    If pDeltaSyncConfig Is Nothing Then
        Set pDeltaSyncConfig = CreateObject("Scripting.Dictionary")
        pDeltaSyncConfig.CompareMode = vbTextCompare
    Else
        pDeltaSyncConfig.RemoveAll
    End If

    If pCatalog Is Nothing Then Exit Sub

    Dim logicalName As Variant
    For Each logicalName In pCatalog.keys
        Dim catEntry As Object
        Set catEntry = pCatalog(logicalName)
        If catEntry Is Nothing Then GoTo NextItem
        If Not catEntry.Exists("Type") Then GoTo NextItem
        If StrComp(CStr(catEntry("Type")), "SP", vbTextCompare) <> 0 Then GoTo NextItem

        Dim enabled As Boolean
        Dim modifiedField As String
        Dim stateKey As String
        enabled = False
        modifiedField = "Modified"
        stateKey = CStr(logicalName)

        If catEntry.Exists("DeltaSyncEnabled") Then enabled = CBool(catEntry("DeltaSyncEnabled"))
        If catEntry.Exists("DeltaSyncField") Then
            If Len(Trim$(CStr(catEntry("DeltaSyncField")))) > 0 Then
                modifiedField = Trim$(CStr(catEntry("DeltaSyncField")))
            End If
        End If
        If catEntry.Exists("DeltaSyncStateKey") Then
            If Len(Trim$(CStr(catEntry("DeltaSyncStateKey")))) > 0 Then
                stateKey = Trim$(CStr(catEntry("DeltaSyncStateKey")))
            End If
        End If

        If enabled Then
            Call ConfigureDeltaSync(CStr(logicalName), modifiedField, stateKey, True)
        End If
NextItem:
    Next logicalName
    Exit Sub

SafeExit:
    Err.Clear
End Sub

'******************************************************************************
' Nome do elemento: GetDeltaSyncConfigForList (Function)
' Resumo:
'   Obtém configuração de Delta Sync para uma lista específica.
'******************************************************************************
Private Function GetDeltaSyncConfigForList(ByVal logicalName As String) As Object
    On Error GoTo SafeExit
    If pDeltaSyncConfig Is Nothing Then Exit Function
    If pDeltaSyncConfig.Exists(logicalName) Then
        Set GetDeltaSyncConfigForList = pDeltaSyncConfig(logicalName)
    End If
    Exit Function

SafeExit:
    Set GetDeltaSyncConfigForList = Nothing
    Err.Clear
End Function

'******************************************************************************
' Nome do elemento: ApplyDeltaSyncWhereClause (Function)
' Resumo:
'   Acrescenta filtro incremental por Modified quando Delta Sync estiver ativo.
' @remarks O comportamento é opt-in por lista e aplicado somente em SELECT.
'******************************************************************************
Private Function ApplyDeltaSyncWhereClause( _
    ByVal listToStage As String, _
    ByVal originalSql As String, _
    ByVal currentWhereClause As String _
) As String
    On Error GoTo SafeExit

    ApplyDeltaSyncWhereClause = currentWhereClause

    If Not pHelpers.IsSelectSql(originalSql) Then Exit Function

    Dim cfg As Object
    Set cfg = GetDeltaSyncConfigForList(listToStage)
    If cfg Is Nothing Then Exit Function
    If Not cfg.Exists("Enabled") Then Exit Function
    If CBool(cfg("Enabled")) = False Then Exit Function

    Dim modifiedField As String
    Dim stateKey As String
    modifiedField = CStr(cfg("ModifiedField"))
    stateKey = CStr(cfg("StateKey"))
    If Len(modifiedField) = 0 Then modifiedField = "Modified"
    If Len(stateKey) = 0 Then stateKey = listToStage

    Dim lastSync As Date
    lastSync = ReadDeltaSyncCheckpoint(stateKey)
    If lastSync <= 0 Then
        Trace "  [DeltaSync] Sem checkpoint prévio para '" & listToStage & "'. Executando carga completa."
        Exit Function
    End If

    Dim predicate As String
    predicate = "[" & pHelpers.Unbracket(modifiedField) & "] > " & pHelpers.FormatDateLiteral(lastSync)

    ApplyDeltaSyncWhereClause = AppendWherePredicate(currentWhereClause, predicate)
    Trace "  [DeltaSync] Filtro incremental aplicado em '" & listToStage & "': " & predicate
    Exit Function

SafeExit:
    ApplyDeltaSyncWhereClause = currentWhereClause
    Err.Clear
End Function

'******************************************************************************
' Nome do elemento: AppendWherePredicate (Function)
' Resumo:
'   Combina predicado adicional com cláusula WHERE existente.
'******************************************************************************
Private Function AppendWherePredicate(ByVal whereClause As String, ByVal predicate As String) As String
    Dim normalizedWhere As String
    normalizedWhere = Trim$(whereClause)
    predicate = Trim$(predicate)
    If Len(predicate) = 0 Then
        AppendWherePredicate = normalizedWhere
        Exit Function
    End If

    If Len(normalizedWhere) = 0 Then
        AppendWherePredicate = "WHERE " & predicate
        Exit Function
    End If

    If Right$(normalizedWhere, 1) = ";" Then normalizedWhere = Left$(normalizedWhere, Len(normalizedWhere) - 1)
    If UCase$(Left$(normalizedWhere, 5)) <> "WHERE" Then
        normalizedWhere = "WHERE " & normalizedWhere
    End If
    AppendWherePredicate = normalizedWhere & " AND " & predicate
End Function

'******************************************************************************
' Nome do elemento: RegisterDeltaSyncSuccess (Sub)
' Resumo:
'   Atualiza checkpoint persistente de Delta Sync após staging bem-sucedido.
'******************************************************************************
Private Sub RegisterDeltaSyncSuccess(ByVal logicalName As String, ByVal originalSql As String)
    On Error GoTo SafeExit

    If Not pHelpers.IsSelectSql(originalSql) Then Exit Sub

    Dim cfg As Object
    Set cfg = GetDeltaSyncConfigForList(logicalName)
    If cfg Is Nothing Then Exit Sub
    If Not cfg.Exists("Enabled") Then Exit Sub
    If CBool(cfg("Enabled")) = False Then Exit Sub

    Dim stateKey As String
    stateKey = CStr(cfg("StateKey"))
    If Len(stateKey) = 0 Then stateKey = logicalName

    Call WriteDeltaSyncCheckpoint(stateKey, now)
    Call pHelpers.AppendPersistentLog("INFO", "fwXLSPConn", "RegisterDeltaSyncSuccess", _
                                      "Checkpoint atualizado para '" & logicalName & "'.", _
                                      "", stateKey)
    Exit Sub

SafeExit:
    Err.Clear
End Sub

'******************************************************************************
' Nome do elemento: EnsureDeltaSyncStateSheet (Function)
' Resumo:
'   Garante planilha oculta para armazenamento de checkpoints de Delta Sync.
'******************************************************************************
Private Function EnsureDeltaSyncStateSheet() As Worksheet
    On Error GoTo SafeExit

    If ActiveWorkbook Is Nothing Then Exit Function

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook

    On Error Resume Next
    Set ws = wb.Worksheets(DELTASYNCSTATESHEETNAME)
    On Error GoTo SafeExit

    If ws Is Nothing Then
        Dim oldDisplayAlerts As Boolean
        oldDisplayAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False

        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = DELTASYNCSTATESHEETNAME
        ws.Visible = xlSheetVeryHidden
        ws.Range("A1:B1").value = Array("StateKey", "LastSyncAt")
        ws.rows(1).Font.Bold = True
        ws.Columns("A:B").EntireColumn.AutoFit

        Application.DisplayAlerts = oldDisplayAlerts
    End If

    Set EnsureDeltaSyncStateSheet = ws
    Exit Function

SafeExit:
    Set EnsureDeltaSyncStateSheet = Nothing
    Err.Clear
End Function

'******************************************************************************
' Nome do elemento: ReadDeltaSyncCheckpoint (Function)
' Resumo:
'   Lê o timestamp de checkpoint para uma chave de estado.
'******************************************************************************
Private Function ReadDeltaSyncCheckpoint(ByVal stateKey As String) As Date
    On Error GoTo SafeExit
    If Len(Trim$(stateKey)) = 0 Then Exit Function

    Dim ws As Worksheet
    Set ws = EnsureDeltaSyncStateSheet()
    If ws Is Nothing Then Exit Function

    Dim rowHit As Long
    rowHit = FindDeltaSyncStateRow(ws, stateKey)
    If rowHit = 0 Then Exit Function

    Dim rawValue As Variant
    rawValue = ws.Cells(rowHit, 2).value
    If IsDate(rawValue) Then
        ReadDeltaSyncCheckpoint = CDate(rawValue)
    End If
    Exit Function

SafeExit:
    ReadDeltaSyncCheckpoint = 0
    Err.Clear
End Function

'******************************************************************************
' Nome do elemento: WriteDeltaSyncCheckpoint (Sub)
' Resumo:
'   Persiste o timestamp de checkpoint para reutilização em execuções futuras.
'******************************************************************************
Private Sub WriteDeltaSyncCheckpoint(ByVal stateKey As String, ByVal syncMoment As Date)
    On Error GoTo SafeExit
    If Len(Trim$(stateKey)) = 0 Then Exit Sub

    Dim ws As Worksheet
    Set ws = EnsureDeltaSyncStateSheet()
    If ws Is Nothing Then Exit Sub

    Dim rowHit As Long
    rowHit = FindDeltaSyncStateRow(ws, stateKey)
    If rowHit = 0 Then
        rowHit = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
        If rowHit < 2 Then rowHit = 2
    End If

    ws.Cells(rowHit, 1).value = stateKey
    ws.Cells(rowHit, 2).value = syncMoment
    ws.Cells(rowHit, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    Exit Sub

SafeExit:
    Err.Clear
End Sub

'******************************************************************************
' Nome do elemento: FindDeltaSyncStateRow (Function)
' Resumo:
'   Localiza a linha de uma chave no armazenamento de checkpoints.
'******************************************************************************
Private Function FindDeltaSyncStateRow(ByVal ws As Worksheet, ByVal stateKey As String) As Long
    On Error GoTo SafeExit
    If ws Is Nothing Then Exit Function
    If Len(Trim$(stateKey)) = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim r As Long
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 1).value)), Trim$(stateKey), vbTextCompare) = 0 Then
            FindDeltaSyncStateRow = r
            Exit Function
        End If
    Next r
    Exit Function

SafeExit:
    FindDeltaSyncStateRow = 0
    Err.Clear
End Function

'******************************************************************************
' Nome do elemento: BuildPushdownWhereClauseForStaging (Function)
' Resumo: Avalia se a cláusula WHERE original pode ser aplicada no staging da
'         lista alvo sem introduzir dependências de outras tabelas.
'******************************************************************************
Private Function BuildPushdownWhereClauseForStaging( _
    ByVal listToStage As String, _
    ByVal originalSql As String, _
    ByVal tablesCollection As Collection _
) As String
    On Error GoTo SafeExit

    Dim whereClause As String
    Dim hasJoin As Boolean
    hasJoin = (InStr(1, originalSql, " JOIN ", vbTextCompare) > 0)
    If Not tablesCollection Is Nothing Then hasJoin = hasJoin Or (tablesCollection.Count > 1)

    If pHelpers.IsSelectSql(originalSql) Then
        whereClause = ExtractSimpleWhereClauseFromSelect(originalSql)
    ElseIf pHelpers.IsUpdateSql(originalSql) Then
        Dim parsedUpdate As Object
        Set parsedUpdate = pHelpers.ParseUpdateJoinSqlDetails(originalSql)
        If Not parsedUpdate Is Nothing Then
            If parsedUpdate.Exists("WhereClause") Then whereClause = Trim$(CStr(parsedUpdate("WhereClause")))
            If parsedUpdate.Exists("Tables") Then hasJoin = hasJoin Or (parsedUpdate("Tables").Count > 1)
        End If
    ElseIf pHelpers.IsDeleteSql(originalSql) Then
        Dim parsedDelete As Object
        Set parsedDelete = pHelpers.ParseDeleteSqlDetails(originalSql)
        If Not parsedDelete Is Nothing Then
            If parsedDelete.Exists("WhereClause") Then whereClause = Trim$(CStr(parsedDelete("WhereClause")))
        End If
    End If

    If Len(whereClause) = 0 Then Exit Function
    If Right$(whereClause, 1) = ";" Then whereClause = Left$(whereClause, Len(whereClause) - 1)
    If UCase$(Left$(whereClause, 5)) <> "WHERE" Then whereClause = "WHERE " & whereClause

    If pHelpers.IsWhereClauseComplex(whereClause) Then
        Trace "  [BuildStagingSqlFor] Pushdown de WHERE ignorado por complexidade: " & whereClause
        Exit Function
    End If

    If hasJoin Then
        Dim targetAlias As String
        Dim tbl As Object
        For Each tbl In tablesCollection
            If StrComp(CStr(tbl("Name")), listToStage, vbTextCompare) = 0 Then
                targetAlias = CStr(tbl("Alias"))
                Exit For
            End If
        Next tbl
        If Len(targetAlias) = 0 Then Exit Function

        Dim aliasDeps As Object
        Set aliasDeps = pHelpers.GetWhereClauseDependencies(whereClause)
        If aliasDeps Is Nothing Then Exit Function

        ' Em JOIN, só empurra se o WHERE depender exclusivamente do alias da lista.
        If aliasDeps.Count = 0 Then Exit Function

        Dim aliasName As Variant
        For Each aliasName In aliasDeps.keys
            If StrComp(CStr(aliasName), targetAlias, vbTextCompare) <> 0 Then Exit Function
        Next aliasName
    End If

    BuildPushdownWhereClauseForStaging = whereClause
    Exit Function

SafeExit:
    BuildPushdownWhereClauseForStaging = vbNullString
End Function

'******************************************************************************
' Nome do elemento: ExtractSimpleWhereClauseFromSelect (Function)
' Resumo: Extrai uma cláusula WHERE simples de SELECT para pushdown no staging.
'******************************************************************************
Private Function ExtractSimpleWhereClauseFromSelect(ByVal sql As String) As String
    On Error GoTo SafeExit

    Dim upperSql As String
    upperSql = UCase$(sql)

    If InStr(1, upperSql, " GROUP BY ", vbTextCompare) > 0 Then Exit Function
    If InStr(1, upperSql, " ORDER BY ", vbTextCompare) > 0 Then Exit Function
    If InStr(1, upperSql, " HAVING ", vbTextCompare) > 0 Then Exit Function
    If InStr(1, upperSql, " UNION ", vbTextCompare) > 0 Then Exit Function

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .IgnoreCase = True
        .Global = False
        .Pattern = "\bWHERE\b\s+([\s\S]+)$"
    End With

    Dim matches As Object
    Set matches = regEx.Execute(sql)
    If matches.Count = 0 Then Exit Function

    Dim whereClause As String
    whereClause = Trim$(CStr(matches(0).SubMatches(0)))
    If Len(whereClause) = 0 Then Exit Function
    If Right$(whereClause, 1) = ";" Then whereClause = Left$(whereClause, Len(whereClause) - 1)

    If UCase$(Left$(whereClause, 5)) <> "WHERE" Then whereClause = "WHERE " & whereClause
    ExtractSimpleWhereClauseFromSelect = whereClause
    Exit Function

SafeExit:
    ExtractSimpleWhereClauseFromSelect = vbNullString
End Function

'******************************************************************************
' Nome do elemento: BuildSqlInListFromKeys (Function)
' Resumo: Constrói uma lista SQL para cláusula IN com tratamento de tipo.
'******************************************************************************
Private Function BuildSqlInListFromKeys(ByVal keys As Variant, ByVal startIndex As Long, ByVal endIndex As Long) As String
    On Error GoTo SafeExit

    Dim itemCount As Long
    itemCount = endIndex - startIndex + 1
    If itemCount <= 0 Then Exit Function

    Dim sqlValues() As String
    ReDim sqlValues(0 To itemCount - 1)

    Dim i As Long
    For i = 0 To itemCount - 1
        sqlValues(i) = pHelpers.FormatSQLValue(keys(startIndex + i), True)
    Next i

    BuildSqlInListFromKeys = Join(sqlValues, ",")
    Exit Function

SafeExit:
    BuildSqlInListFromKeys = vbNullString
End Function

'******************************************************************************
' Nome do elemento: CreateStagingTable (Sub)
' Resumo: Executa a rotina CreateStagingTable.
' @param tableName - Parâmetro da rotina.
' @param dataResultAsObject - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Sub CreateStagingTable(ByVal tableName As String, ByVal dataResultAsObject As Object)
    On Error GoTo ErrHandler
    
    Dim stagingSheet As Worksheet, sheetName As String
    sheetName = "Staging_" & tableName & "_" & Format(now, "hhmmss")
    
    ' 1. Cria a planilha oculta
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets(sheetName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True
    
    Set stagingSheet = ActiveWorkbook.Worksheets.Add(Before:=ActiveWorkbook.activeSheet)
    stagingSheet.Name = sheetName
    stagingSheet.Visible = xlSheetVeryHidden
    
    ' 2. Obtém dados, headers e types
    Dim dataArr As Variant
    Dim headersArr() As String
    Dim typesArr As Variant
    Dim hasData As Boolean
    
    If dataResultAsObject.Exists("Headers") And IsArray(dataResultAsObject("Headers")) Then
        headersArr = dataResultAsObject("Headers")
    Else
        Err.Raise vbObjectError + 860, "CreateStagingTable", "Headers not found or not an array for '" & tableName & "'."
    End If
     If dataResultAsObject.Exists("Types") And IsArray(dataResultAsObject("Types")) Then
        typesArr = dataResultAsObject("Types")
    Else
        Err.Raise vbObjectError + 861, "CreateStagingTable", "Types not found or not an array for '" & tableName & "'."
    End If
    If dataResultAsObject.Exists("Data") And Not pHelpers.IsVarEmpty(dataResultAsObject("Data")) Then
        dataArr = dataResultAsObject("Data")
        hasData = True
    End If

    If LBound(headersArr) > UBound(headersArr) Then
        Err.Raise vbObjectError + 862, "CreateStagingTable", "Array de Headers está vazio para '" & tableName & "'."
    End If

    ' 3. Escreve APENAS os cabeçalhos na planilha A1
    Dim headerRange As Range
    Set headerRange = stagingSheet.Range("A1").Resize(1, UBound(headersArr) - LBound(headersArr) + 1)
    headerRange.value = headersArr
    
    ' 4. Aplica a formatação às COLUNA INTEIRAS
    Trace "[DEBUG] CreateStagingTable: Aplicando pré-formatação às colunas inteiras em '" & stagingSheet.Name & "'..."
    Dim i As Long
    Dim adoType As Long
    Dim excelFormat As String
    Dim targetCol As Range
        
    For i = LBound(headersArr) To UBound(headersArr)
        adoType = -1
        excelFormat = "@" ' Padrão Texto
        
        On Error Resume Next
        adoType = typesArr(i)
        On Error GoTo ErrHandler
            
        If adoType <> -1 Then
            Select Case adoType
                Case 3: excelFormat = "0"
                Case 5: excelFormat = "0.###############"
                Case 7, 135: excelFormat = "dd/mm/yyyy" ' Ajuste formato se necessário
                Case Else: excelFormat = "@"
            End Select
            
            Set targetCol = headerRange.Cells(1, i - LBound(headersArr) + 1).EntireColumn
            On Error Resume Next
            targetCol.NumberFormat = excelFormat
            If Err.Number <> 0 Then
                Trace "[DEBUG] CreateStagingTable: Aviso - Falha ao aplicar formato '" & excelFormat & "' à coluna '" & headersArr(i) & "'. Erro: " & Err.Description
                Err.Clear
            Else
                Trace "[DEBUG] CreateStagingTable: Coluna '" & headersArr(i) & "' pré-formatada como " & excelFormat
            End If
            On Error GoTo ErrHandler
        Else
            Trace "[DEBUG] CreateStagingTable: Aviso - Tipo inválido para coluna '" & headersArr(i) & "'."
        End If
    Next i
    Trace "[DEBUG] CreateStagingTable: Pré-formatação de colunas concluída."

    ' --- INÍCIO DA CORREÇÃO ---
    ' 5. CRIA o ListObject MÍNIMO (APENAS CABEÇALHO)
    Dim newLo As ListObject
    On Error Resume Next ' Trata erro na criação do ListObject
    ' Cria o ListObject usando SOMENTE o headerRange como source
    Set newLo = stagingSheet.ListObjects.Add(sourceType:=xlSrcRange, source:=headerRange, XlListObjectHasHeaders:=xlYes)
    If Err.Number <> 0 Then
        Trace "[DEBUG] CreateStagingTable: ERRO ao criar ListObject '" & tableName & "' no range " & headerRange.Address & ". Erro: " & Err.Description
        Err.Raise Err.Number, "CreateStagingTable", "Erro ao criar ListObject '" & tableName & "'. " & Err.Description
    End If
    On Error GoTo ErrHandler
    newLo.Name = tableName
    If Not pTempTables.Exists(tableName) Then pTempTables.Add tableName, sheetName
    Trace "[DEBUG] CreateStagingTable: ListObject '" & tableName & "' criado apenas com cabeçalho."
    ' --- FIM DA CORREÇÃO ---

    ' 6. Carrega os dados (se houver) usando XLLoadArray
    '    Como o ListObject foi criado vazio (ListRows.Count=0), XLLoadArray
    '    usará a lógica de Resize + Escrita Direta, que é o correto aqui.
    If hasData Then
        Trace "[DEBUG] CreateStagingTable: Carregando dados em '" & tableName & "' via XLLoadArray..."
        Dim loader As New fwXLTable
        loader.TraceOn = Me.TraceOn
        loader.Init tableName
        ' Chama XLLoadArray SEM parâmetros de limpeza/título.
        ' A lógica interna dele fará o resize correto para a carga inicial.
        loader.XLLoadArray dataArr, headersArr ', ClearData:=False, ClearTitle:=False, ResizeTitle:=False ' Parâmetros não são necessários aqui
        Trace "[DEBUG] CreateStagingTable: Carregamento inicial de dados concluído."
    Else
        Trace "[DEBUG] CreateStagingTable: Nenhum dado para carregar em '" & tableName & "'."
        ' A tabela já está vazia (0 linhas de dados), então não precisa fazer nada.
    End If

    Trace "CreateStagingTable: Tabela '" & tableName & "' criada, formatada e populada na planilha '" & sheetName & "'."
    Exit Sub

ErrHandler:
    Trace "[DEBUG] CreateStagingTable: ERRO " & Err.Number & " - " & Err.Description & ". Tentando limpar planilha de staging..."
    Application.DisplayAlerts = False
    On Error Resume Next
    If Not stagingSheet Is Nothing Then stagingSheet.Delete
    If pTempTables.Exists(tableName) Then pTempTables.Remove tableName
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    On Error GoTo 0
    Err.Raise Err.Number, "fwXLSPConn.CreateStagingTable", Err.Description
End Sub

'******************************************************************************
' Nome do elemento: CleanupTempTables (Sub)
' Resumo: Executa a rotina CleanupTempTables.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub CleanupTempTables()
    Const SENHA_ESTRUTURA As String = ""
    
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet
    Dim sheetName As Variant
    Dim wasProtected As Boolean
    
    On Error GoTo RestoreSettings
    
    If pTempTables Is Nothing Or pTempTables.Count = 0 Then Exit Sub
    
    Trace "CleanupTempTables: Limpando " & pTempTables.Count & " tabela(s) temporária(s)..."
    
    ' Apenas o controle de Alertas e Proteção permanece aqui
    Application.DisplayAlerts = False
    
    If wb.ProtectStructure Then
        wasProtected = True
        wb.Unprotect SENHA_ESTRUTURA
    End If
    
    For Each sheetName In pTempTables.Items
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(sheetName))
        On Error GoTo RestoreSettings
        
        If Not ws Is Nothing Then
            ws.Visible = xlSheetHidden
            If ws.ListObjects.Count > 0 Then
                ws.ListObjects(1).Delete
            End If
            ws.Delete
        End If
        Set ws = Nothing
    Next
    
    pTempTables.RemoveAll

RestoreSettings:
    If wasProtected Then
        wb.Protect SENHA_ESTRUTURA, Structure:=True
    End If
    
    Application.DisplayAlerts = True
    
    If Err.Number <> 0 Then
        Trace "CleanupTempTables: Ocorreu um erro durante a limpeza: " & Err.Description
        Err.Clear
    Else
        Trace "CleanupTempTables: Limpeza concluída com sucesso."
    End If
End Sub

'******************************************************************************
' Nome do elemento: ExecuteSelect (Function)
' Resumo: Retorna o resultado calculado por ExecuteSelect.
' @param sql - Parâmetro da rotina.
' @param execReport - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Function ExecuteSelect(ByVal sql As String, Optional ByRef execReport As Variant) As Object
    Dim result As Object
    On Error GoTo ErrHandler
    
    Trace "ExecuteSelect ENTER; SQLLen=" & Len(sql)
    
    ' ETAPA 1: ANÁLISE INICIAL
    Call pHelpers.UpdateStatus("Analisando consulta SELECT...")
    Trace "ExecuteSelect: [ETAPA 1] Analisando SQL..."
    Dim parsedSql As Object: Set parsedSql = pHelpers.ParseSqlQueryDetails(sql)
    If parsedSql Is Nothing Then Err.Raise vbObjectError + 812, "ExecuteSelect", "Falha ao analisar o comando SQL."
    
    Dim tablesInQuery As Collection: Set tablesInQuery = parsedSql("Tables")
    
    Dim allFoundTables As Object: Set allFoundTables = pHelpers.ExtractAllTableNamesFromClauses(sql)
    Dim tempDict As Object: Set tempDict = CreateObject("Scripting.Dictionary")
    tempDict.CompareMode = vbTextCompare
    Dim tbl As Object
    For Each tbl In tablesInQuery
        If Not tempDict.Exists(tbl("Name")) Then tempDict.Add tbl("Name"), tbl("Alias")
    Next
    Dim key As Variant
    For Each key In allFoundTables.keys
        If Not tempDict.Exists(key) Then tempDict.Add key, key
    Next
    Set tablesInQuery = New Collection
    For Each key In tempDict.keys
        Dim tblDict As Object: Set tblDict = CreateObject("Scripting.Dictionary")
        tblDict.Add "Name", key
        tblDict.Add "Alias", tempDict(key)
        tablesInQuery.Add tblDict
    Next
    
    If tablesInQuery.Count = 0 Then Err.Raise vbObjectError + 813, "ExecuteSelect", "Nenhuma tabela encontrada no SQL."
    
    Dim spListsInQuery As Object: Set spListsInQuery = CreateObject("Scripting.Dictionary")
    spListsInQuery.CompareMode = vbTextCompare
    Dim hasXLTable As Boolean
    
    Dim tblName As String, i As Long
    For i = 1 To tablesInQuery.Count
        Set tbl = tablesInQuery(i)
        tblName = tbl("Name")
        Dim sourceType As String: sourceType = pCatalog(tblName)("Type")
        If sourceType = "SP" Then
            If Not spListsInQuery.Exists(tblName) Then spListsInQuery.Add tblName, True
        Else
            hasXLTable = True
        End If
    Next
    
    ' ETAPA 2: DECISÃO DA ESTRATÉGIA
    Dim strategy As String
    If spListsInQuery.Count = 0 Then
        strategy = "Direct-XL"
    ElseIf spListsInQuery.Count <= 1 And Not hasXLTable Then
        strategy = "Direct-SP"
    Else
        strategy = "Staging"
    End If
    Call pHelpers.UpdateStatus("Estratégia decidida: " & strategy)
    
    ' --- INÍCIO DO BLOCO DE FORMATAÇÃO E DIAGNÓSTICO ---
    If InStr(1, sql, " JOIN ", vbTextCompare) > 0 And hasXLTable Then
        Call pHelpers.UpdateStatus("Preparando tabelas para a junção...")
        
        Dim joinCols As Object: Set joinCols = pHelpers.ParseOnClauseColumns(sql)
        If joinCols.Count = 2 Then
            Dim colsToFormat(1) As String
            colsToFormat(0) = joinCols("LeftColumn")
            colsToFormat(1) = joinCols("RightColumn")
            
            Trace "[DIAGNÓSTICO] ==> INICIANDO LOOP DE FORMATAÇÃO DE COLUNAS DE JUNÇÃO"
            Dim colFullName As Variant
            For Each colFullName In colsToFormat
                Trace "[DIAGNÓSTICO] Processando colFullName: '" & colFullName & "' (Tipo: " & TypeName(colFullName) & ")"
                
                ' Parse seguro de alias.[coluna] (aceita ponto no nome da coluna)
                Dim tblAlias As String
                Dim colName As String
                If Not pHelpers.TryParseQualifiedColumnRef(CStr(colFullName), tblAlias, colName) Then
                    Trace "[DIAGNOSTICO]   -> AVISO: referencia de coluna invalida no ON: '" & CStr(colFullName) & "'. Pulando."
                    GoTo FinalizeJoinColumn
                End If

                Trace "[DIAGNOSTICO]   -> Alias extraido: '" & tblAlias & "'"
                Trace "[DIAGNOSTICO]   -> Nome da coluna extraido: '" & colName & "'"
                Dim realTableName As String
                For Each tbl In tablesInQuery
                    If StrComp(tbl("Alias"), tblAlias, vbTextCompare) = 0 Then
                        realTableName = tbl("Name")
                        Exit For
                    End If
                Next
                Trace "[DIAGNÓSTICO]   -> Nome real da tabela encontrado: '" & realTableName & "'"
                
                If Len(realTableName) > 0 And pCatalog(realTableName)("Type") = "XL" Then
                    Trace "[DIAGNÓSTICO]     -> Tabela é do tipo XL. Tentando formatar..."
                    On Error Resume Next
                    Dim xlTable As ListObject, ws As Worksheet
                    Set xlTable = Nothing
                    For Each ws In ActiveWorkbook.Worksheets
                        Set xlTable = ws.ListObjects(realTableName)
                        If Not xlTable Is Nothing Then Exit For
                    Next ws
                    
                    If Not xlTable Is Nothing Then
                        Dim keyCol As ListColumn: Set keyCol = xlTable.ListColumns(colName)
                        If Not keyCol Is Nothing Then
                            If Not keyCol.DataBodyRange Is Nothing Then
                                Trace "[DIAGNÓSTICO]       -> Formato ANTES: " & keyCol.DataBodyRange.Cells(1, 1).NumberFormat
                                keyCol.DataBodyRange.NumberFormat = "0"
                                Trace "[DIAGNÓSTICO]       -> Formato DEPOIS: " & keyCol.DataBodyRange.Cells(1, 1).NumberFormat
                            Else
                                Trace "[DIAGNÓSTICO]       -> AVISO: DataBodyRange para a coluna '" & colName & "' está vazio."
                            End If
                        Else
                             Trace "[DIAGNÓSTICO]       -> ERRO: Coluna '" & colName & "' não encontrada na tabela '" & realTableName & "'."
                        End If
                    Else
                        Trace "[DIAGNÓSTICO]     -> ERRO: Tabela '" & realTableName & "' não encontrada no Workbook."
                    End If
                    On Error GoTo ErrHandler
                Else
                    Trace "[DIAGNÓSTICO]     -> Tabela não é do tipo XL. Formatação ignorada."
                End If
FinalizeJoinColumn:
                Trace "[DIAGNÓSTICO] Fim do processamento para '" & colFullName & "'"
            Next colFullName
            Trace "[DIAGNÓSTICO] <== FIM DO LOOP DE FORMATAÇÃO"
        End If
    End If
    
    ' ETAPA 3: EXECUÇÃO DA ESTRATÉGIA
    Select Case strategy
    Case "Direct-XL"
        Call pHelpers.UpdateStatus("Executando consulta diretamente no Excel...")
        Dim xl As New fwXLConn
        xl.TraceOn = Me.TraceOn
        xl.Init
        Set result = xl.XLRunQuery(sql, execReport)
        
    Case "Direct-SP"
        Call pHelpers.UpdateStatus("Executando consulta diretamente no SharePoint...")
        Dim listName As String: listName = spListsInQuery.keys()(0)
        Dim spDetails As Object: Set spDetails = pCatalog(listName)("Connection")
        Dim sp As New fwSPConn
        sp.TraceOn = Me.TraceOn
        sp.Init spDetails("SiteURL"), spDetails("ListID")
        Set result = sp.SPRunQuery(sql, execReport)
        
    Case "Staging"
        Trace "ExecuteSelect: Iniciando processo de staging..."
        Dim spListName As Variant
        For Each spListName In spListsInQuery.keys
            Call pHelpers.UpdateStatus("Baixando dados da lista '" & spListName & "'...")
            Dim spConnDetails As Object: Set spConnDetails = pCatalog(spListName)("Connection")
            Dim spConn As New fwSPConn
            spConn.TraceOn = Me.TraceOn
            spConn.Init spConnDetails("SiteURL"), spConnDetails("ListID")
            Dim stagingSql As String: stagingSql = BuildStagingSqlFor(spListName, sql, tablesInQuery)
            Dim spData As Object: Set spData = spConn.SPRunQuery(stagingSql)
            Call CreateStagingTable(spListName, spData)
            Call RegisterDeltaSyncSuccess(CStr(spListName), sql)
        Next
        
        Call pHelpers.UpdateStatus("Executando consulta final nos dados locais...")
        Dim xlQuery As New fwXLConn
        xlQuery.TraceOn = Me.TraceOn
        xlQuery.Init
        Set result = xlQuery.XLRunQuery(sql, execReport)
    End Select

    Call pHelpers.UpdateStatus("Operação SELECT concluída.")
    Set ExecuteSelect = result
    Exit Function
    
ErrHandler:
    Trace "ExecuteSelect: ERRO " & Err.Number & " - " & Err.Description
    Set result = CreateObject("Scripting.Dictionary")
    Set ExecuteSelect = result
End Function

'******************************************************************************
' Nome do elemento: ExecuteInsert (Function)
' Resumo: Retorna o resultado calculado por ExecuteInsert.
' @param sql - Parâmetro da rotina.
' @param execReport - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Function ExecuteInsert(ByVal sql As String, ByRef execReport As Variant) As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim rowsAffected As Long
    Dim parsedSql As Object
    Dim targetTable As String
    Dim targetType As String
    Dim spDetails As Object
    Dim sp As fwSPConn
    
    On Error GoTo ErrHandler ' Adicionado handler de erro

    Call pHelpers.UpdateStatus("Analisando comando INSERT...")
    Set parsedSql = pHelpers.ParseInsertSqlDetails(sql)
    If parsedSql Is Nothing Then Err.Raise vbObjectError + 820, "ExecuteInsert", "Falha ao analisar o comando INSERT."
    
    targetTable = parsedSql("TargetTable")
    If Not pCatalog.Exists(targetTable) Then Err.Raise vbObjectError + 821, "ExecuteInsert", "Tabela de destino '" & targetTable & "' não encontrada no catálogo."
    
    targetType = pCatalog(targetTable)("Type")
    
    Select Case parsedSql("InsertType")
        
        Case "VALUES"
            Trace "ExecuteInsert: Estratégia 'INSERT VALUES' detectada."
            
            If targetType = "SP" Then
                Call pHelpers.UpdateStatus("Executando inserção no SharePoint. Aguarde...")
                Set spDetails = pCatalog(targetTable)("Connection")
                Set sp = New fwSPConn
                sp.TraceOn = Me.TraceOn
                sp.Init spDetails("SiteURL"), spDetails("ListID")
                rowsAffected = sp.SPExecuteNonQuery(sql)
            Else ' "XL"
                Call pHelpers.UpdateStatus("Executando INSERT VALUES na tabela Excel...")
                Dim xl As New fwXLConn
                xl.TraceOn = Me.TraceOn
                xl.Init
                Dim xlResult As Object: Set xlResult = xl.XLRunQuery(sql, execReport)
                If xlResult.Exists("rowsAffected") Then
                    rowsAffected = xlResult("rowsAffected")
                End If
            End If

        Case "SELECT"
            Trace "ExecuteInsert: Estratégia 'INSERT SELECT' detectada."
            Call pHelpers.UpdateStatus("Iniciando operação INSERT...SELECT...")
            
            Dim selectSubQuery As String: selectSubQuery = parsedSql("SelectSubQuery")
            Trace "ExecuteInsert: Executando subconsulta SELECT..."
            
            Dim selectResult As Object: Set selectResult = ExecuteSelect(selectSubQuery, execReport) ' Reutiliza ExecuteSelect
            
            If Not selectResult.Exists("Data") Or pHelpers.IsVarEmpty(selectResult("Data")) Then ' Verifica se Data existe E não está vazio
                Trace "ExecuteInsert: A subconsulta SELECT não retornou dados. Nenhuma linha inserida."
                rowsAffected = 0
            Else
                Dim dataArr As Variant: dataArr = selectResult("Data")
                Dim headersArr As Variant ' Pode ser String Array ou Variant
                headersArr = selectResult("Headers") ' Obtém headers do resultado do select
                Dim columnsToInsert As Variant: columnsToInsert = parsedSql("Columns") ' Coluna alvo do INSERT
                
                Dim rowCountToInsert As Long
                On Error Resume Next ' Trata erro se dataArr não for 2D
                rowCountToInsert = UBound(dataArr, 2) - LBound(dataArr, 2) + 1
                If Err.Number <> 0 Then rowCountToInsert = 1 ' Assume 1 se for 1D
                Err.Clear
                On Error GoTo ErrHandler ' Restaura handler
                
                Call pHelpers.UpdateStatus("Dados da subconsulta obtidos. Carregando " & rowCountToInsert & " linha(s)...")

                ' Se o INSERT especificou coluna, usa esses nomes como cabeçalho para o LoadArray/SPInsertRows
                If IsArray(columnsToInsert) Then
                    If Not pHelpers.IsVarEmpty(columnsToInsert) Then ' Verifica se o array não está vazio
                        ' Validação de contagem de coluna
                        If UBound(columnsToInsert) - LBound(columnsToInsert) + 1 <> UBound(headersArr) - LBound(headersArr) + 1 Then
                             Err.Raise vbObjectError + 825, "ExecuteInsert", "O número de colunas da subconsulta SELECT (" & (UBound(headersArr) + 1) & ") não corresponde à lista de colunas do INSERT (" & (UBound(columnsToInsert) + 1) & ")."
                        End If
                        headersArr = columnsToInsert ' Usa os nomes das coluna do INSERT
                        Trace "[DEBUG] ExecuteInsert: Cabeçalhos para carregamento definidos pelas colunas do INSERT."
                    End If
                End If

                Trace "ExecuteInsert: Carregando " & rowCountToInsert & " linha(s) na tabela de destino '" & targetTable & "' (" & targetType & ")..."
                If targetType = "SP" Then
                    Set spDetails = pCatalog(targetTable)("Connection")
                    Set sp = New fwSPConn
                    sp.TraceOn = Me.TraceOn
                    sp.Init spDetails("SiteURL"), spDetails("ListID")
                    Dim spInsertResult As Variant: spInsertResult = sp.SPInsertRows(dataArr, headersArr)
                    rowsAffected = spInsertResult(1) ' Pega o total afetado do resultado
                Else ' "XL"
                    Dim xlTbl As New fwXLTable
                    xlTbl.TraceOn = Me.TraceOn
                    xlTbl.Init targetTable
                    ' *** CORREÇÃO AQUI: Remove o terceiro argumento ***
                    xlTbl.XLLoadArray dataArr, headersArr
                    rowsAffected = rowCountToInsert ' Assume sucesso se não houver erro
                End If
            End If
    End Select
    
    Call pHelpers.UpdateStatus("Operação INSERT concluída.")
    result.Add "rowsAffected", rowsAffected
    Set ExecuteInsert = result
    Exit Function

ErrHandler: ' Handler de erro para a função ExecuteInsert
    Trace "ExecuteInsert: ERRO " & Err.Number & " - " & Err.Description
    ' Preenche o Result com erro e sai
    If result Is Nothing Then Set result = CreateObject("Scripting.Dictionary")
    If Not result.Exists("rowsAffected") Then result.Add "rowsAffected", 0
    result.Add "errorMessage", Err.Description ' Adiciona a mensagem de erro
    result.Add "errorNumber", Err.Number
    Set ExecuteInsert = result
    ' NÃO re-lança o erro, permite que RunQuery finalize a limpeza
End Function

'******************************************************************************
' Nome do elemento: ExecuteDelete (Function)
' Resumo: Retorna o resultado calculado por ExecuteDelete.
' @param sql - Parâmetro da rotina.
' @param execReport - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Function ExecuteDelete(ByVal sql As String, ByRef execReport As Variant) As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim rowsAffected As Long
    
    Call pHelpers.UpdateStatus("Analisando comando DELETE...")
    Dim parsedSql As Object
    Set parsedSql = pHelpers.ParseDeleteSqlDetails(sql)
    If parsedSql Is Nothing Then Err.Raise vbObjectError + 830, "ExecuteDelete", "Falha ao analisar o comando DELETE."
    
    Dim targetTable As String: targetTable = parsedSql("TargetTable")
    If Not pCatalog.Exists(targetTable) Then Err.Raise vbObjectError + 831, "ExecuteDelete", "Tabela de destino '" & targetTable & "' não encontrada no catálogo."
    
    Dim targetType As String: targetType = pCatalog(targetTable)("Type")
    
    Trace "ExecuteDelete: Estratégia 'DELETE' detectada para o tipo '" & targetType & "'."
    
    If targetType = "SP" Then
        ' --- INÍCIO DA CORREÇÃO ---
        Call pHelpers.UpdateStatus("Iniciando exclusão em fonte 'SP'...")
        Dim spDetails As Object: Set spDetails = pCatalog(targetTable)("Connection")
        Dim sp As New fwSPConn
        sp.TraceOn = Me.TraceOn
        sp.Init spDetails("SiteURL"), spDetails("ListID")
        ' Chama a nova função iterativa que gerencia seu próprio progresso
        rowsAffected = sp.SPDeleteRowsByQuery(sql)
        ' --- FIM DA CORREÇÃO ---
    Else ' "XL"
        Call pHelpers.UpdateStatus("Iniciando exclusão em tabela Excel...")
        Dim xlTbl As New fwXLTable
        xlTbl.TraceOn = Me.TraceOn
        xlTbl.Init targetTable
        rowsAffected = xlTbl.DeleteRowsByQuery(sql)
    End If
    
    Trace "[DIAGNÓSTICO fwXLSPConn] Valor de rowsAffected recebido da função de exclusão: " & rowsAffected

    Call pHelpers.UpdateStatus("Operação DELETE concluída.")
    result.Add "rowsAffected", rowsAffected
    Set ExecuteDelete = result
End Function

'******************************************************************************
' Nome do elemento: ExecuteUpdate (Function)
' Resumo: Retorna o resultado calculado por ExecuteUpdate.
' @param sql - Parâmetro da rotina.
' @param execReport - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Function ExecuteUpdate(ByVal sql As String, ByRef execReport As Variant) As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim rowsAffected As Long: rowsAffected = 0 ' Inicializa com 0
    Dim shouldHighlightChanges As Boolean
    Dim highlightedCellCount As Long
    Dim highlightTargetTable As String
    Dim highlightColumns As Object
    Dim highlightSnapshot As Object
    Dim parsedSql As Object
    Dim tablesInQuery As Collection
    Dim spTablesInQuery As Object
    Dim tbl As Object
    Dim sp As fwSPConn ' Staging
    Dim spDetails As Object
    Dim spSync As fwSPConn ' Sync final
    Dim sourceXLTableName As String

    ' Sync / Select IDs vars
    Dim spTableTargetedForSet As String
    Dim targetColumnsForSet As New Collection
    Dim spKeyField As String
    Dim targetAlias As String ' Alias SP Alvo
    Dim sourceAlias As String ' Alias XL Fonte
    Dim idsToUpdateDict As Object ' Armazena IDs coletados ANTES do update local
    Set idsToUpdateDict = CreateObject("Scripting.Dictionary")

    On Error GoTo Cleanup

    ' Parse como JOIN primeiro
    Set parsedSql = pHelpers.ParseUpdateJoinSqlDetails(sql)

    If parsedSql Is Nothing Or InStr(1, sql, " JOIN ", vbTextCompare) = 0 Then
        GoTo SimpleUpdatePath ' Pula para lógica de UPDATE simples
    End If

    ' --- CAMINHO UPDATE COM JOIN ---
    Trace "ExecuteUpdate: Sintaxe JOIN detectada."
    Call pHelpers.UpdateStatus("Analisando UPDATE com JOIN...")

    Set tablesInQuery = parsedSql("Tables")
    Call NormalizeTextFormattedColumnsInJoinTables(tablesInQuery, sql)

    ' Verifica Self-Join
    Dim uniqueTableNames As Object: Set uniqueTableNames = CreateObject("Scripting.Dictionary")
    uniqueTableNames.CompareMode = vbTextCompare
    For Each tbl In tablesInQuery
        If Not uniqueTableNames.Exists(tbl("Name")) Then uniqueTableNames.Add tbl("Name"), pCatalog(tbl("Name"))("Type")
    Next
    If uniqueTableNames.Count = 1 And tablesInQuery.Count > 1 Then
        Dim selfJoinTableName As String: selfJoinTableName = uniqueTableNames.keys()(0)
        Dim selfJoinTableType As String: selfJoinTableType = uniqueTableNames(selfJoinTableName)
        Trace "ExecuteUpdate: Self-JOIN detectado na tabela '" & selfJoinTableName & "' (Tipo: " & selfJoinTableType & ")..."
        If selfJoinTableType = "SP" Then
            Trace "Estratégia: Execução iterativa SPUpdateRowsByQuery."
            Call pHelpers.UpdateStatus("Executando Self-JOIN no SharePoint...")
            Set sp = New fwSPConn: sp.TraceOn = Me.TraceOn
            Set spDetails = pCatalog(selfJoinTableName)("Connection")
            sp.Init spDetails("SiteURL"), spDetails("ListID")
            rowsAffected = sp.SPUpdateRowsByQuery(sql) ' Delega para a função iterativa
            GoTo FinalizeExecution
        Else ' Self-Join XL
            Trace "Estratégia: Execução direta no Excel."
            GoTo DirectExcelUpdatePath
        End If
    End If

    ' Identifica tabela SP, qual é alvo do SET, aliases e PK
    Set spTablesInQuery = CreateObject("Scripting.Dictionary"): spTablesInQuery.CompareMode = vbTextCompare
    Dim setAssignments As Collection: Set setAssignments = pHelpers.ParseSetClause(parsedSql("SetClause"))
    Dim assignmentDict As Object
    spTableTargetedForSet = "": targetAlias = "": sourceAlias = ""

    For Each assignmentDict In setAssignments
        Dim targetFieldExpr As String: targetFieldExpr = assignmentDict("TargetField")
        Dim currentTargetAlias As String
        Dim currentTargetColName As String
        If pHelpers.TryParseQualifiedColumnRef(targetFieldExpr, currentTargetAlias, currentTargetColName) Then
            Dim targetTableName As String
            For Each tbl In tablesInQuery
                If StrComp(tbl("Alias"), currentTargetAlias, vbTextCompare) = 0 Then targetTableName = tbl("Name"): Exit For
            Next tbl
            If Len(targetTableName) > 0 And pCatalog(targetTableName)("Type") = "SP" Then
                If Len(spTableTargetedForSet) = 0 Then
                    spTableTargetedForSet = targetTableName: targetAlias = currentTargetAlias
                    targetColumnsForSet.Add currentTargetColName
                    Trace "ExecuteUpdate: Tabela SP '" & spTableTargetedForSet & "' (Alias '" & targetAlias & "') ID alvo SET. Col: '" & currentTargetColName & "'"
                ElseIf spTableTargetedForSet = targetTableName Then
                    Dim colExists As Boolean: colExists = False: Dim existingCol As Variant
                    For Each existingCol In targetColumnsForSet
                        If existingCol = currentTargetColName Then
                            colExists = True
                            Exit For
                        End If
                    Next
                    If Not colExists Then targetColumnsForSet.Add currentTargetColName
                    Trace "ExecuteUpdate:   Add coluna alvo SP: '" & currentTargetColName & "'"
                Else
                    Err.Raise vbObjectError + 845, "ExecuteUpdate", "Nao suportado: Multiplas listas SP no SET."
                End If
            End If
        End If
    Next assignmentDict

    For Each tbl In tablesInQuery
        If pCatalog(tbl("Name"))("Type") = "SP" Then
            If Not spTablesInQuery.Exists(tbl("Name")) Then spTablesInQuery.Add tbl("Name"), True
            If tbl("Name") = spTableTargetedForSet Then
                 Set sp = New fwSPConn: sp.TraceOn = Me.TraceOn ' Temp
                 Set spDetails = pCatalog(tbl("Name"))("Connection")
                 sp.Init spDetails("SiteURL"), spDetails("ListID")
                 spKeyField = sp.spKeyField: Set sp = Nothing
                 If Len(spKeyField) = 0 Then Err.Raise vbObjectError + 846, "ExecuteUpdate", "PK não encontrada para '" & spTableTargetedForSet & "'."
                 Trace "ExecuteUpdate: PK para '" & spTableTargetedForSet & "' é '" & spKeyField & "'"
            End If
        Else
             sourceAlias = tbl("Alias")
             sourceXLTableName = tbl("Name")
        End If
    Next
    If Len(spTableTargetedForSet) > 0 And (Len(targetAlias) = 0 Or Len(sourceAlias) = 0) Then
        Err.Raise vbObjectError + 847, "ExecuteUpdate", "Aliases SP alvo ('" & targetAlias & "') ou XL fonte ('" & sourceAlias & "') não determinados."
    End If
     Trace "ExecuteUpdate: Alias SP Alvo = '" & targetAlias & "'. Alias XL Fonte = '" & sourceAlias & "'."

    If spTablesInQuery.Count = 0 Or Len(spTableTargetedForSet) = 0 Then
        GoTo DirectExcelUpdatePath ' Vai para XL-XL direto
    End If

    ' --- ESTRATÉGIA STAGE -> SELECT IDs -> UPDATE LOCAL -> SELECT UPDATED -> SYNC ---
    Trace "ExecuteUpdate: JOIN misto com alvo SP. Usando estratégia Stage-SelectIDs-UpdateLocal-SelectUpdated-Sync."

    ' Passo 1: STAGE
    Call pHelpers.UpdateStatus("Passo 1/4: Baixando dados do SharePoint...")
    Dim spListName As Variant
    For Each spListName In spTablesInQuery.keys
        Call pHelpers.UpdateStatus("Baixando lista '" & spListName & "'...")
        Set spDetails = pCatalog(spListName)("Connection")
        Set sp = New fwSPConn: sp.TraceOn = Me.TraceOn
        sp.Init spDetails("SiteURL"), spDetails("ListID")
        Dim stagingSql As String: stagingSql = BuildStagingSqlFor(spListName, sql, tablesInQuery)
        Dim spData As Object: Set spData = sp.SPRunQuery(stagingSql)
        Call CreateStagingTable(spListName, spData)
    Next

    ' Passo 2: Preparação Local
    Call pHelpers.UpdateStatus("Passo 2/4: Preparando tabelas locais...")
     Dim tblToFormat As Object
    For Each tblToFormat In tablesInQuery
        If pCatalog(tblToFormat("Name"))("Type") = "XL" Then
            On Error Resume Next
            Dim xlTable As ListObject, ws As Worksheet: Set xlTable = Nothing
            For Each ws In ActiveWorkbook.Worksheets
                Set xlTable = ws.ListObjects(tblToFormat("Name"))
                If Not xlTable Is Nothing Then Exit For
            Next ws
            If Not xlTable Is Nothing Then
                Dim fmtCol As ListColumn
                Set fmtCol = Nothing: On Error Resume Next: Set fmtCol = xlTable.ListColumns("Uniao"): On Error GoTo 0
                If Not fmtCol Is Nothing Then If Not fmtCol.DataBodyRange Is Nothing Then fmtCol.DataBodyRange.NumberFormat = "@"
                Set fmtCol = Nothing: On Error Resume Next: Set fmtCol = xlTable.ListColumns(spKeyField): On Error GoTo 0
                If Not fmtCol Is Nothing Then If Not fmtCol.DataBodyRange Is Nothing Then fmtCol.DataBodyRange.NumberFormat = "0"
            End If: On Error GoTo 0
        End If
    Next tblToFormat
    
    ' Passo 3: SELECT IDs (Pré-Update)
    Call pHelpers.UpdateStatus("Passo 3/5: Identificando linhas a serem atualizadas...")
    Trace "ExecuteUpdate: Executando SELECT local para capturar IDs ANTES do update..."
    Dim selectIdSql As String
    Dim selectIdFieldExpr As String
    Dim safeTargetAlias As String
    safeTargetAlias = pHelpers.Unbracket(targetAlias)
    If Len(safeTargetAlias) = 0 Then safeTargetAlias = pHelpers.Unbracket(spTableTargetedForSet)
    selectIdFieldExpr = "[" & safeTargetAlias & "].[" & pHelpers.Unbracket(spKeyField) & "]"
    selectIdSql = "SELECT DISTINCT " & selectIdFieldExpr & " FROM " & parsedSql("TablesClause")
    If Len(parsedSql("WhereClause")) > 0 Then selectIdSql = selectIdSql & " " & parsedSql("WhereClause")
    Trace "[DEBUG] ExecuteUpdate: SQL de Coleta de IDs (Lógico): " & selectIdSql

    Dim xlCollectIDs As New fwXLConn: xlCollectIDs.TraceOn = Me.TraceOn: xlCollectIDs.Init
    Dim rewrittenSelectIdSql As String
    rewrittenSelectIdSql = pHelpers.RewriteDataSourceNames(selectIdSql, ActiveWorkbook)
    Trace "[DEBUG] ExecuteUpdate: SQL de Coleta de IDs (Físico): " & rewrittenSelectIdSql

    ' --- INÍCIO DA CORREÇÃO ---
    ' Força IMEX=0 (modo escrita) para este SELECT, para que ele respeite a linha dummy e a formatação de texto
    Dim collectIdResult As Object
    Set collectIdResult = xlCollectIDs.XLRunQuery(rewrittenSelectIdSql, , forceWriteMode:=True)
    ' --- FIM DA CORREÇÃO ---

    Dim collectedIdCount As Long: collectedIdCount = 0
    If collectIdResult.Exists("Data") And Not pHelpers.IsVarEmpty(collectIdResult("Data")) Then
        Dim idsData As Variant: idsData = collectIdResult("Data")
        Set idsToUpdateDict = CreateObject("Scripting.Dictionary") ' Recria para garantir
        Dim idRow As Long, idCol As Long
        idCol = LBound(idsData, 1)
        For idRow = LBound(idsData, 2) To UBound(idsData, 2)
            idsToUpdateDict(idsData(idCol, idRow)) = True
        Next idRow
        collectedIdCount = idsToUpdateDict.Count
        Trace "ExecuteUpdate: Coleta Pré-Update encontrou " & collectedIdCount & " ID(s) distintos para atualizar."
    Else
        Trace "ExecuteUpdate: Coleta Pré-Update não encontrou IDs. Nenhuma linha será atualizada."
        rowsAffected = 0
        GoTo FinalizeExecution
    End If

    ' Passo 4: UPDATE LOCAL
    Call pHelpers.UpdateStatus("Passo 4/5: Executando atualização local...")
    Trace "ExecuteUpdate: Executando UPDATE local na tabela de staging..."
    Dim xlUpdate As New fwXLConn: xlUpdate.TraceOn = Me.TraceOn: xlUpdate.Init
    Dim xlUpdateResult As Object
    Set xlUpdateResult = xlUpdate.XLRunQuery(sql, execReport) ' Usa SQL original (com WHERE)
    If xlUpdateResult.Exists("rowsAffected") Then rowsAffected = xlUpdateResult("rowsAffected") Else rowsAffected = 0
    Trace "ExecuteUpdate: Update local concluído. Linhas afetadas localmente: " & rowsAffected

    ' Validação
    If rowsAffected <> collectedIdCount Then
         Trace "ExecuteUpdate: AVISO - Contagem de IDs (" & collectedIdCount & ") != Linhas afetadas (" & rowsAffected & ")."
         ' Continua, mas usa a contagem de IDs coletados como a fonte da verdade para o Sync
    End If

    If collectedIdCount = 0 Then
         Trace "ExecuteUpdate: Nenhuma linha afetada localmente ou nenhum ID coletado. Pulando Sync."
         GoTo FinalizeExecution
    End If

    ' Passo 5: SELECT UPDATED DATA & SYNC
    Call pHelpers.UpdateStatus("Passo 5/5: Coletando dados atualizados e sincronizando...")
    Trace "ExecuteUpdate: Coletando dados ATUALIZADOS da tabela de staging para os IDs selecionados..."

    Dim syncSelectList As String
    syncSelectList = "[" & spKeyField & "]"
    Dim colName As Variant
    For Each colName In targetColumnsForSet
        syncSelectList = syncSelectList & ", [" & colName & "]"
    Next colName

    Dim stagingTableNamePhysical As String: stagingTableNamePhysical = pHelpers.ResolveTableSqlName(spTableTargetedForSet, ActiveWorkbook)
    Dim xlCollectFinal As New fwXLConn: xlCollectFinal.TraceOn = Me.TraceOn: xlCollectFinal.Init

    Set spSync = New fwSPConn: spSync.TraceOn = Me.TraceOn
    Set spDetails = pCatalog(spTableTargetedForSet)("Connection")
    spSync.Init spDetails("SiteURL"), spDetails("ListID")
    spSync.spKeyField = spKeyField

    Dim idKeys As Variant
    idKeys = idsToUpdateDict.keys
    If Not IsArray(idKeys) Then
         Trace "ExecuteUpdate: ERRO - Chaves de IDs não estão em formato de array."
         Err.Raise vbObjectError + 849, "ExecuteUpdate", "Falha ao obter a lista de IDs para coleta final."
    End If

    Dim idLBound As Long, idUBound As Long, totalKeys As Long
    idLBound = LBound(idKeys)
    idUBound = UBound(idKeys)
    totalKeys = idUBound - idLBound + 1
    If totalKeys <= 0 Then
         Trace "ExecuteUpdate: ERRO - Lista de IDs para cláusula IN está vazia."
         Err.Raise vbObjectError + 849, "ExecuteUpdate", "Falha ao construir lista de IDs para coleta final."
    End If

    Dim totalChunks As Long
    totalChunks = (totalKeys + SYNC_ID_CHUNK_SIZE - 1) \ SYNC_ID_CHUNK_SIZE

    Dim chunkStart As Long, chunkEnd As Long, chunkNumber As Long
    Dim totalSyncedRows As Long

    For chunkStart = idLBound To idUBound Step SYNC_ID_CHUNK_SIZE
        chunkEnd = chunkStart + SYNC_ID_CHUNK_SIZE - 1
        If chunkEnd > idUBound Then chunkEnd = idUBound
        chunkNumber = chunkNumber + 1

        Dim idListForInClause As String
        idListForInClause = BuildSqlInListFromKeys(idKeys, chunkStart, chunkEnd)
        If Len(idListForInClause) = 0 Then
             Trace "ExecuteUpdate: ERRO - Falha ao montar IN para chunk " & chunkNumber & "."
             Err.Raise vbObjectError + 849, "ExecuteUpdate", "Falha ao construir lista de IDs para coleta final."
        End If

        Dim syncCollectSql As String
        syncCollectSql = "SELECT " & syncSelectList & " FROM " & stagingTableNamePhysical & _
                         " WHERE [" & spKeyField & "] IN (" & idListForInClause & ")"
        Trace "[DEBUG] ExecuteUpdate: SQL de Coleta Pós-Update (Chunk " & chunkNumber & "/" & totalChunks & "): " & syncCollectSql

        Dim collectFinalResult As Object
        Set collectFinalResult = xlCollectFinal.XLRunQuery(syncCollectSql, , forceWriteMode:=True)

        If collectFinalResult.Exists("Data") And Not pHelpers.IsVarEmpty(collectFinalResult("Data")) Then
            Dim dataToSync As Variant: dataToSync = collectFinalResult("Data")
            Dim headersToSync As Variant: headersToSync = collectFinalResult("Headers")
            Dim finalRowCount As Long
            On Error Resume Next: finalRowCount = UBound(dataToSync, 2) - LBound(dataToSync, 2) + 1
            If Err.Number <> 0 Then finalRowCount = 0: Err.Clear
            On Error GoTo Cleanup

            Trace "ExecuteUpdate: Coleta Final do chunk " & chunkNumber & " retornou " & finalRowCount & " linha(s)."
            If finalRowCount > 0 Then
                Call pHelpers.UpdateStatus("Sincronizando lote " & chunkNumber & "/" & totalChunks & " com " & finalRowCount & " linha(s)...")
                Dim syncCount As Long
                syncCount = spSync.SPUpdateRows(dataToSync, headersToSync)
                totalSyncedRows = totalSyncedRows + syncCount
                Trace "ExecuteUpdate: SPUpdateRows do chunk " & chunkNumber & " concluído. Afetadas: " & syncCount
            End If
        Else
            Trace "ExecuteUpdate: Coleta Final do chunk " & chunkNumber & " não retornou dados."
        End If
    Next chunkStart

    Trace "ExecuteUpdate: Sync em chunks concluído. Total de linhas sincronizadas: " & totalSyncedRows
    GoTo FinalizeExecution


DirectExcelUpdatePath:
    Trace "ExecuteUpdate: UPDATE JOIN detectado (apenas Excel)."
    Call pHelpers.UpdateStatus("Executando UPDATE com JOIN diretamente no Excel...")

    shouldHighlightChanges = ShouldHighlightChangesForSql(parsedSql, tablesInQuery)
    If shouldHighlightChanges Then
        highlightTargetTable = CStr(parsedSql("TargetTable"))
        Set highlightColumns = BuildUpdateJoinTargetColumns(parsedSql, tablesInQuery)
        If highlightColumns Is Nothing Then
            shouldHighlightChanges = False
        ElseIf highlightColumns.Count = 0 Then
            shouldHighlightChanges = False
        Else
            Set highlightSnapshot = CaptureTableSnapshot(highlightTargetTable, highlightColumns)
            If highlightSnapshot Is Nothing Then
                shouldHighlightChanges = False
            End If
        End If
    End If

    Dim xlDirect As New fwXLConn: xlDirect.TraceOn = Me.TraceOn: xlDirect.Init
    Dim xlDirectResult As Object: Set xlDirectResult = xlDirect.XLRunQuery(sql, execReport)
    If IsExecReportError(execReport) Then
        Trace "ExecuteUpdate: Falha no UPDATE JOIN via ACE. Tentando retentativa com shadow columns..."
        If Not TryExecuteJoinWithShadowColumns(sql, execReport, rowsAffected) Then
            Dim directErr As String
            directErr = GetExecReportErrorMessage(execReport)
            If Len(directErr) = 0 Then directErr = "Erro nao detalhado pelo driver ACE."
            Err.Raise vbObjectError + 860, "ExecuteUpdate", "Falha no UPDATE JOIN em Excel. " & directErr
        End If
    ElseIf xlDirectResult.Exists("rowsAffected") Then
        rowsAffected = xlDirectResult("rowsAffected")
    End If

    If shouldHighlightChanges Then
        highlightedCellCount = HighlightChangedCellsFromSnapshot( _
            highlightTargetTable, _
            highlightColumns, _
            highlightSnapshot, _
            pCellChangeHighlightColor, _
            pClearPreviousHighlightOnUpdate _
        )
        If highlightedCellCount > 0 Then
            Trace "ExecuteUpdate: Marcação visual aplicada em " & highlightedCellCount & " célula(s)."
        End If
        If Not result.Exists("changedCellsHighlighted") Then
            result.Add "changedCellsHighlighted", highlightedCellCount
        Else
            result("changedCellsHighlighted") = highlightedCellCount
        End If
    End If

    GoTo FinalizeExecution

SimpleUpdatePath:
    Call pHelpers.UpdateStatus("Executando UPDATE simples...")
    Set parsedSql = pHelpers.ParseUpdateSqlDetails(sql)
    If parsedSql Is Nothing Then Err.Raise vbObjectError + 840, "ExecuteUpdate", "Falha ao analisar UPDATE simples."
    Dim simpleTargetTable As String: simpleTargetTable = parsedSql("TargetTable")
    Dim simpleTargetType As String: simpleTargetType = pCatalog(simpleTargetTable)("Type")
    If simpleTargetType = "SP" Then
        Trace "Alvo SP. Usando estratégia SPUpdateRowsByQuery."
        Set sp = New fwSPConn: sp.TraceOn = Me.TraceOn
        Set spDetails = pCatalog(simpleTargetTable)("Connection")
        sp.Init spDetails("SiteURL"), spDetails("ListID")
        rowsAffected = sp.SPUpdateRowsByQuery(sql)
    Else ' Alvo XL
        If pHelpers.IsWhereClauseComplex(parsedSql("WhereClause")) Then
            Trace "Alvo XL com WHERE complexa. Usando fwXLTable.UpdateRowsByQuery."
            Dim xlTbl As New fwXLTable: xlTbl.TraceOn = Me.TraceOn
            xlTbl.Init simpleTargetTable
            rowsAffected = xlTbl.UpdateRowsByQuery(sql) ' Precisa de correção em fwXLTable
        Else
            Trace "Alvo XL com WHERE simples. Usando fwXLConn.XLRunQuery."
            Dim xlSimple As New fwXLConn: xlSimple.TraceOn = Me.TraceOn
            xlSimple.Init
            Dim xlResultSimple As Object: Set xlResultSimple = xlSimple.XLRunQuery(sql, execReport)
            If xlResultSimple.Exists("rowsAffected") Then rowsAffected = xlResultSimple("rowsAffected")
        End If
    End If
    GoTo FinalizeExecution

FinalizeExecution:
    Call pHelpers.UpdateStatus("Operação UPDATE concluída.")
    If Not result.Exists("rowsAffected") Then result.Add "rowsAffected", rowsAffected

Cleanup:
    On Error Resume Next
    Call pHelpers.UpdateStatus("Limpando tabelas temporárias...")
    Trace "RunQuery: [ETAPA FINAL] Limpando tabelas temporárias..."
    Call CleanupTempTables
    Application.EnableEvents = True: Application.ScreenUpdating = True
    Application.StatusBar = False
    Dim errNum As Long: errNum = Err.Number: Dim errDesc As String: errDesc = Err.Description
    Dim errSource As String: errSource = Err.source: On Error GoTo 0
    If errNum <> 0 Then
        Trace "ExecuteUpdate/RunQuery: ERRO CAPTURADO " & errNum & " - " & errDesc & " (Source: " & errSource & ")"
        If result Is Nothing Then Set result = CreateObject("Scripting.Dictionary")
        If Not result.Exists("rowsAffected") Then result.Add "rowsAffected", 0 Else result("rowsAffected") = 0
        If Not IsMissing(execReport) And IsObject(execReport) And TypeName(execReport) = "Dictionary" Then
            execReport("status") = "ERROR": execReport("errorNumber") = errNum
            execReport("errorMessage") = errDesc: execReport("errorSource") = errSource
        End If
    End If
    If result Is Nothing Then Set result = CreateObject("Scripting.Dictionary")
    If Not result.Exists("rowsAffected") Then result.Add "rowsAffected", rowsAffected
    Set ExecuteUpdate = result
    Trace "ExecuteUpdate/RunQuery: Fim da execução."
    If errNum <> 0 Then Err.Raise errNum, errSource, errDesc
End Function

'******************************************************************************
' Nome do elemento: IsExecReportError (Function)
' Resumo: Indica se o dicionario execReport foi preenchido com status de erro.
' @param execReport - Relatorio preenchido por fwXLConn.XLRunQuery.
' @return Boolean - True quando status = "ERROR".
'******************************************************************************
Private Function IsExecReportError(ByRef execReport As Variant) As Boolean
    On Error GoTo SafeExit
    If IsMissing(execReport) Then Exit Function
    If Not IsObject(execReport) Then Exit Function
    If TypeName(execReport) <> "Dictionary" Then Exit Function
    If Not execReport.Exists("status") Then Exit Function
    IsExecReportError = (StrComp(CStr(execReport("status")), "ERROR", vbTextCompare) = 0)
    Exit Function

SafeExit:
    IsExecReportError = False
End Function

'******************************************************************************
' Nome do elemento: GetExecReportErrorMessage (Function)
' Resumo: Consolida uma mensagem amigavel a partir do execReport.
' @param execReport - Relatorio preenchido por fwXLConn.XLRunQuery.
' @return String - Mensagem de erro consolidada.
'******************************************************************************
Private Function GetExecReportErrorMessage(ByRef execReport As Variant) As String
    On Error GoTo SafeExit
    If IsMissing(execReport) Then Exit Function
    If Not IsObject(execReport) Then Exit Function
    If TypeName(execReport) <> "Dictionary" Then Exit Function
    If execReport.Exists("errorMessage") Then GetExecReportErrorMessage = CStr(execReport("errorMessage"))
    If Len(GetExecReportErrorMessage) = 0 And execReport.Exists("errorSource") Then _
        GetExecReportErrorMessage = CStr(execReport("errorSource"))
    Exit Function

SafeExit:
    GetExecReportErrorMessage = vbNullString
End Function

'******************************************************************************
' Nome do elemento: ShouldHighlightChangesForSql (Function)
' Resumo:
'   Define se a marcação visual de células deve ser aplicada para o UPDATE JOIN
'   com base apenas na configuração da própria classe.
'******************************************************************************
Private Function ShouldHighlightChangesForSql(ByVal parsedSql As Object, ByVal tablesInQuery As Collection) As Boolean
    ShouldHighlightChangesForSql = pEnableCellChangeHighlight
End Function

'******************************************************************************
' Nome do elemento: BuildUpdateJoinTargetColumns (Function)
' Resumo:
'   Extrai as colunas-alvo da clausula SET para uso no snapshot de comparacao.
'******************************************************************************
Private Function BuildUpdateJoinTargetColumns(ByVal parsedSql As Object, ByVal tablesInQuery As Collection) As Object
    On Error GoTo SafeExit

    If parsedSql Is Nothing Then Exit Function
    If tablesInQuery Is Nothing Then Exit Function
    If Not parsedSql.Exists("SetClause") Then Exit Function
    If Not parsedSql.Exists("TargetTable") Then Exit Function

    Dim targetTable As String
    targetTable = CStr(parsedSql("TargetTable"))

    Dim targetAlias As String
    Dim tbl As Variant
    For Each tbl In tablesInQuery
        If StrComp(CStr(tbl("Name")), targetTable, vbTextCompare) = 0 Then
            targetAlias = CStr(tbl("Alias"))
            Exit For
        End If
    Next tbl
    If Len(targetAlias) = 0 Then targetAlias = targetTable

    Dim setCols As Object
    Set setCols = CreateObject("Scripting.Dictionary")
    setCols.CompareMode = vbTextCompare

    Dim assignments As Collection
    Set assignments = pHelpers.ParseSetClause(CStr(parsedSql("SetClause")))

    Dim assignmentDict As Object
    For Each assignmentDict In assignments
        Dim colAlias As String
        Dim colName As String
        If pHelpers.TryParseQualifiedColumnRef(CStr(assignmentDict("TargetField")), colAlias, colName) Then
            If StrComp(colAlias, targetAlias, vbTextCompare) = 0 Then
                If Not setCols.Exists(colName) Then
                    setCols.Add colName, True
                End If
            End If
        End If
    Next assignmentDict

    Set BuildUpdateJoinTargetColumns = setCols
    Exit Function

SafeExit:
    Set BuildUpdateJoinTargetColumns = Nothing
End Function

'******************************************************************************
' Nome do elemento: CaptureTableSnapshot (Function)
' Resumo:
'   Captura os valores atuais das colunas informadas para comparacao pos-update.
'******************************************************************************
Private Function CaptureTableSnapshot(ByVal tableName As String, ByVal columnNames As Object) As Object
    On Error GoTo SafeExit

    If columnNames Is Nothing Then Exit Function
    If columnNames.Count = 0 Then Exit Function

    Dim lo As ListObject
    Set lo = FindListObjectByName(tableName)
    If lo Is Nothing Then Exit Function

    Dim snapshot As Object
    Set snapshot = CreateObject("Scripting.Dictionary")
    snapshot.CompareMode = vbTextCompare

    Dim colName As Variant
    For Each colName In columnNames.keys
        Dim lc As ListColumn
        Set lc = Nothing
        On Error Resume Next
        Set lc = lo.ListColumns(CStr(colName))
        On Error GoTo SafeExit

        If Not lc Is Nothing Then
            If Not lc.DataBodyRange Is Nothing Then
                Dim valuesBefore As Collection
                Set valuesBefore = New Collection

                Dim c As Range
                For Each c In lc.DataBodyRange.Cells
                    valuesBefore.Add NormalizeValueForCompare(c.Value2)
                Next c

                snapshot.Add CStr(colName), valuesBefore
            End If
        End If
    Next colName

    Set CaptureTableSnapshot = snapshot
    Exit Function

SafeExit:
    Set CaptureTableSnapshot = Nothing
End Function

'******************************************************************************
' Nome do elemento: HighlightChangedCellsFromSnapshot (Function)
' Resumo:
'   Compara snapshot anterior com os valores atuais e colore apenas células
'   alteradas de conteúdo.
'******************************************************************************
Private Function HighlightChangedCellsFromSnapshot( _
    ByVal tableName As String, _
    ByVal columnNames As Object, _
    ByVal snapshot As Object, _
    ByVal highlightColor As Long, _
    Optional ByVal clearPreviousHighlight As Boolean = True _
) As Long
    On Error GoTo SafeExit

    If columnNames Is Nothing Then Exit Function
    If snapshot Is Nothing Then Exit Function

    Dim lo As ListObject
    Set lo = FindListObjectByName(tableName)
    If lo Is Nothing Then Exit Function

    Dim colName As Variant
    For Each colName In columnNames.keys
        If snapshot.Exists(CStr(colName)) Then
            Dim lc As ListColumn
            Set lc = Nothing
            On Error Resume Next
            Set lc = lo.ListColumns(CStr(colName))
            On Error GoTo SafeExit

            If Not lc Is Nothing Then
                If Not lc.DataBodyRange Is Nothing Then
                    Dim valuesBefore As Collection
                    Set valuesBefore = snapshot(CStr(colName))

                    Dim idx As Long
                    idx = 0

                    Dim c As Range
                    For Each c In lc.DataBodyRange.Cells
                        idx = idx + 1

                        If clearPreviousHighlight And c.Interior.Color = highlightColor Then
                            c.Interior.Pattern = xlNone
                        End If

                        If idx <= valuesBefore.Count Then
                            If StrComp(CStr(valuesBefore(idx)), NormalizeValueForCompare(c.Value2), vbBinaryCompare) <> 0 Then
                                c.Interior.Color = highlightColor
                                HighlightChangedCellsFromSnapshot = HighlightChangedCellsFromSnapshot + 1
                            End If
                        End If
                    Next c
                End If
            End If
        End If
    Next colName

SafeExit:
End Function

'******************************************************************************
' Nome do elemento: NormalizeValueForCompare (Function)
' Resumo:
'   Normaliza um valor de célula para comparação textual estável.
'******************************************************************************
Private Function NormalizeValueForCompare(ByVal v As Variant) As String
    On Error GoTo SafeExit

    If IsError(v) Then
        NormalizeValueForCompare = "#ERR:" & CStr(CLng(v))
        Exit Function
    End If

    If IsNull(v) Then
        NormalizeValueForCompare = vbNullString
        Exit Function
    End If

    If IsEmpty(v) Then
        NormalizeValueForCompare = vbNullString
        Exit Function
    End If

    Select Case VarType(v)
        Case vbDate
            NormalizeValueForCompare = Format$(CDate(v), "yyyy-mm-dd hh:nn:ss")
        Case vbBoolean
            If CBool(v) Then
                NormalizeValueForCompare = "TRUE"
            Else
                NormalizeValueForCompare = "FALSE"
            End If
        Case Else
            NormalizeValueForCompare = CStr(v)
    End Select
    Exit Function

SafeExit:
    NormalizeValueForCompare = vbNullString
End Function

'******************************************************************************
' Nome do elemento: NormalizeTextFormattedColumnsInJoinTables (Sub)
' Resumo:
'   Em UPDATE ... JOIN com tabelas Excel, força valores textuais para String
'   nas colunas com formato de texto ("@"), reduzindo falhas de tipagem mista
'   do provider ACE durante a execução do SQL.
'******************************************************************************
Private Sub NormalizeTextFormattedColumnsInJoinTables( _
    ByVal tablesInQuery As Collection, _
    Optional ByVal sql As String = vbNullString _
)
    On Error GoTo SafeExit

    If tablesInQuery Is Nothing Then Exit Sub

    Dim tbl As Object
    Dim tableName As String
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim aliasToTable As Object
    Dim refsByAlias As Object
    Dim colsByAlias As Object
    Dim aliasName As Variant
    Dim colName As Variant
    Dim useReferencedOnly As Boolean

    ' Preferencia: normalizar apenas colunas realmente referenciadas no SQL.
    If Len(Trim$(sql)) > 0 Then
        Set refsByAlias = pHelpers.ExtractAllReferencedColumns(sql)
        useReferencedOnly = Not refsByAlias Is Nothing
        If useReferencedOnly Then
            Set aliasToTable = CreateObject("Scripting.Dictionary")
            aliasToTable.CompareMode = vbTextCompare
            For Each tbl In tablesInQuery
                If Not aliasToTable.Exists(CStr(tbl("Alias"))) Then
                    aliasToTable.Add CStr(tbl("Alias")), CStr(tbl("Name"))
                End If
            Next tbl
        End If
    End If

    If useReferencedOnly Then
        For Each aliasName In refsByAlias.keys
            If aliasToTable.Exists(CStr(aliasName)) Then
                tableName = CStr(aliasToTable(CStr(aliasName)))
                If pCatalog.Exists(tableName) Then
                    If pCatalog(tableName)("Type") = "XL" Then
                        Set lo = FindListObjectByName(tableName)
                        If Not lo Is Nothing Then
                            Set colsByAlias = refsByAlias(aliasName)
                            For Each colName In colsByAlias.keys
                                Set lc = Nothing
                                On Error Resume Next
                                Set lc = lo.ListColumns(CStr(colName))
                                On Error GoTo SafeExit
                                If Not lc Is Nothing Then
                                    If Not lc.DataBodyRange Is Nothing Then
                                        If ColumnLooksTextFormatted(lc.DataBodyRange) Then
                                            CoerceRangeValuesToString lc.DataBodyRange
                                        End If
                                    End If
                                End If
                            Next colName
                        End If
                    End If
                End If
            End If
        Next aliasName
    Else
        ' Fallback: comportamento antigo.
        For Each tbl In tablesInQuery
            tableName = CStr(tbl("Name"))
            If pCatalog.Exists(tableName) Then
                If pCatalog(tableName)("Type") = "XL" Then
                    Set lo = FindListObjectByName(tableName)
                    If Not lo Is Nothing Then
                        For Each lc In lo.ListColumns
                            If Not lc.DataBodyRange Is Nothing Then
                                If ColumnLooksTextFormatted(lc.DataBodyRange) Then
                                    CoerceRangeValuesToString lc.DataBodyRange
                                End If
                            End If
                        Next lc
                    End If
                End If
            End If
        Next tbl
    End If

SafeExit:
End Sub

'******************************************************************************
' Nome do elemento: ColumnLooksTextFormatted (Function)
' Resumo:
'   Detecta se a coluna está formatada para texto (presença de "@").
'******************************************************************************
Private Function ColumnLooksTextFormatted(ByVal dataRng As Range) As Boolean
    On Error GoTo SafeExit

    If dataRng Is Nothing Then Exit Function
    If dataRng.Cells.Count = 0 Then Exit Function

    Dim fmt As String
    fmt = CStr(dataRng.Cells(1, 1).NumberFormat)
    ColumnLooksTextFormatted = (InStr(1, fmt, "@", vbTextCompare) > 0)
    Exit Function

SafeExit:
    ColumnLooksTextFormatted = False
End Function

'******************************************************************************
' Nome do elemento: CoerceRangeValuesToString (Sub)
' Resumo:
'   Converte valores não vazios de um Range para String sem alterar NumberFormat.
'******************************************************************************
Private Sub CoerceRangeValuesToString(ByVal dataRng As Range)
    On Error GoTo SafeExit

    If dataRng Is Nothing Then Exit Sub
    If dataRng.Cells.Count = 0 Then Exit Sub

    Dim constantsRng As Range
    Dim c As Range
    Dim rawValue As Variant

    ' Evita sobrescrever formulas: atua somente em constantes.
    On Error Resume Next
    Set constantsRng = dataRng.SpecialCells(xlCellTypeConstants)
    On Error GoTo SafeExit
    If constantsRng Is Nothing Then Exit Sub

    For Each c In constantsRng.Cells
        rawValue = c.Value2
        If Not IsError(rawValue) Then
            If Len(Trim$(CStr(rawValue))) > 0 Then
                c.Value2 = CStr(rawValue)
            End If
        End If
    Next c

SafeExit:
End Sub

'******************************************************************************
' Nome do elemento: TryExecuteJoinWithShadowColumns (Function)
' Resumo:
'   Retenta UPDATE ... JOIN em Excel quando o ACE falha com nomes de colunas
'   contendo ponto. Cria colunas tecnicas temporarias sem ponto, reescreve o
'   SQL e executa novamente; ao final remove as colunas temporarias.
' @param sql - SQL original do UPDATE JOIN.
' @param execReport - Relatorio de execucao compartilhado.
' @param rowsAffected - Retorno de linhas afetadas.
' @return Boolean - True quando a retentativa conclui com sucesso.
'******************************************************************************
Private Function TryExecuteJoinWithShadowColumns( _
    ByVal sql As String, _
    ByRef execReport As Variant, _
    ByRef rowsAffected As Long _
) As Boolean
    On Error GoTo ErrHandler

    Dim parsedSql As Object
    Dim tablesInQuery As Collection
    Dim aliasToTable As Object
    Dim setAssignments As Collection
    Dim allRefs As Object
    Dim setTargetKeys As Object
    Dim createdColsByTable As Object
    Dim shadowByTableCol As Object
    Dim rewrittenSql As String
    Dim hadRewrite As Boolean

    TryExecuteJoinWithShadowColumns = False
    rowsAffected = 0

    Set parsedSql = pHelpers.ParseUpdateJoinSqlDetails(sql)
    If parsedSql Is Nothing Then Exit Function
    Set tablesInQuery = parsedSql("Tables")
    If tablesInQuery Is Nothing Or tablesInQuery.Count = 0 Then Exit Function

    Set aliasToTable = CreateObject("Scripting.Dictionary")
    aliasToTable.CompareMode = vbTextCompare

    Dim tbl As Variant
    For Each tbl In tablesInQuery
        If Not aliasToTable.Exists(CStr(tbl("Alias"))) Then
            aliasToTable.Add CStr(tbl("Alias")), CStr(tbl("Name"))
        End If
    Next tbl

    Set setAssignments = pHelpers.ParseSetClause(parsedSql("SetClause"))
    Set setTargetKeys = CreateObject("Scripting.Dictionary")
    setTargetKeys.CompareMode = vbTextCompare

    Dim assignmentDict As Object
    For Each assignmentDict In setAssignments
        Dim targetAlias As String
        Dim targetCol As String
        If pHelpers.TryParseQualifiedColumnRef(CStr(assignmentDict("TargetField")), targetAlias, targetCol) Then
            setTargetKeys(targetAlias & "|" & targetCol) = True
            If InStr(1, targetCol, ".", vbBinaryCompare) > 0 Then
                Trace "TryExecuteJoinWithShadowColumns: alvo SET com ponto nao suportado: " & targetAlias & ".[" & targetCol & "]"
                Exit Function
            End If
        End If
    Next assignmentDict

    Set allRefs = pHelpers.ExtractAllReferencedColumns(sql)
    If allRefs Is Nothing Or allRefs.Count = 0 Then Exit Function

    Set createdColsByTable = CreateObject("Scripting.Dictionary")
    createdColsByTable.CompareMode = vbTextCompare
    Set shadowByTableCol = CreateObject("Scripting.Dictionary")
    shadowByTableCol.CompareMode = vbTextCompare

    rewrittenSql = sql

    Dim aliasKey As Variant
    For Each aliasKey In allRefs.keys
        Dim aliasName As String
        aliasName = CStr(aliasKey)
        If Not aliasToTable.Exists(aliasName) Then GoTo NextAlias

        Dim tableName As String
        tableName = CStr(aliasToTable(aliasName))

        Dim colsDict As Object
        Set colsDict = allRefs(aliasKey)

        Dim colKey As Variant
        For Each colKey In colsDict.keys
            Dim originalCol As String
            originalCol = CStr(colKey)

            If InStr(1, originalCol, ".", vbBinaryCompare) = 0 Then GoTo NextCol

            ' Nao reescrever lado alvo do SET para nao desviar a escrita.
            If setTargetKeys.Exists(aliasName & "|" & originalCol) Then GoTo NextCol

            Dim shadowKey As String
            shadowKey = tableName & "|" & originalCol

            Dim shadowCol As String
            If shadowByTableCol.Exists(shadowKey) Then
                shadowCol = CStr(shadowByTableCol(shadowKey))
            Else
                shadowCol = EnsureShadowColumn(tableName, originalCol, createdColsByTable)
                If Len(shadowCol) = 0 Then GoTo Cleanup
                shadowByTableCol.Add shadowKey, shadowCol
            End If

            rewrittenSql = Replace( _
                rewrittenSql, _
                aliasName & ".[" & Replace(originalCol, "]", "]]") & "]", _
                aliasName & ".[" & shadowCol & "]", _
                1, _
                -1, _
                vbTextCompare _
            )
            hadRewrite = True
NextCol:
        Next colKey
NextAlias:
    Next aliasKey

    If Not hadRewrite Then GoTo Cleanup

    Trace "TryExecuteJoinWithShadowColumns: SQL reescrito para contornar nomes com ponto."
    Dim xlRetry As New fwXLConn: xlRetry.TraceOn = Me.TraceOn: xlRetry.Init
    Dim retryResult As Object: Set retryResult = xlRetry.XLRunQuery(rewrittenSql, execReport)
    If retryResult.Exists("rowsAffected") Then rowsAffected = CLng(retryResult("rowsAffected"))

    TryExecuteJoinWithShadowColumns = Not IsExecReportError(execReport)

Cleanup:
    On Error Resume Next
    DropShadowColumns createdColsByTable
    On Error GoTo 0
    Exit Function

ErrHandler:
    Trace "TryExecuteJoinWithShadowColumns: ERRO " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function

'******************************************************************************
' Nome do elemento: EnsureShadowColumn (Function)
' Resumo:
'   Cria uma coluna tecnica temporaria em uma tabela e replica o conteudo da
'   coluna original para uso na retentativa de SQL.
' @param tableName - Nome da tabela do workbook.
' @param sourceColumnName - Nome da coluna original que contem ponto.
' @param createdColsByTable - Registro das colunas tecnicas criadas.
' @return String - Nome da coluna tecnica criada (ou vazio em falha).
'******************************************************************************
Private Function EnsureShadowColumn( _
    ByVal tableName As String, _
    ByVal sourceColumnName As String, _
    ByVal createdColsByTable As Object _
) As String
    On Error GoTo ErrHandler

    Dim lo As ListObject
    Set lo = FindListObjectByName(tableName)
    If lo Is Nothing Then Exit Function

    Dim srcCol As ListColumn
    On Error Resume Next
    Set srcCol = lo.ListColumns(sourceColumnName)
    On Error GoTo ErrHandler
    If srcCol Is Nothing Then Exit Function

    Dim shadowColName As String
    shadowColName = BuildShadowColumnName(lo)
    If Len(shadowColName) = 0 Then Exit Function

    Dim newCol As ListColumn
    Set newCol = lo.ListColumns.Add
    newCol.Name = shadowColName

    If Not lo.DataBodyRange Is Nothing Then
        If Not srcCol.DataBodyRange Is Nothing Then
            newCol.DataBodyRange.NumberFormat = srcCol.DataBodyRange.NumberFormat
            newCol.DataBodyRange.value = srcCol.DataBodyRange.value
        End If
    End If

    Dim colsCreated As Collection
    If createdColsByTable.Exists(tableName) Then
        Set colsCreated = createdColsByTable(tableName)
    Else
        Set colsCreated = New Collection
        createdColsByTable.Add tableName, colsCreated
    End If
    colsCreated.Add shadowColName

    EnsureShadowColumn = shadowColName
    Exit Function

ErrHandler:
    Trace "EnsureShadowColumn: ERRO ao criar shadow para '" & tableName & "." & sourceColumnName & "' -> " & Err.Number & " - " & Err.Description
    EnsureShadowColumn = vbNullString
End Function

'******************************************************************************
' Nome do elemento: BuildShadowColumnName (Function)
' Resumo: Gera nome unico para coluna tecnica temporaria.
' @param lo - Tabela onde a coluna sera criada.
' @return String - Nome tecnico disponivel.
'******************************************************************************
Private Function BuildShadowColumnName(ByVal lo As ListObject) As String
    On Error GoTo SafeExit

    Dim i As Long
    For i = 1 To 9999
        Dim candidate As String
        candidate = "__SPX_DOT_" & Format$(i, "0000")

        On Error Resume Next
        Dim lc As ListColumn
        Set lc = lo.ListColumns(candidate)
        If Err.Number <> 0 Or lc Is Nothing Then
            Err.Clear
            On Error GoTo 0
            BuildShadowColumnName = candidate
            Exit Function
        End If
        Err.Clear
        Set lc = Nothing
        On Error GoTo SafeExit
    Next i

SafeExit:
    If Len(BuildShadowColumnName) = 0 Then BuildShadowColumnName = vbNullString
End Function

'******************************************************************************
' Nome do elemento: FindListObjectByName (Function)
' Resumo: Localiza uma tabela (ListObject) pelo nome em qualquer planilha.
' @param tableName - Nome logico/fisico da tabela.
' @return ListObject - Objeto encontrado; Nothing quando inexistente.
'******************************************************************************
Private Function FindListObjectByName(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ActiveWorkbook.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set FindListObjectByName = lo
            Exit Function
        End If
    Next ws
End Function

'******************************************************************************
' Nome do elemento: DropShadowColumns (Sub)
' Resumo: Remove colunas tecnicas temporarias criadas para retentativa SQL.
' @param createdColsByTable - Dicionario tabela -> Collection de colunas.
'******************************************************************************
Private Sub DropShadowColumns(ByVal createdColsByTable As Object)
    On Error GoTo SafeExit
    If createdColsByTable Is Nothing Then Exit Sub

    Dim tableName As Variant
    For Each tableName In createdColsByTable.keys
        Dim lo As ListObject
        Set lo = FindListObjectByName(CStr(tableName))
        If lo Is Nothing Then GoTo NextTable

        Dim colsCreated As Collection
        Set colsCreated = createdColsByTable(tableName)

        Dim i As Long
        For i = colsCreated.Count To 1 Step -1
            On Error Resume Next
            lo.ListColumns(CStr(colsCreated(i))).Delete
            On Error GoTo 0
        Next i
NextTable:
    Next tableName

SafeExit:
    Exit Sub
End Sub

'******************************************************************************
' Nome do elemento: ValidateQuerySchema (Sub)
' Resumo: Executa a rotina ValidateQuerySchema.
' @param sql - Parâmetro da rotina.
' @param tablesInQuery - Parâmetro da rotina.
' @param parsedSql - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Sub ValidateQuerySchema(ByVal sql As String, ByVal tablesInQuery As Collection, ByVal parsedSql As Object)
    On Error GoTo ErrHandler
    Trace "ValidateQuerySchema: Iniciando validação de esquema..."
    Trace "[DIAGNÓSTICO-VS] ==> Entrando em ValidateQuerySchema"
    
    If parsedSql Is Nothing Then
        Trace "ValidateQuerySchema: Nenhum SQL parseado para validar."
        Trace "[DIAGNÓSTICO-VS] parsedSql é Nothing. Saindo."
        Exit Sub
    End If
    
    If tablesInQuery Is Nothing Then
         Trace "[DIAGNÓSTICO-VS] tablesInQuery É Nothing. Saindo."
         Trace "ValidateQuerySchema: Coleção de tabelas é Nothing."
         Exit Sub
    End If
     Trace "[DIAGNÓSTICO-VS] tablesInQuery.Count = " & tablesInQuery.Count
     If tablesInQuery.Count = 0 Then
        Trace "ValidateQuerySchema: Nenhuma tabela encontrada para validar."
        Trace "[DIAGNÓSTICO-VS] tablesInQuery.Count é 0. Saindo."
        Exit Sub
    End If

    Dim tableSchemas As Object: Set tableSchemas = CreateObject("Scripting.Dictionary")
    tableSchemas.CompareMode = vbTextCompare
    
    Dim actualColumns As Object, sourceType As String, sp As fwSPConn, spDetails As Object
    Dim spField As Variant, xlT As fwXLTable, xlCol As ListColumn
    
    Trace "[DIAGNÓSTICO-VS] --- ETAPA 1: Validando colunas COM ALIAS ---"
    Dim allAliasedColumns As Object: Set allAliasedColumns = pHelpers.ExtractAllReferencedColumns(sql)
    Trace "[DIAGNÓSTICO-VS] Número de aliases com colunas referenciadas: " & allAliasedColumns.Count
    
    Dim tblAlias As Variant
    For Each tblAlias In allAliasedColumns.keys
        Trace "[DIAGNOSTICO-VS] Processando Alias: '" & tblAlias & "'"
        Dim realTableName As String: realTableName = ""
        Dim tbl As Object
        For Each tbl In tablesInQuery
             Trace "[DIAGNÓSTICO-VS]   Comparando com tbl('Alias'): '" & tbl("Alias") & "'"
            If StrComp(tbl("Alias"), CStr(tblAlias), vbTextCompare) = 0 Then
                realTableName = tbl("Name")
                Trace "[DIAGNOSTICO-VS]   Encontrado Nome Real: '" & realTableName & "' para Alias '" & tblAlias & "'"
                Exit For
            End If
        Next tbl
        
        If Len(realTableName) = 0 Then
            Trace "[DIAGNOSTICO-VS]   ERRO: Nome Real NAO encontrado para Alias '" & tblAlias & "'"
            Err.Raise vbObjectError + 901, "ValidateQuerySchema", "O alias de tabela '" & tblAlias & "' usado em uma coluna não foi encontrado na cláusula FROM/JOIN."
        End If
        
        ' Obtém ou carrega o esquema da tabela
        If Not tableSchemas.Exists(realTableName) Then
            Set actualColumns = Nothing

            If pSchemaCache Is Nothing Then
                Set pSchemaCache = CreateObject("Scripting.Dictionary")
                pSchemaCache.CompareMode = vbTextCompare
            End If

            If pSchemaCache.Exists(realTableName) Then
                If IsObject(pSchemaCache(realTableName)) Then
                    Set actualColumns = pSchemaCache(realTableName)
                    tableSchemas.Add realTableName, actualColumns
                    Trace "[DIAGNÓSTICO-VS]   Esquema para '" & realTableName & "' obtido do cache."
                Else
                    pSchemaCache.Remove realTableName
                End If
            End If

            If actualColumns Is Nothing Then
                Trace "[DIAGNÓSTICO-VS]   Carregando esquema para Tabela '" & realTableName & "'..."
                Set actualColumns = CreateObject("Scripting.Dictionary")
                actualColumns.CompareMode = vbTextCompare
                
                ' *** INÍCIO DA VERIFICAÇÃO ADICIONAL ***
                If Not pCatalog.Exists(realTableName) Then
                    Trace "[DIAGNÓSTICO-VS]     ERRO FATAL: Tabela '" & realTableName & "' (usada no SQL) não encontrada no pCatalog!"
                    Err.Raise vbObjectError + 904, "ValidateQuerySchema", "A tabela '" & realTableName & "' referenciada no SQL não foi definida na inicialização (InitSP)."
                End If
                 If Not IsObject(pCatalog(realTableName)) Then
                     Trace "[DIAGNÓSTICO-VS]     ERRO FATAL: Entrada do catálogo para '" & realTableName & "' não é um objeto!"
                     Err.Raise vbObjectError + 905, "ValidateQuerySchema", "A entrada do catálogo para '" & realTableName & "' está corrompida."
                 End If
                 If Not pCatalog(realTableName).Exists("Type") Then
                     Trace "[DIAGNÓSTICO-VS]     ERRO FATAL: Entrada do catálogo para '" & realTableName & "' não possui a chave 'Type'!"
                     Err.Raise vbObjectError + 906, "ValidateQuerySchema", "A entrada do catálogo para '" & realTableName & "' está incompleta (falta 'Type')."
                 End If
                Trace "[DIAGNÓSTICO-VS]     Acesso ao pCatalog('" & realTableName & "') parece OK. Obtendo tipo..."
                ' *** FIM DA VERIFICAÇÃO ADICIONAL ***

                sourceType = pCatalog(realTableName)("Type") ' Esta linha agora deve ser segura
                 Trace "[DIAGNÓSTICO-VS]     Tipo da Fonte: '" & sourceType & "'"
                
                If sourceType = "SP" Then
                    Set sp = New fwSPConn
                    sp.TraceOn = Me.TraceOn
                    Set spDetails = pCatalog(realTableName)("Connection")
                    sp.Init spDetails("SiteURL"), spDetails("ListID")
                    For Each spField In sp.SPGetFields()
                        actualColumns.Add CStr(spField), True
                         Trace "[DIAGNÓSTICO-VS]       Coluna SP adicionada ao esquema: '" & spField & "'"
                    Next
                    Set sp = Nothing
                Else ' "XL"
                    Set xlT = New fwXLTable
                    xlT.TraceOn = Me.TraceOn
                    On Error Resume Next
                    xlT.Init realTableName
                    If Err.Number <> 0 Then
                        Trace "[DIAGNÓSTICO-VS]     ERRO ao inicializar fwXLTable para '" & realTableName & "': " & Err.Description
                        Err.Clear
                        On Error GoTo ErrHandler
                        Err.Raise vbObjectError + 903, "ValidateQuerySchema", "Erro ao inicializar a tabela Excel '" & realTableName & "' para validação. Verifique se ela existe."
                    End If
                    On Error GoTo ErrHandler
                    For Each xlCol In xlT.Self.ListColumns
                        actualColumns.Add xlCol.Name, True
                        Trace "[DIAGNÓSTICO-VS]       Coluna XL adicionada ao esquema: '" & xlCol.Name & "'"
                    Next
                    Set xlT = Nothing
                End If

                tableSchemas.Add realTableName, actualColumns
                Set pSchemaCache(realTableName) = actualColumns
                Trace "[DIAGNÓSTICO-VS]   Esquema para '" & realTableName & "' carregado e armazenado em cache."
            End If
        Else
             Trace "[DIAGNÓSTICO-VS]   Esquema para Tabela '" & realTableName & "' já estava carregado."
            Set actualColumns = tableSchemas(realTableName)
        End If
        
        ' Verifica as coluna extraídas para este alias
        Dim columnsToCheck As Object: Set columnsToCheck = allAliasedColumns(tblAlias)
        Dim colName As Variant
        Trace "[DIAGNOSTICO-VS]   Verificando " & columnsToCheck.Count & " coluna(s) para Alias '" & tblAlias & "'..."
        For Each colName In columnsToCheck.keys
             Trace "[DIAGNÓSTICO-VS]     Verificando Coluna: '" & colName & "' (Tipo: " & TypeName(colName) & ")"
             Dim colNameStr As String: colNameStr = CStr(colName)
             Dim columnFound As Boolean
             columnFound = actualColumns.Exists(colNameStr)
             Trace "[DIAGNÓSTICO-VS]       actualColumns.Exists(CStr('" & colName & "')) = " & columnFound
            If Not columnFound Then
                 Trace "[DIAGNÓSTICO-VS]       ERRO: Coluna '" & colNameStr & "' NÃO ENCONTRADA no esquema de '" & realTableName & "'"
                Err.Raise vbObjectError + 902, "ValidateQuerySchema", "A coluna '" & colNameStr & "' não foi encontrada na tabela '" & realTableName & "' (alias: '" & tblAlias & "')."
            End If
        Next colName
        Trace "[DIAGNOSTICO-VS]   Fim da verificacao para Alias '" & tblAlias & "'"
    Next tblAlias
    Trace "[DIAGNÓSTICO-VS] --- FIM ETAPA 1 ---"
    
    Trace "[DIAGNÓSTICO-VS] --- Verificando condição para ETAPA 2 ---"
    Trace "[DIAGNÓSTICO-VS] tablesInQuery.Count = " & tablesInQuery.Count
    Dim checkColumnsKey As Boolean: checkColumnsKey = False
    If Not parsedSql Is Nothing Then checkColumnsKey = parsedSql.Exists("Columns")
    Trace "[DIAGNÓSTICO-VS] parsedSql.Exists(""Columns"") = " & checkColumnsKey
    
    If tablesInQuery.Count = 1 And checkColumnsKey Then
         Trace "[DIAGNÓSTICO-VS] --- ENTRANDO NA ETAPA 2: Validando colunas SEM ALIAS ---"
         ' ... (código da Etapa 2 permanece o mesmo, incluindo a variável renomeada columnFoundStep2) ...
        Dim singleTableName As String: singleTableName = tablesInQuery(1)("Name")
         Trace "[DIAGNÓSTICO-VS]   Tabela Única: '" & singleTableName & "'"
        If Not tableSchemas.Exists(singleTableName) Then
            Trace "[DIAGNÓSTICO-VS]   Carregando esquema para Tabela Única '" & singleTableName & "'..."
            Set actualColumns = CreateObject("Scripting.Dictionary")
            actualColumns.CompareMode = vbTextCompare
            sourceType = pCatalog(singleTableName)("Type") ' Assumindo que singleTableName existe no pCatalog
             Trace "[DIAGNÓSTICO-VS]     Tipo da Fonte: '" & sourceType & "'"
            If sourceType = "SP" Then
                Set sp = New fwSPConn: sp.TraceOn = Me.TraceOn
                Set spDetails = pCatalog(singleTableName)("Connection")
                sp.Init spDetails("SiteURL"), spDetails("ListID")
                For Each spField In sp.SPGetFields(): actualColumns.Add CStr(spField), True: Next
                Set sp = Nothing
            Else ' "XL"
                Set xlT = New fwXLTable: xlT.TraceOn = Me.TraceOn
                On Error Resume Next: xlT.Init singleTableName
                If Err.Number <> 0 Then
                    Trace "[DIAGNÓSTICO-VS]     ERRO Init fwXLTable Etapa 2: " & Err.Description: Err.Clear: On Error GoTo ErrHandler
                    Err.Raise vbObjectError + 903, "ValidateQuerySchema", "Erro Init XL Tabela Etapa 2: '" & singleTableName & "'."
                End If: On Error GoTo ErrHandler
                For Each xlCol In xlT.Self.ListColumns: actualColumns.Add xlCol.Name, True: Next
                Set xlT = Nothing
            End If
            tableSchemas.Add singleTableName, actualColumns
            Trace "[DIAGNÓSTICO-VS]   Esquema para '" & singleTableName & "' carregado."
        Else
            Trace "[DIAGNÓSTICO-VS]   Esquema para Tabela Única '" & singleTableName & "' já carregado."
            Set actualColumns = tableSchemas(singleTableName)
        End If
        Dim columnsInSelect As Variant
        If IsObject(parsedSql("Columns")) Then
            Set columnsInSelect = parsedSql("Columns")
        Else
            columnsInSelect = parsedSql("Columns")
        End If
        Dim colDict As Variant
        Trace "[DIAGNÓSTICO-VS]   Verificando colunas SEM alias no SELECT/INSERT..."
        If pHelpers.IsVarEmpty(columnsInSelect) Then
            Trace "[DIAGNÓSTICO-VS]   Nenhuma coluna para validar na Etapa 2."
        ElseIf IsArray(columnsInSelect) Then
            Dim i As Long
            For i = LBound(columnsInSelect) To UBound(columnsInSelect)
                Dim colNameToValidate As String: colNameToValidate = CStr(columnsInSelect(i))
                If Len(colNameToValidate) > 0 Then
                    Trace "[DIAGNÓSTICO-VS]     Verificando Coluna sem alias: '" & colNameToValidate & "'"
                    Dim columnFoundStep2 As Boolean
                    columnFoundStep2 = actualColumns.Exists(CStr(colNameToValidate))
                    Trace "[DIAGNÓSTICO-VS]       actualColumns.Exists(CStr('" & colNameToValidate & "')) = " & columnFoundStep2
                    If Not columnFoundStep2 Then
                        Trace "[DIAGNÓSTICO-VS]       ERRO: Coluna '" & colNameToValidate & "' NÃO ENCONTRADA no esquema de '" & singleTableName & "'"
                        Err.Raise vbObjectError + 902, "ValidateQuerySchema", "A coluna '" & colNameToValidate & "' não foi encontrada na tabela '" & singleTableName & "'."
                    End If
                End If
            Next i
        ElseIf TypeName(columnsInSelect) = "Collection" Then
            For Each colDict In columnsInSelect
                If IsObject(colDict) Then
                    If colDict("TableAlias") = "" And colDict("ColumnName") <> "*" Then
                        Dim colNameToValidate2 As String: colNameToValidate2 = colDict("ColumnName")
                        Trace "[DIAGNÓSTICO-VS]     Verificando Coluna sem alias: '" & colNameToValidate2 & "'"
                        Dim columnFoundStep2b As Boolean
                        columnFoundStep2b = actualColumns.Exists(CStr(colNameToValidate2))
                        Trace "[DIAGNÓSTICO-VS]       actualColumns.Exists(CStr('" & colNameToValidate2 & "')) = " & columnFoundStep2b
                        If Not columnFoundStep2b Then
                            Trace "[DIAGNÓSTICO-VS]       ERRO: Coluna '" & colNameToValidate2 & "' NÃO ENCONTRADA no esquema de '" & singleTableName & "'"
                            Err.Raise vbObjectError + 902, "ValidateQuerySchema", "A coluna '" & colNameToValidate2 & "' não foi encontrada na tabela '" & singleTableName & "'."
                        End If
                    End If
                Else
                    Dim colNameToValidate3 As String: colNameToValidate3 = CStr(colDict)
                    If Len(colNameToValidate3) > 0 Then
                        Trace "[DIAGNÓSTICO-VS]     Verificando Coluna sem alias: '" & colNameToValidate3 & "'"
                        Dim columnFoundStep2c As Boolean
                        columnFoundStep2c = actualColumns.Exists(CStr(colNameToValidate3))
                        Trace "[DIAGNÓSTICO-VS]       actualColumns.Exists(CStr('" & colNameToValidate3 & "')) = " & columnFoundStep2c
                        If Not columnFoundStep2c Then
                            Trace "[DIAGNÓSTICO-VS]       ERRO: Coluna '" & colNameToValidate3 & "' NÃO ENCONTRADA no esquema de '" & singleTableName & "'"
                            Err.Raise vbObjectError + 902, "ValidateQuerySchema", "A coluna '" & colNameToValidate3 & "' não foi encontrada na tabela '" & singleTableName & "'."
                        End If
                    End If
                End If
            Next
        Else
            Trace "[DIAGNÓSTICO-VS]   Formato inesperado para Columns: '" & TypeName(columnsInSelect) & "'."
        End If
        Trace "[DIAGNÓSTICO-VS]   Fim da verificação de colunas sem alias."
    Else
         Trace "[DIAGNÓSTICO-VS] --- Condição para ETAPA 2 FALHOU. Pulando Etapa 2. ---"
    End If
    Trace "[DIAGNÓSTICO-VS] --- FIM ETAPA 2 ---"
    
    Trace "ValidateQuerySchema: Validação de esquema concluída com sucesso."
     Trace "[DIAGNÓSTICO-VS] <== Saindo de ValidateQuerySchema (Sucesso)"
    Exit Sub
    
ErrHandler:
    Trace "[DIAGNÓSTICO-VS] !!! ERRO em ValidateQuerySchema: " & Err.Number & " - " & Err.Description & " (Source: " & Err.source & ")"
    Set sp = Nothing
    Set xlT = Nothing
    Err.Raise Err.Number, "fwXLSPConn.ValidateQuerySchema", Err.Description
End Sub

'******************************************************************************
' Nome do elemento: SyncUpdateBatch (Function)
' Resumo: Retorna o resultado calculado por SyncUpdateBatch.
' @param logicalName - Parâmetro da rotina.
' @param dataArray - Parâmetro da rotina.
' @param headers - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function SyncUpdateBatch(ByVal logicalName As String, ByVal dataArray As Variant, ByVal headers As Variant) As Long
    If Not pCatalog.Exists(logicalName) Or pCatalog(logicalName)("Type") <> "SP" Then
        Err.Raise vbObjectError + 850, "SyncUpdateBatch", "Fonte de dados '" & logicalName & "' não é uma lista SP válida."
    End If
    
    Dim spDetails As Object: Set spDetails = pCatalog(logicalName)("Connection")
    Dim sp As New fwSPConn
    sp.TraceOn = Me.TraceOn
    sp.Init spDetails("SiteURL"), spDetails("ListID")
    
    SyncUpdateBatch = sp.SPUpdateRows(dataArray, headers)
End Function

'******************************************************************************
' Nome do elemento: SyncInsertBatch (Function)
' Resumo: Retorna o resultado calculado por SyncInsertBatch.
' @param logicalName - Parâmetro da rotina.
' @param dataArray - Parâmetro da rotina.
' @param headers - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function SyncInsertBatch(ByVal logicalName As String, ByVal dataArray As Variant, ByVal headers As Variant) As Long
    If Not pCatalog.Exists(logicalName) Or pCatalog(logicalName)("Type") <> "SP" Then
        Err.Raise vbObjectError + 851, "SyncInsertBatch", "Fonte de dados '" & logicalName & "' não é uma lista SP válida."
    End If
    
    Dim spDetails As Object: Set spDetails = pCatalog(logicalName)("Connection")
    Dim sp As New fwSPConn
    sp.TraceOn = Me.TraceOn
    sp.Init spDetails("SiteURL"), spDetails("ListID")
    
    Dim result As Variant
    result = sp.SPInsertRows(dataArray, headers)
    ' SPInsertRows retorna um array; o segundo elemento é o total de linhas afetadas
    If IsArray(result) Then SyncInsertBatch = result(1)
End Function


