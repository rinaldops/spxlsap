VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fwSPConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' Nome do elemento: fwSPConn (Classe)
' Resumo:
'   Camada de acesso a Listas SharePoint via OLEDB/ADODB (ACE WSS).
'
' Responsabilidades:
'   - Abrir/fechar conexão OLEDB com a lista SharePoint.
'   - Executar comandos SQL (SELECT, INSERT, UPDATE) na lista.
'   - Abstrair a complexidade da conexão e formatação de dados.
'
' Dependências:
'   - fwHelpers
'   - Microsoft ActiveX Data Objects (late binding)
'   - Provedor ACE OLEDB WSS (Microsoft.ACE.OLEDB.12.0)
'******************************************************************************

Option Explicit

'=== Estado interno ===
Private pSiteURL    As String
Private pListID     As String
Private pConn       As Object ' ADODB.Connection
Private pKeyField   As String ' amigável (default "ID")
Private pFields     As Collection ' nomes amigáveis, carregados sob demanda
Private pTraceOn    As Boolean
Private pHelpers    As fwHelpers

'=== Constantes ADO (late binding) ===
Private Const adOpenForwardOnly  As Long = 0
Private Const adOpenKeyset       As Long = 1
Private Const adOpenStatic       As Long = 3
Private Const adLockReadOnly     As Long = 1
Private Const adLockOptimistic   As Long = 3
Private Const adUseClient        As Long = 3
Private Const adStateOpen        As Long = 1

'******************************************************************************
' Nome do elemento: Init (Sub)
' Resumo: Executa a rotina Init.
' @param siteURL - Parâmetro da rotina.
' @param listID - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Sub Init(ByVal siteURL As String, ByVal listID As String)
    If Len(Trim$(siteURL)) = 0 Or Len(Trim$(listID)) = 0 Then
        Err.Raise vbObjectError + 100, "fwSPConn.Init", "Parâmetros SiteURL e ListID são obrigatórios."
    End If
    pSiteURL = siteURL
    pListID = listID
    pKeyField = "ID"
    Set pFields = Nothing
    EnsureHelpers
    Trace "Init: fwSPConn inicializada para Site='" & pSiteURL & "', ListID='" & pListID & "'"
End Sub

'******************************************************************************
' Nome do elemento: helpers (Property Set)
' Resumo: Define o valor da propriedade helpers.
' @param h - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Set helpers(ByVal h As fwHelpers)
    If h Is Nothing Then
        Set pHelpers = New fwHelpers
    Else
        Set pHelpers = h
    End If
End Property

'******************************************************************************
' Nome do elemento: helpers (Property Get)
' Resumo: Obtem o valor da propriedade helpers.
' @return fwHelpers - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get helpers() As fwHelpers
    If pHelpers Is Nothing Then Set pHelpers = New fwHelpers
    Set helpers = pHelpers
End Property

'******************************************************************************
' Nome do elemento: EnsureHelpers (Sub)
' Resumo: Executa a rotina EnsureHelpers.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub EnsureHelpers()
    If pHelpers Is Nothing Then Set pHelpers = New fwHelpers
End Sub

'******************************************************************************
' Nome do elemento: ValidateConnectionPreflight (Sub)
' Resumo:
'   Valida parametros obrigatorios antes de abrir conexao com SharePoint.
' @param operationName - Nome da operacao para diagnostico e log.
'******************************************************************************
Private Sub ValidateConnectionPreflight(ByVal operationName As String)
    EnsureHelpers

    If Len(Trim$(pSiteURL)) = 0 Then
        Call Me.helpers.AppendPersistentLog("ERROR", "fwSPConn", operationName, _
                                            "SiteURL nao informado. Chame Init(siteURL, listID).")
        Err.Raise vbObjectError + 103, "fwSPConn." & operationName, _
                  "SiteURL nao informado. Chame Init(siteURL, listID)."
    End If

    If Len(Trim$(pListID)) = 0 Then
        Call Me.helpers.AppendPersistentLog("ERROR", "fwSPConn", operationName, _
                                            "ListID nao informado. Chame Init(siteURL, listID).")
        Err.Raise vbObjectError + 104, "fwSPConn." & operationName, _
                  "ListID nao informado. Chame Init(siteURL, listID)."
    End If
End Sub

'******************************************************************************
' Nome do elemento: SPLogEvento (Sub)
' Resumo:
'   Registra evento persistente para as operações de fwSPConn.
'******************************************************************************
Private Sub SPLogEvento(ByVal levelName As String, ByVal actionName As String, ByVal message As String, Optional ByVal sqlText As String = "")
    On Error Resume Next
    EnsureHelpers
    Call Me.helpers.AppendPersistentLog(levelName, "fwSPConn", actionName, message, sqlText)
End Sub

'******************************************************************************
' Nome do elemento: TraceOn (Property Let)
' Resumo: Define o valor da propriedade TraceOn.
' @param v - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Let TraceOn(ByVal v As Boolean)
    pTraceOn = CBool(v)
    Me.helpers.TraceOn = v
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
' Nome do elemento: spKeyField (Property Get)
' Resumo: Obtem o valor da propriedade spKeyField.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get spKeyField() As String
    If Len(pKeyField) = 0 Then pKeyField = "ID"
    spKeyField = pKeyField
End Property

'******************************************************************************
' Nome do elemento: spKeyField (Property Let)
' Resumo: Define o valor da propriedade spKeyField.
' @param fldName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Property Let spKeyField(ByVal fldName As String)
    Dim nm As String: nm = Trim$(fldName)
    If Len(nm) = 0 Then Err.Raise vbObjectError + 101, "fwSPConn.SPKeyField", "Nome de campo chave não pode ser vazio."
    Dim ok As Boolean, f As Variant
    For Each f In SPGetFields()
        If StrComp(CStr(f), nm, vbTextCompare) = 0 Then ok = True: Exit For
    Next
    If Not ok Then Err.Raise vbObjectError + 102, "fwSPConn.SPKeyField", "Campo '" & nm & "' não existe na lista."
    pKeyField = nm
End Property

'******************************************************************************
' Nome do elemento: SPConnect (Sub)
' Resumo: Executa a rotina SPConnect.
' @param isReadOnly - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Sub SPConnect(ByVal isReadOnly As Boolean)
    On Error GoTo ErrHandler
    Call ValidateConnectionPreflight("SPConnect")
    If pConn Is Nothing Then Set pConn = New ADODB.Connection
    If pConn.state = adStateOpen Then Exit Sub
    
    Dim extendedProps As String
    If isReadOnly Then
        extendedProps = ";WSS;IMEX=1;RetrieveIds=Yes;"
    Else
        extendedProps = ";WSS;IMEX=0;RetrieveIds=Yes;"
    End If
    
    With pConn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0" & _
                            extendedProps & _
                            "DATABASE=" & pSiteURL & ";LIST={" & pListID & "}"
        .ConnectionTimeout = 30
        .Open
    End With
    
    Trace "SPConnect: Connected (ReadOnly=" & isReadOnly & ")"
    Exit Sub
ErrHandler:
    Trace "SPConnect ERROR: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "fwSPConn.SPConnect", Err.Description
End Sub

'******************************************************************************
' Nome do elemento: SPDisconnect (Sub)
' Resumo: Executa a rotina SPDisconnect.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub SPDisconnect()
    On Error Resume Next
    If Not pConn Is Nothing Then
        If pConn.state = adStateOpen Then pConn.Close
    End If
    Set pConn = Nothing
End Sub

'******************************************************************************
' Nome do elemento: SPLoadFieldInfoLegacy (Sub)
' Resumo: Executa a rotina SPLoadFieldInfoLegacy.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Sub SPLoadFieldInfoLegacy()
    On Error GoTo ErrHandler
    Dim rs As Object, fld As Object
    Call SPConnect(isReadOnly:=True) ' --- CORREÇÃO ---
    Set rs = New ADODB.Recordset
    rs.Open "SELECT TOP 1 * FROM [" & pListID & "]", pConn, adOpenStatic, adLockReadOnly
    Set pFields = New Collection
    For Each fld In rs.fields
        pFields.Add fld.Name
    Next
    rs.Close
    Call SPDisconnect
    Exit Sub
ErrHandler:
    If Not rs Is Nothing Then If rs.state = adStateOpen Then rs.Close
    Call SPDisconnect
    Err.Raise Err.Number, "fwSPConn.SPLoadFieldInfoLegacy", Err.Description
End Sub

'******************************************************************************
' Nome do elemento: SPGetFields (Function)
' Resumo: Retorna o resultado calculado por SPGetFields.
' @return Collection - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function SPGetFields() As Collection
    If pFields Is Nothing Then Call SPLoadFieldInfoLegacy
    Dim result As New Collection, v As Variant
    For Each v In pFields: result.Add v: Next
    Set SPGetFields = result
End Function

'******************************************************************************
' Nome do elemento: SPRunQuery (Function)
' Resumo: Retorna o resultado calculado por SPRunQuery.
' @param sql - Parâmetro da rotina.
' @param execReport - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function SPRunQuery(ByVal sql As String, Optional ByRef execReport As Variant) As Object
    On Error GoTo ErrHandler
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim i As Long

    If Len(Trim$(sql)) = 0 Then Err.Raise vbObjectError + 9810, "fwSPConn.SPRunQuery", "SQL obrigatório."
    Call SPLogEvento("INFO", "SPRunQuery", "Iniciando SELECT em lista SharePoint.", sql)

    Trace "SPRunQuery ENTER; SQL: " & sql
    Dim isSelect As Boolean: isSelect = Me.helpers.IsSelectSql(sql)

    ' SOMENTE SELECT é suportado por esta versão detalhada
    If Not isSelect Then
        Err.Raise vbObjectError + 9811, "fwSPConn.SPRunQuery", "Esta versão de SPRunQuery só suporta SELECT. Use SPExecuteNonQuery para DML."
    End If

    Call SPConnect(isReadOnly:=True) ' Sempre ReadOnly para SELECT

    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset") ' Late binding
    Trace "[DEBUG] SPRunQuery: Executando rs.Open..."
    rs.Open sql, pConn, adOpenStatic, adLockReadOnly
    Trace "[DEBUG] SPRunQuery: rs.Open concluído."

    Dim fieldCount As Long: fieldCount = 0
    On Error Resume Next ' Caso fields.Count falhe
    fieldCount = rs.fields.Count
    On Error GoTo ErrHandler ' Restaura handler
    Trace "[DEBUG] SPRunQuery: rs.Fields.Count: " & fieldCount
    
    Dim headersArr() As String
    Dim sourceTypes() As Long ' Array para tipos ADO
            
    If fieldCount > 0 Then
        Dim fld As Object ' Late binding
        ReDim headersArr(0 To fieldCount - 1)
        ReDim sourceTypes(0 To fieldCount - 1)
        For i = 0 To fieldCount - 1
            Set fld = rs.fields(i)
            headersArr(i) = fld.Name
            sourceTypes(i) = fld.Type ' Captura o tipo ADO
            Trace "[DEBUG] SPRunQuery: Campo " & i & ": '" & fld.Name & "' (Tipo ADO: " & fld.Type & ")"
        Next i
    Else
         Trace "[DEBUG] SPRunQuery: Recordset retornado sem campos."
         headersArr = Array() ' Array vazio
         sourceTypes = Array() ' Array vazio
    End If
    
    result.Add "Headers", headersArr ' Adiciona headers mesmo se vazio
    result.Add "Types", sourceTypes  ' Adiciona types mesmo se vazio
    
    If Not rs.EOF Then
        Trace "[DEBUG] SPRunQuery: Recordset contém dados. Executando GetRows..."
        Dim dataArr As Variant: dataArr = rs.GetRows()
        Trace "[DEBUG] SPRunQuery: GetRows concluído. Tipo de Var: " & TypeName(dataArr)
        result.Add "Data", dataArr
        Dim rowCount As Long
        On Error Resume Next
        rowCount = UBound(dataArr, 2) - LBound(dataArr, 2) + 1
        If Err.Number <> 0 Then rowCount = 0
        On Error GoTo ErrHandler
        If Not result.Exists("rowsAffected") Then result.Add "rowsAffected", rowCount
    Else
        Trace "[DEBUG] SPRunQuery: Recordset está vazio (EOF=True)."
        result.Add "Data", Empty ' Usa Empty para indicar ausência de dados
        If Not result.Exists("rowsAffected") Then result.Add "rowsAffected", 0
    End If

    If Not rs Is Nothing Then
        If rs.state = adStateOpen Then rs.Close
        Set rs = Nothing
    End If

    Call SPDisconnect
    Set SPRunQuery = result
    If result.Exists("rowsAffected") Then
        Call SPLogEvento("INFO", "SPRunQuery", "SELECT concluído. rowsAffected=" & CStr(result("rowsAffected")) & ".", sql)
    Else
        Call SPLogEvento("INFO", "SPRunQuery", "SELECT concluído sem rowsAffected explícito.", sql)
    End If
    Trace "[DEBUG] SPRunQuery: Saindo."
    Exit Function

ErrHandler:
    Dim errN As Long: errN = Err.Number
    Dim errD As String: errD = Err.Description
    Trace "[DEBUG] SPRunQuery: !!! ERRO: " & errN & " - " & errD
    On Error Resume Next
    If Len(sql) > 0 Then Trace "[DEBUG] SPRunQuery: SQL causando erro: " & sql
    If Not rs Is Nothing Then If rs.state = adStateOpen Then rs.Close
    Set rs = Nothing
    Call SPDisconnect
    Call SPLogEvento("ERROR", "SPRunQuery", "Erro durante SELECT: " & errD, sql)
    On Error GoTo 0
    Err.Raise errN, "fwSPConn.SPRunQuery", errD ' Re-lança o erro
End Function

'******************************************************************************
' Nome do elemento: SPUpdateRows (Function)
' Resumo: Retorna o resultado calculado por SPUpdateRows.
' @param getRowsArr - Parâmetro da rotina.
' @param columnsParam - Parâmetro da rotina.
' @param batchSize - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function SPUpdateRows(ByVal getRowsArr As Variant, ByVal columnsParam As Variant, Optional ByVal batchSize As Long = 100) As Long
    On Error GoTo Cleanup
    Dim totalAffected As Long ' A contagem real de linhas afetadas pelo pConn.Execute
    Dim lastSql As String
    totalAffected = 0
    Call SPLogEvento("INFO", "SPUpdateRows", "Iniciando atualização em lote de registros SharePoint.")
    
    Trace "[DEBUG-SYNC] SPUpdateRows: ENTER - Iniciando sincronização linha a linha."
    
    ' CORREÇÃO DE SINTAXE: Separa Trace e Exit Function
    If Me.helpers.IsVarEmpty(getRowsArr) Then
        Trace "[DEBUG-SYNC] SPUpdateRows: Array de dados vazio. Saindo."
        Call SPLogEvento("WARN", "SPUpdateRows", "Array de dados vazio. Nenhuma atualização executada.")
        Exit Function
    End If
    
    Trace "[DEBUG-SYNC] SPUpdateRows: list=" & pListID & "; batchSize=" & batchSize
    Dim headers() As String: headers = Me.helpers.NormalizeFieldsParam(columnsParam)
    
    Trace "[DEBUG-SYNC] SPUpdateRows: Headers normalizados: " & Join(headers, ", ")

    Dim pkFieldIndex As Long: pkFieldIndex = -1
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If StrComp(headers(i), Me.spKeyField, vbTextCompare) = 0 Then ' Usa SPKeyField
            pkFieldIndex = i
            Exit For
        End If
    Next i
    
    If pkFieldIndex = -1 Then
        Trace "[DEBUG-SYNC] SPUpdateRows: ERRO: Coluna chave '" & Me.spKeyField & "' não encontrada no array de dados de sincronização."
        Err.Raise vbObjectError + 9752, "fwSPConn.SPUpdateRows", "A coluna chave '" & Me.spKeyField & "' não foi encontrada."
    End If
    
    Dim r As Long, c As Long, rowCounter As Long
    
    ' Determina limites do array (assumindo 2D vindo do ADO GetRows)
    Dim lowerBound1 As Long, lowerBound2 As Long, upperBound2 As Long, totalRows As Long
    On Error Resume Next ' Trata caso de array 1D
    lowerBound1 = LBound(getRowsArr, 1) ' Deve ser 0
    lowerBound2 = LBound(getRowsArr, 2) ' Deve ser 0
    upperBound2 = UBound(getRowsArr, 2) ' Último índice de linha
    If Err.Number <> 0 Then
        Err.Clear
        lowerBound1 = LBound(getRowsArr)
        lowerBound2 = 0 ' Assume linha única virtual
        upperBound2 = 0
        totalRows = 1
        Trace "[DEBUG-SYNC] SPUpdateRows: Detectado array 1D (linha única?)."
    Else
        totalRows = (upperBound2 - lowerBound2 + 1)
    End If
    On Error GoTo Cleanup
    
    Trace "[DEBUG-SYNC] SPUpdateRows: Total de linhas para atualizar: " & totalRows
    
    For r = lowerBound2 To upperBound2
        ' Lógica de Batching e Conexão
        If rowCounter = 0 Then
            Trace "[DEBUG-SYNC] SPUpdateRows: Abrindo nova conexão (Batch Start)..."
            Call SPConnect(isReadOnly:=False)
            Trace "[DEBUG-SYNC] SPUpdateRows: Conexão aberta."
        End If
        
        Dim processedRows As Long
        processedRows = (r - lowerBound2 + 1)
        Call Me.helpers.NotifyProgress(processedRows, totalRows, "Sincronizando SharePoint...", 10)
        
        ' 1. Reconstrução da Cláusula SET
        Dim pkValue As Variant: pkValue = getRowsArr(pkFieldIndex + lowerBound1, r)
        Dim setClauses As String: setClauses = ""
        
        For c = LBound(headers) To UBound(headers)
            If c <> pkFieldIndex Then
                Dim colName As String: colName = headers(c)
                Dim newValue As Variant
                On Error Resume Next
                newValue = getRowsArr(c + lowerBound1, r)
                If Err.Number <> 0 Then
                    Trace "[DEBUG-SYNC] SPUpdateRows: Aviso - Coluna '" & colName & "' (índice " & c & ") não encontrada nos dados da linha " & r & ". Ignorando."
                    Err.Clear
                Else
                    If Len(setClauses) > 0 Then setClauses = setClauses & ", "
                    setClauses = setClauses & "[" & colName & "] = " & Me.helpers.FormatSQLValue(newValue)
                End If
                On Error GoTo Cleanup
            End If
        Next c
        
        ' 2. Reconstrução e Execução do SQL
        If Len(setClauses) > 0 Then
            Dim sql As String
            Dim recordsAffected As Long

            sql = "UPDATE [" & pListID & "]" & " SET " & setClauses & " WHERE [" & Me.spKeyField & "] = " & Me.helpers.FormatSQLValue(pkValue, True)
            lastSql = sql
            Trace "[DEBUG-SYNC] SPUpdateRows: Executando SQL: " & sql
            
            On Error Resume Next
            pConn.Execute sql, recordsAffected
            
            If Err.Number <> 0 Then
                Trace "[DEBUG-SYNC] SPUpdateRows: !!! ERRO DURANTE EXECUÇÃO (ID " & pkValue & "): " & Err.Number & " - " & Err.Description
                Err.Clear
            Else
                totalAffected = totalAffected + IIf(IsNull(recordsAffected), 0, recordsAffected)
                Trace "[DEBUG-SYNC] SPUpdateRows: Execução OK. Registros afetados: " & recordsAffected
            End If
            On Error GoTo Cleanup
        Else
             Trace "[DEBUG-SYNC] SPUpdateRows: Nenhuma cláusula SET construída para ID " & pkValue & ". Pulando execução."
        End If
        
        ' 3. Lógica de Batching e Fechamento
        rowCounter = rowCounter + 1
        
        If rowCounter >= batchSize Or r = upperBound2 Then
            Trace "[DEBUG-SYNC] SPUpdateRows: Fim do lote/loop. Fechando conexão."
            Call SPDisconnect
            rowCounter = 0
        End If
    Next r
    
    SPUpdateRows = totalAffected

Cleanup:
    Application.StatusBar = False
    If Err.Number <> 0 Then
        Trace "[DEBUG-SYNC] SPUpdateRows: ERRO FINAL. Desconectando..."
        Call SPDisconnect
        Call Me.helpers.AppendPersistentLog("ERROR", "fwSPConn", "SPUpdateRows", Err.Description, lastSql)
        Trace "SPUpdateRows ERROR: " & Err.Number & " - " & Err.Description
        Err.Raise Err.Number, "fwSPConn.SPUpdateRows", Err.Description
    Else
        Trace "[DEBUG-SYNC] SPUpdateRows: EXIT. Total afetado: " & totalAffected
        Call SPLogEvento("INFO", "SPUpdateRows", "Atualização em lote concluída. Total afetado=" & totalAffected & ".")
    End If
End Function

'******************************************************************************
' Nome do elemento: SPInsertRows (Function)
' Resumo: Retorna o resultado calculado por SPInsertRows.
' @param getRowsArr - Parâmetro da rotina.
' @param columnsParam - Parâmetro da rotina.
' @param execReport - Parâmetro da rotina.
' @param batchSize - Parâmetro da rotina.
' @return Variant - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function SPInsertRows(ByVal getRowsArr As Variant, ByVal columnsParam As Variant, Optional ByRef execReport As Variant, Optional ByVal batchSize As Long = 50) As Variant
    On Error GoTo ErrHandler
    If Me.helpers.IsVarEmpty(getRowsArr) Then
        Call SPLogEvento("WARN", "SPInsertRows", "Array de dados vazio. Nenhum INSERT executado.")
        Exit Function
    End If
    
    Call SPLogEvento("INFO", "SPInsertRows", "Iniciando INSERT em lote na lista SharePoint.")
    Trace "SPInsertRows ENTER; list=" & pListID
    Dim hdrs() As String: hdrs = Me.helpers.NormalizeFieldsParam(columnsParam)
    Dim fL As Long, fU As Long, rL As Long, rU As Long
    fL = LBound(getRowsArr, 1): fU = UBound(getRowsArr, 1)
    rL = LBound(getRowsArr, 2): rU = UBound(getRowsArr, 2)
    Dim arrFldCount As Long: arrFldCount = fU - fL + 1
    
    If (UBound(hdrs) + 1) <> arrFldCount Then
        Err.Raise vbObjectError + 9713, "fwSPConn.SPInsertRows", "Disparidade de colunas."
    End If
    
    Dim colList As String: colList = "[" & Join(hdrs, "],[") & "]"
    Dim processed As Long, totalAffected As Long, r As Long
    Dim totalRows As Long: totalRows = (rU - rL + 1)
    Dim lastSql As String
    
    Call SPConnect(isReadOnly:=False) ' --- CORREÇÃO ---
    
    For r = rL To rU
        processed = processed + 1
        Call Me.helpers.NotifyProgress(processed, totalRows, "Inserindo registros no SharePoint...", 10)

        Dim vals() As String
        ReDim vals(0 To arrFldCount - 1)
        Dim c As Long
        For c = 0 To arrFldCount - 1
            Dim v As Variant: v = getRowsArr(fL + c, r)
            vals(c) = Me.helpers.FormatSQLValue(v, False)
        Next c
        
        Dim sql As String
        sql = "INSERT INTO [" & pListID & "] (" & colList & ") VALUES (" & Join(vals, ",") & ")"
        lastSql = sql
        
        pConn.Execute sql
        totalAffected = totalAffected + 1
    Next r

    Dim out As Variant: out = Array(processed, totalAffected, Array())
    SPInsertRows = out
    Call SPLogEvento("INFO", "SPInsertRows", "INSERT em lote concluído. Processados=" & processed & "; Afetados=" & totalAffected & ".")
    
    Call SPDisconnect
    Application.StatusBar = False
    Exit Function
    
ErrHandler:
    Application.StatusBar = False
    Call SPDisconnect
    Call Me.helpers.AppendPersistentLog("ERROR", "fwSPConn", "SPInsertRows", Err.Description, lastSql)
    Trace "SPInsertRows ERROR: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "fwSPConn.SPInsertRows", Err.Description
End Function

'******************************************************************************
' Nome do elemento: OpenConnection (Sub)
' Resumo: Executa a rotina OpenConnection.
' @param isReadOnly - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Sub OpenConnection(ByVal isReadOnly As Boolean)
    On Error GoTo ErrHandler
    Trace "OpenConnection: Tentando abrir conexão (ReadOnly=" & isReadOnly & ")"
    ' Chama a rotina interna que já existe
    Call ValidateConnectionPreflight("OpenConnection")
    Call SPConnect(isReadOnly)
    Exit Sub
ErrHandler:
    Trace "OpenConnection ERROR: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "fwSPConn.OpenConnection", Err.Description
End Sub

'******************************************************************************
' Nome do elemento: CloseConnection (Sub)
' Resumo: Executa a rotina CloseConnection.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub CloseConnection()
    On Error Resume Next
    Trace "CloseConnection: Fechando conexão."
    ' Chama a rotina interna que já existe
    Call SPDisconnect
End Sub

'******************************************************************************
' Nome do elemento: ExecuteNonQueryOverOpenConnection (Function)
' Resumo: Retorna o resultado calculado por ExecuteNonQueryOverOpenConnection.
' @param sql - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function ExecuteNonQueryOverOpenConnection(ByVal sql As String) As Long
    On Error GoTo ErrHandler
    Dim recordsAffected As Long
    
    If pConn Is Nothing Then
        Err.Raise vbObjectError + 9801, "ExecuteNonQueryOverOpenConnection", "A conexão não está aberta. Chame OpenConnection primeiro."
    End If
    If pConn.state = 0 Then ' adStateClosed
         Err.Raise vbObjectError + 9802, "ExecuteNonQueryOverOpenConnection", "A conexão está fechada. Chame OpenConnection primeiro."
    End If
    
    pConn.Execute sql, recordsAffected
    ExecuteNonQueryOverOpenConnection = recordsAffected
    Exit Function
    
ErrHandler:
    Trace "ExecuteNonQueryOverOpenConnection ERROR: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "fwSPConn.ExecuteNonQueryOverOpenConnection", Err.Description
End Function

'******************************************************************************
' Nome do elemento: SPExecuteNonQuery (Function)
' Resumo: Retorna o resultado calculado por SPExecuteNonQuery.
' @param sql - Parâmetro da rotina.
' @param execReport - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function SPExecuteNonQuery(ByVal sql As String, Optional ByRef execReport As Variant) As Long
    On Error GoTo ErrHandler
    Call SPLogEvento("INFO", "SPExecuteNonQuery", "Iniciando comando DML sobre SharePoint.", sql)
    
    ' Abre a conexão em modo de escrita
    Call Me.OpenConnection(isReadOnly:=False)
    
    ' Executa o comando sobre a conexão aberta
    SPExecuteNonQuery = Me.ExecuteNonQueryOverOpenConnection(sql)
    Call SPLogEvento("INFO", "SPExecuteNonQuery", "Comando DML concluído. rowsAffected=" & CStr(SPExecuteNonQuery) & ".", sql)
    
    ' Fecha a conexão ao final
    Call Me.CloseConnection
    
    Exit Function

ErrHandler:
    ' Garante a desconexão em caso de erro
    Call Me.CloseConnection
    Call SPLogEvento("ERROR", "SPExecuteNonQuery", "Erro durante comando DML: " & Err.Description, sql)
    Err.Raise Err.Number, "fwSPConn.SPExecuteNonQuery", Err.Description
End Function

'******************************************************************************
' Nome do elemento: SPDeleteRowsByQuery (Function)
' Resumo: Retorna o resultado calculado por SPDeleteRowsByQuery.
' @param deleteSql - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function SPDeleteRowsByQuery(ByVal deleteSql As String) As Long
    On Error GoTo Cleanup
    Dim rowsAffected As Long
    Dim lastSql As String
    
    Trace "SPDeleteRowsByQuery ENTER; SQL: " & deleteSql
    Call SPLogEvento("INFO", "SPDeleteRowsByQuery", "Iniciando DELETE por query.", deleteSql)
    lastSql = deleteSql
    
    ' 1. Analisa o SQL para extrair a tabela e a cláusula WHERE
    Dim parsedSql As Object
    Set parsedSql = Me.helpers.ParseDeleteSqlDetails(deleteSql)
    If parsedSql Is Nothing Then
        Err.Raise vbObjectError + 9901, "SPDeleteRowsByQuery", "Falha ao analisar o comando DELETE."
    End If
    
    Dim targetTable As String: targetTable = parsedSql("TargetTable")
    Dim whereClause As String: whereClause = parsedSql("WhereClause")
    
    ' 2. Transforma o DELETE em um SELECT para obter os IDs das linhas a serem deletadas
    Dim selectSql As String
    selectSql = "SELECT [ID] FROM [" & targetTable & "] " & whereClause
    Trace "SPDeleteRowsByQuery: Executando consulta de pré-seleção: " & selectSql
    
    ' Executa a consulta de seleção (usando o próprio SPRunQuery)
    Dim result As Object: Set result = Me.SPRunQuery(selectSql)
    
    If Not result.Exists("Data") Then
        Trace "SPDeleteRowsByQuery: Nenhuma linha encontrada para deletar."
        Call SPLogEvento("INFO", "SPDeleteRowsByQuery", "Nenhuma linha encontrada para DELETE.", deleteSql)
        Exit Function
    End If
    
    ' 3. Monta uma lista com os IDs a serem deletados
    Dim idsToDelete As Variant: idsToDelete = result("Data")
    Dim totalToDelete As Long: totalToDelete = UBound(idsToDelete, 2) + 1
    
    ' 4. Itera sobre os IDs e deleta um por um, com barra de progresso
    ' --- INÍCIO DA CORREÇÃO ---
    ' Especifica que a conexão deve ser aberta em modo de escrita.
    Call SPConnect(isReadOnly:=False)
    ' --- FIM DA CORREÇÃO ---
    
    Dim i As Long
    For i = LBound(idsToDelete, 2) To UBound(idsToDelete, 2)
        Dim currentId As Variant: currentId = idsToDelete(0, i)
        Call Me.helpers.NotifyProgress((i - LBound(idsToDelete, 2) + 1), totalToDelete, "Excluindo registros no SharePoint...", 10)
        
        ' Executa o DELETE simples
        Dim singleDeleteSql As String
        singleDeleteSql = "DELETE FROM [" & targetTable & "] WHERE [ID] = " & Me.helpers.FormatSQLValue(currentId, True)
        lastSql = singleDeleteSql
        
        pConn.Execute singleDeleteSql
        rowsAffected = rowsAffected + 1
    Next i

    SPDeleteRowsByQuery = rowsAffected
    Call SPLogEvento("INFO", "SPDeleteRowsByQuery", "DELETE concluído. Linhas removidas=" & rowsAffected & ".", deleteSql)

Cleanup:
    Application.StatusBar = False
    Call SPDisconnect
    If Err.Number <> 0 Then
        Call Me.helpers.AppendPersistentLog("ERROR", "fwSPConn", "SPDeleteRowsByQuery", Err.Description, lastSql)
        Trace "SPDeleteRowsByQuery ERROR: " & Err.Number & " - " & Err.Description
        Err.Raise Err.Number, "fwSPConn.SPDeleteRowsByQuery", Err.Description
    End If
End Function

'******************************************************************************
' Nome do elemento: SPUpdateRowsByQuery (Function)
' Resumo: Retorna o resultado calculado por SPUpdateRowsByQuery.
' @param updateSql - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function SPUpdateRowsByQuery(ByVal updateSql As String) As Long
    On Error GoTo Cleanup
    Dim rowsAffected As Long
    Dim lastSql As String
    rowsAffected = 0 ' Garantir que começa em zero

    Trace "SPUpdateRowsByQuery ENTER; SQL: " & updateSql
    Call SPLogEvento("INFO", "SPUpdateRowsByQuery", "Iniciando UPDATE por query.", updateSql)
    Trace "[DEBUG] SPUpdateRowsByQuery: Iniciando execução..."
    lastSql = updateSql

    ' 1. Analisa o SQL
    Dim parsedSql As Object
    Set parsedSql = Me.helpers.ParseUpdateJoinSqlDetails(updateSql) ' Usa o parser de JOIN mesmo para UPDATE simples
    If parsedSql Is Nothing Then
        Trace "[DEBUG] SPUpdateRowsByQuery: ERRO FATAL: Falha ao analisar o SQL inicial."
        Err.Raise vbObjectError + 9911, "SPUpdateRowsByQuery", "Falha ao analisar o comando UPDATE."
    End If
    Trace "[DEBUG] SPUpdateRowsByQuery: SQL inicial analisado com sucesso."

    Dim targetTable As String: targetTable = parsedSql("TargetTable")
    Dim setClause As String: setClause = parsedSql("SetClause")
    Dim whereClause As String: whereClause = parsedSql("WhereClause") ' Pode ser "" se não houver WHERE
    Dim tablesClause As String: tablesClause = parsedSql("TablesClause")
    Trace "[DEBUG] SPUpdateRowsByQuery: TargetTable: " & targetTable
    Trace "[DEBUG] SPUpdateRowsByQuery: SetClause: " & setClause
    Trace "[DEBUG] SPUpdateRowsByQuery: WhereClause: " & whereClause
    Trace "[DEBUG] SPUpdateRowsByQuery: TablesClause: " & tablesClause

    Dim targetAlias As String, tbl As Object
    ' Encontra o alias da tabela alvo (mesmo que seja igual ao nome da tabela)
    For Each tbl In parsedSql("Tables")
        If tbl("Name") = targetTable Then
            targetAlias = tbl("Alias")
            Exit For
        End If
    Next
    If Len(targetAlias) = 0 Then targetAlias = targetTable ' Fallback se algo der errado no parser de tabela
    Trace "[DEBUG] SPUpdateRowsByQuery: TargetAlias encontrado: " & targetAlias

    ' 2. Executa o SELECT com JOIN para obter os IDs das linhas alvo
    Dim selectIdSql As String
    Dim selectIdFieldExpr As String
    Dim hasJoinInUpdate As Boolean
    Dim safeTargetQualifier As String
    hasJoinInUpdate = (parsedSql("Tables").Count > 1) Or (InStr(1, tablesClause, " JOIN ", vbTextCompare) > 0)

    If hasJoinInUpdate Then
        safeTargetQualifier = Me.helpers.Unbracket(targetAlias)
        If Len(safeTargetQualifier) = 0 Then safeTargetQualifier = Me.helpers.Unbracket(targetTable)
        selectIdFieldExpr = "[" & safeTargetQualifier & "].[" & Me.helpers.Unbracket(Me.spKeyField) & "]"
    Else
        selectIdFieldExpr = "[" & Me.helpers.Unbracket(Me.spKeyField) & "]"
    End If

    selectIdSql = "SELECT " & selectIdFieldExpr & " FROM " & tablesClause
    If Len(whereClause) > 0 Then selectIdSql = selectIdSql & " " & whereClause
    Trace "SPUpdateRowsByQuery: Executando consulta de pré-seleção de IDs: " & selectIdSql
    Trace "[DEBUG] SPUpdateRowsByQuery: SQL de pré-seleção de IDs: " & selectIdSql

    Dim result As Object: Set result = Me.SPRunQuery(selectIdSql) ' Chama SPRunQuery para obter os IDs

    If Not result.Exists("Data") Or Me.helpers.IsVarEmpty(result("Data")) Then ' Verifica se Data existe E não está vazio/nulo
        Trace "SPUpdateRowsByQuery: Nenhuma linha encontrada para atualizar (Result('Data') não existe ou está vazio)."
        Trace "[DEBUG] SPUpdateRowsByQuery: Nenhuma linha encontrada na pré-seleção. Saindo da função."
        Call SPLogEvento("INFO", "SPUpdateRowsByQuery", "Nenhuma linha encontrada para UPDATE.", updateSql)
        GoTo Cleanup ' Pula para a limpeza, retornará 0
    End If
    Trace "[DEBUG] SPUpdateRowsByQuery: Pré-seleção retornou dados."

    ' 3. Itera sobre os IDs e executa um UPDATE com JOIN para cada um
    Dim idsToUpdate As Variant: idsToUpdate = result("Data")
    Dim totalToUpdate As Long
    Dim is2DArray As Boolean
    Dim lowerBound1 As Long, upperBound1 As Long, lowerBound2 As Long, upperBound2 As Long

    ' Bloco robusto para determinar dimensões e limites do array de IDs
    On Error Resume Next
    upperBound2 = UBound(idsToUpdate, 2) ' Tenta acessar a 2a dimensão
    is2DArray = (Err.Number = 0)
    Err.Clear

    If is2DArray Then
        lowerBound1 = LBound(idsToUpdate, 1) ' Deve ser 0
        upperBound1 = UBound(idsToUpdate, 1) ' Deve ser 0 (só a coluna ID)
        lowerBound2 = LBound(idsToUpdate, 2) ' Deve ser 0
        upperBound2 = UBound(idsToUpdate, 2) ' Índice da última linha
        totalToUpdate = upperBound2 - lowerBound2 + 1
        Trace "[DEBUG] SPUpdateRowsByQuery: Array de IDs é 2D. Dim1: " & lowerBound1 & " To " & upperBound1 & "; Dim2: " & lowerBound2 & " To " & upperBound2
    Else
        lowerBound1 = LBound(idsToUpdate) ' Limite inferior da única dimensão
        upperBound1 = UBound(idsToUpdate) ' Limite superior da única dimensão
        is2DArray = False ' Redundante, mas explícito
        totalToUpdate = upperBound1 - lowerBound1 + 1
        Trace "[DEBUG] SPUpdateRowsByQuery: Array de IDs é 1D. LBound: " & lowerBound1 & "; UBound: " & upperBound1
    End If
    On Error GoTo Cleanup ' Restaura o tratamento de erro padrão

    Trace "[DEBUG] SPUpdateRowsByQuery: Calculado totalToUpdate = " & totalToUpdate

    If totalToUpdate <= 0 Then
        Trace "SPUpdateRowsByQuery: Cálculo resultou em 0 linhas para atualizar."
        Trace "[DEBUG] SPUpdateRowsByQuery: totalToUpdate <= 0. Nenhuma linha será atualizada. Saindo."
        GoTo Cleanup ' Pula para a limpeza, retornará 0
    End If

    Trace "[DEBUG] SPUpdateRowsByQuery: Abrindo conexão em modo de escrita..."
    Call SPConnect(isReadOnly:=False)
    Trace "[DEBUG] SPUpdateRowsByQuery: Conexão de escrita aberta."

    Dim i As Long
    Dim currentId As Variant
    Dim singleUpdateSql As String
    Dim recordsAffectedIteration As Long ' Para capturar o resultado de cada Execute

    Trace "[DEBUG] SPUpdateRowsByQuery: *** INICIANDO LOOP DE UPDATE INDIVIDUAL *** (Total de iterações esperadas: " & totalToUpdate & ")"
    For i = 0 To totalToUpdate - 1 ' Loop baseado em contagem (0 a N-1)
        ' Extração cuidadosa do ID baseado na dimensionalidade
        If is2DArray Then
            currentId = idsToUpdate(lowerBound1, lowerBound2 + i) ' (0, 0+i)
        Else
            currentId = idsToUpdate(lowerBound1 + i) ' (0+i)
        End If
        Trace "[DEBUG] SPUpdateRowsByQuery: Loop i = " & i & ", currentId = " & currentId

        Call Me.helpers.NotifyProgress((i + 1), totalToUpdate, "Atualizando registros no SharePoint...", 10)

        ' # CORREÇÃO #: Remove o prefixo do alias da tabela na cláusula WHERE e usa FormatSQLValue para número.
        singleUpdateSql = "UPDATE " & tablesClause & " SET " & setClause & " WHERE [" & Me.spKeyField & "] = " & Me.helpers.FormatSQLValue(currentId, True)
        lastSql = singleUpdateSql
        ' Linha Anterior (com erro de alias e aspas):
        ' singleUpdateSql = "UPDATE " & tablesClause & " SET " & setClause & " WHERE " & targetAlias & ".[" & Me.SPKeyField & "] = '" & CStr(currentId) & "'"
        ' Linha Original (com erro de alias):
        ' singleUpdateSql = "UPDATE " & tablesClause & " SET " & setClause & " WHERE " & targetAlias & ".[" & Me.SPKeyField & "] = " & Me.helpers.FormatSQLValue(currentId, True)

        Trace "[DEBUG] SPUpdateRowsByQuery: SQL para Execução (i=" & i & "): " & singleUpdateSql

        recordsAffectedIteration = 0 ' Reseta para esta iteração
        Trace "[DEBUG] SPUpdateRowsByQuery: Executando pConn.Execute..."
        On Error Resume Next ' Temporariamente captura erros na execução individual
        pConn.Execute singleUpdateSql, recordsAffectedIteration
        If Err.Number <> 0 Then
             Trace "[DEBUG] SPUpdateRowsByQuery: !!! ERRO DURANTE pConn.Execute (i=" & i & "): " & Err.Number & " - " & Err.Description
             ' Decide o que fazer: parar, continuar, logar... por ora, vamos apenas logar e continuar
             Err.Clear ' Limpa o erro para não ir para o handler principal ainda
        Else
             Trace "[DEBUG] SPUpdateRowsByQuery: pConn.Execute concluído. recordsAffectedIteration = " & recordsAffectedIteration

             Dim rowsAppliedIteration As Long
             rowsAppliedIteration = CLng(IIf(IsNull(recordsAffectedIteration), 0, recordsAffectedIteration))
             If rowsAppliedIteration < 0 Then rowsAppliedIteration = 1
             rowsAffected = rowsAffected + rowsAppliedIteration
             Trace "[DEBUG] SPUpdateRowsByQuery: rowsAffected +=" & rowsAppliedIteration & " (acumulado=" & rowsAffected & ")"
        End If
        On Error GoTo Cleanup ' Restaura o tratamento de erro padrão

    Next i
    Trace "[DEBUG] SPUpdateRowsByQuery: *** FIM DO LOOP DE UPDATE INDIVIDUAL ***"

    Trace "[DEBUG] SPUpdateRowsByQuery: Valor final a ser retornado por SPUpdateRowsByQuery: " & rowsAffected
    SPUpdateRowsByQuery = rowsAffected ' Define o valor de retorno
    Call SPLogEvento("INFO", "SPUpdateRowsByQuery", "UPDATE concluído. Linhas afetadas=" & rowsAffected & ".", updateSql)

Cleanup:
    Dim finalErrNumber As Long
    Dim finalErrDescription As String
    finalErrNumber = Err.Number
    finalErrDescription = Err.Description

    Trace "[DEBUG] SPUpdateRowsByQuery: Entrando no bloco Cleanup..."
    Application.StatusBar = False
    Trace "[DEBUG] SPUpdateRowsByQuery: Fechando conexão..."
    Call SPDisconnect ' Garante que a conexão seja fechada
    Trace "[DEBUG] SPUpdateRowsByQuery: Conexão fechada."

    If finalErrNumber <> 0 Then
        Call Me.helpers.AppendPersistentLog("ERROR", "fwSPConn", "SPUpdateRowsByQuery", finalErrDescription, lastSql)
        Trace "SPUpdateRowsByQuery ERROR: " & finalErrNumber & " - " & finalErrDescription
        Trace "[DEBUG] SPUpdateRowsByQuery: !!! ERRO FINAL CAPTURADO NO HANDLER: " & finalErrNumber & " - " & finalErrDescription
        Err.Raise finalErrNumber, "fwSPConn.SPUpdateRowsByQuery", finalErrDescription
    Else
        Trace "[DEBUG] SPUpdateRowsByQuery: Saindo de SPUpdateRowsByQuery sem erros."
    End If
End Function

'******************************************************************************
' Nome do elemento: Class_Terminate (Sub)
' Resumo: Executa a rotina Class_Terminate.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub Class_Terminate()
    On Error Resume Next
    Call SPDisconnect
    Set pFields = Nothing
    Set pHelpers = Nothing
End Sub


