VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fwXLConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' Nome do elemento: fwXLConn (Classe)
' Resumo:
'   Motor de acesso a dados para um Workbook Excel via OLEDB/ADODB.
'
' Responsabilidades:
'   - Executar comandos SQL (SELECT, INSERT, UPDATE, DELETE) em ListObjects
'     e Ranges do Workbook.
'   - Gerenciar a conexão ADO (Microsoft ACE OLEDB).
'   - Traduzir sintaxe SQL para ser compatível com o dialeto ACE.
'
' Dependências:
'   - fwHelpers
'   - Microsoft Access Database Engine (ACE OLEDB) 12.0 ou superior.
'******************************************************************************

Option Explicit

'--- estado interno da conexão
Private pTraceOn        As Boolean
Private pHelpers        As fwHelpers

'=== ADO infra (late binding) ===
Private pCnn As Object            ' ADODB.Connection
Private pRst As Object            ' ADODB.Recordset
Private pConnString As String     ' ConnectionString efetiva usada na abertura

'******************************************************************************
' Nome do elemento: Init (Sub)
' Resumo: Executa a rotina Init.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub Init()
    ' A inicialização agora é mais simples, apenas prepara os helpers.
    ' O contexto do workbook será pego dinamicamente pelo ActiveWorkbook.
    EnsureHelpers
    Trace "Init: Motor de consulta fwXLConn inicializado."
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
' Nome do elemento: XLRunQuery (Function)
' Resumo: Traduz o SQL para o dialeto ACE, abre a conexão e executa comandos sobre o workbook.
' @param sql - Comando SQL a ser alinhado e executado.
' @param execReport - Dicionário opcional usado para relatar progresso e erros.
' @param forceWriteMode - Mantém a conexão em modo escrita mesmo para SELECT quando True.
' @return Object - Dicionário com Data/Headers/rowsAffected e mensagens auxiliares.
' @remarks Controla toda a pilha ADO, incluindo tradução de nomes e tratamento de erros do driver ACE.
'******************************************************************************
Public Function XLRunQuery( _
    ByVal sql As String, _
    Optional ByRef execReport As Variant, _
    Optional ByVal forceWriteMode As Boolean = False _
) As Object
    Dim t0 As Double
    Dim effSql As String, wbPath As String, isSelect As Boolean
    Dim rowsAffected As Variant, fieldsCount As Long, rowsCount As Long
    
    EnsureHelpers
    t0 = Me.helpers.TickCountMs()
    
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    
    On Error GoTo ErrHandler

    Dim wb As Workbook: Set wb = ActiveWorkbook
    
    wbPath = Me.helpers.GetLocalPathFromSyncUrl(ActiveWorkbook.FullName)
    If Len(wbPath) = 0 Then Err.Raise 702, "fwXLConn.XLRunQuery", "Caminho local do arquivo não encontrado."
    Trace "[DEBUG] XLRunQuery: Caminho do arquivo: " & wbPath
    
    ' 1. Linha sintaxe
    Dim alignedSql As String
    alignedSql = Me.helpers.SqlAlignToAce(sql)
    
    ' 2. Traduz nomes
    effSql = Me.helpers.RewriteDataSourceNames(alignedSql, wb)
    
    ' Determina se é SELECT
    isSelect = Me.helpers.IsSelectSql(effSql)
    
    If Not wb.Saved Then wb.Save

    Trace "XLRunQuery: Início | IsSelect=" & CStr(isSelect) & " | forceWriteMode=" & forceWriteMode & " | SQL=" & effSql
    Trace "[fwXLConn-DEBUG] SQL FINAL ENVIADO AO DRIVER: " & effSql
    
    ' --- CORREÇÃO DA CONEXÃO (IMEX) ---
    ' Conecta em modo ReadOnly (IMEX=1) APENAS SE:
    ' 1. For um SELECT (isSelect = True)
    ' 2. NÃO for forçado a usar modo de escrita (forceWriteMode = False)
    ' Em todos os outros casos (UPDATE/INSERT/DELETE ou SELECT forçado), usa ReadOnly=False (IMEX=0).
    Call XLConnect(wbPath, (isSelect And Not forceWriteMode))

    If isSelect Then
        Set pRst = CreateObject("ADODB.Recordset")
        pRst.Open effSql, pCnn, 0, 1, 1 ' adOpenForwardOnly, adLockReadOnly, adCmdText
        If Not pRst.EOF Then
            Dim dataArr As Variant: dataArr = pRst.GetRows()
            Dim headersArr() As String
            ReDim headersArr(0 To pRst.fields.Count - 1)
            Dim i As Long
            For i = 0 To pRst.fields.Count - 1: headersArr(i) = pRst.fields(i).Name: Next i
            result.Add "Data", dataArr
            result.Add "Headers", headersArr
            rowsCount = UBound(dataArr, 2) + 1
            fieldsCount = UBound(dataArr, 1) + 1
        End If
        result.Add "rowsAffected", rowsCount
        Call XLCloseRecordset
    Else ' DML (INSERT, UPDATE, DELETE)
        On Error Resume Next
        pCnn.Execute effSql, rowsAffected
        Dim errNum As Long: errNum = Err.Number
        Dim errDesc As String: errDesc = Err.Description
        On Error GoTo ErrHandler
        
        result.Add "rowsAffected", CLng(IIf(IsNull(rowsAffected), 0, rowsAffected))
        If errNum <> 0 Then
            Err.Raise errNum, "fwXLConn.XLRunQuery(Execute)", errDesc
        End If
    End If

    Call Me.helpers.FillExecReport(execReport, "OK", effSql, isSelect, rowsCount, fieldsCount, rowsAffected, 0, "", "", t0, Me.helpers.TickCountMs())
    Trace "XLRunQuery: Fim | OK"
    
    Call XLDisconnect
    Set XLRunQuery = result
    Exit Function

ErrHandler:
    Trace "XLRunQuery: ERRO " & Err.Number & " - " & Err.Description
    Call XLCloseRecordset
    Call XLDisconnect
    If result.Exists("rowsAffected") Then rowsAffected = result("rowsAffected")
    Call Me.helpers.FillExecReport(execReport, "ERROR", effSql, isSelect, 0, 0, rowsAffected, Err.Number, "fwXLConn.XLRunQuery", Err.Description, t0, Me.helpers.TickCountMs())
    Set XLRunQuery = result
End Function

'******************************************************************************
' Nome do elemento: XLConnect (Sub)
' Resumo: Executa a rotina XLConnect.
' @param wbFullPath - Parâmetro da rotina.
' @param isReadOnly - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Sub XLConnect(ByVal wbFullPath As String, ByVal isReadOnly As Boolean)
    Dim cs As String, extendedProps As String, errMessage As String
    
    If pCnn Is Nothing Then Set pCnn = CreateObject("ADODB.Connection")
    If pCnn.state <> 0 Then Exit Sub ' Se já estiver aberta, não faz nada

    ' --- INÍCIO DA CORREÇÃO ---
    ' Define propriedade fixas (baseado na cDBExcel)
    ' HDR=Yes: O framework usa nomes de cabeçalho.
    ' IMEX=0: Força o driver a usar os tipos de dados formatados (respeita Texto/@ e DummyRow)
    ' Readonly=False: Garante que a conexão possa ler e escrever.
    extendedProps = "Excel 12.0;HDR=Yes;IMEX=0;Readonly=False"
    
    Trace "XLConnect: Conectando com propriedades fixas (IMEX=0) para consistência..."
    ' --- FIM DA CORREÇÃO ---

    ' --- Lógica de conexão (mantida) ---
    On Error Resume Next ' Inicia o tratamento de erro local
    
    ' Tentativa 1: Provedor 16.0
    cs = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & wbFullPath & ";Extended Properties=""" & extendedProps & """;"
    Trace "XLConnect: Tentando Provedor 16.0..."
    pCnn.Open cs
    
    If Err.Number <> 0 Then
        ' Se falhou, guarda a mensagem de erro e tenta o fallback
        errMessage = "ACE 16.0 Error: " & Err.Description & vbCrLf
        Trace "XLConnect: Falha no 16.0. " & errMessage
        Err.Clear
        
        ' Tentativa 2: Provedor 12.0 (Fallback)
        cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & wbFullPath & ";Extended Properties=""" & extendedProps & """;"
        Trace "XLConnect: Tentando Provedor 12.0 (Fallback)..."
        pCnn.Open cs
        
        If Err.Number <> 0 Then
            ' Se o fallback também falhou, agora sim geramos um erro claro
            errMessage = errMessage & "ACE 12.0 Error: " & Err.Description
            Trace "XLConnect: Falha no 12.0. " & errMessage
            On Error GoTo 0 ' Reseta o handler de erro
            Err.Raise vbObjectError + 711, "fwXLConn.XLConnect", "Falha ao conectar com ambos os provedores ACE (16.0 e 12.0). Verifique a instalação do Microsoft Access Database Engine. Detalhes: " & errMessage
        End If
    End If
    
    On Error GoTo 0 ' Reseta o handler de erro se a conexão foi bem sucedida
    pConnString = cs
    Trace "XLConnect: Conectado com sucesso. (Driver: " & pCnn.Provider & ", IMEX=0)"
End Sub

'******************************************************************************
' Nome do elemento: XLDisconnect (Sub)
' Resumo: Executa a rotina XLDisconnect.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub XLDisconnect()
    On Error Resume Next
    If Not pRst Is Nothing Then
        If pRst.state <> 0 Then pRst.Close
        Set pRst = Nothing
    End If
    If Not pCnn Is Nothing Then
        If pCnn.state <> 0 Then pCnn.Close
        Set pCnn = Nothing
    End If
    pConnString = vbNullString
End Sub

'******************************************************************************
' Nome do elemento: XLCloseRecordset (Sub)
' Resumo: Executa a rotina XLCloseRecordset.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub XLCloseRecordset()
    On Error Resume Next
    If Not pRst Is Nothing Then
        If pRst.state <> 0 Then pRst.Close
        Set pRst = Nothing
    End If
End Sub



