VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fwSAPConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' Nome do elemento: fwSAPConn (Classe)
' Resumo: Inicia SAP Logon se necessário, abre ou localiza a conexão desejada,
' garante uma sessão ociosa (idle) e expõe o objeto Session pronto para uso.
' @remarks Autor: Rinaldo Paulino de Souza; Versão: 1.0 – 2025-06-20.
'******************************************************************************

Option Explicit

'==========================================================================
'== Private state =========================================================
'==========================================================================
Private cSession    As Object   ' live GuiSession entregue aos chamadores
Private cConnection As Object   ' GuiConnection que detém a sessão
Private cSapGuiApp  As Object   ' GuiApplication / ScriptingEngine raiz

'==========================================================================
'== Público API ============================================================
'==========================================================================

'******************************************************************************
' Nome do elemento: Connect (Function)
' Resumo: Estabelece uma sessão SAP GUI com scripting habilitado.
' @param sysDescription — Descrição do sistema no SAP Logon Pad (coluna Description), ex: "PEP" ou "S48 – S/4 DEV".
' @param sysClient — Número do cliente a ser digitado durante o logon.
' @param sysSession — Índice da sessão desejada (0-based, padrão = 0).
' @return Boolean — True se a sessão estiver pronta para uso; False em caso de falha (exibe MsgBox).
' @raises vbObjectError+1 — scripting não disponível no workstation.
' @raises vbObjectError+2 — SAP Logon não apareceu dentro do tempo limite.
' @raises vbObjectError+3 — Descrição do sistema não encontrada no SAP Logon Pad.
' @remarks • Inicia saplogon.exe se não estiver em execução.
'          • Reutiliza conexão existente ou abre nova via OpenConnection.
'          • Garante existência da sessão desejada e executa logon interativo se estiver na tela S000.
'******************************************************************************

Public Function Connect(sysDescription As String, _
                        ByVal sysClient As Long, _
                        Optional ByVal sysSession As Long = 0) As Boolean
    On Error GoTo ErrHandler

    Application.StatusBar = "Verificando SAP Logon…"
    EnsureSapLogonRunning                                ' start saplogon.exe if needed

    Set cSapGuiApp = GetObject("SAPGUI").GetScriptingEngine
    If cSapGuiApp Is Nothing Then Err.Raise vbObjectError + 1, , _
        "ScriptingEngine não disponível – verifique perfil Z:BC_PB001_USUÁRIO_EXEC_SCRIPT"

    Set cConnection = PickOrOpenConnection(sysDescription)
    If cConnection Is Nothing Then GoTo CleanFail

    Set cSession = EnsureWantedSession(cConnection, sysClient, sysSession)
    If cSession Is Nothing Then GoTo CleanFail
    
    On Error Resume Next
    If cSession.findbyid("wnd[1]").Text = "Mensagens de sistema" Then cSession.findbyid("wnd[1]/tbar[0]/btn[12]").press
    If cSession.Info.transaction = "S000" Then cSession.findbyid("wnd[0]/usr/btnSTARTBUTTON").press
    Err.Clear
    On Error GoTo ErrHandler
    If cSession.Info.transaction = "S000" Then PerformInteractiveLogon cSession, sysClient
    Connect = True: Exit Function

CleanFail:
    MsgBox "Não foi possível criar a conexão com o SAP.", vbCritical
    Exit Function

ErrHandler:
    MsgBox Err.Description, vbCritical
End Function

'******************************************************************************
' Nome do elemento: Session (Property Get)
' Resumo: Obtem o valor da propriedade Session.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Session() As Object
    Set Session = cSession
End Property

'******************************************************************************
' Nome do elemento: EnsureSapLogonRunning (Sub)
' Resumo: Executa a rotina EnsureSapLogonRunning.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Sub EnsureSapLogonRunning()
    Const SAP_EXE As String = "C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"
    Dim hwnd As Long, i As Long, appId As Long

    hwnd = FindWindow(vbNullString, "SAP Logon 800")
    If hwnd = 0 Then
        appId = Shell(SAP_EXE, vbNormalFocus)
        For i = 1 To 60                                ' 60 s timeout
            Sleep32 1000
            hwnd = FindWindow(vbNullString, "SAP Logon 800")
            If hwnd <> 0 Then Exit For
            DoEvents
        Next
        If hwnd = 0 Then Err.Raise vbObjectError + 2, , "Timeout iniciando SAP Logon."
    End If
    DeActivateWindow hwnd
End Sub

'******************************************************************************
' Nome do elemento: PickOrOpenConnection (Function)
' Resumo: Retorna o resultado calculado por PickOrOpenConnection.
' @param sysDescription - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Function PickOrOpenConnection(sysDescription As String) As Object
    Dim c As Object

    ' A) Reutiliza conexão existente
    For Each c In cSapGuiApp.Children
        If c.Children.Count > 0 Then Set PickOrOpenConnection = c: Exit Function
    Next

    ' B) Abre nova conexão
    Set PickOrOpenConnection = cSapGuiApp.OpenConnection(sysDescription, True)
    If PickOrOpenConnection Is Nothing Then
        Err.Raise vbObjectError + 3, , _
          "Sistema '" & sysDescription & "' não encontrado no SAP Logon Pad."
    End If
End Function

'******************************************************************************
' Nome do elemento: EnsureWantedSession (Function)
' Resumo: Retorna o resultado calculado por EnsureWantedSession.
' @param conn - Parâmetro da rotina.
' @param sysClient - Parâmetro da rotina.
' @param sysSession - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function EnsureWantedSession(conn As Object, sysClient As Long, sysSession As Long) As Object
    Dim s As Object, prevCnt As Long

    Do
        ' 1) Sessão já existe?
        Set s = Nothing
        On Error Resume Next
        Set s = conn.Sessions(CInt(sysSession))
        On Error GoTo 0
        If Not s Is Nothing Then Exit Do

        ' 2) Se única sessão em logon, interage
        If conn.Sessions.Count = 1 Then
            If conn.Sessions(0).Info.transaction = "S000" Then
                PerformInteractiveLogon conn.Sessions(0), sysClient
            End If
        End If

        ' 3) Cria nova sessão e espera a janela aparecer
        prevCnt = conn.Sessions.Count
        conn.Sessions(0).CreateSession

        Dim t0 As Double: t0 = Timer
        Do
            DoEvents
            If conn.Sessions.Count > prevCnt Then Exit Do
            If Timer - t0 > 10 Then Exit Do  ' 10-s timeout
            Sleep32 200
        Loop

        ' 4) Aguarda até nova sessão não estar Busy
        Set s = conn.Sessions(conn.Sessions.Count - 1)
        WaitUntil s, "Busy", False, 10
    Loop

    Set EnsureWantedSession = s
End Function

'******************************************************************************
' Nome do elemento: WaitUntil (Sub)
' Resumo: Polling genérico: espera até que propName de obj seja igual a wantedValue ou timeout.
' @param obj — Objeto a ser verificado.
' @param propName — Nome da propriedade a ler.
' @param wantedValue — Valor desejado para comparação.
' @param seconds — Limite de tempo em segundos (opcional, default = 30).
'******************************************************************************
Private Sub WaitUntil(obj As Object, propName As String, wantedValue As Variant, _
                       Optional seconds As Long = 30)
    Dim t As Double: t = Timer
    Do
        DoEvents
        If CallByName(obj, propName, VbGet) = wantedValue Then Exit Do
        If Timer - t > seconds Then Exit Do
        Sleep32 200
    Loop
End Sub

'******************************************************************************
' Nome do elemento: PerformInteractiveLogon (Sub)
' Resumo: Executa a rotina PerformInteractiveLogon.
' @param s - Parâmetro da rotina.
' @param Client - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub PerformInteractiveLogon(s As Object, Client As Long)
    With s
        .findbyid("wnd[0]/usr/txtRSYST-MANDT").Text = Client
        .findbyid("wnd[0]/usr/txtRSYST-BNAME").Text = ""
        .findbyid("wnd[0]/usr/pwdRSYST-BCODE").Text = ""
        .sendVKey 0
    End With

    Do
        DoEvents
        Sleep32 1000
    Loop Until s.ActiveWindow.Text Like "SAP Easy Access*"
End Sub




