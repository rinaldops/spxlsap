Attribute VB_Name = "fwSapGuiApi"
'******************************************************************************
' Módulo: fwSapGuiApi
' Descrição:
'   Centraliza todas as declarações Win32 e pequenos helpers necessários à
'   biblioteca de automação SAP. Permite que todas as classes e módulos
'   reutilizem APIs e wrappers sem duplicação de código ou conflitos de
'   compilação.
' Autor: Rinaldo Paulino de Souza
' Data da documentação: 09/07/2025
'*******************************************************************************

'******************************************************************************
' Nome do elemento: fwSapGuiApi (Módulo)
' Resumo:
'   Reúne rotinas deste módulo no padrão documental fwXLTable.
'******************************************************************************

Option Explicit

'******************************************************************************
' Seção: Win32 API Declarations
' Descrição:
'   Declarações públicas de funções e tipos Win32 utilizadas para
'   gerenciamento de janelas, privilégios e memória.
' Observações:
'   Todas as declarações são PtrSafe e compatíveis com VBA 7+ (64 bits).
'******************************************************************************

' Suspende a thread atual por um número de milissegundos
' @param dwMilliseconds As Long — Tempo em milissegundos para Sleep.
Public Declare PtrSafe Sub Sleep32 Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

' Ajusta privilégios no token de acesso (habilita debug, etc.)
Public Declare PtrSafe Function AdjustTokenPrivileges Lib "advapi32.dll" ( _
    ByVal TokenHandle As Long, _
    ByVal DisableAllPrivileges As Long, _
    NewState As TOKEN_PRIVILEGES, _
    ByVal BufferLength As Long, _
    PreviousState As TOKEN_PRIVILEGES, _
    ReturnLength As Long _
) As Long

' Traz janela para o topo da Z-order
Public Declare PtrSafe Function BringWindowToTop Lib "User32" ( _
    ByVal hwnd As Long _
) As Long

' Fecha handle de objeto
Public Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As LongPtr _
) As Long

' Localiza janela pelo nome da classe ou título
Public Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As Long

' Obtém handle da janela em primeiro plano
Public Declare PtrSafe Function GetForegroundWindow Lib "User32" () As Long

' Navega na hierarquia de janelas
Public Declare PtrSafe Function GetWindow Lib "User32" ( _
    ByVal hwnd As Long, _
    ByVal wFlag As Long _
) As Long

' Lê texto do título da janela
Public Declare PtrSafe Function GetWindowText Lib "User32" Alias "GetWindowTextA" ( _
    ByVal hwnd As LongLong, _
    ByVal lpString As String, _
    ByVal cch As Long _
) As Long

' Obtém nome da classe de janela
Public Declare PtrSafe Function GetClassName Lib "User32" Alias "GetClassNameA" ( _
    ByVal hwnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long _
) As Long

' Recebe ID de processo associado à janela
Public Declare PtrSafe Function GetWindowThreadProcessId Lib "User32" ( _
    ByVal hwnd As LongPtr, _
    lpdwProcessId As Long _
) As Long

' Retorna handle do processo atual
Public Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As Long

' Obtém dados de posicionamento e estado da janela
Public Declare PtrSafe Function GetWindowPlacement Lib "User32" ( _
    ByVal hwnd As Long, _
    lpwndpl As WINDOWPLACEMENT _
) As Long

' Define dados de posicionamento e estado da janela
Public Declare PtrSafe Function SetWindowPlacement Lib "User32" ( _
    ByVal hwnd As Long, _
    lpwndpl As WINDOWPLACEMENT _
) As Long

' Posta mensagem para a fila de mensagens da janela
Public Declare PtrSafe Function PostMessage Lib "User32" Alias "PostMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any _
) As Long

' Envia mensagem síncrona para a janela
Public Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any _
) As Long

' Concede permissão de término de processo
Public Declare PtrSafe Function OpenProcessToken Lib "advapi32.dll" ( _
    ByVal ProcessHandle As Long, _
    ByVal DesiredAccess As Long, _
    TokenHandle As Long _
) As Long

' Obtém LUID de privilégio
Public Declare PtrSafe Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" ( _
    ByVal lpSystemName As String, _
    ByVal lpName As String, _
    lpLuid As LUID _
) As Long

' Abre processo com direitos específicos
Public Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long _
) As Long

' Termina o processo especificado
Public Declare PtrSafe Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As Long, _
    ByVal uExitCode As Long _
) As Long

' Lê linha de comando Unicode
Public Declare PtrSafe Function w_commandline Lib "kernel32.dll" Alias "GetCommandLineW" () As LongPtr

' Obtém comprimento de string Unicode
Public Declare PtrSafe Function w_strlen Lib "kernel32.dll" Alias "lstrlenW" ( _
    ByVal lpString As LongPtr _
) As Long

' Copia bloco de memória (Unicode)
Public Declare PtrSafe Sub w_memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    dst As Any, _
    src As Any, _
    ByVal size As LongPtr _
)

' ––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––– '
' Types retornados/aceitos por GetWindowPlacement / SetWindowPlacement      '
' ––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––– '
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Public Type WINDOWPLACEMENT
    Length           As Long
    flags            As Long
    showCmd          As Long
    ptMinPosition    As POINTAPI
    ptMaxPosition    As POINTAPI
    rcNormalPosition As RECT
End Type

Private Type LUID
    LowPart  As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid      As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid        As LUID
    Attributes     As Long
End Type

'******************************************************************************
' Seção: Constants
' Descrição:
'   Flags e mensagens usadas nas chamadas Win32.
'*******************************************************************************
Public Const SW_SHOWNORMAL    As Long = 1
Public Const SW_SHOWMINIMIZED As Long = 2
Public Const WM_CLOSE         As Long = &H10
Public Const WM_SETTEXT       As Long = &HC
Public Const WM_SYSKEYDOWN    As Long = &H104
Public Const WM_RBUTTONDOWN   As Long = &H204
Public Const WS_MAXIMIZE      As Long = &H1000000
Public Const WM_CHAR          As Long = &H102
Public Const BM_CLICK         As Long = &HF5&
Public Const GW_HWNDNEXT      As Long = 2&
Public Const GW_CHILD         As Long = 5&
Public Const VK_RETURN        As Long = &HD

'******************************************************************************
' Seção: Variáveis globais
' Descrição:
'   Variáveis de conexão com o SAP GUI. Instanciamento feito na Sub/Function
'*******************************************************************************
'Público sap As fwSAPConn
'Público Session As Object
'Público w As fwGuiMainWindow

'*******************************************************************************
' @description  Espera em ms, mantendo a UI viva
' @param ms     Tempo em milissegundos
'*******************************************************************************

'******************************************************************************
' Nome do elemento: WaitTime (Sub)
' Resumo: Executa a rotina WaitTime.
' @param ms - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub WaitTime(ByVal ms As Long)
    Dim start As Long, now As Long
    start = GetTickCount
    Do
        DoEvents           ' processa mensagens da UI
        Sleep32 50         ' pausa em pequenos blocos (ajuste se quiser)
        now = GetTickCount
    Loop While now - start < ms
End Sub

'*******************************************************************************
' @description  Minimises the window `xhWnd` regardless of its current
'               state.  Opposite of `ActivateWindow`
' @param xhWnd  Handle (hWnd) of the window to hide
' @returns      True  – window successfully minimised
'               False – inválido handle or API call failed
'*******************************************************************************

'******************************************************************************
' Nome do elemento: DeActivateWindow (Function)
' Resumo: Retorna o resultado calculado por DeActivateWindow.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function DeActivateWindow(xhWnd&) As Boolean
    Dim WndPlcmt As WINDOWPLACEMENT
    WndPlcmt.Length = Len(WndPlcmt)
    If GetWindowPlacement(xhWnd, WndPlcmt) Then
        WndPlcmt.flags = 0: WndPlcmt.showCmd = SW_SHOWMINIMIZED
        DeActivateWindow = CBool(SetWindowPlacement(xhWnd, WndPlcmt))
    End If
End Function

'******************************************************************************
' Nome do elemento: ProcessTerminate (Function)
' Resumo: Retorna o resultado calculado por ProcessTerminate.
' @param lProcessID - Parâmetro da rotina.
' @param lHwndWindow - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ProcessTerminate(Optional lProcessID As Long, Optional lHwndWindow As LongPtr) As Boolean
    Dim lhwndProcess As Long
    Dim lExitCode As Long
    Dim lRetVal As Long
    Dim lhThisProc As Long
    Dim lhTokenHandle As Long
    Dim tLuid As LUID
    Dim tTokenPriv As TOKEN_PRIVILEGES, tTokenPrivNew As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    
    Const PROCESS_ALL_ACCESS = &H1F0FFF, PROCESS_TERMINATE = &H1
    Const ANYSIZE_ARRAY = 1, TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8, SE_DEBUG_NAME As String = "SeDebugPrivilege"
    Const SE_PRIVILEGE_ENABLED = &H2

    On Error Resume Next
    If lHwndWindow Then
        'Get the processo ID from the window handle
        lRetVal = GetWindowThreadProcessId(lHwndWindow, lProcessID)
    End If
    
    If lProcessID Then
        'Give Kill permissions to this processo
        lhThisProc = GetCurrentProcess
        
        OpenProcessToken lhThisProc, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lhTokenHandle
        LookupPrivilegeValue "", SE_DEBUG_NAME, tLuid
        'Set the number of privileges to be change
        tTokenPriv.PrivilegeCount = 1
        tTokenPriv.TheLuid = tLuid
        tTokenPriv.Attributes = SE_PRIVILEGE_ENABLED
        'Enable the kill privilege in the access token of this processo
        AdjustTokenPrivileges lhTokenHandle, False, tTokenPriv, Len(tTokenPrivNew), tTokenPrivNew, lBufferNeeded

        'Open the processo to kill
        lhwndProcess = OpenProcess(PROCESS_TERMINATE, 0, lProcessID)
    
        If lhwndProcess Then
            'Obtained processo handle, kill the processo
            ProcessTerminate = CBool(TerminateProcess(lhwndProcess, lExitCode))
            Call CloseHandle(lhwndProcess)
        End If
    End If
    On Error GoTo 0
End Function

'******************************************************************************
' Nome do elemento: GetCommandLine (Function)
' Resumo: Retorna o resultado calculado por GetCommandLine.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetCommandLine() As String
    GetCommandLine = String$(w_strlen(w_commandline()), 0)
    w_memcpy ByVal StrPtr(GetCommandLine), ByVal w_commandline(), LenB(GetCommandLine)
End Function

'******************************************************************************
' Nome do elemento: initSAP (Function)
' Resumo: Retorna o resultado calculado por initSAP.
' @param Nome - Parâmetro da rotina.
' @param SID - Parâmetro da rotina.
' @param Sessão - Parâmetro da rotina.
' @param w - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function initSAP(ByVal Nome As String, ByVal SID As String, ByVal Sessao As String, ByRef w As fwGuiMainWindow) As Boolean
    Dim sap As fwSAPConn
    On Error GoTo Cleanup
    If sap Is Nothing Then Set sap = New fwSAPConn
    If Not sap.Connect(Nome, SID, Sessao) Then GoTo Cleanup
    If w Is Nothing Then
        Set w = New fwGuiMainWindow: w.Init sap.Session, "wnd[0]"
    End If
    initSAP = True
    Exit Function
Cleanup:
    initSAP = False
End Function

'******************************************************************************
' Nome do elemento: AtualizaBtnSAP (Sub)
' Resumo: Executa a rotina AtualizaBtnSAP.
' @param w - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Sub AtualizaBtnSAP(ByRef w As fwGuiMainWindow)
    If Not w Is Nothing Then
        If w.IsSessionAlive Then
            pVariaveis.Range("imgConexao").value = "AppointmentColor3"
        Else
            If w.IsSessionAlive Then
                pVariaveis.Range("imgConexao").value = "AppointmentColor3"
            Else
                pVariaveis.Range("imgConexao").value = "AppointmentColor1"
            End If
        End If
    Else
        pVariaveis.Range("imgConexao").value = "AppointmentColor1"
    End If
    If Not GetRibbon Is Nothing Then GetRibbon.InvalidateControl "btnConectar"
End Sub

