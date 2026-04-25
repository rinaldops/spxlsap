VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fwGuiMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' Nome do elemento: fwGuiMainWindow (Classe)
' Resumo: Encapsula uma janela SAP GUI (wnd[n]) e fornece método para navegação,
'         busca de controle e execução de transações.
' @remarks: Varre recursivamente a hierarquia de elemento, instanciando wrappers
'           para GuiTableControl e GuiTree.
'******************************************************************************

Option Explicit

Private pSession        As Object           ' GuiSession
Private pWnd            As Object           ' GuiMainWindow (ou pop-up)
Private pWndId          As String           ' ex: "wnd[0]", "wnd[1]"
Private pGuiTable       As fwGuiTableControl ' wrapper único de GuiTableControl
Private pGuiTree        As fwGuiTree         ' wrapper único de GuiTree
Private pManualOn      As Boolean       ' flag do modo manual
Private pManualStep    As Long          ' token de passo manual
Private pLastObjId     As String        ' ultimo objeto marcado
Private pLastIsButton  As Boolean       ' indica se ultimo era botão
Private pTraceOn       As Boolean       ' flag de rastreamento

'******************************************************************************
' Nome do elemento: ManualOn (Property Get)
' Resumo: Obtem o valor da propriedade ManualOn.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get ManualOn() As Boolean
    ManualOn = pManualOn
End Property

'******************************************************************************
' Nome do elemento: ManualOn (Property Let)
' Resumo: Define o valor da propriedade ManualOn.
' @param value - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Let ManualOn(ByVal value As Boolean)
    pManualOn = value
End Property

'******************************************************************************
' Nome do elemento: TraceOn (Property Let)
' Resumo: Define o valor da propriedade TraceOn.
' @param v - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Let TraceOn(ByVal v As Boolean)
    pTraceOn = CBool(v)
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
' Nome do elemento: ManualNext (Sub)
' Resumo: Executa a rotina ManualNext.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub ManualNext()
    Trace "ManualNext"
    pManualStep = pManualStep + 1
End Sub

'******************************************************************************
' Nome do elemento: MarkLastID (Sub)
' Resumo: Executa a rotina MarkLastID.
' @param ctrl - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub MarkLastID(ByVal ctrl As Object)
    If ctrl Is Nothing Then
        Trace "MarkLastID: ctrl is Nothing"
        Exit Sub
    End If
    On Error Resume Next
    pLastObjId = ctrl.ID
    pLastIsButton = IsButtonType(ctrl)
    Trace "MarkLastID: id='" & pLastObjId & "'"
    On Error GoTo 0
End Sub

'******************************************************************************
' Nome do elemento: ManualPause (Sub)
' Resumo: Executa a rotina ManualPause.
' @param ctrl - Parâmetro da rotina.
' @param hint - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub ManualPause(ByVal ctrl As Object, Optional ByVal hint As String = "")
    If Not pManualOn Then Exit Sub
    If ctrl Is Nothing Then Exit Sub

    Trace "ManualPause: hint='" & hint & "'"
    Debug.Print ctrl.ID
    
    On Error Resume Next
    If InStr(1, ctrl.Type, "GuiMainWindow", vbTextCompare) > 0 Then Exit Sub
    On Error GoTo 0
    
    Dim currId As String
    Dim currIsButton As Boolean
    On Error Resume Next
    currId = ctrl.ID
    currIsButton = IsButtonType(ctrl) Or Right(currId, 4) = "okcd"
    On Error GoTo 0
    
    ' Se o objeto anterior for "okcd" anula marcação do anterior e ativa o do objeto atual
    Dim markPrev As Boolean, markCurr As Boolean
    If pLastObjId = "wnd[0]/tbar[0]/okcd" Then
        markCurr = True
        pLastObjId = ""
    End If
    Dim prev As Object
    If LenB(pLastObjId) > 0 Then Set prev = SafeFindById(pLastObjId)
    
    If pLastIsButton And currIsButton Then
        markCurr = True
    ElseIf (Not pLastIsButton) And currIsButton Then
        markPrev = Not prev Is Nothing
        markCurr = True
    ElseIf (Not pLastIsButton) And (Not currIsButton) Then
        markPrev = Not prev Is Nothing
    ElseIf pLastIsButton And (Not currIsButton) Then
        markCurr = True
    End If
    
    Dim token As Long: token = pManualStep
    On Error Resume Next
    If markPrev Then prev.Visualize True
    If markCurr Then ctrl.Visualize True
    On Error GoTo 0

    Do While pManualOn And pManualStep = token
        DoEvents
    Loop

    On Error Resume Next
    If markPrev Then prev.Visualize False
    If markCurr Then ctrl.Visualize False
    On Error GoTo 0

    pLastObjId = currId
    pLastIsButton = currIsButton
End Sub

'******************************************************************************
' Nome do elemento: IsButtonType (Function)
' Resumo: Retorna o resultado calculado por IsButtonType.
' @param ctrl - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function IsButtonType(ByVal ctrl As Object) As Boolean
    On Error Resume Next
    Dim t As String: t = ctrl.Type
    If InStr(1, t, "GuiButton", vbTextCompare) > 0 Then IsButtonType = True
    If InStr(1, t, "ToolbarButton", vbTextCompare) > 0 Then IsButtonType = True
    If InStr(1, t, "GuiToolbar", vbTextCompare) > 0 Then IsButtonType = True
    If IsButtonType = False Then
        Dim cid As String: cid = ctrl.ID
        If LCase$(Left$(cid, 3)) = "btn" Then IsButtonType = True
    End If
    On Error GoTo 0
End Function

'******************************************************************************
' Nome do elemento: Init (Sub)
' Resumo: Inicializa o wrapper de janela para um identificador de janela específico.
' @param sessionObj - Objeto de sessão SAP GUI (GuiSession).
' @param WindowId - Identificador da janela, ex: "wnd[0]" (default).
' @raises Erro vbObjectError+513 se a janela não for encontrada.
'******************************************************************************
Public Sub Init( _
    ByVal sessionObj As Object, _
    Optional ByVal windowId As String = "wnd[0]" _
)
    Trace "Init: windowId='" & windowId & "'"
    Set pSession = sessionObj
    pWndId = windowId

    On Error Resume Next
    Set pWnd = pSession.findbyid(windowId)
    On Error GoTo 0
    If pWnd Is Nothing Then
        Trace "Init: window not found"
        
        Err.Raise vbObjectError + 513, _
                  "cWindow.Init", _
                  "Janela '" & windowId & "' não encontrada."
    End If
End Sub

'******************************************************************************
' Nome do elemento: windowId (Property Get)
' Resumo: Obtem o valor da propriedade windowId.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get windowId() As String
    windowId = pWndId
End Property

'******************************************************************************
' Nome do elemento: ListWindowIds (Function)
' Resumo: Retorna o resultado calculado por ListWindowIds.
' @return Collection - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ListWindowIds() As Collection
    Trace "ListWindowIds: scanning"
    Dim lst As New Collection
    Dim idx As Long, wId As String
    On Error Resume Next
    For idx = 0 To 20
        wId = "wnd[" & idx & "]"
        pSession.findbyid wId
        If Err.Number = 0 Then lst.Add wId Else Err.Clear
    Next
    On Error GoTo 0
    Trace "ListWindowIds: found=" & lst.Count
    Set ListWindowIds = lst
End Function

'******************************************************************************
' Nome do elemento: Enter (Sub)
' Resumo: Executa a rotina Enter.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub Enter()
    Trace "Enter: sendVKey 0"
    On Error Resume Next
    pSession.findbyid(Right(pSession.ActiveWindow.ID, 6)).sendVKey 0
    On Error GoTo 0
End Sub

'******************************************************************************
' Nome do elemento: Self (Property Get)
' Resumo: Obtem o valor da propriedade Self.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Self() As Object
    Set Self = pWnd
End Property

'******************************************************************************
' Nome do elemento: TCode (Sub)
' Resumo: Executa a rotina TCode.
' @param sTCode - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Sub TCode(ByVal sTCode As String)
    Trace "TCode: '" & sTCode & "'"
    On Error GoTo Err_Handler

    If pSession Is Nothing Then
        Err.Raise vbObjectError + 1000, "cWindow.TCode", "Sessão não inicializada."
    End If

    Dim okField As Object

    If Len(Trim$(sTCode)) = 0 Then
        pSession.endTransaction
    Else
        If pManualOn Then
            Set okField = SafeFindById("wnd[0]/tbar[0]/okcd")
            okField.Text = "/n" & sTCode
            ManualPause okField, sTCode
        End If
        pSession.StartTransaction sTCode
        pLastObjId = "wnd[0]/tbar[0]/okcd"
    End If

    Do While pSession.Busy
        DoEvents
    Loop
    
    Set okField = SafeFindById("wnd[0]/tbar[0]/okcd")
    If okField.Text <> "" Then
        okField.Visualize = True
    End If
    
    Exit Sub

Err_Handler:
    Err.Raise Err.Number, Err.source, "cWindow.TCode('" & sTCode & "'): " & Err.Description
End Sub


'******************************************************************************
' Nome do elemento: FindObj (Function)
' Resumo: Encontra um objeto SAP GUI dentro de uma janela usando prefixos técnicos ou FindByName.
' @param sWindowId - Identificador da janela (ex: "wnd[0]" ou prefixo técnico).
' @param sTechName - Nome técnico do controle (ex: "btnOK").
' @param sSearchText - Texto a ser pesquisado na propriedade Text do controle.
' @return Object - Objeto SAP GUI encontrado ou Nothing se não existir.
'******************************************************************************
Public Function FindObj( _
    ByVal sWindowId As String, _
    Optional ByVal sIdOrPrefName As String = vbNullString, _
    Optional ByVal sText As String = vbNullString, _
    Optional ByVal sParentText As String = vbNullString, _
    Optional ByVal sSubstrId As String = vbNullString, _
    Optional ByVal sSubType As String = vbNullString, _
    Optional ByVal sTitle As String = vbNullString, _
    Optional ByVal sTooltip As String = vbNullString _
) As Object

    Dim targetWin      As Object
    Dim targetObj      As Object
    Dim matches        As Object
    Dim candidate      As Object
    Dim resolvedName   As String
    Dim resolvedType   As String
    Dim idx            As Long
    Dim directSearch   As Boolean
    Dim resultObj      As Object
    Static depth       As Long
    Dim hasFilters     As Boolean
    Dim allowDeep      As Boolean

    depth = depth + 1
    On Error GoTo CleanExit
    Trace "FindObj ENTER: depth=" & depth & " windowId='" & sWindowId & "' idOrName='" & sIdOrPrefName & "' text='" & sText & "' parentText='" & sParentText & "' substrId='" & sSubstrId & "' subType='" & sSubType & "' title='" & sTitle & "' tooltip='" & sTooltip & "'"
    hasFilters = (LenB(sText & sParentText & sSubstrId & sSubType & sTitle & sTooltip) > 0)
    directSearch = Not hasFilters
    NormalizeWindowArgs sWindowId, sIdOrPrefName

    ' Branch para menu (usa texto do item e do pai)
    If LenB(sParentText) > 0 Then
        Trace "FindObj: menu branch"
        Set candidate = FindMenuByText(sWindowId, sText, sParentText)
        If CandidateMatches(candidate, sSubType, vbNullString, sSubstrId, sTitle, sTooltip) Then
            Set resultObj = candidate
        End If
        GoTo CleanExit
    End If

    If LenB(sIdOrPrefName) = 0 Then GoTo CleanExit

    If LooksLikeFullId(sIdOrPrefName) Then
        Trace "FindObj: LooksLikeFullId"
        Set candidate = SafeFindById(sIdOrPrefName)
        If CandidateMatches(candidate, sSubType, sText, sSubstrId, sTitle, sTooltip) Then
            Set resultObj = candidate
        End If
        GoTo CleanExit
    End If

    Set targetObj = SafeFindById(sIdOrPrefName)
    If Not targetObj Is Nothing Then
        Trace "FindObj: SafeFindById matched direct id"
        If directSearch Then
            Set resultObj = targetObj
            GoTo CleanExit
        ElseIf CandidateMatches(targetObj, sSubType, sText, sSubstrId, sTitle, sTooltip) Then
            Set resultObj = targetObj
            GoTo CleanExit
        End If
    End If

    If IsWindowToken(sIdOrPrefName) Then
        Set candidate = SafeFindById(sIdOrPrefName)
        If CandidateMatches(candidate, sSubType, sText, sSubstrId, sTitle, sTooltip) Then
            Set resultObj = candidate
        End If
        GoTo CleanExit
    End If

    If Not ResolveTypeAndName(sIdOrPrefName, resolvedName, resolvedType) Then GoTo CleanExit

    allowDeep = IsDeepSearchType(resolvedType, sIdOrPrefName)
    Trace "FindObj: allowDeep=" & allowDeep

    Set targetWin = SafeFindById(sWindowId)
    If targetWin Is Nothing Then GoTo CleanExit

    ' Tenta encontrar objeto na janela principal por FindAllByName
    Trace "FindObj: FindAllByName in targetWin name='" & resolvedName & "' type='" & resolvedType & "'"
    Set matches = SafeFindAllByName(targetWin, resolvedName, resolvedType)
    If Not matches Is Nothing Then
        If matches.Count > 0 Then
            For idx = 0 To SafeCollectionCount(matches) - 1
                Set candidate = SafeCollectionItem(matches, idx)
                If CandidateMatches(candidate, sSubType, sText, sSubstrId, sTitle, sTooltip) Then
                    Set resultObj = candidate
                    GoTo CleanExit
                End If
            Next idx
        End If
    End If
    
    ' Tenta encontrar objeto dentro do container por FindAllByName
    If allowDeep Then
        Dim i As Long, subTargetWin As Object
        For i = 0 To targetWin.Children.Count - 1
            If targetWin.Children(i - 0).Name = "shellcont" Then
                Set subTargetWin = targetWin.Children(i - 0)
                Exit For
            End If
        Next
        Trace "FindObj: FindAllByName in shellcont"
        Set matches = SafeFindAllByName(subTargetWin, resolvedName, resolvedType)
        If Not matches Is Nothing Then
            If matches.Count > 0 Then
                For idx = 0 To SafeCollectionCount(matches) - 1
                    Set candidate = SafeCollectionItem(matches, idx)
                    If CandidateMatches(candidate, sSubType, sText, sSubstrId, sTitle, sTooltip) Then
                        Set resultObj = candidate
                        GoTo CleanExit
                    End If
                Next idx
            End If
        End If
    End If
    
    ' Se não encontrar, tenta pelo FindByName
    Trace "FindObj: FindByName fallback"
    Set candidate = SafeFindByName(targetWin, resolvedName, resolvedType)
    If CandidateMatches(candidate, sSubType, sText, sSubstrId, sTitle, sTooltip) Then
        Set resultObj = candidate
        GoTo CleanExit
    End If

    If Not allowDeep Then GoTo CleanExit

    Trace "FindObj: recursive fallback"
    Set candidate = FindMatchRecursive(targetWin, resolvedName, resolvedType, sSubType, sText, sSubstrId, sTitle, sTooltip)
    If Not candidate Is Nothing Then
        Set resultObj = candidate
    End If

CleanExit:
    If pManualOn And depth = 1 Then ManualPause resultObj, sIdOrPrefName
    Trace "FindObj EXIT: depth=" & depth & " result=" & IIf(resultObj Is Nothing, "Nothing", TryGetStringProperty(resultObj, "Id"))
    Set FindObj = resultObj
    If depth > 0 Then depth = depth - 1
End Function

'******************************************************************************
' Nome do elemento: NormalizeWindowArgs (Sub)
' Resumo: Executa a rotina NormalizeWindowArgs.
' @param sWindowId - Parâmetro da rotina.
' @param sIdOrPrefName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub NormalizeWindowArgs(ByRef sWindowId As String, ByRef sIdOrPrefName As String)
    Dim origWin As String: origWin = sWindowId
    Dim origId As String: origId = sIdOrPrefName

    If LenB(sIdOrPrefName) = 0 Then
        sIdOrPrefName = sWindowId
        sWindowId = Me.windowId
        Trace "NormalizeWindowArgs: swapped windowId/id"
        Exit Sub
    End If

    If Not IsWindowToken(sWindowId) Then
        sIdOrPrefName = sWindowId
        sWindowId = Me.windowId
        Trace "NormalizeWindowArgs: normalized windowId from '" & origWin & "' to '" & sWindowId & "' idOrName from '" & origId & "' to '" & sIdOrPrefName & "'"
    End If
End Sub

'******************************************************************************
' Nome do elemento: IsWindowToken (Function)
' Resumo: Retorna o resultado calculado por IsWindowToken.
' @param candidateId - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function IsWindowToken(ByVal candidateId As String) As Boolean
    If Len(candidateId) >= 6 Then
        If LCase$(Left$(candidateId, 4)) = "wnd[" And Right$(candidateId, 1) = "]" Then
            IsWindowToken = True
        End If
    End If
End Function

'******************************************************************************
' Nome do elemento: LooksLikeFullId (Function)
' Resumo: Retorna o resultado calculado por LooksLikeFullId.
' @param candidateId - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function LooksLikeFullId(ByVal candidateId As String) As Boolean
    If InStr(1, candidateId, "/", vbTextCompare) > 0 Then
        If InStr(1, candidateId, "wnd[", vbTextCompare) > 0 Then
            LooksLikeFullId = True
        End If
    End If
End Function

'******************************************************************************
' Nome do elemento: IsDeepSearchType (Function)
' Resumo: Retorna o resultado calculado por IsDeepSearchType.
' @param resolvedType - Parâmetro da rotina.
' @param sIdOrPrefName - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function IsDeepSearchType(ByVal resolvedType As String, _
                                  ByVal sIdOrPrefName As String) As Boolean
    Dim idLower As String
    Dim typeLower As String

    idLower = LCase$(sIdOrPrefName)
    typeLower = LCase$(resolvedType)

    If typeLower = "guishell" Or typeLower = "guimenu" Then
        IsDeepSearchType = True
        Exit Function
    End If

    If Left$(idLower, 5) = "shell" Or Left$(idLower, 4) = "menu" Then
        IsDeepSearchType = True
    End If
End Function

'******************************************************************************
' Nome do elemento: SafeFindById (Function)
' Resumo: Retorna o resultado calculado por SafeFindById.
' @param fullId - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function SafeFindById(ByVal fullId As String) As Object
    On Error Resume Next
    Set SafeFindById = pSession.findbyid(fullId)
    If Err.Number = 0 Then
        Trace "SafeFindById: found '" & fullId & "'"
    Else
        Trace "SafeFindById: not found '" & fullId & "'"
        Err.Clear
    End If
    On Error GoTo 0
End Function

'******************************************************************************
' Nome do elemento: SafeFindByName (Function)
' Resumo: Retorna o resultado calculado por SafeFindByName.
' @param container - Parâmetro da rotina.
' @param compName - Parâmetro da rotina.
' @param compType - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function SafeFindByName(ByVal container As Object, _
                                ByVal compName As String, _
                                ByVal compType As String) As Object
    On Error Resume Next
    Set SafeFindByName = container.findByName(compName, compType)
    If Err.Number = 0 Then
        Trace "SafeFindByName: name='" & compName & "' type='" & compType & "'"
    Else
        Trace "SafeFindByName: not found name='" & compName & "' type='" & compType & "'"
        Err.Clear
    End If
    On Error GoTo 0
End Function

'******************************************************************************
' Nome do elemento: SafeFindAllByName (Function)
' Resumo: Retorna o resultado calculado por SafeFindAllByName.
' @param container - Parâmetro da rotina.
' @param compName - Parâmetro da rotina.
' @param compType - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function SafeFindAllByName(ByVal container As Object, _
                                   ByVal compName As String, _
                                   ByVal compType As String) As Object
    On Error Resume Next
    Set SafeFindAllByName = container.FindAllByName(compName, compType)
    If Err.Number = 0 Then
        Trace "SafeFindAllByName: name='" & compName & "' type='" & compType & "'"
    Else
        Trace "SafeFindAllByName: not found name='" & compName & "' type='" & compType & "'"
        Err.Clear
    End If
    On Error GoTo 0
End Function

'******************************************************************************
' Nome do elemento: SafeCollectionCount (Function)
' Resumo: Retorna o resultado calculado por SafeCollectionCount.
' @param collectionObj - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function SafeCollectionCount(ByVal collectionObj As Object) As Long
    On Error Resume Next
    SafeCollectionCount = collectionObj.Count
    If Err.Number <> 0 Then
        SafeCollectionCount = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

'******************************************************************************
' Nome do elemento: SafeCollectionItem (Function)
' Resumo: Retorna o resultado calculado por SafeCollectionItem.
' @param collectionObj - Parâmetro da rotina.
' @param index - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function SafeCollectionItem(ByVal collectionObj As Object, _
                                    ByVal index As Long) As Object
    On Error Resume Next
    Set SafeCollectionItem = collectionObj.item(CLng(index))
    If Err.Number <> 0 Then
        Set SafeCollectionItem = Nothing
        Err.Clear
    End If
    On Error GoTo 0
End Function

'******************************************************************************
' Nome do elemento: SafeChildrenCount (Function)
' Resumo: Retorna o resultado calculado por SafeChildrenCount.
' @param container - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function SafeChildrenCount(ByVal container As Object) As Long
    On Error Resume Next
    SafeChildrenCount = container.Children.Count
    If Err.Number <> 0 Then
        SafeChildrenCount = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

'******************************************************************************
' Nome do elemento: SafeChildItem (Function)
' Resumo: Retorna o resultado calculado por SafeChildItem.
' @param container - Parâmetro da rotina.
' @param index - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function SafeChildItem(ByVal container As Object, _
                               ByVal index As Long) As Object
    On Error Resume Next
    Set SafeChildItem = container.Children.item(CLng(index))
    If Err.Number <> 0 Then
        Set SafeChildItem = Nothing
        Err.Clear
    End If
    On Error GoTo 0
End Function

'******************************************************************************
' Nome do elemento: NameTypeMatches (Function)
' Resumo: Retorna o resultado calculado por NameTypeMatches.
' @param candidate - Parâmetro da rotina.
' @param resolvedName - Parâmetro da rotina.
' @param resolvedType - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function NameTypeMatches(ByVal candidate As Object, _
                                 ByVal resolvedName As String, _
                                 ByVal resolvedType As String) As Boolean
    If candidate Is Nothing Then Exit Function

    Dim candName As String
    Dim candType As String
    candName = TryGetStringProperty(candidate, "Name")
    candType = TryGetStringProperty(candidate, "Type")

    If LenB(resolvedType) > 0 Then
        If StrComp(candType, resolvedType, vbTextCompare) <> 0 Then Exit Function
    End If

    If LenB(resolvedName) > 0 Then
        If StrComp(candName, resolvedName, vbTextCompare) <> 0 Then Exit Function
    End If

    NameTypeMatches = True
End Function

'******************************************************************************
' Nome do elemento: FindMatchRecursive (Function)
' Resumo: Retorna o resultado calculado por FindMatchRecursive.
' @param container - Parâmetro da rotina.
' @param resolvedName - Parâmetro da rotina.
' @param resolvedType - Parâmetro da rotina.
' @param sSubType - Parâmetro da rotina.
' @param sText - Parâmetro da rotina.
' @param sSubstrId - Parâmetro da rotina.
' @param sTitle - Parâmetro da rotina.
' @param sTooltip - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function FindMatchRecursive(ByVal container As Object, _
                                    ByVal resolvedName As String, _
                                    ByVal resolvedType As String, _
                                    ByVal sSubType As String, _
                                    ByVal sText As String, _
                                    ByVal sSubstrId As String, _
                                    ByVal sTitle As String, _
                                    ByVal sTooltip As String) As Object
    Dim cnt As Long
    Dim i As Long
    Dim child As Object

    cnt = SafeChildrenCount(container)
    Trace "FindMatchRecursive: containerId='" & TryGetStringProperty(container, "Id") & "' children=" & cnt
    For i = 0 To cnt - 1
        Set child = SafeChildItem(container, i)
        If Not child Is Nothing Then
            If NameTypeMatches(child, resolvedName, resolvedType) Then
                If CandidateMatches(child, sSubType, sText, sSubstrId, sTitle, sTooltip) Then
                    Trace "FindMatchRecursive: match id='" & TryGetStringProperty(child, "Id") & "'"
                    Set FindMatchRecursive = child
                    Exit Function
                End If
            End If

            Set FindMatchRecursive = FindMatchRecursive(child, resolvedName, resolvedType, sSubType, sText, sSubstrId, sTitle, sTooltip)
            If Not FindMatchRecursive Is Nothing Then Exit Function
        End If
    Next i
End Function

'******************************************************************************
' Nome do elemento: ResolveTypeAndName (Function)
' Resumo: Retorna o resultado calculado por ResolveTypeAndName.
' @param sIdOrPrefName - Parâmetro da rotina.
' @param resolvedName - Parâmetro da rotina.
' @param resolvedType - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function ResolveTypeAndName(ByVal sIdOrPrefName As String, _
                                    ByRef resolvedName As String, _
                                    ByRef resolvedType As String) As Boolean
    Dim mapData As Variant
    Dim i As Long
    Dim prefix As String

    mapData = PrefixTypeMap()
    resolvedName = sIdOrPrefName
    resolvedType = vbNullString

    For i = LBound(mapData) To UBound(mapData)
        prefix = mapData(i)(0)
        If Len(prefix) > 0 Then
            If LCase$(Left$(sIdOrPrefName, Len(prefix))) = prefix Then
                resolvedType = mapData(i)(1)
                If InStr(sIdOrPrefName, "[") > 0 Then
                    resolvedName = sIdOrPrefName
                Else
                    resolvedName = Mid$(sIdOrPrefName, Len(prefix) + 1)
                    If LenB(resolvedName) = 0 Then
                        If prefix <> "lbl" Then resolvedName = sIdOrPrefName
                    End If
                End If
                ResolveTypeAndName = (LenB(resolvedType) > 0)
                Trace "ResolveTypeAndName: input='" & sIdOrPrefName & "' name='" & resolvedName & "' type='" & resolvedType & "'"
                Exit Function
            End If
        End If
    Next i
    Trace "ResolveTypeAndName: no match for '" & sIdOrPrefName & "'"
End Function

'******************************************************************************
' Nome do elemento: PrefixTypeMap (Function)
' Resumo: Retorna o resultado calculado por PrefixTypeMap.
' @return Variant - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function PrefixTypeMap() As Variant
    PrefixTypeMap = Array( _
        Array("shellcont", "GuiContainerShell"), _
        Array("shell", "GuiShell"), _
        Array("ssub", "GuiSimpleContainer"), _
        Array("sub", "GuiSubcontainer"), _
        Array("tabs", "GuiTabStrip"), _
        Array("tabp", "GuiTab"), _
        Array("tbar", "GuiToolbar"), _
        Array("mbar", "GuiMenubar"), _
        Array("menu", "GuiMenu"), _
        Array("cntl", "GuiCustomControl"), _
        Array("ctxt", "GuiCTextField"), _
        Array("okcd", "GuiOkCodeField"), _
        Array("sbar", "GuiStatusbar"), _
        Array("txt", "GuiTextField"), _
        Array("btn", "GuiButton"), _
        Array("lbl", "GuiLabel"), _
        Array("chk", "GuiCheckBox"), _
        Array("rad", "GuiRadioButton"), _
        Array("cmb", "GuiComboBox"), _
        Array("tbl", "GuiTableControl"), _
        Array("usr", "GuiUserArea") _
    )
End Function

'******************************************************************************
' Nome do elemento: CandidateMatches (Function)
' Resumo: Retorna o resultado calculado por CandidateMatches.
' @param candidate - Parâmetro da rotina.
' @param sSubType - Parâmetro da rotina.
' @param sText - Parâmetro da rotina.
' @param sSubstrId - Parâmetro da rotina.
' @param sTitle - Parâmetro da rotina.
' @param sTooltip - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function CandidateMatches(ByVal candidate As Object, _
                                  ByVal sSubType As String, _
                                  ByVal sText As String, _
                                  ByVal sSubstrId As String, _
                                  ByVal sTitle As String, _
                                  ByVal sTooltip As String) As Boolean

    If candidate Is Nothing Then
        Trace "CandidateMatches: candidate is Nothing"
        Exit Function
    End If

    If LenB(sSubType) > 0 Then
        If StrComp(TryGetStringProperty(candidate, "SubType"), sSubType, vbTextCompare) <> 0 Then
            Trace "CandidateMatches: SubType mismatch"
            Exit Function
        End If
    End If

    If Not PropertyContains(candidate, "Text", sText) Then Exit Function
    If Not PropertyContains(candidate, "Id", sSubstrId) Then Exit Function
    If Not PropertyContains(candidate, "Title", sTitle) Then Exit Function
    If Not PropertyContains(candidate, "Tooltip", sTooltip) Then Exit Function

    CandidateMatches = True
End Function

'******************************************************************************
' Nome do elemento: PropertyContains (Function)
' Resumo: Retorna o resultado calculado por PropertyContains.
' @param candidate - Parâmetro da rotina.
' @param propName - Parâmetro da rotina.
' @param expectedSubStr - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function PropertyContains(ByVal candidate As Object, _
                                  ByVal propName As String, _
                                  ByVal expectedSubStr As String) As Boolean
    Dim propValue As String

    If LenB(expectedSubStr) = 0 Then
        PropertyContains = True
        Exit Function
    End If

    propValue = TryGetStringProperty(candidate, propName)
    If LenB(propValue) = 0 Then
        Trace "PropertyContains: missing property '" & propName & "'"
        PropertyContains = False
    Else
        PropertyContains = (InStr(1, propValue, expectedSubStr, vbTextCompare) > 0)
        If Not PropertyContains Then Trace "PropertyContains: mismatch prop='" & propName & "' expected='" & expectedSubStr & "'"
    End If
End Function

'******************************************************************************
' Nome do elemento: TryGetStringProperty (Function)
' Resumo: Retorna o resultado calculado por TryGetStringProperty.
' @param candidate - Parâmetro da rotina.
' @param propName - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function TryGetStringProperty(ByVal candidate As Object, _
                                      ByVal propName As String) As String
    Dim value As Variant
    On Error Resume Next
    value = CallByName(candidate, propName, VbGet)
    If Err.Number <> 0 Then
        value = vbNullString
        Err.Clear
    End If
    On Error GoTo 0

    If IsNull(value) Then
        TryGetStringProperty = vbNullString
    Else
        TryGetStringProperty = CStr(value)
    End If
End Function

'******************************************************************************
' Nome do elemento: ExistsObj (Function)
' Resumo: Retorna o controle SAP GUI se existir na janela, ou Nothing caso
'         a janela não exista ou o controle/texto não seja encontrado.
' @param sWindowId   — Identificador da janela (ex: "wnd[0]").
' @param sTechName   — Nome técnico do controle (ex: "txtMESSTXT1").
' @param sSearchText — Texto a procurar no .Text do controle.
' @return Object     — Instância do controle ou Nothing.
'******************************************************************************
Public Function ExistsObj( _
    ByVal sWindowId As String, _
    Optional ByVal sIdOrPrefName As String = vbNullString, _
    Optional ByVal sText As String = vbNullString, _
    Optional ByVal sParentText As String = vbNullString, _
    Optional ByVal sSubstrId As String = vbNullString, _
    Optional ByVal sSubType As String = vbNullString, _
    Optional ByVal sTitle As String = vbNullString, _
    Optional ByVal sTooltip As String = vbNullString _
) As Boolean

    Dim foundObj As Object
    Dim wasManual As Boolean

    Trace "ExistsObj: checking"
    wasManual = pManualOn
    If wasManual Then pManualOn = False

    Set foundObj = Me.FindObj(sWindowId, sIdOrPrefName, sText, sParentText, sSubstrId, sSubType, sTitle, sTooltip)
    ExistsObj = Not foundObj Is Nothing

    If wasManual Then pManualOn = True
End Function

'******************************************************************************
' Nome do elemento: SetIfChangeable (Sub)
' Resumo: Atualiza um campo somente se o controle existir, for editável e o
'         valor for diferente do atual.
' @param controlId - Nome técnico/ID do controle.
' @param newValue  - Valor desejado.
' @param propName  - Propriedade a escrever: "text" (padrão), "key" ou "selected".
'******************************************************************************
Public Function SetIfChangeable( _
    ByVal controlId As String, _
    ByVal newValue As String _
) As Boolean
    Dim ctrl         As Object
    Dim isChangeable As Boolean
    Dim desired      As Boolean
    Dim propName     As String

    Trace "SetIfChangeable: controlId='" & controlId & "'"
    SetIfChangeable = False
    If Not Me.ExistsObj(controlId) Then Exit Function
    Set ctrl = Me.FindObj(controlId)
    If ctrl Is Nothing Then Exit Function

    Select Case LCase$(Left$(controlId, 3))
        Case "cmb": propName = "key"
        Case "chk": propName = "selected"
        Case Else:   propName = "text"
    End Select

    On Error Resume Next
    isChangeable = ctrl.Changeable
    If Err.Number <> 0 Then
        Err.Clear
        isChangeable = True
    End If
    On Error GoTo 0
    If Not isChangeable Then Exit Function

    Select Case propName
        Case "key"
            On Error Resume Next
            If ctrl.key <> newValue Then
                ctrl.key = newValue
                SetIfChangeable = True
            End If
            On Error GoTo 0
        Case "selected"
            desired = (newValue = "X") Or (newValue = "1") Or (newValue = "true") Or (newValue = True)
            On Error Resume Next
            If ctrl.Selected <> desired Then
                ctrl.Selected = desired
                SetIfChangeable = True
            End If
            On Error GoTo 0
        Case Else
            On Error Resume Next
            If ctrl.Text <> newValue Then
                ctrl.Text = newValue
                SetIfChangeable = True
            End If
            On Error GoTo 0
    End Select
End Function



'******************************************************************************
' Nome do elemento: guiTable (Function)
' Resumo: Factory híbrido para criar wrapper fwGuiTableControl a partir de ID ou Name.
' @param sTableName - Nome técnico ou ID da tabela SAP GUI.
' @return fwGuiTableControl - Wrapper inicializado para o controle de tabela.
' @raises Erro vbObjectError+513 se a tabela não for encontrada.
'******************************************************************************
Public Function GuiTable( _
    ByVal sTableName As String _
) As fwGuiTableControl

    Dim foundCtrl As Object
    Trace "GuiTable: name='" & sTableName & "'"
    Set foundCtrl = Me.FindObj(pWndId, sTableName)
    If foundCtrl Is Nothing Then
        Err.Raise vbObjectError + 513, _
                  "cWindow.guiTable", _
                  "Tabela '" & sTableName & "' não encontrada em " & pWndId
    End If

    Dim oWrapper As New fwGuiTableControl
    oWrapper.Init pSession, pWnd, foundCtrl
    Set GuiTable = oWrapper
End Function

'******************************************************************************
' Nome do elemento: guiMenu (Function)
' Resumo: Cria e inicializa um wrapper fwGuiMenu para acesso ao menu SAP GUI.
' @param ParentMenu - Nome do menu pai a ser acessado.
' @param SubMenu - Nome do submenu/opção filho (opcional).
' @return fwGuiMenu - Wrapper inicializado para o menu especificado; Nothing em caso de erro.
'******************************************************************************
Public Function GuiMenu( _
    ByVal ParentMenu As String, _
    Optional ByVal SubMenu As String = "" _
) As fwGuiMenu

    Dim foundMenu As Object
    Dim oMenu As fwGuiMenu
    On Error GoTo SafeReturn

    Trace "GuiMenu: parent='" & ParentMenu & "' submenu='" & SubMenu & "'"
    Set oMenu = New fwGuiMenu
    Set foundMenu = Me.FindObj(pWndId, vbNullString, SubMenu, ParentMenu)
    If Not foundMenu Is Nothing Then oMenu.Init foundMenu

SafeReturn:
    Set GuiMenu = oMenu
End Function

'******************************************************************************
' Nome do elemento: FindMenuByText (Function)
' Resumo: Retorna o resultado calculado por FindMenuByText.
' @param sWindowId - Parâmetro da rotina.
' @param sText - Parâmetro da rotina.
' @param sParentText - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function FindMenuByText( _
    ByVal sWindowId As String, _
    ByVal sText As String, _
    ByVal sParentText As String _
) As Object
    On Error GoTo ErrHandler
    Dim oMenuBar   As Object
    Dim colMenus   As Object
    Dim colParents As Object
    Dim i          As Long
    Dim oItem      As Object

    Set oMenuBar = SafeFindById(sWindowId & "/mbar")
    If oMenuBar Is Nothing Then Exit Function

    On Error Resume Next
    Set colMenus = oMenuBar.FindAllByName(sText, "GuiMenu")
    On Error GoTo ErrHandler

    If Not colMenus Is Nothing And colMenus.Count > 0 Then
        If LenB(sParentText) > 0 Then
            For i = 0 To colMenus.Count - 1
                If TypeName(colMenus.item(i - 0)) = "GuiMenu" _
                   And CStr(colMenus.item(i - 0).Parent.Name) = CStr(sParentText) Then
                    Set FindMenuByText = colMenus.item(i - 0)
                    Exit Function
                End If
            Next i
        Else
            Set FindMenuByText = colMenus.item(0)
            Exit Function
        End If
    End If

    If LenB(sParentText) > 0 Then
        On Error Resume Next
        Set colParents = oMenuBar.FindAllByName(sParentText, "GuiMenu")
        On Error GoTo ErrHandler

        If Not colParents Is Nothing And colParents.Count > 0 Then
            For i = 0 To colParents.item(0).Children.Count - 1
                Set oItem = colParents.item(0).Children.item(i - 0)
                If TypeName(oItem) = "GuiMenu" _
                   And CStr(oItem.Name) = CStr(sText) Then
                    Set FindMenuByText = oItem
                    Exit Function
                End If
            Next i
        End If
    End If

    Set FindMenuByText = FindMenuRecursive(sWindowId, sText, sParentText)
    Exit Function

ErrHandler:
    Set FindMenuByText = Nothing
End Function

'******************************************************************************
' Nome do elemento: FindMenuRecursive (Function)
' Resumo: Busca recursivamente um item de menu por Name e ParentName.
' @param ItemName — Nome do item de menu a localizar.
' @param ParentName — Nome do menu pai imediato (opcional).
' @param prefix — Caminho parcial de IDs usado internamente (opcional).
' @return Object — Objeto GuiMenu localizado ou Nothing se não encontrado.
'******************************************************************************
Private Function FindMenuRecursive( _
    ByVal windowId As String, _
    ByVal ItemName As String, _
    ByVal ParentName As String, _
    Optional ByVal prefix As String = "" _
) As Object
    Dim oParent    As Object
    Dim cnt        As Long
    Dim i          As Long
    Dim childId    As String
    Dim oChild     As Object
    Dim parentId   As String
    Dim oParentObj As Object

    On Error Resume Next
    If prefix = "" Then
        Set oParent = pSession.findbyid(windowId & "/mbar")
    Else
        Set oParent = pSession.findbyid(windowId & "/mbar" & prefix)
    End If
    On Error GoTo 0
    If oParent Is Nothing Then Exit Function

    cnt = oParent.Children.Count
    For i = 0 To cnt - 1
        childId = windowId & "/mbar" & prefix & "/menu[" & i & "]"
        On Error Resume Next
        Set oChild = pSession.findbyid(childId)
        On Error GoTo 0
        If Not oChild Is Nothing Then
            If CStr(oChild.Name) = CStr(ItemName) Then
                If ParentName = "" Then
                    Set FindMenuRecursive = oChild
                    Exit Function
                Else
                    parentId = Left(childId, InStrRev(childId, "/") - 1)
                    On Error Resume Next
                    Set oParentObj = pSession.findbyid(parentId)
                    On Error GoTo 0
                    If Not oParentObj Is Nothing And CStr(oParentObj.Name) = CStr(ParentName) Then
                        Set FindMenuRecursive = oChild
                        Exit Function
                    End If
                End If
            End If
            Set FindMenuRecursive = FindMenuRecursive(windowId, ItemName, ParentName, prefix & "/menu[" & i & "]")
            If Not FindMenuRecursive Is Nothing Then Exit Function
        End If
    Next i
End Function

'******************************************************************************
' Nome do elemento: SbarText (Function)
' Resumo: Retorna o resultado calculado por SbarText.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function SbarText() As String
    Trace "SbarText"
    SbarText = Me.FindObj("sbar").Text
End Function

'******************************************************************************
' Nome do elemento: oGuiTable (Property Get)
' Resumo: Obtem o valor da propriedade oGuiTable.
' @return fwGuiTableControl - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get oGuiTable() As fwGuiTableControl
    Trace "oGuiTable"
    If pGuiTable Is Nothing Then
        Set pGuiTable = New fwGuiTableControl
        pGuiTable.Init pSession, pWnd
    End If
    Set oGuiTable = pGuiTable
End Property

'******************************************************************************
' Nome do elemento: oGuiTree (Property Get)
' Resumo: Obtem o valor da propriedade oGuiTree.
' @return fwGuiTree - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get oGuiTree() As fwGuiTree
    Trace "oGuiTree"
    If pGuiTree Is Nothing Then
        Set pGuiTree = New fwGuiTree
        pGuiTree.SetSession pSession
    End If
    Set oGuiTree = pGuiTree
End Property

'******************************************************************************
' Nome do elemento: IsSessionAlive (Function)
' Resumo: Retorna o resultado calculado por IsSessionAlive.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function IsSessionAlive() As Boolean
    On Error GoTo Cleanup
    Trace "IsSessionAlive"
    ' Tenta ler algo que sempre exista na session:
    Dim dummy As String
    dummy = pSession.Info.SystemName   ' ou: cSession.ActiveWindow.Title
    IsSessionAlive = True
    Exit Function

Cleanup:
    IsSessionAlive = False
    Err.Clear
End Function

'******************************************************************************
' Nome do elemento: ActiveWindow (Property Get)
' Resumo: Obtem o valor da propriedade ActiveWindow.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get ActiveWindow() As Object
    Set ActiveWindow = pSession.ActiveWindow
End Property

'******************************************************************************
' Nome do elemento: Info (Property Get)
' Resumo: Obtem o valor da propriedade Info.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Info() As Object
    Set Info = pSession.Info
End Property


