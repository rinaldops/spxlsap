VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fwGuiTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' Nome do elemento: fwGuiTree (Classe)
' Resumo: Wrapper flexível para GuiTree com late binding e detecção automática de padding.
'******************************************************************************

Option Explicit

Private pSession   As Object   ' Sessão SAP GUI para findById
Private pGuiTreeId    As String   ' ID do controle (ex: "wnd[0]/shellcont/shell")
Private pPadLength As Long     ' Quantidade de dígitos detectada (zeros à esquerda)

'******************************************************************************
' Nome do elemento: SetSession (Sub)
' Resumo: Executa a rotina SetSession.
' @param sessionObj - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub SetSession(ByVal sessionObj As Object)
    Set pSession = sessionObj
    pPadLength = 0
End Sub

'******************************************************************************
' Nome do elemento: ID (Property Let)
' Resumo: Define o valor da propriedade ID.
' @param treeId - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Property Let ID(ByVal treeId As String)
    If pSession Is Nothing Then
        Err.Raise vbObjectError + 1002, "fwGuiTree.ID", "Sessão não inicializada."
    End If
    pGuiTreeId = treeId
    pPadLength = 0
End Property

'******************************************************************************
' Nome do elemento: ID (Property Get)
' Resumo: Obtem o valor da propriedade ID.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get ID() As String
    ID = pGuiTreeId
End Property

'******************************************************************************
' Nome do elemento: GetTree (Function)
' Resumo: Retorna o resultado calculado por GetTree.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Function GetTree() As Object
    On Error GoTo ErrHandler
    Set GetTree = pSession.findbyid(pGuiTreeId)
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "fwGuiTree.GetTree", "Controle '" & pGuiTreeId & "' não encontrado."
End Function

'******************************************************************************
' Nome do elemento: DetectPadLength (Sub)
' Resumo: Executa a rotina DetectPadLength.
' @param tree - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub DetectPadLength(ByVal tree As Object)
    Dim key As String
    If pPadLength = 0 Then
        On Error Resume Next
        key = tree.SelectedItemNode
        On Error GoTo 0
        If Len(key) > 0 Then
            pPadLength = Len(key)
        Else
            Dim coll As Object
            Set coll = tree.GetAllNodeKeys
            If coll.Count > 0 Then
                key = CStr(coll.ElementAt(0))
                pPadLength = Len(key)
            Else
                pPadLength = 4  ' Fallback arbitrário
            End If
        End If
    End If
End Sub

'******************************************************************************
' Nome do elemento: GetPaddedID (Function)
' Resumo: Retorna o resultado calculado por GetPaddedID.
' @param idx - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function GetPaddedID(ByVal idx As Long) As String
    GetPaddedID = Right(String(pPadLength, "0") & CStr(idx), pPadLength)
End Function

'******************************************************************************
' Nome do elemento: SelectItem (Sub)
' Resumo: Executa a rotina SelectItem.
' @param idx - Parâmetro da rotina.
' @param colName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub SelectItem(ByVal idx As Long, ByVal colName As String)
    Dim tree As Object
    Set tree = GetTree()
    DetectPadLength tree
    tree.SelectItem GetPaddedID(idx), colName
End Sub

'******************************************************************************
' Nome do elemento: EnsureVisibleHorizontalItem (Sub)
' Resumo: Executa a rotina EnsureVisibleHorizontalItem.
' @param idx - Parâmetro da rotina.
' @param colName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub EnsureVisibleHorizontalItem(ByVal idx As Long, ByVal colName As String)
    Dim tree As Object
    Set tree = GetTree()
    DetectPadLength tree
    tree.EnsureVisibleHorizontalItem GetPaddedID(idx), colName
End Sub

'******************************************************************************
' Nome do elemento: ClickLink (Sub)
' Resumo: Executa a rotina ClickLink.
' @param idx - Parâmetro da rotina.
' @param colName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub ClickLink(ByVal idx As Long, ByVal colName As String)
    Dim tree As Object
    Set tree = GetTree()
    DetectPadLength tree
    tree.ClickLink GetPaddedID(idx), colName
End Sub

'******************************************************************************
' Nome do elemento: DumpState (Sub)
' Resumo: Executa a rotina DumpState.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub DumpState()
    Dim tree As Object: Set tree = GetTree()
    tree.DumpState
End Sub

'******************************************************************************
' Nome do elemento: SetFocus (Sub)
' Resumo: Executa a rotina SetFocus.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub SetFocus()
    Dim tree As Object: Set tree = GetTree()
    tree.SetFocus
End Sub

'******************************************************************************
' Nome do elemento: Visualize (Sub)
' Resumo: Executa a rotina Visualize.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub Visualize()
    Dim tree As Object: Set tree = GetTree()
    tree.Visualize
End Sub

'******************************************************************************
' Nome do elemento: SelectContextMenuItem (Sub)
' Resumo: Executa a rotina SelectContextMenuItem.
' @param FunctionCode - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub SelectContextMenuItem(ByVal FunctionCode As String)
    Dim tree As Object: Set tree = GetTree()
    tree.SelectContextMenuItem FunctionCode
End Sub

'******************************************************************************
' Nome do elemento: SelectContextMenuItemByPosition (Sub)
' Resumo: Executa a rotina SelectContextMenuItemByPosition.
' @param PositionDesc - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub SelectContextMenuItemByPosition(ByVal PositionDesc As String)
    Dim tree As Object: Set tree = GetTree()
    tree.SelectContextMenuItemByPosition PositionDesc
End Sub

'******************************************************************************
' Nome do elemento: SelectContextMenuItemByText (Sub)
' Resumo: Executa a rotina SelectContextMenuItemByText.
' @param text - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub SelectContextMenuItemByText(ByVal Text As String)
    Dim tree As Object: Set tree = GetTree()
    tree.SelectContextMenuItemByText Text
End Sub

'------------------------------------------------------------------------------
' GuiTree–specific Methods
'------------------------------------------------------------------------------

'******************************************************************************
' Nome do elemento: ChangeCheckbox (Sub)
' Resumo: Altera o estado do checkbox em um item específico.
' @param idx — Índice numérico do item (0-based).
' @param colName — Nome da coluna que contém o checkbox.
' @param Checked — True para marcar; False para desmarcar.
'******************************************************************************
Public Sub ChangeCheckbox( _
    ByVal idx As Long, _
    ByVal colName As String, _
    ByVal Checked As Boolean _
)
    Dim tree As Object: Set tree = GetTree()
    DetectPadLength tree
    tree.ChangeCheckbox GetPaddedID(idx), colName, Checked
End Sub

'******************************************************************************
' Nome do elemento: CollapseNode (Sub)
' Resumo: Executa a rotina CollapseNode.
' @param idx - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub CollapseNode(ByVal idx As Long)
    Dim tree As Object: Set tree = GetTree()
    DetectPadLength tree
    tree.CollapseNode GetPaddedID(idx)
End Sub

'******************************************************************************
' Nome do elemento: DefaultContextMenu (Sub)
' Resumo: Executa a rotina DefaultContextMenu.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub DefaultContextMenu()
    Dim tree As Object: Set tree = GetTree()
    tree.DefaultContextMenu
End Sub

'******************************************************************************
' Nome do elemento: DoubleClickItem (Sub)
' Resumo: Emula duplo clique em um item específico na árvore.
' @param idx — Índice numérico do item (0-based).
' @param colName — Nome da coluna do item.
'******************************************************************************
Public Sub DoubleClickItem( _
    ByVal idx As Long, _
    ByVal colName As String _
)
    Dim tree As Object: Set tree = GetTree()
    DetectPadLength tree
    tree.DoubleClickItem GetPaddedID(idx), colName
End Sub

'******************************************************************************
' Nome do elemento: DoubleClickNode (Sub)
' Resumo: Executa a rotina DoubleClickNode.
' @param idx - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub DoubleClickNode(ByVal idx As Long)
    Dim tree As Object: Set tree = GetTree()
    DetectPadLength tree
    tree.DoubleClickNode GetPaddedID(idx)
End Sub

'******************************************************************************
' Nome do elemento: ExpandNode (Sub)
' Resumo: Executa a rotina ExpandNode.
' @param idx - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub ExpandNode(ByVal idx As Long)
    Dim tree As Object: Set tree = GetTree()
    DetectPadLength tree
    tree.ExpandNode GetPaddedID(idx)
End Sub

'******************************************************************************
' Nome do elemento: FindNodeKeyByPath (Function)
' Resumo: Retorna o resultado calculado por FindNodeKeyByPath.
' @param Path - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function FindNodeKeyByPath(ByVal Path As String) As String
    Dim tree As Object: Set tree = GetTree()
    FindNodeKeyByPath = tree.FindNodeKeyByPath(Path)
End Function

'******************************************************************************
' Nome do elemento: GetNodeKeyByPath (Function)
' Resumo: Retorna o resultado calculado por GetNodeKeyByPath.
' @param Path - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodeKeyByPath(ByVal Path As String) As String
    Dim tree As Object: Set tree = GetTree()
    GetNodeKeyByPath = tree.GetNodeKeyByPath(Path)
End Function

'******************************************************************************
' Nome do elemento: GetAbapImage (Function)
' Resumo: Recupera o código de imagem ABAP em um item.
' @param idx — Índice numérico do item (0-based).
' @param colName — Nome da coluna do item.
' @return String — Código da imagem (ex: "01").
'******************************************************************************
Public Function GetAbapImage( _
    ByVal idx As Long, _
    ByVal colName As String _
) As String
    Dim tree As Object: Set tree = GetTree()
    DetectPadLength tree
    GetAbapImage = tree.GetAbapImage(GetPaddedID(idx), colName)
End Function

'******************************************************************************
' Nome do elemento: GetAllNodeKeys (Function)
' Resumo: Retorna o resultado calculado por GetAllNodeKeys.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetAllNodeKeys() As Object
    Dim tree As Object: Set tree = GetTree()
    Set GetAllNodeKeys = tree.GetAllNodeKeys
End Function

'******************************************************************************
' Nome do elemento: GetCheckBoxState (Function)
' Resumo: Obtém o estado de um checkbox em um item.
' @param idx — Índice numérico do item (0-based).
' @param colName — Nome da coluna do checkbox.
' @return Byte — 1 se marcado; 0 se desmarcado.
'******************************************************************************
Public Function GetCheckBoxState( _
    ByVal idx As Long, _
    ByVal colName As String _
) As Byte
    Dim tree As Object: Set tree = GetTree()
    DetectPadLength tree
    GetCheckBoxState = tree.GetCheckBoxState(GetPaddedID(idx), colName)
End Function

'******************************************************************************
' Nome do elemento: GetColumnCol (Function)
' Resumo: Retorna o resultado calculado por GetColumnCol.
' @param colName - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetColumnCol(ByVal colName As String) As Object
    Dim tree As Object: Set tree = GetTree()
    Set GetColumnCol = tree.GetColumnCol(colName)
End Function

'******************************************************************************
' Nome do elemento: GetColumnHeaders (Function)
' Resumo: Retorna o resultado calculado por GetColumnHeaders.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetColumnHeaders() As Object
    Dim tree As Object: Set tree = GetTree()
    Set GetColumnHeaders = tree.GetColumnHeaders
End Function

'******************************************************************************
' Nome do elemento: GetColumnIndexFromName (Function)
' Resumo: Retorna o resultado calculado por GetColumnIndexFromName.
' @param key - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetColumnIndexFromName(ByVal key As String) As Long
    Dim tree As Object: Set tree = GetTree()
    GetColumnIndexFromName = tree.GetColumnIndexFromName(key)
End Function

'******************************************************************************
' Nome do elemento: GetNodeItemHeaders (Function)
' Resumo: Retorna o resultado calculado por GetNodeItemHeaders.
' @param idx - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodeItemHeaders(ByVal idx As Long) As Object
    Dim tree As Object: Set tree = GetTree()
    DetectPadLength tree
    Set GetNodeItemHeaders = tree.GetNodeItemHeaders(GetPaddedID(idx))
End Function

'******************************************************************************
' Nome do elemento: GetNodeLeft (Function)
' Resumo: Retorna o resultado calculado por GetNodeLeft.
' @param key - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodeLeft(ByVal key As String) As Long
    Dim tree As Object: Set tree = GetTree()
    GetNodeLeft = tree.GetNodeLeft(key)
End Function

'******************************************************************************
' Nome do elemento: GetNodePathByKey (Function)
' Resumo: Retorna o resultado calculado por GetNodePathByKey.
' @param key - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodePathByKey(ByVal key As String) As String
    Dim tree As Object: Set tree = GetTree()
    GetNodePathByKey = tree.GetNodePathByKey(key)
End Function

'******************************************************************************
' Nome do elemento: GetNodesCol (Function)
' Resumo: Retorna o resultado calculado por GetNodesCol.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodesCol() As Object
    Dim tree As Object: Set tree = GetTree()
    Set GetNodesCol = tree.GetNodesCol
End Function

'******************************************************************************
' Nome do elemento: GetNodeStyle (Function)
' Resumo: Retorna o resultado calculado por GetNodeStyle.
' @param NodeKey - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodeStyle(ByVal NodeKey As String) As Long
    Dim tree As Object: Set tree = GetTree()
    GetNodeStyle = tree.GetNodeStyle(NodeKey)
End Function

'******************************************************************************
' Nome do elemento: GetNodeTextByKey (Function)
' Resumo: Retorna o resultado calculado por GetNodeTextByKey.
' @param key - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodeTextByKey(ByVal key As String) As String
    Dim tree As Object: Set tree = GetTree()
    GetNodeTextByKey = tree.GetNodeTextByKey(key)
End Function

'******************************************************************************
' Nome do elemento: GetNodeTextByPath (Function)
' Resumo: Retorna o resultado calculado por GetNodeTextByPath.
' @param Path - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodeTextByPath(ByVal Path As String) As String
    Dim tree As Object: Set tree = GetTree()
    GetNodeTextByPath = tree.GetNodeTextByPath(Path)
End Function

'******************************************************************************
' Nome do elemento: GetNodeTextColor (Function)
' Resumo: Retorna o resultado calculado por GetNodeTextColor.
' @param key - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodeTextColor(ByVal key As String) As Long
    Dim tree As Object: Set tree = GetTree()
    GetNodeTextColor = tree.GetNodeTextColor(key)
End Function

'******************************************************************************
' Nome do elemento: GetNodeToolTip (Function)
' Resumo: Retorna o resultado calculado por GetNodeToolTip.
' @param NodeKey - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodeToolTip(ByVal NodeKey As String) As String
    Dim tree As Object: Set tree = GetTree()
    GetNodeToolTip = tree.GetNodeToolTip(NodeKey)
End Function

'******************************************************************************
' Nome do elemento: GetParent (Function)
' Resumo: Retorna o resultado calculado por GetParent.
' @param CKey - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetParent(ByVal CKey As String) As String
    Dim tree As Object: Set tree = GetTree()
    GetParent = tree.GetParent(CKey)
End Function

'******************************************************************************
' Nome do elemento: GetPreviousNodeKey (Function)
' Resumo: Retorna o resultado calculado por GetPreviousNodeKey.
' @param NodeKey - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetPreviousNodeKey(ByVal NodeKey As String) As String
    Dim tree As Object: Set tree = GetTree()
    GetPreviousNodeKey = tree.GetPreviousNodeKey(NodeKey)
End Function

'******************************************************************************
' Nome do elemento: GetSelectedNodes (Function)
' Resumo: Retorna o resultado calculado por GetSelectedNodes.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetSelectedNodes() As Object
    Dim tree As Object: Set tree = GetTree()
    Set GetSelectedNodes = tree.GetSelectedNodes
End Function

'******************************************************************************
' Nome do elemento: GetSelectionMode (Function)
' Resumo: Retorna o resultado calculado por GetSelectionMode.
' @return Integer - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetSelectionMode() As Integer
    Dim tree As Object: Set tree = GetTree()
    GetSelectionMode = tree.GetSelectionMode
End Function

'******************************************************************************
' Nome do elemento: GetStyleDescription (Function)
' Resumo: Retorna o resultado calculado por GetStyleDescription.
' @param nStyle - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetStyleDescription(ByVal nStyle As Long) As String
    Dim tree As Object: Set tree = GetTree()
    GetStyleDescription = tree.GetStyleDescription(nStyle)
End Function

'******************************************************************************
' Nome do elemento: GetSubNodesCol (Function)
' Resumo: Retorna o resultado calculado por GetSubNodesCol.
' @param NodeKey - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetSubNodesCol(ByVal NodeKey As String) As Object
    Dim tree As Object: Set tree = GetTree()
    Set GetSubNodesCol = tree.GetSubNodesCol(NodeKey)
End Function

'******************************************************************************
' Nome do elemento: GetListTreeNodeItemCount (Function)
' Resumo: Retorna o resultado calculado por GetListTreeNodeItemCount.
' @param NodeKey - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetListTreeNodeItemCount(ByVal NodeKey As String) As Long
    Dim tree As Object: Set tree = GetTree()
    GetListTreeNodeItemCount = tree.GetListTreeNodeItemCount(NodeKey)
End Function

'******************************************************************************
' Nome do elemento: GetNextNodeKey (Function)
' Resumo: Retorna o resultado calculado por GetNextNodeKey.
' @param NodeKey - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNextNodeKey(ByVal NodeKey As String) As String
    Dim tree As Object: Set tree = GetTree()
    GetNextNodeKey = tree.GetNextNodeKey(NodeKey)
End Function

'******************************************************************************
' Nome do elemento: GetNodeAbapImage (Function)
' Resumo: Retorna o resultado calculado por GetNodeAbapImage.
' @param key - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodeAbapImage(ByVal key As String) As String
    Dim tree As Object: Set tree = GetTree()
    GetNodeAbapImage = tree.GetNodeAbapImage(key)
End Function

'******************************************************************************
' Nome do elemento: GetNodeChildrenCount (Function)
' Resumo: Retorna o resultado calculado por GetNodeChildrenCount.
' @param key - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodeChildrenCount(ByVal key As String) As Long
    Dim tree As Object: Set tree = GetTree()
    GetNodeChildrenCount = tree.GetNodeChildrenCount(key)
End Function

'******************************************************************************
' Nome do elemento: GetNodeChildrenCountByPath (Function)
' Resumo: Retorna o resultado calculado por GetNodeChildrenCountByPath.
' @param Path - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodeChildrenCountByPath(ByVal Path As String) As Long
    Dim tree As Object: Set tree = GetTree()
    GetNodeChildrenCountByPath = tree.GetNodeChildrenCountByPath(Path)
End Function

'******************************************************************************
' Nome do elemento: GetNodeIndex (Function)
' Resumo: Retorna o resultado calculado por GetNodeIndex.
' @param key - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetNodeIndex(ByVal key As String) As Long
    Dim tree As Object: Set tree = GetTree()
    GetNodeIndex = tree.GetNodeIndex(key)
End Function

'******************************************************************************
' Nome do elemento: PressHeader (Sub)
' Resumo: Executa a rotina PressHeader.
' @param headerName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub PressHeader(ByVal headerName As String)
    Dim tree As Object: Set tree = GetTree()
    tree.PressHeader headerName
End Sub

'******************************************************************************
' Nome do elemento: PressKey (Sub)
' Resumo: Executa a rotina PressKey.
' @param key - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub PressKey(ByVal key As String)
    Dim tree As Object: Set tree = GetTree()
    tree.PressKey key
End Sub

'******************************************************************************
' Nome do elemento: SelectColumn (Sub)
' Resumo: Executa a rotina SelectColumn.
' @param columnName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub SelectColumn(ByVal columnName As String)
    Dim tree As Object: Set tree = GetTree()
    tree.SelectColumn columnName
End Sub

'******************************************************************************
' Nome do elemento: SelectedItemColumn (Function)
' Resumo: Retorna o resultado calculado por SelectedItemColumn.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function SelectedItemColumn() As String
    Dim tree As Object: Set tree = GetTree()
    SelectedItemColumn = tree.SelectedItemColumn()
End Function

'******************************************************************************
' Nome do elemento: SelectedItemNode (Function)
' Resumo: Retorna o resultado calculado por SelectedItemNode.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function SelectedItemNode() As String
    Dim tree As Object: Set tree = GetTree()
    SelectedItemNode = tree.SelectedItemNode()
End Function

'******************************************************************************
' Nome do elemento: SelectNode (Sub)
' Resumo: Executa a rotina SelectNode.
' @param idx - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub SelectNode(ByVal idx As Long)
    Dim tree As Object: Set tree = GetTree()
    DetectPadLength tree
    tree.SelectNode GetPaddedID(idx)
End Sub

'******************************************************************************
' Nome do elemento: SetCheckBoxState (Sub)
' Resumo: Marca ou desmarca checkbox em célula específica.
' @param idx — Índice numérico do item (0-based).
' @param colName — Nome da coluna do checkbox.
' @param state — 0 para desmarcar; 1 para marcar.
'******************************************************************************
Public Sub SetCheckBoxState( _
    ByVal idx As Long, _
    ByVal colName As String, _
    ByVal state As Long _
)
    Dim tree As Object: Set tree = GetTree()
    DetectPadLength tree
    tree.SetCheckBoxState GetPaddedID(idx), colName, state
End Sub

'******************************************************************************
' Nome do elemento: SetColumnWidth (Sub)
' Resumo: Define a largura de uma coluna em pixels.
' @param ColumnName — Nome da coluna.
' @param Width — Largura em pixels.
'******************************************************************************
Public Sub SetColumnWidth( _
    ByVal columnName As String, _
    ByVal Width As Long _
)
    Dim tree As Object: Set tree = GetTree()
    tree.SetColumnWidth columnName, Width
End Sub

'******************************************************************************
' Nome do elemento: UnselectAll (Sub)
' Resumo: Executa a rotina UnselectAll.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub UnselectAll()
    Dim tree As Object: Set tree = GetTree()
    tree.UnselectAll
End Sub

'******************************************************************************
' Nome do elemento: UnselectColumn (Sub)
' Resumo: Executa a rotina UnselectColumn.
' @param columnName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub UnselectColumn(ByVal columnName As String)
    Dim tree As Object: Set tree = GetTree()
    tree.UnselectColumn columnName
End Sub

'******************************************************************************
' Nome do elemento: UnselectNode (Sub)
' Resumo: Executa a rotina UnselectNode.
' @param idx - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub UnselectNode(ByVal idx As Long)
    Dim tree As Object: Set tree = GetTree()
    DetectPadLength tree
    tree.UnselectNode GetPaddedID(idx)
End Sub

'******************************************************************************
' Nome do elemento: PressButton (Sub)
' Resumo: Emula clique em botão dentro de um item de árvore.
' @param idx — Índice numérico do item (0-based).
' @param ItemName — Nome do botão.
'******************************************************************************
Public Sub PressButton( _
    ByVal idx As Long, _
    ByVal ItemName As String _
)
    Dim tree As Object: Set tree = GetTree()
    DetectPadLength tree
    tree.PressButton GetPaddedID(idx), ItemName
End Sub

'******************************************************************************
' Nome do elemento: Name (Property Get)
' Resumo: Obtem o valor da propriedade Name.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Name() As String
    Name = GetTree.Name
End Property

'******************************************************************************
' Nome do elemento: Self (Property Get)
' Resumo: Obtem o valor da propriedade Self.
' @param sTech - Parâmetro da rotina.
' @return fwGuiTree - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Self(ByVal sTech As String) As fwGuiTree
'    Me.Name = sTech
'    Set Self = Me
End Property

'******************************************************************************
' Nome do elemento: Parent (Property Get)
' Resumo: Obtem o valor da propriedade Parent.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Parent() As Object
    Set Parent = GetTree.Parent
End Property

'******************************************************************************
' Nome do elemento: TypeName (Property Get)
' Resumo: Obtem o valor da propriedade TypeName.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get TypeName() As String
    TypeName = GetTree.Type
End Property

'******************************************************************************
' Nome do elemento: TypeAsNumber (Property Get)
' Resumo: Obtem o valor da propriedade TypeAsNumber.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get TypeAsNumber() As Long
    TypeAsNumber = GetTree.TypeAsNumber
End Property

'******************************************************************************
' Nome do elemento: Changeable (Property Get)
' Resumo: Obtem o valor da propriedade Changeable.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Changeable() As Boolean
    Changeable = GetTree.Changeable
End Property

'******************************************************************************
' Nome do elemento: DefaultTooltip (Property Get)
' Resumo: Obtem o valor da propriedade DefaultTooltip.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get DefaultTooltip() As String
    DefaultTooltip = GetTree.DefaultTooltip
End Property

'******************************************************************************
' Nome do elemento: IconName (Property Get)
' Resumo: Obtem o valor da propriedade IconName.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get IconName() As String
    IconName = GetTree.IconName
End Property

'******************************************************************************
' Nome do elemento: IsSymbolFont (Property Get)
' Resumo: Obtem o valor da propriedade IsSymbolFont.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get IsSymbolFont() As Boolean
    IsSymbolFont = GetTree.IsSymbolFont
End Property

'******************************************************************************
' Nome do elemento: Modified (Property Get)
' Resumo: Obtem o valor da propriedade Modified.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Modified() As Boolean
    Modified = GetTree.Modified
End Property

'******************************************************************************
' Nome do elemento: text (Property Get)
' Resumo: Obtem o valor da propriedade text.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Text() As String
    Text = GetTree.Text
End Property

'******************************************************************************
' Nome do elemento: Tooltip (Property Get)
' Resumo: Obtem o valor da propriedade Tooltip.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Tooltip() As String
    Tooltip = GetTree.Tooltip
End Property

'******************************************************************************
' Nome do elemento: Children (Property Get)
' Resumo: Obtem o valor da propriedade Children.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Children() As Object
    Set Children = GetTree.Children
End Property

'******************************************************************************
' Nome do elemento: DragDropSupported (Property Get)
' Resumo: Obtem o valor da propriedade DragDropSupported.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get DragDropSupported() As Boolean
    DragDropSupported = GetTree.DragDropSupported
End Property

'******************************************************************************
' Nome do elemento: Handle (Property Get)
' Resumo: Obtem o valor da propriedade Handle.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Handle() As Long
    Handle = GetTree.Handle
End Property

'******************************************************************************
' Nome do elemento: SubType (Property Get)
' Resumo: Obtem o valor da propriedade SubType.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get SubType() As String
    SubType = GetTree.SubType
End Property

'******************************************************************************
' Nome do elemento: ColumnOrder (Property Get)
' Resumo: Obtem o valor da propriedade ColumnOrder.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get ColumnOrder() As Object
    Set ColumnOrder = GetTree.ColumnOrder
End Property

'******************************************************************************
' Nome do elemento: SelectedNode (Property Get)
' Resumo: Obtem o valor da propriedade SelectedNode.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get selectedNode() As String
    selectedNode = GetTree.selectedNode
End Property

'******************************************************************************
' Nome do elemento: TopNode (Property Get)
' Resumo: Obtem o valor da propriedade TopNode.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get TopNode() As String
    TopNode = GetTree.TopNode
End Property

'******************************************************************************
' Nome do elemento: GetTreeType (Property Get)
' Resumo: Obtem o valor da propriedade GetTreeType.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get GetTreeType() As Long
    GetTreeType = GetTree.GetTreeType
End Property




