VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fwGuiTableControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' Nome do elemento: fwGuiTableControl (Classe)
' Resumo: Wrapper flexível para GuiTableControl com find-by-ID ou find-by-Name, sem referência obsoletas.
'******************************************************************************

Option Explicit

' Late binding e contexto
Private pSession    As Object       ' GuiSession
Private pWindow     As Object       ' GuiMainWindow (ou subscreen)
Private pTableId    As String       ' ID explícito passado pelo usuário
Private pTechName   As String       ' Nome técnico para search-by-name
Private pTable      As Object       ' Instância atual do GuiTableControl
Private pColIndex   As Object       ' Scripting.Dictionary (colName ? índice)

'******************************************************************************
' Nome do elemento: Init (Sub)
' Resumo: Injeta sessão e janela, opcionalmente já associando o objeto de tabela.
' @param sessionObj — Objeto GuiSession usado para localizar controle.
' @param windowObj — Objeto GuiMainWindow ou subscreen que contém o controle de tabela.
' @param tableObj — Objeto GuiTableControl opcional para inicialização de cache.
' @remarks Inicializa dicionário de coluna se tableObj informado.
'******************************************************************************
Public Sub Init( _
    ByVal sessionObj As Object, _
    ByVal windowObj As Object, _
    Optional ByVal tableObj As Object _
)
    Set pSession = sessionObj
    Set pWindow = windowObj
    Set pTable = tableObj
    pTableId = ""
    pTechName = ""
    Set pColIndex = CreateObject("Scripting.Dictionary")
    If Not tableObj Is Nothing Then RefreshColumns
End Sub

'******************************************************************************
' Nome do elemento: ID (Property Let)
' Resumo: Define o valor da propriedade ID.
' @param tableId - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Let ID(ByVal tableId As String)
    pTableId = tableId
    pTechName = ""
    Set pTable = Nothing
    pColIndex.RemoveAll
End Property

'******************************************************************************
' Nome do elemento: ID (Property Get)
' Resumo: Obtem o valor da propriedade ID.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get ID() As String
    If Len(pTableId) > 0 Then
        ID = pTableId
    Else
        ID = GetTable().ID
    End If
End Property

'******************************************************************************
' Nome do elemento: Name (Property Let)
' Resumo: Define o valor da propriedade Name.
' @param techName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Let Name(ByVal techName As String)
    Dim realName As String
    If LCase$(Left$(techName, 3)) = "tbl" Then
        realName = Mid$(techName, 4)
    Else
        realName = techName
    End If
    pTechName = realName
    pTableId = ""
    Set pTable = Nothing
    pColIndex.RemoveAll
End Property

'******************************************************************************
' Nome do elemento: Name (Property Get)
' Resumo: Obtem o valor da propriedade Name.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Name() As String
    Name = pTechName
End Property

'******************************************************************************
' Nome do elemento: Self (Property Get)
' Resumo: Obtem o valor da propriedade Self.
' @param sTech - Parâmetro da rotina.
' @return fwGuiTableControl - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Self(ByVal sTech As String) As fwGuiTableControl
    Me.Name = sTech
    Set Self = Me
End Property

'******************************************************************************
' Nome do elemento: GetTable (Function)
' Resumo: Retorna o resultado calculado por GetTable.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Function GetTable() As Object
    On Error GoTo ErrHandler

    ' Reutiliza cache se o objeto ainda estiver vivo no frontend
    If Not pTable Is Nothing Then
        Dim cached As Object, tmpId As String
        Set cached = pTable
        On Error Resume Next
        tmpId = cached.ID
        If Err.Number = 0 Then
            On Error GoTo ErrHandler
            Set GetTable = cached
            Exit Function
        End If
        On Error GoTo ErrHandler
        Err.Clear
        Set pTable = Nothing
    End If

    If Len(pTableId) > 0 Then
        Set GetTable = pSession.findbyid(pTableId)
    ElseIf Len(pTechName) > 0 Then
        Set GetTable = pWindow.findByName(pTechName, "GuiTableControl")
    Else
        Err.Raise vbObjectError + 2001, , "Nenhum ID ou Name definido para o wrapper."
    End If
    Set pTable = GetTable
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "fwGuiTableControl.GetTable", Err.Description
End Function

'******************************************************************************
' Nome do elemento: RefreshColumns (Sub)
' Resumo: Executa a rotina RefreshColumns.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub RefreshColumns()
    Dim tbl As Object: Set tbl = GetTable()
    Dim i As Long, hdr As String, tech As String, key As String
    Set pColIndex = CreateObject("Scripting.Dictionary")
    For i = 0 To tbl.Columns.Count - 1
        hdr = Trim$(tbl.Columns.item(i).title)
        tech = GetColumnTechnicalName(i)
        key = IIf(Len(hdr) > 0, hdr, tech)
        If pColIndex.Exists(LCase$(key)) Then
            Dim oldIdx As Long, oldTech As String
            oldIdx = pColIndex(LCase$(key))
            oldTech = GetColumnTechnicalName(oldIdx)
            pColIndex.Remove LCase$(key)
            pColIndex(LCase$(oldTech & hdr)) = oldIdx
            pColIndex(LCase$(tech & hdr)) = i
        Else
            pColIndex(LCase$(key)) = i
        End If
    Next i
End Sub

'******************************************************************************
' Nome do elemento: GetColumnTechnicalName (Function)
' Resumo: Retorna o resultado calculado por GetColumnTechnicalName.
' @param colIdx - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function GetColumnTechnicalName(ByVal colIdx As Long) As String
    Dim tbl As Object: Set tbl = GetTable()
    Dim visRows As Long: visRows = tbl.VisibleRowCount
    Dim r As Long, tmpCell As Object, fullId As String
    Dim dashP As Long, brP As Long
    For r = 0 To visRows - 1
        On Error Resume Next
        Set tmpCell = tbl.GetCell(r, colIdx)
        On Error GoTo 0
        If Not tmpCell Is Nothing Then
            fullId = tmpCell.ID
            dashP = InStrRev(fullId, "-")
            brP = InStrRev(fullId, "[")
            If dashP > 0 And brP > dashP Then
                GetColumnTechnicalName = Mid$(fullId, dashP + 1, brP - dashP - 1)
            Else
                GetColumnTechnicalName = ""
            End If
            Exit Function
        End If
    Next r
    GetColumnTechnicalName = "Col" & colIdx
End Function

'******************************************************************************
' Nome do elemento: GetColIdx (Function)
' Resumo: Retorna o resultado calculado por GetColIdx.
' @param colName - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Function GetColIdx(ByVal colName As String) As Long
    Dim key As String: key = LCase$(Trim$(colName))
    If Not pColIndex.Exists(key) Then
        RefreshColumns
        If Not pColIndex.Exists(key) Then
            Err.Raise vbObjectError + 1002, , _
              "Coluna '" & colName & "' não encontrada."
        End If
    End If
    GetColIdx = pColIndex(key)
End Function

'******************************************************************************
' Nome do elemento: CellExists (Function)
' Resumo: Verifica se a célula em uma linha e coluna especificadas existe no controle.
' @param rowIdx — Índice de linha (0-based).
' @param colName — Rótulo ou nome técnico da coluna.
' @return Boolean — True se a célula existir; False caso contrário.
'******************************************************************************
Public Function CellExists( _
    ByVal rowIdx As Long, _
    ByVal colName As String _
) As Boolean
    On Error GoTo ErrHandler
    Dim tbl       As Object: Set tbl = GetTable()
    Dim cIdx      As Long:   cIdx = GetColIdx(colName)
    Dim totalRows As Long:   totalRows = tbl.rowCount
    If rowIdx < 0 Or rowIdx >= totalRows Then
        CellExists = False: Exit Function
    End If
    Dim relIdx As Long: relIdx = rowIdx Mod tbl.VisibleRowCount
    Dim tmpCell As Object
    On Error Resume Next
      Set tmpCell = tbl.GetCell(relIdx, cIdx)
    If Err.Number <> 0 Or tmpCell Is Nothing Then
        CellExists = False
    Else
        CellExists = True
    End If
    Exit Function
ErrHandler:
    CellExists = False
End Function

'******************************************************************************
' Nome do elemento: CellText (Function)
' Resumo: Retorna o texto da célula em uma linha e coluna especificadas.
' @param rowIdx — Índice de linha (0-based).
' @param colName — Rótulo ou nome técnico da coluna.
' @return String — Texto contido na célula.
' @raises vbObjectError+1005 — se a célula for inválida.
'******************************************************************************
Public Function CellText( _
    ByVal rowIdx As Long, _
    ByVal colName As String _
) As String
    If Not CellExists(rowIdx, colName) Then
        Err.Raise vbObjectError + 1005, , _
          "Célula (" & rowIdx & ",""" & colName & """) inválida."
    End If
    Dim tbl    As Object: Set tbl = GetTable()
    Dim cIdx   As Long:   cIdx = GetColIdx(colName)
    Dim relIdx As Long:   relIdx = rowIdx Mod tbl.VisibleRowCount
    Dim cell   As Object: Set cell = tbl.GetCell(relIdx, cIdx)
    CellText = cell.Text
End Function

'******************************************************************************
' Nome do elemento: DumpColumns (Sub)
' Resumo: Executa a rotina DumpColumns.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub DumpColumns()
    If pColIndex Is Nothing Or pColIndex.Count = 0 Then RefreshColumns
    Dim key As Variant
    Debug.Print "---- fwGuiTableControl Columns Dump ----"
    For Each key In pColIndex.keys
        Dim idx As Long: idx = pColIndex(key)
        Debug.Print "Label: """ & key & """", _
                    " Índice: " & idx, _
                    " Técnico: """ & GetColumnTechnicalName(idx) & """"
    Next
    Debug.Print "----------------------------------------"
End Sub

'******************************************************************************
' Nome do elemento: VerticalScrollbar (Property Get)
' Resumo: Obtem o valor da propriedade VerticalScrollbar.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get VerticalScrollbar() As Object
    Dim tbl As Object
    Set tbl = GetTable()
    Set VerticalScrollbar = tbl.VerticalScrollbar
End Property

'******************************************************************************
' Nome do elemento: rowCount (Property Get)
' Resumo: Obtem o valor da propriedade rowCount.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get rowCount() As Long
    Dim tbl As Object
    Set tbl = GetTable()
    rowCount = tbl.rowCount
End Property

'******************************************************************************
' Nome do elemento: VisibleRowCount (Property Get)
' Resumo: Obtem o valor da propriedade VisibleRowCount.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get VisibleRowCount() As Long
    Dim tbl As Object
    Set tbl = GetTable()
    VisibleRowCount = tbl.VisibleRowCount
End Property

'******************************************************************************
' Nome do elemento: ExistsColumn (Function)
' Resumo: Verifica se uma coluna com o rótulo ou nome técnico especificado existe.
' @param colName — Rótulo ou nome técnico da coluna.
' @return Boolean — True se a coluna existir; False caso contrário.
'******************************************************************************
Public Function ExistsColumn( _
    ByVal colName As String _
) As Boolean
    Dim key As String
    key = LCase$(Trim$(colName))
    If pColIndex Is Nothing Or pColIndex.Count = 0 Then
        RefreshColumns
    End If
    ExistsColumn = pColIndex.Exists(key)
End Function

'******************************************************************************
' Nome do elemento: DumpState (Sub)
' Resumo: Executa a rotina DumpState.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub DumpState()
    Dim tbl As Object: Set tbl = GetTable()
    tbl.DumpState
End Sub

'******************************************************************************
' Nome do elemento: Visualize (Sub)
' Resumo: Executa a rotina Visualize.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub Visualize()
    Dim tbl As Object: Set tbl = GetTable()
    tbl.Visualize
End Sub

'******************************************************************************
' Nome do elemento: ConfigureLayout (Sub)
' Resumo: Executa a rotina ConfigureLayout.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub ConfigureLayout()
    Dim tbl As Object: Set tbl = GetTable()
    tbl.ConfigureLayout
End Sub

'******************************************************************************
' Nome do elemento: DeselectAllColumns (Sub)
' Resumo: Executa a rotina DeselectAllColumns.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub DeselectAllColumns()
    Dim tbl As Object: Set tbl = GetTable()
    tbl.DeselectAllColumns
End Sub

'******************************************************************************
' Nome do elemento: GetAbsoluteRow (Function)
' Resumo: Retorna o resultado calculado por GetAbsoluteRow.
' @param index - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetAbsoluteRow(ByVal index As Long) As Object
    Dim tbl As Object: Set tbl = GetTable()
    On Error Resume Next
    Set GetAbsoluteRow = tbl.GetAbsoluteRow(index)
    On Error GoTo 0
End Function

'******************************************************************************
' Nome do elemento: GetCell (Function)
' Resumo: Retorna o resultado calculado por GetCell.
' @param Row - Parâmetro da rotina.
' @param Column - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetCell(ByVal Row As Long, ByVal Column As Long) As Object
    Dim tbl As Object: Set tbl = GetTable()
    On Error Resume Next
    Set GetCell = tbl.GetCell(Row, Column)
    On Error GoTo 0
End Function

'******************************************************************************
' Nome do elemento: ReorderTable (Sub)
' Resumo: Executa a rotina ReorderTable.
' @param Permutation - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub ReorderTable(ByVal Permutation As String)
    Dim tbl As Object: Set tbl = GetTable()
    tbl.ReorderTable Permutation
End Sub

'******************************************************************************
' Nome do elemento: SelectAllColumns (Sub)
' Resumo: Executa a rotina SelectAllColumns.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub SelectAllColumns()
    Dim tbl As Object: Set tbl = GetTable()
    tbl.SelectAllColumns
End Sub

'******************************************************************************
' Nome do elemento: Parent (Property Get)
' Resumo: Obtem o valor da propriedade Parent.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Parent() As Object
    Dim tbl As Object: Set tbl = GetTable()
    Set Parent = tbl.Parent
End Property

'******************************************************************************
' Nome do elemento: TypeName (Property Get)
' Resumo: Obtem o valor da propriedade TypeName.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get TypeName() As String
    Dim tbl As Object: Set tbl = GetTable()
    TypeName = tbl.Type
End Property

'******************************************************************************
' Nome do elemento: TypeAsNumber (Property Get)
' Resumo: Obtem o valor da propriedade TypeAsNumber.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get TypeAsNumber() As Long
    Dim tbl As Object: Set tbl = GetTable()
    TypeAsNumber = tbl.TypeAsNumber
End Property

'******************************************************************************
' Nome do elemento: Changeable (Property Get)
' Resumo: Obtem o valor da propriedade Changeable.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Changeable() As Boolean
    Dim tbl As Object: Set tbl = GetTable()
    Changeable = tbl.Changeable
End Property

'******************************************************************************
' Nome do elemento: DefaultTooltip (Property Get)
' Resumo: Obtem o valor da propriedade DefaultTooltip.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get DefaultTooltip() As String
    Dim tbl As Object: Set tbl = GetTable()
    DefaultTooltip = tbl.DefaultTooltip
End Property

'******************************************************************************
' Nome do elemento: IconName (Property Get)
' Resumo: Obtem o valor da propriedade IconName.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get IconName() As String
    Dim tbl As Object: Set tbl = GetTable()
    IconName = tbl.IconName
End Property

'******************************************************************************
' Nome do elemento: IsSymbolFont (Property Get)
' Resumo: Obtem o valor da propriedade IsSymbolFont.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get IsSymbolFont() As Boolean
    Dim tbl As Object: Set tbl = GetTable()
    IsSymbolFont = tbl.IsSymbolFont
End Property

'******************************************************************************
' Nome do elemento: Modified (Property Get)
' Resumo: Obtem o valor da propriedade Modified.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Modified() As Boolean
    Dim tbl As Object: Set tbl = GetTable()
    Modified = tbl.Modified
End Property

'******************************************************************************
' Nome do elemento: text (Property Get)
' Resumo: Obtem o valor da propriedade text.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Text() As String
    Dim tbl As Object: Set tbl = GetTable()
    Text = tbl.Text
End Property

'******************************************************************************
' Nome do elemento: Tooltip (Property Get)
' Resumo: Obtem o valor da propriedade Tooltip.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Tooltip() As String
    Dim tbl As Object: Set tbl = GetTable()
    Tooltip = tbl.Tooltip
End Property

'******************************************************************************
' Nome do elemento: Children (Property Get)
' Resumo: Obtem o valor da propriedade Children.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Children() As Object
    Dim tbl As Object: Set tbl = GetTable()
    Set Children = tbl.Children
End Property

'******************************************************************************
' Nome do elemento: ColSelectMode (Property Get)
' Resumo: Obtem o valor da propriedade ColSelectMode.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get ColSelectMode() As Long
    Dim tbl As Object: Set tbl = GetTable()
    ColSelectMode = tbl.ColSelectMode
End Property

'******************************************************************************
' Nome do elemento: Columns (Property Get)
' Resumo: Obtem o valor da propriedade Columns.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Columns() As Object
    Dim tbl As Object: Set tbl = GetTable()
    Set Columns = tbl.Columns
End Property

'******************************************************************************
' Nome do elemento: CurrentCol (Property Get)
' Resumo: Obtem o valor da propriedade CurrentCol.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get CurrentCol() As Long
    Dim tbl As Object: Set tbl = GetTable()
    CurrentCol = tbl.CurrentCol
End Property

'******************************************************************************
' Nome do elemento: CurrentRow (Property Get)
' Resumo: Obtem o valor da propriedade CurrentRow.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get CurrentRow() As Long
    Dim tbl As Object: Set tbl = GetTable()
    CurrentRow = tbl.CurrentRow
End Property

'******************************************************************************
' Nome do elemento: HorizontalScrollbar (Property Get)
' Resumo: Obtem o valor da propriedade HorizontalScrollbar.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get HorizontalScrollbar() As Object
    Dim tbl As Object: Set tbl = GetTable()
    Set HorizontalScrollbar = tbl.HorizontalScrollbar
End Property

'******************************************************************************
' Nome do elemento: rows (Property Get)
' Resumo: Obtem o valor da propriedade rows.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get rows() As Object
    Dim tbl As Object: Set tbl = GetTable()
    Set rows = tbl.rows
End Property

'******************************************************************************
' Nome do elemento: RowSelectMode (Property Get)
' Resumo: Obtem o valor da propriedade RowSelectMode.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get RowSelectMode() As Long
    Dim tbl As Object: Set tbl = GetTable()
    RowSelectMode = tbl.RowSelectMode
End Property

'******************************************************************************
' Nome do elemento: TableFieldName (Property Get)
' Resumo: Obtem o valor da propriedade TableFieldName.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get TableFieldName() As String
    Dim tbl As Object: Set tbl = GetTable()
    TableFieldName = tbl.TableFieldName
End Property



