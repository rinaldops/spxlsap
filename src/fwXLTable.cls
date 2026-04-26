VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fwXLTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' Nome do elemento: fwXLTable (Classe)
' Resumo:
'   Encapsula operações sobre ListObjects do Excel, oferecendo navegação,
'   seleção, formatação e manipulação de dados em tabelas XL estruturadas.
'******************************************************************************
Option Explicit

Private Const STAGING_SHEET_NAME As String = "Staging_fwXLSPConn"

'--- estado interno do ListObject
Private pWorksheet      As Worksheet
Private pListObject     As ListObject
Private pNavColumn      As String
Private pTraceOn        As Boolean
Private pHelpers        As fwHelpers

' offsets/contagens relativas ao DataBody
Private pFirstDataRow     As Long
Private pFirstDataColumn  As Long
Private pTableColsCount   As Long
Private pTableRowsCount   As Long
Private pHeaderRowIndex   As Long

'******************************************************************************
' Nome do elemento: Init (Sub)
' Resumo: Executa a rotina Init.
' @param tableName - Parâmetro da rotina.
' @param navCol - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Sub Init(ByVal tableName As String, Optional ByVal navCol As String = "ID")
    On Error GoTo ErrHandler
    
    Trace "Init ENTER; Procurando pela tabela '" & tableName & "'"
        
    If Len(Trim$(tableName)) = 0 Then
        Err.Raise vbObjectError + 611, "fwXLTable.Init", "O parâmetro 'tableName' é obrigatório."
    End If

    Dim lo As ListObject
    Dim sht As Worksheet
    Dim found As Boolean: found = False

    For Each sht In ActiveWorkbook.Worksheets
        For Each lo In sht.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set pListObject = lo
                Set pWorksheet = sht
                found = True
                Exit For
            End If
        Next lo
        
        If found Then Exit For
    Next sht

    If Not found Then
        Err.Raise vbObjectError + 612, "fwXLTable.Init", "Tabela '" & tableName & "' não foi encontrada em nenhuma planilha do ActiveWorkbook."
    End If
    
    pNavColumn = navCol
    Call UpdateTableMetrics
    Trace "Init: Worksheet=" & pWorksheet.Name & " TableName=" & pListObject.Name & " key=" & pNavColumn
    Exit Sub
    
ErrHandler:
    Trace "Init ERROR: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "fwXLTable.Init", Err.Description
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
' Nome do elemento: Self (Property Get)
' Resumo: Obtem o valor da propriedade Self.
' @return ListObject - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Self() As ListObject
    Set Self = pListObject
End Property

'******************************************************************************
' Nome do elemento: Worksheet (Property Get)
' Resumo: Obtem o valor da propriedade Worksheet.
' @return Worksheet - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property

'******************************************************************************
' Nome do elemento: NavColumn (Property Get)
' Resumo: Obtem o valor da propriedade NavColumn.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get NavColumn() As String
    NavColumn = pNavColumn
End Property

'******************************************************************************
' Nome do elemento: UpdateTableMetrics (Sub)
' Resumo: Executa a rotina UpdateTableMetrics.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub UpdateTableMetrics()
    If pListObject Is Nothing Then
        pHeaderRowIndex = 0
        Exit Sub
    End If
    
    pHeaderRowIndex = pListObject.HeaderRowRange.Row
    
    If pListObject.DataBodyRange Is Nothing Then
        With pListObject.HeaderRowRange
            pFirstDataRow = .Row + 1
            pFirstDataColumn = .Columns(1).Column
        End With
        pTableRowsCount = 0
        pTableColsCount = pListObject.ListColumns.Count
    Else
        With pListObject.DataBodyRange
            pFirstDataRow = .Row
            pFirstDataColumn = .Column
            pTableRowsCount = .rows.Count
            pTableColsCount = .Columns.Count
        End With
    End If
    Trace "UpdateTableMetrics: pFirstDataRow=" & pFirstDataRow & " rowsCount=" & pTableRowsCount
End Sub

'******************************************************************************
' Nome do elemento: CellRel (Property Get)
' Resumo: Obtem o valor da propriedade CellRel.
' @param rowIndex - Parâmetro da rotina.
' @param colName - Parâmetro da rotina.
' @return Range - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Property Get CellRel(ByVal rowIndex As Long, ByVal colName As String) As Range
    Dim relCol As Long
    If pListObject Is Nothing Then Err.Raise vbObjectError + 620, "fwXLTable.Cell", "ListObject não inicializado."
    If rowIndex < 1 Or rowIndex > pListObject.ListRows.Count Then Err.Raise vbObjectError + 621, "fwXLTable.Cell", "Índice de linha fora do intervalo."
    relCol = SafeListColumnIndex(colName)
    If relCol = 0 Then Err.Raise vbObjectError + 622, "fwXLTable.Cell", "Coluna '" & colName & "' não encontrada."
    Set CellRel = pListObject.ListRows(rowIndex).Range.Cells(1, relCol)
End Property

'******************************************************************************
' Nome do elemento: CellAbs (Property Get)
' Resumo: Obtem o valor da propriedade CellAbs.
' @param rowNumber - Parâmetro da rotina.
' @param colName - Parâmetro da rotina.
' @return Range - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Property Get CellAbs(ByVal rowNumber As Long, ByVal colName As String) As Range
    Dim relRow As Long
    Dim relCol As Long
    
    If pListObject Is Nothing Then Err.Raise vbObjectError + 623, "fwXLTable.CellAbs", "ListObject não inicializado."
    If rowNumber <= pHeaderRowIndex Then Err.Raise vbObjectError + 624, "fwXLTable.CellAbs", "Linha fora do intervalo da tabela."
    
    relRow = rowNumber - pHeaderRowIndex
    If relRow < 1 Or relRow > pListObject.ListRows.Count Then
        Err.Raise vbObjectError + 624, "fwXLTable.CellAbs", "Linha fora do intervalo da tabela."
    End If
    
    relCol = SafeListColumnIndex(colName)
    If relCol = 0 Then Err.Raise vbObjectError + 622, "fwXLTable.Cell", "Coluna '" & colName & "' não encontrada."
    
    Set CellAbs = pListObject.ListRows(relRow).Range.Cells(1, relCol)
End Property

'******************************************************************************
' Nome do elemento: QtdRows (Property Get)
' Resumo: Obtem o valor da propriedade QtdRows.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get QtdRows() As Long
    If pListObject Is Nothing Then
        QtdRows = 0
    Else
        QtdRows = pListObject.ListRows.Count
    End If
End Property

'******************************************************************************
' Nome do elemento: QtdVisibleRows (Property Get)
' Resumo: Obtem o valor da propriedade QtdVisibleRows.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get QtdVisibleRows() As Long
    Dim rngVis As Range
    
    If pListObject Is Nothing Then Exit Property
    If pListObject.ListRows.Count = 0 Then Exit Property
    
    On Error Resume Next
    Set rngVis = pListObject.DataBodyRange.Columns(1).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If rngVis Is Nothing Then
        QtdVisibleRows = 0
    Else
        QtdVisibleRows = rngVis.Cells.Count
    End If
End Property

'******************************************************************************
' Nome do elemento: IdxCol (Function)
' Resumo: Retorna o resultado calculado por IdxCol.
' @param colName - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function idxCol(ByVal colName As String) As Long
    On Error Resume Next
    idxCol = pListObject.ListColumns(colName).index
    If Err.Number <> 0 Then Err.Raise vbObjectError + 631, "fwXLTable.IdxCol", "Coluna '" & colName & "' não encontrada."
End Function

'******************************************************************************
' Nome do elemento: IdxRow (Function)
' Resumo: Retorna o resultado calculado por IdxRow.
' @param rng - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function IdxRow(Optional ByVal rng As Range) As Long
    On Error GoTo ErrHandler
    Dim c As Range
    If rng Is Nothing Then Set c = ActiveCell Else Set c = rng
    
    ' Se a célula estiver fora do corpo da tabela, não levanta erro, retorna 0
    On Error Resume Next
    Dim isOutside As Boolean
    isOutside = (Intersect(c, pListObject.DataBodyRange) Is Nothing)
    If isOutside Or Err.Number <> 0 Then
        IdxRow = 0
        Err.Clear
        Exit Function
    End If
    On Error GoTo ErrHandler
    
    IdxRow = c.Row - pFirstDataRow + 1
    Exit Function

ErrHandler:
    IdxRow = 0
End Function

'******************************************************************************
' Nome do elemento: RangeRow (Function)
' Resumo: Retorna o resultado calculado por RangeRow.
' @param startCol - Parâmetro da rotina.
' @param endCol - Parâmetro da rotina.
' @return Range - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function RangeRow(Optional ByVal startCol As String = "", Optional ByVal endCol As String = "") As Range
    Dim rowNum As Long
    rowNum = ActiveCell.Row
    
    If pListObject Is Nothing Then Exit Function
    If rowNum <= pHeaderRowIndex Or rowNum > pHeaderRowIndex + pListObject.ListRows.Count Then
        Set RangeRow = Nothing
        Exit Function
    End If
    
    If Len(startCol) = 0 Then
        Dim firstCol As Long
        Dim lastCol As Long
        firstCol = pListObject.DataBodyRange.Columns(1).Column
        lastCol = pListObject.DataBodyRange.Columns(pListObject.DataBodyRange.Columns.Count).Column
        Set RangeRow = pWorksheet.Range(pWorksheet.Cells(rowNum, firstCol), pWorksheet.Cells(rowNum, lastCol))
    ElseIf Len(endCol) = 0 Then
        Set RangeRow = CellAbs(rowNum, startCol)
    Else
        Set RangeRow = pWorksheet.Range(CellAbs(rowNum, startCol), CellAbs(rowNum, endCol))
    End If
End Function

'******************************************************************************
' Nome do elemento: RangeCol (Function)
' Resumo: Retorna o resultado calculado por RangeCol.
' @param colunaNome - Parâmetro da rotina.
' @return Range - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function RangeCol(Optional ByVal colunaNome As String = "") As Range
    Dim rngDestino As Range
    
    If pListObject Is Nothing Then Err.Raise vbObjectError + 1001, "cListObject.RangeCol", "Tabela não definida"
    If pListObject.ListRows.Count = 0 Then Exit Function
    
    If Len(colunaNome) = 0 Then
        If Intersect(ActiveCell, pListObject.DataBodyRange) Is Nothing Then
            Err.Raise vbObjectError + 1002, "cListObject.RangeCol", "Célula ativa fora da tabela"
        End If
        Dim relCol As Long
        relCol = ActiveCell.Column - pListObject.Range.Columns(1).Column + 1
        Set rngDestino = pListObject.ListColumns(relCol).DataBodyRange
    Else
        On Error Resume Next
        Set rngDestino = pListObject.ListColumns(colunaNome).DataBodyRange
        On Error GoTo 0
        If rngDestino Is Nothing Then
            Err.Raise vbObjectError + 1003, "cListObject.RangeCol", "Coluna '" & colunaNome & "' não encontrada"
        End If
    End If
    
    On Error Resume Next
    Set RangeCol = rngDestino.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
End Function

'******************************************************************************
' Nome do elemento: firstRow (Function)
' Resumo: Retorna o resultado calculado por firstRow.
' @param colName - Parâmetro da rotina.
' @return Range - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function firstRow(Optional ByVal colName As String = "") As Range
    If pListObject Is Nothing Then
        Err.Raise vbObjectError + 604, "cListObjectsSP.FirstRow", _
                  "ListObject não inicializado. Chame Init primeiro."
    End If
    Dim visAreas As Range
    On Error Resume Next
        Set visAreas = pListObject.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If visAreas Is Nothing Then
        Set firstRow = Nothing
        Exit Function
    End If
    ' pega a primeira área de células visíveis
    Dim firstArea As Range: Set firstArea = visAreas.Areas(1)
    ' converte a linha absoluta para índice relativo de ListRows
    Dim relRow As Long: relRow = firstArea.Row - pFirstDataRow + 1
    If colName = "" Then
        Set firstRow = pListObject.ListRows(relRow).Range
    Else
        Set firstRow = CellRel(relRow, colName)
    End If
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "cListObjectsSP.FirstRow", Err.Description
End Function

'******************************************************************************
' Nome do elemento: lastRow (Function)
' Resumo: Retorna o resultado calculado por lastRow.
' @param colName - Parâmetro da rotina.
' @return Range - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function lastRow(Optional ByVal colName As String = "") As Range
    If pListObject Is Nothing Then
        Err.Raise vbObjectError + 605, "cListObjectsSP.LastRow", _
                  "ListObject não inicializado. Chame Init primeiro."
    End If
    Dim visAreas As Range
    On Error Resume Next
        Set visAreas = pListObject.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If visAreas Is Nothing Then
        Set lastRow = Nothing
        Exit Function
    End If
    ' pega a última área de células visíveis
    Dim lastArea As Range: Set lastArea = visAreas.Areas(visAreas.Areas.Count)
    ' converte a última linha absoluta para índice relativo
    Dim relRow As Long: relRow = lastArea.rows(lastArea.rows.Count).Row - pFirstDataRow + 1
    
    If colName = "" Then
        Set lastRow = pListObject.ListRows(relRow).Range
    Else
        Call idxCol(colName)
        Set lastRow = CellRel(relRow, colName)
    End If
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "cListObjectsSP.LastRow", Err.Description
End Function

'******************************************************************************
' Nome do elemento: ExistsNextRow (Function)
' Resumo: Retorna o resultado calculado por ExistsNextRow.
' @param vGo - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ExistsNextRow(Optional vGo As Boolean = False) As Boolean
    Dim currRel As Long, nextRel As Long
    Dim maxRows As Long: maxRows = pTableRowsCount
    Dim navIdx As Long: navIdx = Me.idxCol(pNavColumn)
    Dim lr As ListRow
    
    If pListObject Is Nothing Then
        ExistsNextRow = False
        Exit Function
    End If
    
    ' Índice relativo da linha atual
    currRel = Me.IdxRow(ActiveCell)
    If currRel < 1 Or currRel >= maxRows Then
        ExistsNextRow = False
        Exit Function
    End If
    
    ExistsNextRow = False
    ' Percorre da próxima até o fim da tabela
    For nextRel = currRel + 1 To maxRows
        Set lr = pListObject.ListRows(nextRel)
        With lr.Range.Cells(1, navIdx)
            If Not .EntireRow.Hidden Then
                If vGo Then .Activate
                ExistsNextRow = True
                Exit Function
            End If
        End With
    Next nextRel
End Function

'******************************************************************************
' Nome do elemento: GoToNextRow (Sub)
' Resumo: Executa a rotina GoToNextRow.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub GoToNextRow()
    Call ExistsNextRow(True)
End Sub

'******************************************************************************
' Nome do elemento: CurrentRowRange (Property Get)
' Resumo: Obtem o valor da propriedade CurrentRowRange.
' @return Range - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get CurrentRowRange() As Range
    Dim relRow As Long
    relRow = Me.IdxRow(ActiveCell)
    If relRow > 0 Then
        Set CurrentRowRange = pListObject.ListRows(relRow).Range
    End If
End Property

'******************************************************************************
' Nome do elemento: GetCurrentRowValue (Function)
' Resumo: Retorna o resultado calculado por GetCurrentRowValue.
' @param colName - Parâmetro da rotina.
' @return Variant - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetCurrentRowValue(ByVal colName As String) As Variant
    On Error GoTo ErrHandler
    Dim relRow As Long, relCol As Long
    
    relRow = Me.IdxRow(ActiveCell)
    If relRow = 0 Then Exit Function
    
    relCol = SafeListColumnIndex(colName)
    If relCol = 0 Then Exit Function

    GetCurrentRowValue = pListObject.DataBodyRange.Cells(relRow, relCol).value
    Exit Function
ErrHandler:
    GetCurrentRowValue = CVErr(xlErrNA)
End Function

'******************************************************************************
' Nome do elemento: SetCurrentRowValue (Sub)
' Resumo: Executa a rotina SetCurrentRowValue.
' @param colName - Parâmetro da rotina.
' @param value - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub SetCurrentRowValue(ByVal colName As String, ByVal value As Variant)
    On Error GoTo ErrHandler
    Dim relRow As Long, relCol As Long
    
    relRow = Me.IdxRow(ActiveCell)
    If relRow = 0 Then Exit Sub
    
    relCol = SafeListColumnIndex(colName)
    If relCol = 0 Then Exit Sub

    pListObject.DataBodyRange.Cells(relRow, relCol).value = value
    Exit Sub
ErrHandler:
    ' Ignora o erro se a coluna não for encontrada
End Sub

'******************************************************************************
' Nome do elemento: GetVisibleRows (Function)
' Resumo: Retorna o resultado calculado por GetVisibleRows.
' @return Collection - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetVisibleRows() As Collection
    Dim visibleRows As New Collection
    Dim body As Range
    Dim cell As Range
    Dim rowIdx As Long
    
    On Error Resume Next
    Set body = pListObject.DataBodyRange.Columns(1).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not body Is Nothing Then
        For Each cell In body.Cells
            rowIdx = cell.Row - pListObject.HeaderRowRange.Row
            visibleRows.Add pListObject.ListRows(rowIdx)
        Next cell
    End If
    
    Set GetVisibleRows = visibleRows
End Function

'******************************************************************************
' Nome do elemento: SafeListColumnIndex (Function)
' Resumo: Retorna o resultado calculado por SafeListColumnIndex.
' @param colName - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function SafeListColumnIndex(ByVal colName As String) As Long
    On Error Resume Next
    SafeListColumnIndex = pListObject.ListColumns(colName).index
    On Error GoTo 0
End Function

'******************************************************************************
' Nome do elemento: ActConditionalFormatting (Público Sub)
' Resumo: Aplica ou remove uma regra de formatação condicional via fórmula.
'
' @param ruleFormula   – String com a expressão lógica (sem o “=” inicial),
'                        ex: "$B1='2026'" ou "$B1=$C1".
' @param Optional colourRole As Variant – Long RGB a usar; se Empty, será remoção.
' @param Optional targetRange As Range  – Range a formatar; se Nothing, usa DataBodyRange.
' @param Optional cleanRoles  As Boolean – Se True, limpa as regras antes (ou remove a específica).
'
' @raises Erro vbObjectError+1001 se targetRange final for Nothing.
'******************************************************************************
Public Sub ActConditionalFormatting( _
        ByVal ruleFormula As String, _
        Optional ByVal colourRole As Long = -1, _
        Optional ByVal targetRange As Range, _
        Optional ByVal cleanRoles As Boolean = False _
    )
    On Error GoTo ErrHandler

    ' 1) resolve o intervalo alvo
    Dim rngApply As Range
    If targetRange Is Nothing Then
        If pListObject Is Nothing Then _
            Err.Raise vbObjectError + 1001, "cListObjectsSP.ActConditionalFormatting", "ListObject não inicializado."
        Set rngApply = pListObject.DataBodyRange
    Else
        Set rngApply = targetRange
    End If
    If rngApply Is Nothing Then _
        Err.Raise vbObjectError + 1001, "cListObjectsSP.ActConditionalFormatting", "targetRange não definido."

    ' 2) prepara a fórmula
    Dim formulaFull As String
    formulaFull = IIf(Left$(ruleFormula, 1) = "=", ruleFormula, "=" & ruleFormula)

    ' 3) MODO REMOÇÃO: só entra aqui se cor = -1 e cleanRoles=True
    If colourRole = -1 And cleanRoles Then
        Dim fc As FormatCondition
        For Each fc In rngApply.FormatConditions
            If fc.Type = xlExpression _
               And InStr(1, fc.Formula1, ruleFormula, vbTextCompare) > 0 Then
                fc.Delete
                Exit For
            End If
        Next fc
        Exit Sub
    End If

    ' 4) MODO ADIÇÃO: cor válida (<> -1)
    If colourRole <> -1 Then
        If cleanRoles Then Worksheet.Cells.FormatConditions.Delete

        Dim newFC As FormatCondition
        Set newFC = rngApply.FormatConditions.Add( _
                        Type:=xlExpression, _
                        operator:=xlEqual, _
                        Formula1:=formulaFull _
                    )
        With newFC
            .Interior.Color = colourRole
            .stopIfTrue = False
        End With
    End If

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "cListObjectsSP.ActConditionalFormatting", Err.Description
End Sub

'******************************************************************************
' Nome do elemento: FreezeFormatting (Público Sub)
' Resumo: Converte em cores estáticas a formatação condicional aplicada ao corpo da tabela
'         e limpa todas as regras de CF.
' @param Optional targetRange As Range – DataBodyRange da tabela se omitido.
'******************************************************************************
Public Sub FreezeFormatting( _
        Optional ByVal targetRange As Range _
    )
    On Error GoTo ErrHandler

    ' 1) valida e define intervalo alvo
    If targetRange Is Nothing Then
        If pListObject Is Nothing Then _
            Err.Raise vbObjectError + 1000, "cListObjectsSP.FreezeFormatting", _
                      "ListObject não inicializado."
        Set targetRange = pListObject.DataBodyRange
    End If

    ' 2) copia cor exibida (com CF) para preenchimento estático
    Dim cel As Range
    For Each cel In targetRange.Cells
        If cel.DisplayFormat.Interior.Color <> 16777215 Then   ' 16777215 => Sem cor
            cel.Interior.Color = cel.DisplayFormat.Interior.Color
        End If
    Next cel

    ' 3) remove regras de CF
    targetRange.FormatConditions.Delete

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "cListObjectsSP.FreezeFormatting", Err.Description
End Sub

'******************************************************************************
' Nome do elemento: ActClearData (Function)
' Resumo: Retorna o resultado calculado por ActClearData.
' @return Collection - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function ActClearData() As Collection
    Dim savedFilters As Collection
    Dim sheetWasProtected As Boolean
    Dim originalActiveSheet As Worksheet
    Const SENHA_PLANILHA As String = ""
    
    On Error GoTo ErrHandler
    
    If pWorksheet Is Nothing Or pListObject Is Nothing Then
        Err.Raise vbObjectError + 700, "fwXLTable.ActClearData", "Worksheet ou ListObject não inicializado."
    End If
    
    ' Salva a planilha atualmente ativa e ativa a nossa planilha de trabalho
    Set originalActiveSheet = activeSheet
    If Not pWorksheet Is activeSheet Then
        pWorksheet.Activate
        Trace "ActClearData: Planilha '" & pWorksheet.Name & "' ativada para garantir a operação."
    End If

    ' Etapa 1: Lidar com a Proteção da Planilha
    If pWorksheet.ProtectContents Then
        sheetWasProtected = True
        pWorksheet.Unprotect SENHA_PLANILHA
        Trace "ActClearData: Proteção da planilha removida temporariamente."
    End If
    
    ' Etapa 2: Salvar e Remover Filtros
    If pListObject.AutoFilter.FilterMode Then
        Set savedFilters = SaveFilters()
        pListObject.AutoFilter.ShowAllData
        Trace "ActClearData: Filtro salvo e removido para garantir a limpeza."
    End If
    
    ' --- INÍCIO DA ESTRATÉGIA "DELETAR XLUP" (Anti-Lixo) ---
    ' Etapa 3: Executar a Limpeza
    
    ' Só executa se houver linhas para limpar
    If pListObject.ListRows.Count > 0 Then
        Application.EnableEvents = False
        
        ' Este método deleta todas as linhas do DataBodyRange e
        ' move as células abaixo para cima, sem deixar "lixo".
        ' A tabela é automaticamente redimensionada para 0 linhas de dados.
        pListObject.DataBodyRange.Delete xlUp
        
        Application.EnableEvents = True
    End If
    ' --- FIM DA ESTRATÉGIA ---
    
    Call UpdateTableMetrics
    Trace "ActClearData OK; Tabela limpa com sucesso (0 linhas)."
    
    GoTo RestoreAndExit

ErrHandler:
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    Dim errSource As String: errSource = "fwXLTable.ActClearData"
    
RestoreAndExit:
    ' Etapa Final: Restaura o estado original da planilha
    Application.EnableEvents = True
    
    If sheetWasProtected Then
        On Error Resume Next
        pWorksheet.Protect SENHA_PLANILHA
        Trace "ActClearData: Proteção da planilha restaurada."
        On Error GoTo 0
    End If
    
    If Not originalActiveSheet Is Nothing Then
        If originalActiveSheet.Name <> pWorksheet.Name Then
            On Error Resume Next
            originalActiveSheet.Activate
            On Error GoTo 0
        End If
    End If

    ' Retorna o estado do filtro para a rotina chamadora
    Set ActClearData = savedFilters
    
    ' Se um erro ocorreu, ele é levantado agora, após a limpeza
    If errNum <> 0 Then
        Err.Raise errNum, errSource, errDesc
    End If
    
End Function

'******************************************************************************
' Nome do elemento: ApplySavedFilters (Sub)
' Resumo: Executa a rotina ApplySavedFilters.
' @param savedFilters - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub ApplySavedFilters(ByVal savedFilters As Collection)
    Dim sheetWasProtected As Boolean
    ' ATENÇÃO: Se suas planilha são protegidas com senha, insira-a aqui.
    Const SENHA_PLANILHA As String = ""
    On Error GoTo ErrHandler
    
    If savedFilters Is Nothing Then Exit Sub ' Nada a fazer
    If pListObject Is Nothing Then Exit Sub ' Tabela não existe
    ' Se não há dados, não é possível aplicar filtros
    If pListObject.ListRows.Count = 0 Then Exit Sub
    
    ' Desprotege a planilha para poder aplicar o filtro
    If pWorksheet.ProtectContents Then
        sheetWasProtected = True
        pWorksheet.Unprotect SENHA_PLANILHA
    End If
    
    ' Chama a rotina interna que já temos para fazer o trabalho
    RestoreFilters savedFilters
    
    Trace "ApplySavedFilters: Filtro original restaurado com sucesso."

ErrHandler:
    ' Garante que a proteção seja restaurada no final, com sucesso ou erro
    If sheetWasProtected Then
        On Error Resume Next
        pWorksheet.Protect SENHA_PLANILHA
        On Error GoTo 0
    End If
    
    If Err.Number <> 0 Then
        ' Apenas loga o erro, não o levanta, para não interromper o fluxo principal
        ' por uma falha na restauração de um filtro (que é secundária).
        Trace "ApplySavedFilters: Ocorreu um erro ao restaurar o filtro. " & Err.Description
        Err.Clear
    End If
End Sub

'******************************************************************************
' Nome do elemento: XLLoadArray (Function)
' Resumo: Carrega dados e cabeçalhos no ListObject a partir de arrays em memória.
' @param getRowsArr - Matriz 2D com os registro a serem escritos.
' @param columnsParam - Conjunto de nomes de coluna que define a ordem das ListColumns.
' @param ClearData - Quando True remove dados existentes antes da carga.
' @param ClearTitle - Quando True redefine os cabeçalhos para os valores informados.
' @param ResizeTitle - Ajusta o header para o tamanho do array sem apagar dados.
' @return Variant - Dicionário com status, linhas processadas e mensagem de erro.
'******************************************************************************
Public Function XLLoadArray( _
    ByVal getRowsArr As Variant, _
    ByVal columnsParam As Variant, _
    Optional ByVal ClearData As Boolean = False, _
    Optional ByVal ClearTitle As Boolean = False, _
    Optional ByVal ResizeTitle As Boolean = False _
) As Variant
    Dim rep As Object: Set rep = CreateObject("Scripting.Dictionary")
    rep.Add "status", "error": rep.Add "processedRows", 0

    On Error GoTo ErrHandler
    Trace "XLLoadArray ENTER (vFinal - Parâmetros Independentes)"
    If pListObject Is Nothing Then Err.Raise vbObjectError + 9600, "XLLoadArray", "ListObject não inicializado."

    EnsureHelpers

    ' 1. Normaliza os campos de origem (necessário para todas as operações)
    Dim sourceFields() As String
    sourceFields = Me.helpers.NormalizeFieldsParam(columnsParam)
    Dim arrFieldCount As Long: arrFieldCount = UBound(sourceFields) - LBound(sourceFields) + 1


    ' --- INÍCIO DA LÓGICA DE LIMPEZA E REDIMENSIONAMENTO ---

    ' 2. Lida com o ESQUEMA (Títulos)
    ' ClearTitle tem precedência sobre ResizeTitle
    If ClearTitle Then
        Trace "XLLoadArray: ClearTitle=True. Redefinindo colunas..."
        Application.EnableEvents = False
        ' Usa a função robusta que renomeia, adiciona e remove coluna
        ' para corresponder exatamente ao array 'sourceFields', sem destruir
        ' o objeto ListObject.
        Call Me.EnsureColumnsInOrder(sourceFields)
        Application.EnableEvents = True
        Trace "XLLoadArray: Esquema de colunas redefinido."
        
    ElseIf ResizeTitle Then
        Trace "XLLoadArray: ResizeTitle=True. Verificando colunas ausentes..."
        Dim existingCols As Object
        Set existingCols = CreateObject("Scripting.Dictionary")
        existingCols.CompareMode = vbTextCompare
        Dim lc As ListColumn
        For Each lc In pListObject.ListColumns
            existingCols.Add lc.Name, True
        Next lc
        
        Dim fieldName As String
        Dim i As Long
        For i = LBound(sourceFields) To UBound(sourceFields)
            fieldName = sourceFields(i)
            If Not existingCols.Exists(fieldName) Then
                Application.EnableEvents = False
                pListObject.ListColumns.Add
                pListObject.ListColumns(pListObject.ListColumns.Count).Name = fieldName
                Application.EnableEvents = True
                Trace "XLLoadArray: Coluna '" & fieldName & "' adicionada."
            End If
        Next i
    End If

    ' 3. Lida com os DADOS (Linhas)
    If ClearData Then
        Trace "XLLoadArray: ClearData=True. Limpando dados existentes..."
        ' Chama a nova versão robusta "Anti-Lixo"
        Call Me.ActClearData
        Trace "XLLoadArray: Dados limpos."
    End If

    ' Atualiza as métricas da tabela APÓS as mudanças de esquema/dados
    Call UpdateTableMetrics
    
    ' 4. Lógica de Carregamento (O restante da função original)
    If Me.helpers.IsVarEmpty(getRowsArr) Then
        Trace "XLLoadArray: Array de dados vazio. Nada a carregar."
        rep("status") = "ok"
        Set XLLoadArray = rep
        Exit Function
    End If

    ' Lógica robusta para obter a contagem de linhas e limites
    Dim arrRowCount As Long
    Dim is2D As Boolean
    Dim arrLBound1 As Long, arrUBound1 As Long, arrLBound2 As Long, arrUBound2 As Long
    On Error Resume Next
    arrUBound2 = UBound(getRowsArr, 2)
    is2D = (Err.Number = 0)
    Err.Clear
    On Error GoTo ErrHandler

    If is2D Then
        arrLBound1 = LBound(getRowsArr, 1)
        arrUBound1 = UBound(getRowsArr, 1)
        arrLBound2 = LBound(getRowsArr, 2)
        arrUBound2 = UBound(getRowsArr, 2)
        arrRowCount = arrUBound2 - arrLBound2 + 1
        If (arrUBound1 - arrLBound1 + 1) <> arrFieldCount Then
            Err.Raise vbObjectError + 9601, "XLLoadArray", "Disparidade headers (" & arrFieldCount & ") vs dados 2D (" & (arrUBound1 - arrLBound1 + 1) & ")."
        End If
    Else ' Array 1D
        arrLBound1 = LBound(getRowsArr)
        arrUBound1 = UBound(getRowsArr)
        arrLBound2 = 0: arrUBound2 = 0
        arrRowCount = 1
         If (arrUBound1 - arrLBound1 + 1) <> arrFieldCount Then
            Err.Raise vbObjectError + 9602, "XLLoadArray", "Disparidade headers (" & arrFieldCount & ") vs dados 1D (" & (arrUBound1 - arrLBound1 + 1) & ")."
        End If
    End If

    If arrRowCount = 0 Then
        Trace "XLLoadArray: Array contém 0 linhas."
        rep("status") = "ok"
        Set XLLoadArray = rep
        Exit Function
    End If

    ' --- LÓGICA DE CARREGAMENTO (APPEND) ---
    Dim outArr As Variant
    Dim destColCount As Long: destColCount = pListObject.ListColumns.Count
    ReDim outArr(1 To arrRowCount, 1 To destColCount) ' Array de saída dimensionado para a tabela destino

    ' Mapeia header de entrada -> índice da coluna na tabela destino
    Dim colMap As Object: Set colMap = CreateObject("Scripting.Dictionary")
    colMap.CompareMode = vbTextCompare
    Dim destColIndex As Long
    For i = LBound(sourceFields) To UBound(sourceFields)
        On Error Resume Next
        destColIndex = 0
        destColIndex = pListObject.ListColumns(sourceFields(i)).index
        On Error GoTo ErrHandler
        If destColIndex > 0 Then
            colMap(i) = destColIndex
        Else
             Trace "XLLoadArray: Aviso - Coluna header '" & sourceFields(i) & "' não encontrada na tabela destino. Dados ignorados."
        End If
    Next i

    ' Transpõe manualmente e mapeia os dados para outArr
    Dim r As Long, c As Long
    Dim sourceColIndex As Long
    For r = 1 To arrRowCount ' Linhas do array de SAÍDA (1-based)
        For c = LBound(sourceFields) To UBound(sourceFields) ' Coluna do array de ENTRADA
            sourceColIndex = c
            If colMap.Exists(sourceColIndex) Then
                destColIndex = colMap(sourceColIndex)
                If is2D Then
                    outArr(r, destColIndex) = getRowsArr(sourceColIndex + arrLBound1 - LBound(sourceFields), r + arrLBound2 - 1)
                Else
                    outArr(r, destColIndex) = getRowsArr(sourceColIndex + arrLBound1 - LBound(sourceFields))
                End If
            End If
        Next c
    Next r

    ' Escreve o array transposto e mapeado no Excel
    If arrRowCount > 0 Then
        Application.EnableEvents = False
        On Error GoTo ErrHandlerWrite ' Handler específico para a escrita
        
        Dim targetRange As Range
        
        ' Se ClearData=True,         ' pListObject.ListRows.Count será 0.
        If pListObject.ListRows.Count = 0 Then ' Se a tabela está completamente vazia (só header)
            Trace "XLLoadArray: Tabela vazia (0 linhas). Redimensionando e escrevendo dados..."
            ' Redimensiona para comportar EXATAMENTE os novos dados
            If arrRowCount = 1 And pListObject.DataBodyRange Is Nothing Then
                pListObject.ListRows.Add
            Else
                pListObject.Resize pListObject.HeaderRowRange.Resize(arrRowCount + 1, destColCount)
            End If
            Set targetRange = pListObject.DataBodyRange
        Else ' Se a tabela já tem linhas (Modo Append real)
            Trace "XLLoadArray: Tabela não vazia. Adicionando dados no final..."
            ' Adiciona UMA nova linha e pega o range inicial dela
            Set targetRange = pListObject.ListRows.Add.Range
            ' Redimensiona esse range para comportar todas as novas linhas
            Set targetRange = targetRange.Resize(arrRowCount, destColCount)
        End If
        
        If Not targetRange Is Nothing Then
            targetRange.value = outArr ' Escreve o bloco de dados
            Trace "XLLoadArray: Dados escritos no range " & targetRange.Address
        Else
             Trace "XLLoadArray: ERRO - Não foi possível determinar o range de destino para escrita."
             Err.Raise vbObjectError + 9603, "XLLoadArray", "Falha ao obter range de destino."
        End If

        Application.EnableEvents = True
    End If

    Call UpdateTableMetrics ' Atualiza métricas após carregar

    rep("status") = "ok": rep("processedRows") = arrRowCount
    Set XLLoadArray = rep

    Trace "XLLoadArray OK; processed=" & arrRowCount
    Exit Function

ErrHandlerWrite:
    ' Erro durante a escrita dos dados
    Application.EnableEvents = True ' Garante reabilitação
    Trace "XLLoadArray ERROR during Write: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "fwXLTable.XLLoadArray.Write", Err.Description ' Re-lança com source específico
    
ErrHandler:
    Application.EnableEvents = True ' Garante reabilitação de eventos
    Trace "XLLoadArray ERROR: " & Err.Number & " - " & Err.Description
    rep("errorMessage") = Err.Description
    Set XLLoadArray = rep
End Function

'******************************************************************************
' Nome do elemento: EnsureColumnsInOrder (Sub)
' Resumo: Executa a rotina EnsureColumnsInOrder.
' @param fields - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub EnsureColumnsInOrder(ByRef fields() As String)
    Dim i As Long, cntExisting As Long
    cntExisting = pListObject.ListColumns.Count
    Dim placeholders() As String
    If cntExisting > 0 Then
        ReDim placeholders(1 To cntExisting)
        For i = 1 To cntExisting
            placeholders(i) = "__TMP_COL_" & CStr(i)
            On Error Resume Next
            pListObject.ListColumns(i).Name = placeholders(i)
            On Error GoTo 0
        Next i
    End If
    Dim targetCols As Long: targetCols = UBound(fields) - LBound(fields) + 1
    If cntExisting < targetCols Then
        For i = 1 To targetCols - cntExisting
            pListObject.ListColumns.Add
        Next i
    ElseIf cntExisting > targetCols Then
        For i = pListObject.ListColumns.Count To targetCols + 1 Step -1
            pListObject.ListColumns(i).Delete
        Next i
    End If
    Dim idx As Long: idx = 0
    For i = LBound(fields) To UBound(fields)
        idx = idx + 1
        pListObject.ListColumns(idx).Name = fields(i)
    Next i
End Sub

'******************************************************************************
' Nome do elemento: CreateTableFromData (Function)
' Resumo: Retorna o resultado calculado por CreateTableFromData.
' @param tableName - Parâmetro da rotina.
' @param dataResultAsObject - Parâmetro da rotina.
' @return ListObject - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function CreateTableFromData(ByVal tableName As String, ByVal dataResultAsObject As Object) As ListObject
    On Error GoTo ErrHandler
    
    Dim stagingSheet As Worksheet, sheetName As String
    sheetName = "Staging_" & tableName & "_" & Format(now, "hhmmss")
    
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets(sheetName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True
    
    Set stagingSheet = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    stagingSheet.Name = sheetName
    stagingSheet.Visible = xlSheetVeryHidden
    
    Dim newLo As ListObject
    Set newLo = stagingSheet.ListObjects.Add(sourceType:=xlSrcRange, source:=stagingSheet.Range("A1"), XlListObjectHasHeaders:=xlYes)
    newLo.Name = tableName
    
    Trace "CreateTableFromData: Tabela temporária '" & tableName & "' criada na planilha '" & sheetName & "'."
    
    Dim loader As New fwXLTable
    loader.TraceOn = Me.TraceOn
    loader.Init tableName
    
    Dim dataArr As Variant, headersArr As Variant
    If Not dataResultAsObject Is Nothing Then
        If dataResultAsObject.Exists("Data") Then dataArr = dataResultAsObject("Data")
        If dataResultAsObject.Exists("Headers") Then headersArr = dataResultAsObject("Headers")
    End If
    
    loader.XLLoadArray dataArr, headersArr
    
    Set CreateTableFromData = newLo
    Exit Function

ErrHandler:
    Application.DisplayAlerts = True
    Trace "CreateTableFromData ERROR: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "fwXLTable.CreateTableFromData", Err.Description
End Function

'******************************************************************************
' Nome do elemento: DeleteTable (Sub)
' Resumo: Executa a rotina DeleteTable.
' @param tableName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Sub DeleteTable(ByVal tableName As String)
    On Error GoTo ErrHandler
    Dim sht As Worksheet, lo As ListObject, sheetToDelete As Worksheet
    For Each sht In ActiveWorkbook.Worksheets
        On Error Resume Next
        Set lo = sht.ListObjects(tableName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set sheetToDelete = sht
            Exit For
        End If
    Next sht
    
    If Not sheetToDelete Is Nothing Then
        Application.DisplayAlerts = False
        If ActiveWorkbook.Worksheets.Count > 1 Then
            If sheetToDelete.index = 1 Then
                ActiveWorkbook.Worksheets(2).Activate
            Else
                ActiveWorkbook.Worksheets(1).Activate
            End If
        End If
        sheetToDelete.Delete
        Application.DisplayAlerts = True
        Trace "DeleteTable: Planilha '" & sheetToDelete.Name & "' (com tabela '" & tableName & "') foi excluída."
    Else
        Trace "DeleteTable: Tabela '" & tableName & "' não encontrada para exclusão."
    End If
    Exit Sub
ErrHandler:
    Application.DisplayAlerts = True
    Trace "DeleteTable ERROR: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "fwXLTable.DeleteTable", Err.Description
End Sub

'******************************************************************************
' Nome do elemento: DeleteRowsByQuery (Function)
' Resumo: Retorna o resultado calculado por DeleteRowsByQuery.
' @param deleteSql - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function DeleteRowsByQuery(ByVal deleteSql As String) As Long
    Dim savedFilters As Collection
    On Error GoTo Cleanup
    Dim rowsAffected As Long
    
    Dim idColumnName As String
    Dim idColumnWasTemporary As Boolean
    Dim tempCol As ListColumn
    Dim idColIndex As Long
    Dim i As Long, j As Long, k As Long
    
    ' --- LÓGICA DE PRESERVAÇÃO DE FILTRO (INÍCIO) ---
    If pListObject.AutoFilter.FilterMode Then
        Set savedFilters = SaveFilters()
        pListObject.AutoFilter.ShowAllData
        Trace "DeleteRowsByQuery: Filtro salvo e removido para garantir a exclusão."
    End If
    
    ' 1. Verifica/cria a coluna de ID temporária.
    On Error Resume Next
    idColIndex = pListObject.ListColumns(Me.NavColumn).index
    On Error GoTo Cleanup
    
    If idColIndex = 0 Then
        idColumnWasTemporary = True
        idColumnName = "TEMP_ID_" & Format(Timer, "00000")
        Set tempCol = pListObject.ListColumns.Add
        tempCol.Name = idColumnName
        If pListObject.ListRows.Count > 0 Then
            Dim idRange As Range
            Set idRange = tempCol.DataBodyRange
            Dim idArr() As Variant
            ReDim idArr(1 To pListObject.ListRows.Count, 1 To 1)
            For k = 1 To pListObject.ListRows.Count
                idArr(k, 1) = k
            Next k
            idRange.value = idArr
        End If
    Else
        idColumnWasTemporary = False
        idColumnName = Me.NavColumn
    End If
    
    ' 2. Transforma o DELETE em um SELECT.
    Dim parsedSql As Object: Set parsedSql = Me.helpers.ParseDeleteSqlDetails(deleteSql)
    If parsedSql Is Nothing Then Exit Function
    Dim whereClause As String: whereClause = parsedSql("WhereClause")
    
    Dim selectSql As String
    selectSql = "SELECT [" & idColumnName & "] FROM [" & pListObject.Name & "] " & whereClause
    
    ' 3. Executa o SELECT.
    Dim xl As New fwXLConn
    xl.TraceOn = Me.TraceOn
    xl.Init
    
    ' --- CORREÇÃO AQUI ---
    Dim result As Object: Set result = xl.XLRunQuery(selectSql, , forceWriteMode:=True)
    ' --- FIM DA CORREÇÃO ---

    If Not result.Exists("Data") Then
        Trace "DeleteRowsByQuery: Nenhuma linha encontrada para deletar."
        GoTo Cleanup
    End If
    
    ' 4. Monta a lista de IDs.
    Dim idsToDelete As Object: Set idsToDelete = CreateObject("Scripting.Dictionary")
    Dim dataArr As Variant: dataArr = result("Data")
    For j = LBound(dataArr, 2) To UBound(dataArr, 2)
        idsToDelete(Trim(CStr(dataArr(0, j)))) = True
    Next j
    
    ' 5. Deleta as linhas usando VBA.
    idColIndex = pListObject.ListColumns(idColumnName).index
    Dim currentRowValue As String
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    For i = pListObject.ListRows.Count To 1 Step -1
        Dim cellValue As Variant
        cellValue = pListObject.DataBodyRange.Cells(i, idColIndex).value
        
        If Not IsError(cellValue) Then
            currentRowValue = Trim(CStr(cellValue))
            If idsToDelete.Exists(currentRowValue) Then
                pListObject.ListRows(i).Delete
                rowsAffected = rowsAffected + 1
            End If
        End If
    Next i
    
    DeleteRowsByQuery = rowsAffected

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    On Error Resume Next
    If Not savedFilters Is Nothing Then
        RestoreFilters savedFilters
        Trace "DeleteRowsByQuery: Filtro original restaurado."
    End If
    
    If idColumnWasTemporary And Not tempCol Is Nothing Then
        tempCol.Delete ' Deleta a coluna temporária
    End If
    
    If Err.Number <> 0 Then
        Trace "DeleteRowsByQuery ERROR: " & Err.Number & " - " & Err.Description
        Err.Raise Err.Number, "fwXLTable.DeleteRowsByQuery", Err.Description
    End If
End Function

'******************************************************************************
' Nome do elemento: SaveFilters (Function)
' Resumo: Retorna o resultado calculado por SaveFilters.
' @return Collection - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Function SaveFilters() As Collection
    On Error GoTo SafeExit
    Dim filtersCollection As New Collection
    If pListObject Is Nothing Then Set SaveFilters = filtersCollection: Exit Function
    If Not pListObject.AutoFilter.FilterMode Then Set SaveFilters = filtersCollection: Exit Function
    
    Dim i As Long
    Dim filterCriteria As Object
    
    With pListObject.AutoFilter.filters
        For i = 1 To .Count
            If .item(i).On Then
                Set filterCriteria = CreateObject("Scripting.Dictionary")
                filterCriteria.Add "Index", i
                On Error Resume Next
                filterCriteria.Add "Criteria1", .item(i).Criteria1
                filterCriteria.Add "Operator", .item(i).operator
                If .item(i).operator <> 0 Then
                    filterCriteria.Add "Criteria2", .item(i).Criteria2
                End If
                On Error GoTo 0
                filtersCollection.Add filterCriteria
            End If
        Next i
    End With
    
    Set SaveFilters = filtersCollection
    Exit Function
SafeExit:
    Trace "SaveFilters: Ocorreu um erro ao salvar o estado do filtro."
    Set SaveFilters = New Collection
End Function

'******************************************************************************
' Nome do elemento: RestoreFilters (Sub)
' Resumo: Executa a rotina RestoreFilters.
' @param savedFilters - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub RestoreFilters(ByVal savedFilters As Collection)
    On Error GoTo SafeExit
    If savedFilters Is Nothing Then Exit Sub
    If savedFilters.Count = 0 Then Exit Sub
    If pListObject Is Nothing Then Exit Sub
    
    If pListObject.AutoFilter.FilterMode Then pListObject.AutoFilter.ShowAllData
    
    Dim filterCriteria As Object
    For Each filterCriteria In savedFilters
        If filterCriteria.Exists("Index") And filterCriteria.Exists("Criteria1") Then
            If filterCriteria.Exists("Operator") And filterCriteria("Operator") <> 0 And filterCriteria.Exists("Criteria2") Then
                pListObject.Range.AutoFilter field:=filterCriteria("Index"), Criteria1:=filterCriteria("Criteria1"), operator:=filterCriteria("Operator"), Criteria2:=filterCriteria("Criteria2")
            Else
                pListObject.Range.AutoFilter field:=filterCriteria("Index"), Criteria1:=filterCriteria("Criteria1")
            End If
        End If
    Next filterCriteria
    
    Exit Sub
SafeExit:
    Trace "RestoreFilters: Ocorreu um erro ao restaurar o filtro."
End Sub

'******************************************************************************
' Nome do elemento: UpdateRowsByQuery (Function)
' Resumo: Retorna o resultado calculado por UpdateRowsByQuery.
' @param updateSql - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function UpdateRowsByQuery(ByVal updateSql As String) As Long
    On Error GoTo Cleanup
    Dim rowsAffected As Long
    
    Trace "UpdateRowsByQuery ENTER; SQL: " & updateSql
    
    ' 1. Analisa o SQL para extrair as partes
    Dim parsedSql As Object
    Set parsedSql = Me.helpers.ParseUpdateSqlDetails(updateSql)
    If parsedSql Is Nothing Then
        Err.Raise vbObjectError + 9921, "UpdateRowsByQuery", "Falha ao analisar o comando UPDATE."
    End If
    
    Dim targetTable As String: targetTable = parsedSql("TargetTable")
    Dim setClause As String: setClause = parsedSql("SetClause")
    Dim whereClause As String: whereClause = parsedSql("WhereClause")
    
    ' Garante a existência da coluna de chave 'ID'
    Dim idCol As ListColumn
    On Error Resume Next
    Set idCol = pListObject.ListColumns("ID")
    On Error GoTo Cleanup
    If idCol Is Nothing Then
        Err.Raise vbObjectError + 9922, "UpdateRowsByQuery", "A estratégia 'UpdateRowsByQuery' requer que a tabela '" & targetTable & "' tenha uma coluna 'ID'."
    End If
    
    ' 2. Executa um SELECT com a WHERE complexa para obter os IDs
    Dim selectSql As String
    selectSql = "SELECT [ID] FROM [" & targetTable & "] " & whereClause
    Trace "UpdateRowsByQuery: Executando consulta de pré-seleção de IDs: " & selectSql
    
    Dim xl As New fwXLConn
    xl.TraceOn = Me.TraceOn
    xl.Init
    
    ' --- CORREÇÃO AQUI ---
    Dim result As Object: Set result = xl.XLRunQuery(selectSql, , forceWriteMode:=True)
    ' --- FIM DA CORREÇÃO ---
    
    If Not result.Exists("Data") Or Me.helpers.IsVarEmpty(result("Data")) Then
        Trace "UpdateRowsByQuery: Nenhuma linha encontrada para atualizar."
        Exit Function
    End If
    
    ' 3. Constrói a string de IDs para a cláusula IN
    Dim idsToUpdate As Variant: idsToUpdate = result("Data")
    Dim idList As String
    Dim i As Long
    
    On Error Resume Next
    Dim dummy As Long: dummy = UBound(idsToUpdate, 2)
    If Err.Number = 0 Then ' É um array 2D
        For i = LBound(idsToUpdate, 2) To UBound(idsToUpdate, 2)
            idList = idList & "," & Me.helpers.FormatSQLValue(idsToUpdate(0, i), True) ' Formata IDs (pode ser texto)
        Next i
    Else ' É um array 1D
        Err.Clear
        For i = LBound(idsToUpdate) To UBound(idsToUpdate)
            idList = idList & "," & Me.helpers.FormatSQLValue(idsToUpdate(i), True) ' Formata IDs
        Next i
    End If
    On Error GoTo Cleanup
    
    If Len(idList) = 0 Then Exit Function
    idList = Mid(idList, 2) ' Remove a primeira vírgula
    
    ' 4. Monta e executa um ÚNICO comando UPDATE em massa
    Dim bulkUpdateSql As String
    bulkUpdateSql = "UPDATE [" & targetTable & "] SET " & setClause & " WHERE [ID] IN (" & idList & ")"
    
    Call Me.helpers.UpdateStatus("Executando atualização em massa no Excel...")
    Dim updateResult As Object
    Set updateResult = xl.XLRunQuery(bulkUpdateSql) ' Este já está no modo IMEX=0 (isSelect=False)
    If updateResult.Exists("rowsAffected") Then
        rowsAffected = updateResult("rowsAffected")
    End If

    UpdateRowsByQuery = rowsAffected

Cleanup:
    Application.StatusBar = False
    If Err.Number <> 0 Then
        Trace "UpdateRowsByQuery ERROR: " & Err.Number & " - " & Err.Description
        Err.Raise Err.Number, "fwXLTable.UpdateRowsByQuery", Err.Description
    End If
End Function

'******************************************************************************
' Nome do elemento: AddColumnByName (Sub)
' Resumo: Adiciona ou sobrescreve uma coluna aplicando conteúdo e formatação conforme parâmetros.
' @param columnName - Nome da coluna que será criada ou atualizada.
' @param position - Índice base-1 para inserção quando Overwrite=False.
' @param Content - Instrução textual (Text:/Formula:/Sequence:) usada para preencher as células.
' @param Overwrite - Controla se coluna existentes podem ser redefinidas.
'******************************************************************************
Public Sub AddColumnByName( _
    ByVal columnName As String, _
    Optional ByVal position As Long = -1, _
    Optional Content As Variant, _
    Optional Overwrite As Boolean = False _
)
    On Error GoTo ErrHandler
    
    Dim targetCol As ListColumn ' Renomeado de newCol para clareza
    Dim contentStr As String
    Dim prefix As String
    Dim valuePart As String
    Dim colRange As Range
    Dim i As Long ' Index for loops
    Dim excelFormat As String ' Variável para guardar o formato
    Dim cell As Range ' Variável para o loop célula a célula
    Dim columnExisted As Boolean ' Flag para saber se a coluna já existia
            
    If pListObject Is Nothing Then
        Err.Raise vbObjectError + 650, "fwXLTable.AddColumnByName", "ListObject não inicializado."
    End If
    
    If Len(Trim$(columnName)) = 0 Then
        Err.Raise vbObjectError + 651, "fwXLTable.AddColumnByName", "Nome da coluna não pode ser vazio."
    End If
    
    ' --- Verifica se a coluna já existe ---
    On Error Resume Next
    Set targetCol = pListObject.ListColumns(columnName)
    On Error GoTo ErrHandler
    
    If Not targetCol Is Nothing Then
        ' --- Coluna JÁ EXISTE ---
        columnExisted = True
        If Not Overwrite Then
            Trace "AddColumnByName: Coluna '" & columnName & "' já existe. Overwrite=False. Nenhuma ação realizada."
            Exit Sub ' Comportamento padrão: não sobrescrever
        Else
            Trace "AddColumnByName: Coluna '" & columnName & "' já existe. Overwrite=True. Prosseguindo para atualizar conteúdo (se houver)."
            ' A variável targetCol já aponta para a coluna existente.
            ' A lógica de adição será pulada.
        End If
    Else
        ' --- Coluna NÃO EXISTE - Adicionar ---
        columnExisted = False
        Trace "AddColumnByName: Coluna '" & columnName & "' não encontrada. Adicionando..."
        Application.EnableEvents = False ' Desabilita eventos durante a adição
        
        ' Valida a posição e adiciona a coluna
        On Error Resume Next ' Captura erro se Add falhar
        If position > 0 And position <= pListObject.ListColumns.Count + 1 Then
            Set targetCol = pListObject.ListColumns.Add(position:=position)
        Else
            Set targetCol = pListObject.ListColumns.Add ' Adiciona ao final por padrão
        End If
        If Err.Number <> 0 Then GoTo ErrHandler ' Se Add falhou, vai para o handler
        On Error GoTo ErrHandler ' Restaura handler principal

        ' Define o nome da nova coluna
        targetCol.Name = columnName
        Application.EnableEvents = True ' Reabilita eventos após adicionar
        
        Trace "AddColumnByName: Coluna '" & columnName & "' adicionada" & IIf(position > 0 And position <= targetCol.index, " na posição " & position, " ao final") & "."
    End If
    
    ' --- Formatação e Preenchimento do Conteúdo ---
    ' Executa se a coluna foi recém-criada OU se existia e Overwrite=True
    If Not IsMissing(Content) And Not IsEmpty(Content) Then
        contentStr = Trim$(CStr(Content))
        
        If InStr(1, contentStr, ":") > 0 Then
            prefix = UCase$(Trim$(Left$(contentStr, InStr(1, contentStr, ":") - 1)))
            valuePart = Mid$(contentStr, InStr(1, contentStr, ":") + 1)
            
            ' Verifica se há linhas de dados para preencher E formata a coluna
            If pListObject.ListRows.Count > 0 Then
                
                ' Obtém o DataBodyRange da coluna alvo (nova ou existente)
                On Error Resume Next
                Set colRange = Nothing ' Reseta para garantir
                Set colRange = targetCol.DataBodyRange
                If Err.Number <> 0 Or colRange Is Nothing Then
                    Trace "AddColumnByName: ERRO ou Aviso - Não foi possível obter DataBodyRange para a coluna '" & columnName & "'. Preenchimento/Formatação abortados."
                    Err.Clear
                    GoTo CleanupAndExit ' Pula para o final seguro
                End If
                On Error GoTo ErrHandler ' Restaura handler principal

                ' --- Aplica Formatação ANTES de Preencher ---
                excelFormat = "" ' Reseta o formato
                Select Case prefix
                    Case "TEXT": excelFormat = "@"
                    Case "FORMULA": excelFormat = "General"
                    Case "SEQUENCE": excelFormat = "0" ' Formato para Inteiro
                End Select
                
                If excelFormat <> "" Then
                    Application.EnableEvents = False ' Desabilita para formatar
                    On Error Resume Next ' Ignora erro se a formatação falhar
                    colRange.NumberFormat = excelFormat
                    Application.EnableEvents = True ' Reabilita após formatar
                    If Err.Number <> 0 Then
                        Trace "AddColumnByName: Aviso - Falha ao aplicar NumberFormat '" & excelFormat & "'. Erro: " & Err.Description
                        Err.Clear
                    Else
                        Trace "AddColumnByName: NumberFormat '" & excelFormat & "' aplicado à coluna."
                    End If
                    On Error GoTo ErrHandler ' Restaura handler
                End If

                ' --- Preenchimento do Conteúdo ---
                Application.EnableEvents = False ' Desabilita para preencher
                Select Case prefix
                    Case "TEXT"
                        Trace "AddColumnByName: Preenchendo com Texto: '" & valuePart & "'"
                        colRange.value = valuePart
                        
                    Case "FORMULA"
                        valuePart = Trim$(valuePart)
                        If Left$(valuePart, 1) <> "=" Then valuePart = "=" & valuePart
                        Trace "AddColumnByName: Preenchendo com Fórmula Local (Célula a Célula): '" & valuePart & "'"
                        For Each cell In colRange.Cells
                            cell.FormulaLocal = valuePart
                        Next cell
                        
                    Case "SEQUENCE"
                        valuePart = Trim$(valuePart)
                        Trace "AddColumnByName: Tentando preencher com Sequência: '" & valuePart & "'"
                        Dim seqParts() As String
                        Dim startVal As Double, stepVal As Double
                        
                        If Left$(valuePart, 1) = "[" And Right$(valuePart, 1) = "]" Then
                            valuePart = Mid$(valuePart, 2, Len(valuePart) - 2)
                            seqParts = Split(valuePart, ",")
                            
                            If UBound(seqParts) = 1 Then
                                If IsNumeric(Trim$(seqParts(0))) And IsNumeric(Trim$(seqParts(1))) Then
                                    startVal = CDbl(Trim$(seqParts(0)))
                                    stepVal = CDbl(Trim$(seqParts(1)))
                                    
                                    Trace "AddColumnByName: Preenchendo sequência Start=" & startVal & ", Step=" & stepVal
                                    Dim currentValue As Double: currentValue = startVal
                                    For Each cell In colRange.Cells
                                        cell.value = currentValue
                                        currentValue = currentValue + stepVal
                                    Next cell
                                Else
                                    Trace "AddColumnByName: Aviso - Valores de Start/Step na sequência não são numéricos: '" & valuePart & "'"
                                End If
                            Else
                                Trace "AddColumnByName: Aviso - Formato inválido para Sequence (esperado [start,step]): '" & valuePart & "'"
                            End If
                        Else
                            Trace "AddColumnByName: Aviso - Formato inválido para Sequence (faltam colchetes []): '" & valuePart & "'"
                        End If
                        
                    Case Else
                        Trace "AddColumnByName: Aviso - Prefixo de conteúdo desconhecido: '" & prefix & "'. Coluna permanecerá vazia ou com valor anterior se Overwrite=True."
                        ' Não faz nada se o prefixo for inválido
                        
                End Select
                Application.EnableEvents = True ' Reabilita após preencher
            Else
                Trace "AddColumnByName: Tabela está vazia. Conteúdo '" & prefix & "' não aplicado."
            End If
        Else
             Trace "AddColumnByName: Aviso - Formato de conteúdo inválido (falta ':'). Coluna permanecerá vazia ou com valor anterior se Overwrite=True."
        End If
    Else
         Trace "AddColumnByName: Nenhum conteúdo especificado. Coluna permanecerá vazia ou com valor anterior se Overwrite=True."
         ' Se Overwrite=True e Content estiver vazio, nenhum conteúdo será aplicado (a coluna existente manterá seus dados).
         ' Se a coluna foi recém-criada, ela ficará vazia.
    End If

CleanupAndExit: ' Ponto de saída seguro
    Application.EnableEvents = True ' Garante reabilitação de eventos
    Call UpdateTableMetrics ' Atualiza métricas internas após a operação
    
    Trace "AddColumnByName: Operação concluída para coluna '" & columnName & "'."
    Exit Sub

ErrHandler:
    Application.EnableEvents = True ' Garante reabilitação de eventos
    Trace "AddColumnByName ERROR: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "fwXLTable.AddColumnByName", Err.Description
End Sub

'******************************************************************************
' Nome do elemento: RemoveColumnByName (Sub)
' Resumo: Executa a rotina RemoveColumnByName.
' @param columnName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Sub RemoveColumnByName(ByVal columnName As String)
    On Error GoTo ErrHandler
    
    If pListObject Is Nothing Then
        Err.Raise vbObjectError + 660, "fwXLTable.RemoveColumnByName", "ListObject não inicializado."
    End If
    
    If Len(Trim$(columnName)) = 0 Then
        Err.Raise vbObjectError + 661, "fwXLTable.RemoveColumnByName", "Nome da coluna não pode ser vazio."
    End If
    
    Dim colToRemove As ListColumn
    On Error Resume Next ' Tenta encontrar a coluna
    Set colToRemove = pListObject.ListColumns(columnName)
    On Error GoTo ErrHandler ' Restaura handler principal
    
    If colToRemove Is Nothing Then
        Trace "RemoveColumnByName: Coluna '" & columnName & "' não encontrada. Nenhuma ação realizada."
        Exit Sub
    End If
    
    Application.EnableEvents = False ' Desabilita eventos durante a modificação da estrutura
    
    colToRemove.Delete
    
    Application.EnableEvents = True ' Reabilita eventos
    Call UpdateTableMetrics ' Atualiza métricas internas após remover coluna
    
    Trace "RemoveColumnByName: Coluna '" & columnName & "' removida com sucesso."
    Exit Sub

ErrHandler:
    Application.EnableEvents = True ' Garante reabilitação de eventos
    Trace "RemoveColumnByName ERROR: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, "fwXLTable.RemoveColumnByName", Err.Description
End Sub

'******************************************************************************
' Nome do elemento: LoadResult (Function)
' Resumo: Carrega em uma tabela o resultado padronizado retornado por consultas fwXLSPConn/cXLSP.
' @param result - Dicionário contendo chaves Data e Headers.
' @param ClearData - Quando True limpa as linhas existentes antes de inserir.
' @param ClearTitle - Define se o cabeçalho deve ser redefinido.
' @param ResizeTitle - Ajusta o header para coincidir com as coluna recebidas.
' @return Boolean - True quando o carregamento ocorre sem erros.
'******************************************************************************
Public Function LoadResult( _
    ByVal result As Object, _
    Optional ByVal ClearData As Boolean = False, _
    Optional ByVal ClearTitle As Boolean = False, _
    Optional ByVal ResizeTitle As Boolean = False _
) As Boolean
    LoadResult = False ' Assume falha por padrão
    On Error GoTo ErrHandler
    
    Dim arr As Variant
    Dim fields As Variant
    Dim loadRep As Variant ' Para capturar o retorno de XLLoadArray
    
    Trace "LoadResult ENTER para Tabela '" & pListObject.Name & "'"
    
    ' 1. Valida a instância e o objeto Result
    If pListObject Is Nothing Then
        Err.Raise vbObjectError + 750, "fwXLTable.LoadResult", "Instância de fwXLTable não inicializada (pListObject is Nothing)."
    End If
    
    If result Is Nothing Then
        Trace "LoadResult: Objeto Result é Nothing. Nada a carregar."
        LoadResult = True ' Considera sucesso, pois não há o que carregar
        Exit Function
    End If
    
    ' 2. Extrai dados e headers do Result (com segurança)
    If result.Exists("Data") Then arr = result("Data") Else arr = Empty ' Usa Empty se Data não existir
    If result.Exists("Headers") Then fields = result("Headers") Else fields = Array() ' Usa Array vazio se Headers não existir
    
    ' Verifica se há headers (necessário para XLLoadArray)
    On Error Resume Next ' Temporário para checar UBound
    Dim hasHeaders As Boolean: hasHeaders = (UBound(fields) >= LBound(fields))
    On Error GoTo ErrHandler ' Restaura handler
    
    If Not hasHeaders Then
        Trace "LoadResult: Objeto Result não contém Headers válidos. Carregamento abortado."
        ' Decide se isso é um erro ou não. Vamos considerar que não carregar por falta de headers não é um erro fatal.
        LoadResult = True ' Sucesso, mas nada foi carregado.
        Exit Function
    End If

    ' 3. Chama XLLoadArray com os parâmetros fornecidos
    Trace "LoadResult: Chamando XLLoadArray com ClearData=" & ClearData & ", ClearTitle=" & ClearTitle & ", ResizeTitle=" & ResizeTitle
    
    ' Chama a função principal de carregamento desta classe
    Set loadRep = Me.XLLoadArray(arr, fields, ClearData, ClearTitle, ResizeTitle)
    
    ' 4. Verifica o status retornado por XLLoadArray
    If IsObject(loadRep) Then
        If loadRep("status") = "ok" Then
            LoadResult = True
            Trace "LoadResult SUCESSO: Tabela [" & pListObject.Name & "] populada/atualizada."
        Else
            ' XLLoadArray encontrou um erro e já logou, apenas repassa a falha
            Trace "LoadResult FALHA: XLLoadArray retornou status de erro."
            ' Opcional: Poderia levantar um erro aqui baseado em loadRep("errorMessage") se quisesse
        End If
    Else
        ' Se XLLoadArray não retorno um objeto esperado (erro inesperado)
        Trace "LoadResult FALHA: Retorno inesperado de XLLoadArray."
        Err.Raise vbObjectError + 751, "fwXLTable.LoadResult", "Retorno inválido de XLLoadArray."
    End If

    Exit Function

ErrHandler:
    LoadResult = False ' Garante retorno de Falha
    Trace "LoadResult ERRO: " & Err.Number & " - " & Err.Description
    ' O erro será propagado para a rotina chamadora
    ' Não usamos Resume Next aqui, o erro deve parar a execução ou ser tratado acima.
End Function

'******************************************************************************
' Nome do elemento: InFilter (Sub)
' Resumo: Executa a rotina InFilter.
' @param Tp - Parâmetro da rotina.
' @param colName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Sub InFilter(ByVal Tp As Boolean, Optional ByVal colName As String = "")
    Dim sel As Range, dict As Object, c As Range
    Dim fieldIdx As Long, temp As String, arrCrit As Variant
    Set dict = CreateObject("Scripting.Dictionary")
    
    If colName <> "" Then
        fieldIdx = pListObject.ListColumns(colName).index
        Set sel = Intersect(Selection, pListObject.ListColumns(colName).DataBodyRange)
    Else
        fieldIdx = GetFieldIndexFromColumn(ActiveCell.Column)
        Set sel = Intersect(Selection, pListObject.ListColumns(fieldIdx).DataBodyRange)
    End If
    If sel Is Nothing Then Err.Raise vbObjectError + 520, "cListObject.inFiltro", "Seleção inválida"
    
    For Each c In sel.Cells
        If Not c.EntireRow.Hidden Then
            If CStr(c.value) = "" Then
                dict(CStr(c.value)) = "="
            ElseIf Tp Then
                temp = Replace(CStr(c.value), "*", "~*")
                temp = Replace(temp, "?", "~?")
                dict(CStr(c.value)) = temp
            Else
                dict(CStr(c.value)) = CStr(c.value)
            End If
        End If
    Next c
    
    arrCrit = dict.Items
    pListObject.Range.AutoFilter field:=fieldIdx, Criteria1:=arrCrit, operator:=xlFilterValues
End Sub

'******************************************************************************
' Nome do elemento: OutFilter (Sub)
' Resumo: Executa a rotina OutFilter.
' @param Tp - Parâmetro da rotina.
' @param colName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Sub OutFilter(ByVal Tp As Boolean, Optional ByVal colName As String = "")
    Dim nEvents As Boolean: nEvents = Application.EnableEvents
    Dim sel As Range, c As Range
    Dim fieldIdx As Long, colNum As Long, colLetter As String
    Dim firstRowIdx As Long, lastRowIdx As Long
    Dim filters As Object, dictAll As Object
    Dim i As Long, key As Variant, arrCur As Variant, arrCrit As Variant

    Application.EnableEvents = False
    Set filters = CreateObject("Scripting.Dictionary")
    Set dictAll = CreateObject("Scripting.Dictionary")

    ' Captura filtros existentes
    For i = 1 To pListObject.ListColumns.Count
        With pListObject.AutoFilter.filters(i)
            If .On Then
                If .operator = xlFilterValues Then
                    filters(i) = .Criteria1
                ElseIf .operator = xlAnd Or .operator = xlOr Then
                    filters(i) = Array(.Criteria1, .Criteria2)
                Else
                    filters(i) = Array(.Criteria1)
                End If
            End If
        End With
    Next i

    ' Determina coluna e seleção
    If colName <> "" Then
        fieldIdx = pListObject.ListColumns(colName).index
        Set sel = Intersect(Selection, pListObject.ListColumns(fieldIdx).DataBodyRange)
    Else
        fieldIdx = GetFieldIndexFromColumn(ActiveCell.Column)
        Set sel = Intersect(Selection, pListObject.ListColumns(fieldIdx).DataBodyRange)
    End If
    If sel Is Nothing Then Err.Raise vbObjectError + 521, "cListObject.outFiltro", "Seleção inválida"

    firstRowIdx = pHeaderRowIndex + 1
    lastRowIdx = pHeaderRowIndex + pListObject.ListRows.Count
    colLetter = Split(pWorksheet.Cells(pHeaderRowIndex, fieldIdx).Address(True, False), "$")(0)

    ' Popula dictAll
    If Tp Then
        For Each c In pWorksheet.Range(colLetter & firstRowIdx & ":" & colLetter & lastRowIdx).Cells
            If Not c.EntireRow.Hidden Then dictAll(CStr(c.value)) = CStr(c.value)
        Next c
    Else
        If filters.Exists(fieldIdx) Then
            arrCur = filters(fieldIdx)
            If IsArray(arrCur) Then
                For Each key In arrCur: dictAll(CStr(key)) = CStr(key): Next key
            Else
                dictAll(CStr(arrCur)) = CStr(arrCur)
            End If
        Else
            For Each c In pWorksheet.Range(colLetter & firstRowIdx & ":" & colLetter & lastRowIdx).Cells
                If Not c.EntireRow.Hidden Then dictAll(CStr(c.value)) = CStr(c.value)
            Next c
        End If
    End If

    ' Remove valores selecionados
    For Each c In sel.Cells
        If dictAll.Exists(CStr(c.value)) Then dictAll.Remove CStr(c.value)
    Next c

    ' Reaplica filtros
    If pWorksheet.FilterMode Then pWorksheet.ShowAllData
    For Each key In filters.keys
        If key <> fieldIdx Then
            pListObject.Range.AutoFilter field:=key, Criteria1:=filters(key), operator:=xlFilterValues
        End If
    Next key

    If dictAll.Count > 0 Then
        arrCrit = dictAll.keys
        pListObject.Range.AutoFilter field:=fieldIdx, Criteria1:=arrCrit, operator:=xlFilterValues
    Else
        pListObject.Range.AutoFilter field:=fieldIdx
    End If

    Application.EnableEvents = nEvents
End Sub

'******************************************************************************
' Nome do elemento: GetFieldIndexFromColumn (Function)
' Resumo: Retorna o resultado calculado por GetFieldIndexFromColumn.
' @param absCol - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Function GetFieldIndexFromColumn(ByVal absCol As Long) As Long
    Dim lc As ListColumn
    For Each lc In pListObject.ListColumns
        If lc.Range.Column = absCol Then
            GetFieldIndexFromColumn = lc.index
            Exit Function
        End If
    Next lc
    Err.Raise vbObjectError + 516, "cListObject.GetFieldIndexFromColumn", "Coluna não pertence à tabela"
End Function

'******************************************************************************
' Nome do elemento: ToDateSAP (Sub)
' Resumo: Executa a rotina ToDateSAP.
' @param coluna - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Sub ToDateSAP(ByVal colunas As Variant)
    Dim col As Variant
    Dim lc As ListColumn
    Dim rngAll As Range
    Dim cell As Range

    If pListObject Is Nothing Then Err.Raise vbObjectError + 1004, "cListObject", "Tabela não definida"
    If pListObject.ListRows.Count = 0 Then Exit Sub

    For Each col In colunas
        Set lc = Nothing
        On Error Resume Next
        Set lc = pListObject.ListColumns(CStr(col))
        On Error GoTo 0
        If lc Is Nothing Then Err.Raise vbObjectError + 1005, "cListObject", "Coluna '" & col & "' não encontrada"

        Set rngAll = lc.DataBodyRange
        For Each cell In rngAll
            cell.NumberFormat = "@"
            If cell.value <> "" Then cell.value = Format(cell.value, "dd.mm.yyyy")
        Next cell
    Next col
End Sub

'******************************************************************************
' Nome do elemento: ResetTable (Sub)
' Resumo: Executa a rotina ResetTable.
' @param clearFormats - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Sub ResetTable(Optional ByVal clearFormats As Boolean = False)
    If pListObject Is Nothing Then Err.Raise vbObjectError + 640, "fwXLTable.ResetTable", "ListObject não inicializado."

    Dim lo As ListObject
    Set lo = pListObject

    On Error Resume Next
    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop
    On Error GoTo 0

    If clearFormats Then
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.clearFormats
    End If

    Call UpdateTableMetrics
End Sub

'******************************************************************************
' Nome do elemento: EnsureWorksheetRow (Sub)
' Resumo: Executa a rotina EnsureWorksheetRow.
' @param worksheetRow - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Sub EnsureWorksheetRow(ByVal worksheetRow As Long)
    If pListObject Is Nothing Then Err.Raise vbObjectError + 641, "fwXLTable.EnsureWorksheetRow", "ListObject não inicializado."

    Dim targetCount As Long
    targetCount = worksheetRow - pListObject.HeaderRowRange.Row
    If targetCount <= 0 Then Exit Sub

    Do While pListObject.ListRows.Count < targetCount
        pListObject.ListRows.Add
    Loop
    Call UpdateTableMetrics
End Sub

'******************************************************************************
' Nome do elemento: AppendRow (Function)
' Resumo: Retorna o resultado calculado por AppendRow.
' @param valores - Parâmetro da rotina.
' @return ListRow - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function AppendRow(Optional ByVal valores As Object = Nothing) As ListRow
    If pListObject Is Nothing Then Err.Raise vbObjectError + 642, "fwXLTable.AppendRow", "ListObject não inicializado."

    Dim lo As ListObject
    Dim novaLinha As ListRow
    Dim chave As Variant

    Set lo = pListObject
    Set novaLinha = lo.ListRows.Add

    If Not valores Is Nothing Then
        If TypeName(valores) = "Dictionary" Then
            For Each chave In valores.keys
                On Error Resume Next
                novaLinha.Range.Cells(1, lo.ListColumns(CStr(chave)).index).value = valores(chave)
                On Error GoTo 0
            Next chave
        End If
    End If

    Set AppendRow = novaLinha
    Call UpdateTableMetrics
End Function

'******************************************************************************
' Nome do elemento: SetRowValues (Sub)
' Resumo: Executa a rotina SetRowValues.
' @param worksheetRow - Parâmetro da rotina.
' @param valores - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub SetRowValues(ByVal worksheetRow As Long, ByVal valores As Object)
    If valores Is Nothing Then Exit Sub
    Call EnsureWorksheetRow(worksheetRow)

    Dim chave As Variant
    If TypeName(valores) = "Dictionary" Then
        For Each chave In valores.keys
            Me.CellAbs(worksheetRow, CStr(chave)).value = valores(chave)
        Next chave
    End If
End Sub

'******************************************************************************
' Nome do elemento: FillColumn (Sub)
' Resumo: Executa a rotina FillColumn.
' @param colName - Parâmetro da rotina.
' @param valueOrFormula - Parâmetro da rotina.
' @param fromRow - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Sub FillColumn(ByVal colName As String, ByVal valueOrFormula As Variant, Optional ByVal fromRow As Long = 1)
    If pListObject Is Nothing Then Err.Raise vbObjectError + 643, "fwXLTable.FillColumn", "ListObject não inicializado."
    If fromRow < 1 Then fromRow = 1

    Dim relRow As Long
    Dim lastRow As Long
    Dim colIdx As Long

    colIdx = Me.idxCol(colName)
    lastRow = pListObject.ListRows.Count

    For relRow = fromRow To lastRow
        pListObject.DataBodyRange.Cells(relRow, colIdx).value = valueOrFormula
    Next relRow
End Sub

'******************************************************************************
' Nome do elemento: ApplyArray (Sub)
' Resumo: Copia um array bidimensional para a tabela, respeitando a lista de coluna
'         e expandindo o número de linhas conforme necessário.
' @param colNames - Array com os nomes das coluna de destino.
' @param data - Array de dados (linha/coluna) alinhado aos nomes informados.
'******************************************************************************
Public Sub ApplyArray(ByVal colNames As Variant, ByVal data As Variant, _
                      Optional ByVal startRow As Long = 1, _
                      Optional ByVal clearExisting As Boolean = False)
    If pListObject Is Nothing Then Err.Raise vbObjectError + 644, "fwXLTable.ApplyArray", "ListObject não inicializado."
    If Not IsArray(colNames) Then Err.Raise vbObjectError + 645, "fwXLTable.ApplyArray", "colNames precisa ser um array."
    If Not IsArray(data) Then Err.Raise vbObjectError + 646, "fwXLTable.ApplyArray", "data precisa ser um array."

    Dim totalRows As Long
    Dim totalCols As Long
    Dim linha As Long
    Dim coluna As Long
    Dim targetRow As Long
    Dim colName As String

    totalRows = UBound(data, 1) - LBound(data, 1) + 1
    totalCols = UBound(colNames) - LBound(colNames) + 1

    Call EnsureWorksheetRow(pHeaderRowIndex + startRow - 1 + totalRows)

    If clearExisting Then
        For coluna = LBound(colNames) To UBound(colNames)
            Me.RangeCol CStr(colNames(coluna))
            If Not pListObject.ListColumns(CStr(colNames(coluna))).DataBodyRange Is Nothing Then
                pListObject.ListColumns(CStr(colNames(coluna))).DataBodyRange.ClearContents
            End If
        Next coluna
    End If

    For linha = 0 To totalRows - 1
        targetRow = pHeaderRowIndex + startRow + linha
        For coluna = 0 To totalCols - 1
            colName = CStr(colNames(LBound(colNames) + coluna))
            Me.CellAbs(targetRow, colName).value = data(LBound(data, 1) + linha, LBound(data, 2) + coluna)
        Next coluna
    Next linha
End Sub

'******************************************************************************
' Nome do elemento: FindRowIndex (Function)
' Resumo: Retorna o resultado calculado por FindRowIndex.
' @param keyValue - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function FindRowIndex(ByVal keyValue As Variant) As Long
    If pListObject Is Nothing Then Err.Raise vbObjectError + 647, "fwXLTable.FindRowIndex", "ListObject não inicializado."

    Dim navCol As Range
    Dim cell As Range
    Dim idx As Long

    Set navCol = pListObject.ListColumns(pNavColumn).DataBodyRange
    If navCol Is Nothing Then Exit Function

    For Each cell In navCol.Cells
        If cell.value = keyValue Then
            idx = cell.Row
            Exit For
        End If
    Next cell
    FindRowIndex = idx
End Function

'******************************************************************************
' Nome do elemento: GetValue (Function)
' Resumo: Retorna o resultado calculado por GetValue.
' @param keyValue - Parâmetro da rotina.
' @param colName - Parâmetro da rotina.
' @return Variant - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetValue(ByVal keyValue As Variant, ByVal colName As String) As Variant
    Dim absRow As Long
    absRow = FindRowIndex(keyValue)
    If absRow = 0 Then Exit Function

    GetValue = Me.CellAbs(absRow, colName).value
End Function

'******************************************************************************
' Nome do elemento: SetValue (Sub)
' Resumo: Define o valor de uma coluna para a linha cujo identificador corresponde ao valor informado.
'         Pode criar nova linha automaticamente quando addIfMissing = True.
'******************************************************************************
Public Sub SetValue(ByVal keyValue As Variant, ByVal colName As String, ByVal value As Variant, _
                    Optional ByVal addIfMissing As Boolean = False)
    Dim absRow As Long
    absRow = FindRowIndex(keyValue)
    If absRow = 0 Then
        If Not addIfMissing Then Exit Sub
        Dim lr As ListRow
        Set lr = AppendRow()
        lr.Range.Cells(1, Me.idxCol(pNavColumn)).value = keyValue
        absRow = lr.Range.Row
    End If

    Me.CellAbs(absRow, colName).value = value
End Sub

'******************************************************************************
' Nome do elemento: ApplyFilterValue (Sub)
' Resumo: Abstrai o AutoFilter do Excel permitindo aplicar critérios em uma coluna por nome.
'******************************************************************************
Public Sub ApplyFilterValue(ByVal colName As String, ByVal criteria As Variant, _
                            Optional ByVal operator As XlAutoFilterOperator = xlAnd)
    If pListObject Is Nothing Then Err.Raise vbObjectError + 648, "fwXLTable.ApplyFilterValue", "ListObject não inicializado."

    Dim fieldIdx As Long
    fieldIdx = Me.idxCol(colName)
    pListObject.Range.AutoFilter field:=fieldIdx, Criteria1:=criteria, operator:=operator
End Sub

'******************************************************************************
' Nome do elemento: ClearFilter (Sub)
' Resumo: Executa a rotina ClearFilter.
' @param colName - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub ClearFilter(Optional ByVal colName As String)
    If pListObject Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = pWorksheet

    If Len(colName) = 0 Then
        If ws.FilterMode Then ws.ShowAllData
    Else
        Dim fieldIdx As Long
        fieldIdx = Me.idxCol(colName)
        pListObject.Range.AutoFilter field:=fieldIdx
    End If
End Sub

'******************************************************************************
' Nome do elemento: ClearAllFilters (Sub)
' Resumo: Executa a rotina ClearAllFilters.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub ClearAllFilters()
    Call ClearFilter
End Sub

'******************************************************************************
' Nome do elemento: SortByColor (Sub)
' Resumo: Ordena a tabela com base na cor de preenchimento de uma coluna específica.
'******************************************************************************
Public Sub SortByColor(ByVal colName As String, ByVal rgbColor As Long, _
                        Optional ByVal sortOrder As XlSortOrder = xlAscending)
    If pListObject Is Nothing Then Err.Raise vbObjectError + 649, "fwXLTable.SortByColor", "ListObject não inicializado."

    Dim fieldIdx As Long
    fieldIdx = Me.idxCol(colName)

    With pListObject.Sort
        .SortFields.Clear
        .SortFields.Add key:=pListObject.ListColumns(fieldIdx).Range, SortOn:=xlSortOnCellColor, _
                         Order:=sortOrder, DataOption:=xlSortNormal
        .SortFields(1).SortOnValue.Color = rgbColor
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

'******************************************************************************
' Nome do elemento: Class_Terminate (Sub)
' Resumo: Executa a rotina Class_Terminate.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub Class_Terminate()
    On Error Resume Next
    Set pWorksheet = Nothing
    Set pListObject = Nothing
    Set pHelpers = Nothing
End Sub

