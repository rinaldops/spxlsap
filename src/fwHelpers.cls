VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fwHelpers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' Nome do elemento: fwHelpers (Classe)
' Resumo:
'   Encapsula comportamento da classe com documentacao padronizada.
'******************************************************************************

Option Explicit

Private pTraceOn As Boolean
Private Const LOGWORKSHEETNAME As String = "_fwExecLog"
Private Const LOGMAXROWS As Long = 10000

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
    ' Prefixo: timestamp + nome da classe
    Debug.Print Format$(now, "yyyy-mm-dd hh:nn:ss") & " [fwHelpers] " & msg
End Sub

'******************************************************************************
' Nome do elemento: ToDbValue (Function)
' Resumo: Retorna o resultado calculado por ToDbValue.
' @param v - Parâmetro da rotina.
' @return Variant - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ToDbValue(ByVal v As Variant) As Variant
    If IsNull(v) Then
        ToDbValue = Null
    Else
        Dim s As String: s = CStr(v)
        If Len(Trim$(s)) = 0 Then
            ToDbValue = Null
        Else
            ToDbValue = v
        End If
    End If
End Function

'******************************************************************************
' Nome do elemento: FormatSQLValue (Function)
' Resumo: Formata valores para literais SQL seguros, considerando contexto WHERE/SET.
' @param val - Valor a ser convertido em literal SQL.
' @param forWhere - Quando True aplica regras específicas de comparação (sem aspas em NULL).
' @param typeHint - Letra opcional indicando o tipo (T/N/D).
' @return String - Literal SQL pronto para concatenação.
' @remarks Escapa aspas simples e respeita IMEX quando necessário.
'******************************************************************************
Public Function FormatSQLValue( _
        ByVal val As Variant, _
        Optional ByVal forWhere As Boolean = False, _
        Optional ByVal typeHint As String = "" _
    ) As String
    On Error GoTo ErrHandler

    Dim IsEmpty As Boolean
    IsEmpty = (IsNull(val) Or Len(Trim$(CStr(val))) = 0)

    If Not forWhere Then
        ' --- SET/VALUES (INSERT/UPDATE) => sempre aspas, ou NULL ---
        If IsEmpty Then
            FormatSQLValue = "NULL"
        Else
            FormatSQLValue = "'" & Replace(CStr(val), "'", "''") & "'"
        End If
        Exit Function
    End If

    ' --- WHERE => Texto/Número/Data ---
    If IsEmpty Then
        FormatSQLValue = "NULL" ' o chamador deve preferir IS NULL
        Exit Function
    End If

    Dim t As String
    If Len(typeHint) > 0 Then
        t = UCase$(Left$(typeHint, 1))
    Else
        If IsDate(val) Then
            t = "D"
        ElseIf IsNumeric(val) Then
            t = "N"
        Else
            t = "T"
        End If
    End If

    Select Case t
        Case "D"
            FormatSQLValue = FormatDateLiteral(val)
        Case "N"
            FormatSQLValue = FormatNumberLiteral(val)
        Case Else
            FormatSQLValue = "'" & Replace(CStr(val), "'", "''") & "'"
    End Select
    Exit Function

ErrHandler:
    FormatSQLValue = "NULL"
End Function

'******************************************************************************
' Nome do elemento: FormatNumberLiteral (Function)
' Resumo: Retorna o resultado calculado por FormatNumberLiteral.
' @param v - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function FormatNumberLiteral(ByVal v As Variant) As String
    Dim d As Double
    d = CDbl(v)
    If Abs(d - Round(d, 0)) < 0.0000000005 Then
        FormatNumberLiteral = Format$(Round(d, 0), "0")
    Else
        FormatNumberLiteral = Replace(Format$(d, "0.################"), ",", ".")
    End If
End Function

'******************************************************************************
' Nome do elemento: FormatDateLiteral (Function)
' Resumo: Retorna o resultado calculado por FormatDateLiteral.
' @param v - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function FormatDateLiteral(ByVal v As Variant) As String
    Dim d As Date, hasTime As Boolean
    d = CDate(v)
    hasTime = (d <> Int(d))
    If hasTime Then
        FormatDateLiteral = "#" & Format$(d, "mm/dd/yyyy hh:nn:ss") & "#"
    Else
        FormatDateLiteral = "#" & Format$(d, "mm/dd/yyyy") & "#"
    End If
End Function

'******************************************************************************
' Nome do elemento: NormalizeFieldsParam (Function)
' Resumo: Retorna o resultado calculado por NormalizeFieldsParam.
' @param columnsParam - Parâmetro da rotina.
' @return Variant - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function NormalizeFieldsParam(ByVal columnsParam As Variant) As String()
    Dim tmp As Variant, out() As String
    Dim i As Long, j As Long, lb As Long, ub As Long

    If TypeName(columnsParam) = "Range" Then
        tmp = columnsParam.value
        If IsArray(tmp) Then
            If UBound(tmp, 1) = 1 Then
                ReDim out(0 To UBound(tmp, 2) - 1)
                For j = 1 To UBound(tmp, 2): out(j - 1) = CStr(tmp(1, j)): Next j
            ElseIf UBound(tmp, 2) = 1 Then
                ReDim out(0 To UBound(tmp, 1) - 1)
                For i = 1 To UBound(tmp, 1): out(i - 1) = CStr(tmp(i, 1)): Next i
            Else
                ReDim out(0 To UBound(tmp, 2) - 1)
                For j = 1 To UBound(tmp, 2): out(j - 1) = CStr(tmp(1, j)): Next j
            End If
        Else
            ReDim out(0 To 0): out(0) = CStr(tmp)
        End If
    ElseIf IsArray(columnsParam) Then
        lb = LBound(columnsParam): ub = UBound(columnsParam)
        ReDim out(0 To (ub - lb))
        For i = lb To ub: out(i - lb) = CStr(columnsParam(i)): Next i
    Else
        ReDim out(0 To 0): out(0) = CStr(columnsParam)
    End If

    For i = LBound(out) To UBound(out): out(i) = Trim$(out(i)): Next i
    NormalizeFieldsParam = out
End Function

'******************************************************************************
' Nome do elemento: ToZeroBased (Function)
' Resumo: Retorna o resultado calculado por ToZeroBased.
' @param coll - Parâmetro da rotina.
' @return Variant - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ToZeroBased(ByVal coll As Collection) As Variant
    Dim arr() As Variant, i As Long
    ReDim arr(0 To coll.Count - 1)
    For i = 1 To coll.Count
        arr(i - 1) = coll(i)
    Next
    ToZeroBased = arr
End Function

'******************************************************************************
' Nome do elemento: TickCountMs (Function)
' Resumo: Retorna o resultado calculado por TickCountMs.
' @return Double - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function TickCountMs() As Double
    TickCountMs = Timer * 1000
End Function

'******************************************************************************
' Nome do elemento: IsSelectSql (Function)
' Resumo: Retorna o resultado calculado por IsSelectSql.
' @param sql - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function IsSelectSql(ByVal sql As String) As Boolean
    Dim s As String, i As Long, ch As String
    s = LTrim$(sql)
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "A" And ch <= "z" Then Exit Do
        i = i + 1
    Loop
    s = LCase$(Mid$(s, i))
    IsSelectSql = (Left$(s, 6) = "select")
End Function

'******************************************************************************
' Nome do elemento: IsSQLIdentifierChar (Function)
' Resumo: Retorna o resultado calculado por IsSQLIdentifierChar.
' @param ch - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function IsSQLIdentifierChar(ch As String) As Boolean
    If Len(ch) = 0 Then IsSQLIdentifierChar = False: Exit Function
    Dim c As Long: c = AscW(ch)
    IsSQLIdentifierChar = ((c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 95)
End Function

'******************************************************************************
' Nome do elemento: IsVarEmpty (Function)
' Resumo: Retorna o resultado calculado por IsVarEmpty.
' @param v - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function IsVarEmpty(ByVal v As Variant) As Boolean
    On Error GoTo ErrHandler
    If IsObject(v) Then
        IsVarEmpty = (v Is Nothing)
    ElseIf VBA.IsEmpty(v) Then
        IsVarEmpty = True
    ElseIf IsNull(v) Then
        IsVarEmpty = True
    ElseIf IsArray(v) Then
        Dim lb As Long
        On Error Resume Next
        lb = LBound(v, 1)
        If Err.Number <> 0 Then
            IsVarEmpty = True
            Err.Clear
        Else
            IsVarEmpty = False
        End If
        On Error GoTo ErrHandler
    Else
        IsVarEmpty = (Len(Trim$(CStr(v))) = 0)
    End If
    Exit Function
ErrHandler:
    IsVarEmpty = True
End Function

'******************************************************************************
' Nome do elemento: Unbracket (Function)
' Resumo: Retorna o resultado calculado por Unbracket.
' @param s - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function Unbracket(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    If Len(t) >= 2 Then
        If Left$(t, 1) = "[" And Right$(t, 1) = "]" Then
            t = Mid$(t, 2, Len(t) - 2)
        End If
    End If
    If Right$(t, 1) = "$" Then t = Left$(t, Len(t) - 1)
    Unbracket = t
End Function

'******************************************************************************
' Nome do elemento: FindTopLevelDot (Function)
' Resumo: Localiza o primeiro ponto fora de colchetes em uma expressao SQL.
' @param expr - Expressao a ser analisada.
' @return Long - Posicao do ponto encontrado; 0 quando nao ha separador valido.
'******************************************************************************
Private Function FindTopLevelDot(ByVal expr As String) As Long
    Dim i As Long
    Dim bracketDepth As Long
    Dim ch As String

    For i = 1 To Len(expr)
        ch = Mid$(expr, i, 1)
        Select Case ch
            Case "["
                bracketDepth = bracketDepth + 1
            Case "]"
                If bracketDepth > 0 Then bracketDepth = bracketDepth - 1
            Case "."
                If bracketDepth = 0 Then
                    FindTopLevelDot = i
                    Exit Function
                End If
        End Select
    Next i
End Function

'******************************************************************************
' Nome do elemento: TryParseQualifiedColumnRef (Function)
' Resumo:
'   Interpreta referencias no formato alias.[coluna], preservando pontos
'   dentro do nome da coluna (ex.: alias.[Def.Projeto]).
' @param expr - Expressao SQL de entrada.
' @param tableAlias - Retorno do alias da tabela.
' @param columnName - Retorno do nome da coluna sem colchetes.
' @return Boolean - True quando a expressao foi parseada com sucesso.
'******************************************************************************
Public Function TryParseQualifiedColumnRef( _
    ByVal expr As String, _
    ByRef tableAlias As String, _
    ByRef columnName As String _
) As Boolean
    Dim src As String
    Dim dotPos As Long
    Dim aliasPart As String
    Dim colPart As String

    tableAlias = vbNullString
    columnName = vbNullString

    src = Trim$(expr)
    If Len(src) = 0 Then Exit Function

    dotPos = FindTopLevelDot(src)
    If dotPos = 0 Then Exit Function

    aliasPart = Trim$(Left$(src, dotPos - 1))
    colPart = Trim$(Mid$(src, dotPos + 1))
    If Len(aliasPart) = 0 Or Len(colPart) = 0 Then Exit Function

    tableAlias = Me.Unbracket(aliasPart)

    If Left$(colPart, 1) = "[" And Right$(colPart, 1) = "]" Then
        columnName = Me.Unbracket(colPart)
    Else
        ' Sem colchetes, nao aceitamos novo ponto para evitar ambiguidade.
        If InStr(1, colPart, ".", vbBinaryCompare) > 0 Then Exit Function
        columnName = colPart
    End If

    TryParseQualifiedColumnRef = (Len(tableAlias) > 0 And Len(columnName) > 0)
End Function

'******************************************************************************
' Nome do elemento: TraceArray (Sub)
' Resumo: Executa a rotina TraceArray.
' @param arr - Parâmetro da rotina.
' @param headers - Parâmetro da rotina.
' @param title - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub TraceArray(ByRef arr As Variant, ByRef headers As Variant, ByVal title As String)
    On Error GoTo Cleanup
    If IsVarEmpty(arr) Then
        Trace title & ": Array está vazio ou nulo."
        Exit Sub
    End If
    Dim rL As Long, rU As Long, cL As Long, cU As Long
    On Error Resume Next
    cL = LBound(arr, 1): cU = UBound(arr, 1)
    rL = LBound(arr, 2): rU = UBound(arr, 2)
    If Err.Number <> 0 Then
        Trace title & ": Não foi possível determinar os limites do array."
        Exit Sub
    End If
    On Error GoTo Cleanup
    Dim rowCount As Long: rowCount = rU - rL + 1
    Dim colCount As Long: colCount = cU - cL + 1
    Trace title & ": (" & rowCount & " linhas x " & colCount & " colunas)"
    Dim headerLine As String, c As Long
    headerLine = ""
    For c = LBound(headers) To UBound(headers)
        headerLine = headerLine & headers(c) & vbTab
    Next c
    Trace "  " & headerLine
    Trace "  " & String(Len(headerLine) + 5, "-")
    Dim r As Long, dataLine As String
    For r = rL To IIf(rL + 4 > rU, rU, rL + 4)
        dataLine = ""
        For c = cL To cU
            dataLine = dataLine & arr(c, r) & vbTab
        Next c
        Trace "  " & dataLine
    Next r
    If rU > rL + 4 Then
        Trace "  (... e mais " & (rU - (rL + 4)) & " linhas)"
    End If
Cleanup:
    Exit Sub
End Sub

'******************************************************************************
' Nome do elemento: ParseSqlQueryDetails (Function)
' Resumo: Retorna o resultado calculado por ParseSqlQueryDetails.
' @param sql - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ParseSqlQueryDetails(ByVal sql As String) As Object
    On Error GoTo ErrHandler
    
    Call ValidateSqlTokenSpacing(sql)
    
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    
    Dim tablesColl As New Collection
    Dim columnsColl As New Collection
    
    ' --- INÍCIO DA CORREÇÃO DEFINITIVA DO PARSER ---
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False

    Dim selectClause As String
    Dim fromClause As String
    
    ' Padrão para separar a cláusula SELECT do resto (FROM, WHERE, etc.)
    regEx.Pattern = "^\s*SELECT\s+(.+?)\s+FROM\s+(.+)$"
    Dim matches As Object: Set matches = regEx.Execute(sql)
    
    If matches.Count = 0 Then GoTo MalformedSql
    
    selectClause = Trim(matches(0).SubMatches(0))
    
    ' Remove modificadores iniciais do SELECT para validar apenas as expressões de coluna.
    ' Exemplos: DISTINCT, DISTINCTROW, ALL, TOP n [PERCENT].
    Dim selectClauseNormalized As String
    selectClauseNormalized = selectClause
    
    Dim selectModifierRegEx As Object: Set selectModifierRegEx = CreateObject("VBScript.RegExp")
    selectModifierRegEx.IgnoreCase = True
    selectModifierRegEx.Global = False
    
    selectModifierRegEx.Pattern = "^(DISTINCTROW|DISTINCT|ALL)\s+(.+)$"
    If selectModifierRegEx.Test(selectClauseNormalized) Then
        selectClauseNormalized = Trim$(selectModifierRegEx.Execute(selectClauseNormalized)(0).SubMatches(1))
    End If
    
    selectModifierRegEx.Pattern = "^TOP\s+\d+\s+(PERCENT\s+)?(.+)$"
    If selectModifierRegEx.Test(selectClauseNormalized) Then
        selectClauseNormalized = Trim$(selectModifierRegEx.Execute(selectClauseNormalized)(0).SubMatches(1))
    End If
    
    If Len(selectClauseNormalized) > 0 Then selectClause = selectClauseNormalized
    Set selectModifierRegEx = Nothing
    
    Dim fromAndAfter As String: fromAndAfter = matches(0).SubMatches(1)
    
    ' Lógica robusta para encontrar o fim da cláusula FROM, respeitando parênteses
    Dim parenLevel As Long
    Dim inQuotes As Boolean
    Dim fromEndPos As Long
    Dim i As Long
    fromEndPos = 1
    For i = 1 To Len(fromAndAfter)
        Dim char As String: char = Mid(fromAndAfter, i, 1)
        If char = "'" Then inQuotes = Not inQuotes
        If char = "(" And Not inQuotes Then parenLevel = parenLevel + 1
        If char = ")" And Not inQuotes Then parenLevel = parenLevel - 1
        
        If parenLevel = 0 And Not inQuotes Then
            If i > 5 Then ' Otimização para não checar a cada caractere
                Dim lookahead As String
                lookahead = UCase(Mid(fromAndAfter, i, 11)) ' Pega um pedaço para checar keywords
                If Left(lookahead, 7) = " WHERE " Or _
                   Left(lookahead, 10) = " GROUP BY " Or _
                   Left(lookahead, 10) = " ORDER BY " Then
                    fromEndPos = i
                    Exit For
                End If
            End If
        End If
    Next i
    
    If fromEndPos = 1 Then fromEndPos = Len(fromAndAfter) + 1
    
    fromClause = Trim(Left(fromAndAfter, fromEndPos - 1))
    ' --- FIM DA CORREÇÃO DEFINITIVA DO PARSER ---

    ' O resto da lógica da função, que já estava correta, permanece o mesmo
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    
    Dim fromClauseTablesOnly As String
    Dim onPos As Long: onPos = InStr(1, fromClause, " ON ", vbTextCompare)
    If onPos > 0 Then fromClauseTablesOnly = Trim$(Left$(fromClause, onPos - 1)) Else fromClauseTablesOnly = fromClause
    
    regEx.Pattern = "\[([^\]]+)\](?:\s+AS\s+([a-zA-Z0-9_]+))?"
    Dim tableMatches As Object: Set tableMatches = regEx.Execute(fromClauseTablesOnly)
    Dim tableMatch As Object
    For Each tableMatch In tableMatches
        Dim tblDict As Object: Set tblDict = CreateObject("Scripting.Dictionary")
        Dim tblName As String: tblName = tableMatch.SubMatches(0)
        Dim tblAlias As String
        If tableMatch.SubMatches.Count > 1 Then tblAlias = tableMatch.SubMatches(1)
        If Len(tblAlias) = 0 Then tblAlias = tblName
        tblDict.Add "Name", tblName
        tblDict.Add "Alias", tblAlias
        tablesColl.Add tblDict
    Next
    
    Dim colsRaw As Collection: Set colsRaw = SplitByTopLevelComma(selectClause)
    Dim colDict As Object
    If colsRaw.Count = 1 And Trim$(colsRaw(1)) = "*" Then
        Set colDict = CreateObject("Scripting.Dictionary")
        colDict.Add "SourceExpression", "*": colDict.Add "FinalAlias", "*"
        colDict.Add "TableAlias", "": colDict.Add "ColumnName", "*"
        columnsColl.Add colDict
    Else
        Dim columnExpr As Variant
        For Each columnExpr In colsRaw
            Dim expr As String: expr = Trim(CStr(columnExpr))
            Set colDict = CreateObject("Scripting.Dictionary")
            Dim asPos As Long: asPos = InStr(1, expr, " AS ", vbTextCompare)
            Dim sourceExpr As String, finalAlias As String
            Dim parsedTableAlias As String
            Dim parsedColumnName As String
            Dim hasQualifiedRef As Boolean
            If asPos > 0 Then
                sourceExpr = Trim$(Left$(expr, asPos - 1))
                finalAlias = Trim$(Mid$(expr, asPos + 4))
            Else
                sourceExpr = expr: finalAlias = expr
            End If

            hasQualifiedRef = Me.TryParseQualifiedColumnRef(sourceExpr, parsedTableAlias, parsedColumnName)
            If asPos = 0 And hasQualifiedRef Then finalAlias = parsedColumnName

            finalAlias = Me.Unbracket(finalAlias)
            colDict.Add "SourceExpression", sourceExpr
            colDict.Add "FinalAlias", finalAlias

            If hasQualifiedRef Then
                colDict.Add "TableAlias", parsedTableAlias
                colDict.Add "ColumnName", parsedColumnName
            Else
                colDict.Add "TableAlias", ""
                colDict.Add "ColumnName", Me.Unbracket(sourceExpr)
            End If
            columnsColl.Add colDict
        Next
    End If
    
    result.Add "Tables", tablesColl
    result.Add "Columns", columnsColl
    Set ParseSqlQueryDetails = result
    Exit Function

MalformedSql:
    Trace "ParseSqlQueryDetails ERRO: A estrutura do SQL não pôde ser analisada."
    Set ParseSqlQueryDetails = Nothing
    Exit Function
    
ErrHandler:
    Trace "ParseSqlQueryDetails ERRO INESPERADO: " & Err.Number & " - " & Err.Description
    Set ParseSqlQueryDetails = Nothing
End Function

'******************************************************************************
' Nome do elemento: CollectionToArray (Function)
' Resumo: Retorna o resultado calculado por CollectionToArray.
' @param coll - Parâmetro da rotina.
' @return Variant - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function CollectionToArray(ByVal coll As Collection) As Variant
    If coll Is Nothing Or coll.Count = 0 Then
        CollectionToArray = Array()
        Exit Function
    End If
    Dim arr() As Variant
    ReDim arr(0 To coll.Count - 1)
    Dim i As Long
    For i = 1 To coll.Count
        arr(i - 1) = coll(i)
    Next i
    CollectionToArray = arr
End Function

'******************************************************************************
' Nome do elemento: SplitByTopLevelComma (Function)
' Resumo: Retorna o resultado calculado por SplitByTopLevelComma.
' @param text - Parâmetro da rotina.
' @return Collection - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function SplitByTopLevelComma(ByVal Text As String) As Collection
    Dim result As New Collection
    Dim currentPart As String, parenLevel As Long
    Dim inSingleQuotes As Boolean, inBrackets As Boolean
    Dim i As Long, char As String
    For i = 1 To Len(Text)
        char = Mid$(Text, i, 1)
        If char = "'" And Not inBrackets Then inSingleQuotes = Not inSingleQuotes
        If char = "[" And Not inSingleQuotes Then inBrackets = True
        If char = "]" And inBrackets Then inBrackets = False
        If char = "(" And Not inSingleQuotes And Not inBrackets Then parenLevel = parenLevel + 1
        If char = ")" And Not inSingleQuotes And Not inBrackets Then
            If parenLevel > 0 Then parenLevel = parenLevel - 1
        End If
        If char = "," And parenLevel = 0 And Not inSingleQuotes And Not inBrackets Then
            result.Add Trim$(currentPart)
            currentPart = ""
        Else
            currentPart = currentPart & char
        End If
    Next i
    If Len(Trim$(currentPart)) > 0 Then result.Add Trim$(currentPart)
    Set SplitByTopLevelComma = result
End Function

'******************************************************************************
' Nome do elemento: NormalizeIdentifier (Function)
' Resumo: Retorna o resultado calculado por NormalizeIdentifier.
' @param s - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function NormalizeIdentifier(ByVal s As String) As String
    Dim t As String: t = Trim$(s)
    If Len(t) = 0 Then
        NormalizeIdentifier = t
        Exit Function
    End If
    If Left$(t, 1) = "[" And Right$(t, 1) = "]" Then
        NormalizeIdentifier = t
        Exit Function
    End If
    Dim needs As Boolean, i As Long, ch As String
    If t Like "#*" Then needs = True
    If Not needs Then
        For i = 1 To Len(t)
            ch = Mid$(t, i, 1)
            If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = ".") Then
                needs = True
                Exit For
            End If
        Next i
    End If
    If needs Then
        NormalizeIdentifier = "[" & t & "]"
    Else
        NormalizeIdentifier = t
    End If
End Function

'******************************************************************************
' Nome do elemento: GetDmlTargetTable (Function)
' Resumo: Retorna o resultado calculado por GetDmlTargetTable.
' @param sql - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetDmlTargetTable(ByVal sql As String) As String
    Dim s As String, token As String, nameLen As Long
    s = LCase$(Trim$(sql))
    If Left$(s, 7) = "update " Then
        token = Me.ExtractNameToken(sql, 8, nameLen)
    ElseIf Left$(s, 12) = "delete from " Then
        token = Me.ExtractNameToken(sql, 13, nameLen)
    ElseIf Left$(s, 7) = "delete " Then
        token = Me.ExtractNameToken(sql, 8, nameLen)
        If token = "*" Then token = Me.ExtractNameToken(sql, InStr(8, s, "from") + 5, nameLen)
    End If
    GetDmlTargetTable = Me.Unbracket(token)
End Function

'******************************************************************************
' Nome do elemento: ExtractNameTokenQuickAfterKeyword (Function)
' Resumo: Retorna o resultado calculado por ExtractNameTokenQuickAfterKeyword.
' @param sql - Parâmetro da rotina.
' @param keyword - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ExtractNameTokenQuickAfterKeyword(ByVal sql As String, ByVal keyword As String) As String
    Dim s As String: s = sql
    Dim k As String: k = " " & LCase$(keyword) & " "
    Dim work As String: work = " " & LCase$(s) & " "
    Dim p As Long, startS As Long, nameLen As Long
    p = InStr(1, work, k, vbTextCompare)
    If p = 0 Then Exit Function
    startS = (p - 1) + Len(k)
    ExtractNameTokenQuickAfterKeyword = Me.ExtractNameToken(s, startS, nameLen)
End Function

'******************************************************************************
' Nome do elemento: ReplaceOnce (Function)
' Resumo: Retorna o resultado calculado por ReplaceOnce.
' @param s - Parâmetro da rotina.
' @param oldText - Parâmetro da rotina.
' @param newText - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ReplaceOnce(ByVal s As String, ByVal oldText As String, ByVal newText As String) As String
    Dim p As Long
    p = InStr(1, s, oldText, vbBinaryCompare)
    If p > 0 Then
        ReplaceOnce = Left$(s, p - 1) & newText & Mid$(s, p + Len(oldText))
    Else
        ReplaceOnce = s
    End If
End Function

'******************************************************************************
' Nome do elemento: AceBooleanLiteral (Function)
' Resumo: Retorna o resultado calculado por AceBooleanLiteral.
' @param v - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function AceBooleanLiteral(ByVal v As Variant) As String
    On Error GoTo ErrHandler
    If IsEmpty(v) Or VarType(v) = vbNull Then
        AceBooleanLiteral = "NULL"
    ElseIf VarType(v) = vbBoolean Then
        AceBooleanLiteral = IIf(CBool(v), "1", "0")
    Else
        Dim s As String
        s = LCase$(Trim$(CStr(v)))
        Select Case s
            Case "true", "1", "yes", "sim"
                AceBooleanLiteral = "1"
            Case "false", "0", "no", "nao", "não"
                AceBooleanLiteral = "0"
            Case Else
                AceBooleanLiteral = "NULL"
        End Select
    End If
    Exit Function
ErrHandler:
    AceBooleanLiteral = "NULL"
End Function

'******************************************************************************
' Nome do elemento: FillExecReport (Sub)
' Resumo: Atualiza um dicionário de relatório com status, mensagem e linhas afetadas.
' @param execReport - Dicionário alvo (criado caso necessário).
' @param status - Texto do status atual.
' @param message - Mensagem exibida ao usuário.
' @param rowsAffected - Quantidade de linhas processadas (opcional).
' @remarks Usado por conectores para exibição em UI/StatusBar.
'******************************************************************************
Public Sub FillExecReport( _
    ByRef execReport As Variant, _
    ByVal status As String, _
    ByVal effectiveSql As String, _
    ByVal isSelect As Boolean, _
    Optional ByVal rows As Long = 0, _
    Optional ByVal fields As Long = 0, _
    Optional ByVal rowsAffected As Variant, _
    Optional ByVal errNum As Long = 0, _
    Optional ByVal errSrc As String = "", _
    Optional ByVal errMsg As String = "", _
    Optional ByVal startedAtMs As Double = 0, _
    Optional ByVal endedAtMs As Double = 0 _
)
    On Error Resume Next
    If IsMissing(execReport) Then Exit Sub
    If Me.IsVarEmpty(execReport) Then Set execReport = CreateObject("Scripting.Dictionary")
    With execReport
        .RemoveAll
        .item("status") = status
        .item("effectiveSql") = effectiveSql
        .item("isSelect") = isSelect
        .item("rows") = rows
        .item("fields") = fields
        .item("rowsAffected") = CLng(IIf(IsNull(rowsAffected), 0, rowsAffected))
        .item("errorNumber") = errNum
        .item("errorSource") = errSrc
        .item("errorMessage") = errMsg
        .item("startedAtMs") = startedAtMs
        .item("endedAtMs") = endedAtMs
        .item("elapsedMs") = IIf(endedAtMs >= startedAtMs And startedAtMs > 0, endedAtMs - startedAtMs, 0)
    End With
    On Error GoTo 0
End Sub

'******************************************************************************
' Nome do elemento: GetLocalPathFromSyncUrl (Function)
' Resumo: Retorna o resultado calculado por GetLocalPathFromSyncUrl.
' @param fileUrl - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetLocalPathFromSyncUrl(ByVal fileUrl As String) As String
    Dim objShell As Object
    Dim accounts As Variant
    Dim accountType As Variant
    Dim accountKey As String
    Dim urlPrefix As String
    Dim localRoot As String
    Dim relativePath As String
    Dim testPath As String
    Dim firstSlashPos As Long
    
    On Error GoTo Cleanup ' Lidar com possíveis erros de acesso ao registro

    If InStr(UCase(fileUrl), "C:\") > 0 Then
        GetLocalPathFromSyncUrl = fileUrl
        Exit Function
    End If
    Set objShell = CreateObject("WScript.Shell")
    
    ' Lista de possíveis chaves de conta do OneDrive no registro.
    accounts = Array("Personal", "Business1", "Business2", "Business3", "Business4", "Business5")
    
    For Each accountType In accounts
        On Error Resume Next
        accountKey = "HKEY_CURRENT_USER\Software\Microsoft\OneDrive\Accounts\" & accountType & "\"
        urlPrefix = objShell.RegRead(accountKey & "TeamSiteSPOResourceId")
        If InStr(fileUrl, urlPrefix) = 0 Then urlPrefix = objShell.RegRead(accountKey & "SPOResourceId")
        localRoot = objShell.RegRead(accountKey & "UserFolder")
        On Error GoTo Cleanup
        
        If Len(urlPrefix) > 0 And InStr(1, fileUrl, urlPrefix, vbTextCompare) = 1 Then
            ' Correspondência de conta encontrada! Agora, vamos encontrar o caminho real.
            
            ' 1. Extrai, decodifica e formata o caminho relativo inicial da URL
            relativePath = Mid(fileUrl, Len(urlPrefix) + 1)
            relativePath = Replace(relativePath, "/", "\")
            relativePath = UrlDecode(relativePath)
            
            ' --- INÍCIO DA LÓGICA DE VERIFICAÇÃO ITERATIVA (SUA SUGESTÃO) ---
            Do While Len(relativePath) > 0
                ' 2. Constrói o caminho completo para teste
                testPath = localRoot & "\" & relativePath
                
                ' 3. Verifica se o arquivo existe neste caminho
                '    Usamos Dir() pois é uma forma rápida e nativa de checar a existência.
                If Dir(testPath, vbNormal) <> "" Then
                    ' SUCESSO! O arquivo foi encontrado.
                    GetLocalPathFromSyncUrl = testPath
                    GoTo Cleanup ' Encerra a função e retorna o caminho correto.
                End If
                
                ' 4. Se não encontrou, remove o primeiro nível do caminho relativo.
                firstSlashPos = InStr(1, relativePath, "\")
                
                If firstSlashPos > 0 Then
                    ' Ainda há mais diretórios, então removemos o primeiro.
                    relativePath = Mid$(relativePath, firstSlashPos + 1)
                Else
                    ' Não há mais barras. Significa que já tentamos o último segmento
                    ' (o próprio nome do arquivo) e falhou. Encerramos o loop.
                    relativePath = ""
                End If
            Loop
            ' --- FIM DA LÓGICA DE VERIFICAÇÃO ITERATIVA ---
            
        End If
    Next accountType

Cleanup:
    If Err.Number <> 0 Then
        GetLocalPathFromSyncUrl = ""
    End If
    Set objShell = Nothing
End Function

'******************************************************************************
' Nome do elemento: UrlDecode (Function)
' Resumo: Retorna o resultado calculado por UrlDecode.
' @param encodedString - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function UrlDecode(ByVal encodedString As String) As String
    Dim result As String
    Dim i As Long
    Dim char As String
    Dim hexCode As String
    
    result = ""
    i = 1
    
    Do While i <= Len(encodedString)
        char = Mid$(encodedString, i, 1)
        If char = "%" And i + 2 <= Len(encodedString) Then
            hexCode = Mid$(encodedString, i + 1, 2)
            On Error Resume Next
            ' Converte o código hexadecimal para seu caractere correspondente
            result = result & ChrW(CLng("&H" & hexCode))
            On Error GoTo 0
            i = i + 3
        Else
            result = result & char
            i = i + 1
        End If
    Loop
    
    UrlDecode = result
End Function

'******************************************************************************
' Nome do elemento: SqlAlignToAce (Function)
' Resumo: Retorna o resultado calculado por SqlAlignToAce.
' @param sql - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function SqlAlignToAce(ByVal sql As String) As String
    On Error GoTo SafeExit
    Dim temporarySql As String
    temporarySql = Trim$(sql)
    temporarySql = Replace(temporarySql, vbCrLf, " "): temporarySql = Replace(temporarySql, vbTab, " ")
    Do While InStr(temporarySql, "  ") > 0: temporarySql = Replace(temporarySql, "  ", " "): Loop

    Dim findFullOuter As Long
    findFullOuter = InStr(1, temporarySql, " FULL OUTER JOIN ", vbTextCompare)

    If findFullOuter > 0 Then
        Dim selectClause As String, fromClause As String, whereClause As String
        Dim groupByClause As String, havingClause As String, orderByClause As String
        Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
        regEx.IgnoreCase = True
        
        regEx.Pattern = "^(SELECT .*?)\s+FROM\s+"
        If Not regEx.Test(temporarySql) Then GoTo SafeExit
        selectClause = regEx.Execute(temporarySql)(0).SubMatches(0)
        
        regEx.Pattern = "\s+FROM\s+(.*?)(\s+WHERE|\s+GROUP BY|\s+HAVING|\s+ORDER BY|$)"
        If Not regEx.Test(temporarySql) Then GoTo SafeExit
        fromClause = regEx.Execute(temporarySql)(0).SubMatches(0)
        
        regEx.Pattern = "(\s+WHERE\s+.*?)(\s+GROUP BY|\s+HAVING|\s+ORDER BY|$)"
        If regEx.Test(temporarySql) Then whereClause = regEx.Execute(temporarySql)(0).SubMatches(0)
        
        regEx.Pattern = "(\s+GROUP BY\s+.*?)(\s+HAVING|\s+ORDER BY|$)"
        If regEx.Test(temporarySql) Then groupByClause = regEx.Execute(temporarySql)(0).SubMatches(0)
        
        regEx.Pattern = "(\s+HAVING\s+.*?)(\s+ORDER BY|$)"
        If regEx.Test(temporarySql) Then havingClause = regEx.Execute(temporarySql)(0).SubMatches(0)
        
        Dim orderByPosition As Long: orderByPosition = InStr(1, temporarySql, " ORDER BY ", vbTextCompare)
        If orderByPosition > 0 Then orderByClause = Mid$(temporarySql, orderByPosition)
        
        Dim reWrittenOrderBy As String
        If Len(orderByClause) > 0 Then
            reWrittenOrderBy = orderByClause
            Dim selectFieldsOnly As String: selectFieldsOnly = Trim(Mid(selectClause, Len("SELECT") + 1))
            Dim selectFields As Collection: Set selectFields = SplitByTopLevelComma(selectFieldsOnly)
            Dim field As Variant
            For Each field In selectFields
                Dim asPosition As Long: asPosition = InStr(1, CStr(field), " AS ", vbTextCompare)
                If asPosition > 0 Then
                    Dim originalExpr As String: originalExpr = Trim(Left(field, asPosition - 1))
                    Dim aliasName As String: aliasName = Trim(Mid(field, asPosition + 4))
                    reWrittenOrderBy = Replace(reWrittenOrderBy, originalExpr, aliasName, , , vbTextCompare)
                End If
            Next field
        End If
        
        Dim leftJoinSql As String
        Dim leftJoinFrom As String: leftJoinFrom = Replace(fromClause, "FULL OUTER JOIN", "LEFT JOIN", 1, 1, vbTextCompare)
        leftJoinSql = selectClause & " FROM " & leftJoinFrom & whereClause & groupByClause & havingClause
        
        Dim rightJoinSql As String
        Dim rightJoinFrom As String: rightJoinFrom = Replace(fromClause, "FULL OUTER JOIN", "RIGHT JOIN", 1, 1, vbTextCompare)
        
        Dim onPosition As Long: onPosition = InStr(1, rightJoinFrom, " ON ", vbTextCompare)
        Dim onCondition As String: onCondition = Mid$(rightJoinFrom, onPosition + 4)
        Dim firstTableColumn As String: firstTableColumn = Trim$(Split(onCondition, "=")(0))
        
        Dim rightJoinWhere As String
        If Len(whereClause) > 0 Then
            rightJoinWhere = " WHERE (" & Mid(whereClause, Len(" WHERE ") + 1) & ") AND (" & firstTableColumn & " IS NULL)"
        Else
            rightJoinWhere = " WHERE (" & firstTableColumn & " IS NULL)"
        End If
        
        rightJoinSql = selectClause & " FROM " & rightJoinFrom & rightJoinWhere & groupByClause & havingClause
        
        Dim unionQuery As String
        unionQuery = "(" & leftJoinSql & ") UNION ALL (" & rightJoinSql & ")"
        
        temporarySql = "SELECT * FROM " & unionQuery & reWrittenOrderBy
        
        Trace "SqlAlignToAce: FULL OUTER JOIN traduzido para subconsulta UNION ALL."
    End If
    
    SqlAlignToAce = temporarySql
    Exit Function

SafeExit:
    SqlAlignToAce = sql
End Function

'******************************************************************************
' Nome do elemento: RewriteDataSourceNames (Function)
' Resumo: Retorna o resultado calculado por RewriteDataSourceNames.
' @param sql - Parâmetro da rotina.
' @param wb - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function RewriteDataSourceNames(ByVal sql As String, ByVal wb As Workbook) As String
    On Error GoTo ErrHandler
    Trace "  [HELPER-DEBUG] ==> Entrando em RewriteDataSourceNames (v5 - Autossuficiente)..."
    
    Dim finalSql As String: finalSql = sql
    
    ' Gera um catálogo dinâmico de todas as tabela do Excel na pasta de trabalho
    Dim xlCatalog As Object
    Set xlCatalog = Me.CreateDynamifwXLTableCatalog(wb)
    
    ' Percorre todas as tabela do Excel conhecidas e as substitui onde quer que apareçam
    Dim logicalName As Variant
    For Each logicalName In xlCatalog.keys
        ' Verifica se o nome lógico da tabela está presente no SQL, entre colchetes
        If InStr(1, finalSql, "[" & logicalName & "]", vbTextCompare) > 0 Then
            Trace "  [HELPER-DEBUG]     > Tabela do catálogo encontrada no SQL: '" & logicalName & "'"
            
            Dim resolvedName As String
            resolvedName = Me.ResolveTableSqlName(logicalName, wb)
            
            ' A substituição Replace() é global e segura, pois busca pelo nome exato entre colchetes
            finalSql = Replace(finalSql, "[" & logicalName & "]", resolvedName, 1, -1, vbTextCompare)
        End If
    Next

    Trace "  [HELPER-DEBUG] <== Saindo de RewriteDataSourceNames. SQL Final: " & finalSql
    RewriteDataSourceNames = finalSql
    Exit Function

ErrHandler:
    Trace "  [HELPER-DEBUG]     > ERRO INESPERADO na RewriteDataSourceNames: " & Err.Number & " - " & Err.Description
    RewriteDataSourceNames = sql ' Retorna o SQL original em caso de erro
End Function

'******************************************************************************
' Nome do elemento: ResolveTableSqlName (Function)
' Resumo: Retorna o resultado calculado por ResolveTableSqlName.
' @param logicalName - Parâmetro da rotina.
' @param wb - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ResolveTableSqlName(ByVal logicalName As String, ByVal wb As Workbook) As String
    On Error GoTo ErrHandler
    Trace "    [HELPER-DEBUG] ==> Entrando em ResolveTableSqlName com logicalName: '" & logicalName & "'"
    Dim ln As String: ln = Me.Unbracket(logicalName)
    Dim ws As Worksheet, lo As ListObject
    
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(ln)
        On Error GoTo 0
        
        If Not lo Is Nothing Then
            Trace "    [HELPER-DEBUG]       > SUCESSO: ListObject '" & ln & "' encontrado na planilha '" & ws.Name & "'."
            Dim namedRangeName As String, addr As String
            Dim targetRange As Range
            Dim topLeft As Range, bottomRight As Range

            ' --- LÓGICA DE RECONSTRUÇÃO DE RANGE (mantida) ---
            Set topLeft = lo.HeaderRowRange.Cells(1, 1)
            If lo.DataBodyRange Is Nothing Then
                Set bottomRight = lo.HeaderRowRange.Cells(1, lo.HeaderRowRange.Columns.Count)
            Else
                Set bottomRight = lo.DataBodyRange.Cells(lo.DataBodyRange.rows.Count, lo.DataBodyRange.Columns.Count)
            End If
            Set targetRange = wb.Application.Range(topLeft, bottomRight)
            addr = "'" & lo.Parent.Name & "'!" & targetRange.Address(RowAbsolute:=True, ColumnAbsolute:=True)
            
            ' --- CORREÇÃO FINAL E DEFINITIVA APLICADA AQUI ---
            ' Adiciona um prefixo ao Nome Definido para evitar conflito com o nome do ListObject.
            namedRangeName = "XL_" & lo.Name
            ' --- FIM DA CORREÇÃO ---
            
            Trace "    [HELPER-DEBUG]       > Nome do Named Range (COM PREFIXO): '" & namedRangeName & "'"
            Trace "    [HELPER-DEBUG]       > Endereço alvo final (RefersTo): =" & addr

            On Error Resume Next
            Dim nm As Name: Set nm = wb.Names(namedRangeName)
            If Err.Number <> 0 Then
                Trace "    [HELPER-DEBUG]       > Status: Named Range não existe. Criando agora..."
                Err.Clear
                wb.Names.Add Name:=namedRangeName, RefersTo:="=" & addr
            Else
                Trace "    [HELPER-DEBUG]       > Status: Named Range já existe. Atualizando..."
                nm.RefersTo = "=" & addr
            End If
            On Error GoTo ErrHandler
            
            ' Retorna o nome COM PREFIXO e colchetes
            ResolveTableSqlName = "[" & namedRangeName & "]"
            Trace "    [HELPER-DEBUG] <== Saindo de ResolveTableSqlName. Retornando: '" & ResolveTableSqlName & "'"
            Exit Function
        End If
    Next ws

    Trace "    [HELPER-DEBUG]       > ATENÇÃO: ListObject '" & ln & "' não foi encontrado."
    ResolveTableSqlName = "[" & ln & "]"
    Exit Function

ErrHandler:
    Trace "    [HELPER-DEBUG]       > ERRO INESPERADO na ResolveTableSqlName: " & Err.Number & " - " & Err.Description
    ResolveTableSqlName = "[" & logicalName & "]"
End Function

'******************************************************************************
' Nome do elemento: ExtractNameToken (Function)
' Resumo: Retorna o resultado calculado por ExtractNameToken.
' @param s - Parâmetro da rotina.
' @param startPos - Parâmetro da rotina.
' @param outLen - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ExtractNameToken(ByVal s As String, ByVal startPos As Long, ByRef outLen As Long) As String
    Dim i As Long, k As Long
    Dim ch As String, token As String, tl As Long
    
    outLen = 0
    tl = Len(s)
    
    ' Pula espaços iniciais
    i = startPos
    Do While i <= tl And Mid$(s, i, 1) = " "
        i = i + 1
    Loop
    If i > tl Then Exit Function

    ' Lê o token até encontrar um delimitador (espaço, vírgula, parênteses, etc.)
    k = i
    Do While k <= tl
        ch = Mid$(s, k, 1)
        If ch = " " Or ch = "," Or ch = "(" Or ch = ")" Or ch = vbTab Or ch = vbCr Or ch = vbLf Then Exit Do
        k = k + 1
    Loop

    token = Mid$(s, i, k - i)
    outLen = (k - startPos)
    ExtractNameToken = token
End Function

'******************************************************************************
' Nome do elemento: ExtractAllReferencedColumns (Function)
' Resumo: Retorna o resultado calculado por ExtractAllReferencedColumns.
' @param sql - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ExtractAllReferencedColumns(ByVal sql As String) As Object
    On Error GoTo ErrHandler
    Trace "[DIAGNÓSTICO] ==> Entrando em ExtractAllReferencedColumns"
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Global = True
        .IgnoreCase = True
        .Pattern = "\b([a-zA-Z0-9_]+)\b\.\[([^\]]+)\]" ' alias.[coluna]
    End With
    
    Dim matches As Object
    Set matches = regEx.Execute(sql)
    Trace "[DIAGNÓSTICO] ExtractAllReferencedColumns: Encontrados " & matches.Count & " matches de 'alias.[coluna]'."
    
    Dim columnsByAlias As Object
    Set columnsByAlias = CreateObject("Scripting.Dictionary")
    columnsByAlias.CompareMode = vbTextCompare
    
    Dim match As Object
    For Each match In matches
        Dim tableAlias As String: tableAlias = match.SubMatches(0)
        Dim columnName As String: columnName = match.SubMatches(1)
        Trace "[DIAGNÓSTICO] ExtractAllReferencedColumns:   Match -> Alias: '" & tableAlias & "', Coluna: '" & columnName & "'"
        
        If Not columnsByAlias.Exists(tableAlias) Then
            Trace "[DIAGNÓSTICO] ExtractAllReferencedColumns:     Criando novo dicionário para alias '" & tableAlias & "'"
            Dim uniqueColumns As Object
            Set uniqueColumns = CreateObject("Scripting.Dictionary")
            uniqueColumns.CompareMode = vbTextCompare
            Set columnsByAlias(tableAlias) = uniqueColumns
        End If
        
        If Not columnsByAlias(tableAlias).Exists(columnName) Then
            Trace "[DIAGNÓSTICO] ExtractAllReferencedColumns:     Adicionando coluna '" & columnName & "' ao alias '" & tableAlias & "'"
            columnsByAlias(tableAlias).Add columnName, True
        Else
             Trace "[DIAGNÓSTICO] ExtractAllReferencedColumns:     Coluna '" & columnName & "' já existe para alias '" & tableAlias & "'"
        End If
    Next
    
    Set ExtractAllReferencedColumns = columnsByAlias
     Trace "[DIAGNÓSTICO] <== Saindo de ExtractAllReferencedColumns"
    Exit Function

ErrHandler:
     Trace "[DIAGNÓSTICO] !!! ERRO em ExtractAllReferencedColumns: " & Err.Number & " - " & Err.Description
    Set ExtractAllReferencedColumns = CreateObject("Scripting.Dictionary")
End Function

'******************************************************************************
' Nome do elemento: CleanOldHideWorksheets (Function)
' Resumo: Retorna o resultado calculado por CleanOldHideWorksheets.
' @param wb - Parâmetro da rotina.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function CleanOldHideWorksheets(ByVal wb As Workbook) As Long
    '==========================================================================
    ' NOTA: Se a estrutura do workbook for protegida com senha, a senha
    '       precisará ser fornecida nos comandos Unprotect/Protect.
    '==========================================================================
    Const SENHA_ESTRUTURA As String = "" ' Insira a senha aqui se houver
    '==========================================================================

    Dim ws As Worksheet
    Dim activeSheet As Worksheet
    Dim contador As Long
    Dim i As Long
    Dim wasProtected As Boolean
    
    ' Se nenhum workbook for passado, a função não faz nada.
    If wb Is Nothing Then Exit Function

    ' Salva a planilha ativa para restaurá-la no final
    Set activeSheet = wb.activeSheet
    
    ' Desliga otimizações e alertas para uma execução limpa e rápida
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Garante que as configurações sejam restauradas mesmo em caso de erro
    On Error GoTo Cleanup

    ' 1. Lida com a proteção da estrutura da pasta de trabalho
    If wb.ProtectStructure Then
        wasProtected = True
        wb.Unprotect SENHA_ESTRUTURA
    End If
    
    ' 3. Loop reverso para deletar com segurança
    For i = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(i)
        
        If ws.Name Like "Staging_*" Then
            
            ' Etapa de segurança 1 (sugestão do usuário): Torna a planilha visível
            ws.Visible = xlSheetHidden
            
            ' Etapa de segurança 2 (sugestão do usuário): Remove a tabela (ListObject) da planilha
            If ws.ListObjects.Count > 0 Then
                On Error Resume Next ' Garante que continue mesmo se a exclusão da tabela falhar
                ws.ListObjects(1).Delete
                On Error GoTo Cleanup
            End If
            
            ' Etapa final: Deleta a planilha (agora visível e, idealmente, vazia)
            ws.Delete
            contador = contador + 1
        End If
    Next i

Cleanup:
    ' 4. Bloco de Limpeza: Restaura todas as configurações originais

    ' Reprotege a estrutura se ela estava protegida antes
    If wasProtected Then
        wb.Protect SENHA_ESTRUTURA, Structure:=True
    End If
    
    ' Restaura a planilha que estava ativa originalmente
    If Not activeSheet Is Nothing Then
        On Error Resume Next ' Evita erro se a planilha ativa original foi deletada
        activeSheet.Activate
        On Error GoTo 0
    End If

    ' Reabilita as configurações do Excel
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Se ocorreu um erro durante o processo, o contador não será preciso, mas retornamos o que foi feito
    If Err.Number <> 0 Then
        ' Limpa o erro para não interromper o código que chamou a função
        Trace "CleanOldHideWorksheets encontrou um erro: " & Err.Description
        Err.Clear
    End If

    ' Retorna o número de planilha deletadas
    CleanOldHideWorksheets = contador
End Function

'******************************************************************************
' Nome do elemento: GetExclusionZones (Function)
' Resumo: Retorna o resultado calculado por GetExclusionZones.
' @param sql - Parâmetro da rotina.
' @return Collection - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetExclusionZones(ByVal sql As String) As Collection
    Dim zones As New Collection
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim zone As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    
    ' Padrão para encontrar tudo dentro de colchetes.
    regEx.Pattern = "\[[^\]]+\]"
    Set matches = regEx.Execute(sql)
    For Each match In matches
        Set zone = CreateObject("Scripting.Dictionary")
        zone.Add "Start", match.FirstIndex + 1 ' Posição 1-based
        zone.Add "End", match.FirstIndex + match.Length
        zones.Add zone
    Next
    
    ' Padrão para encontrar tudo dentro de aspas simples.
    regEx.Pattern = "'[^']+'"
    Set matches = regEx.Execute(sql)
    For Each match In matches
        Set zone = CreateObject("Scripting.Dictionary")
        zone.Add "Start", match.FirstIndex + 1 ' Posição 1-based
        zone.Add "End", match.FirstIndex + match.Length
        zones.Add zone
    Next
    
    Set GetExclusionZones = zones
End Function

'******************************************************************************
' Nome do elemento: ValidateSqlTokenSpacing (Sub)
' Resumo: Executa a rotina ValidateSqlTokenSpacing.
' @param sql - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Sub ValidateSqlTokenSpacing(ByVal sql As String)
    On Error GoTo ErrHandler

    ' Etapa 1: Mapear todas as zonas que devem ser ignoradas.
    Dim exclusionZones As Collection
    Set exclusionZones = GetExclusionZones(sql)
    
    ' Listas de tokens e suas regras de espaçamento
    Dim primaryTokens() As Variant: primaryTokens = Array("SELECT", "UPDATE", "DELETE", "INSERT")
    Dim secondaryTokens() As Variant: secondaryTokens = Array("FROM", "WHERE", "AS", "JOIN", "ON", "BY", "GROUP", "ORDER", "INNER", "LEFT", "RIGHT", "OUTER", "HAVING", "IN", "LIKE", "AND", "OR", "INTO")
    
    Dim token As Variant
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim isInsideZone As Boolean
    Dim zone As Object
    Dim tokenStart As Long, tokenEnd As Long
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True

    ' Etapa 2: Validar tokens primários (só precisam de espaço DEPOIS)
    For Each token In primaryTokens
        regEx.Pattern = "\b" & token & "\b"
        Set matches = regEx.Execute(sql)
        For Each match In matches
            tokenStart = match.FirstIndex + 1
            tokenEnd = tokenStart + match.Length - 1
            
            isInsideZone = False
            For Each zone In exclusionZones
                If tokenStart >= zone("Start") And tokenEnd <= zone("End") Then
                    isInsideZone = True
                    Exit For
                End If
            Next
            
            If Not isInsideZone Then
                ' É um token real. Validar o espaçamento.
                If tokenEnd < Len(sql) Then
                    If Mid$(sql, tokenEnd + 1, 1) <> " " Then
                        Err.Raise vbObjectError + 1021, "ValidateSqlTokenSpacing", "Erro de sintaxe: A palavra-chave '" & UCase$(token) & "' deve ser seguida por um espaço."
                    End If
                End If
            End If
        Next
    Next

    ' Etapa 2 (cont.): Validar tokens secundários (precisam de espaço ANTES e DEPOIS)
    For Each token In secondaryTokens
        ' --- INÍCIO DA CORREÇÃO ---
        ' Lógica especial para o token "IN" para evitar correspondência com "INNER" ou "INTO"
        If UCase$(token) = "IN" Then
            ' Usa uma "negative lookahead" (?!\w) para garantir que "IN" não é seguido por outra letra/número.
            regEx.Pattern = "\bIN\b(?!\w)"
        Else
            regEx.Pattern = "\b" & token & "\b"
        End If
        ' --- FIM DA CORREÇÃO ---
        
        Set matches = regEx.Execute(sql)
        For Each match In matches
            tokenStart = match.FirstIndex + 1
            tokenEnd = tokenStart + match.Length - 1
            
            isInsideZone = False
            For Each zone In exclusionZones
                If tokenStart >= zone("Start") And tokenEnd <= zone("End") Then
                    isInsideZone = True
                    Exit For
                End If
            Next
            
            If Not isInsideZone Then
                ' É um token real. Validar o espaçamento.
                If tokenStart > 1 Then
                    If Mid$(sql, tokenStart - 1, 1) <> " " Then
                        Err.Raise vbObjectError + 1022, "ValidateSqlTokenSpacing", "Erro de sintaxe: A palavra-chave '" & UCase$(token) & "' deve ser precedida por um espaço."
                    End If
                End If
                If tokenEnd < Len(sql) Then
                     If Mid$(sql, tokenEnd + 1, 1) <> " " Then
                        Err.Raise vbObjectError + 1023, "ValidateSqlTokenSpacing", "Erro de sintaxe: A palavra-chave '" & UCase$(token) & "' deve ser seguida por um espaço."
                    End If
                End If
            End If
        Next
    Next
    
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "fwHelpers." & Err.source, Err.Description
End Sub

'******************************************************************************
' Nome do elemento: IsInsertSql (Function)
' Resumo: Retorna o resultado calculado por IsInsertSql.
' @param sql - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function IsInsertSql(ByVal sql As String) As Boolean
    ' A lógica é similar à IsSelectSql, procurando pelo token no início da string.
    Dim s As String
    s = LTrim$(sql)
    If LCase$(Left$(s, 7)) = "insert " Then
        IsInsertSql = True
    Else
        IsInsertSql = False
    End If
End Function

'******************************************************************************
' Nome do elemento: ParseInsertSqlDetails (Function)
' Resumo: Retorna o resultado calculado por ParseInsertSqlDetails.
' @param sql - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function ParseInsertSqlDetails(ByVal sql As String) As Object
    On Error GoTo ErrHandler
    
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    Dim matches As Object
    
    With regEx
        .IgnoreCase = True
        .Global = False
        .Pattern = "^\s*INSERT\s+INTO\s+(\[[^\]]+\])\s*(?:\(([^)]+)\))?\s*(VALUES|SELECT)"
    End With
    
    Set matches = regEx.Execute(sql)
    
    If matches.Count = 0 Then
        Err.Raise vbObjectError + 1030, "ParseInsertSqlDetails", "Sintaxe do comando INSERT inválida. Estrutura 'INSERT INTO [Tabela] ... VALUES/SELECT' não encontrada."
    End If
    
    Dim mainMatch As Object: Set mainMatch = matches(0)
    
    result.Add "TargetTable", Me.Unbracket(mainMatch.SubMatches(0))
    result.Add "InsertType", UCase$(mainMatch.SubMatches(2))
    
    Dim columnsArray() As String
    If mainMatch.SubMatches(1) <> "" Then
        Dim columnListStr As String: columnListStr = mainMatch.SubMatches(1)
        
        regEx.Pattern = "\[[^\]]+\]"
        regEx.Global = True
        Set matches = regEx.Execute(columnListStr)
        
        If matches.Count > 0 Then
            ReDim columnsArray(0 To matches.Count - 1)
            Dim i As Long
            For i = 0 To matches.Count - 1
                columnsArray(i) = Me.Unbracket(matches(i).value)
            Next i
        End If
    End If
    result.Add "Columns", columnsArray
    
    If result("InsertType") = "SELECT" Then
        Dim selectKeywordPos As Long
        selectKeywordPos = InStr(mainMatch.FirstIndex + 1, UCase$(sql), "SELECT")
        
        Dim selectSubQuery As String
        selectSubQuery = Trim$(Mid$(sql, selectKeywordPos))
        
        result.Add "SelectSubQuery", selectSubQuery
    End If

    Set ParseInsertSqlDetails = result
    Exit Function

ErrHandler:
    Set ParseInsertSqlDetails = Nothing
    Err.Raise Err.Number, "fwHelpers.ParseInsertSqlDetails", Err.Description
End Function

'******************************************************************************
' Nome do elemento: IsDeleteSql (Function)
' Resumo: Retorna o resultado calculado por IsDeleteSql.
' @param sql - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function IsDeleteSql(ByVal sql As String) As Boolean
    Dim s As String
    s = LTrim$(sql)
    If LCase$(Left$(s, 7)) = "delete " Then
        IsDeleteSql = True
    Else
        IsDeleteSql = False
    End If
End Function

'******************************************************************************
' Nome do elemento: ParseDeleteSqlDetails (Function)
' Resumo: Retorna o resultado calculado por ParseDeleteSqlDetails.
' @param sql - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function ParseDeleteSqlDetails(ByVal sql As String) As Object
    On Error GoTo ErrHandler
    
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    Dim matches As Object
    
    With regEx
        .IgnoreCase = True
        .Global = False
        ' Padrão para encontrar: DELETE [FROM] [Tabela] [WHERE ...]
        ' Grupo 1: (\\[[^\\]]+\\]) -> Captura o nome da tabela
        ' Grupo 2: (WHERE .*)?   -> Captura a cláusula WHERE (opcional)
        .Pattern = "^\s*DELETE\s+(?:FROM\s+)?(\[[^\]]+\])\s*(WHERE\s+.*)?$"
    End With
    
    Set matches = regEx.Execute(sql)
    
    If matches.Count = 0 Then
        Err.Raise vbObjectError + 1040, "ParseDeleteSqlDetails", "Sintaxe do comando DELETE inválida. Estrutura 'DELETE FROM [Tabela]' não encontrada."
    End If
    
    result.Add "TargetTable", Me.Unbracket(matches(0).SubMatches(0))
    result.Add "WhereClause", matches(0).SubMatches(1)
    
    Set ParseDeleteSqlDetails = result
    Exit Function

ErrHandler:
    Set ParseDeleteSqlDetails = Nothing
    Err.Raise Err.Number, "fwHelpers.ParseDeleteSqlDetails", Err.Description
End Function

'******************************************************************************
' Nome do elemento: IsUpdateSql (Function)
' Resumo: Retorna o resultado calculado por IsUpdateSql.
' @param sql - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function IsUpdateSql(ByVal sql As String) As Boolean
    Dim s As String
    s = LTrim$(sql)
    If LCase$(Left$(s, 7)) = "update " Then
        IsUpdateSql = True
    Else
        IsUpdateSql = False
    End If
End Function

'******************************************************************************
' Nome do elemento: ParseUpdateSqlDetails (Function)
' Resumo: Retorna o resultado calculado por ParseUpdateSqlDetails.
' @param sql - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function ParseUpdateSqlDetails(ByVal sql As String) As Object
    On Error GoTo ErrHandler

    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    Dim matches As Object

    With regEx
        .IgnoreCase = True
        .Global = False
        ' --- INÍCIO DA CORREÇÃO ---
        ' Padrão mais explícito para separar a cláusula SET do WHERE.
        ' Grupo 1: Tabela Alvo
        ' Grupo 2: Conteúdo da Cláusula SET
        ' Grupo 3: Conteúdo da Cláusula WHERE (sem a palavra "WHERE")
        .Pattern = "^\s*UPDATE\s+(\S+|\[[^\]]+\])\s+SET\s+(.+?)(?:\s+WHERE\s+(.*))?$"
        ' --- FIM DA CORREÇÃO ---
    End With

    Set matches = regEx.Execute(sql)

    If matches.Count = 0 Then
        Err.Raise vbObjectError + 1050, "ParseUpdateSqlDetails", "Sintaxe do comando UPDATE inválida. Estrutura 'UPDATE [Tabela] SET' não encontrada."
    End If

    ' --- INÍCIO DA CORREÇÃO ---
    result.Add "TargetTable", Me.Unbracket(matches(0).SubMatches(0))
    result.Add "SetClause", Trim(matches(0).SubMatches(1))
    
    ' Adiciona a palavra "WHERE" de volta se a cláusula foi encontrada
    If matches(0).SubMatches(2) <> "" Then
        result.Add "WhereClause", "WHERE " & matches(0).SubMatches(2)
    Else
        result.Add "WhereClause", ""
    End If
    ' --- FIM DA CORREÇÃO ---

    Set ParseUpdateSqlDetails = result
    Exit Function

ErrHandler:
    Set ParseUpdateSqlDetails = Nothing
    Err.Raise Err.Number, "fwHelpers.ParseUpdateSqlDetails", Err.Description
End Function

'******************************************************************************
' Nome do elemento: CreateDynamifwXLTableCatalog (Function)
' Resumo: Retorna o resultado calculado por CreateDynamifwXLTableCatalog.
' @param wb - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function CreateDynamifwXLTableCatalog(ByVal wb As Workbook) As Object
    Dim cat As Object: Set cat = CreateObject("Scripting.Dictionary")
    cat.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If Not cat.Exists(lo.Name) Then
                Dim catEntry As Object: Set catEntry = CreateObject("Scripting.Dictionary")
                catEntry.Add "Type", "XL"
                ' O dicionário de conexão pode ser vazio para fontes XL
                catEntry.Add "Connection", CreateObject("Scripting.Dictionary")
                cat.Add lo.Name, catEntry
            End If
        Next lo
    Next ws

    Set CreateDynamifwXLTableCatalog = cat
End Function

'******************************************************************************
' Nome do elemento: GetWhereClauseDependencies (Function)
' Resumo: Retorna o resultado calculado por GetWhereClauseDependencies.
' @param whereClauseText - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function GetWhereClauseDependencies(ByVal whereClauseText As String) As Object
    On Error GoTo ErrHandler
    
    ' Usa um dicionário para garantir que cada alias seja listado apenas uma vez.
    Dim aliases As Object
    Set aliases = CreateObject("Scripting.Dictionary")
    aliases.CompareMode = vbTextCompare
    
    If Me.IsVarEmpty(whereClauseText) Then
        Set GetWhereClauseDependencies = aliases
        Exit Function
    End If

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Global = True
        .IgnoreCase = True
        ' A expressão procura pelo padrão "alias.[coluna]" para identifica
        ' de forma inequívoca uma referência a uma coluna de uma tabela com alias.
        .Pattern = "\b([a-zA-Z0-9_]+)\b\.\[([^\]]+)\]"
    End With
    
    Dim matches As Object
    Set matches = regEx.Execute(whereClauseText)
    
    Dim match As Object
    For Each match In matches
        Dim tableAlias As String: tableAlias = match.SubMatches(0)
        
        ' Adiciona o alias ao dicionário. Se já existir, nada acontece.
        If Not aliases.Exists(tableAlias) Then
            aliases.Add tableAlias, True
        End If
    Next
    
    Set GetWhereClauseDependencies = aliases
    Exit Function

ErrHandler:
    ' Em caso de erro, retorna um dicionário vazio para não quebrar o fluxo.
    Set GetWhereClauseDependencies = CreateObject("Scripting.Dictionary")
End Function

'******************************************************************************
' Nome do elemento: ParseUpdateJoinSqlDetails (Function)
' Resumo: Retorna o resultado calculado por ParseUpdateJoinSqlDetails.
' @param sql - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Public Function ParseUpdateJoinSqlDetails(ByVal sql As String) As Object
    On Error GoTo ErrHandler
    Trace "[DIAGNÓSTICO] ==> ENTRANDO em ParseUpdateJoinSqlDetails com SQL: " & sql ' Adicionado SQL na entrada
    
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    Dim matches As Object
    
    With regEx
        .IgnoreCase = True
        .Global = False
        ' Tenta pegar com WHERE primeiro
        .Pattern = "^\s*UPDATE\s+(.+?)\s+SET\s+(.+?)\s+(WHERE\s+.+)$"
    End With
    
    Set matches = regEx.Execute(sql)
    
    ' Fallback para queries sem cláusula WHERE
    If matches.Count = 0 Then
        Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails: Pattern com WHERE falhou. Tentando pattern sem WHERE..."
        regEx.Pattern = "^\s*UPDATE\s+(.+?)\s+SET\s+(.+)$" ' Usar +? (não guloso) aqui também
        Set matches = regEx.Execute(sql)
    End If
    
    If matches.Count = 0 Then
        Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails: ERRO - Nenhum dos patterns correspondeu."
        Err.Raise vbObjectError + 1060, "ParseUpdateJoinSqlDetails", "Sintaxe do comando UPDATE inválida. Estrutura 'UPDATE [Tabelas] SET [Valores] [WHERE ...]' não encontrada."
    End If
    
    Dim mainMatch As Object: Set mainMatch = matches(0)
    Dim tablesClause As String: tablesClause = Trim(mainMatch.SubMatches(0))
    Dim setClause As String: setClause = Trim(mainMatch.SubMatches(1))
    Dim whereClause As String
    If mainMatch.SubMatches.Count > 2 Then
         whereClause = Trim(mainMatch.SubMatches(2)) ' Captura WHERE Clause
    Else
         whereClause = "" ' Define como vazio se não houver WHERE
    End If

    Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails -> tablesClause: '" & tablesClause & "'"
    Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails -> setClause: '" & setClause & "'"
    Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails -> whereClause: '" & whereClause & "'"
    
    ' --- Lógica para extrair tabela e aliases (mantida) ---
    Dim tablesToParse As String
    Dim onPosition As Long: onPosition = InStr(1, tablesClause, " ON ", vbTextCompare)
    
    ' Ajuste: Parsear tudo antes do PRIMEIRO " ON " se houver JOIN
    If onPosition > 0 Then
        tablesToParse = Trim(Left(tablesClause, onPosition - 1))
         Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails: JOIN detectado. tablesToParse (antes do ON): '" & tablesToParse & "'"
    Else
        ' Se não houver JOIN explícito (UPDATE simples?), pega só a primeira parte
        tablesToParse = tablesClause ' Assume que é só "Tabela AS Alias" ou "[Tabela]"
         Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails: JOIN não detectado via ' ON '. tablesToParse: '" & tablesToParse & "'"
         ' Adiciona lógica para pegar só a primeira tabela em caso de UPDATE simples
         Dim spacePos As Long: spacePos = InStr(tablesToParse, " ")
         If spacePos > 0 And InStr(1, tablesToParse, " AS ", vbTextCompare) = 0 Then
             tablesToParse = Trim(Left(tablesToParse, spacePos - 1))
              Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails: Ajustado tablesToParse para UPDATE simples: '" & tablesToParse & "'"
         End If
    End If

    Dim tablesRegEx As Object: Set tablesRegEx = CreateObject("VBScript.RegExp")
    With tablesRegEx
        .Global = True
        .IgnoreCase = True
        ' Pattern ajustado para ser mais robusto:
        ' (\[[^\]]+\]|\b\w+\b)  -> Captura nome com colchetes OU palavra simples (Nome)
        ' \s*(?:AS\s+(\b\w+\b))? -> Captura alias opcional APENAS se precedido por AS (Alias)
        .Pattern = "(\[[^\]]+\]|\b\w+\b)\s*(?:AS\s+(\b\w+\b))?"
    End With
    
    Dim tableMatches As Object: Set tableMatches = tablesRegEx.Execute(tablesToParse)
    Dim tablesColl As New Collection
    Dim isFirstTable As Boolean: isFirstTable = True
    
    If tableMatches.Count = 0 Then
         Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails: ERRO - Nenhuma tabela/alias encontrado em tablesToParse: '" & tablesToParse & "'"
        Err.Raise vbObjectError + 1061, "ParseUpdateJoinSqlDetails", "Nenhuma tabela encontrada na cláusula UPDATE após análise: '" & tablesToParse & "'"
    End If
    
    Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails: Encontrados " & tableMatches.Count & " matches de tabela/alias em '" & tablesToParse & "'"
    Dim tableMatch As Object
    For Each tableMatch In tableMatches
         Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails: Match bruto: '" & tableMatch.value & "'"
         Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails: SubMatch 0 (Nome): '" & tableMatch.SubMatches(0) & "'"
         If tableMatch.SubMatches.Count > 1 Then Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails: SubMatch 1 (Alias): '" & tableMatch.SubMatches(1) & "'"

        ' Ignora palavras-chave SQL que podem ser capturadas acidentalmente
        Dim rawName As String: rawName = tableMatch.SubMatches(0)
        Select Case UCase(Trim(rawName))
            Case "INNER", "LEFT", "RIGHT", "OUTER", "JOIN", "AS"
                 Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails: Ignorando keyword '" & rawName & "'"
                GoTo NextMatch
        End Select
        
        Dim tblName As String: tblName = Me.Unbracket(rawName)
        Dim tblAlias As String
        If tableMatch.SubMatches.Count > 1 Then tblAlias = tableMatch.SubMatches(1) Else tblAlias = ""

        ' Fallback para alias se não houver 'AS' explícito
        If Len(tblAlias) = 0 Then tblAlias = tblName
         Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails: Processado -> Nome: '" & tblName & "', Alias: '" & tblAlias & "'"
        
        Dim tblDict As Object: Set tblDict = CreateObject("Scripting.Dictionary")
        tblDict.Add "Name", tblName
        tblDict.Add "Alias", tblAlias
        tablesColl.Add tblDict
        
        If isFirstTable Then
            result.Add "TargetTable", tblName ' A primeira tabela é o alvo do UPDATE
             Trace "[DIAGNÓSTICO] ParseUpdateJoinSqlDetails: Definido TargetTable como '" & tblName & "'"
            isFirstTable = False
        End If
NextMatch:
    Next tableMatch
    
    result.Add "Tables", tablesColl
    result.Add "SetClause", setClause
    result.Add "WhereClause", whereClause
    result.Add "TablesClause", tablesClause ' Adiciona a cláusula completa original
    
    Set ParseUpdateJoinSqlDetails = result
    Trace "[DIAGNÓSTICO] <== SAINDO de ParseUpdateJoinSqlDetails"
    Exit Function

ErrHandler:
    Trace "[DIAGNÓSTICO] !!! ERRO em ParseUpdateJoinSqlDetails: " & Err.Number & " - " & Err.Description
    Set ParseUpdateJoinSqlDetails = Nothing
    Err.Raise Err.Number, "fwHelpers.ParseUpdateJoinSqlDetails", Err.Description
End Function

'******************************************************************************
' Nome do elemento: UpdateStatus (Sub)
' Resumo: Executa a rotina UpdateStatus.
' @param msg - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub UpdateStatus(ByVal msg As String)
    On Error Resume Next
    
    ' Constrói uma mensagem padronizada
    Dim fullMsg As String
    fullMsg = "[Orquestrador] " & msg
    
    ' Atualiza a barra de status do Excel
    Application.StatusBar = fullMsg
    
    ' Se o trace estiver ligado, também envia para o log
    If pTraceOn Then
        Trace Format$(now, "yyyy-mm-dd hh:nn:ss") & " [STATUS] " & msg
    End If
    
    ' Permite que a interface do Excel (incluindo a barra de status) se atualize
    DoEvents
End Sub

'******************************************************************************
' Nome do elemento: NotifyProgress (Sub)
' Resumo:
'   Atualiza a barra de status com progresso percentual em intervalos
'   controlados para evitar overhead de UI em loops longos.
' @param current - Índice corrente da execução.
' @param total - Total esperado de itens.
' @param message - Texto base da mensagem.
' @param updateEvery - Frequência de atualização (padrão: 10).
'******************************************************************************
Public Sub NotifyProgress(ByVal current As Long, ByVal total As Long, ByVal message As String, Optional ByVal updateEvery As Long = 10)
    On Error GoTo SafeExit

    If total <= 0 Then
        Call UpdateStatus(message)
        Exit Sub
    End If

    If updateEvery < 1 Then updateEvery = 1
    If current < 0 Then current = 0
    If current > total Then current = total

    If current > 1 And current < total Then
        If (current Mod updateEvery) <> 0 Then Exit Sub
    End If

    Dim pct As Long
    pct = CLng((CDbl(current) / CDbl(total)) * 100#)
    Call UpdateStatus(message & " " & current & "/" & total & " (" & pct & "%).")
    Exit Sub

SafeExit:
    Err.Clear
End Sub

'******************************************************************************
' Nome do elemento: ClearPersistentLog (Sub)
' Resumo:
'   Limpa os registros da planilha de log persistente, preservando o cabeçalho.
'******************************************************************************
Public Sub ClearPersistentLog()
    On Error GoTo SafeExit

    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = EnsureLogWorksheet()
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then
        ws.rows("2:" & CStr(lastRow)).Delete Shift:=xlShiftUp
    End If
    Exit Sub

SafeExit:
    Err.Clear
End Sub

'******************************************************************************
' Nome do elemento: AppendPersistentLog (Sub)
' Resumo:
'   Registra eventos de execução em planilha oculta para suporte e auditoria.
' @param level - Severidade (INFO/WARN/ERROR).
' @param sourceName - Origem (classe/módulo).
' @param actionName - Ação em execução.
' @param message - Mensagem consolidada.
' @param sqlText - SQL associado ao evento.
' @param sapTCode - Transação SAP associada ao evento.
'******************************************************************************
Public Sub AppendPersistentLog( _
    ByVal level As String, _
    ByVal sourceName As String, _
    ByVal actionName As String, _
    ByVal message As String, _
    Optional ByVal sqlText As String = "", _
    Optional ByVal sapTCode As String = "" _
)
    On Error GoTo SafeExit

    Dim ws As Worksheet
    Set ws = EnsureLogWorksheet()
    If ws Is Nothing Then Exit Sub

    Dim nextRow As Long
    nextRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    If nextRow > (LOGMAXROWS + 1) Then
        ws.rows(2).Delete Shift:=xlShiftUp
        nextRow = LOGMAXROWS + 1
    End If

    ws.Cells(nextRow, 1).value = now
    ws.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(nextRow, 2).value = UCase$(Trim$(level))
    ws.Cells(nextRow, 3).value = sourceName
    ws.Cells(nextRow, 4).value = actionName
    ws.Cells(nextRow, 5).value = message
    ws.Cells(nextRow, 6).value = sqlText
    ws.Cells(nextRow, 7).value = sapTCode

    If pTraceOn Then
        Trace "[LOG] " & UCase$(Trim$(level)) & " | " & sourceName & "." & actionName & " | " & message
    End If
    Exit Sub

SafeExit:
    Err.Clear
End Sub

'******************************************************************************
' Nome do elemento: EnsureLogWorksheet (Function)
' Resumo:
'   Garante a existência da planilha oculta de log persistente.
' @return Worksheet - Planilha de log ou Nothing em caso de indisponibilidade.
'******************************************************************************
Private Function EnsureLogWorksheet() As Worksheet
    On Error GoTo SafeExit

    If ActiveWorkbook Is Nothing Then Exit Function

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook

    On Error Resume Next
    Set ws = wb.Worksheets(LOGWORKSHEETNAME)
    On Error GoTo SafeExit

    If ws Is Nothing Then
        Dim oldDisplayAlerts As Boolean
        oldDisplayAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False

        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = LOGWORKSHEETNAME
        ws.Visible = xlSheetVeryHidden
        ws.Range("A1:G1").value = Array("Timestamp", "Level", "Source", "Action", "Message", "Sql", "SapTCode")
        ws.rows(1).Font.Bold = True
        ws.Columns("A:G").EntireColumn.AutoFit

        Application.DisplayAlerts = oldDisplayAlerts
    End If

    Set EnsureLogWorksheet = ws
    Exit Function

SafeExit:
    Set EnsureLogWorksheet = Nothing
    Err.Clear
End Function

'******************************************************************************
' Nome do elemento: ExtractAllTableNamesFromClauses (Function)
' Resumo: Retorna o resultado calculado por ExtractAllTableNamesFromClauses.
' @param sql - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ExtractAllTableNamesFromClauses(ByVal sql As String) As Object
    On Error GoTo ErrHandler
    
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    Dim matches As Object
    Dim match As Object
    Dim uniqueNames As Object: Set uniqueNames = CreateObject("Scripting.Dictionary")
    uniqueNames.CompareMode = vbTextCompare
    
    With regEx
        .Global = True
        .IgnoreCase = True
        ' Este padrão procura pelas palavras-chave FROM ou JOIN, seguidas de espaço,
        ' e captura o que estiver dentro de colchetes logo a seguir.
        .Pattern = "\b(?:FROM|JOIN)\s+\[([^\]]+)\]"
    End With
    
    Set matches = regEx.Execute(sql)
    
    For Each match In matches
        ' O nome da tabela é o primeiro (e único) submatch capturado.
        Dim tblName As String: tblName = match.SubMatches(0)
        
        If Not uniqueNames.Exists(tblName) Then
            uniqueNames.Add tblName, True
        End If
    Next
    
    Set ExtractAllTableNamesFromClauses = uniqueNames
    Exit Function
    
ErrHandler:
    Set ExtractAllTableNamesFromClauses = CreateObject("Scripting.Dictionary") ' Retorna dicionário vazio em caso de erro
End Function

'******************************************************************************
' Nome do elemento: BuildSelectForUpdate (Function)
' Resumo: Retorna o resultado calculado por BuildSelectForUpdate.
' @param updateSql - Parâmetro da rotina.
' @param targetTableAlias - Parâmetro da rotina.
' @param pkFieldName - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function BuildSelectForUpdate(ByVal updateSql As String, ByVal targetTableAlias As String, ByVal pkFieldName As String) As String
    On Error GoTo SafeExit
    
    Dim parsedSql As Object: Set parsedSql = Me.ParseUpdateJoinSqlDetails(updateSql)
    If parsedSql Is Nothing Then Exit Function
    
    Dim setClause As String: setClause = parsedSql("SetClause")
    Dim tablesClause As String: tablesClause = parsedSql("TablesClause")
    Dim whereClause As String: whereClause = parsedSql("WhereClause")
    
    Dim fieldsToSelect As String
    fieldsToSelect = targetTableAlias & ".[" & pkFieldName & "]"
    
    Dim assignments As Variant: assignments = Split(setClause, ",")
    Dim assignment As Variant
    For Each assignment In assignments
        Dim parts As Variant: parts = Split(Trim(assignment), "=")
        If UBound(parts) = 1 Then
            Dim targetField As String: targetField = Trim(parts(0))
            
            ' Verifica se o campo que está sendo atualizado pertence ao alias da tabela alvo
            If InStr(1, targetField, targetTableAlias & ".", vbTextCompare) = 1 Then
                Dim newAlias As String
                ' Cria um nome de alias "limpo", removendo todos os caracteres especiais.
                newAlias = "New_" & targetField
                newAlias = Replace(newAlias, "[", "")
                newAlias = Replace(newAlias, "]", "")
                newAlias = Replace(newAlias, ".", "_")
                newAlias = Replace(newAlias, " ", "")
                
                ' Adiciona a coluna-alvo (ex: h.[Ano]) e o novo valor (ex: p.[Ano])
                fieldsToSelect = fieldsToSelect & ", " & targetField
                fieldsToSelect = fieldsToSelect & ", " & Trim(parts(1)) & " AS " & newAlias
            End If
        End If
    Next
    
    BuildSelectForUpdate = "SELECT " & fieldsToSelect & " FROM " & tablesClause & " " & whereClause
    Exit Function
    
SafeExit:
    BuildSelectForUpdate = ""
End Function

'******************************************************************************
' Nome do elemento: ApplyOnClauseCasting (Function)
' Resumo: Retorna o resultado calculado por ApplyOnClauseCasting.
' @param sql - Parâmetro da rotina.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ApplyOnClauseCasting(ByVal sql As String) As String
    On Error Resume Next
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .IgnoreCase = True
        .Global = True
        .Pattern = "(\bON\s+)(\S+)(\s*=\s*)(\S+)"
    End With
    
    Dim result As String
    result = regEx.Replace(sql, "$1CLng($2)$3CLng($4)")
    
    If Len(result) > 0 And StrComp(result, sql, vbTextCompare) <> 0 Then
        Trace "ApplyOnClauseCasting: Casting CLng() aplicado à cláusula ON."
        ApplyOnClauseCasting = result
    Else
        ApplyOnClauseCasting = sql
    End If
End Function

'******************************************************************************
' Nome do elemento: ParseSPInitParams (Function)
' Resumo: Retorna o resultado calculado por ParseSPInitParams.
' @param params - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ParseSPInitParams(ByVal params As Variant) As Object
    On Error GoTo SafeExit
    Set ParseSPInitParams = CreateObject("Scripting.Dictionary")
    ParseSPInitParams.CompareMode = vbTextCompare

    Dim lb As Long
    Dim ub As Long
    lb = LBound(params)
    ub = UBound(params)
    If ub < lb Then Exit Function

    Dim i As Long
    i = lb
    Do While i <= ub
        If i + 2 > ub Then Exit Do

        Dim logicalName As String: logicalName = CStr(params(i))
        Dim siteURL As String: siteURL = CStr(params(i + 1))
        Dim listID As String: listID = CStr(params(i + 2))
        i = i + 3

        Dim spDetails As Object: Set spDetails = CreateObject("Scripting.Dictionary")
        spDetails.Add "SiteURL", siteURL
        spDetails.Add "ListID", listID

        Dim deltaSyncEnabled As Boolean
        Dim deltaSyncField As String
        Dim deltaSyncStateKey As String
        deltaSyncEnabled = False
        deltaSyncField = "Modified"
        deltaSyncStateKey = logicalName

        If i <= ub Then
            If IsObject(params(i)) Then
                Dim optionsDict As Object
                Set optionsDict = params(i)
                If TypeName(optionsDict) = "Dictionary" Then
                    If optionsDict.Exists("DeltaSyncEnabled") Then
                        deltaSyncEnabled = CBool(optionsDict("DeltaSyncEnabled"))
                    End If
                    If optionsDict.Exists("DeltaSyncField") Then
                        If Len(Trim$(CStr(optionsDict("DeltaSyncField")))) > 0 Then
                            deltaSyncField = Trim$(CStr(optionsDict("DeltaSyncField")))
                        End If
                    End If
                    If optionsDict.Exists("DeltaSyncStateKey") Then
                        If Len(Trim$(CStr(optionsDict("DeltaSyncStateKey")))) > 0 Then
                            deltaSyncStateKey = Trim$(CStr(optionsDict("DeltaSyncStateKey")))
                        End If
                    End If
                    i = i + 1
                End If
            End If
        End If

        Dim catEntry As Object: Set catEntry = CreateObject("Scripting.Dictionary")
        catEntry.Add "Type", "SP"
        catEntry.Add "Connection", spDetails
        catEntry.Add "DeltaSyncEnabled", deltaSyncEnabled
        catEntry.Add "DeltaSyncField", deltaSyncField
        catEntry.Add "DeltaSyncStateKey", deltaSyncStateKey

        If ParseSPInitParams.Exists(logicalName) Then
            Set ParseSPInitParams(logicalName) = catEntry
        Else
            ParseSPInitParams.Add logicalName, catEntry
        End If
    Loop

    Exit Function

SafeExit:
    Err.Clear
End Function

'******************************************************************************
' Nome do elemento: ParseOnClauseColumns (Function)
' Resumo: Retorna o resultado calculado por ParseOnClauseColumns.
' @param sql - Parâmetro da rotina.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ParseOnClauseColumns(ByVal sql As String) As Object
    On Error GoTo SafeExit
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    Dim matches As Object
    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    
    With regEx
        .IgnoreCase = True
        .Global = False ' Pega apenas a primeira cláusula ON
        ' Padrão para encontrar: ON <alias.campo> = <alias.campo>
        ' Captura a primeira coluna (ex: a.[ID]) e a segunda (ex: b.[IDProduto])
        .Pattern = "\bON\s+([\w\.]+\[[^\]]+\])\s*=\s*([\w\.]+\[[^\]]+\])"
    End With
    
    Set matches = regEx.Execute(sql)
    
    If matches.Count > 0 Then
        result.Add "LeftColumn", matches(0).SubMatches(0) ' Ex: "a.[ID]"
        result.Add "RightColumn", matches(0).SubMatches(1) ' Ex: "b.[IDProduto]"
        Trace "ParseOnClauseColumns: Colunas de junção encontradas: " & result("LeftColumn") & ", " & result("RightColumn")
    End If
    
    Set ParseOnClauseColumns = result
    Exit Function
    
SafeExit:
    Set ParseOnClauseColumns = CreateObject("Scripting.Dictionary") ' Retorna vazio em caso de erro
End Function

'******************************************************************************
' Nome do elemento: IsWhereClauseComplex (Function)
' Resumo: Retorna o resultado calculado por IsWhereClauseComplex.
' @param whereClause - Parâmetro da rotina.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function IsWhereClauseComplex(ByVal whereClause As String) As Boolean
    ' Verifica a presença do operador de concatenação
    If InStr(1, whereClause, "&", vbTextCompare) > 0 Then
        IsWhereClauseComplex = True
        Exit Function
    End If
    
    ' Verifica a presença de funções de conversão comuns
    If InStr(1, whereClause, "CStr(", vbTextCompare) > 0 Then
        IsWhereClauseComplex = True
        Exit Function
    End If
    
    If InStr(1, whereClause, "Format(", vbTextCompare) > 0 Then
        IsWhereClauseComplex = True
        Exit Function
    End If
    
    ' Se nenhuma das condições complexas for encontrada, retorna False
    IsWhereClauseComplex = False
End Function

'******************************************************************************
' Nome do elemento: ParseSetClause (Function)
' Resumo: Retorna o resultado calculado por ParseSetClause.
' @param setClause - Parâmetro da rotina.
' @return Collection - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function ParseSetClause(ByVal setClause As String) As Collection
    Dim result As New Collection
    
    ' Usa nosso helper que já entende parênteses para separar as atribuições
    Dim assignments As Collection
    Set assignments = Me.SplitByTopLevelComma(setClause)
    
    Dim assignment As Variant
    For Each assignment In assignments
        Dim assignmentDict As Object: Set assignmentDict = CreateObject("Scripting.Dictionary")
        
        ' Encontra a posição do primeiro '=' que não está dentro de parênteses
        Dim parenLevel As Long
        Dim eqPos As Long
        Dim i As Long
        For i = 1 To Len(assignment)
            Dim char As String: char = Mid(assignment, i, 1)
            If char = "(" Then parenLevel = parenLevel + 1
            If char = ")" Then parenLevel = parenLevel - 1
            If char = "=" And parenLevel = 0 Then
                eqPos = i
                Exit For
            End If
        Next i
        
        If eqPos > 0 Then
            assignmentDict.Add "TargetField", Trim(Left(assignment, eqPos - 1))
            assignmentDict.Add "SourceExpression", Trim(Mid(assignment, eqPos + 1))
            result.Add assignmentDict
        End If
    Next assignment
    
    Set ParseSetClause = result
End Function


