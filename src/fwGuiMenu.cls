VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fwGuiMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' Nome do elemento: fwGuiMenu (Classe)
' Resumo: Wrapper para itens de menu SAP GUI com busca híbrida e recursiva.
' @remarks Combina buscas rápidas por Name e fallback recursivo via FindById em /mbar.
'******************************************************************************

Option Explicit

Private pSession    As Object
Private pWindowId   As String
Private pSubMenu    As String
Private pParentMenu As String
Private pMenu       As Object
Private pMenuCache  As Object

'******************************************************************************
' Nome do elemento: Class_Initialize (Sub)
' Resumo: Executa a rotina Class_Initialize.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
' @raises Pode propagar erros levantados por Err.Raise.
'******************************************************************************
Private Sub Class_Initialize()
    On Error GoTo ErrHandler
    Set pMenuCache = CreateObject("Scripting.Dictionary")
    Exit Sub
ErrHandler:
    Err.Raise vbObjectError + 1000, "fwGuiMenu.Class_Initialize", _
              "Falha ao criar cache de menus: " & Err.Description
End Sub

'******************************************************************************
' Nome do elemento: Class_Terminate (Sub)
' Resumo: Executa a rotina Class_Terminate.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Private Sub Class_Terminate()
    On Error Resume Next
    If Not pMenuCache Is Nothing Then
        pMenuCache.RemoveAll
        Set pMenuCache = Nothing
    End If
End Sub

'******************************************************************************
' Nome do elemento: Init (Sub)
' Resumo: Executa a rotina Init.
' @param menuObj - Parâmetro da rotina.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub Init(ByVal menuObj As Object)
    Set pMenu = menuObj
End Sub

'******************************************************************************
' Nome do elemento: SelectMenu (Function)
' Resumo:
'   Tenta selecionar o menu encapsulado e retorna sucesso da operação.
' @return Boolean - True quando o menu existe e o Select executa sem erro.
'******************************************************************************
Public Function SelectMenu() As Boolean
    On Error GoTo SafeExit

    If pMenu Is Nothing Then Exit Function
    pMenu.Select
    SelectMenu = True
    Exit Function

SafeExit:
    SelectMenu = False
End Function

'******************************************************************************
' Nome do elemento: Self (Function)
' Resumo: Retorna o resultado calculado por Self.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Function Self() As Object
    Set Self = pMenu
End Function

'******************************************************************************
' Nome do elemento: Exists (Function)
' Resumo:
'   Indica se o wrapper está vinculado a um item de menu SAP válido.
' @return Boolean - True quando o menu foi encontrado.
'******************************************************************************
Public Function Exists() As Boolean
    On Error GoTo SafeExit

    Exists = Not pMenu Is Nothing
    Exit Function

SafeExit:
    Exists = False
End Function

'******************************************************************************
' Nome do elemento: DumpState (Sub)
' Resumo: Executa a rotina DumpState.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub DumpState()
    If Not pMenu Is Nothing Then pMenu.DumpState
End Sub

'******************************************************************************
' Nome do elemento: Visualize (Sub)
' Resumo: Executa a rotina Visualize.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Sub Visualize()
    If Not pMenu Is Nothing Then pMenu.Visualize
End Sub

'******************************************************************************
' Nome do elemento: ID (Property Get)
' Resumo: Obtem o valor da propriedade ID.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get ID() As String
    If Not pMenu Is Nothing Then ID = pMenu.ID
End Property

'******************************************************************************
' Nome do elemento: Name (Property Get)
' Resumo: Obtem o valor da propriedade Name.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Name() As String
    If Not pMenu Is Nothing Then Name = pMenu.Name
End Property

'******************************************************************************
' Nome do elemento: Parent (Property Get)
' Resumo: Obtem o valor da propriedade Parent.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Parent() As Object
    If Not pMenu Is Nothing Then Set Parent = pMenu.Parent
End Property

'******************************************************************************
' Nome do elemento: MenuTypeName (Property Get)
' Resumo: Obtem o valor da propriedade MenuTypeName.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get MenuTypeName() As String
    If Not pMenu Is Nothing Then MenuTypeName = pMenu.Type
End Property

'******************************************************************************
' Nome do elemento: TypeAsNumber (Property Get)
' Resumo: Obtem o valor da propriedade TypeAsNumber.
' @return Long - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get TypeAsNumber() As Long
    If Not pMenu Is Nothing Then TypeAsNumber = pMenu.TypeAsNumber
End Property

'******************************************************************************
' Nome do elemento: Changeable (Property Get)
' Resumo: Obtem o valor da propriedade Changeable.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Changeable() As Boolean
    Changeable = (Not pMenu Is Nothing And pMenu.Changeable)
End Property

'******************************************************************************
' Nome do elemento: DefaultTooltip (Property Get)
' Resumo: Obtem o valor da propriedade DefaultTooltip.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get DefaultTooltip() As String
    If Not pMenu Is Nothing Then DefaultTooltip = pMenu.DefaultTooltip
End Property

'******************************************************************************
' Nome do elemento: IconName (Property Get)
' Resumo: Obtem o valor da propriedade IconName.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get IconName() As String
    If Not pMenu Is Nothing Then IconName = pMenu.IconName
End Property

'******************************************************************************
' Nome do elemento: IsSymbolFont (Property Get)
' Resumo: Obtem o valor da propriedade IsSymbolFont.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get IsSymbolFont() As Boolean
    IsSymbolFont = (Not pMenu Is Nothing And pMenu.IsSymbolFont)
End Property

'******************************************************************************
' Nome do elemento: Modified (Property Get)
' Resumo: Obtem o valor da propriedade Modified.
' @return Boolean - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Modified() As Boolean
    Modified = (Not pMenu Is Nothing And pMenu.Modified)
End Property

'******************************************************************************
' Nome do elemento: text (Property Get)
' Resumo: Obtem o valor da propriedade text.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Text() As String
    If Not pMenu Is Nothing Then Text = pMenu.Text
End Property

'******************************************************************************
' Nome do elemento: Tooltip (Property Get)
' Resumo: Obtem o valor da propriedade Tooltip.
' @return String - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Tooltip() As String
    If Not pMenu Is Nothing Then Tooltip = pMenu.Tooltip
End Property

'******************************************************************************
' Nome do elemento: Children (Property Get)
' Resumo: Obtem o valor da propriedade Children.
' @return Object - Valor retornado pelo elemento.
' @remarks Bloco documental inserido automaticamente no padrão fwXLTable.
'******************************************************************************
Public Property Get Children() As Object
    If Not pMenu Is Nothing Then Set Children = pMenu.Children
End Property



