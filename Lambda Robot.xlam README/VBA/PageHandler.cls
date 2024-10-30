VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PageHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, MissingMemberAnnotation, ProcedureNotUsed
'@Folder "Lambda.Editor.Metadata.View"

'@PredeclaredId
Option Explicit

Public WithEvents NewButton As MSForms.CommandButton
Attribute NewButton.VB_VarHelpID = -1
Public WithEvents ListViewer As MSForms.ListBox
Attribute ListViewer.VB_VarHelpID = -1
Public WithEvents UpdateButton As MSForms.CommandButton
Attribute UpdateButton.VB_VarHelpID = -1
Public WithEvents DeleteButton As MSForms.CommandButton
Attribute DeleteButton.VB_VarHelpID = -1
Public WithEvents UpArrowButton As MSForms.CommandButton
Attribute UpArrowButton.VB_VarHelpID = -1
Public WithEvents DownArrowButton As MSForms.CommandButton
Attribute DownArrowButton.VB_VarHelpID = -1

Private Type TPageHandler
    ButtonFromGroup As Group
    CaptionTextBox As MSForms.TextBox
    ValueTextBox As MSForms.TextBox
    CrudOperator As ICRUD
    IsChangeByCode As Boolean
End Type

Private This As TPageHandler

Public Property Get IsChangeByCode() As Boolean
    IsChangeByCode = This.IsChangeByCode
End Property

Public Property Let IsChangeByCode(ByVal RHS As Boolean)
    This.IsChangeByCode = RHS
End Property

Public Property Get CrudOperator() As ICRUD
    Set CrudOperator = This.CrudOperator
End Property

Public Property Set CrudOperator(ByVal RHS As ICRUD)
    Set This.CrudOperator = RHS
End Property

Public Property Get ButtonFromGroup() As Group
    ButtonFromGroup = This.ButtonFromGroup
End Property

Public Property Let ButtonFromGroup(ByVal RHS As Group)
    This.ButtonFromGroup = RHS
End Property

Public Property Get CaptionTextBox() As MSForms.TextBox
    Set CaptionTextBox = This.CaptionTextBox
End Property

Public Property Set CaptionTextBox(ByVal RHS As MSForms.TextBox)
    Set This.CaptionTextBox = RHS
End Property

Public Property Get ValueTextBox() As MSForms.TextBox
    Set ValueTextBox = This.ValueTextBox
End Property

Public Property Set ValueTextBox(ByVal RHS As MSForms.TextBox)
    Set This.ValueTextBox = RHS
End Property

Public Function Create(ByVal GivenButtonFromGroup As Group _
                       , ByVal GivenCaptionTextBox As MSForms.TextBox _
                        , ByVal GivenValueTextBox As MSForms.TextBox) As PageHandler

    Logger.Log TRACE_LOG, "Enter PageHandler.Create"
    Dim CurrentPageHandler As PageHandler
    Set CurrentPageHandler = New PageHandler
    With CurrentPageHandler
        .ButtonFromGroup = GivenButtonFromGroup
        Set .CaptionTextBox = GivenCaptionTextBox
        Set .ValueTextBox = GivenValueTextBox
    End With
    ' Creates new PageHandler and sets its properties
    Set Create = CurrentPageHandler
    Logger.Log TRACE_LOG, "Exit PageHandler.Create"

End Function

Private Sub DeleteButton_Click()

    Logger.Log TRACE_LOG, "Enter PageHandler.DeleteButton_Click"
    Dim SelectedIndex As Long
    SelectedIndex = CustomizeListBox.GetSelectedItemIndex(ListViewer)
    If SelectedIndex = -1 Then
        MsgBox "Please select which one to delete.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword PageHandler.DeleteButton_Click"
        Exit Sub
    End If
    Dim Key As String
    Key = ListViewer.List(SelectedIndex, 0)
    This.CrudOperator.Delete Key
    ' Clears the Caption and Value text boxes and updates the ListViewer
    ClearCaptionAndValue
    UpdateListViewer
    Logger.Log TRACE_LOG, "Exit PageHandler.DeleteButton_Click"

End Sub

Private Sub DownArrowButton_Click()

    Logger.Log TRACE_LOG, "Enter PageHandler.DownArrowButton_Click"
    Dim SelectedIndex As Long
    SelectedIndex = CustomizeListBox.GetSelectedItemIndex(ListViewer)
    If SelectedIndex = -1 Then
        MsgBox "Please select which one to reorder.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword PageHandler.DownArrowButton_Click"
        Exit Sub
    End If
    If SelectedIndex = ListViewer.ListCount - 1 Then Exit Sub
    ' Reorders the ListViewer based on the selected index
    UpAndDown SelectedIndex, SelectedIndex + 1
    ListViewer.Selected(SelectedIndex + 1) = True
    Logger.Log TRACE_LOG, "Exit PageHandler.DownArrowButton_Click"

End Sub

Private Sub UpAndDown(ByVal MoveDownIndex As Long, ByVal MoveUpIndex As Long)

    Logger.Log TRACE_LOG, "Enter PageHandler.UpAndDown"
    Dim MoveDownKey As String
    MoveDownKey = ListViewer.List(MoveDownIndex, 0)
    Dim MoveUpKey As String
    MoveUpKey = ListViewer.List(MoveUpIndex, 0)
    This.CrudOperator.Swap MoveDownKey, MoveUpKey
    ' Clears the Caption and Value text boxes and updates the ListViewer
    ClearCaptionAndValue
    UpdateListViewer
    Logger.Log TRACE_LOG, "Exit PageHandler.UpAndDown"

End Sub

Private Sub NewButton_Click()

    Logger.Log TRACE_LOG, "Enter PageHandler.NewButton_Click"
    On Error GoTo DuplicateKey
    If Me.CaptionTextBox.Value = vbNullString Then
        MsgBox "Please add caption and try again.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword PageHandler.NewButton_Click"
        Exit Sub
    End If
    ' Adds a new item to the ListViewer based on input Caption and Value
    This.CrudOperator.AddNew Me.CaptionTextBox.Value, Me.ValueTextBox.Value
    ListViewer.Clear
    ListViewer.List = This.CrudOperator.ReadAll
    TryAdaptingScrollBarHeight ListViewer
    ClearCaptionAndValue
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword PageHandler.NewButton_Click"
    Exit Sub
DuplicateKey:
    MsgBox Err.Number & " : " & Err.Description, vbCritical, "Error"
    Logger.Log TRACE_LOG, "Exit PageHandler.NewButton_Click"

End Sub

Private Sub ClearCaptionAndValue()

    Logger.Log TRACE_LOG, "Enter PageHandler.ClearCaptionAndValue"
    Me.CaptionTextBox.Value = vbNullString
    Me.ValueTextBox.Value = vbNullString
    Logger.Log TRACE_LOG, "Exit PageHandler.ClearCaptionAndValue"

End Sub

Private Sub ListViewer_Change()

    Logger.Log TRACE_LOG, "Enter PageHandler.ListViewer_Change"
    If This.IsChangeByCode Then Exit Sub
    Dim SelectedIndex As Long
    SelectedIndex = CustomizeListBox.GetSelectedItemIndex(ListViewer)
    If SelectedIndex = -1 Then Exit Sub
    ' Changes the Caption and Value text boxes based on the selected item in ListViewer
    This.CaptionTextBox.Value = ListViewer.List(SelectedIndex, 0)
    This.ValueTextBox.Value = ListViewer.List(SelectedIndex, 1)
    Logger.Log TRACE_LOG, "Exit PageHandler.ListViewer_Change"

End Sub

Private Sub UpArrowButton_Click()

    Logger.Log TRACE_LOG, "Enter PageHandler.UpArrowButton_Click"
    Dim SelectedIndex As Long
    SelectedIndex = CustomizeListBox.GetSelectedItemIndex(ListViewer)
    If SelectedIndex = -1 Then
        MsgBox "Please select which one to reorder.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword PageHandler.UpArrowButton_Click"
        Exit Sub
    End If
    If SelectedIndex = 0 Then Exit Sub
    ' Reorders the ListViewer based on the selected index
    UpAndDown SelectedIndex - 1, SelectedIndex
    ListViewer.Selected(SelectedIndex - 1) = True
    Logger.Log TRACE_LOG, "Exit PageHandler.UpArrowButton_Click"

End Sub

Private Sub UpdateButton_Click()

    Logger.Log TRACE_LOG, "Enter PageHandler.UpdateButton_Click"
    If Me.CaptionTextBox.Value = vbNullString Then
        MsgBox "Caption can't be blank. Press delete if you want to remove it.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword PageHandler.UpdateButton_Click"
        Exit Sub
    End If
    Dim SelectedIndex As Long
    SelectedIndex = CustomizeListBox.GetSelectedItemIndex(ListViewer)
    If SelectedIndex = -1 Then Exit Sub
    Dim OldKey As String
    OldKey = ListViewer.List(SelectedIndex, 0)
    ' Updates the ListViewer based on input Caption and Value
    This.CrudOperator.Update OldKey, Me.CaptionTextBox.Value, Me.ValueTextBox.Value
    UpdateListViewer
    ListViewer.Selected(SelectedIndex) = True
    ClearCaptionAndValue
    Logger.Log TRACE_LOG, "Exit PageHandler.UpdateButton_Click"

End Sub

Private Sub UpdateListViewer()

    Logger.Log TRACE_LOG, "Enter PageHandler.UpdateListViewer"
    This.IsChangeByCode = True
    ListViewer.Clear
    Dim AllData As Variant
    AllData = This.CrudOperator.ReadAll
    If IsArray(AllData) Then
        ' Updates ListViewer with the latest data
        ListViewer.List = AllData
        TryAdaptingScrollBarHeight ListViewer
    End If
    This.IsChangeByCode = False
    Logger.Log TRACE_LOG, "Exit PageHandler.UpdateListViewer"

End Sub


