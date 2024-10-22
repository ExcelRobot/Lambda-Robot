VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LETManagerUI 
   Caption         =   "LET Steps Manager"
   ClientHeight    =   8115
   ClientLeft      =   -360
   ClientTop       =   -1755
   ClientWidth     =   16140
   OleObjectBlob   =   "LETManagerUI.frx":0000
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "LETManagerUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, ImplicitActiveSheetReference

Option Explicit
Private Enum CloseBy
    User = 0
    Code = 1
    WindowsOS = 2
    TaskManager = 3
End Enum

Private Type TLETManagerUI
    Parser As FormulaParser
    DependencyObjects As Collection
    Counter As Long
End Type

Private This As TLETManagerUI

Public Property Get DependencyObjects() As Collection
    Set DependencyObjects = This.DependencyObjects
End Property

Public Property Set DependencyObjects(ByVal RHS As Collection)
    Set This.DependencyObjects = RHS
End Property

Public Property Get Parser() As FormulaParser
    Set Parser = This.Parser
End Property

Public Property Set Parser(ByVal RHS As FormulaParser)
    Set This.Parser = RHS
End Property

Private Sub CancelButton_Click()
    Logger.Log TRACE_LOG, "Enter LETManagerUI.CancelButton_Click"
    This.Parser.IsProcessTerminatedByUser = True
    Me.Hide
    Logger.Log TRACE_LOG, "Exit LETManagerUI.CancelButton_Click"
End Sub

Private Sub SelectAgainAfterExclude(ByVal ForListBox As MSForms.ListBox _
                                    , ByVal SelectedRowIndex As Long)

    Logger.Log TRACE_LOG, "Enter LETManagerUI.SelectAgainAfterExclude"

    ' Ensure that a valid row index is selected after an exclusion
    If ForListBox.ListCount = 0 Then Exit Sub
    If SelectedRowIndex = -1 Then Exit Sub

    With ForListBox
        If SelectedRowIndex >= .ListCount Then
            .Selected(.ListCount - 1) = True
        Else
            .Selected(SelectedRowIndex) = True
        End If
    End With

    Logger.Log TRACE_LOG, "Exit LETManagerUI.SelectAgainAfterExclude"

End Sub

Private Function GetSelectedItemVarName(ByVal ForListBox As MSForms.ListBox) As String

    Logger.Log TRACE_LOG, "Enter LETManagerUI.GetSelectedItemVarName"
    ' Get the variable name of the selected item in the given ListBox
    If ForListBox.ListIndex = -1 Then Exit Function
    GetSelectedItemVarName = GetItemVarName(ForListBox, ForListBox.ListIndex)
    Logger.Log TRACE_LOG, "Exit LETManagerUI.GetSelectedItemVarName"

End Function

Private Function GetItemVarName(ByVal ForListBox As MSForms.ListBox _
                                , ByVal FromIndex As Long) As String

    Logger.Log TRACE_LOG, "Enter LETManagerUI.GetItemVarName"
    ' Get the variable name of the item from the given ListBox at the specified index
    GetItemVarName = ForListBox.List(FromIndex, 0)
    Logger.Log TRACE_LOG, "Exit LETManagerUI.GetItemVarName"

End Function

Private Sub ExcludeStepButton_Click()

    Logger.Log TRACE_LOG, "Enter LETManagerUI.ExcludeStepButton_Click"
    ' Get the variable name of the selected item in the StepsListBox
    Dim SelectedDependencyVarName As String

    Dim IsAnyItemSelected As Boolean
    Dim Index As Long
    For Index = 0 To Me.StepsListBox.ListCount - 1
        If Me.StepsListBox.Selected(Index) Then
            IsAnyItemSelected = True
            SelectedDependencyVarName = GetItemVarName(Me.StepsListBox, Index)
            ' Get the DependencyInfo object that matches the selected variable name
            Dim SelectedDependency As DependencyInfo
            Set SelectedDependency = GetMatchingVarNameDependency(SelectedDependencyVarName _
                                                                  , This.DependencyObjects)

            ' Mark the selected dependency as not being a Let statement by the user
            SelectedDependency.IsMarkAsNotLetStatementByUser = True
        End If
    Next Index

    If IsAnyItemSelected Then
        ' Update the StepsListBox after the exclusion
        UpdateListBoxAfterExclude Me.StepsListBox
    End If
    Logger.Log TRACE_LOG, "Exit LETManagerUI.ExcludeStepButton_Click"

End Sub

Private Sub UpdateListBoxAfterExclude(ByVal ForListBox As MSForms.ListBox)

    Logger.Log TRACE_LOG, "Enter LETManagerUI.UpdateListBoxAfterExclude"
    ' Update the ListBox after excluding a dependency
    Dim SelectedRowIndex As Long
    SelectedRowIndex = ForListBox.ListIndex
    RecalculateAndUpdateDependencyCollection
    UpdateListBoxFromCollection
    SelectAgainAfterExclude ForListBox, SelectedRowIndex
    Logger.Log TRACE_LOG, "Exit LETManagerUI.UpdateListBoxAfterExclude"

End Sub

Private Sub RecalculateAndUpdateDependencyCollection()

    Logger.Log TRACE_LOG, "Enter LETManagerUI.RecalculateAndUpdateDependencyCollection"
    ' Recalculate the precedence and update the DependencyObjects
    This.Parser.RecalculatePrecedencyAgain This.DependencyObjects, LET_STATEMENT_GENERATION
    Set This.DependencyObjects = This.Parser.PrecedencyExtractor.AllDependency
    Logger.Log TRACE_LOG, "Exit LETManagerUI.RecalculateAndUpdateDependencyCollection"

End Sub

Private Sub ExpandButton_Click()

    Logger.Log TRACE_LOG, "Enter LETManagerUI.ExpandButton_Click"
    ' Get the variable name of the selected item in the StepsListBox
    Dim SelectedDependencyVarName As String
    SelectedDependencyVarName = GetSelectedItemVarName(Me.StepsListBox)
    If SelectedDependencyVarName = vbNullString Then Exit Sub

    ' Get the DependencyInfo object that matches the selected variable name
    Dim SelectedDependency As DependencyInfo
    Set SelectedDependency = GetMatchingVarNameDependency(SelectedDependencyVarName _
                                                          , This.DependencyObjects)

    ' Mark the selected dependency as expanded by the user
    With SelectedDependency
        .IsInsideNamedRangeOrTable = False
        .IsReferByNamedRange = False
        .IsExpandByUser = True
    End With

    ' Recalculate and update the DependencyObjects
    RecalculateAndUpdateDependencyCollection
    UpdateListBoxFromCollection
    Logger.Log TRACE_LOG, "Exit LETManagerUI.ExpandButton_Click"

End Sub

Private Sub OkButton_Click()
    Me.Hide
End Sub

'@EntryPoint
Public Sub UpdateListBoxFromCollection()

    Logger.Log TRACE_LOG, "Enter LETManagerUI.UpdateListBoxFromCollection"

    ' Update the ListBox with the Lambda preview based on the DependencyObjects
    Me.Preview.Value = This.Parser.GetLetPreview(This.DependencyObjects)

    ' Update the LetStepsListBox
    UpdateLetStepsListBox

    ' Update the selection if it's the first time
    UpdateSelectionIfForTheFirstTime

    Logger.Log TRACE_LOG, "Exit LETManagerUI.UpdateListBoxFromCollection"

End Sub

Private Sub UpdateSelectionIfForTheFirstTime()

    Logger.Log TRACE_LOG, "Enter LETManagerUI.UpdateSelectionIfForTheFirstTime"
    ' If it's the first time, disable certain buttons and select the first item in a ListBox if available
    If This.Counter = 0 Then
        Me.ResetButton.Enabled = False
        Me.RenameStepButton.Enabled = False
        Me.ExcludeStepButton.Enabled = False
        Me.ValueButton.Enabled = False
        Me.ExpandButton.Enabled = False

        ' Deselect all items
        CustomizeListBox.SelectOptionAllOfListbox Me.StepsListBox, False
        CustomizeListBox.SelectOrDeselectFirstItemOfListbox Me.StepsListBox
    Else
        ' Enable the ResetButton after the first time
        Me.ResetButton.Enabled = True
    End If
    This.Counter = This.Counter + 1
    Logger.Log TRACE_LOG, "Exit LETManagerUI.UpdateSelectionIfForTheFirstTime"

End Sub

Private Sub UpdateLetStepsListBox()

    Logger.Log TRACE_LOG, "Enter LETManagerUI.UpdateLetStepsListBox"
    ' Update the LetStepsListBox with non-input Let steps' variable names and range references
    Dim VarsName As Variant
    VarsName = GetLetStepsVarNameAndRangeReference(This.DependencyObjects)

    ' Clear the ListBox if VarsName is not an array
    If Not IsArray(VarsName) Then
        Me.StepsListBox.Clear
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword ParamSelector.UpdateLetStepsListBox"
        Exit Sub
    End If

    ' Populate the ListBox with the variable names and range references
    Me.StepsListBox.List = VarsName
    Logger.Log TRACE_LOG, "Exit LETManagerUI.UpdateLetStepsListBox"

End Sub

Private Sub SelectLastFocusRange(ByVal ForListBox As MSForms.ListBox)

    Logger.Log TRACE_LOG, "Enter LETManagerUI.SelectLastFocusRange"

    ' Select the last focused range in the given ListBox
    Dim FirstSelectedIndex As Long
    FirstSelectedIndex = FindFirstSelectedIndex(ForListBox)
    If FirstSelectedIndex = -1 Then Exit Sub

    Dim RangeReference As String
    RangeReference = ForListBox.List(FirstSelectedIndex, 1)
    Dim FocusAbleRange As Range

    On Error Resume Next
    Set FocusAbleRange = RangeResolver.GetRange(RangeReference)
    If FocusAbleRange.Address <> Selection.Address Then
        FocusAbleRange.Parent.Activate
        FocusAbleRange.Cells(1).Select
    End If
    Application.ScreenUpdating = True
    On Error GoTo 0

    Logger.Log TRACE_LOG, "Exit LETManagerUI.SelectLastFocusRange"

End Sub

Private Function FindFirstSelectedIndex(ByVal ForListBox As MSForms.ListBox) As Long

    Logger.Log TRACE_LOG, "Enter LETManagerUI.FindFirstSelectedIndex"
    Dim Index As Long
    For Index = 0 To ForListBox.ListCount - 1
        If ForListBox.Selected(Index) Then
            FindFirstSelectedIndex = Index
            Exit Function
        End If
    Next Index

    FindFirstSelectedIndex = -1
    Logger.Log TRACE_LOG, "Exit LETManagerUI.FindFirstSelectedIndex"

End Function

Private Sub RenameForListBox(ByVal ForListBox As MSForms.ListBox)

    Logger.Log TRACE_LOG, "Enter LETManagerUI.RenameForListBox"

    ' Get the selected row index in the ListBox
    Dim SelectedRowIndex As Long
    SelectedRowIndex = ForListBox.ListIndex

    ' Get the variable name of the selected item in the ListBox
    Dim SelectedDependencyVarName As String
    SelectedDependencyVarName = GetSelectedItemVarName(ForListBox)

    ' Update the variable name and valid variable name for the selected item
    UpdateForNewName SelectedDependencyVarName

    ' Recalculate and update the DependencyObjects
    RecalculateAndUpdateDependencyCollection

    ' Update the ListBox based on the updated collection
    UpdateListBoxFromCollection

    ' Select the row back in the ListBox if it was previously selected
    If SelectedRowIndex <> -1 Then ForListBox.Selected(SelectedRowIndex) = True

    Logger.Log TRACE_LOG, "Exit LETManagerUI.RenameForListBox"

End Sub

Private Sub UpdateForNewName(ByVal SelectedDependencyVarName As String)

    Logger.Log TRACE_LOG, "Enter LETManagerUI.UpdateForNewName"
    ' Check if a valid item is selected in the ListBox
    If SelectedDependencyVarName = vbNullString Then Exit Sub

    ' Get the DependencyInfo for the selected item
    Dim SelectedDependency As DependencyInfo
    Set SelectedDependency = GetMatchingVarNameDependency(SelectedDependencyVarName _
                                                          , This.DependencyObjects)

    ' Prompt the user to enter a new name for the variable
    Dim NewName As String
    NewName = InputBox("Enter new name:", "Parameter/Step Name", SelectedDependencyVarName)

    ' Check if the user entered a new name or canceled the input box
    If NewName = vbNullString Or NewName = "False" Then Exit Sub

    ' Update the range label and valid variable name for the selected item
    With SelectedDependency
        .RangeLabel = NewName
        .IsUserSpecifiedName = True
        .ValidVarName = ConvertToValidLetVarName(.RangeLabel)
    End With
    Logger.Log TRACE_LOG, "Exit LETManagerUI.UpdateForNewName"

End Sub

Private Sub RenameStepButton_Click()

    Logger.Log TRACE_LOG, "Enter LETManagerUI.RenameStepButton_Click"
    ' Call RenameForListBox for StepsListBox
    RenameForListBox Me.StepsListBox
    Logger.Log TRACE_LOG, "Exit LETManagerUI.RenameStepButton_Click"

End Sub

Private Sub ResetButton_Click()

    Logger.Log TRACE_LOG, "Enter LETManagerUI.ResetButton_Click"

    ' Reset the DependencyObjects to the initial state
    Set This.DependencyObjects = This.Parser.DependencyDataForReset
    This.Counter = 0

    ' Recalculate and update the DependencyObjects
    RecalculateAndUpdateDependencyCollection

    ' Update the ListBoxes based on the updated collection
    UpdateListBoxFromCollection
    Logger.Log TRACE_LOG, "Exit LETManagerUI.ResetButton_Click"

End Sub

Private Sub StepsListBox_Change()

    Logger.Log TRACE_LOG, "Enter LETManagerUI.StepsListBox_Change"

    Dim SelectedItemCount As Long
    SelectedItemCount = NumberOfItemSelected(Me.StepsListBox)
    Dim IsEnableExceptExcludeButton As Boolean
    Dim IsEnableExcludeButton As Boolean

    If SelectedItemCount = 0 Then

        IsEnableExceptExcludeButton = False
        IsEnableExcludeButton = False

    ElseIf SelectedItemCount = 1 Then

        IsEnableExceptExcludeButton = True
        IsEnableExcludeButton = True

    ElseIf SelectedItemCount > 1 Then

        IsEnableExceptExcludeButton = False
        IsEnableExcludeButton = True

    End If

    ' Enable or disable buttons based on selection
    Me.RenameStepButton.Enabled = IsEnableExceptExcludeButton
    Me.ExcludeStepButton.Enabled = IsEnableExcludeButton
    Me.ValueButton.Enabled = IsEnableExceptExcludeButton

    ' Enable or disable the ExpandButton based on the selected item in the ListBox
    EnableOrDisableExpandButton SelectedItemCount

    ' Select the last focused range in the ListBox
    SelectLastFocusRange Me.StepsListBox

    ' Disable the ValueButton and MakeParamButton if the last item is selected in the ListBox
    If SelectedItemCount = 1 And Me.StepsListBox.ListIndex = Me.StepsListBox.ListCount - 1 Then
        Me.ValueButton.Enabled = False
    End If

    Logger.Log TRACE_LOG, "Exit LETManagerUI.StepsListBox_Change"

End Sub

Private Function NumberOfItemSelected(ByVal ForListBox As MSForms.ListBox) As Long

    Logger.Log TRACE_LOG, "Enter LETManagerUI.NumberOfItemSelected"
    Dim SelectedItemCount As Long
    If ForListBox.ListIndex = -1 Then
        SelectedItemCount = 0
    Else
        Dim Index As Long
        For Index = 0 To ForListBox.ListCount - 1
            If ForListBox.Selected(Index) Then
                SelectedItemCount = SelectedItemCount + 1
            End If
        Next Index
    End If

    NumberOfItemSelected = SelectedItemCount
    Logger.Log TRACE_LOG, "Exit LETManagerUI.NumberOfItemSelected"

End Function

Private Sub EnableOrDisableExpandButton(SelectedItemCount As Long)

    Logger.Log TRACE_LOG, "Enter LETManagerUI.EnableOrDisableExpandButton"

    If SelectedItemCount = 0 Or SelectedItemCount > 1 Then
        Me.ExpandButton.Enabled = False
        Exit Sub
    End If

    ' Get the variable name of the selected item in the StepsListBox
    Dim SelectedDependencyVarName As String
    SelectedDependencyVarName = GetSelectedItemVarName(Me.StepsListBox)
    If SelectedDependencyVarName = vbNullString Then
        ' Disable the ExpandButton if no item is selected
        Me.ExpandButton.Enabled = False
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword ParamSelector.EnableOrDisableExpandButton"
        Exit Sub
    End If

    ' Get the DependencyInfo for the selected item
    Dim SelectedDependency As DependencyInfo
    Set SelectedDependency = GetMatchingVarNameDependency(SelectedDependencyVarName _
                                                          , This.DependencyObjects)

    With SelectedDependency

        ' Check if the ExpandButton should be enabled or disabled based on the selected item's properties
        If .IsUserMarkAsValue Then
            Me.ExpandButton.Enabled = False
            Logger.Log TRACE_LOG, "Exit Due to Exit Keyword ParamSelector.EnableOrDisableExpandButton"
            Exit Sub
        ElseIf Not .IsInsideNamedRangeOrTable And Not .IsDemotedFromParameterCellToLetStep Then
            Me.ExpandButton.Enabled = False
            Logger.Log TRACE_LOG, "Exit Due to Exit Keyword ParamSelector.EnableOrDisableExpandButton"
            Exit Sub
        End If

        Me.ExpandButton.Enabled = IsExpandAble(RangeResolver.GetRange(.RangeReference))

    End With

    Logger.Log TRACE_LOG, "Exit LETManagerUI.EnableOrDisableExpandButton"

End Sub

Private Sub UserForm_Activate()

    Logger.Log TRACE_LOG, "Enter LETManagerUI.UserForm_Activate"
    ' Set the initial height and width of the UserForm
    Me.Height = 434
    Me.Width = 818
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) * 0.5
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    Logger.Log TRACE_LOG, "Exit LETManagerUI.UserForm_Activate"

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Logger.Log TRACE_LOG, "Enter LETManagerUI.UserForm_QueryClose"
    ' Hide the UserForm and set IsProcessTerminatedByUser to True on close by user
    If CloseMode = CloseBy.User Then
        This.Parser.IsProcessTerminatedByUser = True
        Me.Hide
        Cancel = True
    End If
    Logger.Log TRACE_LOG, "Exit LETManagerUI.UserForm_QueryClose"

End Sub

Private Sub ValueButton_Click()

    Logger.Log TRACE_LOG, "Enter LETManagerUI.ValueButton_Click"
    ' Handles the click event of the ValueButton.
    Dim SelectedDependencyVarName As String
    SelectedDependencyVarName = GetSelectedItemVarName(Me.StepsListBox)

    ' Check if any item is selected in the Steps ListBox
    If SelectedDependencyVarName = vbNullString Then Exit Sub

    Dim SelectedDependency As DependencyInfo
    Set SelectedDependency = GetMatchingVarNameDependency(SelectedDependencyVarName _
                                                          , This.DependencyObjects)

    ' Check if the item is already marked as a "Value" step, if yes, exit the sub.
    If SelectedDependency.IsUserMarkAsValue Then Exit Sub

    ' Update the selected DependencyInfo as a "Value" step and set its properties accordingly.
    With SelectedDependency

        .IsUserMarkAsValue = True

        Dim ResolvedRange As Range
        Set ResolvedRange = RangeResolver.GetRange(.RangeReference)

        Dim FormulaText As String
        If IsNothing(ResolvedRange) And .IsReferByNamedRange Then
            '@TODO: What if i have error on const named range.
            FormulaText = modUtility.ConvertToValueFormula(Evaluate(.NameInFormula))
        Else
            FormulaText = modUtility.ConvertToValueFormula(ResolvedRange.Value)
        End If

        ' Check if the cell value can be treated as an array constant.
        ' If it can, mark it as a formula, else, mark it as a constant value.
        If Left$(FormulaText, 1) = LEFT_BRACE Then
            .FormulaText = EQUAL_SIGN & FormulaText
            .HasFormula = True
        Else
            .FormulaText = FormulaText
            .HasFormula = False
        End If

        ' Since it is a "Value" step, it has no dependencies.
        .HasAnyDependency = False
    End With

    ' Update the DependencyObjects and the ListBox.
    RecalculateAndUpdateDependencyCollection
    UpdateListBoxFromCollection
    Logger.Log TRACE_LOG, "Exit LETManagerUI.ValueButton_Click"

End Sub

