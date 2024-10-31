VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ParamSelector 
   Caption         =   "Lambda Parameter Setup"
   ClientHeight    =   8120
   ClientLeft      =   -360
   ClientTop       =   -1755
   ClientWidth     =   16140
   OleObjectBlob   =   "ParamSelector.frx":0000
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "ParamSelector"
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

Private Type TParamSelector
    Parser As FormulaParser
    DependencyObjects As Collection
    Counter As Long
    LastActiveListBox As MSForms.ListBox
End Type

Private This As TParamSelector

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
    Logger.Log TRACE_LOG, "Enter ParamSelector.CancelButton_Click"
    This.Parser.IsProcessTerminatedByUser = True
    Me.Hide
    Logger.Log TRACE_LOG, "Exit ParamSelector.CancelButton_Click"
End Sub

Private Sub MakeOptionalButton_Click()
    
    If Me.ParametersListBox.ListIndex = -1 Then Exit Sub
    If Me.ParametersListBox.ListCount = 0 Then Exit Sub
    
    Dim Index As Long
    Dim SelectedDependencyVarName As String
    For Index = 0 To Me.ParametersListBox.ListCount - 1
        
        SelectedDependencyVarName = GetItemVarName(Me.ParametersListBox, Index)
        ' Get the DependencyInfo object that matches the selected variable name
        Dim SelectedDependency As DependencyInfo
        Set SelectedDependency = GetMatchingVarNameDependency(SelectedDependencyVarName _
                                                              , This.DependencyObjects)
        SelectedDependency.IsOptional = (Index >= Me.ParametersListBox.ListIndex)
    Next Index
    
    RecalculateAndUpdateDependencyCollection
    UpdateListBoxFromCollection
    
End Sub

Private Sub MakeStepButton_Click()
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.Exclude_Click"
    ' Get the variable name of the selected item in the ParametersListBox
    Dim SelectedDependencyVarName As String
    
    SelectedDependencyVarName = GetSelectedItemVarName(Me.ParametersListBox)
    If SelectedDependencyVarName = vbNullString Then Exit Sub
    
    ' Get the DependencyInfo object that matches the selected variable name
    Dim SelectedDependency As DependencyInfo
    Set SelectedDependency = GetMatchingVarNameDependency(SelectedDependencyVarName _
                                                          , This.DependencyObjects)
    
    ' Exclude the selected dependency from being treated as a parameter
    With SelectedDependency
        .IsLabelAsInputCell = False
        .IsDemotedFromParameterCellToLetStep = True
    End With
    
    ' Update the ParametersListBox after the exclusion
    UpdateListBoxAfterExclude Me.ParametersListBox
    EnableOrDisableExpandButton NumberOfItemSelected(Me.StepsListBox)
    Logger.Log TRACE_LOG, "Exit ParamSelector.Exclude_Click"
    
End Sub

Private Sub SelectAgainAfterExclude(ByVal ForListBox As MSForms.ListBox, ByVal SelectedRowIndex As Long)
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.SelectAgainAfterExclude"
    
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
    
    Logger.Log TRACE_LOG, "Exit ParamSelector.SelectAgainAfterExclude"
    
End Sub

Private Function GetSelectedItemVarName(ByVal ForListBox As MSForms.ListBox) As String
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.GetSelectedItemVarName"
    ' Get the variable name of the selected item in the given ListBox
    If ForListBox.ListIndex = -1 Then Exit Function
    GetSelectedItemVarName = GetItemVarName(ForListBox, ForListBox.ListIndex)
    Logger.Log TRACE_LOG, "Exit ParamSelector.GetSelectedItemVarName"
    
End Function

Private Function GetItemVarName(ByVal ForListBox As MSForms.ListBox, ByVal FromIndex As Long) As String
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.GetItemVarName"
    ' Get the variable name of the item from the given ListBox at the specified index
    GetItemVarName = ForListBox.List(FromIndex, 0)
    Logger.Log TRACE_LOG, "Exit ParamSelector.GetItemVarName"
    
End Function

Private Sub ExcludeStepButton_Click()
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.ExcludeStepButton_Click"
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
    Logger.Log TRACE_LOG, "Exit ParamSelector.ExcludeStepButton_Click"
    
End Sub

Private Sub UpdateListBoxAfterExclude(ByVal ForListBox As MSForms.ListBox)
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.UpdateListBoxAfterExclude"
    ' Update the ListBox after excluding a dependency
    Dim SelectedRowIndex As Long
    SelectedRowIndex = ForListBox.ListIndex
    RecalculateAndUpdateDependencyCollection
    UpdateListBoxFromCollection
    SelectAgainAfterExclude ForListBox, SelectedRowIndex
    Logger.Log TRACE_LOG, "Exit ParamSelector.UpdateListBoxAfterExclude"
    
End Sub

Private Sub RecalculateAndUpdateDependencyCollection()
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.RecalculateAndUpdateDependencyCollection"
    ' Recalculate the precedence and update the DependencyObjects
    This.Parser.RecalculatePrecedencyAgain This.DependencyObjects, LAMBDA_STATEMENT_GENERATION
    Set This.DependencyObjects = This.Parser.PrecedencyExtractor.AllDependency
    Logger.Log TRACE_LOG, "Exit ParamSelector.RecalculateAndUpdateDependencyCollection"
    
End Sub

Private Sub ExpandButton_Click()
        
    Logger.Log TRACE_LOG, "Enter ParamSelector.ExpandButton_Click"
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
    Logger.Log TRACE_LOG, "Exit ParamSelector.ExpandButton_Click"
    
End Sub

Private Sub MakeParamButton_Click()
        
    Logger.Log TRACE_LOG, "Enter ParamSelector.MakeParamButton_Click"
    ' Get the variable name of the selected item in the StepsListBox
    Dim SelectedDependencyVarName As String
    SelectedDependencyVarName = GetSelectedItemVarName(Me.StepsListBox)
    If SelectedDependencyVarName = vbNullString Then Exit Sub
    
    ' Get the DependencyInfo object that matches the selected variable name
    Dim SelectedDependency As DependencyInfo
    Set SelectedDependency = GetMatchingVarNameDependency(SelectedDependencyVarName _
                                                          , This.DependencyObjects)
    
    ' Mark the selected dependency as a parameter cell
    With SelectedDependency
        .IsLabelAsInputCell = True
        .IsMarkAsNotLetStatementByUser = False
        .IsUserMarkAsParameterCell = True
        .IsDemotedFromParameterCellToLetStep = False
    End With
    
    ' Update the StepsListBox after the exclusion
    UpdateListBoxAfterExclude Me.StepsListBox
    Logger.Log TRACE_LOG, "Exit ParamSelector.MakeParamButton_Click"
    
End Sub

Private Sub OkButton_Click()
    Me.Hide
End Sub

'@EntryPoint
Public Sub UpdateListBoxFromCollection()
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.UpdateListBoxFromCollection"
    
    ' Update the ListBox with the Lambda preview based on the DependencyObjects
    Me.Preview.Value = This.Parser.GetLambdaPreview(This.DependencyObjects)
    
    ' Update the ParametersListBox
    UpdateParametersListBox
    
    ' Update the LetStepsListBox
    UpdateLetStepsListBox
    
    ' Update the selection if it's the first time
    UpdateSelectionIfForTheFirstTime
    
    Logger.Log TRACE_LOG, "Exit ParamSelector.UpdateListBoxFromCollection"
    
End Sub

Private Sub UpdateSelectionIfForTheFirstTime()
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.UpdateSelectionIfForTheFirstTime"
    ' If it's the first time, disable certain buttons and select the first item in a ListBox if available
    If This.Counter = 0 Then
        Me.ResetButton.Enabled = False
        Me.RenameParamButton.Enabled = False
        Me.MakeStepButton.Enabled = False
        Me.RenameStepButton.Enabled = False
        Me.ExcludeStepButton.Enabled = False
        Me.MakeParamButton.Enabled = False
        Me.ValueButton.Enabled = False
        Me.ExpandButton.Enabled = False
        
        ' Deselect all items in both ListBoxes
        CustomizeListBox.SelectOptionAllOfListbox Me.ParametersListBox, False
        CustomizeListBox.SelectOptionAllOfListbox Me.StepsListBox, False
        ' Select the first item in ParametersListBox if available, otherwise, select the first item in StepsListBox if available
        If Me.ParametersListBox.ListCount > 0 Then
            CustomizeListBox.SelectOrDeselectFirstItemOfListbox Me.ParametersListBox
        ElseIf Me.StepsListBox.ListCount > 0 Then
            CustomizeListBox.SelectOrDeselectFirstItemOfListbox Me.StepsListBox
        End If
    Else
        ' Enable the ResetButton after the first time
        Me.ResetButton.Enabled = True
    End If
    This.Counter = This.Counter + 1
    Logger.Log TRACE_LOG, "Exit ParamSelector.UpdateSelectionIfForTheFirstTime"
    
End Sub

Private Sub UpdateLetStepsListBox()
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.UpdateLetStepsListBox"
    ' Update the LetStepsListBox with non-input Let steps' variable names and range references
    Dim VarsName As Variant
    If This.Parser.IsLetNeededInLambda Then
        VarsName = GetNonInputLetStepsVarNameAndRangeReference(This.DependencyObjects)
    End If
    
    ' Clear the ListBox if VarsName is not an array
    If Not IsArray(VarsName) Then
        Me.StepsListBox.Clear
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword ParamSelector.UpdateLetStepsListBox"
        Exit Sub
    End If
    
    ' Populate the ListBox with the variable names and range references
    Me.StepsListBox.List = VarsName
    TryAdaptingScrollBarHeight Me.StepsListBox
    Logger.Log TRACE_LOG, "Exit ParamSelector.UpdateLetStepsListBox"
    
End Sub

Private Sub UpdateParametersListBox()
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.UpdateParametersListBox"
    
    ' Update the ParametersListBox with input cells' variable names and range references
    Dim VarsName As Variant
    VarsName = GetInputCellsVarNameAndRangeReference(This.DependencyObjects)
    
    ' Clear the ListBox if VarsName is not an array
    If Not IsArray(VarsName) Then
        Me.ParametersListBox.Clear
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword ParamSelector.UpdateParametersListBox"
        Exit Sub
    End If
    
    ' Populate the ListBox with the variable names and range references
    Me.ParametersListBox.List = VarsName
    TryAdaptingScrollBarHeight Me.ParametersListBox
    Logger.Log TRACE_LOG, "Exit ParamSelector.UpdateParametersListBox"
    
End Sub

Private Sub ParametersListBox_Change()
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.ParametersListBox_Change"
    
    ' Check if any item is selected in the ParametersListBox
    Dim IsAnyItemSelected As Boolean
    IsAnyItemSelected = (Me.ParametersListBox.ListIndex <> -1)
    
    ' Enable or disable buttons based on selection
    Me.RenameParamButton.Enabled = IsAnyItemSelected
    Me.MakeStepButton.Enabled = IsAnyItemSelected
    Me.MakeOptionalButton.Enabled = IsAnyItemSelected
    
    ' Select the last focused range in the ListBox
    SelectLastFocusRange Me.ParametersListBox
    Logger.Log TRACE_LOG, "Exit ParamSelector.ParametersListBox_Change"
    
End Sub

Private Sub SelectLastFocusRange(ByVal ForListBox As MSForms.ListBox)
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.SelectLastFocusRange"
    
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
        FocusAbleRange.Worksheet.Activate
        FocusAbleRange.Cells(1).Select
    End If
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    Logger.Log TRACE_LOG, "Exit ParamSelector.SelectLastFocusRange"
    
End Sub

Private Function FindFirstSelectedIndex(ByVal ForListBox As MSForms.ListBox) As Long
    
    Dim Index As Long
    For Index = 0 To ForListBox.ListCount - 1
        If ForListBox.Selected(Index) Then
            FindFirstSelectedIndex = Index
            Exit Function
        End If
    Next Index
    
    FindFirstSelectedIndex = -1
    
End Function

Private Sub RenameParamButton_Click()
    RenameForListBox Me.ParametersListBox
End Sub

Private Sub RenameForListBox(ByVal ForListBox As MSForms.ListBox)
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.RenameForListBox"
    
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
    
    Logger.Log TRACE_LOG, "Exit ParamSelector.RenameForListBox"
    
End Sub

Private Sub UpdateForNewName(ByVal SelectedDependencyVarName As String)
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.UpdateForNewName"
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
    Logger.Log TRACE_LOG, "Exit ParamSelector.UpdateForNewName"
    
End Sub

Private Sub RenameStepButton_Click()
    
    ' Call RenameForListBox for StepsListBox
    RenameForListBox Me.StepsListBox
    
End Sub

Private Sub ResetButton_Click()
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.ResetButton_Click"
    
    ' Reset the DependencyObjects to the initial state
    Set This.DependencyObjects = This.Parser.DependencyDataForReset(LAMBDA_STATEMENT_GENERATION)
    This.Counter = 0
    
    ' Recalculate and update the DependencyObjects
    RecalculateAndUpdateDependencyCollection
    
    ' Update the ListBoxes based on the updated collection
    UpdateListBoxFromCollection
    Logger.Log TRACE_LOG, "Exit ParamSelector.ResetButton_Click"
    
End Sub

Private Sub StepsListBox_Change()
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.StepsListBox_Change"
    
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
    Me.MakeParamButton.Enabled = IsEnableExceptExcludeButton
    Me.ValueButton.Enabled = IsEnableExceptExcludeButton
    
    ' Enable or disable the ExpandButton based on the selected item in the ListBox
    EnableOrDisableExpandButton SelectedItemCount
    
    ' Select the last focused range in the ListBox
    SelectLastFocusRange Me.StepsListBox
    
    ' Disable the ValueButton and MakeParamButton if the last item is selected in the ListBox
    If SelectedItemCount = 1 And Me.StepsListBox.ListIndex = Me.StepsListBox.ListCount - 1 Then
        Me.ValueButton.Enabled = False
        Me.MakeParamButton.Enabled = False
    End If
    
    Logger.Log TRACE_LOG, "Exit ParamSelector.StepsListBox_Change"
    
End Sub

Private Function NumberOfItemSelected(ByVal ForListBox As MSForms.ListBox) As Long
    
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
    
End Function

Private Sub EnableOrDisableExpandButton(SelectedItemCount As Long)
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.EnableOrDisableExpandButton"
    
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
    
    Logger.Log TRACE_LOG, "Exit ParamSelector.EnableOrDisableExpandButton"
    
End Sub

Private Sub UpButton_Click()
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.UpButton_Click"
    ' In valid cases
    If This.Counter = 0 Then Exit Sub
    If Me.ParametersListBox.ListIndex <= 0 Then Exit Sub
    
    ' Get the selected and previous DependencyInfo items
    Dim SelectedItem As DependencyInfo
    Set SelectedItem = GetMatchingVarNameDependency(GetItemVarName(Me.ParametersListBox _
    , Me.ParametersListBox.ListIndex), This.DependencyObjects)
    
    Dim PreviousItem As DependencyInfo
    Set PreviousItem = GetMatchingVarNameDependency(GetItemVarName(Me.ParametersListBox _
    , Me.ParametersListBox.ListIndex - 1), This.DependencyObjects)
    
    ' Check if the move is valid, and if not, show a message box
    If (SelectedItem.IsOptional And Not PreviousItem.IsOptional) Then
        PreviousItem.IsOptional = True
    End If
    
    ' Move the selected item up in the DependencyObjects
    This.DependencyObjects.Remove SelectedItem.RangeReference
    This.DependencyObjects.Add SelectedItem, SelectedItem.RangeReference, PreviousItem.RangeReference
    
    ' Recalculate and update the DependencyObjects
    RecalculateAndUpdateDependencyCollection
    
    ' Update the ListBoxes based on the updated collection
    UpdateListBoxFromCollection
    Logger.Log TRACE_LOG, "Exit ParamSelector.UpButton_Click"
    
End Sub

Private Sub DownButton_Click()

    Logger.Log TRACE_LOG, "Enter ParamSelector.DownButton_Click"
    ' In valid cases
    If This.Counter = 0 Then Exit Sub
    If Me.ParametersListBox.ListIndex = -1 Then Exit Sub
    If Me.ParametersListBox.ListIndex = Me.ParametersListBox.ListCount - 1 Then Exit Sub
    
    ' Get the selected and next DependencyInfo items
    Dim SelectedItem As DependencyInfo
    Set SelectedItem = GetMatchingVarNameDependency(GetItemVarName(Me.ParametersListBox _
    , Me.ParametersListBox.ListIndex), This.DependencyObjects)
    
    Dim NextItem As DependencyInfo
    Set NextItem = GetMatchingVarNameDependency(GetItemVarName(Me.ParametersListBox _
    , Me.ParametersListBox.ListIndex + 1), This.DependencyObjects)
    
    ' If next item is option and we are moving down then make the SelectedItem optional as well.
    If (Not SelectedItem.IsOptional And NextItem.IsOptional) Then
        SelectedItem.IsOptional = True
    End If
    
    ' Move the selected item down in the DependencyObjects
    This.DependencyObjects.Remove NextItem.RangeReference
    This.DependencyObjects.Add NextItem, NextItem.RangeReference, SelectedItem.RangeReference
    
    ' Update the ListBoxes based on the updated collection
    UpdateListBoxFromCollection
    Logger.Log TRACE_LOG, "Exit ParamSelector.DownButton_Click"
    
End Sub

Private Sub UserForm_Activate()
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.UserForm_Activate"
    ' Set the initial height and width of the UserForm
    Me.Height = 434
    Me.Width = 818
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) * 0.5
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    Logger.Log TRACE_LOG, "Exit ParamSelector.UserForm_Activate"

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.UserForm_QueryClose"
    ' Hide the UserForm and set IsProcessTerminatedByUser to True on close by user
    If CloseMode = CloseBy.User Then
        This.Parser.IsProcessTerminatedByUser = True
        Me.Hide
        Cancel = True
    End If
    Logger.Log TRACE_LOG, "Exit ParamSelector.UserForm_QueryClose"
    
End Sub

Private Sub ValueButton_Click()
    
    Logger.Log TRACE_LOG, "Enter ParamSelector.ValueButton_Click"
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
    Logger.Log TRACE_LOG, "Exit ParamSelector.ValueButton_Click"
    
End Sub


