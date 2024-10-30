Attribute VB_Name = "modUserParameter"
'@IgnoreModule UndeclaredVariable
'@Folder "ParameterTaker"
' @IgnoreModule ImplicitActiveSheetReference
Option Explicit
Option Private Module

Public Function GetOrAskForParameterValue(ByVal ParameterCell As Range _
                                          , ByVal ParameterPrompt As String _
                                           , ByVal ParameterType As String) As String
    
    Logger.Log TRACE_LOG, "Enter modUserParameter.GetOrAskForParameterValue"
    ' If the ParameterCell is not empty, return its text value and exit the function
    If ParameterCell.Value <> vbNullString Then
        GetOrAskForParameterValue = ParameterCell.Text
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modUserParameter.GetOrAskForParameterValue"
        Exit Function
    End If
    
    Dim UserGivenValue As Variant
    
    If ParameterType = "List" Then
        ' If ParameterCell has list validation, obtain the validation list from the cell
        If ParameterCell.Validation.Type = XlDVType.xlValidateList Then
            Dim ValidationList As Variant
            ValidationList = GetValidationListFromRange(ParameterCell)
            ' Prompt the user to select a value from the validation list
            UserGivenValue = GetSelectedValueFromList(ParameterPrompt, ValidationList)
            ' If user doesn't select any value (i.e., returns False), clear the ParameterCell and exit the function
            If UserGivenValue = False Then
                ParameterCell.ClearContents
                GetOrAskForParameterValue = vbNullString
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modUserParameter.GetOrAskForParameterValue"
                Exit Function
            End If
        Else
            ' If ParameterCell doesn't have list validation, treat the parameter type as "Text"
            ParameterType = "Text"
        End If
    End If
    
    ' If ParameterType is "Range", prompts user for a range input
    If ParameterType = "Range" Then
        ' Ignore any error temporarily
        On Error Resume Next
        ' Prompt user for a range input
        Set UserGivenValue = Application.InputBox(ParameterPrompt, "Range Selector", Type:=8)
        ' If user does not provide input, clear ParameterCell and return an empty string, then exit function
        If Not UserGivenValue Is Nothing Then
            ParameterCell.ClearContents
            GetOrAskForParameterValue = vbNullString
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modUserParameter.GetOrAskForParameterValue"
            Exit Function
        End If
        ' Reset error handling to default behavior
        On Error GoTo 0
        ' If ParameterType is not "List", prompts user for a text input
    ElseIf ParameterType <> "List" Then
        ' Prompt user for a text input
        UserGivenValue = Application.InputBox(ParameterPrompt, "List Selector", Type:=10)
        ' If user does not provide input, clear ParameterCell and return an empty string, then exit function
        If UserGivenValue = False Then
            ParameterCell.ClearContents
            GetOrAskForParameterValue = vbNullString
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modUserParameter.GetOrAskForParameterValue"
            Exit Function
        End If
    End If

    
    ' Based on ParameterType, handle input differently
    Select Case ParameterType
            ' For "Text" type, no validation needed, directly assign value
        Case "Text"
            ParameterCell.Value = UserGivenValue
            ' For "Date" type, ensure UserGivenValue can be interpreted as a date
        Case "Date"
            ' Check if it is a probable date, a whole number, and > 1
            If IsProbablyDate(UserGivenValue) And IsWholeNumber(UserGivenValue) And CLng(UserGivenValue) > 1 Then
                ParameterCell.Value = UserGivenValue
            Else
                ' User input not valid date, provide a prompt and re-call the function
                MsgBox "You haven't given a proper Date", vbOKOnly, "Date Taker"
                GetOrAskForParameterValue = GetOrAskForParameterValue(ParameterCell _
                                                                      , ParameterPrompt, ParameterType)
            End If
            ' For "Date/Time" type, ensure UserGivenValue can be interpreted as a date/time
        Case "Date/Time"
            ' Check if it is a probable date, a whole or decimal number, and > 1
            If IsProbablyDate(UserGivenValue) And (IsWholeNumber(UserGivenValue) Or _
                                                   IsDecialNumber(UserGivenValue)) And CLng(UserGivenValue) > 1 Then
                ParameterCell.Value = UserGivenValue
            Else
                ' User input not valid date/time, provide a prompt and re-call the function
                MsgBox "You haven't given a proper Date/Time", vbOKOnly, "Date/Time Taker"
                GetOrAskForParameterValue = GetOrAskForParameterValue(ParameterCell _
                                                                      , ParameterPrompt, ParameterType)
            End If
            ' For "Time" type, ensure UserGivenValue can be interpreted as a time
        Case "Time"
            ' Check if it is a probable date, a decimal number, and < 1
            If IsProbablyDate(UserGivenValue) And IsDecialNumber(UserGivenValue) And CLng(UserGivenValue) < 1 Then
                ParameterCell.Value = UserGivenValue
            Else
                ' User input not valid time, provide a prompt and re-call the function
                MsgBox "You haven't given a proper Time", vbOKOnly, "Time Taker"
                GetOrAskForParameterValue = GetOrAskForParameterValue(ParameterCell _
                                                                      , ParameterPrompt, ParameterType)
            End If
            ' For "Integer" type, ensure UserGivenValue can be interpreted as an integer
        Case "Integer"
            ' Check if it is a whole number
            If IsWholeNumber(UserGivenValue) Then
                ParameterCell.Value = UserGivenValue
            Else
                ' User input not valid integer, provide a prompt and re-call the function
                MsgBox "You haven't given a proper Integer number", vbOKOnly, "Integer Taker"
                GetOrAskForParameterValue = GetOrAskForParameterValue(ParameterCell _
                                                                      , ParameterPrompt, ParameterType)
            End If
            ' For "Decimal" type, ensure UserGivenValue can be interpreted as a decimal
        Case "Decimal"
            ' Check if it is a decimal number
            If IsDecialNumber(UserGivenValue) Then
                ParameterCell.Value = UserGivenValue
            Else
                ' User input not valid decimal, provide a prompt and re-call the function
                MsgBox "You haven't given a proper Decimal", vbOKOnly, "Decimal Taker"
                GetOrAskForParameterValue = GetOrAskForParameterValue(ParameterCell _
                                                                      , ParameterPrompt, ParameterType)
            End If
            ' For "Percent" type, ensure UserGivenValue can be interpreted as a number
        Case "Percent"
            ' Check if it is a numeric value
            If IsNumeric(UserGivenValue) Then
                ' Convert the percentage to a decimal
                ParameterCell.Value = UserGivenValue / 100
            Else
                ' User input not valid percent, provide a prompt and re-call the function
                MsgBox "You haven't given a proper percent", vbOKOnly, "Percent Taker"
                GetOrAskForParameterValue = GetOrAskForParameterValue(ParameterCell _
                                                                      , ParameterPrompt, ParameterType)
            End If
            ' For "List", "File", "Folder", and "Range" types, directly assign value or call function to retrieve it
        Case "List"
            ParameterCell.Value = UserGivenValue
        Case "File"
            ParameterCell.Value = GetSelectedFilePath(ParameterPrompt, "*")
        Case "Folder"
            ParameterCell.Value = GetSelectedFolderPath(ParameterPrompt)
        Case "Range"
            ' Convert the UserGivenValue to a Range object
            Dim Temp As Range
            Set Temp = UserGivenValue
            ParameterCell.Value = GetRangeRefWithSheetName(Temp)
            ' For "True/False" type, show a MessageBox to get user input
        Case "True/False"
            Dim MessageBoxAnswer As VbMsgBoxResult
            MessageBoxAnswer = MsgBox(ParameterPrompt, vbYesNoCancel, "Boolean Taker")
            ' Based on MessageBox result, set ParameterCell value
            If MessageBoxAnswer = vbYes Then
                ParameterCell.Value = True
            ElseIf MessageBoxAnswer = vbNo Then
                ParameterCell.Value = False
            Else
                ParameterCell.Value = vbNullString
            End If
            ' If none of the above cases, raise an error for wrong ParameterType
        Case Else
            Err.Raise 13, "GetOrAskForParameterValue Function", "Wrong ParameterType"
    End Select
    
    ' Call Recursively for validation.
    If Not ParameterCell.Validation.Value Then
        ParameterCell.Value = vbNullString
        MsgBox "Your input doesn't pass parameter cell validation rule.", vbOKOnly, "Invalid Value"
        GetOrAskForParameterValue = GetOrAskForParameterValue(ParameterCell, ParameterPrompt, ParameterType)
    Else
        GetOrAskForParameterValue = ParameterCell.Text
    End If
    Logger.Log TRACE_LOG, "Exit modUserParameter.GetOrAskForParameterValue"
    
End Function

' Function to get selected value from a list
Private Function GetSelectedValueFromList(ByVal PromptTitle As String _
                                          , ByVal ValidationList As Variant) As Variant
    
    Logger.Log TRACE_LOG, "Enter modUserParameter.GetSelectedValueFromList"
    ' Create a new instance of ListItemPicker UserForm
    Dim UF As ListItemPicker
    Set UF = New ListItemPicker
    ' Set the prompt for the UserForm
    UF.PromptLabel.Caption = PromptTitle
    ' Set the list items in the UserForm
    UF.ValidationListItems.List = ValidationList
    TryAdaptingScrollBarHeight UF.ValidationListItems
    ' Display the UserForm
    UF.Show
    ' After the UserForm is closed, get the selected item from the UserForm
    GetSelectedValueFromList = UF.SelectedItem
    ' Unload the UserForm from memory
    Unload UF
    ' Set the UserForm object to Nothing
    Set UF = Nothing
    Logger.Log TRACE_LOG, "Exit modUserParameter.GetSelectedValueFromList"
    
End Function

' Function to get validation list from a range
Private Function GetValidationListFromRange(ByVal GivenCell As Range) As Variant

    Logger.Log TRACE_LOG, "Enter modUserParameter.GetValidationListFromRange"
    ' Check if GivenCell has list validation
    If GivenCell.Validation.Type <> XlDVType.xlValidateList Then Exit Function
    
    ' Get validation formula from GivenCell
    Dim FormulaText As String
    FormulaText = GivenCell.Validation.Formula1
    ' Check if the formula refers to a range or is a comma separated list
    If Text.IsStartsWith(FormulaText, EQUAL_SIGN) Then
        ' If the formula refers to a range, get the values from the range
        Dim ValidationRange As Range
        Set ValidationRange = RangeResolver.GetRange( _
                              Text.RemoveFromStartIfPresent(FormulaText, EQUAL_SIGN) _
                              , GivenCell.Worksheet.Parent _
                               , GivenCell.Worksheet _
                                )
            
        GetValidationListFromRange = ValidationRange.Value
    Else
        ' If the formula is a comma separated list, split it into an array
        GetValidationListFromRange = Application.WorksheetFunction.Transpose(Split(FormulaText, LIST_SEPARATOR))
    End If
    Logger.Log TRACE_LOG, "Exit modUserParameter.GetValidationListFromRange"
    
End Function

Private Function IsProbablyDate(ByVal GivenValue As Variant) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modUserParameter.IsProbablyDate"
    If IsNumeric(GivenValue) Then
        IsProbablyDate = IsDate(CDate(GivenValue))
    Else
        IsProbablyDate = False
    End If
    Logger.Log TRACE_LOG, "Exit modUserParameter.IsProbablyDate"
    
End Function

Private Function IsWholeNumber(ByVal GivenValue As Variant) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modUserParameter.IsWholeNumber"
    On Error GoTo ErrorHandler
    If IsNumeric(GivenValue) Then
        IsWholeNumber = (GivenValue = CStr(Int(GivenValue)))
    End If
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modUserParameter.IsWholeNumber"
    Exit Function
    
ErrorHandler:
    Logger.Log DEBUG_LOG, "Error Details :" & vbNewLine & vbTab & "Number : " & Err.Number _
                         & vbNewLine & vbTab & "Description : " & Err.Description & vbNewLine _
                         & vbTab & "Source : " & Err.Source & vbNewLine & vbTab & "HelpContext : " _
                         & Err.HelpContext & vbNewLine & vbTab & "HelpFile : " & Err.HelpFile _
                         & vbNewLine & vbTab & "LastDllError : " & Err.LastDllError
    Err.Clear
    IsWholeNumber = False
    Logger.Log TRACE_LOG, "Exit modUserParameter.IsWholeNumber"
              
End Function

Private Function IsDecialNumber(ByVal GivenValue As Variant) As Boolean
    Logger.Log TRACE_LOG, "Enter modUserParameter.IsDecialNumber"
    If IsNumeric(GivenValue) Then
        IsDecialNumber = (GivenValue <> Int(GivenValue))
    End If
    Logger.Log TRACE_LOG, "Exit modUserParameter.IsDecialNumber"
End Function

' This will give the selected file path as string.
' Example call : GetSelectedFilePath("Select Correct CSV","*.csv")
Public Function GetSelectedFilePath(ByVal GivenTitle As String, ByVal GivenFilter As String) As String
    
    Logger.Log TRACE_LOG, "Enter modUserParameter.GetSelectedFilePath"
    Dim SelectedFilePath As FileDialogSelectedItems
    Set SelectedFilePath = GetSelectedFilesPath(GivenTitle, GivenFilter, False)
    If SelectedFilePath.Count = 0 Then
        GetSelectedFilePath = vbNullString
    Else
        GetSelectedFilePath = SelectedFilePath.Item(1)
    End If
    Logger.Log TRACE_LOG, "Exit modUserParameter.GetSelectedFilePath"

End Function

' This will give the selected file path as string.
' Example call : GetSelectedFilePath("Select Correct CSV","*.csv",True)
Public Function GetSelectedFilesPath(ByVal GivenTitle As String _
                                     , ByVal GivenFilter As String _
                                      , Optional ByVal IsMultiSelected As Boolean = False) As FileDialogSelectedItems

    Logger.Log TRACE_LOG, "Enter modUserParameter.GetSelectedFilesPath"
    Dim FilePicker As FileDialog
    Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)

    With FilePicker
        .AllowMultiSelect = IsMultiSelected
        .Title = GivenTitle
        .InitialFileName = CurDir$()
        .Filters.Clear
        .Filters.Add "Filter : ", GivenFilter
        .Show
        Set GetSelectedFilesPath = .SelectedItems
    End With
    Logger.Log TRACE_LOG, "Exit modUserParameter.GetSelectedFilesPath"

End Function

' This will give all the selected file path as string.
' Example call : GetAllSelectedFilePath("Select Correct CSV","*.csv")
Public Function GetAllSelectedFilePath(ByVal GivenTitle As String, ByVal GivenFilter As String) As Variant

    Logger.Log TRACE_LOG, "Enter modUserParameter.GetAllSelectedFilePath"
    Const FILE_SELECTION_SUCCESSFUL As Long = -1
    Dim FilePicker As FileDialog
    Set FilePicker = Application.FileDialog(msoFileDialogFilePicker)
    Dim Result As Variant
    With FilePicker
        .AllowMultiSelect = True
        .Title = GivenTitle
        .InitialFileName = CurDir$()
        .Filters.Clear
        .Filters.Add "Filter : ", GivenFilter
        If .Show = FILE_SELECTION_SUCCESSFUL Then
            ReDim Result(1 To .SelectedItems.Count, 1 To 1)
            Dim Counter As Long
            For Counter = 1 To .SelectedItems.Count
                Result(Counter, 1) = .SelectedItems(Counter)
            Next Counter
        End If

    End With
    GetAllSelectedFilePath = Result
    Logger.Log TRACE_LOG, "Exit modUserParameter.GetAllSelectedFilePath"

End Function

' This will give the selected folder path as string.
' Example call : GetSelectedFolderPath("Select Correct Folder")
Public Function GetSelectedFolderPath(ByVal GivenTitle As String) As String

    Logger.Log TRACE_LOG, "Enter modUserParameter.GetSelectedFolderPath"
    Const FOLDER_SELECTION_SUCCESSFUL As Long = -1
    Dim FolderPicker As FileDialog
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FolderPicker
        .AllowMultiSelect = False
        .Title = GivenTitle
        .InitialFileName = CurDir$()

        If .Show = FOLDER_SELECTION_SUCCESSFUL Then
            GetSelectedFolderPath = .SelectedItems(1)
        End If

    End With
    Logger.Log TRACE_LOG, "Exit modUserParameter.GetSelectedFolderPath"

End Function


