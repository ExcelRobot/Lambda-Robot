Attribute VB_Name = "modLambdaEditor"
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed, ImplicitActiveWorkbookReference, UnrecognizedAnnotation, ProcedureNotUsed
'@Folder "Lambda.Editor.XLRobot"
Option Explicit
Option Private Module

'--------------------------------------------< OA HOTKEY >--------------------------------------------
' Command Name:           Edit Lambda
' Description:            Converts a custom Lambda function to it's definition for editing.
' Example:                Converts active cell containing =AddAPlusB(1,2)
'                         with =LAMBDA(_editing_,a,b,a+b)("Lambda: AddAPlusB",1,2)
' Macro Expression:       modLambdaEditor.EditLambda()
' Generated:              03/06/2022 05:43
'-----------------------------------------------------------------------------------------------------
Public Sub EditLambda(ByVal OfCell As Range)
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.EditLambda"
    ' Validate if it's appropriate to run the command
    If IsInvalidToRunCommand(OfCell, "Edit Lambda") Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Dim LambdaName As String
    ' Extract the Lambda function's name
    LambdaName = modUtility.ExtractStartFormulaName(OfCell.Formula2)
    
    If IsLambdaCreatedInExcelLabs(OfCell.Worksheet.Parent, LambdaName) Then
        Dim Answer As VbMsgBoxResult
        Answer = MsgBox("'" & LambdaName & "' lambda is only editable using Excel Labs' Advanced Formula Editor.  Would you like to see it read-only?" _
                        , vbYesNoCancel + vbExclamation, "Edit Lambda")
        
        If Answer = vbNo Or Answer = vbCancel Then
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.EditLambda"
            Exit Sub
        End If
        
    End If
    
    Dim Parameters As String
    
    ' Extract parameters from the formula
    Parameters = Text.AfterDelimiter(OfCell.Formula2, FIRST_PARENTHESIS_OPEN)
    ' If parameters exist, format them
    If Parameters <> vbNullString Then Parameters = FIRST_PARENTHESIS_OPEN & Parameters
    
    ' If name doesn't exist or the named range formula doesn't start with =LAMBDA(, it's not editable.
    Dim IsEditable As Boolean
    IsEditable = modUtility.IsCellHasSavedLambdaFormula(OfCell)
    
    ' If the Lambda function is not editable, proceed to error handler
    If Not IsEditable Then GoTo ErrorHandler
    
    ' Retrieve the existing Lambda function formula
    Dim FormulaText As String
    FormulaText = OfCell.Worksheet.Parent.Names(LambdaName).RefersTo
    
    Dim NewFormulaText As String
    ' Create new formula text by appending parameters to the existing formula
    NewFormulaText = FormatFormula(FormulaText & Parameters)
    
    ' Update the formula in the active cell with the new formula text
    RemoveMetadataAndAddNote OfCell, NewFormulaText, LambdaName
    ' Adjust formula bar to fit new formula text
    AutofitFormulaBar OfCell
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.EditLambda"
    Exit Sub
    
ErrorHandler:

    Dim errorNumber As Long
    errorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description
    
    ' If an error occurred, raise the error and proceed for debugging
    If errorNumber <> 0 Then
        Err.Raise errorNumber, Err.Source, ErrorDescription
        ' This is only for debugging purpose.
        Resume
    End If
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.EditLambda"
    
End Sub

' Critical part of this function is handling Tokenization and updating the formula with error handling.
Private Sub RemoveMetadataAndAddNote(ByVal ToCell As Range, ByVal LambdaFormula As String _
                                                           , ByVal LambdaName As String)
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.RemoveMetadataAndAddNote"
    Dim FormulaText As String
    FormulaText = RemoveMetadataFromFormula(LambdaFormula)
    
    ' Update or Add a note to the cell with the lambda name
    UpdateOrAddLambdaNameNote ToCell, LambdaName, LAMBDA_NAME_NOTE_PREFIX
    
    ' Attempt to update the cell formula with the new formula text.
    ' If any error occurs, it is handled in the 'PutFormulaUsingSendKeys' part.
    On Error GoTo PutFormulaUsingSendKeys
    FormulaText = ReplaceNewlineWithChar10(FormulaText)
    UpdateFormulaAndCalculate ToCell, FormulaText
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.RemoveMetadataAndAddNote"
    Exit Sub
    
PutFormulaUsingSendKeys:
    ' If an error occurs while trying to update the cell formula,
    ' we use SendKeys to put the formula
    PutFormulaWhichHasError ToCell, FormulaText
    ToCell.Calculate
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.RemoveMetadataAndAddNote"
    
End Sub

' Critical path involves handling of the formula input and changing cell's
' format temporarily to insert formula with errors.
Private Sub PutFormulaWhichHasError(ByVal GivenRange As Range, ByVal FormulaText As String)
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.PutFormulaWhichHasError"
    ' Storing the previous format of the cell, so it can be restored later
    Dim PreviousNumberFormat As String
    PreviousNumberFormat = GivenRange.NumberFormat
    
    ' Temporarily change the number format to text.
    ' This is done to prevent Excel from auto-formatting the formula
    GivenRange.NumberFormat = "@"
    
    ' Insert the formula text into the cell. Since the number format is set to text, even formulas with errors can be inserted without issues
    GivenRange.Formula2 = ReplaceInvalidCharFromFormulaWithValid(FormulaText)
    
    ' After inserting the formula, restore the cell's original number format
    GivenRange.NumberFormat = PreviousNumberFormat
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.PutFormulaWhichHasError"
    
End Sub

Public Sub UpdateOrAddLambdaNameNote(ByVal ToCell As Range, ByVal LambdaName As String _
                                                           , ByVal Prefix As String)
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.UpdateOrAddLambdaNameNote"
    On Error Resume Next
    DeleteComment ToCell
    ToCell.Cells(1).AddComment Prefix & LambdaName
    On Error GoTo 0
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.UpdateOrAddLambdaNameNote"
    
End Sub

Public Sub DeleteComment(ByVal ToCell As Range)
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.DeleteComment"
    Dim CurrentComment As Comment
    Set CurrentComment = ToCell.Comment
    On Error GoTo ExitSub
    If Text.IsStartsWith(CurrentComment.Text, LAMBDA_NAME_NOTE_PREFIX) Then
        CurrentComment.Delete
    End If
    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.DeleteComment"
    Exit Sub
    
ExitSub:
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.DeleteComment"
    
End Sub

'--------------------------------------------< OA HOTKEY >--------------------------------------------
' Command Name:           Save Lambda
' Description:            Saves the Lambda definition in the active cell as a defined name.
' Example:                Converts active cell containing =LAMBDA(_editing_,a,b,SUM(a,b))("Lambda: AddAPlusB",3,4)
'                         to =AddAPlusB(3,4) and updates the named range formula for AddAPlusB
'                         to =LAMBDA(a,b,SUM(a,b))
' Macro Expression:       modLambdaEditor.SaveLambda()
' Generated:              03/06/2022 05:48
'-----------------------------------------------------------------------------------------------------
Public Sub SaveLambda(ByVal OfCell As Range)
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.SaveLambda"
    ' Check if the command is valid
    If IsInvalidToRunCommand(OfCell, "Save Lambda") Then Exit Sub
    
    ' Storing the original formula of the cell
    Dim OldFormula As String
    OldFormula = OfCell.Formula2
    
    On Error GoTo ErrorHandler
    Application.StatusBar = "Saving Lambda... (please wait)"
    
    ' Removing any erroneous LET Step Named Ranges
    DeleteLETStepNamedRangesHavingError OfCell.Worksheet.Parent
    
    ' Attempt to generate a lambda function if one is not already present
    TryToGenerateLambdaIfNotGeneratedAlready OfCell
    
    ' Save and update the lambda metadata
    SaveAndUpdateLambdaMetadata OfCell, True, OldFormula
    
    ' Resize the formula bar to fit the formula
    AutofitFormulaBar OfCell
    Application.StatusBar = False
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.SaveLambda"
    Exit Sub
    
ErrorHandler:
    
    Dim errorNumber As Long
    errorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description
    
    Application.StatusBar = False
    
    ' Raising the error to be handled by the calling procedure
    If errorNumber <> 0 Then
        Err.Raise errorNumber, Err.Source, ErrorDescription
        ' This is only for debugging purpose.
        Resume
    End If
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.SaveLambda"
    
End Sub

Private Sub TryToGenerateLambdaIfNotGeneratedAlready(ByVal ForCell As Range)
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.TryToGenerateLambdaIfNotGeneratedAlready"
    Dim OldFormulaText As String
    OldFormulaText = ForCell.Formula
    
    If Not Text.IsStartsWith(OldFormulaText, EQUAL_SIGN) Then Exit Sub
    
    ' If formula in active cell formula doesn't start with =LAMBDA( then Generate lambda.
    If Not IsLambdaFunction(OldFormulaText) Then
        ' In the process of creating lambda FormulaStartCell loop back to other cell and if we set to different var
        ' then it will hold the actual cell for var ForCell.
        Dim FormulaStartCell As Range
        Set FormulaStartCell = ForCell
        modLambdaBuilder.GenerateLambdaStatement FormulaStartCell, ForCell
    End If
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.TryToGenerateLambdaIfNotGeneratedAlready"
    
End Sub

Private Sub SaveAndUpdateLambdaMetadata(ByVal ForCell As Range _
                                        , Optional ByVal AppendMetadata As Boolean = True _
                                         , Optional ByVal OldFormulaIfPopUpCancelled As String = vbNullString)
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.SaveAndUpdateLambdaMetadata"
    ' Retrieving the old formula
    Dim OldFormulaText As String
    OldFormulaText = ForCell.Formula2
    
    ' Checking if the formula is a Lambda
    If Not IsLambdaFunction(OldFormulaText) Or Not Text.IsStartsWith(OldFormulaText, EQUAL_SIGN) Then Exit Sub
    
    ' Retrieving the Lambda name from the cell comment
    Dim LambdaName As String
    LambdaName = GetOldNameFromComment(ForCell, LAMBDA_NAME_NOTE_PREFIX)
    
    ' Checking if the Lambda name is empty
    If LambdaName = vbNullString Then
        ' Saving Lambda if no name was previously defined
        SaveLambdaAsAfterTakingUserInput ForCell, OldFormulaIfPopUpCancelled
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.SaveAndUpdateLambdaMetadata"
        Exit Sub
    End If
    
    ' Saving and updating metadata if the Lambda name was previously defined
    SaveAndUpdateLambdaMetadataConsideringLambdaNameDefined ForCell, AppendMetadata
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.SaveAndUpdateLambdaMetadata"
    
End Sub

Private Sub SaveAndUpdateLambdaMetadataConsideringLambdaNameDefined(ByVal ForCell As Range _
                                                                    , Optional ByVal AppendMetadata As Boolean = True)
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.SaveAndUpdateLambdaMetadataConsideringLambdaNameDefined"
    Dim NewLambdaFormulaText As String
    NewLambdaFormulaText = ForCell.Formula2
    
    ' Getting the old Lambda name from the cell comment
    Dim LambdaName As String
    LambdaName = GetOldNameFromComment(ForCell, LAMBDA_NAME_NOTE_PREFIX)
    
    If IsLambdaCreatedInExcelLabs(ForCell.Worksheet.Parent, LambdaName) Then
        MsgBox "A lambda named '" & LambdaName _
               & "' already exists in Excel Labs' Advanced Formula Editor.  You must edit it there to avoid conflicts." _
               , vbOKOnly + vbInformation, "Save Lambda"
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.SaveAndUpdateLambdaMetadataConsideringLambdaNameDefined"
        Exit Sub
    End If
    
    
    If AppendMetadata Then NewLambdaFormulaText = AddMetadataFromNameManager(NewLambdaFormulaText _
                                                                             , LambdaName, ForCell)
    
    Dim DefPart As String
    DefPart = GetLambdaDefPart(NewLambdaFormulaText)
    Dim InvocationPart As String
    InvocationPart = GetLambdaInvocationPart(NewLambdaFormulaText)

    DefPart = ConvertDependencisToFullyQualifiedRef(RemoveSpaceBetweenEqualAndLambdaText(DefPart) _
                                                    , ForCell.Worksheet)
    
    ' Get the comment from name manager using the lambda name and cell
    Dim CommentInNameManager As String
    CommentInNameManager = GetCommentForNameManager(LambdaName, ForCell, NewLambdaFormulaText)
    
    ' Add new name referring to the new lambda formula text
    On Error Resume Next
    Dim CurrentName As name
    Set CurrentName = ForCell.Worksheet.Parent.Names.Add(name:=LambdaName, RefersTo:=DefPart)
    
    ' We have noticed that sometimes changing comment change the RefersTo of the Lambda (Portuguese/Brazil).
    ' That's why we are resetting it again to the original.
    If CommentInNameManager <> vbNullString Then
        CurrentName.RefersTo = "=$A$1"
        CurrentName.Comment = CommentInNameManager
        CurrentName.RefersTo = DefPart
    End If
    
    ' If an error occurred during adding new name, show a message box and exit the subroutine
    If Err.Number <> 0 Then
        MsgBox "Unable to save lambda to name '" & LambdaName & "'.", vbOKOnly + vbExclamation, "Save Lambda"
        On Error GoTo 0
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.SaveAndUpdateLambdaMetadataConsideringLambdaNameDefined"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Replace the lambda formula with the name of the lambda and remove "Lambda: <lambda name>",
    Dim NewFormulaText As String
    NewFormulaText = EQUAL_SIGN & LambdaName & InvocationPart
    NewFormulaText = FormatFormula(NewFormulaText)
    
    ' Update the formula in the cell and calculate
    UpdateFormulaAndCalculate ForCell, NewFormulaText
    ' Delete comment containing lambda name from the cell
    DeleteComment ForCell
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.SaveAndUpdateLambdaMetadataConsideringLambdaNameDefined"
    
End Sub

Private Function AddMetadataFromNameManager(ByVal NewFormula As String _
                                            , ByVal LambdaName As String _
                                             , ByVal ForCell As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.AddMetadataFromNameManager"
    On Error GoTo NoMetadata
    Dim OldFormula As String
    OldFormula = ForCell.Worksheet.Parent.Names(LambdaName).RefersTo
    If OldFormula = vbNullString Then
        AddMetadataFromNameManager = NewFormula
    Else
        Dim LetParts As Variant
        LetParts = GetDependencyFunctionResult(OldFormula, LET_PARTS, True)
        If Not IsArrayAllocated(LetParts) Then
            AddMetadataFromNameManager = NewFormula
        Else
            Dim FirstColumnIndex  As Long
            FirstColumnIndex = LBound(LetParts, 2)
            Dim CurrentRowIndex As Long
            For CurrentRowIndex = UBound(LetParts, 1) To LBound(LetParts, 1) Step -1
                Dim StepName As String
                StepName = LetParts(CurrentRowIndex, FirstColumnIndex)
                If Text.IsStartsWith(StepName, METADATA_IDENTIFIER) Then
                    Dim StepCalculation As String
                    StepCalculation = LetParts(CurrentRowIndex _
                                               , FirstColumnIndex + LET_PARTS_VALUE_COL_INDEX - 1)
                    NewFormula = InsertLetStep(NewFormula, 1, StepName, StepCalculation)
                End If
            Next CurrentRowIndex
        End If
    End If
    AddMetadataFromNameManager = NewFormula
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.AddMetadataFromNameManager"
    Exit Function
    
NoMetadata:
    AddMetadataFromNameManager = NewFormula
    Err.Clear
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.AddMetadataFromNameManager"
    
End Function

Private Function GetCommentForNameManager(ByVal LambdaName As String _
                                          , ByVal FromCell As Range _
                                           , ByVal NewFormulaText As String) As String
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.GetCommentForNameManager"
    ' Try to get the current name from workbook's names collection
    On Error GoTo NoMetadata
    
    If IsMetadataPresent(NewFormulaText) Then
        GetCommentForNameManager = GetCommentForNameManagerFromFormulaText(NewFormulaText)
    Else
        
        Dim CurrentName As name
        Set CurrentName = FromCell.Worksheet.Parent.Names(LambdaName)
        ' If comment already exists, return it as the result
        If CurrentName.Comment <> vbNullString Then
            GetCommentForNameManager = CurrentName.Comment
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.GetCommentForNameManager"
            Exit Function
        End If
        
        GetCommentForNameManager = GetCommentForNameManagerFromFormulaText(CurrentName.RefersTo)
    End If
    
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.GetCommentForNameManager"
    Exit Function

    ' If no metadata is found, return null string
NoMetadata:
    GetCommentForNameManager = vbNullString
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.GetCommentForNameManager"
    
End Function

Private Function GetCommentForNameManagerFromFormulaText(ByVal FormulaText As String) As String
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.GetCommentForNameManagerFromFormulaText"
    If FormulaText = vbNullString Then
        GetCommentForNameManagerFromFormulaText = vbNullString
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.GetCommentForNameManagerFromFormulaText"
        Exit Function
    End If
    
    Dim LambdaParts As Variant
    LambdaParts = GetDependencyFunctionResult(FormulaText, LAMBDA_PARTS)
    Dim LetParts As Variant
    LetParts = GetDependencyFunctionResult(FormulaText, LET_PARTS)
    Dim CurrentMetadata As Metadata
    Set CurrentMetadata = Metadata.CreateLambdaMetadata(LambdaParts, LetParts, vbNullString, vbNullString)
    GetCommentForNameManagerFromFormulaText = CurrentMetadata.NameManagerComment

    ' Clean up
    Set CurrentMetadata = Nothing
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.GetCommentForNameManagerFromFormulaText"
    
End Function

Private Function RemoveSpaceBetweenEqualAndLambdaText(ByVal LambadaFormula As String) As String
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.RemoveSpaceBetweenEqualAndLambdaText"
    Dim FindText As String
    FindText = EQUAL_SIGN & ONE_SPACE & LAMBDA_FX_NAME & FIRST_PARENTHESIS_OPEN
    
    Dim ReplaceText As String
    ReplaceText = EQUAL_SIGN & LAMBDA_FX_NAME & FIRST_PARENTHESIS_OPEN
    
    RemoveSpaceBetweenEqualAndLambdaText = Replace(LambadaFormula, FindText, ReplaceText, 1, 1, vbTextCompare)
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.RemoveSpaceBetweenEqualAndLambdaText"
    
End Function

'--------------------------------------------< OA HOTKEY >--------------------------------------------
' Command Name:           Save Lambda As
' Description:            Saves the Lambda definition in the active cell as a new defined name specified by user.
' Example:                If user specifies new lambda name is SumAAndB, converts active cell
'                         containing =LAMBDA(_editing_,a,b,SUM(a,b))("Editing: AddAPlusB",3,4)
'                         to =SumAAndB(3,4) and creates/updates the named range formula for SumAAndB
'                         to =LAMBDA(a,b,SUM(a,b))
' Macro Expression:       modLambdaEditor.SaveLambdaAs()
' Generated:              03/06/2022 05:51
'-----------------------------------------------------------------------------------------------------
Public Sub SaveLambdaAs(ByVal OfCell As Range, Optional ByVal OldFormulaIfPopUpCancelled As String = vbNullString)
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.SaveLambdaAs"
    ' Check for command validity, if invalid then exit
    If IsInvalidToRunCommand(OfCell, "Save Lambda As") Then Exit Sub
    
    ' If old formula not provided, take the current cell formula
    If OldFormulaIfPopUpCancelled = vbNullString Then OldFormulaIfPopUpCancelled = OfCell.Formula2

    On Error GoTo ErrorHandler

    ' Initializing the save process and acquiring necessary resources
    Application.StatusBar = "Saving Lambda As... (please wait)"

    DeleteLETStepNamedRangesHavingError OfCell.Worksheet.Parent

    ' Check and generate lambda if not already generated
    TryToGenerateLambdaIfNotGeneratedAlready OfCell
    
    ' Perform save operation considering the available resources
    SaveLambdaAsAfterTakingUserInput OfCell, OldFormulaIfPopUpCancelled

    ' Auto-fit the formula bar for the cell
    AutofitFormulaBar OfCell

    ' Reset the status bar
    Application.StatusBar = False
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.SaveLambdaAs"
    Exit Sub
    
ErrorHandler:
    ' Record error number and description if an error occurs
    Dim errorNumber As Long
    errorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description

    ' Reset status bar, release resources and clean up
    Application.StatusBar = False
    
    ' If error occurred, raise it
    If errorNumber <> 0 Then
        Err.Raise errorNumber, Err.Source, ErrorDescription
        ' This is only for debugging purpose.
        Resume
    End If
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.SaveLambdaAs"
    
End Sub

Public Sub SaveLambdaAsAfterTakingUserInput(ByVal OfCell As Range _
                                            , Optional ByVal OldFormulaIfPopUpCancelled As String = vbNullString)

    Logger.Log TRACE_LOG, "Enter modLambdaEditor.SaveLambdaAsAfterTakingUserInput"
    Dim OldFormulaText As String
    OldFormulaText = OfCell.Formula2
    Dim DefaultName As String
    DefaultName = modUtility.FindLetVarName(OfCell)
    DefaultName = modUtility.MakeValidLetVarName(DefaultName, GetNamingConv(False))

    ' Ensure that the formula starts with "=LAMBDA("
    OldFormulaText = RemoveSpaceBetweenEqualAndLambdaText(OldFormulaText)
    ' If not, log the absence of a lambda and exit the subroutine
    If Not IsLambdaFunction(OldFormulaText) Then
        Logger.Log DEBUG_LOG, OfCell.Address & " has no Lambda"
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.SaveLambdaAsAfterTakingUserInput"
        Exit Sub
    End If

    ' Create a Presenter object to handle metadata editing
    Dim CurrentPresenter As Presenter
    Set CurrentPresenter = EditMetadataForCell(OfCell, True, DefaultName)
    ' If the Presenter was cancelled, revert to the old formula and exit
    If CurrentPresenter.IsProcessCancelled Then
        OfCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormulaIfPopUpCancelled)
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.SaveLambdaAsAfterTakingUserInput"
        Exit Sub
    End If

    ' Validate the new name of the lambda function
    Dim NewName As String
    NewName = CurrentPresenter.LambdaMetadata.LambdaName
    ' If invalid, alert the user and exit the subroutine
    If NewName = vbNullString Then
        MsgBox "Unable to save Lambda formula.  Lambda Name is not valid.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.SaveLambdaAsAfterTakingUserInput"
        Exit Sub
    End If

    ' Update or add a note containing the lambda name
    UpdateOrAddLambdaNameNote OfCell, NewName, LAMBDA_NAME_NOTE_PREFIX
    ' Finally, save and update the lambda and its metadata, considering the lambda name definition
    SaveAndUpdateLambdaMetadataConsideringLambdaNameDefined OfCell, False
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.SaveLambdaAsAfterTakingUserInput"

End Sub

'--------------------------------------------< OA HOTKEY >--------------------------------------------
' Command Name:           Remove Lambda
' Description:            Removes the defined name for the Lambda in active cell and reverts back to Lambda definition.
' Example:                Converts active cell containing =AddAPlusB(3,4) to =LAMBDA(a,b,SUM(a,b))(3,4)
'                         and deletes the named range AddAPlusB
' Macro Expression:       modLambdaEditor.RemoveLambda()
' Generated:              03/06/2022 05:53
'-----------------------------------------------------------------------------------------------------
Public Sub RemoveLambda(ByVal OfCell As Range)
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.RemoveLambda"
    ' Check if the command is valid for the cell
    If IsInvalidToRunCommand(OfCell, "Remove Lambda") Then Exit Sub
    
    ' Get the formula name from the cell's comment
    Dim FormulaName As String
    FormulaName = GetOldNameFromComment(OfCell, LAMBDA_NAME_NOTE_PREFIX)
    
    ' If formula name is not found, edit the lambda
    If FormulaName = vbNullString Then
        EditLambda OfCell
    End If
    
    ' Try to get the formula name again
    FormulaName = GetOldNameFromComment(OfCell, LAMBDA_NAME_NOTE_PREFIX)
    
    ' If still not found, exit the subroutine
    If FormulaName = vbNullString Then
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.RemoveLambda"
        Exit Sub
    End If

    ' Delete the named range matching the lambda name
    ' Error handling is used to ignore errors if the named range does not exist
    On Error Resume Next
    OfCell.Worksheet.Parent.Names(FormulaName).Delete
    On Error GoTo 0
    
    ' Delete the comment that contains the lambda name
    DeleteComment OfCell
    AutofitFormulaBar OfCell
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.RemoveLambda"
    
End Sub

'--------------------------------------------< OA HOTKEY >--------------------------------------------
' Command Name:           Cancel Edit Lambda
' Description:            Cancel any edits to Lambda definition in active cell and revert back to custom Lambda call.
' Macro Expression:       modLambdaEditor.CancelEditLambda()
' Generated:              03/06/2022 06:22
'-----------------------------------------------------------------------------------------------------
Public Sub CancelEditLambda(ByVal OfCell As Range)
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.CancelEditLambda"
    ' Check if the command is valid for the cell
    If IsInvalidToRunCommand(OfCell, "Cancel Edit Lambda") Then Exit Sub

    ' Creating a Resource Handler to manage Excel resources and adding resource
    On Error GoTo ErrorHandler

    ' Get the name of the Lambda from the cell's comment
    Dim LambdaName As String
    LambdaName = GetOldNameFromComment(OfCell, LAMBDA_NAME_NOTE_PREFIX)
    
    ' If Lambda name is not found, proceed to resource cleanup
    If LambdaName = vbNullString Then
        Resume ErrorHandler
    End If
    
    ' Delete the comment in the cell containing the Lambda name
    DeleteComment OfCell
    
    ' Replace the Lambda formula with the Lambda name and get the invocation part
    Dim NewFormula As String
    NewFormula = EQUAL_SIGN & LambdaName & GetLambdaInvocationPart(OfCell.Formula2)
    NewFormula = FormatFormula(NewFormula)
    
    ' Update the formula in the active cell and recalculate it
    UpdateFormulaAndCalculate OfCell, NewFormula
    ' Resize formula bar to fit new formula
    AutofitFormulaBar OfCell
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.CancelEditLambda"
    Exit Sub
    
ErrorHandler:

    ' Collect error info if any
    Dim errorNumber As Long
    errorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description
    
    ' Debug print the new formula for troubleshooting
    Debug.Print "New Formula : " & NewFormula

    ' Re-raise error for debugging
    If errorNumber <> 0 Then
        Err.Raise errorNumber, Err.Source, ErrorDescription
        'This is only for debugging purpose.
        Resume
    End If
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.CancelEditLambda"

End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Autofit Formula Bar
' Description:            Autofit formula bar height based on formula length so that whole formula is visible.
' Macro Expression:       modLambdaEditor.AutofitFormulaBar([ActiveCell])
' Generated:              07/16/2023 02:17 PM
'----------------------------------------------------------------------------------------------------
Public Sub AutofitFormulaBar(ByVal FormulaCell As Range)
    
    Logger.Log TRACE_LOG, "Enter modLambdaEditor.AutofitFormulaBar"
    ' Constants for the minimum and maximum height of the formula bar
    Const MIN_HEIGHT As Long = 4
    Const MAX_HEIGHT As Long = 10
    
    ' Calculate the number of new lines in the cell's formula
    Dim NewLineCount As Long
    NewLineCount = Len(FormulaCell.Formula) - Len(VBA.Replace(FormulaCell.Formula, Chr$(10), vbNullString)) + 1

    On Error GoTo TryOnceAgain
    ' Adjust the height of the formula bar based on the number of new lines
    ' If the number of lines is less than the minimum height, set it to the minimum height
    If NewLineCount < MIN_HEIGHT Then
        Application.FormulaBarHeight = MIN_HEIGHT
        ' If the number of lines is more than the maximum height, set it to the maximum height
    ElseIf NewLineCount > MAX_HEIGHT Then
        Application.FormulaBarHeight = MAX_HEIGHT
        ' If the number of lines is between the minimum and maximum heights, set the height equal to the number of lines
    Else
        Application.FormulaBarHeight = NewLineCount
    End If
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaEditor.AutofitFormulaBar"
    Exit Sub
    
TryOnceAgain:
    
    ' After openning excel and before activating VBE if we try to run this then it doesn't work for the first time.
    ' After using Resume it may not work too. But trying.
    ' But after that every time we run this command then it will work.
    Dim ErrorCount As Long
    ErrorCount = ErrorCount + 1
    If ErrorCount = 1 Then Resume
    Err.Clear
    Logger.Log TRACE_LOG, "Exit modLambdaEditor.AutofitFormulaBar"
    
End Sub


