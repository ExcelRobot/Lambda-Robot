Attribute VB_Name = "modLambdaBuilder"
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed, FunctionReturnValueDiscarded, IndexedDefaultMemberAccess, EmptyMethod, UnrecognizedAnnotation, ProcedureNotUsed
'@Folder "Lambda.Editor.Driver"

Option Explicit

Public Sub TestGenerateDependencyInfo()
    
    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.TestGenerateDependencyInfo"
    Dim PutOn As Range
    Set PutOn = ActiveCell
    modLambdaBuilder.GenerateLambdaStatement ActiveCell, PutOn
    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.TestGenerateDependencyInfo"
    
End Sub

Public Sub GenerateDependencyInfo(ByVal FormulaRange As Range, ByVal PutOnRange As Range _
                                                              , Optional ByVal DependencySearchInRegion As Range _
                                                               , Optional ByVal IsUndo As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.GenerateDependencyInfo"
    ' Keep the data in a static ListObject
    Static Table As ListObject
    ' Keep track of the formula range to use for undo
    Static PutFormulaOnUndo As Range
    
    If IsUndo Then
        ' If table exists, delete it
        If IsNotNothing(Table) Then Table.Delete
        ' If undo, select the previously stored formula range
        If IsNotNothing(Table) Then PutFormulaOnUndo.Select
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.GenerateDependencyInfo"
        Exit Sub
    Else
        ' If not undo, store the formula range for future use
        Set PutFormulaOnUndo = FormulaRange
    End If

    ' Generate the information about the lambda formula
    Dim CurrentLambdaInfo As LambdaInfo
    Set CurrentLambdaInfo = GetLambdaInfo(FormulaRange, PutOnRange _
                                           , OperationType.DEPENDENCY_INFO_GENERATION _
                                            , DependencySearchInRegion)
    
    ' If the operation is not undo, then set the table to hold the dependency information and assign the current procedure to undo stack
    If Not IsUndo Then
        Set Table = CurrentLambdaInfo.PutDependencyOnTable
        AssingOnUndo "GenerateDependencyInfo"
    End If
    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.GenerateDependencyInfo"

End Sub

Private Sub GenerateDependencyInfo_Undo()
    GenerateDependencyInfo Nothing, Nothing, Nothing, True
End Sub

Public Sub GenerateLetStatement(ByVal FormulaRange As Range, ByVal PutLetOnCell As Range _
                                                            , Optional ByVal DependencySearchInRegion As Range _
                                                             , Optional ByVal IsUndo As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.GenerateLetStatement"
    
    ' Keep track of the formula range and old formula for undo operation
    Static PutFormulaOnUndo As Range
    Static OldFormula As String

    If IsUndo Then
        ' If undo, restore the old formula
        If IsNotNothing(PutFormulaOnUndo) Then PutFormulaOnUndo.Formula2 = OldFormula
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.GenerateLetStatement"
        Exit Sub
    Else
        OldFormula = FormulaRange.Formula2
    End If

    ' Generate the information about the lambda formula
    Dim CurrentLambdaInfo As LambdaInfo
    Set CurrentLambdaInfo = GetLambdaInfo(FormulaRange, Nothing, OperationType.LET_STATEMENT_GENERATION _
                                                                , DependencySearchInRegion)
    
    If IsNotNothing(CurrentLambdaInfo) Then
        ' Assign the generated let statement to the cell and print it in the debug window if there's an error
        AssignFormulaIfErrorPrintIntoDebugWindow PutLetOnCell, CurrentLambdaInfo.LetFormula _
                                                              , "Generated Let Statement : "
        ' Force calculation after assigning formula
        PutLetOnCell.Calculate
        
        If Not IsUndo Then
            ' Check if FormulaRange and PutLetOnCell are the same, if so, store FormulaRange for future use, otherwise store PutLetOnCell
            If GetRangeRefWithSheetName(FormulaRange) = GetRangeRefWithSheetName(PutLetOnCell) Then
                Set PutFormulaOnUndo = FormulaRange
            Else
                Set PutFormulaOnUndo = PutLetOnCell
            End If
            
            ' Assign the current procedure to undo stack
            AssingOnUndo "GenerateLetStatement"
        End If
    End If

    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.GenerateLetStatement"
    
End Sub

Private Sub GenerateLetStatement_Undo()
    GenerateLetStatement Nothing, Nothing, Nothing, True
End Sub

Public Sub GenerateLambdaStatement(ByVal FormulaRange As Range _
                                   , ByVal PutLambdaOnCell As Range _
                                    , Optional ByVal DependencySearchInRegion As Range _
                                     , Optional ByVal IsCreateInNameManager As Boolean = False _
                                      , Optional ByVal IsExportable As Boolean = False _
                                       , Optional ByVal IsUndo As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.GenerateLambdaStatement"

    ' Keeping track of previous state for undo functionality
    Static PutFormulaOnUndo As Range
    Static OldFormula As String

    If IsUndo Then
        If IsNotNothing(PutFormulaOnUndo) Then PutFormulaOnUndo.Formula2 = OldFormula
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.GenerateLambdaStatement"
        Exit Sub
    Else
        OldFormula = FormulaRange.Cells(1).Formula2
    End If

    ' Exit if there is no formula
    If FormulaRange.Cells(1).Formula2 = vbNullString Then Exit Sub
    
    Dim FormulaRangeFullAddress As String
    FormulaRangeFullAddress = GetRangeRefWithSheetName(PutLambdaOnCell)

    Dim CurrentLambdaInfo As LambdaInfo
    Set CurrentLambdaInfo = GetLambdaInfo(FormulaRange, Nothing _
                                                       , LAMBDA_STATEMENT_GENERATION _
                                                        , DependencySearchInRegion _
                                                         , IsCreateInNameManager)

    ' If creating in the name manager, update the PutLambdaOnCell reference
    If IsCreateInNameManager Then
        Set PutLambdaOnCell = RangeResolver.GetRange(FormulaRangeFullAddress, PutLambdaOnCell.Worksheet.Parent)
    End If

    ' If a lambda info object was created, process it
    If IsNotNothing(CurrentLambdaInfo) Then
        Dim FormulaText As String
        FormulaText = FormatFormula(ReplaceNewlineWithChar10(CurrentLambdaInfo.LambdaFormula & CurrentLambdaInfo.InvocationArgument))
        AssignFormulaIfErrorPrintIntoDebugWindow PutLambdaOnCell, FormulaText, "Formula : "

        ' Include any lambda dependencies
        IncludeLambdaDependencies PutLambdaOnCell, IsUndo, True
        PutLambdaOnCell.Calculate
        PutLambdaOnCell.Activate
        
        ' Handle undo state
        If Not IsUndo Then
            If GetRangeRefWithSheetName(FormulaRange) = GetRangeRefWithSheetName(PutLambdaOnCell) Then
                Set PutFormulaOnUndo = FormulaRange
            Else
                Set PutFormulaOnUndo = PutLambdaOnCell
            End If
            AssingOnUndo "GenerateLambdaStatement"
        End If
    End If
    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.GenerateLambdaStatement"
    
End Sub

Private Sub GenerateLambdaStatement_Undo()
    GenerateLambdaStatement Nothing, Nothing, Nothing, , , True
End Sub

Public Sub GenerateAFEScript(ByVal FormulaRange As Range, ByVal PutAFEScriptOnCell As Range _
                                                         , Optional ByVal DependencySearchInRegion As Range)
    
    GetLambdaInfo FormulaRange, PutAFEScriptOnCell, AFE_SCRIPT_GENERATION, DependencySearchInRegion
    
End Sub

Private Function GetLambdaInfo(ByVal FormulaRange As Range _
                               , ByVal PutOnRange As Range _
                                , ByVal TypeOfOperation As OperationType _
                                 , ByVal DependencySearchInRegion As Range _
                                  , Optional ByVal IsCreateInNameManager As Boolean = False _
                                   , Optional ByVal IsExportable As Boolean = False) As LambdaInfo

    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.GetLambdaInfo"

    ' Checks if the workbook has been saved. If not, sends a message and exits function
    If IsWorksheetProtected(FormulaRange.Worksheet) Then
        MsgBox "Unable to generate " & GetTypeOfOperationText(TypeOfOperation) & _
               " on a protected worksheet. Unprotect the worksheet and try again." _
               , vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.GetLambdaInfo"
        Exit Function
        ' Checks if the workbook is protected. If it is, sends a message and exits function
    ElseIf IsWorkbookProtected(FormulaRange.Worksheet.Parent) Then
        MsgBox "Unable to generate " & GetTypeOfOperationText(TypeOfOperation) & _
               " on a protected workbook. Unprotect the workbook and try again." _
               , vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.GetLambdaInfo"
        Exit Function
    End If

    On Error GoTo ErrorHandler
    
    ' Creating a new instance of the FormulaParser
    Dim CurrentParser As FormulaParser
    Set CurrentParser = New FormulaParser

    ' Setting properties of the FormulaParser based on provided parameters
    CurrentParser.IsAddToNameManager = IsCreateInNameManager
    CurrentParser.IsExportable = IsExportable

    ' Calling the method to generate details of the lambda function based on the provided parameters
    CurrentParser.CreateLambdaDetails FormulaRange, PutOnRange, TypeOfOperation, DependencySearchInRegion
    
    ' If the process was terminated by the user, set the return value to Nothing
    If CurrentParser.IsProcessTerminatedByUser Then
        Set GetLambdaInfo = Nothing
    Else
        ' Otherwise, retrieve the LambdaInfo from the parser
        Set GetLambdaInfo = CurrentParser.GetLambdaInfo
    End If
    
    ' Release the parser object
    Set CurrentParser = Nothing
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.GetLambdaInfo"
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Resume
    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.GetLambdaInfo"
    
End Function

Private Function GetTypeOfOperationText(ByVal TypeOfOperation As OperationType) As String
    
    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.GetTypeOfOperationText"
    Select Case TypeOfOperation

        Case OperationType.DEPENDENCY_INFO_GENERATION
            GetTypeOfOperationText = "Dependency info"
        Case OperationType.LET_STATEMENT_GENERATION
            GetTypeOfOperationText = "LET statement"
        Case OperationType.LAMBDA_STATEMENT_GENERATION
            GetTypeOfOperationText = "LAMBDA statement"
        Case OperationType.AFE_SCRIPT_GENERATION
            GetTypeOfOperationText = "AFE Script"
        Case Else
            Err.Raise 13, "Wrong Input Argument"
    End Select
    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.GetTypeOfOperationText"

End Function

Private Sub TestConvertLetToLambda()
    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.TestConvertLetToLambda"
    ' Logger.Log DEBUG_LOG, ConvertLetToLambda(Sheet7.Range("F13").Formula2)
    LetToLambda ActiveCell, ActiveCell.Offset(1, 0)
    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.TestConvertLetToLambda"
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Let To Lambda
' Description:            Let to lambda.
' Macro Expression:       modLambdaBuilder.LetToLambda([ActiveCell],[ActiveCell.Offset(0,1)])
' Generated:              06/14/2022 10:54 AM
'----------------------------------------------------------------------------------------------------
Public Sub LetToLambda(ByVal LetFormulaCell As Range, ByVal PutLambdaOnCell As Range _
                                                     , Optional ByVal IsUndo As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.LetToLambda"
    
    ' If PutLambdaOnCell is not defined, set it to LetFormulaCell
    If IsNothing(PutLambdaOnCell) Then Set PutLambdaOnCell = LetFormulaCell

    ' Declare variables for undo operation
    Static PutFormulaOnUndo As Range
    Static OldFormula As String
    
    ' If IsUndo is true, revert to old formula and exit subroutine
    If IsUndo Then
        If IsNotNothing(PutFormulaOnUndo) Then PutFormulaOnUndo.Formula2 = OldFormula
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.LetToLambda"
        Exit Sub
    Else
        Dim InvalidReason As String
        InvalidReason = LetToLambdaInvalidMessage(LetFormulaCell, PutLambdaOnCell)
        If InvalidReason = vbNullString Then
            ' If not undo operation, save current formula for potential future undo
            OldFormula = PutLambdaOnCell.Formula2
        Else
            MsgBox InvalidReason, vbExclamation + vbOKOnly, "LET To LAMBDA"
            Exit Sub
        End If
        
    End If
    
    On Error GoTo ErrorHandler
    ' Convert LetFormula to Lambda and assign it to target cell
    Dim FormulaText As String
    FormulaText = LETToLAMBDAConverter.ConvertLetToLambda(LetFormulaCell.Formula2)
    AssignFormulaIfErrorPrintIntoDebugWindow PutLambdaOnCell, FormulaText, "Formula : "
    
    If Not IsUndo Then
        ' Saving the cell that may need to be reverted to in the future
        If GetRangeRefWithSheetName(LetFormulaCell) = GetRangeRefWithSheetName(PutLambdaOnCell) Then
            Set PutFormulaOnUndo = LetFormulaCell
        Else
            Set PutFormulaOnUndo = PutLambdaOnCell
        End If
        ' Saving the current action for potential future undo
        AssingOnUndo "LetToLambda"
    End If
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.LetToLambda"
    Exit Sub
    
ErrorHandler:
    Dim errorNumber As Long
    errorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description
    ' Raise the error again for further processing
    If errorNumber <> 0 Then
        Err.Raise errorNumber, Err.Source, ErrorDescription
        ' This is only for debugging purpose.
        Resume
    End If
    
    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.LetToLambda"
    
End Sub

Private Sub LetToLambda_Undo()
    LetToLambda Nothing, Nothing, True
End Sub

Private Function LetToLambdaInvalidMessage(ByVal LetFormulaCell As Range _
                                           , ByVal PutLambdaOnCell As Range) As String

    Dim Reason As String
    ' Check if more than one cell is selected in LetFormulaCell
    ' If more than one cell is selected, display a message box and exit the function
    If LetFormulaCell.Cells.Count > 1 Then
        Reason = "Unable to convert " & LET_FX_NAME & " to " & LAMBDA_FX_NAME _
                 & ". Only one cell at a time allowed."
    ElseIf Not LetFormulaCell.HasFormula Then
        Reason = "No formula found on " & LetFormulaCell.Address
    ElseIf Not IsLetFunction(LetFormulaCell.Formula2) Then
        Reason = "The formula is not a LET formula.  Procedure aborted."
    ElseIf PutLambdaOnCell.Address(, , , True) <> LetFormulaCell.Address(, , , True) Then
        ' Check if the address of PutLambdaOnCell is different from the LetFormulaCell
        ' Also check if the cell is not empty or already contains a formula
        ' If it's not empty or contains a formula, display a message box and exit the function
        If IsError(PutLambdaOnCell) Then
            Reason = "Unable to convert " & LET_FX_NAME & " to " & LAMBDA_FX_NAME _
                   & ".  Destination range is not empty."
        ElseIf PutLambdaOnCell.Value <> vbNullString Or PutLambdaOnCell.HasFormula Then
            Reason = "Unable to convert " & LET_FX_NAME & " to " & LAMBDA_FX_NAME _
                   & ".  Destination range is not empty."
        End If
    End If

    ' If all checks pass, the arguments are valid for the LET to LAMBDA conversion
    LetToLambdaInvalidMessage = Reason

End Function

'Public Function ConvertLetToLambda(ByVal LetFormula As String) As String
'
'    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.ConvertLetToLambda"
'    ' Check if the formula in the cell starts with "LAMBDA", if so, no conversion is needed
'    If IsLambdaFunction(LetFormula) Then
'        ConvertLetToLambda = LetFormula
'        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.ConvertLetToLambda"
'        Exit Function
'    End If
'
'    Dim LetParts As Variant
'    LetParts = GetDependencyFunctionResult(LetFormula, LET_PARTS)
'    ' If the returned result is not an array, that means no LET function was found in the cell, hence a message box pops up to notify the user
'    If Not IsArray(LetParts) Then
'        MsgBox "Unable to convert " & LET_FX_NAME & " to " & LAMBDA_FX_NAME _
'               & ".  No " & LET_FX_NAME & " function found.", vbExclamation + vbOKOnly, APP_NAME
'        ' Since no LET function was found, the function logs its exit and terminates
'        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.ConvertLetToLambda"
'        Exit Function
'    End If
'
'
'    ' Retrieve the index of the first column of the "LetParts" array
'    Dim FirstColIndex As Long
'    FirstColIndex = LBound(LetParts, 2)
'
'    ' Initialize the lambda argument part of the final string to be built
'    Dim LambdaArgumentPart As String
'    LambdaArgumentPart = EQUAL_SIGN & LAMBDA_FX_NAME & FIRST_PARENTHESIS_OPEN
'
'    ' Initialize the lambda invocation part of the final string to be built
'    Dim LambdaInvocationPart As String
'    LambdaInvocationPart = FIRST_PARENTHESIS_OPEN
'
'    ' Initialize the LET statement of the final lambda to be built
'    Dim LetOfFinalLambda As String
'    LetOfFinalLambda = LET_AND_OPEN_PAREN
'
'    ' Create a WorksheetFunction object to be able to use Excel worksheet functions in VBA
'    Dim AppFunction As WorksheetFunction
'    Set AppFunction = Application.WorksheetFunction
'
'    Dim CountOfNonInputStep As Long
'
'    ' Extract the address of the first LET variable and check if it is an input cell
'    Dim VarName As String
'    Dim CellAddress As String
'    Dim CurrentRowIndex As Long
'    Dim IsFirstLetVarIsInputCell As Boolean
'    CellAddress = LetParts(FirstColIndex, FirstColIndex + modSharedConstant.LET_PARTS_VALUE_COL_INDEX - 1)
'    IsFirstLetVarIsInputCell = IsRangeAddress(CellAddress)
'
'    ' If the first LET variable is a reference to a cell, determine if it's an input cell
'    If IsFirstLetVarIsInputCell Then IsFirstLetVarIsInputCell = IsInputCell(RangeResolver.GetRange(CellAddress), Nothing)
'
'
'    ' Iterate through the "LetParts" array
'    For CurrentRowIndex = LBound(LetParts, 1) To UBound(LetParts, 1)
'        ' Clean and trim the current variable name
'        VarName = VBA.LTrim$(AppFunction.Clean(AppFunction.Trim(LetParts(CurrentRowIndex, FirstColIndex))))
'
'        ' Trim the current cell address
'        CellAddress = Trim$(LetParts(CurrentRowIndex _
'                                     , FirstColIndex + modSharedConstant.LET_PARTS_VALUE_COL_INDEX - 1))
'
'        ' Check if the cell address is a range address
'        If modUtility.IsRangeAddress(CellAddress) Then
'
'            If IsFirstLetVarIsInputCell Then
'                ' If the current cell is an input cell, add it to the lambda argument and invocation part
'                If IsInputCell(RangeResolver.GetRange(CellAddress), Nothing) Then
'                    LambdaArgumentPart = LambdaArgumentPart & UpdateForOptionalArgument(LetFormula, VarName)
'                    LambdaInvocationPart = LambdaInvocationPart & CellAddress & LIST_SEPARATOR
'                    ' If it's not an input cell, increment the non-input step counter and add the LET variable to the final lambda
'                Else
'                    CountOfNonInputStep = CountOfNonInputStep + 1
'                    LetOfFinalLambda = LetOfFinalLambda & NEW_LINE & THREE_SPACE & VarName & _
'                                       LIST_SEPARATOR & _
'                                       Text.PadIfNotPresent(CellAddress, ONE_SPACE, FROM_START) & LIST_SEPARATOR
'                End If
'                ' If the first LET variable was not an input cell, add it to the lambda argument and invocation part
'            Else
'                LambdaArgumentPart = LambdaArgumentPart & UpdateForOptionalArgument(LetFormula, VarName)
'                LambdaInvocationPart = LambdaInvocationPart & CellAddress & LIST_SEPARATOR
'            End If
'            ' If the variable is an optional argument, create it and add it to the lambda argument and invocation part
'        ElseIf modUtility.IsOptionalArgument(LetFormula, VarName) Then
'            LambdaArgumentPart = LambdaArgumentPart & CreateOptionalArgument(VarName)
'            LambdaInvocationPart = LambdaInvocationPart & CellAddress & LIST_SEPARATOR
'            ' If it's not an optional argument or a range address, increment the non-input step counter and add the LET variable to the final lambda
'        Else
'            CountOfNonInputStep = CountOfNonInputStep + 1
'            LetOfFinalLambda = LetOfFinalLambda & NEW_LINE & THREE_SPACE & VarName & _
'                               LIST_SEPARATOR & _
'                               Text.PadIfNotPresent(CellAddress, ONE_SPACE, FROM_START) & LIST_SEPARATOR
'        End If
'    Next CurrentRowIndex
'
'    ' Remove trailing comma from the LambdaInvocationPart
'    LambdaInvocationPart = modUtility.RemoveEndingText(LambdaInvocationPart, LIST_SEPARATOR)
'
'    ' Remove trailing comma from the LetOfFinalLambda
'    LetOfFinalLambda = VBA.RTrim$(Text.RemoveFromEndIfPresent(LetOfFinalLambda, LIST_SEPARATOR))
'
'    ' If only one non-input step was encountered, remove the opening parenthesis from LetOfFinalLambda
'    If CountOfNonInputStep = 1 Then
'        LetOfFinalLambda = VBA.Mid$(LetOfFinalLambda, Len(LET_AND_OPEN_PAREN) + 1)
'        ' Otherwise, append closing parenthesis to LetOfFinalLambda
'    Else
'        LetOfFinalLambda = LetOfFinalLambda & NEW_LINE & THREE_SPACE & FIRST_PARENTHESIS_CLOSE
'    End If
'
'    LambdaInvocationPart = LambdaInvocationPart & FIRST_PARENTHESIS_CLOSE
'
'    ' Form the final lambda function by concatenating LambdaArgumentPart, LetOfFinalLambda and LambdaInvocationPart
'    ConvertLetToLambda = LambdaArgumentPart & LetOfFinalLambda & Chr$(10) _
'                         & FIRST_PARENTHESIS_CLOSE & LambdaInvocationPart
'
'    ' Replace new line character with line feed in the final lambda function so that we can use in cell.
'    ConvertLetToLambda = ReplaceNewlineWithChar10(ConvertLetToLambda)
'
'    Logger.Log DEBUG_LOG, "Result Lambda : " & ConvertLetToLambda
'    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.ConvertLetToLambda"
'
'    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.ConvertLetToLambda"
'
'End Function

'Private Function UpdateForOptionalArgument(ByVal LetFormula As String, ByVal VarName As String) As String
'
'    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.UpdateForOptionalArgument"
'    If modUtility.IsOptionalArgument(LetFormula, VarName) Then
'        ' If it is an optional argument, create an optional argument
'        UpdateForOptionalArgument = CreateOptionalArgument(VarName)
'    Else
'        ' If not, return a cleaned and trimmed VarName, followed by a comma and a space
'        UpdateForOptionalArgument = Application.WorksheetFunction.Trim(Application.WorksheetFunction.Clean(VarName)) _
'                                    & LIST_SEPARATOR & ONE_SPACE
'    End If
'
'    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.UpdateForOptionalArgument"
'
'End Function

'Private Function CreateOptionalArgument(ByVal VarName As String) As String
'
'    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.CreateOptionalArgument"
'
'    ' Establish a reference to Application.WorksheetFunction
'    Dim AppFunction As WorksheetFunction
'    Set AppFunction = Application.WorksheetFunction
'
'    ' Create an optional argument by adding parentheses around a cleaned and trimmed VarName,
'    ' followed by a comma and a space
'    CreateOptionalArgument = LEFT_BRACKET & _
'                             AppFunction.Trim(AppFunction.Clean(VarName)) & _
'                             RIGHT_BRACKET & LIST_SEPARATOR & ONE_SPACE
'    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.CreateOptionalArgument"
'
'End Function

Private Function ConcatenateArray(ByVal GivenArray As Variant, ByVal StartFromIndex As Long _
                                                             , ByVal Delimiter As String) As String
    
    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.ConcatenateArray"

    ' Initialising indices and the formula string
    Dim CurrentIndex As Long
    Dim LetFormula As String
    LetFormula = LET_AND_OPEN_PAREN

    ' Checking if start index is beyond the array bounds
    If StartFromIndex > UBound(GivenArray) Then
        ' Logging premature exit due to exceeding bounds of GivenArray
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.ConcatenateArray"
        Exit Function
    Else
        ' Appending first element from the GivenArray (from StartFromIndex) to the LetFormula string
        LetFormula = LetFormula & VBA.LTrim$(GivenArray(StartFromIndex)) & Delimiter
    End If

    ' Iterating over GivenArray from the element next to StartFromIndex to the end of GivenArray
    For CurrentIndex = StartFromIndex + 1 To UBound(GivenArray)
        ' Appending elements from GivenArray to the LetFormula string
        LetFormula = LetFormula & GivenArray(CurrentIndex) & Delimiter
    Next CurrentIndex

    ' Removing ending delimiter from LetFormula string
    LetFormula = modUtility.RemoveEndingText(LetFormula, Delimiter)

    ' Appending closing parenthesis to LetFormula string
    LetFormula = LetFormula & THREE_SPACE & FIRST_PARENTHESIS_CLOSE

    ' Returning the constructed LetFormula string as the result of the function
    ConcatenateArray = LetFormula

    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.ConcatenateArray"

End Function

' --------------------------------------------< OA Robot >--------------------------------------------
'  Command Name:           Mark As Input Cells
'  Description:            Mark as input cells.
'  Macro Expression:       modLambdaBuilder.MarkAsInputCells([Selection])
'  Generated:              06/13/2022 11:08 AM
' ----------------------------------------------------------------------------------------------------
Public Sub MarkAsInputCells(ByVal GivenRange As Range, Optional ByVal InteriorOnly As Boolean = True)
    
    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.MarkAsInputCells"
    GivenRange.Interior.Color = INPUT_CELL_BACKGROUND_COLOR
    If Not InteriorOnly Then
        GivenRange.Font.Color = INPUT_CELL_FONT_COLOR
    End If
    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.MarkAsInputCells"
    
End Sub

' --------------------------------------------< OA Robot >--------------------------------------------
'  Command Name:           Mark Lambda As LET Step
'  Description:            Create LETStep_FX and LETStepRef_FX named range for activecell formula so that we can use them for further calculation for generating lambda statement.
'  Macro Expression:       modLambdaBuilder.MarkLambdaAsLETStep([ActiveCell])
'  Generated:              07/22/2023 12:41 PM
' ----------------------------------------------------------------------------------------------------
Public Sub MarkLambdaAsLETStep(ByVal ForCell As Range)
    
    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.MarkLambdaAsLETStep"
    Dim FormulaText As String
    FormulaText = ForCell.Formula2
    Dim FxName As String
    Dim RangeReference As String
    RangeReference = GetRangeRefWithSheetName(ForCell, True)
    
    ' Check if the formula does not start with LAMBDA
    If Not IsLambdaFunction(FormulaText) Then
        ' If not, inform the user and exit the subroutine
        MsgBox ForCell.Address & " doesn't contains any " & LAMBDA_FX_NAME & " def." _
               , vbCritical + vbInformation, "Mark Lambda As LET Step"
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.MarkLambdaAsLETStep"
        Exit Sub
    End If
    
    ' Retrieve the function name associated with the range
    FxName = FindRangeLabel(RangeReference, ForCell, True)
    
    ' Check if a function name is found
    If FxName = vbNullString Then
        ' If not, inform the user and request a label
        MsgBox "Couldn't find proper label for the LAMBDA. Please add a label and run again." _
               , vbCritical + vbInformation, "Mark Lambda As LET Step"
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.MarkLambdaAsLETStep"
        Exit Sub
    End If
    
    ' Remove any leading underscore from the function name
    FxName = Text.RemoveFromStartIfPresent(FxName, UNDER_SCORE)
    FxName = MakeValidDefinedName(FxName, Len(FxName) <> 1, False)
    
    Dim AddToBook As Workbook
    Set AddToBook = ForCell.Worksheet.Parent
    
    ' Clean up any errors in LETStep named ranges
    DeleteLETStepNamedRangesHavingError AddToBook
    FormulaText = ConvertDependencisToFullyQualifiedRef(ForCell.Formula2, ForCell.Worksheet)
    FormulaText = GetLambdaDefPart(FormulaText)
    ' Working with named ranges in the workbook
    With AddToBook
        ' Check if a named range exists for the LETStep function
        If IsNamedRangeExist(AddToBook, LETSTEP_UNDERSCORE_PREFIX & FxName) Then
            ' If it exists, update the reference
            .Names(LETSTEP_UNDERSCORE_PREFIX & FxName).RefersTo = FormulaText
        Else
            ' If it doesn't exist, create a new named range
            .Names.Add LETSTEP_UNDERSCORE_PREFIX & FxName, FormulaText
        End If
        
        ' Repeat the process for the LETStepRef named range
        If IsNamedRangeExist(AddToBook, LETSTEPREF_UNDERSCORE_PREFIX & FxName) Then
            .Names(LETSTEPREF_UNDERSCORE_PREFIX & FxName).RefersTo = EQUAL_SIGN & RangeReference
        Else
            .Names.Add LETSTEPREF_UNDERSCORE_PREFIX & FxName, EQUAL_SIGN & RangeReference
        End If
        
    End With
    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.MarkLambdaAsLETStep"
    
End Sub
