Attribute VB_Name = "modAuditLambdaSteps"
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed
'@Folder "Lambda.Debugger.Driver"
'@Ignore ProcedureNotUsed

Option Explicit
Option Private Module

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Generate Lambda Steps
' Description:            Generate lambda steps.
' Macro Expression:       modAuditLambdaSteps.GenerateLambdaSteps([ActiveCell],[ActiveCell.Offset(0,1)])
' Generated:              06/16/2022 11:55 AM
'----------------------------------------------------------------------------------------------------
Public Sub GenerateLambdaSteps(ByVal LambdaInvocationCell As Range, Optional ByVal StepStartCell As Range)
    
    Logger.Log TRACE_LOG, "Enter modAuditLambdaSteps.GenerateLambdaSteps"
    On Error GoTo ErrorHandler
    Const METHOD_NAME As String = "GenerateLambdaSteps"
    Context.ExtractContextFromCell LambdaInvocationCell, METHOD_NAME
    
    If StepStartCell Is Nothing Then
        
        With LambdaInvocationCell.Worksheet
            Dim LastUsedCell As Range
            Set LastUsedCell = .UsedRange.Cells(.UsedRange.Rows.CountLarge, .UsedRange.Columns.CountLarge)
            Set StepStartCell = .Range(.Cells(1, 1), LastUsedCell).Cells(LambdaInvocationCell.Row, LastUsedCell.Column).Offset(0, 2)
            
            ' In case of formula in ROW 1 add the Label in Row 1 and the calculation in Row 2
            If StepStartCell.Row = 1 Then Set StepStartCell = StepStartCell.Offset(1)
        End With
        
    End If
    
    ' Initialize a builder for constructing steps
    Dim Builder As BuildLambdaSteps
    Set Builder = BuildLambdaSteps.Create(LambdaInvocationCell, StepStartCell)
    Builder.ConstructSteps
    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modAuditLambdaSteps.GenerateLambdaSteps"
    Context.ClearContext METHOD_NAME
    Exit Sub
    
ErrorHandler:
    Context.ClearContext METHOD_NAME
    Dim ErrorNumber As Long
    ErrorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description
    
    If Err.Number <> 0 Then
        Err.Raise ErrorNumber, Err.Source, ErrorDescription
        Resume
    End If
    Logger.Log TRACE_LOG, "Exit modAuditLambdaSteps.GenerateLambdaSteps"
    
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Lambda To Let
' Description:            Lambda to let.
' Macro Expression:       modAuditLambdaSteps.LambdaToLet([ACTIVECELL],[ACTIVECELL.Offset(0,1)])
' Generated:              06/15/2022 01:22 PM
'----------------------------------------------------------------------------------------------------
'@Ignore ProcedureNotUsed
Public Sub LambdaToLet(ByVal LambdaFormulaCell As Range _
                       , Optional ByVal PutLetOnCell As Range = Nothing _
                        , Optional ByVal IsUndo As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter modAuditLambdaSteps.LambdaToLet"
    
    Const METHOD_NAME As String = "LambdaToLet"
    Context.ExtractContextFromCell LambdaFormulaCell, METHOD_NAME
    
    ' If no target cell specified, set it to the same as the source cell.
    If IsNothing(PutLetOnCell) Then Set PutLetOnCell = LambdaFormulaCell

    ' Define variables to store the original formula and its cell location for potential undo action.
    Static PutFormulaOnUndo As Range
    Static OldFormula As String
    
    ' Check if it's an undo operation.
    If IsUndo Then
        ' If there's a stored formula from a previous operation, put it back to its original cell.
        If IsNotNothing(PutFormulaOnUndo) Then PutFormulaOnUndo.Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modAuditLambdaSteps.LambdaToLet"
        GoTo ExitMethod
    Else
        
        ' If saved lambda then run Edit lambda first.
        With LambdaFormulaCell
            If IsSavedLambdaInCellFormula(GetCellFormula(LambdaFormulaCell), .Worksheet) Then
                If IsLambdaCreatedInExcelLabs(.Worksheet.Parent, GetSavedNamedNameFromCellFormula(GetCellFormula(LambdaFormulaCell), .Worksheet)) Then
                    GoTo ExitMethod
                Else
                    EditLambda LambdaFormulaCell
                End If
            End If
        End With
        
        Dim InvalidOptReason As String
        InvalidOptReason = LambdaToLetOperationInvalidMessage(LambdaFormulaCell, PutLetOnCell)
        If InvalidOptReason <> vbNullString Then
            MsgBox InvalidOptReason, vbExclamation + vbOKOnly, "LAMBDA To LET"
            GoTo ExitMethod
        Else
            ' If not an undo operation, store the current formula before conversion.
            OldFormula = PutLetOnCell.Formula2
        End If
        
    End If
    
    DeleteComment LambdaFormulaCell
    ' Begin error handling. If an error occurs anywhere in the code that follows,
    ' the program will jump to the ErrorHandler label.
    On Error GoTo ErrorHandler
    ' Convert the formula in the LambdaFormulaCell from LAMBDA to LET.
    Dim FormulaText As String
    FormulaText = ConvertLambdaToLet(GetCellFormula(LambdaFormulaCell))
    ' Assign the converted formula to the target cell.
    PutLetOnCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(FormulaText)
    ' If it's not an undo operation, set up for a possible future undo operation.
    If Not IsUndo Then
        ' Store the cell that has the converted formula for potential future undo.
        If GetRangeRefWithSheetName(LambdaFormulaCell) = GetRangeRefWithSheetName(PutLetOnCell) Then
            Set PutFormulaOnUndo = LambdaFormulaCell
        Else
            Set PutFormulaOnUndo = PutLetOnCell
        End If
        ' Assign the action name "LambdaToLet" for future undo.
        AssingOnUndo "LambdaToLet"
    End If
    ' Exit the subroutine successfully.
    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modAuditLambdaSteps.LambdaToLet"
    
ExitMethod:
    Context.ClearContext METHOD_NAME
    Exit Sub
    
ErrorHandler:
    
    Context.ClearContext METHOD_NAME
    ' Capture the error number and description, if any
    Dim ErrorNumber As Long
    ErrorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description

    ' Print a message to the immediate window to help with debugging
    Debug.Print "Converted Lambda To Let : " & FormulaText
    ' If an error was raised, re-raise it with its original details and resume execution for debugging
    If ErrorNumber <> 0 Then
        Err.Raise ErrorNumber, Err.Source, ErrorDescription
        Resume
    End If
    Logger.Log TRACE_LOG, "Exit modAuditLambdaSteps.LambdaToLet"
    
End Sub

Private Function LambdaToLetOperationInvalidMessage(ByVal LambdaFormulaCell As Range _
                                                    , ByVal PutLetOnCell As Range) As String
    
    Dim Result As String
    ' Validation: Ensure only one cell is selected to convert from LAMBDA to LET.
    If LambdaFormulaCell.Cells.CountLarge > 1 Then
        Result = "Unable to convert " & LAMBDA_FN_NAME & " to " & LET_FN_NAME _
                 & ". Only one cell at a time allowed."
    ElseIf Not LambdaFormulaCell.HasFormula Then
        Result = "No formula found on " & LambdaFormulaCell.Address & " ."
    ElseIf Not IsLambdaFunction(GetCellFormula(LambdaFormulaCell)) Then
        Result = "The formula is not a LAMBDA formula.  Procedure aborted."
    ElseIf LambdaFormulaCell.Address(, , , True) <> PutLetOnCell.Address(, , , True) Then
        
        ' If target cell is not empty or contains errors, show a message and exit the method.
        If IsError(PutLetOnCell) Then
            Result = "Unable to convert " & LAMBDA_FN_NAME & " to " & LET_FN_NAME _
                     & ". Destination cell not empty."
        
        ElseIf PutLetOnCell.Value <> vbNullString Or PutLetOnCell.HasFormula Then
            Result = "Unable to convert " & LAMBDA_FN_NAME & " to " & LET_FN_NAME _
                     & ". Destination cell not empty."
        End If
        
    End If
    
    LambdaToLetOperationInvalidMessage = Result
    
End Function

Private Sub LambdaToLet_Undo()
    LambdaToLet Nothing, Nothing, True
End Sub

Public Function ConvertLambdaToLet(ByVal FormulaText As String) As String
    
    Logger.Log TRACE_LOG, "Enter modAuditLambdaSteps.ConvertLambdaToLet"

    ' If the formula already starts with "=LET", there's no need for conversion
    If IsLetFunction(FormulaText) Then
        ConvertLambdaToLet = FormulaText
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modAuditLambdaSteps.ConvertLambdaToLet"
        Exit Function
    End If

    ' Break down the lambda formula into parts
    Dim LambdaParts As Variant
    LambdaParts = GetDependencyFunctionResult(FormulaText, LAMBDA_PARTS)

    ' If the breakdown doesn't return an array, it means no LAMBDA function was found in the cell
    If Not IsArray(LambdaParts) Then
        MsgBox "No " & LAMBDA_FN_NAME & " function found.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modAuditLambdaSteps.ConvertLambdaToLet"
        Exit Function
    End If

    Dim FirstColIndex As Long
    FirstColIndex = LBound(LambdaParts, 2)
    Dim ArgumentAsLetVar As String

    Dim VarName As String
    Dim CellAddress As Variant
    Dim CurrentRowIndex As Long
    
    Dim LetPart As String
    ' Retrieve the corresponding part of the LambdaParts array for conversion
    LetPart = EQUAL_SIGN & LambdaParts(UBound(LambdaParts, 1), FirstColIndex)
    
    Dim Counter As Long
    ' Loop through each row in the LambdaParts array except for the last one
    For CurrentRowIndex = LBound(LambdaParts, 1) To UBound(LambdaParts, 1) - 1
        ' Clean up the variable name by removing unwanted characters
        VarName = modUtility.CleanVarName(CStr(LambdaParts(CurrentRowIndex, FirstColIndex)))
        VarName = Text.RemoveFromEndIfPresent(VarName, RIGHT_BRACKET)
        VarName = Text.RemoveFromStartIfPresent(VarName, LEFT_BRACKET)

        ' Get the cell address from the LambdaParts array
        CellAddress = LambdaParts(CurrentRowIndex, FirstColIndex + LAMBDA_PARTS_VALUE_COL_INDEX - 1)
        LetPart = InsertLetStep(LetPart, Counter + 1, VarName, CStr(CellAddress))
        Counter = Counter + 1
    Next CurrentRowIndex
    
    ' Replace new lines in LetPart with a line break, and remove leading spaces or new lines
    ConvertLambdaToLet = ReplaceNewlineWithChar10(LetPart)
    
    ' Log the final LetPart and exit function
    Logger.Log DEBUG_LOG, "Final LET part : " & NEW_LINE & LetPart
    Logger.Log TRACE_LOG, "Exit modAuditLambdaSteps.ConvertLambdaToLet"
    
End Function


