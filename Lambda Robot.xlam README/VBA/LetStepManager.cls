VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LetStepManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed
'@PredeclaredId
'@Folder "Step.Manager"
Option Explicit

Private Const ADD_LET_STEP_SUB_NAME As String = "LetStepManager.AddLETStep"
Private Const REMOVE_LET_STEP_SUB_NAME As String = "LetStepManager.RemoveLastLETStep"
Private Const CYCLE_LET_STEPS_SUB_NAME As String = "LetStepManager.CycleLETSteps"
Private Const CYCLE_LAST_STEP_IDENTIFIER As String = "CycleLETResultStep"

Private Const PREVIOUS_STEP_MARKER As String = "{{LastStep}}"
Private Const DEFAULT_STEP_NAME As String = "_Step1"
Private Const DEFAULT_FIRST_STEP_NAME As String = "_Source"

Private Type TLetStepManager
    AllNamedRanges As Collection
    OnBook As Workbook
    FormulaText As String
    IsLambdaFormulaWithLet As Boolean
    IsCellStartWithLetFormula As Boolean
    IsValidForProcessingFurther As Boolean
    LetParts As Variant
    LambdaParts As Variant
    LambdaAndLetStepsName As Variant
    LastStepName As String
End Type

Private this As TLetStepManager

'@EntryPoint
Public Sub AddLetStep(ByVal FormulaCell As Range, ByVal StepName As String _
                                                 , ByVal StepFormula As String _
                                                  , Optional ByVal TargetCell As Range)

    Logger.Log TRACE_LOG, "Enter LetStepManager.AddLetStep"
    ' Validating the input arguments
    ValidateInputArguments FormulaCell, StepFormula
    If StepName = vbNullString Then StepName = DEFAULT_STEP_NAME
    If Not Text.IsStartsWith(StepName, UNDER_SCORE) Then StepName = UNDER_SCORE & StepName
    
    On Error GoTo ErrorHandler
    ' Retrieving the context data
    RetriveContextData FormulaCell
    ' Getting a valid step name
    StepName = GetValidStepName(StepName, this.AllNamedRanges, this.LambdaAndLetStepsName)
    Dim FormulaText As String
    FormulaText = modDependencyLambdaResult.AddLetStep(this.FormulaText, StepName, StepFormula)
    ' Checking if TargetCell is null
    If IsNothing(TargetCell) Then Set TargetCell = FormulaCell
    ' Replacing the formula in the target cell
    TargetCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(FormulaText)
    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword LetStepManager.AddLetStep"
    Exit Sub
    
ErrorHandler:
    
    Dim ErrorNumber As Long
    ErrorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description
    
    ' Raising the error if ErrorNumber is not 0
    If ErrorNumber <> 0 Then
        Err.Raise ErrorNumber, Err.Source, ErrorDescription
        ' This is only for debugging purpose.
        Resume
    End If
    Logger.Log TRACE_LOG, "Exit LetStepManager.AddLetStep"

End Sub

Private Function ConvertToALetStatement(ByVal PlainFormula As String _
                                        , Optional ByVal LastStepName As String = DEFAULT_FIRST_STEP_NAME) As String

    Logger.Log TRACE_LOG, "Enter LetStepManager.ConvertToALetStatement"
    ' Converting a plain formula to a LET statement
    Dim FinalFormula As String
    FinalFormula = EQUAL_LET_FIRST_PAREN & NEW_LINE _
                   & THREE_SPACE & LastStepName & LIST_SEPARATOR _
                   & ONE_SPACE & Text.RemoveFromStartIfPresent(PlainFormula, EQUAL_SIGN)

    ' Completing the LET statement
    FinalFormula = FinalFormula & LIST_SEPARATOR & NEW_LINE & THREE_SPACE _
                   & LastStepName & NEW_LINE & FIRST_PARENTHESIS_CLOSE

    ' Returning the LET statement
    ConvertToALetStatement = FinalFormula
    Logger.Log TRACE_LOG, "Exit LetStepManager.ConvertToALetStatement"

End Function

'@Pure
Private Function GetCleanLetVarName(ByVal LetVarName As String) As String
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.GetCleanLetVarName"
    ' Cleaning up the input variable name by removing potential undesirable characters and spaces
    Dim CleanLetVarName As String
    CleanLetVarName = Text.RemoveFromStartIfPresent(LetVarName, vbNewLine)
    CleanLetVarName = Text.RemoveFromStartIfPresent(CleanLetVarName, Chr$(10))
    CleanLetVarName = Text.Trim(CleanLetVarName)
    GetCleanLetVarName = CleanLetVarName
    Logger.Log TRACE_LOG, "Exit LetStepManager.GetCleanLetVarName"
    
End Function

'@Impure
Private Sub RetriveContextData(ByVal FormulaCell As Range)
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.RetriveContextData"
    ' Retrieving all context data necessary for processing the input formula cell
    Set this.OnBook = FormulaCell.Worksheet.Parent
    Set this.AllNamedRanges = Context.GetAllNamedRangeCollectionFromBook(this.OnBook)
    this.FormulaText = GetCellFormula(FormulaCell)
    this.IsLambdaFormulaWithLet = IsLambdaWithLet(this.FormulaText)
    this.IsCellStartWithLetFormula = IsLetFunction(this.FormulaText)

    this.IsValidForProcessingFurther = IsStartingFormulaIsTheEntireFormula(this.FormulaText)
    ' If the formula is not valid for processing, it's converted to a LET statement
    If Not this.IsValidForProcessingFurther Then
        this.IsLambdaFormulaWithLet = False
        this.FormulaText = ConvertToALetStatement(this.FormulaText)
    End If

    ' Retrieving more context data
    this.LetParts = GetDependencyFunctionResult(this.FormulaText, LET_PARTS)
    this.LambdaParts = GetDependencyFunctionResult(this.FormulaText, LAMBDA_PARTS)
    this.LambdaAndLetStepsName = GetParametersAndStepsName(this.FormulaText)
    this.LastStepName = FindLastStepName(this.LambdaAndLetStepsName)
    Logger.Log TRACE_LOG, "Exit LetStepManager.RetriveContextData"
    
End Sub

'@Pure
Private Function FindLastStepName(ByVal LambdaAndLetStepsName As Variant) As String
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.FindLastStepName"
    ' Iterating through all the step names and finding the last valid (non-error) step name
    If Not IsArray(LambdaAndLetStepsName) Then
        FindLastStepName = vbNullString
        Exit Function
    End If
    
    Dim LastStepName As String
    Dim CurrentStepName As Variant
    For Each CurrentStepName In LambdaAndLetStepsName
        If Not IsError(CurrentStepName) Then
            If Trim$(CurrentStepName) <> vbNullString Then LastStepName = CurrentStepName
        End If
    Next CurrentStepName
    FindLastStepName = GetCleanLetVarName(LastStepName)
    Logger.Log TRACE_LOG, "Exit LetStepManager.FindLastStepName"
    
End Function

'@Pure
Private Function GetValidStepName(ByVal StepName As String, ByVal AllNamedRanges As Collection _
                                                           , ByVal LambdaAndLetStepsName As Variant) As String
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.GetValidStepName"
    ' Generating a valid step name and ensuring its uniqueness
    Dim ValidStepName As String
    ValidStepName = MakeValidLetVarName(StepName, GetNamingConv(False))
    Dim AllNameMap As Collection
    Set AllNameMap = CombineNamedRangesWithStepsName(AllNamedRanges, LambdaAndLetStepsName)
    ValidStepName = FindUniqueNameByIncrementingNumber(AllNameMap, ValidStepName)
    GetValidStepName = ValidStepName
    Logger.Log TRACE_LOG, "Exit LetStepManager.GetValidStepName"
    
End Function

'@Pure
Private Function CombineNamedRangesWithStepsName(ByVal NamedRangeColl As Collection _
                                                 , ByVal LambdaAndLetStepsName As Variant) As Collection
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.CombineNamedRangesWithStepsName"
    ' Combining Named Ranges and Step Names into a single collection
    Dim Result As Collection
    Set Result = New Collection
    If IsNotNothing(NamedRangeColl) Then Set Result = NamedRangeColl
    Dim CurrentStepName As Variant
    For Each CurrentStepName In LambdaAndLetStepsName
        CurrentStepName = CleanVarName(CStr(CurrentStepName))
        If Trim$(CurrentStepName) <> vbNullString Then
            Result.Add CStr(CurrentStepName), CStr(CurrentStepName)
        End If
    Next CurrentStepName
    Set CombineNamedRangesWithStepsName = Result
    Logger.Log TRACE_LOG, "Exit LetStepManager.CombineNamedRangesWithStepsName"
    
End Function

'@Pure
Private Sub ValidateInputArguments(ByVal FormulaCell As Range, ByVal StepFormula As String)
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.ValidateInputArguments"
    ValidateFormulaCellParam FormulaCell, ADD_LET_STEP_SUB_NAME
    
    ' Validate StepFormula parameter
    If StepFormula = vbNullString Then Err.Raise 13, ADD_LET_STEP_SUB_NAME, "Invalid StepFormula param."
    If Not Text.Contains(StepFormula, PREVIOUS_STEP_MARKER) Then Err.Raise 13 _
       , ADD_LET_STEP_SUB_NAME, "Invalid StepFormula param."
    Logger.Log TRACE_LOG, "Exit LetStepManager.ValidateInputArguments"
    
End Sub

'@EntryPoint
Public Sub RemoveLastLETStep(ByVal FormulaCell As Range, Optional ByVal TargetCell As Range)
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.RemoveLastLETStep"
    ' First validate the formula cell parameters and retrieve the context data
    ValidateFormulaCellParam FormulaCell, REMOVE_LET_STEP_SUB_NAME
    On Error GoTo ErrorHandler
    RetriveContextData FormulaCell
    
    If Not (this.IsLambdaFormulaWithLet Or this.IsCellStartWithLetFormula) Then
        MsgBox "Remove Last LET Steps command is only for cell having " _
               & LET_FX_NAME & " or " & LAMBDA_FX_NAME & " formula." _
               , vbCritical, "Remove Last LET Steps"
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword LetStepManager.RemoveLastLETStep"
        Exit Sub
    End If
    
    Dim ReplaceOnFormula As String
    Dim ReplacedFormula As String
    ' Depending on whether cell starts with a lambda formula, the removal of the step varies
    If this.IsLambdaFormulaWithLet Then
        ReplaceOnFormula = this.LambdaParts(UBound(this.LambdaParts, 1), LBound(this.LambdaParts, 2))
        ReplacedFormula = GetFormulaAfterRemovingLastStep(EQUAL_SIGN & ReplaceOnFormula)
        ReplacedFormula = GetUptoLambdaParamDefPart(this.FormulaText) _
                          & ReplacedFormula & FIRST_PARENTHESIS_CLOSE _
                          & GetLambdaInvocationPart(this.FormulaText)
    Else
        ReplaceOnFormula = this.FormulaText
        ReplacedFormula = EQUAL_SIGN & GetFormulaAfterRemovingLastStep(ReplaceOnFormula)
    End If
    Logger.Log DEBUG_LOG, "Removing last let step from formula:" & ReplaceOnFormula
    Logger.Log DEBUG_LOG, "Final formula after Removing last let step:" & ReplacedFormula
    If IsNothing(TargetCell) Then Set TargetCell = FormulaCell
    TargetCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(ReplacedFormula)
    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword LetStepManager.RemoveLastLETStep"
    Exit Sub
    
ErrorHandler:
    ' If an error occurs, clean up resources and raise error
    Dim ErrorNumber As Long
    ErrorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description
    
    If ErrorNumber <> 0 Then
        Err.Raise ErrorNumber, Err.Source, ErrorDescription
        Resume
    End If
    Logger.Log TRACE_LOG, "Exit LetStepManager.RemoveLastLETStep"

End Sub

'@EntryPoint
Public Sub CycleLETSteps(ByVal FormulaCell As Range _
                         , Optional ByVal TargetCell As Range _
                          , Optional ByVal IsReset As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.CycleLETSteps"
    ' First validate the formula cell parameters and retrieve the context data
    ValidateFormulaCellParam FormulaCell, CYCLE_LET_STEPS_SUB_NAME
    On Error GoTo ErrorHandler
    RetriveContextData FormulaCell
    
    If Not (this.IsLambdaFormulaWithLet Or this.IsCellStartWithLetFormula) Then
        MsgBox "Cycle LET Steps command is only for cells having a " & LET_FX_NAME _
               & " or " & LAMBDA_FX_NAME & "/" _
               & LET_FX_NAME & " formula.", vbCritical, "Cycle LET Steps"
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword LetStepManager.CycleLETSteps"
        Exit Sub
    End If
    
    Dim ReplaceOnFormula As String
    Dim ReplacedFormula As String
    ' Depending on whether cell starts with a lambda formula, the removal of the step varies
    If this.IsLambdaFormulaWithLet Then
        ReplaceOnFormula = this.LambdaParts(UBound(this.LambdaParts, 1), LBound(this.LambdaParts, 2))
        ReplacedFormula = GetFormulaAfterCycleLetStep(EQUAL_SIGN & ReplaceOnFormula, IsReset)
        ReplacedFormula = GetUptoLambdaParamDefPart(this.FormulaText) _
                          & ReplacedFormula & FIRST_PARENTHESIS_CLOSE _
                          & GetLambdaInvocationPart(this.FormulaText)
    Else
        ReplaceOnFormula = this.FormulaText
        ReplacedFormula = EQUAL_SIGN & GetFormulaAfterCycleLetStep(ReplaceOnFormula, IsReset)
    End If
    
    Logger.Log DEBUG_LOG, "Removing last let step from formula:" & ReplaceOnFormula
    Logger.Log DEBUG_LOG, "Final formula after Removing last let step:" & ReplacedFormula
    If IsNothing(TargetCell) Then Set TargetCell = FormulaCell
    
    TargetCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(ReplacedFormula)
    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword LetStepManager.CycleLETSteps"
    Exit Sub
    
ErrorHandler:
    ' If an error occurs, clean up resources and raise error
    Dim ErrorNumber As Long
    ErrorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description
    
    If ErrorNumber <> 0 Then
        Err.Raise ErrorNumber, Err.Source, ErrorDescription
        Resume
    End If
    Logger.Log TRACE_LOG, "Exit LetStepManager.CycleLETSteps"

End Sub

'@Pure
Private Function GetFormulaAfterCycleLetStep(ByVal LetFormula As String, ByVal IsReset As Boolean) As String
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.GetFormulaAfterCycleLetStep"
    If Not IsLetFunction(LetFormula) Then
        GetFormulaAfterCycleLetStep = Text.RemoveFromStartIfPresent(LetFormula, EQUAL_SIGN)
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword LetStepManager.GetFormulaAfterCycleLetStep"
        Exit Function
    End If

    Dim LetParts As Variant
    LetParts = GetLetParts(LetFormula)
    
    If IsCyclingNotNeeded(LetParts) Then
        GetFormulaAfterCycleLetStep = Text.RemoveFromStartIfPresent(LetFormula, EQUAL_SIGN)
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword LetStepManager.GetFormulaAfterCycleLetStep"
        Exit Function
    End If

    LetParts = UpdateLetPartsForCycling(LetParts, IsReset)
    GetFormulaAfterCycleLetStep = ReconstructFormula(LetParts)
    Logger.Log TRACE_LOG, "Exit LetStepManager.GetFormulaAfterCycleLetStep"
    
End Function

Private Function GetLetParts(ByVal LetFormula As String) As Variant
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.GetLetParts"
    Dim LetParts As Variant
    LetParts = GetDependencyFunctionResult(LetFormula, LET_PARTS)
    TrimAndCleanStepsName LetParts
    GetLetParts = LetParts
    Logger.Log TRACE_LOG, "Exit LetStepManager.GetLetParts"
    
End Function

Private Sub TrimAndCleanStepsName(ByRef LetParts As Variant)
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.TrimAndCleanStepsName"
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(LetParts, 2)
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(LetParts, 1) To UBound(LetParts, 1) - 1
        Dim StepName As String
        StepName = LetParts(CurrentRowIndex, FirstColumnIndex)
        StepName = GetCleanLetVarName(StepName)
        LetParts(CurrentRowIndex, FirstColumnIndex) = StepName
    Next CurrentRowIndex
    StepName = LetParts(UBound(LetParts, 1), FirstColumnIndex)
    StepName = RTrim$(LTrim$(Text.Clean(StepName)))
    LetParts(UBound(LetParts, 1), FirstColumnIndex) = StepName
    Logger.Log TRACE_LOG, "Exit LetStepManager.TrimAndCleanStepsName"
    
End Sub

Private Sub GetRowAndColumnDetails(ByVal LetParts As Variant, ByRef FirstColumnIndex As Long _
                                                             , ByRef LastRowIndex As Long _
                                                              , ByRef FirstRowIndex As Long _
                                                               , ByRef NumberOfRow As Long)
    Logger.Log TRACE_LOG, "Enter LetStepManager.GetRowAndColumnDetails"
    FirstColumnIndex = LBound(LetParts, 2)
    LastRowIndex = UBound(LetParts, 1)
    FirstRowIndex = LBound(LetParts, 1)
    NumberOfRow = LastRowIndex - FirstRowIndex + 1
    Logger.Log TRACE_LOG, "Exit LetStepManager.GetRowAndColumnDetails"
    
End Sub

Private Sub GetStepNames(ByVal LetParts As Variant, ByVal LastRowIndex As Long _
                                                   , ByVal FirstColumnIndex As Long _
                                                    , ByRef LastStepName As String _
                                                     , ByRef SecondLastStepName As String)
    Logger.Log TRACE_LOG, "Enter LetStepManager.GetStepNames"
    LastStepName = LetParts(LastRowIndex, FirstColumnIndex)
    SecondLastStepName = LetParts(LastRowIndex - 1, FirstColumnIndex)
    Logger.Log TRACE_LOG, "Exit LetStepManager.GetStepNames"
    
End Sub

Private Function IsCyclingNotNeeded(ByVal LetParts As Variant) As Boolean
                                                        
    Logger.Log TRACE_LOG, "Enter LetStepManager.IsCyclingNotNeeded"
    Dim NumberOfRow As Long
    NumberOfRow = UBound(LetParts, 1) - LBound(LetParts, 1) + 1
    Dim LastStepName As String
    LastStepName = LetParts(UBound(LetParts, 1), LBound(LetParts, 2))
    Dim SecondLastStepName As String
    SecondLastStepName = LetParts(UBound(LetParts, 1) - 1, LBound(LetParts, 2))
    
    IsCyclingNotNeeded = (NumberOfRow = 2 And LastStepName = SecondLastStepName)
    Logger.Log TRACE_LOG, "Exit LetStepManager.IsCyclingNotNeeded"
    
End Function

Private Function UpdateLetPartsForCycling(ByVal LetParts As Variant, ByVal IsReset As Boolean) As Variant
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.UpdateLetPartsForCycling"
    Dim FirstColumnIndex As Long
    Dim FirstRowIndex As Long
    Dim LastRowIndex As Long
    Dim NumberOfRow As Long
    GetRowAndColumnDetails LetParts, FirstColumnIndex, LastRowIndex, FirstRowIndex, NumberOfRow
    
    Dim LastStepName As String
    LastStepName = LetParts(LastRowIndex, FirstColumnIndex)
    Dim SecondLastStepName As String
    SecondLastStepName = LetParts(LastRowIndex - 1, FirstColumnIndex)
    
    Dim CurrentLastStepIndex As Long
    CurrentLastStepIndex = FirstIndexOf(LetParts, LastStepName, FirstColumnIndex, True)
    
    ' If the process is for reset then just refer the last step and exit
    If IsReset Then
        If Not (CurrentLastStepIndex = LastRowIndex Or CurrentLastStepIndex = LastRowIndex - 1) Then
            UpdateLetPartsIfReset LetParts, FirstColumnIndex, LastRowIndex
        End If
        
        UpdateLetPartsForCycling = LetParts
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword LetStepManager.UpdateLetPartsForCycling"
        Exit Function
    End If
    
    ' If we can find LastStepName  at the end that means it's not referring to any step but it's a calculation.
    ' So we have to add our special LETStep name to use that as calculation.
    If CurrentLastStepIndex = LastRowIndex Then
        LetParts(LastRowIndex, FirstColumnIndex + LET_PARTS_VALUE_COL_INDEX - 1) = LetParts(LastRowIndex, FirstColumnIndex)
        SecondLastStepName = GetCorrectCycleLetStepResultName(SecondLastStepName)
        LetParts(LastRowIndex, FirstColumnIndex) = SecondLastStepName
        LetParts = AddNRowsTo2DArray(LetParts, 1)
        ' As we are adding one row so now LastRow is incremented by one.
        LetParts(LastRowIndex + 1, FirstColumnIndex) = LetParts(FirstRowIndex, FirstColumnIndex)
        ' If it is referring the last step then we need to refer it back to first one.
    ElseIf CurrentLastStepIndex = LastRowIndex - 1 And LastStepName = SecondLastStepName Then
        LetParts(LastRowIndex, FirstColumnIndex) = LetParts(FirstRowIndex, FirstColumnIndex)
        ' But if it is referring the the third last step and second last step is our identifier then we have to remove our identifier and calc
    ElseIf CurrentLastStepIndex = LastRowIndex - 2 And Text.RemoveFromStartIfPresent(SecondLastStepName, UNDER_SCORE) = CYCLE_LAST_STEP_IDENTIFIER Then
        LetParts(LastRowIndex, FirstColumnIndex) = LTrim$(LetParts(LastRowIndex - 1, FirstColumnIndex + LET_PARTS_VALUE_COL_INDEX - 1))
        LetParts(LastRowIndex - 1, FirstColumnIndex + LET_PARTS_VALUE_COL_INDEX - 1) = vbNullString
        LetParts(LastRowIndex - 1, FirstColumnIndex) = vbNullString
    Else
        LetParts(LastRowIndex, FirstColumnIndex) = LetParts(CurrentLastStepIndex + 1, FirstColumnIndex)
    End If

    UpdateLetPartsForCycling = LetParts
    Logger.Log TRACE_LOG, "Exit LetStepManager.UpdateLetPartsForCycling"
    
End Function

Private Sub UpdateLetPartsIfReset(ByRef LetParts As Variant, ByVal FirstColumnIndex As Long _
                                                            , ByVal LastRowIndex As Long)
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.UpdateLetPartsIfReset"
    Dim SecondLastStepName As String
    SecondLastStepName = LetParts(LastRowIndex - 1, FirstColumnIndex)
    If Text.RemoveFromStartIfPresent(SecondLastStepName, UNDER_SCORE) = CYCLE_LAST_STEP_IDENTIFIER Then
        LetParts(LastRowIndex, FirstColumnIndex) = LTrim$(LetParts(LastRowIndex - 1, FirstColumnIndex + LET_PARTS_VALUE_COL_INDEX - 1))
        LetParts(LastRowIndex - 1, FirstColumnIndex + LET_PARTS_VALUE_COL_INDEX - 1) = vbNullString
        LetParts(LastRowIndex - 1, FirstColumnIndex) = vbNullString
    Else
        LetParts(LastRowIndex, FirstColumnIndex) = SecondLastStepName
    End If
    Logger.Log TRACE_LOG, "Exit LetStepManager.UpdateLetPartsIfReset"
    
End Sub

Private Function GetCorrectCycleLetStepResultName(ByVal PreviousStepName As String) As String
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.GetCorrectCycleLetStepResultName"
    If Text.IsStartsWith(PreviousStepName, UNDER_SCORE) Then
        GetCorrectCycleLetStepResultName = UNDER_SCORE & CYCLE_LAST_STEP_IDENTIFIER
    Else
        GetCorrectCycleLetStepResultName = CYCLE_LAST_STEP_IDENTIFIER
    End If
    Logger.Log TRACE_LOG, "Exit LetStepManager.GetCorrectCycleLetStepResultName"
    
End Function

Private Function ReconstructFormula(ByVal LetParts As Variant) As String
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.ReconstructFormula"
    Dim FinalFormula As String
    FinalFormula = LET_AND_OPEN_PAREN & NEW_LINE
    
    Dim FirstColumnIndex As Long
    FirstColumnIndex = LBound(LetParts, 2)

    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(LetParts, 1) To UBound(LetParts, 1) - 1
        
        Dim LetVarName As String
        LetVarName = CStr(LetParts(CurrentRowIndex, FirstColumnIndex))
        If Text.Trim(LetVarName) <> vbNullString Then
            LetVarName = GetCleanLetVarName(LetVarName)
            Dim LetVarValue As String
            LetVarValue = LTrim$(LetParts(CurrentRowIndex, FirstColumnIndex + LET_PARTS_VALUE_COL_INDEX - 1))
            FinalFormula = FinalFormula & THREE_SPACE & LetVarName & LIST_SEPARATOR & ONE_SPACE & LetVarValue _
                           & LIST_SEPARATOR & NEW_LINE
        End If
        
    Next CurrentRowIndex
    
    ' Finalize the reconstructed formula and return
    FinalFormula = FinalFormula & THREE_SPACE & LetParts(UBound(LetParts, 1), FirstColumnIndex) _
                   & NEW_LINE & FIRST_PARENTHESIS_CLOSE
                   
    ReconstructFormula = FinalFormula
    Logger.Log TRACE_LOG, "Exit LetStepManager.ReconstructFormula"
    
End Function

'@Pure
Private Function GetFormulaAfterRemovingLastStep(ByVal LetFormula As String) As String
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.GetFormulaAfterRemovingLastStep"
    ' Check if the outermost function is 'Let'. If not, remove the equal sign and return
    If Not IsLetFunction(LetFormula) Then
        GetFormulaAfterRemovingLastStep = Text.RemoveFromStartIfPresent(LetFormula, EQUAL_SIGN)
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword LetStepManager.GetFormulaAfterRemovingLastStep"
        Exit Function
    End If
    
    ' If 'Let' is the outermost function, get the function result and parse it
    Dim LetParts As Variant
    LetParts = GetDependencyFunctionResult(LetFormula, LET_PARTS)
    
    Dim NumberOfRow As Long
    NumberOfRow = UBound(LetParts, 1) - LBound(LetParts, 1) + 1
    
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(LetParts, 2)
    
    ' If there are only 2 rows, remove the outermost function and return
    If NumberOfRow = 2 Then
        GetFormulaAfterRemovingLastStep = LTrim$(LetParts(LBound(LetParts, 1) _
                                                          , FirstColumnIndex + LET_PARTS_VALUE_COL_INDEX - 1))
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword LetStepManager.GetFormulaAfterRemovingLastStep"
        Exit Function
    End If
    
    ' If there are more than 2 rows, reconstruct the formula without the last step
    Dim CurrentRowIndex As Long
    Dim FinalFormula As String
    FinalFormula = LET_AND_OPEN_PAREN & NEW_LINE
    Dim LetVarName As String
    Dim LetVarValue As String
    For CurrentRowIndex = LBound(LetParts, 1) To UBound(LetParts, 1) - 2
        
        LetVarName = GetCleanLetVarName(CStr(LetParts(CurrentRowIndex, FirstColumnIndex)))
        LetVarValue = LTrim$(LetParts(CurrentRowIndex, FirstColumnIndex + LET_PARTS_VALUE_COL_INDEX - 1))
        FinalFormula = FinalFormula & THREE_SPACE & LetVarName & LIST_SEPARATOR _
                       & ONE_SPACE & LetVarValue & LIST_SEPARATOR & NEW_LINE
        
    Next CurrentRowIndex
                   
    ' Finalize the reconstructed formula and return
    FinalFormula = FinalFormula & THREE_SPACE & LetVarName _
                   & NEW_LINE & FIRST_PARENTHESIS_CLOSE
                   
    GetFormulaAfterRemovingLastStep = FinalFormula
    Logger.Log TRACE_LOG, "Exit LetStepManager.GetFormulaAfterRemovingLastStep"
    
End Function

'@Pure
Private Sub ValidateFormulaCellParam(ByVal FormulaCell As Range, ByVal ErrorSource As String)
    
    Logger.Log TRACE_LOG, "Enter LetStepManager.ValidateFormulaCellParam"
    ' Validate FormulaCell parameter
    Const ERROR_DESCRIPTION As String = "Invalid FormulaCell param."
    If IsNothing(FormulaCell) Then Err.Raise 13, ErrorSource, ERROR_DESCRIPTION
    If FormulaCell.Cells.Count > 1 Then Err.Raise 13, ErrorSource, ERROR_DESCRIPTION
    If Not FormulaCell.HasFormula Then Err.Raise 13, ErrorSource, ERROR_DESCRIPTION
    Logger.Log TRACE_LOG, "Exit LetStepManager.ValidateFormulaCellParam"
    
End Sub

Private Sub Class_Initialize()
    Set this.AllNamedRanges = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.AllNamedRanges = Nothing
End Sub


