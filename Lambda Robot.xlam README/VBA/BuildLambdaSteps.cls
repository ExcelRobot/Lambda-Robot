VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BuildLambdaSteps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable
'@Folder "Step.Utility"
'@PredeclaredId
Option Explicit

Private Enum FunctionType
    LAMBDA_FROM_NAME_MANAGER = 1
    LAMBDA_IN_CELL = 2
    LET_IN_CELL = 3
    OTHER_FUNCTION_OR_BLANK = 4
End Enum

Private Type TBuildLambdaSteps
    LambdaInvocationCell As Range
    StepStartCell As Range
    LetParts As Variant
    LambdaArgumentsName As Collection
    LambdaArgumentsRefCell As Collection
    ReasonToBeNotValid As String
    VarNameVsCalculationAreaMap As Collection
    CellFunctionType As FunctionType
    LambdaName As String
    LambdaCalculationPart As String
    FormulaToAudit As String
    ResultBeforeCommand As Variant
    ResultAfterCommand As Variant
    UsedRangeRequiredFXsName As Collection
    InvocationCellTableName As String
End Type
Private this As TBuildLambdaSteps

Public Property Get InvocationCellTableName() As String
    InvocationCellTableName = this.InvocationCellTableName
End Property

Public Property Let InvocationCellTableName(ByVal RHS As String)
    this.InvocationCellTableName = RHS
End Property

Public Property Get ReasonToBeNotValid() As String
    ReasonToBeNotValid = this.ReasonToBeNotValid
End Property

Public Property Get LambdaInvocationCell() As Range
    Set LambdaInvocationCell = this.LambdaInvocationCell
End Property

Public Property Set LambdaInvocationCell(ByVal RHS As Range)
    
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.LambdaInvocationCell"
    Set this.LambdaInvocationCell = RHS
    this.ResultBeforeCommand = GetFormulaResult(RHS)
    UpdateBasicStats RHS
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.LambdaInvocationCell"
    
End Property

Private Sub UpdateBasicStats(ByRef LambdaInvocationCell As Range)
    
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.UpdateBasicStats"
    Dim FormulaText As String
    FormulaText = GetCellFormula(LambdaInvocationCell)
    If Not IsStartingFormulaIsTheEntireFormula(FormulaText) Then
        this.CellFunctionType = OTHER_FUNCTION_OR_BLANK
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword BuildLambdaSteps.UpdateBasicStats"
        Exit Sub
    End If
    
    Set this.UsedRangeRequiredFXsName = GetUsedRangeRequiredFXsName(FormulaText)
    
    Dim FunctionName As String
    FunctionName = modUtility.ExtractStartFormulaName(FormulaText)
    Dim FunctionDef As String
    ' Identifying the function type and setting up the function definition
    Select Case FunctionName
        Case LET_FN_NAME
            this.CellFunctionType = LET_IN_CELL
            FunctionDef = FormulaText
        Case LAMBDA_FN_NAME
            this.CellFunctionType = LAMBDA_IN_CELL
            FunctionDef = FormulaText
            this.LambdaName = modUtility.GetOldNameFromComment(LambdaInvocationCell, LAMBDA_NAME_NOTE_PREFIX)
        Case Else
            FunctionDef = GetFunctionDefFromNamedRange(FunctionName, FormulaText)
            
    End Select
    
    If this.CellFunctionType = OTHER_FUNCTION_OR_BLANK Then Exit Sub
    this.FormulaToAudit = FunctionDef
    If this.CellFunctionType = LAMBDA_FROM_NAME_MANAGER _
       Or this.CellFunctionType = LAMBDA_IN_CELL Then
        UpdateLambdaArgumentsName FunctionDef
    End If
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.UpdateBasicStats"
    
End Sub

'@Impure
Private Function GetFunctionDefFromNamedRange(ByVal FunctionName As String _
                                              , ByVal FormulaText As String) As String
     
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.GetFunctionDefFromNamedRange"
    Dim FunctionDef As String
    Dim CurrentName As Name
    On Error Resume Next
    Set CurrentName = LambdaInvocationCell.Worksheet.Parent.Names.Item(FunctionName)
    On Error GoTo 0
    ' Checking if the function type is neither Lambda nor Let
    If IsNothing(CurrentName) Then
        this.CellFunctionType = OTHER_FUNCTION_OR_BLANK
        this.ReasonToBeNotValid = "No " & LAMBDA_FN_NAME & " or " & LET_FN_NAME & " Found"
    Else
        Dim RefersTo As String
        RefersTo = CurrentName.RefersTo
        '                WriteStringToTextFile RefersTo, ActiveWorkbook.Path & "\temp.txt"
        FunctionName = ExtractStartFormulaName(RefersTo)
        If IsLambdaFunction(RefersTo) Or IsLetFunction(RefersTo) Then
            
            this.LambdaName = CurrentName.Name
            this.CellFunctionType = LAMBDA_FROM_NAME_MANAGER
            RefersTo = Text.RemoveFromStartIfPresent(RefersTo, EQUAL_SIGN)
            If IsLocalScopeNamedRange(CurrentName.Name) Then
                FunctionDef = Replace(FormulaText, CurrentName.Name, RefersTo)
                FunctionDef = Replace(FunctionDef, ExtractNameFromLocalNameRange(CurrentName.Name), RefersTo)
            Else
                FunctionDef = Replace(FormulaText, CurrentName.NameLocal, RefersTo)
            End If
        Else
            this.CellFunctionType = OTHER_FUNCTION_OR_BLANK
            this.ReasonToBeNotValid = "No " & LAMBDA_FN_NAME & " or " & LET_FN_NAME & " Found"
        End If
    End If
    GetFunctionDefFromNamedRange = FunctionDef
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.GetFunctionDefFromNamedRange"
    
End Function

Private Sub UpdateLambdaArgumentsName(ByVal FormulaText As String)
    
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.UpdateLambdaArgumentsName"
    ' Extract parts from Lambda function
    Dim LambdaParts As Variant
    LambdaParts = GetDependencyFunctionResult(FormulaText, LAMBDA_PARTS, True)
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(LambdaParts, 2)
    
    this.LambdaCalculationPart = LambdaParts(UBound(LambdaParts, 1), FirstColumnIndex)

    Dim CurrentRowIndex As Long
    Dim CurrentArgumentName As String

    ' Loop through all parts of the Lambda function
    For CurrentRowIndex = LBound(LambdaParts, 1) To UBound(LambdaParts, 1) - 1
        CurrentArgumentName = LambdaParts(CurrentRowIndex, FirstColumnIndex)

        ' Clean the extracted argument name
        CurrentArgumentName = Text.RemoveFromEndIfPresent(CurrentArgumentName, RIGHT_BRACKET)
        CurrentArgumentName = Text.RemoveFromStartIfPresent(CurrentArgumentName, LEFT_BRACKET)
        
        Dim ArgumentInvocation As String
        ArgumentInvocation = LambdaParts(CurrentRowIndex, FirstColumnIndex + LAMBDA_PARTS_VALUE_COL_INDEX - 1)
        
        ' Store the cleaned argument name
        this.LambdaArgumentsName.Add Mapper.Create(CurrentArgumentName, ArgumentInvocation), CurrentArgumentName
        
    Next CurrentRowIndex
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.UpdateLambdaArgumentsName"
    
End Sub

Public Property Get StepStartCell() As Range
    Set StepStartCell = this.StepStartCell
End Property

Public Property Set StepStartCell(ByVal RHS As Range)
    Set this.StepStartCell = RHS
End Property

' @EntryPoint
Public Function Create(ByRef GivenLambdaInvocationCell As Range _
                       , ByRef GivenStepStartCell As Range) As BuildLambdaSteps
    
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.Create"
    ' Create new instance of BuildLambdaSteps
    Dim CurrentBuildLambdaSteps As BuildLambdaSteps
    Set CurrentBuildLambdaSteps = New BuildLambdaSteps

    ' Initialize instance with given inputs
    With CurrentBuildLambdaSteps
        Set .LambdaInvocationCell = GivenLambdaInvocationCell
        
        If IsInsideTable(GivenLambdaInvocationCell) Then
            .InvocationCellTableName = GivenLambdaInvocationCell.ListObject.Name
        Else
            .InvocationCellTableName = vbNullString
        End If
        
        Set .StepStartCell = GivenStepStartCell
    End With
    
    ' Return initialized instance
    Set Create = CurrentBuildLambdaSteps
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.Create"

End Function

Public Sub ConstructSteps()
    
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.ConstructSteps"
    ' Start of ConstructSteps procedure

    ' Checking if the cell function type is other or blank
    If this.CellFunctionType = OTHER_FUNCTION_OR_BLANK Then
        ' Displaying a message box if there's no LET or LAMBDA functions
        MsgBox "Unable to deconstruct formula steps. " & this.LambdaInvocationCell.Address(False, False) _
               & " has no " & LET_FN_NAME & " or " & LAMBDA_FN_NAME _
               & " functions or formula has extra part.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword BuildLambdaSteps.ConstructSteps"
        Exit Sub
    End If
    
    ' Deleting LET step named ranges having error
    DeleteLETStepNamedRangesHavingError this.StepStartCell.Worksheet.Parent
    Dim PutOnCell As Range
    Set PutOnCell = this.StepStartCell
    
    On Error GoTo HandleError
    Dim CalcMode As XlCalculation
    CalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    Dim HeadersRange As Range
    ' If it is lambda either in cell or in name manager.
    If this.CellFunctionType = LAMBDA_FROM_NAME_MANAGER Or this.CellFunctionType = LAMBDA_IN_CELL Then
        AuditLambdaParameters PutOnCell, HeadersRange
        ' update formula part that we need to audit
        this.FormulaToAudit = EQUAL_SIGN & this.LambdaCalculationPart
    End If
    
    Dim CellReference As String
    
    this.LetParts = GetDependencyFunctionResult(this.FormulaToAudit, LET_PARTS, True)
    
    If Not IsArrayAllocated(this.LetParts) Then
        PlaceStepFormula PutOnCell, HeadersRange, CellReference _
                                                 , EQUAL_SIGN & this.LambdaCalculationPart _
                                                  , LAST_STEP_NAME, False
        Dim LetVarName As String
        this.VarNameVsCalculationAreaMap.Add Mapper.Create(LetVarName, CellReference), LetVarName
        FormatAndMoveWindowToLastCell this.StepStartCell, PutOnCell.Address
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword BuildLambdaSteps.ConstructSteps"
        GoTo CleanExit
    End If
    
    ' Iterating through the LetParts array
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(this.LetParts, 1) To UBound(this.LetParts, 1)
        
        Dim CalculationFunction As String
        UpdateLetVarNameAndCalculation LetVarName, CalculationFunction, CurrentRowIndex
        If Not IsMetadataLetVarName(LetVarName) Then
            PlaceStepFormula PutOnCell, HeadersRange, CellReference, CalculationFunction, LetVarName
    
            ' If we have same LET Step name like param for sanitization purpose or for default value then use updated cell ref.
            ' Check if the collection contains LetVarName
            If IsExistInCollection(this.VarNameVsCalculationAreaMap, LetVarName) Then
                ' If it does, remove it
                this.VarNameVsCalculationAreaMap.Remove LetVarName
            ElseIf IsExistInCollection(this.LambdaArgumentsName, LetVarName) Then
                ' If LetVarName is found in LambdaArgumentsName, add it to LambdaArgumentsRefCell
                this.LambdaArgumentsRefCell.Add CellReference, LetVarName
            End If
            ' Add LetVarName to VarNameVsCalculationAreaMap regardless of the above conditions
            this.VarNameVsCalculationAreaMap.Add Mapper.Create(LetVarName, CellReference), LetVarName
        End If
        
    Next CurrentRowIndex
    FormatAndMoveWindowToLastCell HeadersRange, CellReference
    
CleanExit:
    Application.Calculation = CalcMode
    Exit Sub
    

HandleError:
    Application.Calculation = CalcMode
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.ConstructSteps"
    
End Sub

Private Sub AuditLambdaParameters(ByRef PutOnCell As Range, ByRef HeadersRange As Range)
    
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.AuditLambdaParameters"
    Dim CurrentParam As Mapper
    Dim CellReference As String
    For Each CurrentParam In this.LambdaArgumentsName
        PlaceStepFormula PutOnCell, HeadersRange, CellReference, CurrentParam.Value, CurrentParam.Key
        this.LambdaArgumentsRefCell.Add CellReference, CurrentParam.Key
        this.VarNameVsCalculationAreaMap.Add Mapper.Create(CurrentParam.Key, CellReference), CurrentParam.Key
    Next CurrentParam
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.AuditLambdaParameters"
    
End Sub

Private Sub PlaceStepFormula(ByRef PutOnCell As Range _
                             , ByRef HeadersRange As Range _
                              , ByRef CellReference As String _
                               , ByVal CalculationFunction As String _
                                , ByVal VarName As String _
                                 , Optional UpdatePutOnCellRef As Boolean = True)
    
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.PlaceStepFormula"
    ' Setting the cell value
    PutOnCell.Offset(-1, 0).Value = Text.RemoveFromStartIfPresent(VarName, FormulaFormatConfig.LetVarPrefix, IGNORE_CASE)
    
    ' Setting HeadersRange
    If IsNothing(HeadersRange) Then
        Set HeadersRange = PutOnCell.Offset(-1, 0)
    Else
        Set HeadersRange = Union(HeadersRange, PutOnCell.Offset(-1, 0))
    End If
    
    ' Checking if the calculation function is an array constant
    If Text.IsStartsWith(CalculationFunction, LEFT_BRACE) Then
        Dim Temp As Variant
        Temp = modUtility.SplitArrayConstantTo2DArray(CalculationFunction)
        ' Filling the cells with the array values
        Dim RowCount As Long
        RowCount = UBound(Temp, 1) - LBound(Temp, 1) + 1
        Dim ColCount As Long
        ColCount = UBound(Temp, 2) - LBound(Temp, 2) + 1
        PutOnCell.Resize(RowCount, ColCount).Value = Temp
        CellReference = PutOnCell.Resize(RowCount, ColCount).Address(False, False)
        ' Checking if the calculation function is numeric
    ElseIf IsNumeric(CalculationFunction) Or Text.Trim(CalculationFunction) = vbNullString Then
        PutOnCell.Value = CalculationFunction
        CellReference = PutOnCell.Address(False, False)
        ' Checking if the calculation function is a formula
    Else
        Dim Formula As String
        Formula = UpdateForPreviousLetVar(CalculationFunction, this.VarNameVsCalculationAreaMap)
        ' Updating the formula of the cell
        Formula = ReplaceInvalidCharFromFormulaWithValid(UpdateForIsOmitted(Formula))
        Formula = QualifyImplicitColRef(Formula, this.InvocationCellTableName)
        '        Debug.Assert PutOnCell.Offset(-1, 0).Value <> "AllPaths"
        PutOnCell.Formula2 = Formula
        PutOnCell.Calculate
        ' Handling cases where the formula is a lambda
        If PutOnCell.Text = CALC_ERR_KEYWORD And IsLambdaFunction(EQUAL_SIGN & CalculationFunction) Then
            Formula = UpdateForPreviousLetVar(CalculationFunction, this.VarNameVsCalculationAreaMap, True)
            Formula = UpdateForIsOmitted(Formula)
            Dim Prefix As String
            Prefix = LETSTEP_PREFIX & IIf(Text.IsStartsWith(VarName, FormulaFormatConfig.LetVarPrefix), vbNullString, FormulaFormatConfig.LetVarPrefix)
            CreateLETStepNames PutOnCell, Prefix, VarName, Formula
            CellReference = Prefix & VarName
        ElseIf PutOnCell.HasSpill Then
            CellReference = PutOnCell.Address(False, False) & HASH_SIGN
        Else
            CellReference = PutOnCell.Address(False, False)
        End If
    End If
    
    If Not UpdatePutOnCellRef Then Exit Sub
    ' Adjust the PutOnCell according to CellReference and LetVarName
    If CellReference = Prefix & VarName Then
        ' If CellReference matches the concatenation of Prefix and LetVarName, move one column to the right
        Set PutOnCell = PutOnCell.Offset(0, 1)
    Else
        ' Otherwise, move as many columns to the right as there are in the range specified by CellReference
        Set PutOnCell = PutOnCell.Offset(0, LambdaInvocationCell.Worksheet.Range(CellReference).Columns.Count)
    End If
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.PlaceStepFormula"
    
End Sub

Private Function QualifyImplicitColRef(ByVal Formula As String, ByVal DefaultTableName As String) As String
    
    Dim Result As String
    Result = Formula
    If DefaultTableName <> vbNullString Then
        
        Dim DirectPrecedents As Variant
        DirectPrecedents = modDependencyLambdaResult.GetDirectPrecedents(Formula, this.LambdaInvocationCell.Worksheet)
            
        Dim CurrentPrecedent As Variant
        For Each CurrentPrecedent In DirectPrecedents
            If Text.IsStartsWith(CurrentPrecedent, "[@") And Text.IsEndsWith(CurrentPrecedent, "]") Then
                Result = ReplaceTokenWithNewToken(Result, CurrentPrecedent, DefaultTableName & CurrentPrecedent)
            End If
        Next CurrentPrecedent
        
    End If
    
    QualifyImplicitColRef = Result

End Function

Private Sub FormatAndMoveWindowToLastCell(ByVal HeadersRange As Range, CellReference As String)
    
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.FormatAndMoveWindowToLastCell"
    MarkAllInputCells
    HeadersRange.Font.Bold = True
    modUtility.AutoFitRange HeadersRange, MaximumColumnWidth:=50, MinimumColumnWidth:=8
    
    Dim EndCalcCell As Range
    Set EndCalcCell = LambdaInvocationCell.Worksheet.Range(CellReference).Cells(1)
    

    this.ResultAfterCommand = GetFormulaResult(EndCalcCell)
    
    ShowAlertIfBeforeAndAfterCommandResultIsDifferent this.ResultBeforeCommand _
                                                      , this.ResultAfterCommand _
                                                       , "Audit Formula Steps" _
                                                        , GetAFSCommandRangeRequiredFXAlertMessage(this.UsedRangeRequiredFXsName)
    
    MoveColumnToRightOfScreen EndCalcCell
    If Not Application.ScreenUpdating Then Application.ScreenUpdating = True
    If this.LambdaName <> vbNullString Then
        UpdateOrAddLambdaNameNote LambdaInvocationCell.Worksheet.Range(CellReference).Cells(1) _
                                  , this.LambdaName, LAMBDA_NAME_AUDIT_PREFIX
    End If
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.FormatAndMoveWindowToLastCell"
    
End Sub

'@Impure
Private Sub UpdateLetVarNameAndCalculation(ByRef LetVarName As String, ByRef CalculationFunction As String _
                                                                      , ByVal CurrentRowIndex As Long)
    
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.UpdateLetVarNameAndCalculation"
    Dim FirstColIndex As Long
    FirstColIndex = LBound(this.LetParts, 2)
    ' Cleaning the variable name
    LetVarName = modUtility.CleanVarName(CStr(this.LetParts(CurrentRowIndex, FirstColIndex)))
    CalculationFunction = this.LetParts(CurrentRowIndex _
                                        , FirstColIndex + modSharedConstant.LET_PARTS_VALUE_COL_INDEX - 1)
        
    ' Checking for the last step
    If CurrentRowIndex = UBound(this.LetParts, 1) Then
        ' Making changes to variable name and calculation function for the last step
        CalculationFunction = LetVarName
        If LetVarName = modUtility.CleanVarName(CStr(this.LetParts(CurrentRowIndex - 1, FirstColIndex))) Then
            LetVarName = vbNullString
        Else
            LetVarName = modUtility.FindUniqueNameByIncrementingNumber(this.VarNameVsCalculationAreaMap, LAST_STEP_NAME)
        End If
    End If
    
    ' Cleaning the calculation function
    CalculationFunction = Text.RemoveFromStartIfPresent(CalculationFunction, ONE_SPACE)
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.UpdateLetVarNameAndCalculation"
        
End Sub

Private Sub CreateLETStepNames(ByVal PutOnCell As Range _
                               , ByVal Prefix As String _
                                , ByVal LetVarName As String _
                                 , ByVal Formula As String)
                                                  
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.CreateLETStepNames"
    Dim AddToSheet As Worksheet
    Set AddToSheet = PutOnCell.Worksheet
    
    Dim QualifiedSheetName As String
    QualifiedSheetName = GetSheetRefForRangeReference(AddToSheet.Name, False)
    
    Dim LetStepFXName As String
    LetStepFXName = Prefix & LetVarName
    ' Check if a named range exists with the same name as LetStepFxName
    If Context.IsLocalScopedNamedRangeExists(AddToSheet, LetStepFXName) Then
        ' If it does exist, delete it
        AddToSheet.Names(QualifiedSheetName & LetStepFXName).Delete
    End If
                
    ' Add a new named range with LetStepFxName and the formula
    Formula = ConvertDependencisToFullyQualifiedRef(Formula, AddToSheet)
    AddToSheet.Names.Add LetStepFXName, Formula
                
    Dim LetStepRefName As String
    LetStepRefName = VBA.Replace(LetStepFXName, LETSTEP_UNDERSCORE_PREFIX, LETSTEPREF_UNDERSCORE_PREFIX)
    ' Check if a named range exists with the same name as LetStepRefName
    If Context.IsLocalScopedNamedRangeExists(AddToSheet, LetStepRefName) Then
        ' If it does exist, delete it
        AddToSheet.Names(QualifiedSheetName & LetStepRefName).Delete
    End If
    ' Assign the name of the cell as LetStepRefName
    AddToSheet.Names.Add LetStepRefName, EQUAL_SIGN & GetRangeRefWithSheetName(PutOnCell, True)
    UpdateOrAddLambdaNameNote PutOnCell, LetStepFXName, LAMBDA_NAME_NOTE_PREFIX
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.CreateLETStepNames"
    
End Sub

Private Sub MarkAllInputCells()
    
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.MarkAllInputCells"
    Dim CellReference As Variant
    Dim CurrentInputCells As Range
    ' Iterate over all references in the LambdaArgumentsRefCell collection
    For Each CellReference In this.LambdaArgumentsRefCell
        'If it is a lambda cell then use that.
        If Text.IsStartsWith(CStr(CellReference), LETSTEP_PREFIX) Then
            CellReference = Replace(CellReference, LETSTEP_PREFIX, LETSTEPREF_PREFIX, , 1)
        End If
        
        ' Get the range for the current cell reference
        Set CurrentInputCells = this.LambdaInvocationCell.Worksheet.Range(CStr(CellReference))
        ' Set the font color for the cells in the current range
        CurrentInputCells.Font.Color = FONT_COLOR_INDEX
        ' Mark the current range as input cells
        MarkAsInputCells CurrentInputCells
    Next CellReference
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.MarkAllInputCells"
    
End Sub

Private Function UpdateForPreviousLetVar(ByVal CalculationFunction As String _
                                         , ByRef PreviousVarNameAndCellMap As Collection _
                                          , Optional ByVal UseAbsoluteRangeRef As Boolean = False) As String
    
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.UpdateForPreviousLetVar"
    
    Dim UpdatedFormula As String
    UpdatedFormula = Text.PadIfNotPresent(CalculationFunction, EQUAL_SIGN)
    
    Dim CurrentMapper As Mapper
    For Each CurrentMapper In PreviousVarNameAndCellMap
        
        Dim CellRef As String
        CellRef = CurrentMapper.Value
        ' If it is LETStep_FUNCTION then value will hold the named range name. So no need to check for range.
        If Not IsStartsWithLetStepPrefix(CurrentMapper.Value) Then
            CellRef = ConvertFormulaToRef(CellRef, UseAbsoluteRangeRef _
                                                  , this.StepStartCell.Worksheet.Name)
        End If
        
        UpdatedFormula = modDependencyLambdaResult.ReplaceTokenWithNewToken(UpdatedFormula, CurrentMapper.Key, CellRef)
        
    Next CurrentMapper
    
    UpdateForPreviousLetVar = UpdatedFormula
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.UpdateForPreviousLetVar"
    
End Function

Private Function ConvertFormulaToRef(ByVal RangeRef As String _
                                     , ByVal UseAbsoluteRangeRef As Boolean _
                                      , ByVal SheetName As String) As String
    
    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.ConvertFormulaToRef"
    If RangeRef = vbNullString Then Exit Function
    
    ' If not using absolute references, simply return the original reference
    If Not UseAbsoluteRangeRef Then
        ConvertFormulaToRef = RangeRef
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword BuildLambdaSteps.ConvertFormulaToRef"
        Exit Function
    End If
    
    ' If using absolute references, convert the reference to the absolute format _
    ' , taking into account whether it is a hash range or not
    
    If Text.IsEndsWith(RangeRef, HASH_SIGN) Then
        ConvertFormulaToRef = GetSheetRefForRangeReference(SheetName, False) _
                              & RangeResolver.GetRange(RangeRef).Cells(1).Address & HASH_SIGN
    Else
        ConvertFormulaToRef = GetSheetRefForRangeReference(SheetName, False) _
                              & RangeResolver.GetRange(RangeRef).Address
    End If
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.ConvertFormulaToRef"
    
End Function

Private Sub Class_Initialize()

    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.Class_Initialize"
    ' Initialize collections to track LET variable names, lambda argument names, and reference cells
    Set this.VarNameVsCalculationAreaMap = New Collection
    Set this.LambdaArgumentsName = New Collection
    Set this.LambdaArgumentsRefCell = New Collection
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.Class_Initialize"

End Sub

Private Sub Class_Terminate()

    Logger.Log TRACE_LOG, "Enter BuildLambdaSteps.Class_Terminate"
    ' Clean up by setting the collections to Nothing when the class instance is terminated
    Set this.VarNameVsCalculationAreaMap = Nothing
    Set this.LambdaArgumentsName = Nothing
    Set this.LambdaArgumentsRefCell = Nothing
    Logger.Log TRACE_LOG, "Exit BuildLambdaSteps.Class_Terminate"

End Sub




