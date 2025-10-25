VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LETORLAMBDAGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TLETORLAMBDAGenerator
    LambdaDef As String
    LambdaInvocation As String
    LetFormula As String
    IsLetNeededInLambda As Boolean
    PrecedencyExtractor As FormulaPrecedencyExtractor
    IsLetNeededInLetFormula As Boolean
    RangeReqFXList As Collection
End Type

Private this As TLETORLAMBDAGenerator

Public Property Get RangeReqFXList() As Collection
    Set RangeReqFXList = this.RangeReqFXList
End Property

Public Property Get IsLetNeededInLetFormula() As Boolean
    IsLetNeededInLetFormula = this.IsLetNeededInLetFormula
End Property

Public Property Get IsLetNeededInLambda() As Boolean
    IsLetNeededInLambda = this.IsLetNeededInLambda
End Property

Public Property Get LambdaDef() As String
    LambdaDef = this.LambdaDef
End Property

Public Property Get LambdaInvocation() As String
    LambdaInvocation = this.LambdaInvocation
End Property

Public Property Get LetFormula() As String
    LetFormula = this.LetFormula
End Property

Public Sub CreateLetOrLambdaStatement(ByVal DependencyObjects As Collection _
                                      , ByVal FormulaStartCell As Range _
                                       , ByVal TypeOfOperation As OperationType _
                                        , ByVal PrecedencyExtractor As FormulaPrecedencyExtractor)
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.CreateLetAndLambdaStatements"
    Set this.PrecedencyExtractor = PrecedencyExtractor
    
    If TypeOfOperation = LET_STATEMENT_GENERATION Then
        CreateLETStatement DependencyObjects, FormulaStartCell
       Set this.RangeReqFXList = GetUsedRangeRequiredFXsName(this.LetFormula)
    ElseIf TypeOfOperation = LAMBDA_STATEMENT_GENERATION Then
        CreateLambdaStatement DependencyObjects, FormulaStartCell
        Set this.RangeReqFXList = GetUsedRangeRequiredFXsName(this.LambdaDef)
    End If
    
    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.CreateLetAndLambdaStatements"
    
End Sub

Private Sub CreateLambdaStatement(ByVal DependencyObjects As Collection _
                                  , ByVal FormulaStartCell As Range)
    
    ' Create Let statements without input cells and get the LetFormula.
    ' Check if Let statement is needed for the Lambda function.
    this.IsLetNeededInLambda = IsLETNeeded(DependencyObjects)
    Dim LetFormula As String
    
    If Not this.IsLetNeededInLambda Then
        ' If Let statement is not needed, create the Let statement without input cells.
        LetFormula = CreateLetStatementsWithoutInputCellsIfLetNotNeeded(DependencyObjects, True)
    Else
        LetFormula = CreateLetStatementsWithoutInputCells(DependencyObjects, FormulaStartCell, True)
    End If
    
    Logger.Log DEBUG_LOG, LetFormula

    ' Create the FullLambda statement based on the LetFormula and DependencyObjects.
    Dim FullLambda As String
    this.LambdaDef = CreateFullLambda(LetFormula, DependencyObjects)
        
End Sub

Private Function CreateLetStatementsWithoutInputCells(ByVal DependencyObjects As Collection _
                                                      , ByVal FormulaStartCell As Range _
                                                       , ByVal IsConvertToAbsoluteRangeRef As Boolean) As String
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.CreateLetStatementsWithoutInputCells"
    ' Creates the Let statement for the Lambda function excluding input cells.
    
    Dim LetFormula As String
    Dim CurrentDependencyInfo As DependencyInfo
    LetFormula = LET_AND_OPEN_PAREN

    For Each CurrentDependencyInfo In DependencyObjects
        LetFormula = LetFormula & _
                     GetLetStepFormulaForCurrDependency(DependencyObjects, FormulaStartCell, CurrentDependencyInfo, IsConvertToAbsoluteRangeRef)
    Next CurrentDependencyInfo

    Set CurrentDependencyInfo = DependencyObjects.Item(DependencyObjects.Count)
    
    ' Check if the FormulaStartCell has a variable name or marked as not Let statement by the user.
    If this.PrecedencyExtractor.VarNameOfFormulaStartCell = vbNullString _
       Or CurrentDependencyInfo.IsMarkAsNotLetStatementByUser _
       Or Not FormulaFormatConfig.IncludeResultStep Then
        Dim StructuralFormulaText As String
        StructuralFormulaText = ReplaceDependencyRefWithStepName(CurrentDependencyInfo.FormulaText, CurrentDependencyInfo, DependencyObjects)
        ' Append the remaining part of the Let statement or the StructuralFormulaText if LetFormulaWithoutInputCells is empty.
        If LetFormula <> LET_AND_OPEN_PAREN Then
            LetFormula = LetFormula & NEW_LINE & THREE_SPACE & _
                         StructuralFormulaText _
                         & NEW_LINE & FIRST_PARENTHESIS_CLOSE
        Else
            LetFormula = StructuralFormulaText
        End If
        
    Else
        ' Append the closing parenthesis for the Let statement.
        LetFormula = LetFormula & NEW_LINE & THREE_SPACE & _
                     CurrentDependencyInfo.ValidVarName & NEW_LINE & FIRST_PARENTHESIS_CLOSE
    End If
    
    Logger.Log TRACE_LOG, vbNewLine
    Logger.Log DEBUG_LOG, "Let Statement Part : " & LetFormula

    ' Call the function to update the blank range label and name in the formula with the Let variable name.
    LetFormula = UpdateLetFormulaForNotLetStatementDepWithProperValue(LetFormula, DependencyObjects, IsConvertToAbsoluteRangeRef)
    CreateLetStatementsWithoutInputCells = LetFormula

    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.CreateLetStatementsWithoutInputCells"
    
End Function

Private Function GetLetStepFormulaForCurrDependency(ByVal DependencyObjects As Collection _
                                                    , ByVal FormulaStartCell As Range _
                                                     , ByVal CurrentDependencyInfo As DependencyInfo _
                                                      , ByVal IsMakeAbsoluteRef As Boolean) As String
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.GetLetStepFormulaForCurrDependency"
    ' Check if the current dependency info is not an input cell and doesn't have a user mark to exclude from Let statement.
    Dim IsValidForLetStep As Boolean
    
    Dim LastDependency As DependencyInfo
    Set LastDependency = DependencyObjects.Item(DependencyObjects.Count)
    
    If CurrentDependencyInfo.RangeReference = LastDependency.RangeReference And Not FormulaFormatConfig.IncludeResultStep Then
        IsValidForLetStep = False
    Else
        IsValidForLetStep = (Not CurrentDependencyInfo.IsLabelAsInputCell _
                             And Not CurrentDependencyInfo.IsMarkAsNotLetStatementByUser)
    End If
    
    If Not IsValidForLetStep Then Exit Function
    
    Dim Result As String
    
    With CurrentDependencyInfo
    
        Dim CorrectedNameInFormula As String
        CorrectedNameInFormula = .NameInFormula
                
        Dim PossibleLetVarValue As String
        If .IsUserMarkAsValue Then
            PossibleLetVarValue = RemoveStartingSingleQuoteAndEqualSign(.FormulaText)
        ElseIf .IsReferByNamedRange Then
            If IsStartsWithLetStepPrefix(.NameInFormula) Then
                PossibleLetVarValue = RemoveStartingSingleQuoteAndEqualSign(.FormulaText)
            Else
                PossibleLetVarValue = .NameInFormula
            End If
        ElseIf .IsTableRef Then
            PossibleLetVarValue = .NameInFormula
        ElseIf .HasFormula And (.HasAnyDependency Or Not .IsDemotedFromParameterCellToLetStep) Then
            PossibleLetVarValue = modUtility.RemoveStartingSingleQuoteAndEqualSign(.FormulaText)
        Else
            If IsMakeAbsoluteRef Then
                PossibleLetVarValue = MakeAbsoluteReference(.RangeReference, FormulaStartCell)
            Else
                PossibleLetVarValue = .NameInFormula
            End If
            CorrectedNameInFormula = PossibleLetVarValue
        End If
        PossibleLetVarValue = FindLetVarValue(.RangeReference, PossibleLetVarValue _
                                                              , CorrectedNameInFormula _
                                                               , .IsUserMarkAsValue)
                        
        If PossibleLetVarValue = vbNullString Then
            ' As excel blank cell treat as zero.
            PossibleLetVarValue = 0
        Else
            ' Replace Previous Cell Refs
            PossibleLetVarValue = ReplaceDependencyRefWithStepName(PossibleLetVarValue _
                                                                   , CurrentDependencyInfo, DependencyObjects)
        End If
        
        ' Append the Let statement part for the current dependency info to LetFormulaWithoutInputCells.
        Result = NEW_LINE & THREE_SPACE _
                 & .ValidVarName & LIST_SEPARATOR & ONE_SPACE _
                 & PossibleLetVarValue & LIST_SEPARATOR
    End With
    
    GetLetStepFormulaForCurrDependency = Result
    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.GetLetStepFormulaForCurrDependency"
    
End Function

Private Function UpdateLetFormulaForNotLetStatementDepWithProperValue(ByVal LetFormula As String _
                                                                      , ByVal DependencyObjects As Collection _
                                                                       , ByVal IsConvertToAbsoluteRangeRef As Boolean) As String
                                                              
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.UpdateLetFormulaForNotLetStatementDepWithProperValue"
    Dim CurrentDependencyInfo As DependencyInfo
    Dim UpdatedFormula As String
    UpdatedFormula = LetFormula
    ' Reverse loop is necessary to handle cases where a cell reference is dependent on another below it.
    Dim Counter As Long
    For Counter = DependencyObjects.Count To 1 Step -1
    
        Set CurrentDependencyInfo = DependencyObjects.Item(Counter)
        With CurrentDependencyInfo
            If Not .IsLabelAsInputCell And .IsMarkAsNotLetStatementByUser Then
                UpdatedFormula = UpdateLetFormulaForCellRefWithProperFormulaText(UpdatedFormula, CurrentDependencyInfo, DependencyObjects, IsConvertToAbsoluteRangeRef)
                DoEvents
            End If
        End With
        
    Next Counter

    ' Concatenate the modified formula back into a single string.
    UpdateLetFormulaForNotLetStatementDepWithProperValue = UpdatedFormula
    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.UpdateLetFormulaForNotLetStatementDepWithProperValue"

End Function

Private Function UpdateLetFormulaForCellRefWithProperFormulaText(ByVal LetFormula As String _
                                                                 , ByVal CurrentDependencyInfo As DependencyInfo _
                                                                  , ByVal DependencyObjects As Collection _
                                                                   , ByVal IsConvertToAbsoluteRangeRef As Boolean) As String

    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.UpdateLetFormulaForCellRefWithProperFormulaText"
    Dim FormulaText As String
    With CurrentDependencyInfo

        Dim ReplacedRefFormula As String
        ' Formulatext can be blank for multi cell dependency, table column ref etc.
        If .FormulaText <> vbNullString Then
            ReplacedRefFormula = ReplaceDependencyRefWithStepName(.FormulaText, CurrentDependencyInfo, DependencyObjects)
        End If

        ' Determine the formula text based on the dependency's attributes.
        If .IsUserMarkAsValue Then
            FormulaText = RemoveStartingEqualSign(ReplacedRefFormula)

        ElseIf .IsReferByNamedRange Then
            If Text.IsStartsWith(.NameInFormula, LETSTEP_PREFIX) Then
                FormulaText = ReplacedRefFormula
            Else
                FormulaText = .NameInFormula
            End If
        ElseIf .IsInsideNamedRangeOrTable Then
            FormulaText = .NameInFormula
        ElseIf .HasFormula And (.HasAnyDependency Or Not .IsDemotedFromParameterCellToLetStep Or .IsExpandByUser) Then
            FormulaText = modUtility.RemoveStartingSingleQuoteAndEqualSign(ReplacedRefFormula)
            FormulaText = EncloseWithParenForMultiTerm(FormulaText)
        ElseIf IsConvertToAbsoluteRangeRef Then
            FormulaText = GetAbsoluteRefForNameInFormula(.RangeReference, .NameInFormula)
        Else
            FormulaText = .NameInFormula
        End If
        
        Dim Result As String
        If FormulaText <> .NameInFormula Then
            ' ref with blank. We need to use zero(excel default for blank cell) for that.
            If FormulaText = vbNullString Then FormulaText = 0
            Result = ReplaceCellRefWithStepName(LetFormula, FormulaText, .NameInFormula, .SheetName)
        Else
            Result = LetFormula
        End If
        
        UpdateLetFormulaForCellRefWithProperFormulaText = Result
        
    End With
    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.UpdateLetFormulaForCellRefWithProperFormulaText"

End Function

Private Function CreateFullLambda(ByVal LetFormula As String _
                                  , ByVal DependencyObjects As Collection) As String
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.CreateFullLambda"
    ' Creates the full Lambda function based on the Let formula and dependency objects collection.

    ' Check if there is only one dependency marked as value.
    If IsOneDependencyAndMarkAsValue(DependencyObjects) Then
        CreateFullLambda = EQUAL_SIGN & LAMBDA_AND_OPEN_PAREN _
                           & Text.RemoveFromStartIfPresent(LetFormula, EQUAL_SIGN) _
                           & FIRST_PARENTHESIS_CLOSE
        this.LambdaInvocation = "()"
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword LETORLAMBDAGenerator.CreateFullLambda"
        Exit Function
    End If

    ' Update Let formula if necessary.
    If Right$(LetFormula, 1) = FIRST_PARENTHESIS_CLOSE And Text.IsStartsWith(LetFormula _
                                                                             , EQUAL_LET_FIRST_PAREN) _
                                                                             And this.IsLetNeededInLambda Then
        LetFormula = Left$(LetFormula, Len(LetFormula) - 1) & THREE_SPACE & FIRST_PARENTHESIS_CLOSE
    End If
    
    Dim LambdaArgsPartAndInvocationPart As Variant
    LambdaArgsPartAndInvocationPart = CreateLambdaArgumentAndInvocationPart(DependencyObjects, LetFormula)
    
    ' Create the Lambda formula argument.
    Dim LambdaFormula As String
    LambdaFormula = LambdaArgsPartAndInvocationPart(LBound(LambdaArgsPartAndInvocationPart))
    this.LambdaInvocation = LambdaArgsPartAndInvocationPart(LBound(LambdaArgsPartAndInvocationPart) + 1)

    ' Update the Let formula with the Lambda argument names.
    Dim FinalLetWithLamdaArgumentName As String
    FinalLetWithLamdaArgumentName = UpdateFormulaForInputCells(LetFormula _
                                                               , DependencyObjects, True)

    ' Combine the Lambda formula and Let formula to create the full Lambda.
    Dim FullLambda As String
    FullLambda = EQUAL_SIGN & LambdaFormula
    
    If LambdaFormula <> LAMBDA_AND_OPEN_PAREN Then
        FullLambda = FullLambda & LIST_SEPARATOR
    End If
    
    CreateFullLambda = FullLambda & modUtility.RemoveStartingEqualSign(FinalLetWithLamdaArgumentName) _
                       & IIf(this.IsLetNeededInLambda, vbNewLine, vbNullString) _
                       & FIRST_PARENTHESIS_CLOSE

    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.CreateFullLambda"
    
End Function

Private Sub CreateLETStatement(ByVal DependencyObjects As Collection _
                               , ByVal FormulaStartCell As Range)
    
    this.PrecedencyExtractor.AddPrefixOnInputCells DependencyObjects
    
    this.IsLetNeededInLetFormula = IsLetNeededInLetStatement(DependencyObjects)
    
    Dim LetFormula As String
    
    If Not this.IsLetNeededInLetFormula Then
        ' If Let statement is not needed, create the Let statement without input cells.
        LetFormula = CreateLetStatementsWithoutInputCellsIfLetNotNeeded(DependencyObjects, False)
    Else
        LetFormula = CreateLetStatementsWithoutInputCells(DependencyObjects, FormulaStartCell, False)
    End If
    
    LetFormula = Text.PadIfNotPresent(LetFormula, EQUAL_SIGN, FROM_START)
    LetFormula = CreateFullLet(LetFormula, DependencyObjects)
    this.LetFormula = FormatFormula(LetFormula)
    
End Sub

Private Function IsLetNeededInLetStatement(ByVal DependencyObjects As Collection) As Boolean
    
    If DependencyObjects.Count = 1 Then
        IsLetNeededInLetStatement = False
        Exit Function
    End If
    
    Dim CurrentDependencyInfo As DependencyInfo
    Dim Counter As Long

    For Each CurrentDependencyInfo In DependencyObjects
        
        With CurrentDependencyInfo
            If Not .IsMarkAsNotLetStatementByUser Then
                Counter = Counter + 1
            End If
        End With
        
        If Counter > 1 Then Exit For
        
    Next CurrentDependencyInfo
    
    Dim Result As Boolean
    
    If Counter = 1 Then
        ' If only one let statement needed and if that is for the last step then don't add let
        Set CurrentDependencyInfo = DependencyObjects.Item(DependencyObjects.Count)
        Result = (CurrentDependencyInfo.IsMarkAsNotLetStatementByUser _
                  Or Not FormulaFormatConfig.IncludeResultStep)
    Else
        Result = True
    End If
    
    IsLetNeededInLetStatement = Result
    
End Function

Private Function CreateFullLet(ByVal LetFormula As String, ByVal DependencyObjects As Collection) As String
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.CreateFullLet"
    ' Creates the full Let statement based on the Let formula and dependency objects collection.
    Dim LetPartForInputCells As String
    LetPartForInputCells = CreateLetPartForInputCells(DependencyObjects)
    LetFormula = UpdateFormulaForInputCells(LetFormula, DependencyObjects, False)
    If LetPartForInputCells <> vbNullString And Not Text.IsStartsWith(LetFormula, EQUAL_LET_FIRST_PAREN) Then
        LetFormula = Replace(LetFormula, EQUAL_SIGN, EQUAL_LET_FIRST_PAREN _
                                                    & NEW_LINE & THREE_SPACE, 1, 1) _
                                                    & NEW_LINE & FIRST_PARENTHESIS_CLOSE
    End If
    
    Dim FullLet As String
    If LetPartForInputCells = vbNullString Then
        FullLet = Text.PadIfNotPresent(LetFormula, EQUAL_SIGN, FROM_START)
    Else
        FullLet = Replace(LetFormula, EQUAL_LET_FIRST_PAREN _
                                     , EQUAL_LET_FIRST_PAREN & NEW_LINE & LetPartForInputCells _
                                      , 1, 1, vbTextCompare)
        FullLet = modUtility.RemoveEndingText(FullLet, FIRST_PARENTHESIS_CLOSE)
        FullLet = modUtility.RemoveEndingText(FullLet, THREE_SPACE)
    
        If Text.IsStartsWith(FullLet, EQUAL_LET_FIRST_PAREN) Then
            FullLet = FullLet & FIRST_PARENTHESIS_CLOSE
        End If
    
    End If
    
    CreateFullLet = FullLet
    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.CreateFullLet"
    
End Function

Private Function CreateLetPartForInputCells(ByVal DependencyObjects As Collection) As String
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.CreateLetPartForInputCells"
    Dim LetArgument As String
    LetArgument = vbNullString
    Dim CurrentDependencyInfo As DependencyInfo
    Dim Prefix As String
    Prefix = FormulaFormatConfig.LetVarPrefix
    For Each CurrentDependencyInfo In DependencyObjects
        With CurrentDependencyInfo
            If .IsLabelAsInputCell And Not .IsMarkAsNotLetStatementByUser Then
                LetArgument = LetArgument & THREE_SPACE _
                              & .ValidVarName & LIST_SEPARATOR _
                              & ONE_SPACE & .StepOrArgumentCalcFunction(this.PrecedencyExtractor.FormulaStartCell) _
                              & LIST_SEPARATOR & vbNewLine
            End If
        End With
    Next CurrentDependencyInfo
    LetArgument = modUtility.RemoveEndingText(LetArgument, vbNewLine)
    CreateLetPartForInputCells = LetArgument
    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.CreateLetPartForInputCells"
    
End Function

Private Function IsLETNeeded(ByVal DependencyObjects As Collection) As Boolean
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.IsLETNeeded"
    ' Checks if a Let statement is needed for the Lambda function based on the number of non-input dependencies.

    Dim NonInputDependencyCount As Long
    NonInputDependencyCount = GetNumberOfNonInputDependency(DependencyObjects)

    Dim NonInputLetDependencyCount As Long
    NonInputLetDependencyCount = GetNumberOfNonInputLetStepDependency(DependencyObjects)

    ' A Let statement is needed if the number of non-input dependencies is not equal to 1 or the number of non-input Let step dependencies is not equal to 0.
    IsLETNeeded = Not (NonInputDependencyCount = 1 Or NonInputLetDependencyCount = 0)

    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.IsLETNeeded"
    
End Function

Private Function CreateLetStatementsWithoutInputCellsIfLetNotNeeded(ByVal DependencyObjects As Collection _
                                                                    , ByVal IsConvertToAbsoluteRangeRef As Boolean) As String
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.CreateLetStatementsWithoutInputCellsIfLetNotNeeded"
    ' Creates the Let statement for the Lambda function without input cells when Let statement is not needed.
    
    Dim CurrentDependencyInfo As DependencyInfo
    Set CurrentDependencyInfo = DependencyObjects.Item(DependencyObjects.Count)

    ' Get the formula from the last dependency info in the collection.
    Dim Formula As String
    Formula = modUtility.RemoveStartingSingleQuoteAndEqualSign(CurrentDependencyInfo.FormulaText)

    ' Call the function to update the blank range label and name in the formula with the Let variable name.
    Formula = UpdateLetFormulaForNotLetStatementDepWithProperValue(Formula, DependencyObjects, IsConvertToAbsoluteRangeRef)
    CreateLetStatementsWithoutInputCellsIfLetNotNeeded = Formula

    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.CreateLetStatementsWithoutInputCellsIfLetNotNeeded"
    
End Function

Private Function MakeAbsoluteReference(ByVal RangeAddress As String _
                                       , ByVal HelperCell As Range) As String
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.MakeAbsoluteReference"
    ' Convert the given range address to an absolute reference.
    
    Dim Result As String
    If Is3DReference(RangeAddress) Then
        Result = Text.BeforeDelimiter(RangeAddress, SHEET_NAME_SEPARATOR, , FROM_end) & _
                 SHEET_NAME_SEPARATOR & _
                 GetRangeReference(Range(Text.AfterDelimiter(RangeAddress, SHEET_NAME_SEPARATOR, , FROM_end)), True)
    Else
        Dim CurrentRange As Range
        Set CurrentRange = RangeResolver.GetRangeForDependency(RangeAddress, HelperCell)
    
        ' Check if the parent worksheet of the current range is the same as the parent worksheet of the helper cell.
        If CurrentRange.Worksheet.Name = HelperCell.Worksheet.Name Then
            ' If they are the same, return the range reference with absolute references.
            Result = GetRangeReference(CurrentRange, True)
        Else
            ' If they are different, return the range reference with absolute references and sheet name.
            Result = GetRangeRefWithSheetName(CurrentRange, , True)
        End If
    End If
    
    MakeAbsoluteReference = Result
    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.MakeAbsoluteReference"
    
End Function

Private Function FindLetVarValue(ByVal RangeReference As String _
                                 , ByVal GivenFormulaText As String _
                                  , ByVal GivenNameInFormula As String _
                                   , ByVal IsMarkAsValue As Boolean) As String
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.FindLetVarValue"
    ' Finds the value for the Let variable in the formula.
    Dim Value As String

    ' Check if the range is multi-cell and not dynamic.
    If IsMultiCellNotDynamic(RangeReference) And Not IsMarkAsValue Then
        ' If multi-cell and not dynamic, use the name in the formula as the value.
        Value = GivenNameInFormula
    Else
        ' Otherwise, use the given formula text as the value.
        Value = GivenFormulaText
    End If

    FindLetVarValue = Value
    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.FindLetVarValue"
    
End Function

Private Function ReplaceDependencyRefWithStepName(ByVal Formula As String _
                                                  , ByVal ForDependency As DependencyInfo _
                                                   , ByVal DependencyObjects As Collection) As String
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.ReplaceDependencyRefWithStepName"
    Dim UpdatedFormula As String
    UpdatedFormula = Text.PadIfNotPresent(Formula, EQUAL_SIGN, FROM_START)
    Dim CurrentDependency As DependencyInfo
    Dim ValidVarName As String
    For Each CurrentDependency In DependencyObjects
        
        With CurrentDependency
            
            ValidVarName = .ValidVarName
            
            If .RangeReference = ForDependency.RangeReference Then
                Exit For
            End If
            
            Dim SheetPrefix As String
            SheetPrefix = GetSheetRefForRangeReference(.SheetName)
            If IsStartsWithLetStepPrefix(.NameInFormula) Then
                UpdatedFormula = ReplaceTokenWithNewToken(UpdatedFormula, SheetPrefix & .NameInFormula, ValidVarName)
                UpdatedFormula = ReplaceTokenWithNewToken(UpdatedFormula, .NameInFormula, ValidVarName)
            ElseIf Text.IsStartsWith(.NameInFormula, SheetPrefix & LETSTEP_UNDERSCORE_PREFIX) Then
                UpdatedFormula = ReplaceTokenWithNewToken(UpdatedFormula, .NameInFormula, ValidVarName)
            ElseIf Not .IsMarkAsNotLetStatementByUser Then
                UpdatedFormula = ReplaceCellRefWithStepName(UpdatedFormula, ValidVarName, .NameInFormula, .SheetName)
            End If
            
        End With
        
    Next CurrentDependency
    
    ReplaceDependencyRefWithStepName = Text.RemoveFromStartIfPresent(UpdatedFormula, EQUAL_SIGN)
    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.ReplaceDependencyRefWithStepName"

End Function

Private Function IsOneDependencyAndMarkAsValue(ByVal DependencyObjects As Collection) As Boolean
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.IsOneDependencyAndMarkAsValue"
    ' Checks if there is only one dependency in the collection and it is marked as value.

    Dim CurrentDependencyInfo As DependencyInfo
    If DependencyObjects.Count = 1 Then
        Set CurrentDependencyInfo = DependencyObjects.Item(1)
        IsOneDependencyAndMarkAsValue = CurrentDependencyInfo.IsUserMarkAsValue
    End If
    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.IsOneDependencyAndMarkAsValue"
    
End Function

Private Function CreateLambdaArgumentAndInvocationPart(ByVal DependencyObjects As Collection _
                                                       , ByVal LetFormula As String) As Variant
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.CreateLambdaArgumentAndInvocationPart"
    ' Creates the Lambda argument for the Lambda formula.

    Dim ArgumentPart As String
    ArgumentPart = LAMBDA_AND_OPEN_PAREN
    Dim InvocationPart As String
    InvocationPart = FIRST_PARENTHESIS_OPEN

    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In DependencyObjects
        With CurrentDependencyInfo
            If .IsLabelAsInputCell Then
                ' Check if the argument is optional based on LetFormula.
                If modUtility.IsOptionalArgument(LetFormula, .ValidVarName) Or .IsOptional Then
                    ArgumentPart = ArgumentPart & LEFT_BRACKET & .ValidVarName _
                                   & RIGHT_BRACKET & LIST_SEPARATOR
                Else
                    ArgumentPart = ArgumentPart & .ValidVarName & LIST_SEPARATOR
                End If
                
                InvocationPart = InvocationPart _
                                 & .StepOrArgumentCalcFunction(this.PrecedencyExtractor.FormulaStartCell) _
                                 & LIST_SEPARATOR
                
            End If
        End With
    Next CurrentDependencyInfo

    ArgumentPart = modUtility.RemoveEndingText(ArgumentPart, LIST_SEPARATOR)
    InvocationPart = modUtility.RemoveEndingText(InvocationPart, LIST_SEPARATOR)
    InvocationPart = InvocationPart & FIRST_PARENTHESIS_CLOSE

    CreateLambdaArgumentAndInvocationPart = Array(ArgumentPart, InvocationPart)

    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.CreateLambdaArgumentAndInvocationPart"
    
End Function

Private Function UpdateFormulaForInputCells(ByVal LetFormula As String _
                                            , ByVal DependencyObjects As Collection _
                                             , ByVal IsForLambda As Boolean) As String
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.UpdateFormulaForInputCells"
    ' Update let formula with input cell valid var name.
    
    Dim LetVarPrefix As String
    LetVarPrefix = FormulaFormatConfig.LetVarPrefix
    
    Dim UpdatedFormula As String
    UpdatedFormula = LetFormula
    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In DependencyObjects
        With CurrentDependencyInfo
            If .IsLabelAsInputCell And Not .IsMarkAsNotLetStatementByUser Then
                
                Dim VarName As String
                Dim CellRef As String
                If IsForLambda Then
                    VarName = .ValidVarName
                    CellRef = .NameInFormula
                    UpdatedFormula = ReplaceCellRefWithStepName(UpdatedFormula, VarName, CellRef, .SheetName)
                Else
                    ' Change the naming convention to use Let Steps.
                    VarName = LetVarPrefix & MakeValidLetVarName(.ValidVarName, GetNamingConv(False))
                    CellRef = .NameInFormula
                    ' CellRef = .ValidVarName
                    UpdatedFormula = ReplaceTokenWithNewToken(UpdatedFormula, CellRef, VarName)
                End If
                
            End If
        End With
    Next CurrentDependencyInfo

    UpdateFormulaForInputCells = UpdatedFormula
    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.UpdateFormulaForInputCells"
    
End Function

Private Function IsMultiCellNotDynamic(ByVal RangeReference As String) As Boolean
    
    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.IsMultiCellNotDynamic"
    ' Checks if the given range reference represents a multi-cell range that is not dynamic.

    ' If the range reference ends with the DYNAMIC_CELL_REFERENCE_SIGN,
    ' it indicates a dynamic range, so return False.
    Dim Result As Boolean
    If IsSpilledRangeRef(RangeReference) Then
        Result = False
    Else
        ' Use the modUtility.GetRange function to get the actual Range object from the given range reference.
        ' Check if the cell count of the range is greater than 1 to determine if it's a multi-cell range.
        Dim ResolvedRange As Range
        Set ResolvedRange = RangeResolver.GetRange(RangeReference _
                                                   , this.PrecedencyExtractor.FormulaInWorkbook _
                                                    , this.PrecedencyExtractor.FormulaStartCellInSheet)
        If IsNothing(ResolvedRange) Then
            Result = False
        Else
            Result = (ResolvedRange.Cells.Count > 1)
        End If
        
    End If
    
    IsMultiCellNotDynamic = Result
    
    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.IsMultiCellNotDynamic"
    
End Function

Private Function GetAbsoluteRefForNameInFormula(ByVal RangeRef As String _
                                                , ByVal NameInFormula As String) As String

    Logger.Log TRACE_LOG, "Enter LETORLAMBDAGenerator.GetAbsoluteRefForNameInFormula"
    Dim Ref As String
    If Is3DReference(NameInFormula) Then
       Ref = MakeAbsoluteReference(NameInFormula, Nothing)
    ElseIf Text.Contains(RangeRef, NameInFormula) Then
        Dim ResolvedRange As Range
        Set ResolvedRange = RangeResolver.GetRange(RangeRef)
        If Text.IsEndsWith(NameInFormula, HASH_SIGN) Then
            Ref = ResolvedRange.Cells(1).Address & HASH_SIGN
        Else
            Ref = ResolvedRange.Address
        End If
    Else
        Ref = NameInFormula
    End If
    
    If Text.Contains(NameInFormula, SHEET_NAME_SEPARATOR) _
       And Not Text.Contains(Ref, SHEET_NAME_SEPARATOR) Then
        Ref = Text.BeforeDelimiter(NameInFormula, SHEET_NAME_SEPARATOR, , FROM_end) _
              & SHEET_NAME_SEPARATOR & Ref
    End If
    
    GetAbsoluteRefForNameInFormula = Ref
    Logger.Log TRACE_LOG, "Exit LETORLAMBDAGenerator.GetAbsoluteRefForNameInFormula"
    
End Function


