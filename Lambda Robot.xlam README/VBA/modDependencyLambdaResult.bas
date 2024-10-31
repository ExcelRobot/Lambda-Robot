Attribute VB_Name = "modDependencyLambdaResult"
'@Exposed
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed
'@Folder "COM.Wrapper"

Option Explicit
Option Private Module

#Const DEVELOPMENT_MODE = True
 
Public Enum DependencyFunctions
    LET_PARTS = 1
    LAMBDA_PARTS = 2
End Enum

Private Sub Test()
    
    Dim TestFormula As String
    TestFormula = ActiveCell.Formula2
    Dim V As Variant
    V = GetDependencyFunctionResult(ActiveCell.Formula2, LAMBDA_PARTS, False)
    V = GetDependencyFunctionResult("=LAMBDA(a,a*2)(5)", LET_PARTS, True)
    
End Sub

' Find all the dependency workbook is necessary for removing Lambda which is in the name manager
Public Function GetDirectPrecedents(ByVal Formula As String, ByVal FormulaInSheet As Worksheet) As Variant
    
    Dim Dependencies As Collection
    Set Dependencies = GetDirectPrecedentsFromExpr(Formula, FormulaInSheet)
    
    Dim Result As Variant
    If Dependencies.Count = 0 Then
        ReDim Result(1 To 1, 1 To 1) As String
        Result(1, 1) = vbNullString
        GetDirectPrecedents = Result
        Exit Function
    End If
    
    Dim Lambdas As Collection
    Set Lambdas = FindLambdas(FormulaInSheet.Parent)
    Dim ValidDependencies As Collection
    Set ValidDependencies = New Collection
    Dim CurrentDependency As Variant
    Dim QualifiedSheetName As String
    QualifiedSheetName = GetSheetRefForRangeReference(FormulaInSheet.name, False)
    
    For Each CurrentDependency In Dependencies
        ' Check if local or global lambdas present or not
        If Not (IsExistInCollection(Lambdas, CStr(CurrentDependency)) _
                Or IsExistInCollection(Lambdas, QualifiedSheetName & CurrentDependency)) Then
            ValidDependencies.Add CurrentDependency
        End If
    Next CurrentDependency
    
    If ValidDependencies.Count = 0 Then
        ReDim Result(1 To 1, 1 To 1) As String
        Result(1, 1) = vbNullString
        GetDirectPrecedents = Result
    Else
        GetDirectPrecedents = CollectionToArray(ValidDependencies)
    End If
    
    Set ValidDependencies = Nothing
    
End Function

Private Function GetDirectPrecedentsFromExpr(ByVal Formula As String _
                                             , ByVal FormulaInSheet As Worksheet) As Collection
    
    If Formula = vbNullString Then
        Set GetDirectPrecedentsFromExpr = New Collection
        Exit Function
    End If
    
    #If DEVELOPMENT_MODE Then
        Dim ParseResult As OARobot.FormulaParseResult
        Dim CurrentExpr As OARobot.Expr
    #Else
        Dim ParseResult As Object
        Dim CurrentExpr As Object
    #End If
    
    Dim FormulaInBook As Workbook
    Set FormulaInBook = FormulaInSheet.Parent
    
    Set ParseResult = ParseFormula(Formula, FormulaInBook)
    
    If Not ParseResult.ParseSuccess Then Err.Raise 13, "DirectPrecedents", "Formula parsing failed."
    
    Dim Precedents As Collection
    Set Precedents = New Collection
    
    Dim Counter As Long
    For Counter = 0 To ParseResult.Expr.DirectPrecedents.Count - 1
        Set CurrentExpr = ParseResult.Expr.DirectPrecedents.Item(Counter)
        Precedents.Add CurrentExpr.Formula
    Next Counter
    
    Set GetDirectPrecedentsFromExpr = Precedents
           
End Function

Private Function GetUsedFunctions(ByVal Formula As String, Optional ByVal IsR1C1 As Boolean = False) As Variant
    
    If Formula = vbNullString Then
        GetUsedFunctions = vbEmpty
        Exit Function
    End If
    
    #If DEVELOPMENT_MODE Then
        Dim ParseResult As OARobot.FormulaParseResult
    #Else
        Dim ParseResult As Object
    #End If
    
    Set ParseResult = ParseFormula(Formula, , IsR1C1)
    
    If Not ParseResult.ParseSuccess Then Err.Raise 13, "UsedFunctions", "Formula parsing failed."
    
    GetUsedFunctions = ParseResult.Expr.UsedFunctions
           
End Function

Private Sub PrintActiveCellDependency()
    
    Debug.Print "Dependencies for : " & ActiveCell.Address
    Debug.Print "Dependencies for formula : " & ActiveCell.Formula2
    Dim Dependencies As Variant
    Dependencies = GetDirectPrecedents(ActiveCell.Formula2, ActiveSheet)
    Dim CurrentDependency As Variant
    For Each CurrentDependency In Dependencies
        Debug.Print CurrentDependency
    Next CurrentDependency
    Debug.Print
    
End Sub

Public Function GetDependencyFunctionResult(ByVal Formula As String _
                                            , ByVal DependencyFunctionName As DependencyFunctions _
                                             , Optional ByVal RemoveHeaderRow As Boolean = True) As Variant
    
    #If DEVELOPMENT_MODE Then
        Dim ParseResult As OARobot.FormulaParseResult
        Dim Processor As OARobot.ExprProcessing
        Set Processor = New OARobot.ExprProcessing
    #Else
        Dim ParseResult As Object
        Dim Processor As Object
        Set Processor = CreateObject("OARobot.ExprProcessing")
    #End If
    
    Set ParseResult = ParseFormula(Formula)
    
    Dim Result As Variant
    Select Case DependencyFunctionName
    
        Case DependencyFunctions.LET_PARTS
            Result = Processor.LetParts(ParseResult.Expr)

        Case DependencyFunctions.LAMBDA_PARTS
            Result = Processor.LambdaParts(ParseResult.Expr)
            
        Case Else
            Err.Raise 13, "Wrong Input Argument"

    End Select
    
    If RemoveHeaderRow Then Result = RemoveTopRowHeader(Result)
    GetDependencyFunctionResult = Result

End Function

Public Function GetLambdaDefPart(ByVal LambdaFormula As String) As String
    
    Dim SplittedPart As Variant
    SplittedPart = SplitLambdaDef(LambdaFormula)
    GetLambdaDefPart = SplittedPart(LBound(SplittedPart))
    
End Function

Public Function GetLambdaInvocationPart(ByVal LambdaFormula As String) As String
    
    Dim SplittedPart As Variant
    SplittedPart = SplitLambdaDef(LambdaFormula)
    GetLambdaInvocationPart = SplittedPart(LBound(SplittedPart) + 1)
    
End Function

Public Function FormatFormula(ByVal FormulaText As String) As String
    
    If FormulaText = vbNullString Then Exit Function
    #If DEVELOPMENT_MODE Then
        Dim Formatter As OARobot.FormulaFormatter
    #Else
        Dim Formatter As Object
    #End If
    
    Set Formatter = GetFormulaFormatter()
    
    With Formatter
    
        ' Set configuration from user context.
        If GetBoModeParamValue() Then
            .CompactConfig
        Else
            .IndentChar = GetIndentCharParamValue()
            .IndentSize = GetIndentSizeParamValue()
            .MultiLine = GetMultilineParamValue()
            .OnlyWrapFunctionAfterNChars = GetOnlyWrapFunctionAfterNCharsParamValue()
            .SpacesAfterArgumentSeparators = GetSpacesAfterArgumentSeparatorsParamValue()
            .SpacesAfterArrayColumnSeparators = GetSpacesAfterArrayColumnSeparatorsParamValue()
            .SpacesAfterArrayRowSeparators = GetSpacesAfterArrayRowSeparatorsParamValue()
            .SpacesAroundInfixOperators = GetSpacesAroundInfixOperatorsParamValue()
        End If
        
        ' Format the formula
        FormatFormula = .Format(FormulaText)
    End With
    Set Formatter = Nothing
    
End Function

' This will extract upto the function definition. Meaning if the lambda is =LAMBDA(a,b,a*2)(10,5) >> =LAMBDA(a,b,
' It handled optional arguments, no param lambda as well.
Public Function GetUptoLambdaParamDefPart(ByVal LambdaFormula As String) As String
    
    Dim DefPart As String
    DefPart = GetLambdaDefPart(LambdaFormula)
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
        Dim Parameters As OARobot.TokenCollection
    #Else
        Dim ParsedFormulaResult As Object
        Dim Parameters As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(DefPart)
    If ParsedFormulaResult.Expr.IsLambda Then
        
        Set Parameters = ParsedFormulaResult.Expr.AsLambda.Parameters
        Dim ParamDefPart As String
        ParamDefPart = EQUAL_SIGN & LAMBDA_FX_NAME & FIRST_PARENTHESIS_OPEN
        Dim CurrentParam As Object
        
        Dim Counter As Long
        For Counter = 0 To Parameters.Count - 1
            Set CurrentParam = Parameters.Item(Counter)
            ParamDefPart = ParamDefPart & CurrentParam.String & LIST_SEPARATOR
        Next Counter
        
        GetUptoLambdaParamDefPart = ParamDefPart
        
    Else
        Err.Raise 13, "This function was expecting a lambda function."
    End If
    
End Function

Private Sub TestSplitLambdaDef()
    
    Dim SplittedDef As Variant
    SplittedDef = SplitLambdaDef(ActiveCell.Formula2)
    Debug.Print "Lambda Def:" & vbNewLine & SplittedDef(LBound(SplittedDef)) & vbNewLine
    Debug.Print "Invocation:" & vbNewLine & SplittedDef(UBound(SplittedDef)) & vbNewLine
    
End Sub

' This will check if starting formula is the entire formula.
' And this is only Applicable to LET Or LAMBDA Formula only. For example
'=Lambda(a,a*2) >> Valid
'=Lambda(a,a*2)(a4) >> Valid
'=Lambda(a,a*2)+1 >> Invalid
'=Lambda(a,a*2)(a4)+1 >>Invalid
'=Let(a,2,a*2) >>Valid
'=Let(a,2,a*2)+1 >>Invalid

Public Function IsStartingFormulaIsTheEntireFormula(ByVal Formula As String) As Boolean
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
        Dim InputExpr As OARobot.Expr
    #Else
        Dim ParsedFormulaResult As Object
        Dim InputExpr As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(Formula)
    
    If Not ParsedFormulaResult.ParseSuccess Then
        IsStartingFormulaIsTheEntireFormula = False
        Exit Function
    End If
    
    Set InputExpr = ParsedFormulaResult.Expr
    
    If InputExpr.IsFunction Then
        Set InputExpr = InputExpr.AsFunction.FunctionName
    End If
    
    With InputExpr
        IsStartingFormulaIsTheEntireFormula = (.IsLambda Or .IsLet Or .IsName)
    End With
    
End Function

Private Function SplitLambdaDef(ByVal LambdaFormula As String) As String()
    
    Dim SplittedFormula(0 To 1) As String
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
        Dim InputExpr As OARobot.Expr
        Dim InputFunctions As OARobot.ExprFunction
    #Else
        Dim ParsedFormulaResult As Object
        Dim InputExpr As Object
        Dim InputFunctions As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(LambdaFormula)
    
    If Not ParsedFormulaResult.ParseSuccess Then
        SplitLambdaDef = SplittedFormula
        Logger.Log DEBUG_LOG, "Failed to parse: " & ParsedFormulaResult.Formula
        Exit Function
    End If
    
    Set InputExpr = ParsedFormulaResult.Expr
    
    If InputExpr.IsFunction Then
        ' If lambda with invocation and only if the formula doesn't have any operands.
        ' For example =Lambda(a,a*2)+1 so here + is the operator and it will not be IsFunction
        
        Set InputFunctions = InputExpr.AsFunction
        If Not InputFunctions.FunctionName.IsLambda Then
            SplitLambdaDef = SplittedFormula
            Exit Function
        End If
                
        SplittedFormula(0) = InputFunctions.FunctionName.Formula(True)
        SplittedFormula(1) = GenerateInvocationPart(InputFunctions)
            
    ElseIf InputExpr.IsLambda Then
        ' If only lambda is present and no invocation.
        SplittedFormula(0) = InputExpr.Formula(True)
    End If
    
    SplitLambdaDef = SplittedFormula
    
End Function

#If DEVELOPMENT_MODE Then
Private Function GenerateInvocationPart(ByVal InputFunctions As OARobot.ExprFunction) As String
#Else
Private Function GenerateInvocationPart(ByVal InputFunctions As Object) As String
#End If
    
    #If DEVELOPMENT_MODE Then
        Dim Args As OARobot.ExprCollection
        Dim ArgSeps As OARobot.TokenCollection
    #Else
        Dim Args As Object
        Dim ArgSeps As Object
    #End If
    
    Dim InvocationPart As String
    InvocationPart = InputFunctions.LeftParen.String
            
    Set Args = InputFunctions.Args
    Set ArgSeps = InputFunctions.ArgSeparators
    Dim Counter As Long
                            
    For Counter = 0 To Args.Count - 1
        InvocationPart = InvocationPart & Args.Item(Counter).Formula(False)
        If Counter <= ArgSeps.Count - 1 Then
            InvocationPart = InvocationPart & ArgSeps.Item(Counter).String
        End If
    Next Counter
    
    InvocationPart = InvocationPart & InputFunctions.RightParen.String
    GenerateInvocationPart = InvocationPart
    
End Function

Public Function AddLetStep(ByVal Formula As String _
                           , ByVal NewStepName As String _
                            , Optional ByVal NewLetStepExpression As String = "{{LastStep}}") As String
    
    
    #If DEVELOPMENT_MODE Then
        Dim FormulaProcessor As OARobot.FormulaProcessing
    #Else
        Dim FormulaProcessor As Object
    #End If
    
    Set FormulaProcessor = GetFormulaProcessor()
    
    AddLetStep = FormulaProcessor.AddLetStep(Formula _
                                             , NewStepName _
                                              , NewLetStepExpression _
                                               , Application.ReferenceStyle = xlR1C1)
    
    Set FormulaProcessor = Nothing
    
End Function

Public Function InsertLetStep(ByVal Formula As String _
                              , ByVal Index As Integer _
                               , ByVal NewStepName As String _
                                , ByVal NewValue As String) As String
    
    #If DEVELOPMENT_MODE Then
        Dim FormulaProcessor As OARobot.FormulaProcessing
    #Else
        Dim FormulaProcessor As Object
    #End If
    
    Set FormulaProcessor = GetFormulaProcessor()
    InsertLetStep = FormulaProcessor.InsertLetStep(Formula _
                                                   , Index _
                                                    , NewStepName _
                                                     , NewValue _
                                                      , Application.ReferenceStyle = xlR1C1)
    Set FormulaProcessor = Nothing
    
End Function

Public Function RemoveLetStep(ByVal Formula As String _
                              , ByVal NameToRemove As String) As String
    
    #If DEVELOPMENT_MODE Then
        Dim FormulaProcessor As OARobot.FormulaProcessing
    #Else
        Dim FormulaProcessor As Object
    #End If
    
    Set FormulaProcessor = GetFormulaProcessor()
    
    RemoveLetStep = FormulaProcessor.RemoveLetStep(Formula _
                                                   , NameToRemove _
                                                    , Application.ReferenceStyle = xlR1C1)
    Set FormulaProcessor = Nothing
    
End Function

Public Function MoveParamToLetStep(ByVal Formula As String _
                                   , ByVal NameToMove As String) As String
    
    #If DEVELOPMENT_MODE Then
        Dim FormulaProcessor As OARobot.FormulaProcessing
    #Else
        Dim FormulaProcessor As Object
    #End If
    
    Set FormulaProcessor = GetFormulaProcessor()
    MoveParamToLetStep = FormulaProcessor.MoveParamToLetStep(Formula _
                                                             , NameToMove _
                                                              , Application.ReferenceStyle = xlR1C1)
    Set FormulaProcessor = Nothing
    
End Function

Public Function MoveLetStepToParam(ByVal Formula As String _
                                   , ByVal NameToMove As String _
                                    , ByVal MakeOptional As Boolean) As String
    
    #If DEVELOPMENT_MODE Then
        Dim FormulaProcessor As OARobot.FormulaProcessing
    #Else
        Dim FormulaProcessor As Object
    #End If
    
    Set FormulaProcessor = GetFormulaProcessor()
    MoveLetStepToParam = FormulaProcessor.MoveLetStepToParam(Formula _
                                                             , NameToMove _
                                                              , MakeOptional _
                                                               , Application.ReferenceStyle = xlR1C1)
    Set FormulaProcessor = Nothing
    
End Function

Public Function IsInvocationExist(ByVal LambdaFormula As String) As Boolean
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
    #Else
        Dim ParsedFormulaResult As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(LambdaFormula)
    
    If ParsedFormulaResult.ParseSuccess Then
        If ParsedFormulaResult.Expr.IsFunction Then
            IsInvocationExist = ParsedFormulaResult.Expr.AsFunction.FunctionName.IsLambda
        End If
    End If
    
End Function

Private Sub TestGetParametersAndStepsName()
    Dim Names As Variant
    Names = GetParametersAndStepsName(ActiveCell.Formula2)
End Sub

Public Function GetParametersAndStepsName(ByVal LetOrLambdaFormula As String) As Variant
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
        Dim InputExpr As OARobot.Expr
        Dim InputFunctions As OARobot.ExprFunction
    #Else
        Dim ParsedFormulaResult As Object
        Dim InputExpr As Object
        Dim InputFunctions As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(LetOrLambdaFormula)
    
    If Not ParsedFormulaResult.ParseSuccess Then
        GetParametersAndStepsName = vbEmpty
        Exit Function
    End If
    
    Set InputExpr = ParsedFormulaResult.Expr
    
    Dim Result As Collection
    Set Result = New Collection
    
    If InputExpr.IsFunction Then
        
        ' If lambda with invocation and only if the formula doesn't have any operands.
        ' For example =Lambda(a,a*2)+1 so here + is the operator and it will not be IsFunction
        
        Set InputFunctions = InputExpr.AsFunction
        
        If Not InputFunctions.FunctionName.IsLambda Then
            GetParametersAndStepsName = vbEmpty
            Exit Function
        End If
        AddParametersAndLetStepsToCollection Result, InputFunctions.FunctionName.AsLambda
        
    ElseIf InputExpr.IsLambda Then
        AddParametersAndLetStepsToCollection Result, InputExpr.AsLambda
    ElseIf InputExpr.IsLet Then
        AddLetStepsNameToCollection Result, InputExpr.AsLet
    Else
        GetParametersAndStepsName = vbEmpty
        Exit Function
    End If
    GetParametersAndStepsName = CollectionToArray(Result)
    
    Set Result = Nothing
    
End Function

#If DEVELOPMENT_MODE Then
Private Sub AddParametersAndLetStepsToCollection(ByRef Result As Collection _
                                                 , ByVal LambdaExpr As OARobot.ExprLambda)
#Else
Private Sub AddParametersAndLetStepsToCollection(ByRef Result As Collection _
                                                 , ByVal LambdaExpr As Object)
#End If
    
    Dim CurrentParam As Object
    Dim ParamName As String
    ' Generate parameter part.
        
    Dim Counter As Long
    For Counter = 0 To LambdaExpr.Parameters.Count - 1
        
        Set CurrentParam = LambdaExpr.Parameters.Item(Counter)
        ' Remove square bracket from optional arguments.
        ParamName = Text.RemoveFromStartIfPresent(CurrentParam.String, LEFT_BRACKET)
        ParamName = Text.RemoveFromEndIfPresent(ParamName, RIGHT_BRACKET)
        If Not IsExistInCollection(Result, ParamName) Then
            Result.Add ParamName, ParamName
        End If
        
    Next Counter
        
    If LambdaExpr.Body.IsLet Then
        AddLetStepsNameToCollection Result, LambdaExpr.Body.AsLet
    End If
    
End Sub

#If DEVELOPMENT_MODE Then
Private Sub AddLetStepsNameToCollection(ByRef Result As Collection, ByVal LetExpr As OARobot.ExprLet)
#Else
Private Sub AddLetStepsNameToCollection(ByRef Result As Collection, ByVal LetExpr As Object)
#End If
        
    Dim Counter As Long
    For Counter = 0 To LetExpr.Names.Count - 1
        Dim StepName As String
        StepName = LetExpr.Names.Item(Counter).String
        If Not IsExistInCollection(Result, StepName) Then
            Result.Add StepName, StepName
        End If
            
    Next Counter
    
End Sub

Public Function RemoveMetadataFromFormulaCOM(ByVal LetOrLambdaFormula As String) As String
    
    Dim LetParts As Variant
    LetParts = GetDependencyFunctionResult(LetOrLambdaFormula, LET_PARTS, True)
    If Not IsArrayAllocated(LetParts) Then
        RemoveMetadataFromFormulaCOM = LetOrLambdaFormula
        Exit Function
    End If
    
    Dim FinalFormula As String
    FinalFormula = LetOrLambdaFormula
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(LetParts, 2)
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(LetParts, 1) To UBound(LetParts, 1) - 1
        Dim StepName As String
        StepName = LetParts(CurrentRowIndex, FirstColumnIndex)
        If IsMetadataLetVarName(StepName) Then
            FinalFormula = RemoveLetStep(FinalFormula, StepName)
        End If
    Next CurrentRowIndex
    RemoveMetadataFromFormulaCOM = FormatFormula(FinalFormula)
    
    
End Function

' Remove metadata from the formula and return back same formula if no metadata is present or if parsing failed.
Public Function RemoveMetadataFromFormula(ByVal LetOrLambdaFormula As String) As String
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
        Dim InputExpr As OARobot.Expr
        Dim InputFunctions As OARobot.ExprFunction
    #Else
        Dim ParsedFormulaResult As Object
        Dim InputExpr As Object
        Dim InputFunctions As Object
    #End If
    
    
    Set ParsedFormulaResult = ParseFormula(LetOrLambdaFormula)
    
    If Not ParsedFormulaResult.ParseSuccess Then
        RemoveMetadataFromFormula = LetOrLambdaFormula
        Exit Function
    End If
    
    Set InputExpr = ParsedFormulaResult.Expr
    
    Dim Formula As String
    If InputExpr.IsFunction Then
        ' If lambda with invocation and only if the formula doesn't have any operands.
        ' For example =Lambda(a,a*2)+1 so here + is the operator and it will not be IsFunction
        
        Set InputFunctions = InputExpr.AsFunction
        
        If Not InputFunctions.FunctionName.IsLambda Then
            RemoveMetadataFromFormula = LetOrLambdaFormula
            Exit Function
        End If
        Formula = RemoveMetadataFromLambdaExpr(InputFunctions.FunctionName.AsLambda)
        Formula = Formula & GetLambdaInvocationPart(LetOrLambdaFormula)
    ElseIf InputExpr.IsLambda Then
        ' If only lambda is present and no invocation.
        Formula = RemoveMetadataFromLambdaExpr(InputExpr.AsLambda)
    ElseIf InputExpr.IsLet Then
        Formula = RemoveMetadataFromLetExpr(InputExpr.AsLet)
    Else
        Formula = LetOrLambdaFormula
    End If
    
    If Formula <> vbNullString Then
        Formula = Text.PadIfNotPresent(Formula, EQUAL_SIGN, FROM_START)
        Formula = FormatFormula(Formula)
    End If
    
    RemoveMetadataFromFormula = Formula
    
End Function

' Check if the outer function is LAMBDA or not and it is the entire function.
Public Function IsLambdaFunction(ByVal Formula As String) As Boolean
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
    #Else
        Dim ParsedFormulaResult As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(Formula)
    
    If Not ParsedFormulaResult.ParseSuccess Then
        IsLambdaFunction = False
        Exit Function
    End If
    
    Dim Result As Boolean
    If ParsedFormulaResult.Expr.IsLambda Then
        Result = True
    ElseIf ParsedFormulaResult.Expr.IsFunction Then
        Result = ParsedFormulaResult.Expr.AsFunction.FunctionName.IsLambda
    Else
        Result = False
    End If
    
    IsLambdaFunction = Result
    
End Function

Public Function GetSavedNamedNameFromCellFormula(ByVal FormulaText As String _
                                           , ByVal FormulaSheet As Worksheet) As String
    
    ' Extract the saved name from the cell formula. Like if we have a formula =AnySavedName or lambda invocation
    ' =MyLambda(Param1..) then it will extract the Lambda Name or Named range name.
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
    #Else
        Dim ParsedFormulaResult As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(FormulaText, FormulaSheet.Parent)
    
    Dim Result As String
    
    If Not ParsedFormulaResult.ParseSuccess Then
        Result = vbNullString
    ElseIf ParsedFormulaResult.Expr.IsName Then
        Result = ParsedFormulaResult.Expr.AsName.name.String
    ElseIf ParsedFormulaResult.Expr.IsFunction Then
        Result = ParsedFormulaResult.Expr.AsFunction.FunctionName.AsName.name.String
    End If
    
    GetSavedNamedNameFromCellFormula = Result
    
End Function

Public Function IsSavedLambdaInCellFormula(ByVal FormulaText As String _
                                           , ByVal FormulaSheet As Worksheet) As Boolean
    
    Dim LambdaName As String
    LambdaName = GetSavedNamedNameFromCellFormula(FormulaText, FormulaSheet)
    
    Dim Result As Boolean
    Dim RefersTo As String
    
    If LambdaName = vbNullString Then
        Result = False
    Else
        RefersTo = FormulaSheet.Parent.Names(LambdaName).RefersTo
        Result = IsLambdaFunction(RefersTo)
    End If
    
    IsSavedLambdaInCellFormula = Result
    
End Function

Public Function IsLambdaWithLet(ByVal Formula As String) As Boolean
    
    Dim Result As Boolean
    If Not IsLambdaFunction(Formula) Then
        Result = False
    Else
        
        #If DEVELOPMENT_MODE Then
            Dim ParsedFormulaResult As OARobot.FormulaParseResult
        #Else
            Dim ParsedFormulaResult As Object
        #End If
    
        Set ParsedFormulaResult = ParseFormula(Formula)
        
        With ParsedFormulaResult.Expr
            If .IsLambda Then
                Result = .AsLambda.Body.IsLet
            ElseIf .IsFunction Then
                Result = .AsFunction.FunctionName.AsLambda.Body.IsLet
            End If
        End With
        
    End If
    
    IsLambdaWithLet = Result
    
End Function

' Check if outer function is LET or not and it is the entire function.
Public Function IsLetFunction(ByVal Formula As String) As Boolean
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
    #Else
        Dim ParsedFormulaResult As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(Formula)
    
    If Not ParsedFormulaResult.ParseSuccess Then
        IsLetFunction = False
        Exit Function
    End If
    
    If ParsedFormulaResult.Expr.IsLet Then
        IsLetFunction = True
    ElseIf ParsedFormulaResult.Expr.IsFunction Then
        IsLetFunction = ParsedFormulaResult.Expr.AsFunction.FunctionName.IsLet
    Else
        IsLetFunction = False
    End If
    
    Set ParsedFormulaResult = Nothing
    
End Function

' If a let function return a lambda as an output then we can invoke that. This will return that invocation part
' Example formula =LET(a,1,LAMBDA(x,y,LET(z,SEQUENCE(y,x),result,TRANSPOSE(z),2*(result+1))))(5,8)
' Output will be (5,8)
'?GetLetFormulaInvocation(ActiveCell.Formula2)
Public Function GetLetFormulaInvocation(ByVal Formula As String) As String
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
        Dim Args As OARobot.ExprCollection
    #Else
        Dim ParsedFormulaResult As Object
        Dim Args As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(Formula)
    
    Dim Invocation As String
    
    If Not ParsedFormulaResult.ParseSuccess Then
        ' If parse failed
        Invocation = vbNullString
    ElseIf ParsedFormulaResult.Expr.IsLet Then
        ' Parsing successful but Let statement
        Invocation = vbNullString
    ElseIf Not ParsedFormulaResult.Expr.IsFunction Then
        ' If not a function then
        Invocation = vbNullString
    ElseIf ParsedFormulaResult.Expr.AsFunction.FunctionName.IsLet _
           And ParsedFormulaResult.Expr.AsFunction.Args.Count > 0 Then
        
        Invocation = FIRST_PARENTHESIS_OPEN
        Set Args = ParsedFormulaResult.Expr.AsFunction.Args
        Dim Counter As Long
        For Counter = 0 To Args.Count - 1
            Invocation = Invocation & Args.Item(Counter).Formula & LIST_SEPARATOR
        Next Counter
        Invocation = Text.RemoveFromEndIfPresent(Invocation, LIST_SEPARATOR) & FIRST_PARENTHESIS_CLOSE
    
    End If
    
    GetLetFormulaInvocation = Invocation
    
    Set ParsedFormulaResult = Nothing
    
End Function

' This will remove metadata from a Let statement.
#If DEVELOPMENT_MODE Then
Private Function RemoveMetadataFromLetExpr(ByVal LetExpr As OARobot.ExprLet) As String
#Else
Private Function RemoveMetadataFromLetExpr(ByVal LetExpr As Object) As String
#End If
    
    #If DEVELOPMENT_MODE Then
        Dim Names As OARobot.TokenCollection
        Dim Values As OARobot.ExprCollection
    #Else
        Dim Names As Object
        Dim Values As Object
    #End If
        
    Set Names = LetExpr.Names
        
    Dim NonMetadataStepsIndex As Collection
    Set NonMetadataStepsIndex = New Collection
        
    Dim Counter As Long
    For Counter = 0 To Names.Count - 1
        Dim StepName As String
        StepName = Names.Item(Counter).String
        If Not Text.IsStartsWith(StepName, METADATA_IDENTIFIER) Then
            NonMetadataStepsIndex.Add Counter, CStr(Counter)
        End If
    Next Counter
        
        
    Set Values = LetExpr.Values
                
    Dim Formula As String
    Dim ValidIndex As Variant
    Select Case NonMetadataStepsIndex.Count
        
        Case 0
            Formula = GetBodyExpression(LetExpr)
                
            ' If only one non metadata then no need to use LET
        Case 1
            ValidIndex = NonMetadataStepsIndex.Item(1)
            Dim BodyExpression As String
            BodyExpression = GetBodyExpression(LetExpr, CLng(ValidIndex))
            If BodyExpression = Names.Item(ValidIndex).String Then
                Formula = Values.Item(NonMetadataStepsIndex.Item(1)).Formula(False)
            Else
                Formula = LetExpr.FunctionName.String & LetExpr.LeftParen.String _
                          & Names.Item(ValidIndex).String & LIST_SEPARATOR & Values.Item(ValidIndex).Formula _
                          & LIST_SEPARATOR & BodyExpression & LetExpr.RightParen.String
            End If
            
            ' If no metadata found
        Case Names.Count
            Formula = LetExpr.AsExpr.Formula(False)
                
        Case Else
            Formula = LetExpr.FunctionName.String
            Formula = Formula & LetExpr.LeftParen.String
                
            For Each ValidIndex In NonMetadataStepsIndex
                    
                ' Calling Formula() on a sub-expression returns a string containing only the sub-expression
                Formula = Formula & Names.Item(ValidIndex).String _
                          & LIST_SEPARATOR & Values.Item(ValidIndex).Formula() _
                          & LIST_SEPARATOR
                    
            Next ValidIndex
                
            Formula = Formula & GetBodyExpression(LetExpr _
                                                  , NonMetadataStepsIndex.Item(NonMetadataStepsIndex.Count))
            Formula = Formula & LetExpr.RightParen.String
                
    End Select
        
    RemoveMetadataFromLetExpr = Formula
    Set NonMetadataStepsIndex = Nothing
        
End Function

#If DEVELOPMENT_MODE Then
Private Function GetBodyExpression(ByVal LetExpr As OARobot.ExprLet _
                                   , Optional ByVal LastValidStepIndex As Long = 0) As String
#Else
Private Function GetBodyExpression(ByVal LetExpr As Object _
                                   , Optional ByVal LastValidStepIndex As Long = 0) As String
#End If
    'If it is a step name
    If LetExpr.Body.IsName Then
        'If name but it was a metadata then use the last step.
        If Text.IsStartsWith(LetExpr.Body.AsName.name.String, METADATA_IDENTIFIER) Then
            GetBodyExpression = Names(LastValidStepIndex).String
        Else
            ' Else use the name
            GetBodyExpression = LetExpr.Body.AsName.name.String
        End If
    Else
        ' If it was a expr then use .Formula to get the whole sub expression.
        GetBodyExpression = LetExpr.Body.Formula()
    End If
    
End Function

' This will remove metadata from a Lambda. This will only return the def part. No invocation will be returned.
#If DEVELOPMENT_MODE Then
Private Function RemoveMetadataFromLambdaExpr(ByVal LambdaExpr As OARobot.ExprLambda) As String
#Else
Private Function RemoveMetadataFromLambdaExpr(ByVal LambdaExpr As Object) As String
#End If
    
    Dim Formula As String
    Formula = LambdaExpr.FunctionName.String
    Formula = Formula & LambdaExpr.LeftParen.String
            
    #If DEVELOPMENT_MODE Then
        Dim Params As OARobot.TokenCollection
    #Else
        Dim Params As Object
    #End If

    Set Params = LambdaExpr.Parameters
    
    ' Generate parameter part.
    Dim Counter As Long
    For Counter = 0 To Params.Count - 1
        Formula = Formula & Params.Item(Counter).String
        Formula = Formula & LIST_SEPARATOR
    Next Counter
            
    If LambdaExpr.Body.IsLet Then
        Formula = Formula & RemoveMetadataFromLetExpr(LambdaExpr.Body.AsLet)
    Else
        Formula = Formula & LambdaExpr.Body.Formula(False)
    End If
            
    Formula = Formula & LambdaExpr.RightParen.String
    RemoveMetadataFromLambdaExpr = Formula
        
End Function

' Check if any metadata is present in a formula or not
Public Function IsMetadataPresent(ByVal LambdaFormula As String) As Boolean
    
    Dim LetParts As Variant
    LetParts = GetDependencyFunctionResult(LambdaFormula, LET_PARTS)
    If Not IsArrayAllocated(LetParts) Then Exit Function
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(LetParts, 2)
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(LetParts, 1) To UBound(LetParts, 1)
        Dim StepName As String
        StepName = LetParts(CurrentRowIndex, FirstColumnIndex)
        If IsMetadataLetVarName(StepName) Then
            IsMetadataPresent = True
            Exit Function
        End If
    Next CurrentRowIndex
    IsMetadataPresent = False
    
End Function

' Return Expr object by parsing the formula. If parsing fail then it will return nothing
#If DEVELOPMENT_MODE Then
Public Function GetExpr(ByVal Formula As String) As OARobot.Expr
#Else
Public Function GetExpr(ByVal Formula As String) As Object
#End If
        
    #If DEVELOPMENT_MODE Then
        Dim ParsedResult As OARobot.FormulaParseResult
    #Else
        Dim ParsedResult As Object
    #End If
    
    Set ParsedResult = ParseFormula(Text.PadIfNotPresent(Formula, EQUAL_SIGN, FROM_START))
    If ParsedResult.ParseSuccess Then
        Set GetExpr = ParsedResult.Expr
    Else
        Set GetExpr = Nothing
    End If
    
End Function

Private Sub ReplaceStepNameWithCellRefExample()

    Dim Formula As String
    Formula = "=MMULT(" & vbNewLine & _
              "INDEX(D6:M13, MID(AA6#, StepName * 2 - 1, 2), StepName)," & vbNewLine & _
              "TOCOL(StepName) ^ 0" & vbNewLine & _
              ")"
    
    Debug.Print "Input Formula : " & vbNewLine & vbNewLine & Formula & vbNewLine
    
    Dim ReplacedFormula As String
    ReplacedFormula = ReplaceTokenWithNewToken(Formula, "StepName", "Q6#")
    Debug.Print "Output Formula : " & vbNewLine & vbNewLine & ReplacedFormula & vbNewLine
    
End Sub

Private Sub TestReplaceTokenWithNewToken()
    
    Dim TestFormula As String
    TestFormula = "=LETStep_FillRowVectorFromLeft(TAKE(DataMatrix,1))"
    Dim ActualFormula As String
    ActualFormula = ReplaceTokenWithNewToken(TestFormula, "LETStep_FillRowVectorFromLeft", "_FillRowVectorFromLeft")
    
    Dim ExpectedFormula As String
    ExpectedFormula = "=_FillRowVectorFromLeft(TAKE(DataMatrix,1))"
    
    If ExpectedFormula = ActualFormula Then
        Logger.Log DEBUG_LOG, "Test Passed."
    Else
        Logger.Log DEBUG_LOG, "Test Failed."
        Logger.Log DEBUG_LOG, "Actual Formula : " & ActualFormula
    End If
    
End Sub

Public Function ReplaceTokenWithNewToken(ByVal OnFormula As String _
                                         , ByVal OldToken As String, NewToken As String) As String
    
    #If DEVELOPMENT_MODE Then
        Dim ExprReplacer As OARobot.ExpressionReplacer
        Dim InputExpr As OARobot.Expr
    #Else
        Dim ExprReplacer As Object
        Dim InputExpr As Object
    #End If
    
    Set ExprReplacer = GetExpressionReplacer()
    
    With ExprReplacer
        .FindWhat = OldToken
        .ReplaceWith = NewToken
    End With

    Set InputExpr = GetExpr(OnFormula)
    Set InputExpr = InputExpr.Rewrite(ExprReplacer)
    ReplaceTokenWithNewToken = InputExpr.Formula(True)
    
End Function

Private Sub ReplaceCellRefWithStepNameExample()

    Dim Formula As String
    
    Formula = "=MMULT(" & vbNewLine & _
              "INDEX(D6:M13, MID(AA6#, Q6# * 2 - 1, 2), Q6#)," & vbNewLine & _
              "TOCOL(Q6#) ^ 0" & vbNewLine & _
              ")"
    
    RunReplaceCellRefWithStepNameTest Formula, "StepName", "Q6#", "Sheet 1", Replace(Formula, "Q6#", "StepName")
    
    Dim ReplacedFormula As String
    Formula = "=FILTER(Table1[Balance],(Table1[State]=Q8) * (Table1[Balance]>=P8))"
    RunReplaceCellRefWithStepNameTest Formula, "_State", "Table1[State]", "Sheet 1", Replace(Formula, "Table1[State]", "_State")
    
    Formula = "=FILTER(Balance,(State=Q8)*(Balance>=P8))"
    
    RunReplaceCellRefWithStepNameTest Formula, "_StateName", "State", "Sheet 1", Replace(Formula, "State", "_StateName")
    
    Formula = "='Blank Range Label'!TestResult"
    RunReplaceCellRefWithStepNameTest Formula, "_LocalNamedRange", "'Blank Range Label'!TestResult", "Blank Range Label", "=_LocalNamedRange"
    
    Formula = "=FILTER('MultiCell No Marked Input Cells 2'!K8:K61,('MultiCell No Marked Input Cells'!K8:K61=ForState)*('MultiCell No Marked Input Cells'!N8:N61>=MinimumBalance))"
    
    RunReplaceCellRefWithStepNameTest Formula, "_State", "K8:K61" _
                                                        , "MultiCell No Marked Input Cells" _
                                                         , Replace(Formula, "'MultiCell No Marked Input Cells'!K8:K61", "_State")
    
    RunReplaceCellRefWithStepNameTest "=Sheet2!TestResult+QuickBook.xlsx!TestResult", "SampleName", "QuickBook.xlsx!TestResult", "Sheet2", "=SampleName+QuickBook.xlsx!TestResult"
    
    RunReplaceCellRefWithStepNameTest "='Input Sheet'!A1+A1+Sheet2!A1", "Step", "Sheet2!A1", "Sheet3", "='Input Sheet'!A1+Step+Sheet2!A1"
    
End Sub

Private Sub RunReplaceCellRefWithStepNameTest(OnFormula As String _
                                              , StepName As String _
                                               , CellRef As String _
                                                , SheetName As String _
                                                 , ExpectedFormula As String)
    
    Dim ActualFormula As String
    ActualFormula = ReplaceCellRefWithStepName(OnFormula, StepName, CellRef, SheetName)
    
    Debug.Print
    Debug.Print "Actual Formula : " & vbNewLine & ActualFormula
    Debug.Print
    Debug.Print "Expected Formula : " & vbNewLine & ExpectedFormula
    Debug.Print "Test Pass ? : " & (ActualFormula = ExpectedFormula)

End Sub

Private Sub TestReplaceCellRefWithStepName()
    RunReplaceCellRefWithStepNameTest "=FILTER(Table1[Balance],(Table1[State]=Q8) * (Table1[Balance]>=P8))" _
                                      , "minimum_balance", "P8", "Table As Precedency" _
                                      , "=FILTER(Table1[Balance],(Table1[State]=Q8) * (Table1[Balance]>=minimum_balance))"
End Sub

Public Function ReplaceCellRefWithStepName(ByVal OnFormula As String _
                                           , ByVal StepName As String _
                                            , ByVal CellRef As String _
                                             , ByVal SheetName As String) As String
    
    If OnFormula = vbNullString Then Exit Function
    
    #If DEVELOPMENT_MODE Then
        Dim ExprReplacer As OARobot.ExpressionReplacer
        Dim InputExpr As OARobot.Expr
    #Else
        Dim ExprReplacer As Object
        Dim InputExpr As Object
    #End If
    
    Set ExprReplacer = GetExpressionReplacer()
            
    With ExprReplacer
        .FindWhat = CellRef
        .ReplaceWith = StepName
        .IsFindRangeRef = True
        .SheetName = SheetName
    End With
    
    Set InputExpr = GetExpr(OnFormula)
    Set InputExpr = InputExpr.Rewrite(ExprReplacer)
    ReplaceCellRefWithStepName = InputExpr.Formula(True)
    
End Function

Public Function RemoveSheetNameFromFormula(ByVal Formula As String _
, ByVal SheetName As String) As String
    
    #If DEVELOPMENT_MODE Then
        Dim InputExpr As OARobot.Expr
        Dim SheetNameRemover As OARobot.SheetNameRemover
    #Else
        Dim InputExpr As Object
        Dim SheetNameRemover As Object
    #End If
    
    Set SheetNameRemover = GetSheetNameRemover()
    SheetNameRemover.SheetName = EscapeSingeQuote(SheetName)
    
    Set InputExpr = GetExpr(Formula)
    Set InputExpr = InputExpr.Rewrite(SheetNameRemover)
    RemoveSheetNameFromFormula = InputExpr.Formula(True)

End Function

' Find all used functions and named ranges.
Public Function GetNamesAndFunctions(ByVal Formula As String) As Collection
    
    #If DEVELOPMENT_MODE Then
        Dim FormulaExpr As OARobot.Expr
        Dim Filterer As OARobot.NamesOrFunctionsFilter
    #Else
        Dim FormulaExpr As Object
        Dim Filterer As Object
    #End If
    
    Set Filterer = GetNamesOrFunctionsFilter()
    
    Dim NameAndFunctionFiltered As Object
    Set FormulaExpr = GetExpr(Formula)
    Set NameAndFunctionFiltered = FormulaExpr.Descendants(True, Filterer)
    
    Dim UsedFunctions As Collection
    Set UsedFunctions = New Collection
    
    On Error Resume Next
    Dim Counter As Long
    For Counter = 0 To NameAndFunctionFiltered.Count - 1
        Dim UsedFXOrName As String
        UsedFXOrName = NameAndFunctionFiltered.Item(Counter).Formula
        UsedFunctions.Add UsedFXOrName, UsedFXOrName
    Next Counter
    On Error GoTo 0
    
    Set GetNamesAndFunctions = UsedFunctions
    
    Set UsedFunctions = Nothing
    
End Function

' Get Naming Convention from Default OA Param
Public Function GetNamingConventionParam(ByVal IsForParam As Boolean) As String
    
    If IsForParam Then
        GetNamingConventionParam = Text.Proper(GetOAParamValue("FormulaFormat_LambdaParamStyle", "Snake_Case"))
    Else
        GetNamingConventionParam = Text.Proper(GetOAParamValue("FormulaFormat_VariableStyle", "Pascal"))
    End If
    
End Function

Public Function GetNamingConv(IsForParam As Boolean) As VarNamingStyle
    
    Dim NamingConv As String
    NamingConv = GetNamingConventionParam(IsForParam)
    Select Case NamingConv
        
        Case "Pascal"
            GetNamingConv = VarNamingStyle.PASCAL_CASE
        
        Case "Camel"
            GetNamingConv = VarNamingStyle.CAMEL_CASE
        
        Case "Snake_Case"
            GetNamingConv = VarNamingStyle.SNAKE_CASE
            
        Case Else
            Err.Raise 5, "modDependencyLambdaResult.GetNamingConv", "Invalid naming convention."
        
    End Select
    
End Function

' Get LET Var Prefix(_ or other) from Default OA Param
Public Function GetLetVarPrefix() As String
    GetLetVarPrefix = GetOAParamValue("FormulaFormat_LetVarPrefix", UNDER_SCORE)
End Function

Public Function GetAddPrefixOnParamValue() As Boolean
    GetAddPrefixOnParamValue = GetOAParamValue("FormulaFormat_AddPrefixOnParam", False)
End Function

Public Function GetIndentCharParamValue() As String
    GetIndentCharParamValue = GetOAParamValue("FormulaFormat_IndentChar", ONE_SPACE)
End Function

Public Function GetIndentSizeParamValue() As Integer
    Const DEFAULT_INDENT_SIZE As Long = 3
    GetIndentSizeParamValue = GetOAParamValue("FormulaFormat_IndentSize", DEFAULT_INDENT_SIZE)
End Function

Public Function GetMultilineParamValue() As Boolean
    GetMultilineParamValue = GetOAParamValue("FormulaFormat_Multiline", True)
End Function

Public Function GetOnlyWrapFunctionAfterNCharsParamValue() As Integer
    Const DEFAULT_LINE_CHAR_COUNT As Long = 80
    GetOnlyWrapFunctionAfterNCharsParamValue = GetOAParamValue("FormulaFormat_OnlyWrapFunctionAfterNChars", DEFAULT_LINE_CHAR_COUNT)
End Function

Public Function GetSpacesAfterArgumentSeparatorsParamValue() As Boolean
    GetSpacesAfterArgumentSeparatorsParamValue = GetOAParamValue("FormulaFormat_SpacesAfterArgumentSeparators", True)
End Function

Public Function GetSpacesAfterArrayColumnSeparatorsParamValue() As Boolean
    GetSpacesAfterArrayColumnSeparatorsParamValue = GetOAParamValue("FormulaFormat_SpacesAfterArrayColumnSeparators", True)
End Function

Public Function GetSpacesAfterArrayRowSeparatorsParamValue() As Boolean
    GetSpacesAfterArrayRowSeparatorsParamValue = GetOAParamValue("FormulaFormat_SpacesAfterArrayRowSeparators", True)
End Function

Public Function GetSpacesAroundInfixOperatorsParamValue() As Boolean
    GetSpacesAroundInfixOperatorsParamValue = GetOAParamValue("FormulaFormat_SpacesAroundInfixOperators", True)
End Function

Public Function GetBoModeParamValue() As Boolean
    GetBoModeParamValue = GetOAParamValue("FormulaFormat_BoMode", False)
End Function

Public Function GetUsedLambdas(ByVal Formula As String, ByVal AllLambdas As Collection _
                                                       , Optional ByVal IsR1C1 As Boolean = False) As Variant
    
    Dim UsedFunctions As Variant
    UsedFunctions = GetUsedFunctions(Formula, IsR1C1)
    
    If Not IsArrayAllocated(UsedFunctions) Then Exit Function
    
    Dim AllUsedLambdas As Collection
    Set AllUsedLambdas = New Collection
    
    Dim CurrentUsedFX As Variant
    For Each CurrentUsedFX In UsedFunctions
        If IsExistInCollection(AllLambdas, CStr(CurrentUsedFX)) Then
            AddToCollectionIfNotExist AllUsedLambdas, CurrentUsedFX, CStr(CurrentUsedFX)
        End If
    Next CurrentUsedFX
    
    GetUsedLambdas = CollectionToArray(AllUsedLambdas)
    
End Function

Public Function GetExcelLabsLambdas(ByVal FromBook As Workbook) As Collection
    
    #If DEVELOPMENT_MODE Then
        Dim AFE As OARobot.AFEProjectFactory
        Set AFE = New OARobot.AFEProjectFactory
        Dim Project As OARobot.AFEProject
    #Else
        Dim AFE As Object
        Set AFE = CreateObject("OARobot.AFEProjectFactory")
        Dim Project As Object
    #End If
    
    On Error Resume Next
    Set Project = AFE.FromWorkbook(FromBook)
    
    If Err.Number <> 0 Then
        Set GetExcelLabsLambdas = New Collection
        On Error GoTo 0
        Exit Function
    End If
    
    On Error GoTo 0
    
    Dim Lambdas As Variant
    Lambdas = Project.ProjectNames
    
    Dim LambdasColl As Collection
    Set LambdasColl = New Collection
    
    Dim CurrentLambdaName As Variant
    For Each CurrentLambdaName In Lambdas
        LambdasColl.Add CurrentLambdaName, CStr(CurrentLambdaName)
    Next CurrentLambdaName
    
    Set GetExcelLabsLambdas = LambdasColl
    
End Function

#If DEVELOPMENT_MODE Then
 Public Function GetFormulaProcessor() As OARobot.FormulaProcessing
#Else
Public Function GetFormulaProcessor() As Object
#End If

    #If DEVELOPMENT_MODE Then
        Set GetFormulaProcessor = New OARobot.FormulaProcessing
    #Else
        Set GetFormulaProcessor = CreateObject("OARobot.FormulaProcessing")
    #End If

End Function

#If DEVELOPMENT_MODE Then
Public Function GetFormulaFormatter() As OARobot.FormulaFormatter
#Else
Public Function GetFormulaFormatter() As Object
#End If

    #If DEVELOPMENT_MODE Then
        Set GetFormulaFormatter = New OARobot.FormulaFormatter
    #Else
        Set GetFormulaFormatter = CreateObject("OARobot.FormulaFormatter")
    #End If

End Function

#If DEVELOPMENT_MODE Then
Public Function GetFormulaParser() As OARobot.FormulaParser
#Else
Public Function GetFormulaParser() As Object
#End If

    #If DEVELOPMENT_MODE Then
        Set GetFormulaParser = New OARobot.FormulaParser
    #Else
        Set GetFormulaParser = CreateObject("OARobot.FormulaParser")
    #End If

End Function

#If DEVELOPMENT_MODE Then
Public Function GetExpressionReplacer() As OARobot.ExpressionReplacer
#Else
Public Function GetExpressionReplacer() As Object
#End If

    #If DEVELOPMENT_MODE Then
        Set GetExpressionReplacer = New OARobot.ExpressionReplacer
    #Else
        Set GetExpressionReplacer = CreateObject("OARobot.ExpressionReplacer")
    #End If

End Function

#If DEVELOPMENT_MODE Then
Public Function GetSheetNameRemover() As OARobot.SheetNameRemover
#Else
Public Function GetSheetNameRemover() As Object
#End If

    #If DEVELOPMENT_MODE Then
        Set GetSheetNameRemover = New OARobot.SheetNameRemover
    #Else
        Set GetSheetNameRemover = CreateObject("OARobot.SheetNameRemover")
    #End If

End Function

#If DEVELOPMENT_MODE Then
Public Function GetNamesOrFunctionsFilter() As OARobot.NamesOrFunctionsFilter
#Else
Public Function GetNamesOrFunctionsFilter() As Object
#End If

    #If DEVELOPMENT_MODE Then
        Set GetNamesOrFunctionsFilter = New OARobot.NamesOrFunctionsFilter
    #Else
        Set GetNamesOrFunctionsFilter = CreateObject("OARobot.NamesOrFunctionsFilter")
    #End If

End Function

#If DEVELOPMENT_MODE Then
Public Function GetLocale() As OARobot.FormulaLocaleInfo
#Else
Public Function GetLocale() As Object
#End If
    
    #If DEVELOPMENT_MODE Then
        Dim LocaleFactory As New OARobot.FormulaLocaleInfoFactory
        Set GetLocale = LocaleFactory.CreateFromExcel(Application)
    #Else
        Set GetLocale = CreateObject("OARobot.FormulaLocaleInfoFactory").CreateFromExcel(Application)
    #End If

End Function

#If DEVELOPMENT_MODE Then
Public Function GetScope(Optional ByVal ForBook As Workbook) As OARobot.FormulaScopeInfo
#Else
Public Function GetScope(Optional ByVal ForBook As Workbook) As Object
#End If
    
    If ForBook Is Nothing Then Set ForBook = ActiveWorkbook
    
    #If DEVELOPMENT_MODE Then
        Dim ScopeFactory As New OARobot.FormulaScopeFactory
        Set GetScope = ScopeFactory.CreateWorkbook(ForBook.name)
    #Else
        Set GetScope = CreateObject("OARobot.FormulaScopeFactory").CreateWorkbook(ForBook.name)
    #End If
    
End Function

#If DEVELOPMENT_MODE Then
Public Function GetNames(Optional ByVal ForBook As Workbook) As OARobot.XLDefinedNames
#Else
Public Function GetNames(Optional ByVal ForBook As Workbook) As Object
#End If
    
    If ForBook Is Nothing Then Set ForBook = ActiveWorkbook
    
    #If DEVELOPMENT_MODE Then
        Dim NamesFactory As New OARobot.DefinedNamesFactory
        Set GetNames = NamesFactory.Create(ForBook)
    #Else
        Set GetNames = CreateObject("OARobot.DefinedNamesFactory").Create(ForBook)
    #End If
    
End Function

#If DEVELOPMENT_MODE Then
Public Function ParseFormula(ByVal Formula As String _
                             , Optional ByVal ForBook As Workbook _
                              , Optional ByVal IsR1C1 As Boolean = False) As OARobot.FormulaParseResult
#Else
Public Function ParseFormula(ByVal Formula As String _
                             , Optional ByVal ForBook As Workbook _
                              , Optional ByVal IsR1C1 As Boolean = False) As Object
#End If
    
Static ScopeBookName As String
If ForBook Is Nothing Then Set ForBook = ActiveWorkbook
    
#If DEVELOPMENT_MODE Then
    Static Scope As OARobot.FormulaScopeInfo
    Static DefinedNames As OARobot.XLDefinedNames
    Dim Parser As OARobot.FormulaParser
    Dim LocaleFactory As FormulaLocaleInfoFactory
    Set LocaleFactory = New FormulaLocaleInfoFactory
#Else
    Static Scope As Object
    Static DefinedNames As Object
    Dim Parser As Object
    Dim LocaleFactory As Object
    Set LocaleFactory = CreateObject("FormulaLocaleInfoFactory")
#End If
    
If ScopeBookName <> ForBook.name Or ScopeBookName = vbNullString Then
    Set Scope = GetScope(ForBook)
    Set DefinedNames = GetNames(ForBook)
    ScopeBookName = ForBook.name
End If
    
Set Parser = GetFormulaParser()
    
If ForBook Is Nothing Then Set ForBook = ActiveWorkbook
    
Set ParseFormula = Parser.Parse(Formula, IsR1C1, LocaleFactory.EN_US, Scope, DefinedNames)
    
End Function




