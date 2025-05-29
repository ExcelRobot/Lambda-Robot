Attribute VB_Name = "MultiColumnLookupLambda"
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed
'@Folder "MultiColumnLookup"
Option Explicit
Option Private Module

' --------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Generate Multi Column LookUp Lambda
' Description:            Generate multi-column lookup lambda from ActiveCell. It will return nth column value by using n-1 parameter and filter by them.
' Macro Expression:       MultiColumnLookupLambda.GenerateMultiColumnLookUpLambda([ActiveCell])
' Generated:              07/02/2023 01:21 PM
' ----------------------------------------------------------------------------------------------------

Public Sub GenerateMultiColumnLookUpLambda(ByVal FromRange As Range, Optional ByVal IsUndo As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter MultiColumnLookupLambda.GenerateMultiColumnLookUpLambda"
    Const METHOD_NAME As String = "GenerateMultiColumnLookUpLambda"
    Context.ExtractContextFromCell FromRange, METHOD_NAME
    ' Generate the multi-column lookup lambda formula for the specified FromRange.
    
    ' Static variables to store old formula and the cell containing the formula
    Static OldFormula As String
    Static FormulaInCell As Range
    
    ' If IsUndo is True, restore the old formula and exit
    If IsUndo Then
        FormulaInCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
        AutofitFormulaBar FormulaInCell
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword MultiColumnLookupLambda.GenerateMultiColumnLookUpLambda"
        GoTo ExitMethod
    End If
    
    Dim LambdaFormula As String
    LambdaFormula = GetMultiColumnLookUpLambda(FromRange)
    If LambdaFormula = vbNullString Then GoTo ExitMethod
    
    ' Set the formula in the specified cell
    Set FormulaInCell = FromRange.SpillParent
    OldFormula = GetCellFormula(FromRange)
    FormulaInCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(LambdaFormula)
    AutofitFormulaBar FormulaInCell
    AssingOnUndo "GenerateMultiColumnLookUpLambda"
    Logger.Log TRACE_LOG, "Exit MultiColumnLookupLambda.GenerateMultiColumnLookUpLambda"
    
ExitMethod:
    Context.ClearContext METHOD_NAME
    Exit Sub
    
End Sub

Public Sub GenerateMultiColumnLookUpLambda_Undo()
    Logger.Log TRACE_LOG, "Enter MultiColumnLookupLambda.GenerateMultiColumnLookUpLambda_Undo"
    ' Undo the action of generating the multi-column lookup lambda formula.
    GenerateMultiColumnLookUpLambda Nothing, True
    Logger.Log TRACE_LOG, "Exit MultiColumnLookupLambda.GenerateMultiColumnLookUpLambda_Undo"
End Sub

Private Function GetMultiColumnLookUpLambda(ByVal FromRange As Range) As String
    
    Logger.Log TRACE_LOG, "Enter MultiColumnLookupLambda.GetMultiColumnLookUpLambda"
    ' Get the multi-column lookup lambda formula for the specified FromRange.
    
    ' Check if the FromRange is a spill range
    If Not FromRange.HasSpill Then Exit Function
    
    ' Create a TextConcatenator to build the filter part of the lambda formula
    Dim FilterPartFormula As TextConcatenator
    Set FilterPartFormula = TextConcatenator.Create(LET_FX_NAME & FIRST_PARENTHESIS_OPEN & vbNewLine)
    With FilterPartFormula
        ' Get the table or named range ref. Formula need to be like =TableOrNamedRangeOrSpillRangeRef
        .Concatenate THREE_SPACE & "_Table" & LIST_SEPARATOR _
                     & Mid$(GetCellFormula(FromRange.SpillParent), 2) _
                     & LIST_SEPARATOR & vbNewLine
                     
        .Concatenate THREE_SPACE & "_LastColumnData" & LIST_SEPARATOR & ONE_SPACE _
                     & CHOOSECOLS_FX_NAME & "(_Table" & LIST_SEPARATOR _
                     & COLUMNS_FX_NAME & "(_Table))" & LIST_SEPARATOR & vbNewLine
    End With

    ' Create the Lambda parameter part of the formula
    Dim LambdaParamPartFormula As String
    LambdaParamPartFormula = EQUAL_SIGN & LAMBDA_FX_NAME & FIRST_PARENTHESIS_OPEN
    
    ' Get the header row of the spilling range
    Dim HeaderRow As Range
    Set HeaderRow = FromRange.SpillParent.SpillingToRange.Resize(1, FromRange.SpillParent.SpillingToRange.Columns.Count)
    
    ' Create the invocation part of the formula
    Dim InvocationPart As String
    InvocationPart = FIRST_PARENTHESIS_OPEN
    
    Dim IncludeAllColumnsStepPart As String
    IncludeAllColumnsStepPart = OR_FX_NAME & FIRST_PARENTHESIS_OPEN
    Dim FilterCriteriaStepPart As String
    
    
    ' Loop through each column in the header row (except the last column)
    Dim ColIndex As Long
    For ColIndex = 1 To HeaderRow.Columns.Count - 1
        
        Dim CurrentCell As Range
        Set CurrentCell = HeaderRow.Columns(ColIndex)
        
        ' Get the header name and make it a valid variable name
        Dim CurrentHeader As String
        CurrentHeader = CurrentCell.Value
        CurrentHeader = MakeValidLetVarName(CurrentHeader, GetNamingConv(False))
        LambdaParamPartFormula = LambdaParamPartFormula & LEFT_BRACKET _
                                 & CurrentHeader & RIGHT_BRACKET & LIST_SEPARATOR
        
        IncludeAllColumnsStepPart = IncludeAllColumnsStepPart & ISOMITTED_FX_NAME _
                                    & FIRST_PARENTHESIS_OPEN & CurrentHeader _
                                    & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR
        
        ' Build the filter part for the current column
        FilterCriteriaStepPart = FilterCriteriaStepPart & IF_FX_NAME _
                                 & FIRST_PARENTHESIS_OPEN & ISOMITTED_FX_NAME _
                                 & FIRST_PARENTHESIS_OPEN & CurrentHeader & FIRST_PARENTHESIS_CLOSE _
                                 & LIST_SEPARATOR & " 1" & LIST_SEPARATOR _
                                 & ONE_SPACE & FIRST_PARENTHESIS_OPEN & CHOOSECOLS_FX_NAME & "(_Table" _
                                 & LIST_SEPARATOR & ColIndex & ")=" & CurrentHeader & "))"
                                 
        If ColIndex <> HeaderRow.Columns.Count - 1 Then FilterCriteriaStepPart = FilterCriteriaStepPart & "*"
        
        ' Build the invocation part for the current column
        InvocationPart = InvocationPart & GetValueForInvocation(CurrentCell.Offset(1, 0)) & LIST_SEPARATOR
        
    Next ColIndex
    
    ' Remove the trailing comma and complete the invocation part
    InvocationPart = Text.RemoveFromEnd(InvocationPart, 1) & FIRST_PARENTHESIS_CLOSE
    IncludeAllColumnsStepPart = Text.RemoveFromEndIfPresent(IncludeAllColumnsStepPart, LIST_SEPARATOR) & FIRST_PARENTHESIS_CLOSE
    
    With FilterPartFormula
        .Concatenate THREE_SPACE & "_IncludeAllColumns" & LIST_SEPARATOR & ONE_SPACE _
                     & IncludeAllColumnsStepPart & LIST_SEPARATOR & vbNewLine
                     
        .Concatenate THREE_SPACE & "_FilterCriteria" & LIST_SEPARATOR & ONE_SPACE _
                     & FilterCriteriaStepPart & LIST_SEPARATOR & vbNewLine
                     
        .Concatenate THREE_SPACE & "_Result" & LIST_SEPARATOR & ONE_SPACE _
                     & FILTER_FX_NAME & FIRST_PARENTHESIS_OPEN & IF_FX_NAME _
                     & "(_IncludeAllColumns" & LIST_SEPARATOR _
                     & " _Table" & LIST_SEPARATOR & " _LastColumnData)" _
                     & LIST_SEPARATOR & " _FilterCriteria)" _
                     & LIST_SEPARATOR & vbNewLine
                     
        .Concatenate THREE_SPACE & "_Result" & vbNewLine
        .Concatenate THREE_SPACE & FIRST_PARENTHESIS_CLOSE
    End With
    
    Dim FullFormula As String
    FullFormula = LambdaParamPartFormula & FilterPartFormula.Text & vbNewLine & FIRST_PARENTHESIS_CLOSE & InvocationPart
    FullFormula = ReplaceNewlineWithChar10(FullFormula)
    GetMultiColumnLookUpLambda = FullFormula
    Logger.Log TRACE_LOG, "Exit MultiColumnLookupLambda.GetMultiColumnLookUpLambda"

End Function

Public Function GetValueForInvocation(ByVal FromCell As Range) As String

    Logger.Log TRACE_LOG, "Enter MultiColumnLookupLambda.GetValueForInvocation"
    ' Get the value of a cell as a string that can be used in the formula invocation.
    Select Case VarType(FromCell.Value)
        Case vbEmpty
            GetValueForInvocation = 0
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency _
                                                   , vbDate, vbBoolean, vbDecimal, vbByte, vbLongLong, vbError
            GetValueForInvocation = FromCell.Value
        Case vbString
            GetValueForInvocation = """" & VBA.Replace(FromCell.Value, """", """""") & """"
    End Select
    Logger.Log TRACE_LOG, "Exit MultiColumnLookupLambda.GetValueForInvocation"
    
End Function


