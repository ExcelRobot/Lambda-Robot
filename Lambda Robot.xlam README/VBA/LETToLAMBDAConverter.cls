VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LETToLAMBDAConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredID

Option Explicit

Private LetParts As Variant
Private FirstColIndex As Long
Private LambdaArgumentPart As String
Private LambdaInvocationPart As String
Private LetOfFinalLambda As String
Private AppFunction As WorksheetFunction
Private CountOfNonInputStep As Long
Private IsFirstLetVarIsInputCell As Boolean

Private Sub Initialize(ByVal LetFormula As String)
    
    LetParts = GetDependencyFunctionResult(LetFormula, LET_PARTS)
    
    If Not IsArray(LetParts) Then Exit Sub
    
    ' Retrieve the index of the first column of the "LetParts" array
    FirstColIndex = LBound(LetParts, 2)
    
    ' Initialize the lambda argument part of the final string to be built
    LambdaArgumentPart = EQUAL_SIGN & LAMBDA_FX_NAME & FIRST_PARENTHESIS_OPEN
    
    ' Initialize the lambda invocation part of the final string to be built
    LambdaInvocationPart = FIRST_PARENTHESIS_OPEN
    ' Initialize the LET statement of the final lambda to be built
    
    LetOfFinalLambda = LET_AND_OPEN_PAREN
    
    ' Create a WorksheetFunction object to be able to use Excel worksheet functions in VBA
    Set AppFunction = Application.WorksheetFunction
    
    Dim Temp As String
    Temp = LetParts(FirstColIndex, FirstColIndex + modSharedConstant.LET_PARTS_VALUE_COL_INDEX - 1)
    IsFirstLetVarIsInputCell = IsRangeAddress(Temp)
    
    ' If the first LET variable is a reference to a cell, determine if it's an input cell
    If IsFirstLetVarIsInputCell Then IsFirstLetVarIsInputCell = IsInputCell(RangeResolver.GetRange(Temp), Nothing)
    
End Sub

Public Function ConvertLetToLambda(ByVal LetFormula As String) As String

    ' Check if the formula in the cell starts with "LAMBDA", if so, no conversion is needed
    If IsLambdaFunction(LetFormula) Then
        ConvertLetToLambda = LetFormula
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modLambdaBuilder.ConvertLetToLambda"
        Exit Function
    End If
    
    Initialize LetFormula
    
    ' If the returned result is not an array, that means no LET function was found in the cell, hence a message box pops up to notify the user
    If Not IsArray(LetParts) Then
        MsgBox "Unable to convert " & LET_FX_NAME & " to " & LAMBDA_FX_NAME _
               & ".  No " & LET_FX_NAME & " function found.", vbExclamation + vbOKOnly, APP_NAME
        Exit Function
    End If
    
    ' Extract the address of the first LET variable and check if it is an input cell
    Dim CurrentRowIndex As Long
    
    ' Iterate through the "LetParts" array
    For CurrentRowIndex = LBound(LetParts, 1) To UBound(LetParts, 1) - 1
        UpdateForCurrentRow CurrentRowIndex, LetFormula
    Next CurrentRowIndex
    
    LetOfFinalLambda = LetOfFinalLambda _
                       & AppFunction.Clean(AppFunction.Trim(LetParts(UBound(LetParts, 1), FirstColIndex)))

    ' Remove trailing comma from the LambdaInvocationPart
    LambdaInvocationPart = modUtility.RemoveEndingText(LambdaInvocationPart, LIST_SEPARATOR)
    
    ' If only one non-input step was encountered, remove the opening parenthesis from LetOfFinalLambda
    If CountOfNonInputStep = 0 Then
        LetOfFinalLambda = VBA.Mid$(LetOfFinalLambda, Len(LET_AND_OPEN_PAREN) + 1)
        ' Otherwise, append closing parenthesis to LetOfFinalLambda
    Else
        LetOfFinalLambda = LetOfFinalLambda & FIRST_PARENTHESIS_CLOSE
    End If

    LambdaInvocationPart = LambdaInvocationPart & FIRST_PARENTHESIS_CLOSE

    ' Form the final lambda function by concatenating LambdaArgumentPart, LetOfFinalLambda and LambdaInvocationPart
    ConvertLetToLambda = LambdaArgumentPart & LetOfFinalLambda & Chr$(10) _
                         & FIRST_PARENTHESIS_CLOSE & LambdaInvocationPart

    ' Replace new line character with line feed in the final lambda function so that we can use in cell.
    ConvertLetToLambda = FormatFormula(ReplaceNewlineWithChar10(ConvertLetToLambda))

    Logger.Log DEBUG_LOG, "Result Lambda : " & ConvertLetToLambda
    
End Function

Private Sub UpdateForCurrentRow(ByVal CurrentRowIndex As Long, ByVal LetFormula As String)
    
    Dim VarName As String
    Dim CellAddress As String
    ' Clean and trim the current variable name
    VarName = VBA.LTrim$(AppFunction.Clean(AppFunction.Trim(LetParts(CurrentRowIndex, FirstColIndex))))

    ' Trim the current cell address
    CellAddress = Trim$(LetParts(CurrentRowIndex _
                                 , FirstColIndex + modSharedConstant.LET_PARTS_VALUE_COL_INDEX - 1))
        
    ' Check if the cell address is a range address
    If modUtility.IsRangeAddress(CellAddress) Then
            
        If IsFirstLetVarIsInputCell Then
            ' If the current cell is an input cell, add it to the lambda argument and invocation part
            If IsInputCell(RangeResolver.GetRange(CellAddress), Nothing) Then
                LambdaArgumentPart = LambdaArgumentPart & UpdateForOptionalArgument(LetFormula, VarName)
                LambdaInvocationPart = LambdaInvocationPart & CellAddress & LIST_SEPARATOR
                ' If it's not an input cell, increment the non-input step counter and add the LET variable to the final lambda
            Else
                CountOfNonInputStep = CountOfNonInputStep + 1
                LetOfFinalLambda = LetOfFinalLambda & VarName & LIST_SEPARATOR & CellAddress & LIST_SEPARATOR
            End If
            
        Else
            ' If the first LET variable was not an input cell
            ' then consider all cell address as input cell. And add it to the lambda argument and invocation part
            LambdaArgumentPart = LambdaArgumentPart & UpdateForOptionalArgument(LetFormula, VarName)
            LambdaInvocationPart = LambdaInvocationPart & CellAddress & LIST_SEPARATOR
        End If
    ElseIf modUtility.IsOptionalArgument(LetFormula, VarName) Then
        ' If the variable is an optional argument, create it and add it to the lambda argument and invocation part
        LambdaArgumentPart = LambdaArgumentPart & CreateOptionalArgument(VarName)
        LambdaInvocationPart = LambdaInvocationPart & CellAddress & LIST_SEPARATOR
    Else
        ' If it's not an optional argument or a range address, increment the non-input step counter and add the LET variable to the final lambda
        CountOfNonInputStep = CountOfNonInputStep + 1
        LetOfFinalLambda = LetOfFinalLambda & VarName & LIST_SEPARATOR & CellAddress & LIST_SEPARATOR
    End If
    
End Sub

Private Function UpdateForOptionalArgument(ByVal LetFormula As String, ByVal VarName As String) As String
    
    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.UpdateForOptionalArgument"
    If modUtility.IsOptionalArgument(LetFormula, VarName) Then
        ' If it is an optional argument, create an optional argument
        UpdateForOptionalArgument = CreateOptionalArgument(VarName)
    Else
        ' If not, return a cleaned and trimmed VarName, followed by a comma and a space
        UpdateForOptionalArgument = AppFunction.Trim(AppFunction.Clean(VarName)) _
                                    & LIST_SEPARATOR
    End If

    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.UpdateForOptionalArgument"
    
End Function

Private Function CreateOptionalArgument(ByVal VarName As String) As String
    
    Logger.Log TRACE_LOG, "Enter modLambdaBuilder.CreateOptionalArgument"
    
    ' Establish a reference to Application.WorksheetFunction
    Dim AppFunction As WorksheetFunction
    Set AppFunction = Application.WorksheetFunction

    ' Create an optional argument by adding parentheses around a cleaned and trimmed VarName,
    ' followed by a comma and a space
    CreateOptionalArgument = LEFT_BRACKET & _
                             AppFunction.Trim(AppFunction.Clean(VarName)) & _
                             RIGHT_BRACKET & LIST_SEPARATOR
    Logger.Log TRACE_LOG, "Exit modLambdaBuilder.CreateOptionalArgument"
                             
End Function

