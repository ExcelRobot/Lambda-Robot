VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaFormatConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("LambdaRobot")
'@PredeclaredId
Option Explicit

Private Const DEFAULT_INDENT_SIZE As Long = 3
Private Const DEFAULT_LINE_CHAR_COUNT As Long = 80

Public AddPrefixOnParam As Boolean
Public BoMode As Boolean
Public IndentChar As String
Public IndentSize As Integer
Public LambdaParamStyle As String
Public LetVarPrefix As String
Public Multiline As Boolean
Public OnlyWrapFunctionAfterNChars As Integer
Public SpacesAfterArgumentSeparators As Boolean
Public SpacesAfterArrayColumnSeparators As Boolean
Public SpacesAfterArrayRowSeparators As Boolean
Public SpacesAroundInfixOperators As Boolean
Public VariableStyle As String
Public IncludeResultStep As Boolean
Public ClearInputFormatting As Boolean

Private OARobotAddIn As Object

Public Sub Initialize()
    
'    Dim StartTime As Double
'    StartTime = Timer()
    
    Set OARobotAddIn = CreateObject("OARobot.ExcelAddin")
    
    BoMode = GetOAParamValue("FormulaFormat_BoMode", False)
    If BoMode Then
        LambdaParamStyle = "Snake_Case"
        VariableStyle = "Camel"
        LetVarPrefix = vbNullString
        AddPrefixOnParam = False
        IncludeResultStep = False
    Else
        LambdaParamStyle = GetOAParamValue("FormulaFormat_LambdaParamStyle", "Snake_Case")
        VariableStyle = GetOAParamValue("FormulaFormat_VariableStyle", "Pascal")
        LetVarPrefix = GetOAParamValue("FormulaFormat_LetVarPrefix", UNDER_SCORE)
        AddPrefixOnParam = GetOAParamValue("FormulaFormat_AddPrefixOnParam", False)
        IndentChar = GetOAParamValue("FormulaFormat_IndentChar", ONE_SPACE)
        IndentSize = GetOAParamValue("FormulaFormat_IndentSize", DEFAULT_INDENT_SIZE)
        Multiline = GetOAParamValue("FormulaFormat_Multiline", True)
        OnlyWrapFunctionAfterNChars = GetOAParamValue("FormulaFormat_OnlyWrapFunctionAfterNChars", DEFAULT_LINE_CHAR_COUNT)
        SpacesAfterArgumentSeparators = GetOAParamValue("FormulaFormat_SpacesAfterArgumentSeparators", True)
        SpacesAfterArrayColumnSeparators = GetOAParamValue("FormulaFormat_SpacesAfterArrayColumnSeparators", True)
        SpacesAfterArrayRowSeparators = GetOAParamValue("FormulaFormat_SpacesAfterArrayRowSeparators", True)
        SpacesAroundInfixOperators = GetOAParamValue("FormulaFormat_SpacesAroundInfixOperators", True)
        IncludeResultStep = GetOAParamValue("FormulaFormat_IncludeResultStep", True)
    End If
    
    ClearInputFormatting = GetOAParamValue("FormulaFormat_ClearInputFormatting", False)
    
    Set OARobotAddIn = Nothing
    
'    Debug.Print "Total time to retrieve formula format config: " & Timer() - StartTime
    
End Sub

Private Function GetOAParamValue(ParamName As String, DefaultValue As Variant) As Variant
    
    On Error GoTo ReturnDefaultValue
    GetOAParamValue = OARobotAddIn.GetParamValueByName(ParamName)
    Exit Function
    
ReturnDefaultValue:
    GetOAParamValue = DefaultValue
    
End Function

