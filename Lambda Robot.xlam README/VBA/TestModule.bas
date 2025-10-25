Attribute VB_Name = "TestModule"
'@IgnoreModule UndeclaredVariable, ImplicitActiveSheetReference, UnrecognizedAnnotation, ProcedureNotUsed
'@Folder "Lambda.Editor.Driver"
Option Explicit
Option Private Module

Public Sub DeleteColumnsToRight(ByVal GivenRange As Range)
    
    Dim FromSheet As Worksheet
    Set FromSheet = GivenRange.Worksheet
    FromSheet.Range(GivenRange.Cells(1, 1), FromSheet.Cells(GivenRange.Row, FromSheet.Columns.Count)).EntireColumn.Delete

End Sub

Public Sub FindReferenceInChartSeries(ByVal GivenRange As Range, ByVal PutOnCell As Range)
    
    On Error GoTo ErrorHandler
    RangeDependencyInChart.SendDataToSheet GivenRange, PutOnCell
    Exit Sub
    
ErrorHandler:
    
    Dim ErrorNumber As Long
    ErrorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description
    If ErrorNumber <> 0 Then
        Err.Raise ErrorNumber, Err.Source, ErrorDescription
        ' This is only for debugging purpose.
        Resume
    End If
    
End Sub

Public Sub SpecialCellsGroup()
    
    Dim CurrentCell As Range
    Dim TestRange As Range
    Set TestRange = ActiveWorkbook.Worksheets("LETStep_FX Named Range").Range("G4:L12")
    For Each CurrentCell In TestRange.Cells
        If CurrentCell.HasFormula Then
            Logger.Log DEBUG_LOG, CurrentCell.Address & " has a formula."
        End If
    Next CurrentCell
    
    Dim RangeHavingFormulas As Range
    On Error Resume Next
    Set RangeHavingFormulas = FilterUsingSpecialCells(TestRange, xlCellTypeFormulas)
    Logger.Log DEBUG_LOG, "Having Formula Cell Address : " & RangeHavingFormulas.Address
    On Error GoTo 0
    
End Sub

Public Sub PrintDirectPrecedentsOfActiveCell()
    
    Dim Dependency As Variant
    Dependency = GetDirectPrecedents(GetCellFormula(ActiveCell), ActiveCell.Worksheet)
    
    ' Ensure the Dependency is an array
    If Not IsArray(Dependency) Then Dependency = Array(Dependency)
    
    Dim CurrentDependency As Variant
    For Each CurrentDependency In Dependency
        If CurrentDependency <> vbNullString Then
            Debug.Print CurrentDependency
        End If
    Next CurrentDependency
    
End Sub

Public Sub ParserFailedToParse()
    
    Dim TestFormula As String
    TestFormula = "=LAMBDA(List,TEXTJOIN("","", TRUE(), List))"
    Dim Parser As Object
    Set Parser = CreateObject("OARobot.FormulaParser")
    Dim ParsedFormulaResult As Object
    
    Set ParsedFormulaResult = Parser.Parse(TestFormula)
    
    If ParsedFormulaResult.ParseSuccess Then
        MsgBox "Parsed successfully.", vbOKOnly, "Formula Parsing"
    Else
        MsgBox "Parser failed to parse." & vbNewLine & "Formula: " & TestFormula, vbOKOnly, "Formula Parsing"
    End If
    
End Sub

'Private Sub JSONTester()
'
'    Dim Map As Scripting.Dictionary
'    Set Map = New Scripting.Dictionary
'
'    With Map
'        .Add "Key1", "Value1"
'        .Add "Key2", "Value"
'    End With
'
'    Debug.Print JsonConverter.ConvertToJson(Map, 2)
'
'End Sub

Private Sub TestNamedRangeEvaluation()
    
    Dim V As Variant
    V = Evaluate("ShortMonthOfTheYear")
    
End Sub

Private Sub TestGenLambda()
    GenerateLambdaStatement ActiveCell, ActiveCell
End Sub

