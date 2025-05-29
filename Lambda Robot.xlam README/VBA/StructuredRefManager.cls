VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StructuredRefManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TStructuredRefManager
    QualifiedCellRefVsStructuredFormulaMap As Collection
    FormulaStartCellInSheet As Worksheet
    AllNamedRange As Dictionary
    CellRefToLetStepLambdaRefMap As Collection
End Type

Private this As TStructuredRefManager

Public Property Get AllNamedRange() As Dictionary
    Set AllNamedRange = this.AllNamedRange
End Property

Public Property Set AllNamedRange(ByVal RHS As Dictionary)
    Set this.AllNamedRange = RHS
End Property

Public Property Get FormulaStartCellInSheet() As Worksheet
    Set FormulaStartCellInSheet = this.FormulaStartCellInSheet
End Property

Public Property Set FormulaStartCellInSheet(ByVal RHS As Worksheet)
    Set this.FormulaStartCellInSheet = RHS
End Property

Public Function GetStructuredFormula(ByVal ForCell As Range) As String
    
    Logger.Log TRACE_LOG, "Enter StructuredRefManager.GetStructuredFormula"
    Dim FullyQualifiedCellRef As String
    FullyQualifiedCellRef = ConvertToFullyQualifiedCellRef(ForCell.Cells(1))
    
    If Not IsExistInCollection(this.QualifiedCellRefVsStructuredFormulaMap, FullyQualifiedCellRef) Then
        ConvertRangeRefToStructuralRef ForCell
    End If
    
    GetStructuredFormula = this.QualifiedCellRefVsStructuredFormulaMap.Item(FullyQualifiedCellRef)
    Logger.Log TRACE_LOG, "Exit StructuredRefManager.GetStructuredFormula"
    
End Function

Public Function GetLetStepLambdaRefs(ByVal ForCell As Range) As Dictionary
    
    Logger.Log TRACE_LOG, "Enter StructuredRefManager.GetLetStepLambdaRefs"
    Dim FullyQualifiedCellRef As String
    FullyQualifiedCellRef = ConvertToFullyQualifiedCellRef(ForCell.Cells(1))
    
    If Not IsExistInCollection(this.QualifiedCellRefVsStructuredFormulaMap, FullyQualifiedCellRef) Then
        ConvertRangeRefToStructuralRef ForCell
    End If
    
    Set GetLetStepLambdaRefs = this.CellRefToLetStepLambdaRefMap.Item(FullyQualifiedCellRef)
    Logger.Log TRACE_LOG, "Exit StructuredRefManager.GetLetStepLambdaRefs"
    
End Function

Public Sub ConvertRangeRefToStructuralRef(ByVal ForCell As Range)
    
    Logger.Log TRACE_LOG, "Enter StructuredRefManager.ConvertRangeRefToStructuralRef"
    ' Converts the cell's formula from a range reference to a structural formula reference.
    ' ForCell: The cell for which to convert the formula.
    
    Dim CellRef As String
    CellRef = modUtility.ConvertToFullyQualifiedCellRef(ForCell.Cells(1))
    
    ' Debug.Assert CellRef <> "'[TextOffset lambda testing.xlsm]Sheet1'!$C$15"
    
    ' If the structural formula is not already stored in the map, calculate the structural formula.
    If Not modUtility.IsExistInCollection(this.QualifiedCellRefVsStructuredFormulaMap, CellRef) Then
        Dim Formula As String
        '        Debug.Assert ForCell.Address <> "$J$3"
        Formula = GetConvertedStructuredFormula(ForCell, True)
        Formula = modUtility.GetLambdaDefIfLETStepRefCell(ForCell, Formula, this.FormulaStartCellInSheet)
        UpdateLETStepLambdaRefCell Formula, CellRef
        this.QualifiedCellRefVsStructuredFormulaMap.Add Formula, CellRef
    End If
    Logger.Log TRACE_LOG, "Exit StructuredRefManager.ConvertRangeRefToStructuralRef"
    
End Sub

'@Impure: This will modify LETStepLambdaRefs Dictionary and send back.
Private Sub UpdateLETStepLambdaRefCell(ByVal Formula As String, ByVal CellRef As String)
    
    Logger.Log TRACE_LOG, "Enter StructuredRefManager.UpdateLETStepLambdaRefCell"
    ' Replaces LET-Step-Lambda references in the formula with corresponding cell references.
    ' Formula: The formula string to process.
    ' CellRef: The cell reference of the formula.

    Dim LetStepLambdaRefs As Scripting.Dictionary
    Set LetStepLambdaRefs = New Scripting.Dictionary
    Dim FunctionName As String
    Dim Suffix As String

    ' Loop to find and replace LET-Step-Lambda references in the formula.
    Do While Text.Contains(Formula, LETSTEP_UNDERSCORE_PREFIX)
        FunctionName = Text.AfterDelimiter(Formula, LETSTEP_UNDERSCORE_PREFIX, 1, FROM_START)
        Dim Index As Long
        For Index = 1 To Len(FunctionName)
            Dim CurrentChar As String
            CurrentChar = Mid$(FunctionName, Index, 1)

            Select Case CurrentChar
                    ' Lambda can be invoked("("), can be passed as a param(",", ")")
                    ' If we pass as a param, then it can be at the end (")") else at any other position)
                Case FIRST_PARENTHESIS_OPEN, LIST_SEPARATOR, FIRST_PARENTHESIS_CLOSE
                    FunctionName = Mid$(FunctionName, 1, Index - 1)
                    Suffix = Right$(FunctionName, Index - 1 - Len(RTrim$(FunctionName))) & CurrentChar
                    FunctionName = RTrim$(FunctionName)
                    FunctionName = LETSTEP_UNDERSCORE_PREFIX & FunctionName
                    Exit For
            End Select
        Next Index

        ' Check if the LET-Step-Lambda reference exists in the AllNamedRange collection.
        Dim QualifiedSheetName As String
        QualifiedSheetName = GetSheetRefForRangeReference(this.FormulaStartCellInSheet.Name, False)
        ' Check for local named range first.
        Dim LetStepRefNamedRange As String
        Dim LetStepRefCell As Range
        If this.AllNamedRange.Exists(QualifiedSheetName & FunctionName) Then
            LetStepRefNamedRange = QualifiedSheetName & VBA.Replace(FunctionName _
                                                                    , LETSTEP_UNDERSCORE_PREFIX _
                                                                     , LETSTEPREF_UNDERSCORE_PREFIX)
            Set LetStepRefCell = this.AllNamedRange.Item(LetStepRefNamedRange).RefersToRange
            Formula = VBA.Replace(Formula, FunctionName & Suffix, LetStepRefCell.Address & Suffix)
            LetStepLambdaRefs.Item(GetRangeRefWithSheetName(LetStepRefCell, False)) = FunctionName
        
            ' Check for global named range as well
        ElseIf this.AllNamedRange.Exists(FunctionName) Then
            LetStepRefNamedRange = VBA.Replace(FunctionName _
                                               , LETSTEP_UNDERSCORE_PREFIX _
                                                , LETSTEPREF_UNDERSCORE_PREFIX)
            Set LetStepRefCell = this.AllNamedRange.Item(LetStepRefNamedRange).RefersToRange
            Formula = VBA.Replace(Formula, FunctionName & Suffix, LetStepRefCell.Address & Suffix)
            LetStepLambdaRefs.Item(GetRangeRefWithSheetName(LetStepRefCell, False)) = FunctionName
        End If
    Loop

    this.CellRefToLetStepLambdaRefMap.Add LetStepLambdaRefs, CellRef
    Logger.Log TRACE_LOG, "Exit StructuredRefManager.UpdateLETStepLambdaRefCell"
    
End Sub

'@Exposed
'@Impure
Public Function FindStructuralFormulaText(ByVal ForCell As Range, ByVal FormulaText As String) As String
    
    Logger.Log TRACE_LOG, "Enter StructuredRefManager.FindStructuralFormulaText"
    ' Finds the structural formula text for the given cell.
    ' If the cell's structural formula is already stored in the map, it is retrieved; otherwise, the cell's structural formula is calculated.
    ' ForCell: The cell for which the structural formula text is to be found.
    ' FormulaText: The formula text of the cell.
    
    Dim Result As String
    Result = GetStructuredFormula(ForCell)
    
    ' If the cell has more than one cell and the formula text is not prefixed with '=', set it to an empty string.
    If ForCell.Count > 1 And Not Text.IsStartsWith(FormulaText, EQUAL_SIGN) Then
        Result = vbNullString
    End If

    FindStructuralFormulaText = Result
    Logger.Log TRACE_LOG, "Exit StructuredRefManager.FindStructuralFormulaText"
    
End Function

Private Sub Class_Initialize()
    
    Logger.Log TRACE_LOG, "Enter StructuredRefManager.Class_Initialize"
    Set this.QualifiedCellRefVsStructuredFormulaMap = New Collection
    Set this.CellRefToLetStepLambdaRefMap = New Collection
    Logger.Log TRACE_LOG, "Exit StructuredRefManager.Class_Initialize"
    
End Sub

Private Sub Class_Terminate()
    
    Logger.Log TRACE_LOG, "Enter StructuredRefManager.Class_Terminate"
    Set this.QualifiedCellRefVsStructuredFormulaMap = Nothing
    Set this.CellRefToLetStepLambdaRefMap = Nothing
    Logger.Log TRACE_LOG, "Exit StructuredRefManager.Class_Terminate"
    
End Sub


