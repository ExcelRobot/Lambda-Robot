Attribute VB_Name = "modCombineArray"
Option Explicit
Option Private Module

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Paste Combine Arrays
' Description:            Paste a dynamic array referencing the copied cells by combining any dynamic arrays automatically.
' Macro Expression:       modCombineArray.PasteCombineArrays([[Clipboard]],[[ActiveCell]])
' Generated:              09/15/2024 06:06 PM
'----------------------------------------------------------------------------------------------------
Sub PasteCombineArrays(SourceRange As Range, cellDestination As Range)
    cellDestination.Cells(1, 1).Formula2 = ReplaceInvalidCharFromFormulaWithValid("=" & SplitAreaByAddress(SourceRange, cellDestination))
    AutofitFormulaBar cellDestination.Cells(1, 1)
End Sub

Private Sub TestSplitAreaByAddress()
    Debug.Print SplitAreaByAddress(Sheet1.Range("A1:C7"), ActiveCell)
End Sub

Public Function SplitCombinedCellsOfFormulaDep(ByVal FormulaCell As Range) As String
    
    If Not FormulaCell.HasFormula Then Exit Function
    
    '    Debug.Assert FormulaCell.Address <> "$B$2"
    'Start the timer for this operation
    Logger.Log DEBUG_LOG, "Formula To Split Dependency: " & FormulaCell.Cells(1).Formula2
    
    Dim FinalFormula As String
    FinalFormula = FormulaCell.Cells(1).Formula2
    
    Dim Dependencies As Variant
    Dependencies = GetDirectPrecedents(FinalFormula, FormulaCell.Worksheet)
    
    ' Ensure the Dependency is an array
    If Not IsArray(Dependencies) Then Dependencies = Array(Dependencies)
    
    ' For each dependency, convert to a structured reference and add to the map
    Dim CurrentRange As Range
    Dim CurrentDependency As Variant
    For Each CurrentDependency In Dependencies
        If CurrentDependency <> vbNullString Then
            
            Set CurrentRange = RangeResolver.GetRangeForDependency(CStr(CurrentDependency), FormulaCell)
            
            Dim SplitFormulaRef As String
            If CurrentRange Is Nothing Then
                SplitFormulaRef = CurrentDependency
            ElseIf CurrentRange.Address(False, False) = RemoveDollarSign(RemoveSheetQualifierIfPresent(CurrentDependency)) Then
                ' Use SplitAreaByAddress only if range ref like A1:D5, Not for named range or table or spill range with #
                SplitFormulaRef = SplitAreaByAddress(CurrentRange, FormulaCell)
            Else
                SplitFormulaRef = CurrentDependency
            End If
        
            If RemoveDollarSign(SplitFormulaRef) <> RemoveDollarSign(CStr(CurrentDependency)) Then
                FinalFormula = modDependencyLambdaResult.ReplaceTokenWithNewToken(FinalFormula, CStr(CurrentDependency), SplitFormulaRef)
            End If
            
        End If
    Next CurrentDependency
    
    SplitCombinedCellsOfFormulaDep = FinalFormula
    
End Function

Private Function SplitAreaByAddress(ByVal AreaRange As Range, ByVal Destination As Range) As String
    
    If Not IsSplitNeeded(AreaRange) Then
        ' Base case scenario
        SplitAreaByAddress = GetRangeRefWithSheetNameIfContextIsDiff(AreaRange, Destination)
        Exit Function
    End If
    
    Dim FormulaExpression As String
    
    Dim RowSplits As Long
    RowSplits = CountRowSplits(AreaRange)
    
    Dim ColumnSplits As Long
    ColumnSplits = CountColumnSplits(AreaRange)
    
    If RowSplits = 1 And ColumnSplits = 1 Then
        ' Base case scenario
        FormulaExpression = GetFormulaIfOneRowAndColSplit(AreaRange, Destination)
    Else
        
        Dim StackFXName As String
        Dim AddressArray() As String
        If RowSplits >= ColumnSplits Then
            AddressArray() = Split(SplitByRows(AreaRange), ",")
            StackFXName = VSTACK_FX_NAME
        Else
            AddressArray() = Split(SplitByColumns(AreaRange), ",")
            StackFXName = HSTACK_FX_NAME
        End If
        
        Dim Index As Long
        Dim SplittedAddress As String
        
        Dim SplitAreaRange As Range
        SplittedAddress = ""
        For Index = LBound(AddressArray) To UBound(AddressArray)
            Set SplitAreaRange = Intersect(AreaRange, AreaRange.Worksheet.Range(AddressArray(Index)))
            ' Here is the recursive call.
            SplittedAddress = SplittedAddress & IIf(SplittedAddress = "", "", ",") _
                              & SplitAreaByAddress(SplitAreaRange, Destination)
        Next
        
        FormulaExpression = StackFXName & "(" & SplittedAddress & ")"
        
    End If

    SplitAreaByAddress = FormulaExpression
    
End Function

Public Function IsSplitNeeded(ByVal AreaRange As Range) As Boolean
    
    Dim Result As Boolean
    
    With AreaRange.Worksheet
        If AreaRange.Rows.CountLarge = .Rows.CountLarge Then
            Result = False
        ElseIf AreaRange.Columns.CountLarge = .Columns.CountLarge Then
            Result = False
        Else
            ' If any formula cell present then
            If IsAnyFormulaCellPresent(AreaRange) Then
                ' This number comes from VSTACK and HSTACK maximum number of parameters count.
                Const MAX_NUMBER_OF_FORMULA_CELLS_ALLOWED = 254
                If IsNull(AreaRange.HasSpill) Then
                    Result = (AreaRange.SpecialCells(xlCellTypeFormulas).Cells.Count <= MAX_NUMBER_OF_FORMULA_CELLS_ALLOWED)
                ElseIf AreaRange.HasSpill Then
                    Result = True
                Else
                    Result = (AreaRange.SpecialCells(xlCellTypeFormulas).Cells.Count <= MAX_NUMBER_OF_FORMULA_CELLS_ALLOWED)
                End If
            Else
                Result = False
            End If
        End If
    End With
    
    IsSplitNeeded = Result
    
End Function

Public Function IsAnyFormulaCellPresent(ByVal CheckOnCell As Range) As Boolean
    
    Dim Result As Boolean
    If CheckOnCell.Cells.CountLarge = 1 Then
        Result = (CheckOnCell.HasFormula Or CheckOnCell.HasSpill)
    ElseIf IsNull(CheckOnCell.HasSpill) Then
        ' If the range contains a spill along with other cells then it returns null.
        ' Ref: https://learn.microsoft.com/en-us/office/vba/api/excel.range.hasspill
        Result = True
    ElseIf CheckOnCell.HasSpill Then
        Result = True
    Else
        Dim FormulaCells As Range
        On Error Resume Next
        Set FormulaCells = CheckOnCell.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        Result = (Not FormulaCells Is Nothing)
    End If
    
    IsAnyFormulaCellPresent = Result
    
End Function

Private Function GetFormulaIfOneRowAndColSplit(ByVal AreaRange As Range _
                                               , ByVal Destination As Range) As String
    
    Dim FormulaExpression As String
    If AreaRange.Cells(1).HasSpill Then
        FormulaExpression = SpillRangePartFormulaCreator.GetFormula(AreaRange, Destination)
    Else
        FormulaExpression = GetRangeRefWithSheetNameIfContextIsDiff(AreaRange, Destination)
    End If

    GetFormulaIfOneRowAndColSplit = FormulaExpression
    
End Function

Private Function GetRangeRefWithSheetNameIfContextIsDiff(ByVal AreaRange As Range, ByVal ContextCell As Range) As String
    
    Dim Ref As String
    If AreaRange.Worksheet.name <> ContextCell.Worksheet.name Then
        Ref = GetRangeRefWithSheetName(AreaRange, , False)
    Else
        Ref = AreaRange.Address(False, False)
    End If
    
    GetRangeRefWithSheetNameIfContextIsDiff = Ref
    
End Function

Private Function SplitByRows(AreaRange As Range) As String
    
    Dim RowAreas As String
    Dim NextRowIndex As Long
    NextRowIndex = 1
    Dim RowIndex As Long
    
    For RowIndex = 1 To AreaRange.Rows.Count
        If RowIndex = NextRowIndex Then
            
            Dim RowHeight As Long
            RowHeight = MaxRowHeight(AreaRange.Rows(RowIndex))
            NextRowIndex = RowIndex + RowHeight
            
            Dim RowAreaRange As Range
            Set RowAreaRange = AreaRange.Rows(RowIndex).Resize(RowHeight)
            
            Dim AreaAddress As String
            AreaAddress = RowAreaRange.Address(False, False)
            
            If RowAreaRange.Cells(1, 1).HasSpill Then
                If RowAreaRange.Cells(1, 1).SpillParent.SpillingToRange.Address = RowAreaRange.Address Then
                    AreaAddress = RowAreaRange.Cells(1, 1).SpillParent.Address(False, False) & "#"
                End If
            End If
            
            RowAreas = RowAreas & IIf(RowAreas = "", "", ",") & AreaAddress
            
        End If
    Next
    
    SplitByRows = RowAreas
    
End Function

Private Function SplitByColumns(AreaRange As Range) As String
    
    Dim ColAreas As String
    
    Dim NextColIndex As Long
    NextColIndex = 1
    Dim ColIndex As Long
    For ColIndex = 1 To AreaRange.Columns.Count
        If ColIndex = NextColIndex Then
            
            Dim ColWidth As Long
            ColWidth = MaxColumnWidth(AreaRange.Columns(ColIndex))
            NextColIndex = ColIndex + ColWidth
            
            Dim ColAreaRange As Range
            Set ColAreaRange = AreaRange.Columns(ColIndex).Resize(, ColWidth)
            
            Dim AreaAddress As String
            AreaAddress = ColAreaRange.Address(False, False)
            
            If ColAreaRange.Cells(1, 1).HasSpill Then
                If ColAreaRange.Cells(1, 1).SpillParent.SpillingToRange.Address = ColAreaRange.Address Then
                    AreaAddress = ColAreaRange.Cells(1, 1).SpillParent.Address(False, False) & "#"
                End If
            End If
            
            ColAreas = ColAreas & IIf(ColAreas = "", "", ",") & AreaAddress
            
        End If
    Next ColIndex
    
    SplitByColumns = ColAreas
    
End Function

Private Function CountRowSplits(ByVal AreaRange As Range) As Long
    
    Dim RowCount As Long
    Dim NextRowIndex As Long
    
    NextRowIndex = 1
    Dim RowIndex As Long
    For RowIndex = 1 To AreaRange.Rows.Count
        If RowIndex = NextRowIndex Then
            NextRowIndex = RowIndex + MaxRowHeight(AreaRange.Rows(RowIndex).Cells)
            RowCount = RowCount + 1
        End If
    Next RowIndex
    
    CountRowSplits = RowCount
    
End Function

Private Function CountColumnSplits(ByVal AreaRange As Range) As Long
    
    Dim ColCount As Long
    Dim NextColIndex As Long
    
    NextColIndex = 1
    Dim ColIndex As Long
    For ColIndex = 1 To AreaRange.Columns.Count
        If ColIndex = NextColIndex Then
            NextColIndex = ColIndex + MaxColumnWidth(AreaRange.Columns(ColIndex))
            ColCount = ColCount + 1
        End If
    Next ColIndex
    
    CountColumnSplits = ColCount
    
End Function

Private Function MaxRowHeight(ByVal RowRange As Range) As Long
    
    Dim CurrentCell As Range
    Dim MaxHeight As Long
    
    MaxHeight = 1
    For Each CurrentCell In RowRange.Cells
        
        If CurrentCell.HasSpill Then
            Dim SpillRowCount As Long
            SpillRowCount = CurrentCell.SpillParent.SpillingToRange.Rows.Count
                
            If CurrentCell.SpillParent.Address = CurrentCell.Address Then
                MaxHeight = modUtility.MaxValue(MaxHeight, SpillRowCount)
            Else
                MaxHeight = modUtility.MaxValue(MaxHeight, SpillRowCount - CurrentCell.Row + CurrentCell.SpillParent.Row)
            End If
        End If
        
    Next CurrentCell
    
    MaxRowHeight = MaxHeight
    
End Function

Private Function MaxColumnWidth(ByVal ColRange As Range) As Long
    
    Dim CurrentCell As Range
    Dim MaxWidth As Long
    
    MaxWidth = 1
    For Each CurrentCell In ColRange.Cells
        
        If CurrentCell.HasSpill Then
            Dim SpillColCount As Long
            SpillColCount = CurrentCell.SpillParent.SpillingToRange.Columns.Count
            
            If CurrentCell.SpillParent.Address = CurrentCell.Address Then
                MaxWidth = modUtility.MaxValue(MaxWidth, SpillColCount)
            Else
                MaxWidth = modUtility.MaxValue(MaxWidth, SpillColCount - CurrentCell.Column + CurrentCell.SpillParent.Column)
            End If
        End If
        
    Next CurrentCell
    
    MaxColumnWidth = MaxWidth
    
End Function


