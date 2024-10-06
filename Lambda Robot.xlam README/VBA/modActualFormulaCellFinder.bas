Attribute VB_Name = "modActualFormulaCellFinder"
Option Explicit

Public Function LoopBackToActualCell(ByVal FormulaStartCell As Range) As Range
    
    Logger.Log TRACE_LOG, "Enter modActualFormulaCellFinder.LoopBackToActualCell"
    ' Loops back to the actual cell in case of a spill range.
    ' FormulaStartCell: The cell containing the formula.

    Dim ActualFormulaStartCell As Range
    Set ActualFormulaStartCell = FormulaStartCell
    Logger.Log DEBUG_LOG, "Formula Start cell before loop back : " & ActualFormulaStartCell.Address

    Dim CurrentRange As Range
    Do While True
        Set CurrentRange = LoopBackToCell(ActualFormulaStartCell)
        If IsNothing(CurrentRange) Then
            Exit Do
        Else
            Set ActualFormulaStartCell = CurrentRange
        End If
    Loop

    Logger.Log DEBUG_LOG, "Formula Start cell after loop back : " & ActualFormulaStartCell.Address
    Set LoopBackToActualCell = ActualFormulaStartCell
    Logger.Log TRACE_LOG, "Exit modActualFormulaCellFinder.LoopBackToActualCell"
    
End Function

Private Function LoopBackToCell(ByVal FromCell As Range) As Range
    
    Logger.Log TRACE_LOG, "Enter modActualFormulaCellFinder.LoopBackToCell"
    ' Loops back to the cell directly referenced by the FromCell.
    ' FromCell: The cell to start the loop back.

    On Error GoTo ExitFunction
    
    Dim Result As Range
    Dim DirectPrecedents As Variant
    DirectPrecedents = GetDirectPrecedents(FromCell.Cells(1).Formula2, FromCell.Worksheet)
    
    Dim IsValidToLoopBack As Boolean
    If IsArrayAllocated(DirectPrecedents) Then
        IsValidToLoopBack = (UBound(DirectPrecedents, 1) = LBound(DirectPrecedents, 1) _
                             And FromCell.Cells(1).Formula2 = EQUAL_SIGN & DirectPrecedents(1, 1))
    End If
    
    If IsValidToLoopBack Then
        
        Set Result = RangeResolver.GetRangeForDependency(DirectPrecedents(1, 1), FromCell)
        
        ' Don't loop back to a cell where we don't have a formula.
        Dim TopCell As Range
        Set TopCell = Result.Cells(1)
        If Not Result.HasFormula Then
            Set Result = Nothing
        ElseIf TopCell.HasFormula And TopCell.HasSpill Then
            ' If not entire spill range included or more than spill range is included then ignore it.
            If TopCell.SpillParent.SpillingToRange.Address <> Result.Address Then
                Set Result = Nothing
            End If
        ElseIf TopCell.HasFormula And Result.Cells.CountLarge > 1 Then
            ' If no spill and top cell has a formula but more than one cell then ignore it.
            Set Result = Nothing
        End If
        
    End If
    
    Set LoopBackToCell = Result
    
    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modActualFormulaCellFinder.LoopBackToCell"
    Exit Function

ExitFunction:
    ' Log the error and clear it.
    Logger.Log ERROR_LOG, Err.Number & "-" & Err.Description
    Err.Clear
    Logger.Log TRACE_LOG, "Exit modActualFormulaCellFinder.LoopBackToCell"
    
End Function


