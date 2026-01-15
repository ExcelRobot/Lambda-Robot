VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpillRangePartFormulaCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredID
Option Explicit

Private Type TSpillRangePartFormulaCreator
    SpillParentCell As Range
    SpillParentSheetName As String
    ParentCellFormulaRef As String
    SpillingRange As Range
    TopLeftCell As Range
    TopRightCell As Range
    BottomLeftCell As Range
    BottomRightCell As Range
    IsSpillingRangeIncluded As Boolean
    IsTopLeftCellIncluded As Boolean
    IsTopRightCellIncluded As Boolean
    IsBottomLeftCellIncluded As Boolean
    IsBottomRightCellIncluded As Boolean
    IsAnyCellOfFirstCol As Boolean
    IsAnyCellOfLastCol As Boolean
    IsAnyCellOfFirstRow As Boolean
    IsAnyCellOfLastRow As Boolean
    TopCellRowDiff As Long
    TopCellColDiff As Long
    CellsCount As Long
    RowCount As Long
    ColCount As Long
    SpillRowCount As Long
    SpillColCount As Long
End Type

Private this As TSpillRangePartFormulaCreator

Public Function GetFormula(ByVal SelectedCells As Range, ByVal DestinationCell As Range) As String
    
    If Not SelectedCells.Cells(1).HasSpill Then Exit Function
    
    ExtractInfoFromSelectedCells SelectedCells, DestinationCell
    
    Dim Formula As String
    Const MIN_ROW_OR_COL_COUNT As Long = 3
    
    Dim ImplicitPrefix As String
    ImplicitPrefix = IIf(this.CellsCount = 1, IMPLICIT_INTERSECTION_OPERATOR, vbNullString)
    
    Dim TakeWithParentCellRef As String
    TakeWithParentCellRef = ImplicitPrefix _
                            & TAKE_FN_NAME & FIRST_PARENTHESIS_OPEN & this.ParentCellFormulaRef
    
    Dim TakeAndDropWithParentCellRef As String
    TakeAndDropWithParentCellRef = ImplicitPrefix _
                                   & TAKE_FN_NAME & FIRST_PARENTHESIS_OPEN _
                                   & DROP_FN_NAME & FIRST_PARENTHESIS_OPEN & this.ParentCellFormulaRef
    
    With this
        ' Here order of these cases are important. If you want to make it independent then you have to change code
        ' and add validation. But as select case statement will go one by one, we can keep in this way.
        Select Case True
            Case .IsSpillingRangeIncluded
                ' Whole spill range is included
                Formula = this.ParentCellFormulaRef
            
            Case .IsTopLeftCellIncluded And .IsTopRightCellIncluded
                
                ' First x row is included
                Formula = TakeWithParentCellRef & "," & .RowCount & FIRST_PARENTHESIS_CLOSE
        
            Case .IsBottomLeftCellIncluded And .IsBottomRightCellIncluded
                
                ' Last x row is included.
                Formula = TakeWithParentCellRef & "," & "-" & .RowCount & FIRST_PARENTHESIS_CLOSE
        
            Case .IsTopLeftCellIncluded And .IsBottomLeftCellIncluded
                
                ' first x col included.
                Formula = TakeWithParentCellRef & ",," & .ColCount & FIRST_PARENTHESIS_CLOSE
                
            Case .IsTopRightCellIncluded And .IsBottomRightCellIncluded
                
                ' last x col is included.
                Formula = TakeWithParentCellRef & ",," & "-" & .ColCount & FIRST_PARENTHESIS_CLOSE
                
            Case .IsTopLeftCellIncluded
                
                ' Spill Parent with x row and y col included.
                Formula = TakeWithParentCellRef & "," & .RowCount & "," & .ColCount & FIRST_PARENTHESIS_CLOSE
                
            Case .IsTopRightCellIncluded
                
                ' first row last cell with x row and y col included.
                Formula = TakeWithParentCellRef & "," & .RowCount _
                          & "," & "-" & .ColCount & FIRST_PARENTHESIS_CLOSE
                
            Case .IsBottomLeftCellIncluded
                
                ' first col last cell with x row and y col included.
                Formula = TakeWithParentCellRef & "," & "-" & .RowCount _
                          & "," & .ColCount & FIRST_PARENTHESIS_CLOSE
                
            Case .IsBottomRightCellIncluded
                
                ' last row last cell with x row and y col included.
                Formula = TakeWithParentCellRef & "," & "-" & .RowCount _
                          & "," & "-" & .ColCount & FIRST_PARENTHESIS_CLOSE
            
            Case .IsAnyCellOfFirstCol And .IsAnyCellOfLastCol
                
                ' middle x rows included.
                If .RowCount = 1 Then
                    Formula = ImplicitPrefix & INDEX_FN_NAME & FIRST_PARENTHESIS_OPEN & .ParentCellFormulaRef & "," & .TopCellRowDiff + 1 & ",)"
                Else
                    Formula = TakeAndDropWithParentCellRef & "," & .TopCellRowDiff & FIRST_PARENTHESIS_CLOSE _
                              & "," & .RowCount & FIRST_PARENTHESIS_CLOSE
                End If
                
            Case .IsAnyCellOfFirstRow And .IsAnyCellOfLastRow
                
                'middle x column included.
                If .ColCount = 1 Then
                    Formula = ImplicitPrefix & INDEX_FN_NAME & FIRST_PARENTHESIS_OPEN & .ParentCellFormulaRef & ",," & .TopCellColDiff + 1 & FIRST_PARENTHESIS_CLOSE
                Else
                    Formula = TakeAndDropWithParentCellRef & ",," & .TopCellColDiff & FIRST_PARENTHESIS_CLOSE _
                              & ",," & .ColCount & FIRST_PARENTHESIS_CLOSE
                End If
                
            Case .IsAnyCellOfFirstCol
                
                ' first col any cell with x rows and y col
                Formula = TakeAndDropWithParentCellRef & "," & .TopCellRowDiff & FIRST_PARENTHESIS_CLOSE _
                          & "," & .RowCount & "," & .ColCount & FIRST_PARENTHESIS_CLOSE
            
            Case .IsAnyCellOfFirstRow
                
                ' first row any cell with x rows and y col
                Formula = TakeAndDropWithParentCellRef & ",," & .TopCellColDiff & FIRST_PARENTHESIS_CLOSE _
                          & "," & .RowCount & "," & .ColCount & FIRST_PARENTHESIS_CLOSE
                         
            Case .IsAnyCellOfLastCol
                
                ' last col any cell with x rows and y col
                Formula = TakeAndDropWithParentCellRef & "," & .TopCellRowDiff & FIRST_PARENTHESIS_CLOSE _
                          & "," & .RowCount & "," & "-" & .ColCount & FIRST_PARENTHESIS_CLOSE
            
            Case .IsAnyCellOfLastRow
                ' last row any cell with x rows and y col
                Formula = TakeAndDropWithParentCellRef & ",," & .TopCellColDiff & FIRST_PARENTHESIS_CLOSE _
                          & "," & "-" & .RowCount & "," & .ColCount & FIRST_PARENTHESIS_CLOSE
            
                        
            Case .RowCount = 1 And .ColCount = 1
            
                ' Any middle single cell selection.
                Formula = INDEX_FN_NAME & FIRST_PARENTHESIS_OPEN & .ParentCellFormulaRef & "," & .TopCellRowDiff + 1 & "," & .TopCellColDiff + 1 & FIRST_PARENTHESIS_CLOSE
            
            Case Else
                
                ' Any middle multiple cells selection.
                Formula = TakeAndDropWithParentCellRef & "," & .TopCellRowDiff & "," _
                          & .TopCellColDiff & FIRST_PARENTHESIS_CLOSE & "," & .RowCount & "," & .ColCount & FIRST_PARENTHESIS_CLOSE
                
        End Select
    End With
    
    GetFormula = Formula
    
End Function

Private Sub ExtractInfoFromSelectedCells(ByVal SelectedCells As Range, ByVal DestinationCell As Range)
    
    With this
        
        Set .SpillParentCell = SelectedCells.Cells(1).SpillParent
        .SpillParentSheetName = .SpillParentCell.Worksheet.Name
        
        ' We are considering that workbook is same. If need to support for different workbook as well then need
        ' to change code accordingly.
        .ParentCellFormulaRef = vbNullString
        If DestinationCell.Worksheet.Name <> .SpillParentSheetName Then
            .ParentCellFormulaRef = GetSheetRefForRangeReference(.SpillParentSheetName)
        End If
        
        .ParentCellFormulaRef = .ParentCellFormulaRef _
                                & .SpillParentCell.Address(False, False) & "#"
        
        Set .TopLeftCell = .SpillParentCell
        .IsTopLeftCellIncluded = IsSubRange(SelectedCells, .TopLeftCell)
        
        Set .SpillingRange = .SpillParentCell.SpillingToRange
        this.SpillRowCount = .SpillingRange.Rows.CountLarge
        this.SpillColCount = .SpillingRange.Columns.CountLarge
        
        .IsSpillingRangeIncluded = (.SpillingRange.Address = SelectedCells.Address)
        Set .TopRightCell = .SpillingRange.Rows(1).Cells(.SpillingRange.Columns.CountLarge)
        .IsTopRightCellIncluded = IsSubRange(SelectedCells, .TopRightCell)
        
        Set .BottomLeftCell = .SpillingRange.Rows(.SpillingRange.Rows.CountLarge).Cells(1)
        .IsBottomLeftCellIncluded = IsSubRange(SelectedCells, .BottomLeftCell)
        
        Set .BottomRightCell = .BottomLeftCell.Offset(, .SpillingRange.Columns.CountLarge - 1)
        .IsBottomRightCellIncluded = IsSubRange(SelectedCells, .BottomRightCell)
        
        .IsAnyCellOfFirstCol = IsAnyCommonCell(SelectedCells, .SpillingRange.Columns(1))
        .IsAnyCellOfFirstRow = IsAnyCommonCell(SelectedCells, .SpillingRange.Rows(1))
        .IsAnyCellOfLastCol = IsAnyCommonCell(SelectedCells, .SpillingRange.Columns(.SpillingRange.Columns.CountLarge))
        .IsAnyCellOfLastRow = IsAnyCommonCell(SelectedCells, .SpillingRange.Rows(.SpillingRange.Rows.CountLarge))
        
        this.CellsCount = SelectedCells.Cells.CountLarge
        this.RowCount = SelectedCells.Rows.CountLarge
        this.ColCount = SelectedCells.Columns.CountLarge
        this.TopCellRowDiff = SelectedCells.Row - .SpillParentCell.Row
        this.TopCellColDiff = SelectedCells.Column - .SpillParentCell.Column
        
    End With
    
End Sub

Private Function IsAnyCommonCell(ByVal FirstRange As Range, ByVal SecondRange As Range) As Boolean
    
    Dim IntersectingCells As Range
    Set IntersectingCells = Intersect(FirstRange, SecondRange)
    
    IsAnyCommonCell = (Not IntersectingCells Is Nothing)
    
End Function


