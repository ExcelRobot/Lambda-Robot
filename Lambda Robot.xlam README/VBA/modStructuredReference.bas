Attribute VB_Name = "modStructuredReference"
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed
'@Folder "StructuredRef"
' @Folder "Lambda.Editor.Utility"
' @IgnoreModule IndexedDefaultMemberAccess
Option Explicit
Option Private Module

Private Const TABLE_COL_SEPARATOR As String = ":"

Public Sub PasteAsReferences(ByVal CopyFrom As Range, ByVal PasteTo As Range)
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.PasteAsReferences"
    ' Converting the range from which data is being copied to a structured reference
    Dim StructuredRef As String
    StructuredRef = ConvertToStructuredReference(CopyFrom, PasteTo)
    
    ' If the structured reference is not empty, assign formula
    ' Print error in debug window if any occurs during assignment
    If StructuredRef <> vbNullString Then
        AssignFormulaIfErrorPrintIntoDebugWindow False, PasteTo, EQUAL_SIGN & StructuredRef _
                                                         , "Converted Structured Reference : "
    End If
    Logger.Log TRACE_LOG, "Exit modStructuredReference.PasteAsReferences"
    
End Sub

Public Sub ConvertFormulaToStructuredRef(ByVal FromCell As Range)
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertFormulaToStructuredRef"
    If Not FromCell.HasFormula Then Exit Sub
    Dim FinalFormula As String
    FinalFormula = GetConvertedStructuredFormula(FromCell, True)
    'Only assign formula if something is changed.
    If FinalFormula <> GetCellFormula(FromCell.Cells(1)) Then
        AssignFormulaIfErrorPrintIntoDebugWindow False, FromCell.Cells(1), FinalFormula _
                                                                   , "Converted Structured Reference : "
    End If
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertFormulaToStructuredRef"
    
End Sub

Public Function GetConvertedStructuredFormula(ByVal FromCell As Range _
                                              , Optional ByVal IgnoreIfInputCell As Boolean = False) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.GetConvertedStructuredFormula"
    If Not FromCell.HasFormula Then Exit Function
    '    Debug.Assert FromCell.Address <> "$B$2"
    'Start the timer for this operation
    Logger.Log DEBUG_LOG, "Formula To Convert To Structured Ref: " & GetCellFormula(FromCell.Cells(1))
    Dim StartTime As Double
    StartTime = Timer()
    Dim FinalFormula As String
    FinalFormula = SplitCombinedCellsOfFormulaDep(FromCell, IgnoreIfInputCell)
    ' Get necessary information for structured reference conversion
    Dim Dependencies As Variant
    Dependencies = GetDirectPrecedents(FinalFormula, FromCell.Worksheet)
    
    ' Ensure the Dependency is an array
    If Not IsArray(Dependencies) Then Dependencies = Array(Dependencies)
    
    ' Initialize a collection to map dependencies to structured references
    Dim DependencyToStructuredRefMap As Collection
    Set DependencyToStructuredRefMap = New Collection
    
    ' For each dependency, convert to a structured reference and add to the map
    Dim CurrentRange As Range
    Dim CurrentDependency As Variant
    For Each CurrentDependency In Dependencies
        If CurrentDependency <> vbNullString And Not Is3DReference(CurrentDependency) Then
            Set CurrentRange = RangeResolver.GetRangeForDependency(CStr(CurrentDependency), FromCell)
            Dim StructuredRef As String
            Dim AreaStartTime As Double
            AreaStartTime = Timer()
            If CurrentRange Is Nothing Then
                StructuredRef = CurrentDependency
            ElseIf IsInputCell(CurrentRange, Nothing) And IgnoreIfInputCell Then
                ' Ref Issue: Unexpected behavior when marked inputs contain a mix of values and formulas #125
                StructuredRef = CurrentDependency
            Else
                StructuredRef = ConvertToStructuredReference(CurrentRange, FromCell)
            End If
            Logger.Log DEBUG_LOG, "Run Time for Current Area : " & Timer() - AreaStartTime
            If RemoveDollarSign(StructuredRef) <> RemoveDollarSign(CStr(CurrentDependency)) Then
                DependencyToStructuredRefMap.Add CurrentDependency, CStr(CurrentDependency)
                FinalFormula = modDependencyLambdaResult.ReplaceTokenWithNewToken(FinalFormula, CStr(CurrentDependency), StructuredRef)
            End If
        End If
    Next CurrentDependency
    
    GetConvertedStructuredFormula = FinalFormula
    Logger.Log DEBUG_LOG, "Run Time for Convert Formula To Structured Ref: " & Timer() - StartTime
    Logger.Log TRACE_LOG, "Exit modStructuredReference.GetConvertedStructuredFormula"
    
End Function

Public Function ConvertToStructuredReference(ByVal CopyFrom As Range, ByVal PasteTo As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertToStructuredReference"
    ' Checking whether the input ranges are valid
    If IsNothing(CopyFrom) Or IsNothing(PasteTo) Then
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.ConvertToStructuredReference"
        Exit Function
    End If
    ' Preparing to convert the ranges to structured references
    Dim AllAreas As Collection
    Set AllAreas = New Collection
    ' Iterating over each area in the CopyFrom range
    Dim CurrentRange As Range
    For Each CurrentRange In CopyFrom.Areas
        Dim StructuredRefForCurrentArea As String
        StructuredRefForCurrentArea = ConvertToStructuredReferenceForCurrentArea(CurrentRange, PasteTo)
        AllAreas.Add StructuredRefForCurrentArea
    Next CurrentRange
    ' Determining whether we need to use a stacking formula based on the number of areas
    Dim StackingFormulaName As String
    If AllAreas.Count > 1 Then StackingFormulaName = GetStackingFormulaName(CopyFrom)
    ' Joining all structured references and returning the result
    Dim Result As String
    Result = JoinCollection(AllAreas)
    ConvertToStructuredReference = StackingFormulaName & Result _
                                   & IIf(StackingFormulaName <> vbNullString _
                                         , FIRST_PARENTHESIS_CLOSE, vbNullString)
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertToStructuredReference"
End Function

Private Function ConvertToStructuredReferenceForCurrentArea(ByVal CopyFrom As Range _
                                                            , ByVal PasteTo As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertToStructuredReferenceForCurrentArea"
    ' Logging the entrance and exit into the module, along with CopyFrom and PasteTo addresses
    Logger.Log DEBUG_LOG, "Copy From : " & CopyFrom.Address
    '    Debug.Assert CopyFrom.Address <> "$N$8:$N$61"
    Logger.Log DEBUG_LOG, "Paste To : " & PasteTo.Address
    
    Dim StartTime As Double
    Dim StructuredRef As String
    StartTime = Timer()

    ' Try to convert to a structured reference for a named range
    StructuredRef = ConvertToStructuredReferenceForNamedRange(CopyFrom, PasteTo)
    Logger.Log DEBUG_LOG, "Run Time For Convert To Structured Reference For Named Range: " & Timer() - StartTime
    
    If StructuredRef = vbNullString Then
        StartTime = Timer()

        ' If not a named range, try to convert to a structured reference for a spill range
        StructuredRef = ConvertToStructuredReferenceForSpillRange(CopyFrom, PasteTo)
        Logger.Log DEBUG_LOG, "Run Time For Convert To Structured Reference For Spill Range : " _
                             & Timer() - StartTime
    End If
    
    If StructuredRef = vbNullString Then
        StartTime = Timer()

        ' If not a spill range, try to convert to a structured reference for a table
        StructuredRef = ConvertToStructuredReferenceForTable(CopyFrom, PasteTo)
        Logger.Log DEBUG_LOG, "Run Time For Convert To Structured Reference For Table : " & Timer() - StartTime
    End If
    
    StartTime = Timer()
    Logger.Log DEBUG_LOG, "Run Time for : Get Prefix For Normal Ref " & Timer() - StartTime
    
    If StructuredRef = vbNullString Then
        StartTime = Timer()

        ' If not a table, convert normal range
        StructuredRef = ConvertNormalRange(CopyFrom, PasteTo)
        Logger.Log DEBUG_LOG, "Run Time For Convert Normal Range : " & Timer() - StartTime
    End If
    
    ConvertToStructuredReferenceForCurrentArea = StructuredRef
    
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertToStructuredReferenceForCurrentArea"
    
End Function

Private Function ConvertNormalRange(ByVal CopyFrom As Range, ByVal PasteTo As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertNormalRange"
    Dim StartTime As Double
    StartTime = Timer()

    ' Get prefix for normal ref
    Dim Prefix As String
    Prefix = GetPrefixForNormalRef(CopyFrom, PasteTo)
    
    Dim IsFormulaPresent As Boolean
    IsFormulaPresent = HasFormulaInAnyCell(CopyFrom)
    
    Dim FormulaName As String
    If CopyFrom.Cells.CountLarge = 1 Then
        ConvertNormalRange = Prefix & CopyFrom.Address(False, False)
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.ConvertNormalRange"
        Exit Function
        
    ElseIf CopyFrom.Rows.CountLarge = 1 And IsFormulaPresent Then
        FormulaName = HSTACK_FN_NAME & FIRST_PARENTHESIS_OPEN
    ElseIf CopyFrom.Columns.CountLarge = 1 And IsFormulaPresent Then
        FormulaName = VSTACK_FN_NAME & FIRST_PARENTHESIS_OPEN
    Else
        ConvertNormalRange = Prefix & CopyFrom.Address(False, False)
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.ConvertNormalRange"
        Exit Function
        
    End If
    
    Const MAX_ARGUMENT_IN_STACK_FORMULA As Long = 254
    If CopyFrom.Cells.CountLarge > MAX_ARGUMENT_IN_STACK_FORMULA Then
        FormulaName = CopyFrom.Address(False, False)
    Else
        ' Iterate over each cell and add it to formula
        Dim CurrentCell As Range
        For Each CurrentCell In CopyFrom.Cells
            FormulaName = FormulaName & Prefix & CurrentCell.Address(False, False) & LIST_SEPARATOR
        Next CurrentCell
    End If
    
    ' Close the formula and finalize it
    FormulaName = Text.RemoveFromEndIfPresent(FormulaName, LIST_SEPARATOR) & FIRST_PARENTHESIS_CLOSE
    ConvertNormalRange = FormulaName
    Logger.Log DEBUG_LOG, "Time to Convert Normal Range : " & Timer() - StartTime
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertNormalRange"
    
End Function

Private Function HasFormulaInAnyCell(ByVal GivenRange As Range) As Boolean

    Logger.Log TRACE_LOG, "Enter modStructuredReference.HasFormulaInAnyCell"
    ' Check if any cell in the given range contains a formula
    Dim RangeHavingFormula As Range
    Dim SheetRef As Worksheet
    Set SheetRef = GivenRange.Worksheet
    Dim UsedRangeIntersectingGivenRange As Range
    ' Find the intersection between the given range and the used range of the sheet
    Set UsedRangeIntersectingGivenRange = FindIntersection(GivenRange, SheetRef.UsedRange)
    On Error Resume Next

    'SpecialCells(xlCellTypeFormulas) will set RangeHavingFormula to the range of cells that have a formula.
    ' If no cells have a formula, this will cause an error, so we use On Error Resume Next to ignore it
    Set RangeHavingFormula = FilterUsingSpecialCells(UsedRangeIntersectingGivenRange, xlCellTypeFormulas)
    On Error GoTo 0

    ' Return True if there is any cell in the range that has a formula
    HasFormulaInAnyCell = IsNotNothing(RangeHavingFormula)
    Logger.Log TRACE_LOG, "Exit modStructuredReference.HasFormulaInAnyCell"
    
End Function

Private Function GetPrefixForNormalRef(ByVal CopyFrom As Range, ByVal PasteTo As Range) As String

    Logger.Log TRACE_LOG, "Enter modStructuredReference.GetPrefixForNormalRef"
    ' Check if the CopyFrom and PasteTo ranges are in the same workbook and sheet
    Dim IsSameWorkbook As Boolean
    IsSameWorkbook = (CopyFrom.Worksheet.Parent.Name = PasteTo.Worksheet.Parent.Name)
    Dim IsSameSheet As Boolean
    IsSameSheet = (IsSameWorkbook And (CopyFrom.Worksheet.Name = PasteTo.Worksheet.Name))

    ' If they're not in the same workbook, include the workbook and sheet name in the prefix
    If Not IsSameWorkbook Then
        GetPrefixForNormalRef = SINGLE_QUOTE & LEFT_BRACKET _
                                & EscapeSingeQuote(CopyFrom.Worksheet.Parent.Name) & RIGHT_BRACKET _
                                & EscapeSingeQuote(CopyFrom.Worksheet.Name) & SINGLE_QUOTE & EXCLAMATION_SIGN
                                
        ' If they're in the same workbook but different sheets, include only the sheet name in the prefix
    ElseIf IsSameWorkbook And Not IsSameSheet Then
        GetPrefixForNormalRef = GetSheetRefForRangeReference(CopyFrom.Worksheet.Name, False)
    End If
    Logger.Log TRACE_LOG, "Exit modStructuredReference.GetPrefixForNormalRef"
    
End Function

Public Function ConvertToStructuredReferenceForTable(ByVal CopyFrom As Range, ByVal PasteTo As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertToStructuredReferenceForTable"
    ' Initialize the logging process
    
    ' Try to get the table that encompasses the CopyFrom range
    Dim Table As ListObject
    Set Table = modUtility.GetTableFromRange(CopyFrom)
    ' If the CopyFrom range isn't within a table, exit the function
    If IsNothing(Table) Then Exit Function
    
    ' Check various conditions to understand how to reference the table correctly
    ' These conditions evaluate the specific part of the table that CopyFrom range is referring to
    ' and return the appropriate structured reference
    If Table.Range.Address = CopyFrom.Address Then
        ConvertToStructuredReferenceForTable = Table.Name & TABLE_ALL_MARKER
    ElseIf Table.DataBodyRange.Address = CopyFrom.Address Then
        ConvertToStructuredReferenceForTable = Table.Name
    ElseIf IsTwoRangeEqual(Table.HeaderRowRange, CopyFrom) Then
        ConvertToStructuredReferenceForTable = Table.Name & TABLE_HEADERS_MARKER
    ElseIf IsTotalRange(Table, CopyFrom) Then
        ConvertToStructuredReferenceForTable = Table.Name & TABLE_TOTALS_MARKER
    ElseIf IsOnlyInsideHeader(Table, CopyFrom) Then
        ConvertToStructuredReferenceForTable = Table.Name & ConvertHeaderReference(Table, CopyFrom)
    ElseIf IsOnlyInsideTotalRow(Table, CopyFrom) Then
        ConvertToStructuredReferenceForTable = Table.Name & ConvertTotalRowReference(Table, CopyFrom)
    ElseIf IsOnlyInsideDatabody(Table, CopyFrom) Then
        If FindIntersection(Table.DataBodyRange, CopyFrom).Rows.CountLarge = Table.DataBodyRange.Rows.CountLarge Then
            ConvertToStructuredReferenceForTable = Table.Name & ConvertDataBodyReference(Table, CopyFrom)
        Else
            ConvertToStructuredReferenceForTable = ConvertPartOfDatabodyWithOrWithoutHeaderAndTotal(CopyFrom _
                                                                                                    , PasteTo)
            Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.ConvertToStructuredReferenceForTable"
            Exit Function
        End If
    Else
        Dim Prefix As String
        Prefix = FindColumnNamePrefixes(Table, CopyFrom)
        ' Use the result of the prefix finding function to conditionally construct the table reference
        If Prefix <> "Not TableRef" Then
            Dim ColumnsRef As String
            ColumnsRef = ConvertTableReference(Table, CopyFrom)
            If ColumnsRef = vbNullString Then Prefix = Text.RemoveFromEndIfPresent(Prefix, LIST_SEPARATOR)
            If Prefix <> vbNullString Then
                ConvertToStructuredReferenceForTable = Table.Name & LEFT_BRACKET & Prefix & ColumnsRef & RIGHT_BRACKET
            Else
                ConvertToStructuredReferenceForTable = Table.Name & ColumnsRef
            End If
        Else
            ConvertToStructuredReferenceForTable = ConvertPartOfDatabodyWithOrWithoutHeaderAndTotal(CopyFrom _
                                                                                                    , PasteTo)
            Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.ConvertToStructuredReferenceForTable"
            Exit Function
        End If
    End If
    
    ' If the CopyFrom and PasteTo ranges are in different workbooks, add the workbook name to the reference
    If CopyFrom.Worksheet.Parent.Name <> PasteTo.Worksheet.Parent.Name Then
        ConvertToStructuredReferenceForTable = SINGLE_QUOTE & EscapeSingeQuote(CopyFrom.Worksheet.Parent.Name) _
                                               & SINGLE_QUOTE & EXCLAMATION_SIGN _
                                               & ConvertToStructuredReferenceForTable
    End If
    
    ' End the logging process
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertToStructuredReferenceForTable"
    
End Function

Private Function IsTotalRange(ByVal Table As ListObject, ByVal CopyFrom As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.IsTotalRange"
    
    ' Check if the CopyFrom range is equal to the totals row of the table
    IsTotalRange = IsTwoRangeEqual(Table.TotalsRowRange, CopyFrom)
    
    Logger.Log TRACE_LOG, "Exit modStructuredReference.IsTotalRange"
    
End Function

Private Function ConvertHeaderReference(ByVal Table As ListObject, ByVal CopyFrom As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertHeaderReference"

    ' We check the number of cells in CopyFrom to determine how to convert it to structured reference
    Dim Result As String
    If CopyFrom.Cells.CountLarge = 1 Then
        Result = LEFT_BRACKET & TABLE_HEADERS_MARKER & LIST_SEPARATOR _
                 & ConvertToProperColumnName(CopyFrom.Cells(1).Value) & RIGHT_BRACKET
    Else
        ' We use the first and last column names to create a reference
        Dim FirstColumnName As String
        FirstColumnName = CopyFrom.Cells(1, 1).Value
        Dim LastColumnName As String
        LastColumnName = CopyFrom.Cells(1, CopyFrom.Cells.CountLarge).Value
        
        ' We check if we're referring to the entire header
        If FirstColumnName = Table.ListColumns(1).Name And LastColumnName = Table.ListColumns(Table.ListColumns.Count).Name Then
            Result = LEFT_BRACKET & TABLE_ALL_MARKER & RIGHT_BRACKET
        Else
            Result = LEFT_BRACKET & TABLE_HEADERS_MARKER & LIST_SEPARATOR _
                     & ConvertToProperColumnName(FirstColumnName) & TABLE_COL_SEPARATOR _
                     & ConvertToProperColumnName(LastColumnName) & RIGHT_BRACKET
        End If
    End If
    
    ConvertHeaderReference = Result
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertHeaderReference"
    
End Function

Private Function IsOnlyInsideHeader(ByVal Table As ListObject, ByVal CopyFrom As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.IsOnlyInsideHeader"
    
    ' We check if CopyFrom is only in the table header
    If IsNotNothing(Table.HeaderRowRange) Then
        IsOnlyInsideHeader = IsTwoRangeEqual(FindIntersection(Table.HeaderRowRange, CopyFrom), CopyFrom)
    End If
    
    Logger.Log TRACE_LOG, "Exit modStructuredReference.IsOnlyInsideHeader"
    
End Function

Private Function ConvertToProperColumnName(ByVal GivenColumnName As String) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertToProperColumnName"
    
    ' We convert the GivenColumnName to a correct structured reference format
    Dim SpecialCharsToPutEscapeChar As Variant
    ' Sequence is important here as escape character is single quote
    
    SpecialCharsToPutEscapeChar = Array(SINGLE_QUOTE, HASH_SIGN, LEFT_BRACKET, RIGHT_BRACKET)
    ' Ref : https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
    Dim CurrentChar As Variant
    For Each CurrentChar In SpecialCharsToPutEscapeChar
        GivenColumnName = VBA.Replace(GivenColumnName, CurrentChar, SINGLE_QUOTE & CurrentChar)
    Next CurrentChar
    
    ConvertToProperColumnName = LEFT_BRACKET & GivenColumnName & RIGHT_BRACKET
    
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertToProperColumnName"
    
End Function

Private Function IsOnlyInsideTotalRow(ByVal Table As ListObject, ByVal CopyFrom As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.IsOnlyInsideTotalRow"
    
    ' Check if CopyFrom is entirely within the totals row of the table
    If IsNothing(Table.TotalsRowRange) Then
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.IsOnlyInsideTotalRow"
        Exit Function
    End If
    
    Dim Temp As Range
    Set Temp = FindIntersection(Table.TotalsRowRange, CopyFrom)
    IsOnlyInsideTotalRow = IsTwoRangeEqual(Temp, CopyFrom)

    Logger.Log TRACE_LOG, "Exit modStructuredReference.IsOnlyInsideTotalRow"
    
End Function

Private Function ConvertTotalRowReference(ByVal Table As ListObject, ByVal CopyFrom As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertTotalRowReference"
    
    ' Get the index of the column in the table's totals row
    Dim ColIndex As Long
    ColIndex = CopyFrom.Cells(1).Column - Table.TotalsRowRange.Cells(1).Column + 1
    
    Dim Result As String
    If CopyFrom.Cells.CountLarge = 1 Then
        ' If CopyFrom is a single cell, create a structured reference to that cell in the table's totals row
        Result = LEFT_BRACKET & TABLE_TOTALS_MARKER & LIST_SEPARATOR _
                 & ConvertToProperColumnName(Table.ListColumns.Item(ColIndex).Name) & RIGHT_BRACKET
    Else
        ' If CopyFrom is a range of cells, determine the first and last column names in the table's totals row
        Dim FirstColumnName As String
        FirstColumnName = Table.ListColumns.Item(ColIndex).Name
        
        Dim LastColumnName As String
        ColIndex = CopyFrom.Cells(1, CopyFrom.Cells.CountLarge).Column - Table.TotalsRowRange.Cells(1).Column + 1
        LastColumnName = Table.ListColumns.Item(ColIndex).Name
        
        ' Check if we're referring to the entire totals row
        If FirstColumnName = Table.ListColumns(1).Name And LastColumnName = Table.ListColumns(Table.ListColumns.Count).Name Then
            Result = LEFT_BRACKET & TABLE_TOTALS_MARKER & RIGHT_BRACKET
        Else
            Result = LEFT_BRACKET & TABLE_TOTALS_MARKER & LIST_SEPARATOR _
                     & ConvertToProperColumnName(FirstColumnName) & TABLE_COL_SEPARATOR _
                     & ConvertToProperColumnName(LastColumnName) & RIGHT_BRACKET
        End If
    End If
    
    ConvertTotalRowReference = Result
    
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertTotalRowReference"
    
End Function

Private Function IsOnlyInsideDatabody(ByVal Table As ListObject, ByVal CopyFrom As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.IsOnlyInsideDatabody"
    
    ' Check if CopyFrom is entirely within the table's data body
    Dim Temp As Range
    Set Temp = FindIntersection(Table.DataBodyRange, CopyFrom)
    IsOnlyInsideDatabody = IsTwoRangeEqual(Temp, CopyFrom)
    
    Logger.Log TRACE_LOG, "Exit modStructuredReference.IsOnlyInsideDatabody"
    
End Function

Public Function ConvertDataBodyReference(ByVal Table As ListObject, ByVal CopyFrom As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertDataBodyReference"
    
    ' Get the index of the column in the table's data body
    Dim ColIndex As Long
    ColIndex = CopyFrom.Cells(1).Column - Table.DataBodyRange.Cells(1).Column + 1
    
    Dim Result As String
    If CopyFrom.Columns.CountLarge = 1 Then
        ' If CopyFrom is a single column, create a structured reference to that column in the table's data body
        Result = ConvertToProperColumnName(Table.ListColumns.Item(ColIndex).Name)
    Else
        ' If CopyFrom is a range of columns, determine the first and last column names in the table's data body
        Dim FirstColumnName As String
        FirstColumnName = Table.ListColumns.Item(ColIndex).Name
        
        Dim LastColumnName As String
        ColIndex = CopyFrom.Cells(1, CopyFrom.Columns.Count).Column - Table.DataBodyRange.Cells(1).Column + 1
        LastColumnName = Table.ListColumns.Item(ColIndex).Name
        
        ' Check if we're referring to the entire data body
        If FirstColumnName = Table.ListColumns(1).Name _
           And LastColumnName = Table.ListColumns(Table.ListColumns.Count).Name Then
           
            Result = LEFT_BRACKET & TABLE_DATA_MARKER & RIGHT_BRACKET
        Else
            Result = LEFT_BRACKET & ConvertToProperColumnName(FirstColumnName) & TABLE_COL_SEPARATOR _
                     & ConvertToProperColumnName(LastColumnName) & RIGHT_BRACKET
        End If
    End If
    
    ConvertDataBodyReference = Result
    
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertDataBodyReference"
    
End Function

Private Function FindColumnNamePrefixes(ByVal Table As ListObject, ByVal CopyFrom As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.FindColumnNamePrefixes"
    Dim Prefix As String
    
    ' Get the first column of the CopyFrom range
    Dim GivenRangeFirstCol As Range
    Set GivenRangeFirstCol = CopyFrom.Columns(1)
    
    ' Check if the number of rows in GivenRangeFirstCol is greater than the number of rows in the table's range
    If GivenRangeFirstCol.Rows.CountLarge > Table.Range.Rows.CountLarge Then
        FindColumnNamePrefixes = "Not TableRef"
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.FindColumnNamePrefixes"
        Exit Function
    End If
    
    ' Check if CopyFrom range includes the entire table
    If IsAllIncluded(Table, CopyFrom) Then
        Prefix = TABLE_ALL_MARKER & LIST_SEPARATOR
    Else
        ' Check if CopyFrom range is inside the header row of the table
        If IsNotNothing(Table.HeaderRowRange) Then
            If HasIntersection(Table.HeaderRowRange, GivenRangeFirstCol) Then Prefix = TABLE_HEADERS_MARKER
        End If
        
        ' Check if CopyFrom range is inside the data body of the table
        If HasIntersection(Table.DataBodyRange, GivenRangeFirstCol) Then
            If FindIntersection(Table.DataBodyRange _
                                , GivenRangeFirstCol).Rows.CountLarge = Table.DataBodyRange.Rows.CountLarge Then
                Prefix = IIf(Prefix <> vbNullString, Prefix & LIST_SEPARATOR, vbNullString) & TABLE_DATA_MARKER
            Else
                FindColumnNamePrefixes = "Not TableRef"
                Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.FindColumnNamePrefixes"
                Exit Function
            End If
        End If
        
        ' Check if CopyFrom range is inside the totals row of the table
        If IsNotNothing(Table.TotalsRowRange) Then
            If HasIntersection(Table.TotalsRowRange, GivenRangeFirstCol) Then
                Prefix = IIf(Prefix <> vbNullString, Prefix & LIST_SEPARATOR, vbNullString) & TABLE_TOTALS_MARKER
            End If
        End If
        
        Prefix = Prefix & LIST_SEPARATOR
    End If
    
    FindColumnNamePrefixes = Prefix
    Logger.Log TRACE_LOG, "Exit modStructuredReference.FindColumnNamePrefixes"
    
End Function

Private Function IsAllIncluded(ByVal Table As ListObject, ByVal FirstColumnRange As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.IsAllIncluded"
    
    ' Check if Table.HeaderRowRange is nothing
    If IsNotNothing(Table.HeaderRowRange) Then
        IsAllIncluded = False
        ' Check if FirstColumnRange has no intersection with the table's header row
    ElseIf IsNoIntersection(Table.HeaderRowRange, FirstColumnRange) Then
        IsAllIncluded = False
        ' Check if FirstColumnRange has no intersection with the table's data body
    ElseIf IsNoIntersection(Table.DataBodyRange, FirstColumnRange) Then
        IsAllIncluded = False
        ' Check if Table.TotalsRowRange is nothing
    ElseIf IsNothing(Table.TotalsRowRange) Then
        IsAllIncluded = False
        ' Check if FirstColumnRange has no intersection with the table's totals row
    ElseIf IsNoIntersection(Table.TotalsRowRange, FirstColumnRange) Then
        IsAllIncluded = False
    Else
        IsAllIncluded = True
    End If
    
    ' If all conditions are met, check if FirstColumnRange has the same number of rows as the table's range
    If IsAllIncluded Then
        IsAllIncluded = (FirstColumnRange.Rows.CountLarge = Table.Range.Rows.CountLarge)
    End If
    
    Logger.Log TRACE_LOG, "Exit modStructuredReference.IsAllIncluded"
    
End Function

Private Function ConvertTableReference(ByVal Table As ListObject, ByVal CopyFrom As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertTableReference"
    
    ' Get the index of the column in the table's header row
    Dim ColIndex As Long
    ColIndex = CopyFrom.Cells(1).Column - Table.HeaderRowRange.Cells(1).Column + 1
    
    Dim Result As String
    If CopyFrom.Columns.CountLarge = 1 Then
        ' If CopyFrom is a single column, create a structured reference to that column in the table
        Result = ConvertToProperColumnName(Table.ListColumns.Item(ColIndex).Name)
    Else
        ' If CopyFrom is a range of columns, determine the first and last column names in the table's header row
        Dim FirstColumnName As String
        FirstColumnName = Table.ListColumns.Item(ColIndex).Name
        
        Dim LastColumnName As String
        ColIndex = CopyFrom.Cells(1, CopyFrom.Columns.CountLarge).Column - Table.HeaderRowRange.Cells(1).Column + 1
        LastColumnName = Table.ListColumns.Item(ColIndex).Name
        
        ' Check if we're referring to the entire table, if so, return an empty string
        If FirstColumnName = Table.ListColumns(1).Name _
           And LastColumnName = Table.ListColumns(Table.ListColumns.Count).Name Then
            Result = vbNullString
        Else
            ' Otherwise, create a structured reference for the range of columns
            Result = ConvertToProperColumnName(FirstColumnName) & TABLE_COL_SEPARATOR & ConvertToProperColumnName(LastColumnName)
        End If
    End If
    
    ConvertTableReference = Result
    
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertTableReference"
    
End Function

' Below part is for named range.
Private Function ConvertToStructuredReferenceForNamedRange(ByVal CopyFrom As Range _
                                                           , ByVal PasteTo As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertToStructuredReferenceForNamedRange"
    
    ' Get the prefix for the named range reference
    Dim Prefix As String
    Prefix = GetPrefixForNamedRange(CopyFrom, PasteTo)
    
    ' Find the closest named range to CopyFrom
    Dim CurrentName As Name
    Set CurrentName = FindClosestNamedRange(CopyFrom)
    
    ' If no named range is found, combine multiple named ranges if applicable
    If IsNothing(CurrentName) Then
        ConvertToStructuredReferenceForNamedRange = CombineMultipleNamedRange(CopyFrom, PasteTo)
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.ConvertToStructuredReferenceForNamedRange"
        Exit Function
    End If
    
    ' If a named range is found, get the structured reference for the named range
    ConvertToStructuredReferenceForNamedRange = GetStructuredRefForNamedRange(CopyFrom _
                                                                              , PasteTo, CurrentName, Prefix)
    
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertToStructuredReferenceForNamedRange"

End Function

Private Function GetStructuredRefForNamedRange(ByVal CopyFrom As Range, ByVal PasteTo As Range _
                                                                       , ByVal CurrentName As Name _
                                                                        , ByVal Prefix As String) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.GetStructuredRefForNamedRange"
    ' Check if the whole named range is present in the CopyFrom range
    If IsWholeNamedRangePresent(CurrentName, CopyFrom) Then
        ' Return the structured reference for the whole named range
        GetStructuredRefForNamedRange = Prefix & GetProperNameForNamedRange(CurrentName, PasteTo)
    Else
        ' Convert columns to formula text for the named range
        GetStructuredRefForNamedRange = ConvertColumnsToFormulaTextForNamedRange(CurrentName, CopyFrom, PasteTo)
    End If
    Logger.Log TRACE_LOG, "Exit modStructuredReference.GetStructuredRefForNamedRange"
    
End Function

Private Function GetProperNameForNamedRange(ByVal CurrentName As Name, ByVal PasteTo As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.GetProperNameForNamedRange"
    ' Check if the named range is local and if it's in the same sheet as the PasteTo range
    If IsLocalScopeNamedRange(CurrentName.NameLocal) _
       And CurrentName.RefersToRange.Worksheet.Name = PasteTo.Worksheet.Name Then
        ' Extract the name from the local named range
        GetProperNameForNamedRange = ExtractNameFromLocalNameRange(CurrentName.NameLocal)
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.GetProperNameForNamedRange"
        Exit Function
    End If
    ' Return the original name if it's not a local named range or in a different sheet
    GetProperNameForNamedRange = CurrentName.Name
    Logger.Log TRACE_LOG, "Exit modStructuredReference.GetProperNameForNamedRange"
    
End Function

Private Function GetPrefixForNamedRange(ByVal CopyFrom As Range, ByVal PasteTo As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.GetPrefixForNamedRange"
    ' Check if CopyFrom and PasteTo are in different workbooks
    If CopyFrom.Worksheet.Parent.Name <> PasteTo.Worksheet.Parent.Name Then
        ' Return the workbook prefix
        GetPrefixForNamedRange = SINGLE_QUOTE & EscapeSingeQuote(CopyFrom.Worksheet.Parent.Name) _
                                 & SINGLE_QUOTE & EXCLAMATION_SIGN
    End If
    Logger.Log TRACE_LOG, "Exit modStructuredReference.GetPrefixForNamedRange"
    
End Function

Private Function CombineMultipleNamedRange(ByVal CopyFrom As Range, ByVal PasteTo As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.CombineMultipleNamedRange"
    ' Create a collection to store all named ranges intersecting with CopyFrom
    Dim AllNamedRangeAddress As Collection
    Set AllNamedRangeAddress = New Collection
    
    ' Find the closest named range to CopyFrom
    Dim Temp  As Name
    Set Temp = FindClosestNamedRange(CopyFrom.Cells(1))
    Dim Index As Long
    
    ' Get the prefix for the named range reference
    Dim Prefix As String
    Prefix = GetPrefixForNamedRange(CopyFrom, PasteTo)
    
    ' Exit if no named range is found
    If IsNothing(Temp) Then Exit Function
    
    ' If CopyFrom and named range refer to the same range, return the structured reference for the named range
    If Temp.RefersToRange.Address = CopyFrom.Address Then
        CombineMultipleNamedRange = Prefix & Temp.Name
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.CombineMultipleNamedRange"
        Exit Function
    End If
    
    ' Initialize variables for building the formula text for combined named ranges
    Dim FormulaName As String
    Dim NamedRangeRangeInColumns As Boolean
    FormulaName = GetStackFormulaForMultiNamedRange(CopyFrom, Temp.RefersToRange)
    If FormulaName = vbNullString Then Exit Function
    NamedRangeRangeInColumns = (FormulaName = HSTACK_FN_NAME & FIRST_PARENTHESIS_OPEN)
    Dim CombinedRange As Range
    
    ' Loop through the named ranges and combine their addresses
    Do While True
        Dim CellIntersectingNamedRangeAndCopyFrom As Range
        Set CellIntersectingNamedRangeAndCopyFrom = FindIntersection(CopyFrom, Temp.RefersToRange)
        Set CombinedRange = UnionOfNonExistableRange(CombinedRange, CellIntersectingNamedRangeAndCopyFrom)
        Dim StructuredRef As String
        StructuredRef = GetStructuredRefForNamedRange(CellIntersectingNamedRangeAndCopyFrom _
                                                      , PasteTo, Temp, Prefix)
        AllNamedRangeAddress.Add StructuredRef
        
        ' Determine the next named range to process based on whether the current range is in columns or rows
        If NamedRangeRangeInColumns Then
            Index = Temp.RefersToRange.Columns(Temp.RefersToRange.Columns.CountLarge).Cells(1).Column - CopyFrom.Cells(1).Column + 2
            Set Temp = FindClosestNamedRange(CopyFrom.Columns(Index).Cells(1))
            If Index > CopyFrom.Columns.CountLarge Then Exit Do
        Else
            Index = Temp.RefersToRange.Rows(Temp.RefersToRange.Rows.CountLarge).Cells(1).Row - CopyFrom.Cells(1).Row + 2
            Set Temp = FindClosestNamedRange(CopyFrom.Rows(Index).Cells(1))
            If Index > CopyFrom.Rows.CountLarge Then Exit Do
        End If
        
        If IsNothing(Temp) Then Exit Function
        FormulaName = GetStackFormulaForMultiNamedRange(CopyFrom, Temp.RefersToRange)
        If FormulaName = vbNullString Then Exit Function
        If CombinedRange.Address = CopyFrom.Address Then Exit Do
        
    Loop
    
    ' Join the addresses of the named ranges with the formula text for combined named ranges
    Dim Result As String
    Result = JoinCollection(AllNamedRangeAddress)
    CombineMultipleNamedRange = FormulaName & Result & FIRST_PARENTHESIS_CLOSE
    Logger.Log TRACE_LOG, "Exit modStructuredReference.CombineMultipleNamedRange"
    
End Function

Private Function FindClosestNamedRange(ByVal CopyFrom As Range) As Name
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.FindClosestNamedRange"
    ' Create a collection to store all named ranges intersecting with CopyFrom
    Dim AllMatchNamedRanges As Collection
    Set AllMatchNamedRanges = New Collection
    
    Dim Book  As Workbook
    Set Book = CopyFrom.Worksheet.Parent
    Dim CurrentName As Name
    For Each CurrentName In Context.GetAllRangeRefNamedRangeFromBook(Book)
        If IsRangeInsideNamedRange(CurrentName, CopyFrom) And CurrentName.Visible Then
            Logger.Log DEBUG_LOG, "Matched Named Range : " & CurrentName.Name
            AllMatchNamedRanges.Add CurrentName, CurrentName.NameLocal
        End If
    Next CurrentName
    
    Const LONG_MAX As Long = 2147483647
    Dim MinimumCellCount As Long
    MinimumCellCount = LONG_MAX
    Dim Temp  As Long
    Dim FinalName As Name
    
    ' Find the closest named range to CopyFrom based on the number of cells it covers
    For Each CurrentName In AllMatchNamedRanges
        Temp = CurrentName.RefersToRange.Cells.CountLarge - CopyFrom.Cells.CountLarge
        If Temp < MinimumCellCount Then
            MinimumCellCount = Temp
            Set FinalName = CurrentName
        End If
    Next CurrentName
    
    Set FindClosestNamedRange = FinalName
    Logger.Log TRACE_LOG, "Exit modStructuredReference.FindClosestNamedRange"
    
End Function

Public Function IsRangeInsideNamedRange(ByVal CurrentNameRange As Name, ByVal CopyFrom As Range) As Boolean
    
    ' Check if the named range refers to a valid range
    Dim ReferredRange As Range
    On Error Resume Next
    Set ReferredRange = CurrentNameRange.RefersToRange
    On Error GoTo 0
    If IsNothing(ReferredRange) Then
        IsRangeInsideNamedRange = False
    ElseIf CopyFrom.Worksheet.Name = ReferredRange.Worksheet.Name Then
        ' Check if CopyFrom is fully inside the named range
        IsRangeInsideNamedRange = IsTwoRangeEqual(FindIntersection(ReferredRange, CopyFrom), CopyFrom)
        If IsRangeInsideNamedRange Then Exit Function
    End If
    
End Function

Private Function IsWholeNamedRangePresent(ByVal CurrentName As Name, ByVal CopyFrom As Range) As Boolean
    IsWholeNamedRangePresent = (CurrentName.RefersToRange.Address = CopyFrom.Address)
End Function

Private Function ConvertColumnsToFormulaTextForNamedRange(ByVal CurrentName As Name _
                                                          , ByVal CopyFrom As Range _
                                                           , ByVal PasteTo As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertColumnsToFormulaTextForNamedRange"
    ' Check if the named range is in the same workbook as PasteTo
    Dim IsSameWorkbook As Boolean
    IsSameWorkbook = (CopyFrom.Worksheet.Parent.Name = PasteTo.Worksheet.Parent.Name)
    Dim Prefix As String
    If Not IsSameWorkbook Then
        Prefix = SINGLE_QUOTE & EscapeSingeQuote(CopyFrom.Worksheet.Parent.Name) _
                 & SINGLE_QUOTE & EXCLAMATION_SIGN
    End If
    ' Convert columns to formula text for the named range
    ConvertColumnsToFormulaTextForNamedRange = ConvertColumnsToFormulaText(CurrentName.RefersToRange _
                                                                           , Prefix _
                                                                            & GetProperNameForNamedRange(CurrentName _
                                                                                                         , PasteTo), CopyFrom)
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertColumnsToFormulaTextForNamedRange"

End Function

Private Function ConvertColumnsToFormulaText(ByVal RefersToRange As Range _
                                             , ByVal NameToRefer As String _
                                              , ByVal CopyFrom As Range _
                                               , Optional ByVal IsForTable As Boolean = False) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertColumnsToFormulaText"
    ' Get the column and row index for CopyFrom relative to RefersToRange
    Dim ColIndex As Long
    ColIndex = CopyFrom.Cells(1).Column - RefersToRange.Cells(1).Column + 1
    Dim RowIndex As Long
    RowIndex = CopyFrom.Cells(1).Row - RefersToRange.Cells(1).Row + 1
    
    Dim ImplicitPrefix As String
    ImplicitPrefix = IIf(CopyFrom.Cells.CountLarge = 1, IMPLICIT_INTERSECTION_OPERATOR, vbNullString)
    
    Dim Formula As String
    ' Check if CopyFrom covers the whole columns or rows of RefersToRange
    If CopyFrom.Rows.CountLarge = RefersToRange.Rows.CountLarge Then
        If ColIndex = 1 Then
            ' Return structured reference for the whole columns
            Formula = TAKE_FN_NAME & FIRST_PARENTHESIS_OPEN _
                                          & NameToRefer & LIST_SEPARATOR _
                                          & LIST_SEPARATOR & CopyFrom.Columns.CountLarge _
                                          & FIRST_PARENTHESIS_CLOSE
                                          
        ElseIf ColIndex + CopyFrom.Columns.CountLarge - 1 = RefersToRange.Columns.CountLarge Then
            ' Return structured reference for the whole columns
            Formula = DROP_FN_NAME & FIRST_PARENTHESIS_OPEN _
                                          & NameToRefer & LIST_SEPARATOR _
                                          & LIST_SEPARATOR & (ColIndex - 1) _
                                          & FIRST_PARENTHESIS_CLOSE
        Else
            ' Return structured reference for part of the columns
            Formula = TAKE_FN_NAME & FIRST_PARENTHESIS_OPEN _
                                          & DROP_FN_NAME & FIRST_PARENTHESIS_OPEN _
                                          & NameToRefer & LIST_SEPARATOR _
                                          & LIST_SEPARATOR & (ColIndex - 1) _
                                          & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR _
                                          & LIST_SEPARATOR & CopyFrom.Columns.CountLarge _
                                          & FIRST_PARENTHESIS_CLOSE
        End If
    ElseIf CopyFrom.Columns.CountLarge = RefersToRange.Columns.CountLarge Then
        If RowIndex = 1 Then
            ' Return structured reference for the whole rows
            Formula = TAKE_FN_NAME & FIRST_PARENTHESIS_OPEN _
                                          & NameToRefer & LIST_SEPARATOR _
                                          & CopyFrom.Rows.CountLarge & FIRST_PARENTHESIS_CLOSE
                                          
        ElseIf RowIndex + CopyFrom.Rows.CountLarge - 1 = RefersToRange.Rows.CountLarge Then
            ' Return structured reference for the whole rows
            Formula = TAKE_FN_NAME & FIRST_PARENTHESIS_OPEN _
                                          & NameToRefer & LIST_SEPARATOR & "-" _
                                          & CopyFrom.Rows.CountLarge & FIRST_PARENTHESIS_CLOSE
        ElseIf RowIndex = 2 And RefersToRange.Rows.CountLarge - CopyFrom.Rows.CountLarge = 2 _
               And CopyFrom.Rows.CountLarge > 1 Then
            ' Special case for structured reference when CopyFrom is in the middle of the rows
            If IsForTable Then
                Formula = TAKE_FN_NAME & FIRST_PARENTHESIS_OPEN & DROP_FN_NAME _
                                              & FIRST_PARENTHESIS_OPEN & NameToRefer & LIST_SEPARATOR & "1)" _
                                              & LIST_SEPARATOR _
                                              & RefersToRange.Rows.CountLarge - 2 & FIRST_PARENTHESIS_CLOSE
            Else
                Formula = DROP_FN_NAME & FIRST_PARENTHESIS_OPEN & DROP_FN_NAME _
                                              & FIRST_PARENTHESIS_OPEN & NameToRefer & LIST_SEPARATOR & "1)" _
                                              & LIST_SEPARATOR & "-1)"
            End If
        Else
            ' Return structured reference for part of the rows
            Formula = TAKE_FN_NAME & FIRST_PARENTHESIS_OPEN & DROP_FN_NAME _
                                          & FIRST_PARENTHESIS_OPEN & NameToRefer & LIST_SEPARATOR _
                                          & (RowIndex - 1) & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR _
                                          & CopyFrom.Rows.CountLarge & FIRST_PARENTHESIS_CLOSE
                                          
        End If
    End If
    
    If Formula <> vbNullString Then Formula = ImplicitPrefix & Formula
    
    ConvertColumnsToFormulaText = Formula
    
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertColumnsToFormulaText"
    
End Function

Private Function ConvertSequenceToArrayConstant(ByVal StartCol As Long, ByVal NumberOfCol As Long) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertSequenceToArrayConstant"
    ' Convert a sequence of columns to an array constant
    Dim SequenceResult As Variant
    If NumberOfCol = 1 Then
        ConvertSequenceToArrayConstant = StartCol
    Else
        SequenceResult = Application.WorksheetFunction.Sequence(1, NumberOfCol, StartCol)
        ConvertSequenceToArrayConstant = LEFT_BRACE _
                                         & Join(SequenceResult, ARRAY_CONST_COLUMN_SEPARATOR) _
                                         & RIGHT_BRACE
    End If
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertSequenceToArrayConstant"
    
End Function

' Below part is for Spill Range.
Private Function ConvertToStructuredReferenceForSpillRange(ByVal CopyFrom As Range _
                                                           , ByVal PasteTo As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertToStructuredReferenceForSpillRange"
    ' Check if CopyFrom and PasteTo are in the same workbook and the same sheet
    Dim IsSameWorkbook As Boolean
    IsSameWorkbook = (CopyFrom.Worksheet.Parent.Name = PasteTo.Worksheet.Parent.Name)
    Dim IsSameSheet As Boolean
    IsSameSheet = (IsSameWorkbook And (CopyFrom.Worksheet.Name = PasteTo.Worksheet.Name))
    
    ' Try to get the spilling range for CopyFrom
    Dim SpillRange As Range
    On Error Resume Next
    Set SpillRange = CopyFrom.Cells(1).SpillParent.SpillingToRange
    On Error GoTo 0
    
    ' If CopyFrom is part of a spilling range
    If IsNotNothing(SpillRange) Then
        If FindIntersection(SpillRange, CopyFrom).Address = CopyFrom.Address Then
            Dim ReferToName As String
            
            ' Generate the structured reference for the spilling range
            If IsSameWorkbook And IsSameSheet Then
                ReferToName = SpillRange.Cells(1).Address(False, False) & HASH_SIGN
            ElseIf IsSameWorkbook And Not IsSameSheet Then
                ReferToName = GetRangeRefWithSheetName(SpillRange.Cells(1), False) & HASH_SIGN
            ElseIf Not IsSameWorkbook Then
                ReferToName = SINGLE_QUOTE & LEFT_BRACKET _
                              & EscapeSingeQuote(CopyFrom.Worksheet.Parent.Name) & RIGHT_BRACKET _
                              & EscapeSingeQuote(CopyFrom.Worksheet.Name) & SINGLE_QUOTE & EXCLAMATION_SIGN _
                              & SpillRange.Cells(1).Address(False, False) & HASH_SIGN
                              
            End If
            
            If SpillRange.Address = CopyFrom.Address Then
                ConvertToStructuredReferenceForSpillRange = ReferToName
            ElseIf SpillRange.Rows.CountLarge > 1 And SpillRange.Columns.CountLarge > 1 _
                   And SpillRange.Cells(1).Address = CopyFrom.Address Then
                ' If spill parent and Spill Range is a 2D grid then return VbNullString
                ConvertToStructuredReferenceForSpillRange = vbNullString
                Exit Function
            Else
                ' Convert columns to formula text for the spilling range
                ConvertToStructuredReferenceForSpillRange = ConvertColumnsToFormulaText(SpillRange _
                                                                                        , ReferToName, CopyFrom)
            End If
        End If
    End If
    
    ' If the spilling range is not directly available, try to combine spill ranges
    If ConvertToStructuredReferenceForSpillRange = vbNullString Then
        ConvertToStructuredReferenceForSpillRange = CombineSpillRanges(CopyFrom, PasteTo)
    End If
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertToStructuredReferenceForSpillRange"
    
End Function

Private Function CombineSpillRanges(ByVal CopyFrom As Range, ByVal PasteTo As Range) As String
    Logger.Log TRACE_LOG, "Enter modStructuredReference.CombineSpillRanges"
    ' Delegate the task to CombineSpillRangesForContigiousArea
    CombineSpillRanges = CombineSpillRangesForContigiousArea(CopyFrom, PasteTo)
    Logger.Log TRACE_LOG, "Exit modStructuredReference.CombineSpillRanges"
End Function

Private Function CombineSpillRangesForContigiousArea(ByVal CopyFrom As Range, ByVal PasteTo As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.CombineSpillRangesForContigiousArea"
    ' Create a collection to store all spill range addresses
    Dim AllSpillAddress As Collection
    Set AllSpillAddress = New Collection
    
    Dim Temp As Range
    Set Temp = CopyFrom.Cells(1).SpillingToRange
    Dim Index As Long
    
    ' Get the prefix based on whether the ranges are in the same workbook
    Dim Prefix As String
    Prefix = GetPrefixForNormalRef(CopyFrom, PasteTo)
    
    ' If CopyFrom is not part of a spill range, exit
    If IsNothing(Temp) Then Exit Function
    If Temp.Address = CopyFrom.Address Then
        ' If CopyFrom is the entire spilling range, return the structured reference
        CombineSpillRangesForContigiousArea = Prefix & Temp.Cells(1).Address(False, False) & HASH_SIGN
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.CombineSpillRangesForContigiousArea"
        Exit Function
    End If
    
    Dim FormulaName As String
    Dim SpillRangeInColumns As Boolean
    SpillRangeInColumns = IsSpillRangeInColumns(CopyFrom)
    FormulaName = IIf(SpillRangeInColumns, HSTACK_FN_NAME, VSTACK_FN_NAME) & FIRST_PARENTHESIS_OPEN
    
    Do While True
        
        Dim StructuredRef As String
        Set Temp = FindIntersection(Temp, CopyFrom)
        ' Get the structured reference for the current spill range and add it to the collection
        StructuredRef = ConvertToStructuredReferenceForSpillRange(Temp, PasteTo)
        AllSpillAddress.Add StructuredRef
        If SpillRangeInColumns Then
            Index = Temp.Columns(Temp.Columns.Count).Cells(1).Column - CopyFrom.Cells(1).Column + 2
            Set Temp = CopyFrom.Columns(Index).Cells(1).SpillingToRange
            If Index > CopyFrom.Columns.CountLarge Then Exit Do
        Else
            Index = Temp.Rows(Temp.Rows.Count).Cells(1).Row - CopyFrom.Cells(1).Row + 2
            Set Temp = CopyFrom.Rows(Index).Cells(1).SpillingToRange
            If Index > CopyFrom.Rows.CountLarge Then Exit Do
        End If
        
        If IsNothing(Temp) Then Exit Do
        
    Loop
    
    ' Join the addresses of the spill ranges with the formula text for combined spill ranges
    Dim Result As String
    Result = JoinCollection(AllSpillAddress)
    CombineSpillRangesForContigiousArea = FormulaName & Result & FIRST_PARENTHESIS_CLOSE
    Logger.Log TRACE_LOG, "Exit modStructuredReference.CombineSpillRangesForContigiousArea"
    
End Function

Private Function IsSpillRangeInColumns(ByVal CopyFrom As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.IsSpillRangeInColumns"
    Dim Temp As Range
    Set Temp = CopyFrom.Cells(1).SpillingToRange
    IsSpillRangeInColumns = (FindIntersection(Temp, CopyFrom).Rows.CountLarge = CopyFrom.Rows.CountLarge _
                             And CopyFrom.Columns.CountLarge > Temp.Columns.CountLarge)
    Logger.Log TRACE_LOG, "Exit modStructuredReference.IsSpillRangeInColumns"
    
End Function

Private Function IsSpillRangeInRows(ByVal CopyFrom As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.IsSpillRangeInRows"
    Dim Temp As Range
    Set Temp = CopyFrom.Cells(1).SpillingToRange
    IsSpillRangeInRows = (Temp.Columns.CountLarge = CopyFrom.Columns.CountLarge And CopyFrom.Rows.CountLarge > Temp.Rows.CountLarge)
    Logger.Log TRACE_LOG, "Exit modStructuredReference.IsSpillRangeInRows"
    
End Function

Private Function GetStackingFormulaName(ByVal CopyFrom As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.GetStackingFormulaName"
    ' Determine the formula name for the stacking function (HSTACK or VSTACK)
    If IsInColumns(CopyFrom) Then
        GetStackingFormulaName = HSTACK_FN_NAME & FIRST_PARENTHESIS_OPEN
    ElseIf IsInRows(CopyFrom) Then
        GetStackingFormulaName = VSTACK_FN_NAME & FIRST_PARENTHESIS_OPEN
    Else
        GetStackingFormulaName = vbNullString
    End If
    Logger.Log TRACE_LOG, "Exit modStructuredReference.GetStackingFormulaName"
    
End Function

Private Function IsInColumns(ByVal Source As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.IsInColumns"
    ' Check if the Source range is stacked horizontally (in columns)
    If Source.Areas.Count = 1 Then
        IsInColumns = False
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.IsInColumns"
        Exit Function
    End If
    
    Dim Area As Range
    Dim RowCount As Long
    Dim FirstRow As Long
    RowCount = Source.Areas(1).Rows.Count
    FirstRow = Source.Areas(1).Row
    IsInColumns = True
    For Each Area In Source.Areas
        If Area.Rows.CountLarge <> RowCount Or Area.Row <> FirstRow Then
            IsInColumns = False
            Exit For
        End If
    Next
    Logger.Log TRACE_LOG, "Exit modStructuredReference.IsInColumns"
    
End Function

Private Function IsInRows(ByVal Source As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.IsInRows"
    ' Check if the Source range is stacked vertically (in rows)
    If Source.Areas.Count = 1 Then
        IsInRows = False
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modStructuredReference.IsInRows"
        Exit Function
    End If
    
    Dim Area As Range
    Dim ColumnCount As Long
    Dim FirstColumn As Long
    ColumnCount = Source.Areas(1).Columns.Count
    FirstColumn = Source.Areas(1).Column
    IsInRows = True
    For Each Area In Source.Areas
        If Area.Columns.CountLarge <> ColumnCount Or Area.Column <> FirstColumn Then
            IsInRows = False
            Exit For
        End If
    Next
    Logger.Log TRACE_LOG, "Exit modStructuredReference.IsInRows"
    
End Function

Private Function JoinCollection(ByVal AllSpillAddress As Collection) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.JoinCollection"
    
    ' Join the addresses from the collection into a single string
    Dim Result As String
    Dim CurrentItem As Variant
    
    For Each CurrentItem In AllSpillAddress
        Result = Result & CurrentItem & LIST_SEPARATOR
    Next CurrentItem
    
    Result = Text.RemoveFromEndIfPresent(Result, LIST_SEPARATOR)
    JoinCollection = Result
    
    Logger.Log TRACE_LOG, "Exit modStructuredReference.JoinCollection"
    
End Function

Private Function IsTwoRangeEqual(ByVal FirstRange As Range, ByVal SecondRange As Range) As Boolean
    
    ' Check if two ranges are equal
    If IsNothing(FirstRange) And IsNothing(SecondRange) Then
        IsTwoRangeEqual = False
    ElseIf IsNothing(FirstRange) And IsNotNothing(SecondRange) Then
        IsTwoRangeEqual = False
    ElseIf IsNotNothing(FirstRange) And IsNothing(SecondRange) Then
        IsTwoRangeEqual = False
    Else
        IsTwoRangeEqual = (FirstRange.Address = SecondRange.Address)
    End If
    
End Function

Private Function ConvertPartOfDatabodyWithOrWithoutHeaderAndTotal(ByVal CopyFrom As Range, ByVal PasteTo As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.ConvertPartOfDatabodyWithOrWithoutHeaderAndTotal"
    ' Convert a part of the data body with or without header and total for a structured reference
    
    ' Get the table from the CopyFrom range
    Dim Table As ListObject
    Set Table = modUtility.GetTableFromRange(CopyFrom)
    If IsNothing(Table) Then Exit Function
    
    ' Find the intersection of the data body range with CopyFrom
    Dim ReferredTo As Range
    Set ReferredTo = FindIntersection(Table.DataBodyRange, CopyFrom)
    
    ' If the number of columns is not equal to the number of table columns, exit
    If ReferredTo.Columns.CountLarge <> Table.ListColumns.Count Then Exit Function
    
    ' Get the name to refer in the structured reference
    Dim NameToRefer As String
    NameToRefer = GetDatabodyRefName(CopyFrom, PasteTo, Table)
    
    ' Convert the data body part to formula text
    Dim DataBodyFormulaPart As String
    DataBodyFormulaPart = ConvertColumnsToFormulaText(Table.DataBodyRange, NameToRefer, ReferredTo, True)
    
    Dim FxName As String
    Dim FormulaText As String
    FormulaText = DataBodyFormulaPart & LIST_SEPARATOR
    Dim Temp As Range
    
    ' Check if CopyFrom intersects with the table header row
    If IsNotNothing(Table.HeaderRowRange) Then Set Temp = FindIntersection(CopyFrom, Table.HeaderRowRange)
    
    If IsTwoRangeEqual(Temp, Table.HeaderRowRange) Then
        FormulaText = Replace(NameToRefer, TABLE_DATA_MARKER, TABLE_HEADERS_MARKER) _
                      & LIST_SEPARATOR & FormulaText
                      
        FxName = VSTACK_FN_NAME & FIRST_PARENTHESIS_OPEN
    End If
    
    ' Check if CopyFrom intersects with the table totals row
    If IsNotNothing(Table.TotalsRowRange) Then Set Temp = FindIntersection(CopyFrom, Table.TotalsRowRange)
    
    If IsTwoRangeEqual(Temp, Table.TotalsRowRange) Then
        FormulaText = FormulaText & Replace(NameToRefer, TABLE_DATA_MARKER, TABLE_TOTALS_MARKER) & LIST_SEPARATOR
        FxName = VSTACK_FN_NAME & FIRST_PARENTHESIS_OPEN
    End If
    
    ConvertPartOfDatabodyWithOrWithoutHeaderAndTotal = FxName _
                                                       & Text.RemoveFromEndIfPresent(FormulaText, LIST_SEPARATOR) _
                                                       & IIf(FxName = vbNullString, vbNullString, FIRST_PARENTHESIS_CLOSE)
                                                       
    Logger.Log TRACE_LOG, "Exit modStructuredReference.ConvertPartOfDatabodyWithOrWithoutHeaderAndTotal"
    
End Function

Private Function GetDatabodyRefName(ByVal CopyFrom As Range _
                                    , ByVal PasteTo As Range, ByVal Table As ListObject) As String
    Logger.Log TRACE_LOG, "Enter modStructuredReference.GetDatabodyRefName"
    ' Get the reference name for the data body
    
    ' Check if CopyFrom and PasteTo are in the same workbook
    Dim IsSameWorkbook As Boolean
    IsSameWorkbook = (CopyFrom.Worksheet.Parent.Name = PasteTo.Worksheet.Parent.Name)

    ' Get the prefix for the reference name based on the workbook
    If Not IsSameWorkbook Then
        GetDatabodyRefName = SINGLE_QUOTE & EscapeSingeQuote(CopyFrom.Worksheet.Parent.Name) _
                             & SINGLE_QUOTE & EXCLAMATION_SIGN
    End If
    
    ' Return the reference name for the data body
    GetDatabodyRefName = GetDatabodyRefName & Table.Name & TABLE_DATA_MARKER
    Logger.Log TRACE_LOG, "Exit modStructuredReference.GetDatabodyRefName"
End Function

Private Function GetStackFormulaForMultiNamedRange(ByVal CopyFrom As Range _
                                                   , ByVal NamedRangeRefersToRange As Range) As String
    Logger.Log TRACE_LOG, "Enter modStructuredReference.GetStackFormulaForMultiNamedRange"
    ' Get the stacking formula name for a multi-named range
    
    Dim CellIntersectingNamedRangeAndCopyFrom As Range
    Set CellIntersectingNamedRangeAndCopyFrom = FindIntersection(CopyFrom, NamedRangeRefersToRange)
    
    ' Determine if the CopyFrom range is stacked horizontally or vertically
    If CopyFrom.Rows.CountLarge = CellIntersectingNamedRangeAndCopyFrom.Rows.CountLarge _
       And CopyFrom.Columns.CountLarge > CellIntersectingNamedRangeAndCopyFrom.Columns.CountLarge Then
        GetStackFormulaForMultiNamedRange = HSTACK_FN_NAME & FIRST_PARENTHESIS_OPEN
    ElseIf CopyFrom.Columns.CountLarge = CellIntersectingNamedRangeAndCopyFrom.Columns.CountLarge _
           And CopyFrom.Rows.CountLarge > CellIntersectingNamedRangeAndCopyFrom.Rows.CountLarge Then
        GetStackFormulaForMultiNamedRange = VSTACK_FN_NAME & FIRST_PARENTHESIS_OPEN
    Else
        GetStackFormulaForMultiNamedRange = vbNullString
    End If
    Logger.Log TRACE_LOG, "Exit modStructuredReference.GetStackFormulaForMultiNamedRange"
    
End Function

' @Description("Normal Union throw error if any of those range is nothing. This will not. That is the use of this function.")
' @Dependency("No Dependency")
' @ExampleCall :UnionOfNonExistableRange(FirstRange,Nothing)
Public Function UnionOfNonExistableRange(ByVal FirstRange As Range, ByVal SecondRange As Range) As Range
    
    Logger.Log TRACE_LOG, "Enter modStructuredReference.UnionOfNonExistableRange"
    If IsNothing(FirstRange) And IsNothing(SecondRange) Then
        Set UnionOfNonExistableRange = Nothing
    ElseIf IsNothing(FirstRange) Then
        Set UnionOfNonExistableRange = SecondRange
    ElseIf IsNothing(SecondRange) Then
        Set UnionOfNonExistableRange = FirstRange
    Else
        Set UnionOfNonExistableRange = Union(FirstRange, SecondRange)
    End If
    Logger.Log TRACE_LOG, "Exit modStructuredReference.UnionOfNonExistableRange"
    
End Function


