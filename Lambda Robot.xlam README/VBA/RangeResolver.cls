VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Utility.RangeResolver")
'@PredeclaredId
Option Explicit

Public Function GetRangeRefFromNamedRangeOrTable(ByVal FromBook As Workbook _
                                                 , ByVal GivenName As String _
                                                  , ByVal HelperSheet As Worksheet) As String
    
    Logger.Log TRACE_LOG, "Enter RangeResolver.GetRangeRefFromNamedRangeOrTable"
    ' Extracts the range reference from a named range or a table.

    If GivenName = vbNullString Or Is3DReference(GivenName) Then
        GetRangeRefFromNamedRangeOrTable = GivenName
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword RangeResolver.GetRangeRefFromNamedRangeOrTable"
        Exit Function
    End If
    
    Dim CurrentName As Name
    On Error Resume Next
    'Check for local scoped named range first and if not then for the global one.
    Set CurrentName = FromBook.Names(GetSheetRefForRangeReference(HelperSheet.Name, False) & GivenName)
    If IsNothing(CurrentName) Then
        Set CurrentName = FromBook.Names(GivenName)
    End If
    On Error GoTo 0
    
    Dim Result As String
    If IsNotNothing(CurrentName) Then
        ' The given name is a named range.
        Dim ReferredRange As Range
        If IsRefersToRangeIsNothing(CurrentName) Then
            Result = GivenName
        Else
            Set ReferredRange = CurrentName.RefersToRange
            Result = GetRangeRefWithSheetName(ReferredRange, False)
        End If
    Else
        ' The given name is not a named range, check if it is part of a table.
        Dim CurrentTableRange As Range
        Set CurrentTableRange = GetRangeForDependency(GivenName, HelperSheet.Cells(1))

        If IsNothing(CurrentTableRange) Then
            Result = GivenName
        ElseIf IsNothing(CurrentTableRange.ListObject) Then
            ' Return the original name if it is not inside a table.
            Result = GivenName
        Else
            Result = GetRangeRefWithSheetName(CurrentTableRange, False)
        End If
    End If
    
    GetRangeRefFromNamedRangeOrTable = Result
    Logger.Log TRACE_LOG, "Exit RangeResolver.GetRangeRefFromNamedRangeOrTable"
    
End Function

Public Function FindRangeFromText(ByVal FromBook As Workbook _
                                  , ByVal FormulaInSheet As Worksheet _
                                   , ByVal RangeReference As String) As Range
    
    Logger.Log TRACE_LOG, "Enter RangeResolver.FindRangeFromText"
    ' Finds the range based on the given range reference.
    Dim Result As Range
    If IsSpilledRangeRef(RangeReference) Then
        ' The range reference is dynamic, remove the trailing DYNAMIC_CELL_REFERENCE_SIGN sign and get the range.
        Set Result = GetRange( _
                     Text.RemoveFromEndIfPresent(RangeReference, DYNAMIC_CELL_REFERENCE_SIGN) _
                     , FromBook _
                      , FormulaInSheet _
                       )
        Set Result = Result.SpillingToRange
    Else
        ' The range reference is not dynamic, directly get the range.
        Set Result = GetRange(RangeReference, FromBook, FormulaInSheet)
    End If
    
    Set FindRangeFromText = Result
    Logger.Log TRACE_LOG, "Exit RangeResolver.FindRangeFromText"
    
End Function

' Is [ is fixed for Workbook qualified range ref or not.
' for workbook reference from another workbook table we may have range like 'Lambda Robot Extensive Tests.xlsx'!ArrowMapTable[Arrow List]
' and we may have single quote for global scoped named range as well. 'TextOffset lambda testing.xlsm'!TestName

' @Description("This will parse text to range based on different Condition. Currently the last parameter is not being used")
' @Dependency("Text.BetweenDelimiter,Text.AfterDelimiter")
' @ExampleCall : GetRange ("'[ORDER FORM V1.2.xlsm]MAINORDERFORM'!$L$25") >> This will give range from ORDER FORM V1.2.xlsm open workbook and MAINORDERFORM sheet and L25 cell
' @Date : 20 April 2022 05:57:23 PM
' @PossibleError : 1004 (If not a proper cell address)
'                         : 9 (If workbook or worksheet is not valid input)

Public Function GetRange(ByVal RangeReference As String _
                         , Optional ByVal DefaultWorkbook As Workbook _
                          , Optional ByVal DefaultSheet As Worksheet) As Range
    
    Logger.Log TRACE_LOG, "Enter RangeResolver.GetRange"
    ' Get the range object from the given range reference.
    
    If IsClosedWorkbookRef(RangeReference) Or Is3DReference(RangeReference) Then
        Set GetRange = Nothing
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword RangeResolver.GetRange"
        Exit Function
    End If
    
    On Error GoTo NoRangeFound
    
    ' Single quote and [ is needed only if we have space in sheet name or in the workbook name.
    ' Otherwise it can start with just [ name.
    ' now we can have reference with single quote and '! in two different scenario.
    ' Same workbook but different sheet with local scoped named range.
    ' another workbook with global named range or table named range.
    ' And if we have two scenario (Active book sheet name = another workbook name and that sheet local scoped named range = global named range of other book)
    ' then active workbook will get the preference.
    ' We can have simple reference with sheet name and normal cell address as well.
    ' We can have the reference without single quote as well.
    ' So we can have only WorkbookName!NamedRange or WorkbookName!TableName[ColIfNeeded]
    
    Dim Result As Range
    
    Dim CurrentBook As Workbook
    Set CurrentBook = GetWorkbookFromRangeRef(RangeReference, DefaultWorkbook)
    
    If IsAnotherWorkbookTableOrNamedRangeRef(RangeReference, DefaultWorkbook) Then
        Set Result = Range(RangeReference)
    Else
        Dim CurrentSheet As Worksheet
        Set CurrentSheet = GetSheetRefFromRangeRef(RangeReference, CurrentBook, DefaultSheet)
        If Text.Contains(RangeReference, SHEET_NAME_SEPARATOR) Then
            Set Result = CurrentSheet.Range(Text.AfterDelimiter(RangeReference, SHEET_NAME_SEPARATOR, , FROM_END))
        ElseIf Context.IsNamedRangeExists(CurrentBook, RangeReference) Then
            Set Result = CurrentBook.Names(RangeReference).RefersToRange
        ElseIf IsTableExist(CurrentBook, RangeReference) Then
            Set Result = GetTable(CurrentBook, RangeReference).DataBodyRange
        ElseIf IsPossibleTableColRef(RangeReference) Then
            Set Result = Range(RangeReference)
        Else
            Set Result = CurrentSheet.Range(RangeReference)
        End If
    
    End If
    
    Logger.Log DEBUG_LOG, "Book Name: " & Result.Worksheet.Parent.Name _
                         & vbNewLine & "Worksheet Name: " & Result.Worksheet.Name _
                         & vbNewLine & "Address: " & Result.Address & vbNewLine
    
    Set GetRange = Result
    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword RangeResolver.GetRange"
    Exit Function

NoRangeFound:
    Set GetRange = Nothing
    Logger.Log TRACE_LOG, "Exit RangeResolver.GetRange"

End Function

Private Function GetSheetRefFromRangeRef(ByVal RangeRef As String _
                                         , ByVal SearchInBook As Workbook _
                                          , ByVal DefaultSheet As Worksheet) As Worksheet
    
    Logger.Log TRACE_LOG, "Enter RangeResolver.GetSheetRefFromRangeRef"
    Dim Result As Worksheet
    If DefaultSheet Is Nothing Then
        Set Result = ActiveSheet
    Else
        Set Result = DefaultSheet
    End If
    
    Dim SheetName As String
    If Not Text.Contains(RangeRef, SHEET_NAME_SEPARATOR) Then
        SheetName = vbNullString
    ElseIf Text.IsStartsWith(RangeRef, SINGLE_QUOTE & LEFT_BRACKET) _
           Or Text.IsStartsWith(RangeRef, LEFT_BRACKET) Then
        SheetName = Text.BetweenDelimiter(RangeRef, RIGHT_BRACKET, SHEET_NAME_SEPARATOR, , , , FROM_END)
        SheetName = Text.RemoveFromEndIfPresent(SheetName, SINGLE_QUOTE)
    Else
        SheetName = Text.BeforeDelimiter(RangeRef, SHEET_NAME_SEPARATOR, , FROM_END)
    End If
    
    SheetName = Text.RemoveFromBothEndIfPresent(SheetName, SINGLE_QUOTE)
    
    ' If we have single quote in sheet name then we had to use double single quote to escape.
    ' so remove that.
    
    SheetName = UnEscapeSingleQuote(SheetName)
    
    If SheetName <> vbNullString And IsSheetExist(SheetName, SearchInBook) Then
        Set Result = SearchInBook.Worksheets(SheetName)
    End If
    
    Set GetSheetRefFromRangeRef = Result
    Logger.Log TRACE_LOG, "Exit RangeResolver.GetSheetRefFromRangeRef"
    
End Function

Private Function IsAnotherWorkbookTableOrNamedRangeRef(ByVal RangeRef As String _
                                                       , ByVal DefaultBook As Workbook) As Boolean
    
    Logger.Log TRACE_LOG, "Enter RangeResolver.IsAnotherWorkbookTableOrNamedRangeRef"
    ' This Range reference can mean two things: 'ABC D.xlsm'!SomeName
    ' ABC D.xlsm is sheet name and SomeName is a local scoped named range or
    ' ABC D.xlsm is another workbook name and SomeName is a global scoped named range.
    ' 'Email Manager V10.xlsm'!CompetingShipsData[IMO] these could mean only one thing.
    ' Email Manager V10.xlsm is another workbook name and CompetingShipsData is the table name.
    ' This function is handling these two cases.
    
    Dim TempBook As Workbook
    If DefaultBook Is Nothing Then
        Set TempBook = ActiveWorkbook
    Else
        Set TempBook = DefaultBook
    End If
    
    Dim IsValidForOtherBookStructuredRef As Boolean
    If Text.IsStartsWith(RangeRef, SINGLE_QUOTE & LEFT_BRACKET) Then
        IsValidForOtherBookStructuredRef = False
    ElseIf Text.IsStartsWith(RangeRef, LEFT_BRACKET) Then
        IsValidForOtherBookStructuredRef = False
    ElseIf Not Text.Contains(RangeRef, SHEET_NAME_SEPARATOR) Then
        IsValidForOtherBookStructuredRef = False
    Else
        IsValidForOtherBookStructuredRef = True
    End If
    
    Dim BookName As String
    Dim NamedOrTableOrRangePart As String
    
    Dim Result As Boolean
    
    If IsValidForOtherBookStructuredRef Then
        NamedOrTableOrRangePart = Text.AfterDelimiter(RangeRef, SHEET_NAME_SEPARATOR, , FROM_END)
            
        BookName = Text.BeforeDelimiter(RangeRef, SHEET_NAME_SEPARATOR, , FROM_END)
        BookName = Text.RemoveFromBothEndIfPresent(BookName, SINGLE_QUOTE)
        BookName = UnEscapeSingleQuote(BookName)
            
        ' Now this could be sheet name or book name.
        ' It could be only sheet name if we have a sheet by that name and a local scoped named range.
        If IsOpenWorkbookExists(BookName) Then
            If IsSheetExist(BookName, TempBook) Then
                If Context.IsLocalScopedNamedRangeExists(TempBook.Worksheets(BookName), NamedOrTableOrRangePart) Then
                    Result = False
                End If
            Else
                Set TempBook = Application.Workbooks.Item(BookName)
                If Context.IsNamedRangeExists(TempBook, NamedOrTableOrRangePart) Then
                    Result = True
                ElseIf IsTableExist(TempBook, NamedOrTableOrRangePart) Then
                    Result = True
                ElseIf IsTableExist(TempBook, Text.BeforeDelimiter(NamedOrTableOrRangePart, LEFT_BRACKET)) Then
                    Result = True
                Else
                    Result = False
                End If
        
            End If
        Else
            Result = False
        End If
    End If
    
    IsAnotherWorkbookTableOrNamedRangeRef = Result
    Logger.Log TRACE_LOG, "Exit RangeResolver.IsAnotherWorkbookTableOrNamedRangeRef"
    
End Function

Private Function GetWorkbookFromRangeRef(ByVal RangeRef As String _
                                         , ByVal DefaultBook As Workbook) As Workbook
    
    Logger.Log TRACE_LOG, "Enter RangeResolver.GetWorkbookFromRangeRef"
    Dim Result As Workbook
    If DefaultBook Is Nothing Then
        Set Result = ActiveWorkbook
    Else
        Set Result = DefaultBook
    End If
    
    Dim BookName As String
    Dim NamedOrTableOrRangePart As String
    
    If Text.Contains(RangeRef, SHEET_NAME_SEPARATOR) Then
        NamedOrTableOrRangePart = Text.AfterDelimiter(RangeRef, SHEET_NAME_SEPARATOR, , FROM_END)
        
        If Text.IsStartsWith(RangeRef, SINGLE_QUOTE & LEFT_BRACKET) _
           Or Text.IsStartsWith(RangeRef, LEFT_BRACKET) Then
            BookName = Text.BetweenDelimiter(RangeRef, LEFT_BRACKET, RIGHT_BRACKET)
            BookName = UnEscapeSingleQuote(BookName)
        Else
            
            BookName = Text.BeforeDelimiter(RangeRef, SHEET_NAME_SEPARATOR, , FROM_END)
            BookName = Text.RemoveFromBothEndIfPresent(BookName, SINGLE_QUOTE)
            BookName = UnEscapeSingleQuote(BookName)
            
            ' Now this could be sheet name or book name.
            ' It could be only sheet name if we have a sheet by that name and a local scoped named range.
            If IsSheetExist(BookName, Result) And IsOpenWorkbookExists(BookName) Then
                If Context.IsLocalScopedNamedRangeExists(Result.Worksheets(BookName), NamedOrTableOrRangePart) Then
                    BookName = vbNullString
                End If
            End If
            
        End If
    
    End If
    
    If BookName <> vbNullString And IsOpenWorkbookExists(BookName) Then
        Set Result = Application.Workbooks(BookName)
    End If
    
    Set GetWorkbookFromRangeRef = Result
    Logger.Log TRACE_LOG, "Exit RangeResolver.GetWorkbookFromRangeRef"
    
End Function

Public Function GetRangeForDependency(ByVal DependencyText As String _
                                      , ByVal HelperCell As Range) As Range
    
    Logger.Log TRACE_LOG, "Enter RangeResolver.GetRangeForDependency"
    ' Retrieves the range object corresponding to the given dependency text.
    ' The function checks if the dependency text refers to a named range or a cell reference.
    
    If Is3DReference(DependencyText) Then Exit Function
    
    Dim Result As Range
    Dim CurrentName As Name
    
    Dim InBook As Workbook
    Set InBook = HelperCell.Worksheet.Parent
    On Error Resume Next
    ' Attempt to find the named range in the parent workbook.
    If Context.IsLocalScopedNamedRangeExists(HelperCell.Worksheet, DependencyText) Then
        Set CurrentName = HelperCell.Worksheet.Names(DependencyText)
    Else
        Set CurrentName = InBook.Names(DependencyText)
    End If
    On Error GoTo 0
    
    Dim RangeRefSheet As Worksheet
    
    If IsImplicitColumnRef(DependencyText, InBook) Then
        Set Result = Nothing
    ElseIf IsTableExist(InBook:=InBook, TableName:=DependencyText) Then
        Set Result = GetTable(InBook, DependencyText).DataBodyRange
    ElseIf IsPossibleTableColRef(DependencyText) _
           And IsTableExist(InBook, Text.BeforeDelimiter(DependencyText, LEFT_BRACKET)) Then
        
        Dim Table As ListObject
        Set Table = GetTable(InBook, Text.BeforeDelimiter(DependencyText, LEFT_BRACKET))
        Set Result = Table.Range.Worksheet.Range(DependencyText)
        
    ElseIf IsNothing(CurrentName) Then
        
        Dim IsDynamicRef As Boolean
        IsDynamicRef = Text.IsEndsWith(DependencyText, HASH_SIGN)
        ' Sometimes we refer to a cell with "#" although there is no spill.
        ' In that case, Range(DependencyText) throws an error.
        
        DependencyText = Text.RemoveFromEndIfPresent(DependencyText, HASH_SIGN)
        
        Set Result = GetRange(DependencyText, InBook, HelperCell.Worksheet)
    
        ' If it is a dynamic reference and has a spill, use the spilling range.
        If IsDynamicRef And Result.HasSpill Then
            Set Result = Result.SpillingToRange
        End If
        
    Else
        
        ' If it is a named range, return the referred range.
        If IsRefersToRangeIsNothing(CurrentName) Then
            Set Result = Nothing
        Else
            Set Result = CurrentName.RefersToRange
        End If
        
    End If
    
    Set GetRangeForDependency = Result
    Logger.Log TRACE_LOG, "Exit RangeResolver.GetRangeForDependency"
    
End Function

Private Function IsPossibleTableColRef(ByVal DependencyText As String) As Boolean
    
    Logger.Log TRACE_LOG, "Enter RangeResolver.IsPossibleTableColRef"
    IsPossibleTableColRef = ( _
                            Text.Contains(DependencyText, LEFT_BRACKET) _
                            And Text.IsEndsWith(DependencyText, RIGHT_BRACKET) _
                            )
    Logger.Log TRACE_LOG, "Exit RangeResolver.IsPossibleTableColRef"

End Function

Private Function IsTableColumnRef(ByVal InBook As Workbook, ByVal DependencyText As String) As Boolean
    
    Logger.Log TRACE_LOG, "Enter RangeResolver.IsTableColumnRef"
    Dim IsBothTableColIdentifierExists As Boolean
    IsBothTableColIdentifierExists = IsPossibleTableColRef(DependencyText)
    
    Dim IsValidRef As Boolean
    If IsBothTableColIdentifierExists Then
        
        Dim TableName As String
        TableName = Text.BeforeDelimiter(DependencyText, LEFT_BRACKET)
        If IsTableExist(InBook, TableName) Then
            Dim ColName As String
            
            ' Ref: https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
            ' As we have already removed first one when using Text.AfterDelimiter then we need to try only once.
            ColName = Text.AfterDelimiter(DependencyText, LEFT_BRACKET)
            ColName = Text.RemoveFromStartIfPresent(ColName, LEFT_BRACKET)
            
            ' For special character we may need double close bracket.
            ColName = Text.RemoveFromEndIfPresent(ColName, RIGHT_BRACKET)
            ColName = Text.RemoveFromEndIfPresent(ColName, RIGHT_BRACKET)
            
            Const COL_SEPERATOR As String = ":"  ' Like [Col1:Col3]
            Dim Table As ListObject
            
            Set Table = GetTable(InBook, TableName)
            If Text.Contains(ColName, COL_SEPERATOR) Then
                IsValidRef = ( _
                             IsTableColExists(Table, Text.BeforeDelimiter(ColName, COL_SEPERATOR)) _
                             And IsTableColExists(Table, Text.AfterDelimiter(ColName, COL_SEPERATOR)) _
                             )
                
            Else
                IsValidRef = IsTableColExists(Table, ColName)
            End If
            
        Else
            IsValidRef = False
        End If
        
    Else
        IsValidRef = False
    End If
    
    IsTableColumnRef = IsValidRef
    Logger.Log TRACE_LOG, "Exit RangeResolver.IsTableColumnRef"
    
End Function

Private Function GetTable(ByVal FromBook As Workbook, ByVal TableName As String) As ListObject
    
    Logger.Log TRACE_LOG, "Enter RangeResolver.GetTable"
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In FromBook.Worksheets
        
        Dim CurrentTable As ListObject
        For Each CurrentTable In CurrentSheet.ListObjects
            If CurrentTable.Name = TableName Then
                Set GetTable = CurrentTable
                Logger.Log TRACE_LOG, "Exit Due to Exit Keyword RangeResolver.GetTable"
                Exit Function
            End If
        Next CurrentTable
        
    Next CurrentSheet
    Logger.Log TRACE_LOG, "Exit RangeResolver.GetTable"
    
End Function

Private Function IsTableColExists(ByVal Table As ListObject, ByVal ColName As String) As Boolean
    
    Logger.Log TRACE_LOG, "Enter RangeResolver.IsTableColExist"
    IsTableColExists = False
    Dim CurrentCol As ListColumn
    For Each CurrentCol In Table.ListColumns
        If CurrentCol.Name = ColName Then
            IsTableColExists = True
            Logger.Log TRACE_LOG, "Exit Due to Exit Keyword RangeResolver.IsTableColExist"
            Exit Function
        End If
    Next CurrentCol
    Logger.Log TRACE_LOG, "Exit RangeResolver.IsTableColExist"
    
End Function


