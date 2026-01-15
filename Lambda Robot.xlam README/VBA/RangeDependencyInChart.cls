VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeDependencyInChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, ImplicitActiveSheetReference
'@PredeclaredId

Option Explicit

Private Type TRangeDependencyInChart
    Book As Workbook
    BeforeRangePrefix As String
    AllDependency As Collection
    PutDependencyOnTable As ListObject
End Type

Private this As TRangeDependencyInChart

Public Property Get PutDependencyOnTable() As ListObject
    Set PutDependencyOnTable = this.PutDependencyOnTable
End Property

Public Property Set PutDependencyOnTable(ByVal RHS As ListObject)
    Set this.PutDependencyOnTable = RHS
End Property

Public Property Get AllDependency() As Collection
    Set AllDependency = this.AllDependency
End Property

Public Property Set AllDependency(ByVal RHS As Collection)
    Set this.AllDependency = RHS
End Property

Public Property Get Book() As Workbook
    Set Book = this.Book
End Property

Public Property Set Book(ByVal RHS As Workbook)
    Set this.Book = RHS
End Property

Public Property Get BeforeRangePrefix() As String
    BeforeRangePrefix = this.BeforeRangePrefix
End Property

Public Property Let BeforeRangePrefix(ByVal RHS As String)
    this.BeforeRangePrefix = RHS
End Property

Private Sub UpdateBasicStat(ByVal GivenRange As Range)
    
    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.UpdateBasicStat"
    ' Set up basic statistics for the given range
    
    Set this.Book = GivenRange.Worksheet.Parent
    this.BeforeRangePrefix = Text.RemoveFromEndIfPresent( _
                             GetSheetRefForRangeReference(GivenRange.Worksheet.Name) _
                             , SHEET_NAME_SEPARATOR)
    Set this.AllDependency = New Collection
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.UpdateBasicStat"
    
End Sub

Private Sub ClearCatches()
    
    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.ClearCatches"
    ' Clear all the cached variables and objects
    
    Set this.Book = Nothing
    this.BeforeRangePrefix = vbNullString
    Set this.AllDependency = Nothing
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.ClearCatches"

End Sub

Public Sub ApplyNameRange(ByVal CurrentName As Name)
    
    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.ApplyNameRange"
    ' Apply the named range to the charts that reference it
    
    Dim RefersToRange As Range
    Set RefersToRange = CurrentName.RefersToRange
    FindAllDependency RefersToRange
    If this.AllDependency.Count = 0 Then Exit Sub
    Dim ReplaceWith As String
    ReplaceWith = GetNamedRangeRefText(CurrentName)
    Dim CurrentItem As ChartInfo
    For Each CurrentItem In this.AllDependency
        If CurrentItem.IsWholeSearchRangeBeingUsed Then
            Dim FindAddress As String
            If CurrentItem.OnSheet = RefersToRange.Worksheet.Name Then
                FindAddress = GetSheetRefForRangeReference(CurrentItem.OnSheet) & CurrentItem.RangeAddress
            Else
                FindAddress = CurrentItem.RangeAddress
            End If
            Dim PropertyValue As String
            PropertyValue = CallByName(CurrentItem.FoundOnObject, CurrentItem.PropertyName, VbGet)
            CallByName CurrentItem.FoundOnObject, CurrentItem.PropertyName _
                                                 , VbLet, VBA.Replace(PropertyValue, FindAddress, ReplaceWith)
        End If
    Next CurrentItem
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.ApplyNameRange"
    
End Sub

Private Function GetNamedRangeRefText(ByVal CurrentName As Name) As String
    
    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.GetNamedRangeRefText"
    ' Get the reference text for the named range based on its scope (local or global)
    
    Dim ReplaceWith As String
    If IsLocalScopeNamedRange(CurrentName.NameLocal) Then
        ReplaceWith = CurrentName.NameLocal
    Else
        Dim WorkbookName As String
        WorkbookName = CurrentName.RefersToRange.Worksheet.Parent.Name
        If IsAnyNonAlphanumeric(WorkbookName) Then
            ReplaceWith = SINGLE_QUOTE & EscapeSingeQuote(WorkbookName) _
                          & SINGLE_QUOTE & EXCLAMATION_SIGN & CurrentName.Name
        Else
            ReplaceWith = WorkbookName & EXCLAMATION_SIGN & CurrentName.Name
        End If
    End If
    GetNamedRangeRefText = ReplaceWith
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.GetNamedRangeRefText"
    
End Function

Private Sub FindAllDependency(ByVal GivenRange As Range)

    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.FindAllDependency"
    ' Find all the chart objects that reference the given range and update basic statistics
    
    UpdateBasicStat GivenRange
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In Book.Worksheets
        ProcessSheetForSeries CurrentSheet, GivenRange
    Next CurrentSheet
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.FindAllDependency"
    
End Sub

Public Sub SendDataToSheet(ByVal GivenRange As Range, ByVal PutToRange As Range)
    
    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.SendDataToSheet"
    ' Send the dependency data to a specified range on a sheet
    
    FindAllDependency GivenRange
    If this.AllDependency.Count = 0 Then
        MsgBox "No Reference found on any chart", vbInformation, "SendDataToSheet"
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword RangeDependencyInChart.SendDataToSheet"
        Exit Sub
    End If
    Dim AllData As Variant
    AllData = GetDependencyData(this.AllDependency, True)
    Dim PutDependencyDataOnRange As Range
    Set PutDependencyDataOnRange = PutToRange.Resize(AllDependency.Count + 1, UBound(AllData, 2))
    If Not IsAllCellBlank(PutDependencyDataOnRange) Then
        MsgBox "Unable to write dependency info to worksheet. Existing data in range " _
               & PutDependencyDataOnRange.Address & _
               " would be overwritten. Please clear enough space and run again.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword RangeDependencyInChart.SendDataToSheet"
        Exit Sub
    End If
    
    PutDependencyDataOnRange.Value = AllData
    Set this.PutDependencyOnTable = PutToRange.Worksheet.ListObjects.Add(xlSrcRange, PutDependencyDataOnRange, , xlYes)
    this.PutDependencyOnTable.ListColumns(3).Range.ColumnWidth = 50
    this.PutDependencyOnTable.ListColumns(5).Range.ColumnWidth = 50
    this.PutDependencyOnTable.Range.WrapText = False
    ScrollToDependencyDataRange this.PutDependencyOnTable
    ClearCatches
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.SendDataToSheet"
    
End Sub

Private Function GetDependencyData(ByVal AllDependency As Collection, ByVal IsWithHeader As Boolean) As Variant
    
    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.GetDependencyData"
    ' Prepare the dependency data in a 2D array
    
    Dim AllData As Variant
    If IsWithHeader Then
        ReDim AllData(0 To AllDependency.Count, 1 To 6)
        AllData(0, 1) = "On Sheet"
        AllData(0, 2) = "Chart Object Name"
        AllData(0, 3) = "Formula"
        AllData(0, 4) = "VBA Syntax"
        AllData(0, 5) = "Is Whole Search Range Being Used?"
        AllData(0, 6) = "Range Address"
    Else
        ReDim AllData(1 To AllDependency.Count, 1 To 6)
    End If
    
    Dim CurrentChartInfo As ChartInfo
    Dim CurrentItemIndex As Long
    For CurrentItemIndex = 1 To AllDependency.Count
        Set CurrentChartInfo = AllDependency.Item(CurrentItemIndex)
        AllData(CurrentItemIndex, 1) = CurrentChartInfo.OnSheet
        AllData(CurrentItemIndex, 2) = CurrentChartInfo.ChartObjectName
        AllData(CurrentItemIndex, 3) = SINGLE_QUOTE & CurrentChartInfo.Formula
        AllData(CurrentItemIndex, 4) = CurrentChartInfo.VBASyntax
        AllData(CurrentItemIndex, 5) = CurrentChartInfo.IsWholeSearchRangeBeingUsed
        AllData(CurrentItemIndex, 6) = CurrentChartInfo.RangeAddress
    Next CurrentItemIndex
    GetDependencyData = AllData
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.GetDependencyData"
    
End Function

Private Sub ProcessSheetForSeries(ByVal CurrentSheet As Worksheet, ByVal GivenRange As Range)
    
    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.ProcessSheetForSeries"
    ' Process the worksheet for chart series
    
    Dim CurrentChartObject As ChartObject
    For Each CurrentChartObject In CurrentSheet.ChartObjects
        Dim CurrentSeries As Series
        Dim Counter As Long
        For Counter = 1 To CurrentChartObject.Chart.SeriesCollection.Count
            Set CurrentSeries = CurrentChartObject.Chart.SeriesCollection.Item(Counter)
            Dim VBASyntax As String
            VBASyntax = GetChartSeriesProperty(CurrentSheet.CodeName, CurrentChartObject.Name, "FormulaLocal", Counter)
            Dim FormulaText As String
            FormulaText = GetSeriesFormulaIfErrorVBNullString(CurrentSeries)
            FindAndCreateDependency FormulaText, GivenRange, CurrentSheet, VBASyntax, CurrentChartObject, CurrentSeries
            ProcessSheetForDataLabels CurrentSheet, GivenRange, CurrentChartObject, CurrentSeries, Counter
        Next Counter
    Next CurrentChartObject
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.ProcessSheetForSeries"
    
End Sub

Private Sub FindAndCreateDependency(ByVal FormulaText As String, ByVal GivenRange As Range _
                                                                , ByVal CurrentSheet As Worksheet _
                                                                 , ByVal VBASyntax As String _
                                                                  , ByVal CurrentChartObject As ChartObject _
                                                                   , ByVal FoundOnObject As Object)
    
    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.FindAndCreateDependency"
    ' Find and create the dependency info for the given formula text
    
    If Text.Contains(FormulaText, this.BeforeRangePrefix) Then
        Dim AllRangeAddress As Variant
        AllRangeAddress = GetDirectPrecedents(FormulaText, GivenRange.Worksheet)
        Dim TempRange As Range
        Dim RangeAddress As Variant
        
        For Each RangeAddress In AllRangeAddress
            If RangeAddress <> vbNullString Then
                Set TempRange = RangeResolver.GetRange(CStr(RangeAddress))
                If HasIntersection(TempRange, GivenRange) Then
                    Dim CurrentChartInfo As ChartInfo
                    Set CurrentChartInfo = CreateNewChartInfo(CurrentSheet, CurrentChartObject _
                                                                           , TempRange.Address = GivenRange.Address _
                                                                            , CStr(RangeAddress) _
                                                                             , FoundOnObject, "FormulaLocal", VBASyntax)
                    this.AllDependency.Add CurrentChartInfo
                End If
            End If
        Next RangeAddress
    End If
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.FindAndCreateDependency"
    
End Sub

Public Sub ProcessSheetForDataLabels(ByVal CurrentSheet As Worksheet, ByVal GivenRange As Range _
                                                                     , ByVal CurrentChartObject As ChartObject _
                                                                      , ByVal CurrentSeries As Series _
                                                                       , ByVal SeriesCounter As Long)
    
    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.ProcessSheetForDataLabels"
    ' Process the worksheet for data labels in the chart series
    
    If GivenRange.Cells.CountLarge > 1 Then Exit Sub
    If Not CurrentSeries.HasDataLabels Then Exit Sub
    Dim Counter As Long
    For Counter = 1 To CurrentSeries.DataLabels.Count
        Dim CurrentDataLabel As DataLabel
        Set CurrentDataLabel = CurrentSeries.DataLabels(Counter)
        Dim VBASyntax As String
        VBASyntax = GetChartSeriesProperty(CurrentSheet.CodeName, CurrentChartObject.Name _
                                                                 , "DataLabels(" & Counter & ").FormulaLocal" _
                                                                  , SeriesCounter)
        FindAndCreateDependency CurrentDataLabel.FormulaLocal, GivenRange _
                                                              , CurrentSheet _
                                                               , VBASyntax _
                                                                , CurrentChartObject _
                                                                 , CurrentDataLabel
    Next Counter
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.ProcessSheetForDataLabels"
    
End Sub

Private Function GetSeriesFormulaIfErrorVBNullString(ByVal CurrentSeries As Series) As String
    
    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.GetSeriesFormulaIfErrorVBNullString"
    ' Get the formula of the series or return empty string in case of an error
    
    On Error Resume Next
    GetSeriesFormulaIfErrorVBNullString = CurrentSeries.FormulaLocal
    On Error GoTo 0
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.GetSeriesFormulaIfErrorVBNullString"
    
End Function

Private Function CreateNewChartInfo(ByVal CurrentSheet As Worksheet, ByVal CurrentChartObject As ChartObject _
                                                                    , ByVal IsWholeSearchRangeBeingUsed As Boolean _
                                                                     , ByVal RangeAddress As String _
                                                                      , ByVal FoundOnObject As Object _
                                                                       , ByVal PropertyName As String _
                                                                        , ByVal VBASyntax As String) As ChartInfo
    
    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.CreateNewChartInfo"
    ' Create a new ChartInfo object with the provided information
    
    Dim CurrentChartInfo As ChartInfo
    Set CurrentChartInfo = New ChartInfo
    With CurrentChartInfo
        .OnSheet = CurrentSheet.Name
        .ChartObjectName = CurrentChartObject.Name
        .Formula = CallByName(FoundOnObject, PropertyName, VbGet)
        .VBASyntax = VBASyntax
        .IsWholeSearchRangeBeingUsed = IsWholeSearchRangeBeingUsed
        .RangeAddress = RangeAddress
        .PropertyName = PropertyName
        Set .FoundOnObject = FoundOnObject
    End With
    
    Set CreateNewChartInfo = CurrentChartInfo
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.CreateNewChartInfo"

End Function

Public Function GetChartSeriesProperty(ByVal SheetCodeName As String _
                                       , ByVal ChartObjectName As String _
                                        , ByVal PropertyName As String, ByVal Counter As Long) As String
    
    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.GetChartSeriesProperty"
    ' Get the property of the chart series
    
    If Counter = 0 Then
        GetChartSeriesProperty = Me.GetChartProperty(SheetCodeName, ChartObjectName, "SeriesCollection." & PropertyName)
    Else
        GetChartSeriesProperty = Me.GetChartProperty(SheetCodeName, ChartObjectName, "SeriesCollection.Item(" & Counter & ")." & PropertyName)
    End If
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.GetChartSeriesProperty"
    
End Function

Public Function GetChartProperty(ByVal SheetCodeName As String _
                                 , ByVal ChartObjectName As String _
                                  , ByVal PropertyName As String) As String
    
    Logger.Log TRACE_LOG, "Enter RangeDependencyInChart.GetChartProperty"
    ' Get the chart property
    
    GetChartProperty = SheetCodeName & ".ChartObjects(""" & ChartObjectName & QUOTES_AND_FIRST_PAREN_CLOSE _
                       & ".Chart." & PropertyName
    Logger.Log TRACE_LOG, "Exit RangeDependencyInChart.GetChartProperty"
End Function


