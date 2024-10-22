VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed, ImplicitActiveWorkbookReference, SuperfluousAnnotationArgument, UnrecognizedAnnotation, ProcedureNotUsed
'@Folder "Lambda.Editor"
Option Explicit

Public Enum OperationType
    DEPENDENCY_INFO_GENERATION = 1
    LET_STATEMENT_GENERATION = 2
    LAMBDA_STATEMENT_GENERATION = 3
    AFE_SCRIPT_GENERATION = 4
End Enum

Private Type TFormulaParser
    PrecedencyExtractor As FormulaPrecedencyExtractor
    CurrentLambdaInfo As LambdaInfo
    LambdaInvocationArgument As String
    DependencySearchInRegion As Range
    AFEScriptCode As Variant
    IsAddToNameManager As Boolean
    IsExportable As Boolean
    PutDependencyOnTable As ListObject
    IsProcessTerminatedByUser As Boolean
    IsLetNeededInLambda As Boolean
    InitialDependencyDataForReset As Variant
    FormulaGen As LETORLAMBDAGenerator
End Type

Private This  As TFormulaParser

Public Property Get PrecedencyExtractor() As FormulaPrecedencyExtractor
    Set PrecedencyExtractor = This.PrecedencyExtractor
End Property

Public Property Get DependencyDataForReset() As Collection
    Set DependencyDataForReset = This.PrecedencyExtractor.GetDependencyObjectsFromData(This.InitialDependencyDataForReset)
End Property

Public Property Get IsLetNeededInLambda() As Boolean
    IsLetNeededInLambda = This.IsLetNeededInLambda
End Property

Public Property Let IsLetNeededInLambda(ByVal RHS As Boolean)
    This.IsLetNeededInLambda = RHS
End Property

Public Property Let IsProcessTerminatedByUser(RHS As Boolean)
    This.IsProcessTerminatedByUser = RHS
End Property

Public Property Get IsProcessTerminatedByUser() As Boolean
    IsProcessTerminatedByUser = This.IsProcessTerminatedByUser
End Property

Public Property Get PutDependencyOnTable() As ListObject
    Set PutDependencyOnTable = This.PutDependencyOnTable
End Property

Public Property Set PutDependencyOnTable(ByVal RHS As ListObject)
    Set This.PutDependencyOnTable = RHS
End Property

Public Property Get IsExportable() As Boolean
    IsExportable = This.IsExportable
End Property

Public Property Let IsExportable(ByVal RHS As Boolean)
    This.IsExportable = RHS
End Property

Public Property Get IsAddToNameManager() As Boolean
    IsAddToNameManager = This.IsAddToNameManager
End Property

Public Property Let IsAddToNameManager(ByVal RHS As Boolean)
    This.IsAddToNameManager = RHS
End Property

Public Property Get GetLambdaInfo() As LambdaInfo
    Set GetLambdaInfo = This.CurrentLambdaInfo
End Property

'@EntryPoint
Public Sub CreateLambdaDetails(ByVal FormulaStartCell As Range _
                               , ByVal PutDependencyInfoOnRange As Range _
                                , ByVal TypeOfOperation As OperationType _
                                 , ByVal DependencySearchInRegion As Range)
                                 
    Logger.Log TRACE_LOG, "Enter FormulaParser.CreateLambdaDetails"
    ' This sub creates Lambda details by generating dependency information based on the provided parameters.

    ' Check if the FormulaStartCell is not provided (i.e., it is nothing)
    If IsNothing(FormulaStartCell) Then
        MsgBox "Unable to generate dependency info. Formula cell not provided.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaParser.CreateLambdaDetails"
        Exit Sub
    End If

    ' Check if PutDependencyInfoOnRange is not provided when TypeOfOperation is DEPENDENCY_INFO_GENERATION
    If IsNothing(PutDependencyInfoOnRange) And TypeOfOperation = DEPENDENCY_INFO_GENERATION Then
        MsgBox "Unable to generate dependency info. Destination range not provided." _
               , vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaParser.CreateLambdaDetails"
        Exit Sub
    End If

    ' Set the DependencySearchInRegion property to the provided DependencySearchInRegion range
    Set This.DependencySearchInRegion = DependencySearchInRegion

    ' Clear any previous catches (if any)
    ClearCatches

    ' Create basic info (like Lambda name and formula text) from the FormulaStartCell
    CreateBasicInfoFromFormulaStartCell

    ' Generate the dependency information based on the FormulaStartCell, PutDependencyInfoOnRange, and TypeOfOperation
    GenerateDependencyInfo FormulaStartCell, PutDependencyInfoOnRange, TypeOfOperation
    Logger.Log TRACE_LOG, "Exit FormulaParser.CreateLambdaDetails"

End Sub

Private Sub ClearCatches()

    Logger.Log TRACE_LOG, "Enter FormulaParser.ClearCatches"
    ' This sub clears any previously set catch variables and objects.

    ' Set the PrecedencyExtractor and CurrentLambdaInfo objects to nothing, releasing any references.
    Set This.PrecedencyExtractor = Nothing
    Set This.CurrentLambdaInfo = Nothing

    ' Set the LambdaInvocationArgument to an empty string (vbNullString).
    This.LambdaInvocationArgument = vbNullString

    Logger.Log TRACE_LOG, "Exit FormulaParser.ClearCatches"
    
End Sub

Private Sub CreateBasicInfoFromFormulaStartCell()

    Logger.Log TRACE_LOG, "Enter FormulaParser.CreateBasicInfoFromFormulaStartCell"
    ' This sub creates basic information from the FormulaStartCell.

    ' Create a new FormulaPrecedencyExtractor and assign it to This.PrecedencyExtractor.
    Set This.PrecedencyExtractor = New FormulaPrecedencyExtractor
    Set This.FormulaGen = New LETORLAMBDAGenerator
    ' Create a new LambdaInfo and assign it to This.CurrentLambdaInfo.
    Set This.CurrentLambdaInfo = New LambdaInfo
    Logger.Log TRACE_LOG, "Exit FormulaParser.CreateBasicInfoFromFormulaStartCell"

End Sub

Private Sub GenerateDependencyInfo(ByVal FormulaStartCell As Range _
                                   , ByVal PutDependencyInfoOnRange As Range _
                                    , ByVal TypeOfOperation As OperationType)
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.GenerateDependencyInfo"
    ' This sub generates dependency information based on the provided formula.

    On Error GoTo ResetCalculation
    
    With This.PrecedencyExtractor
        ' Extract all precedence information using the FormulaPrecedencyExtractor.
        .ExtractAllPrecedencyInfo FormulaStartCell, This.DependencySearchInRegion, Nothing
        If TypeOfOperation = LET_STATEMENT_GENERATION Then .MarkAllAsNonInputCells .AllDependency
        ' Add underscore on valid variable names for non-input cells.
        .AddPrefixOnValidVarNameForNonInputCells .AllDependency
        .UpdateValidVarForUsedNamedRanges .AllDependency
    End With
    
    ' Check the type of operation requested.
    If TypeOfOperation = DEPENDENCY_INFO_GENERATION Then
        ' If it's a dependency info generation, send the dependency data to the sheet.
        SendDependencyDataToSheet This.PrecedencyExtractor.AllDependency, PutDependencyInfoOnRange
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaParser.GenerateDependencyInfo"
        Exit Sub
    End If
    
    ' Get the dependency data and sort it.
    This.PrecedencyExtractor.DependencyData = GetAllDependencyToArrayAndSort(This.PrecedencyExtractor.AllDependency)

    ' Store the initial dependency data for resetting purposes.
    This.InitialDependencyDataForReset = This.PrecedencyExtractor.DependencyData

    ' Create Let and Lambda statements based on the dependency data.
    CreateLetAndLambdaStatements This.PrecedencyExtractor.DependencyData, FormulaStartCell, TypeOfOperation

    ' If the process is terminated by the user, exit the sub.
    If This.IsProcessTerminatedByUser Then Exit Sub

    ' Log the generated lambda details.
    LogGeneratedLambdaDetails

    ' Convert the full Let formula to AFEScript code.
    This.CurrentLambdaInfo.AFEScriptCode = ConvertFullLetToAFEScript(This.CurrentLambdaInfo.LetFormula)

    ' If the type of operation is AFE script generation, put the AFE code to the sheet.
    If TypeOfOperation = AFE_SCRIPT_GENERATION Then
        PutAFECodeToSheet PutDependencyInfoOnRange, This.CurrentLambdaInfo.AFEScriptCode
    End If

    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaParser.GenerateDependencyInfo"
    Exit Sub
    Resume

ResetCalculation:
    ' Handle errors and exceptions.

    If Err.Description = "Formula Lenth Cross Excel Limit" Then
        MsgBox "Unable to generate Lambda formula. Resulting formula too long.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaParser.GenerateDependencyInfo"
        Exit Sub
    End If

    ' If there is an error other than the formula length, raise the error.
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
        Resume
    End If

    ' Close the log file and log the exiting of the function.
    Logger.CloseLogFile
    Logger.Log TRACE_LOG, "Exit FormulaParser.GenerateDependencyInfo"

End Sub

Public Sub RecalculatePrecedencyAgain(ByVal UserModifiedDependency As Collection, ByVal TypeOfOperation As OperationType)
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.RecalculatePrecedencyAgain"
    ' This sub recalculates the precedence again after user modification of dependencies.
    
    With This.PrecedencyExtractor
        ' Extract all precedence information again, considering the user-modified dependencies.
        .ExtractAllPrecedencyInfo .FormulaStartCell, _
                                  This.DependencySearchInRegion, _
                                  UserModifiedDependency
                                  
        If TypeOfOperation = LET_STATEMENT_GENERATION Then .MarkAllAsNonInputCells .AllDependency
        ' Add an underscore on valid variable names for non-input cells.
        .AddPrefixOnValidVarNameForNonInputCells .AllDependency

        ' Get the dependency data and sort it.
        .DependencyData = GetAllDependencyToArrayAndSort(.AllDependency)

        ' Reorder the dependency data to accommodate user-modified dependencies.
        .DependencyData = ReorderForUserModifiedDependency(.DependencyData, UserModifiedDependency)

        ' Get the updated dependency objects collection.
        Set .AllDependency = .GetDependencyObjectsFromData(.DependencyData)
    End With
    Logger.Log TRACE_LOG, "Exit FormulaParser.RecalculatePrecedencyAgain"

End Sub

Private Function ReorderForUserModifiedDependency(ByVal DependencyData As Variant _
                                                  , ByVal UserModifiedDependency As Collection) As Variant
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.ReorderForUserModifiedDependency"
    ' This function reorders the dependency data to accommodate user-modified dependencies.

    ReorderForUserModifiedDependency = DependencyData
    Dim ModifiedParametersVarNameAndRangeRef As Variant
    ModifiedParametersVarNameAndRangeRef = GetInputCellsVarNameAndRangeReference(UserModifiedDependency)

    ' Remove any dependency that is not present in the modified parameters.
    ModifiedParametersVarNameAndRangeRef = RemoveNotPresentDependency(ModifiedParametersVarNameAndRangeRef _
                                                                      , DependencyData)

    ' Check if any previous parameter is found.
    If Not IsArray(ModifiedParametersVarNameAndRangeRef) Then
        Logger.Log DEBUG_LOG, "No previous parameter is found."
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaParser.ReorderForUserModifiedDependency"
        Exit Function
    End If

    ' If only one param then no need to think of the order.
    If LBound(ModifiedParametersVarNameAndRangeRef, 1) = UBound(ModifiedParametersVarNameAndRangeRef, 1) Then
        Logger.Log DEBUG_LOG, "Only one parameter cell found."
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaParser.ReorderForUserModifiedDependency"
        Exit Function
    End If

    Dim FirstColumnIndex As Long
    FirstColumnIndex = LBound(ModifiedParametersVarNameAndRangeRef, 2)
    Dim CurrentRowIndex As Long

    ' Loop through the modified parameters and check if reorder is required.
    For CurrentRowIndex = LBound(ModifiedParametersVarNameAndRangeRef, 1) To _
        UBound(ModifiedParametersVarNameAndRangeRef, 1) - 1
        
        Dim FirstRangeRef As String
        FirstRangeRef = ModifiedParametersVarNameAndRangeRef(CurrentRowIndex, FirstColumnIndex + 1)

        Dim SecondRangeRef As String
        SecondRangeRef = ModifiedParametersVarNameAndRangeRef(CurrentRowIndex + 1, FirstColumnIndex + 1)

        Dim FirstRangeRefIndex As Long
        FirstRangeRefIndex = FirstIndexOf(DependencyData, FirstRangeRef, 2, True)

        Dim SecondRangeRefIndex As Long
        SecondRangeRefIndex = FirstIndexOf(DependencyData, SecondRangeRef, 2, True)

        ' Check if reorder is required or not.
        If FirstRangeRefIndex > SecondRangeRefIndex _
           And FirstRangeRefIndex <> -1 And SecondRangeRefIndex <> -1 Then
            ' We need to bring
            MoveDownRowInPlace DependencyData, SecondRangeRefIndex, FirstRangeRefIndex - 1
        End If

    Next CurrentRowIndex

    ReorderForUserModifiedDependency = DependencyData
    Logger.Log TRACE_LOG, "Exit FormulaParser.ReorderForUserModifiedDependency"
    
End Function

Private Sub MoveDownRowInPlace(OnArray As Variant, ByVal RowIndex As Long, ByVal MoveDownAfterIndex As Long)
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.MoveDownRowInPlace"
    Dim Index As Long
    For Index = RowIndex To MoveDownAfterIndex
        SwapTwoRowsInPlace OnArray, Index, Index + 1
    Next Index
    Logger.Log TRACE_LOG, "Exit FormulaParser.MoveDownRowInPlace"
    
End Sub

Private Function RemoveNotPresentDependency(ByVal ModifiedParametersVarNameAndRangeRef As Variant _
                                            , ByVal DependencyData As Variant) As Variant
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.RemoveNotPresentDependency"
    ' This function removes any dependency that is not present in the modified parameters.

    If Not IsArray(ModifiedParametersVarNameAndRangeRef) Then
        RemoveNotPresentDependency = vbEmpty
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaParser.RemoveNotPresentDependency"
        Exit Function
    End If

    ' Create a collection to store the present rows.
    Dim PresentRows As Collection
    Set PresentRows = New Collection

    Dim FirstColumnIndex As Long
    FirstColumnIndex = LBound(ModifiedParametersVarNameAndRangeRef, 2)

    Dim CurrentRowIndex As Long

    ' Loop through the modified parameters to find the present rows in the DependencyData.
    For CurrentRowIndex = LBound(ModifiedParametersVarNameAndRangeRef, 1) To _
        UBound(ModifiedParametersVarNameAndRangeRef, 1)
        Dim Index As Long
        Index = FirstIndexOf(DependencyData _
                             , CStr(ModifiedParametersVarNameAndRangeRef(CurrentRowIndex, FirstColumnIndex + 1)) _
                              , 2, True)
        If Index <> -1 Then
            PresentRows.Add CurrentRowIndex, CStr(CurrentRowIndex)
        End If
    Next CurrentRowIndex

    ' Check if no present rows found.
    If PresentRows.Count = 0 Then
        RemoveNotPresentDependency = vbEmpty
    Else
        Dim Result As Variant
        ReDim Result(1 To PresentRows.Count, 1 To 2)

        Dim Counter As Long
        Dim CurrentIndex As Variant

        ' Populate the Result array with present rows.
        For Each CurrentIndex In PresentRows
            Counter = Counter + 1
            Result(Counter, 1) = ModifiedParametersVarNameAndRangeRef(CLng(CurrentIndex), FirstColumnIndex)
            Result(Counter, 2) = ModifiedParametersVarNameAndRangeRef(CLng(CurrentIndex), FirstColumnIndex + 1)
        Next CurrentIndex

        RemoveNotPresentDependency = Result
    End If
    Logger.Log TRACE_LOG, "Exit FormulaParser.RemoveNotPresentDependency"
    
End Function

Private Sub LogGeneratedLambdaDetails()

    Logger.Log TRACE_LOG, "Enter FormulaParser.LogGeneratedLambdaDetails"
    ' Logs the generated Let and Lambda formulas to the debug log.

    ' Log the generated Let formula.
    Logger.Log DEBUG_LOG, THREE_NEW_LINE
    Logger.Log DEBUG_LOG, This.CurrentLambdaInfo.LetFormula
    Logger.Log DEBUG_LOG, THREE_NEW_LINE

    ' Log the generated Lambda formula.
    Logger.Log DEBUG_LOG, This.CurrentLambdaInfo.LambdaFormula
    Logger.Log TRACE_LOG, "Exit FormulaParser.LogGeneratedLambdaDetails"

End Sub

Private Sub PutAFECodeToSheet(ByVal FormulaStartCell As Range, ByVal AFEData As Variant)
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.PutAFECodeToSheet"
    ' Writes the generated AFE Script data to the worksheet.

    ' Resize the target range to fit the AFEData array.
    Dim PutAFEDataOnRange As Range
    Set PutAFEDataOnRange = FormulaStartCell.Resize(UBound(AFEData, 1))

    ' Check if there is any existing data in the target range.
    ' If data is present, show a message box indicating the need to clear space and exit the subroutine.
    If Not IsAllCellBlank(PutAFEDataOnRange) Then
        MsgBox "Unable to write AFE Script to worksheet. You have data where you want to store the AFE Script." _
               & " Clear enough space and then run again. Needed Range : " _
               & PutAFEDataOnRange.Address, vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaParser.PutAFECodeToSheet"
        Exit Sub
    End If

    ' Write the AFEData array to the target range.
    PutAFEDataOnRange.Value = AFEData
    Logger.Log TRACE_LOG, "Exit FormulaParser.PutAFECodeToSheet"

End Sub

Private Sub SendDependencyDataToSheet(ByVal AllDependency As Collection, ByVal FirstCellToPutDependencyData As Range)
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.SendDependencyDataToSheet"
    ' Sends the sorted dependency data to the worksheet starting from the specified range.

    ' Sort the dependency data and convert it back to an object collection.
    Dim DependencyData As Variant
    DependencyData = GetAllDependencyToArrayAndSort(AllDependency)
    Set AllDependency = This.PrecedencyExtractor.GetDependencyObjectsFromData(DependencyData)
    DependencyData = GetAllDependencyToArrayAndSort(AllDependency)

    ' Resize the target range to fit the dependency data.
    Dim PutDependencyDataOnRange As Range
    Set PutDependencyDataOnRange = FirstCellToPutDependencyData.Resize(AllDependency.Count + 1 _
                                                                       , UBound(DependencyData, 2))

    ' Check if there is any existing data in the target range.
    ' If data is present, show a message box indicating the need to clear space and exit the subroutine.
    If Not IsAllCellBlank(PutDependencyDataOnRange) Then
        MsgBox "Unable to write dependency info to worksheet. Existing data in range " _
               & PutDependencyDataOnRange.Address & _
               " would be overwritten. Please clear enough space and run again.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaParser.SendDependencyDataToSheet"
        Exit Sub
    End If

    ' Set the number format of the target range to text to preserve any leading zeros.
    PutDependencyDataOnRange.NumberFormat = "@"

    ' Write the dependency data array to the target range.
    PutDependencyDataOnRange.Value = DependencyData

    ' Create a ListObject (Excel Table) from the target range to enable filtering and sorting.
    Set This.PutDependencyOnTable = This.PrecedencyExtractor.FormulaStartCellInSheet.ListObjects.Add(xlSrcRange _
                                                                                                     , PutDependencyDataOnRange, , xlYes)

    ' Set the column width of the FORMULATEXT_COLUMN (4th column) to make the formulas visible.
    Const FORMULATEXT_COLUMN As Long = 4
    This.PutDependencyOnTable.ListColumns(FORMULATEXT_COLUMN).Range.ColumnWidth = 50

    ' Disable text wrapping in the table range to display long formulas without truncation.
    This.PutDependencyOnTable.Range.WrapText = False

    ' Scroll to the table range to make it visible to the user.
    ScrollToDependencyDataRange This.PutDependencyOnTable

    ' Store the table reference in the CurrentLambdaInfo object for future use.
    Set This.CurrentLambdaInfo.PutDependencyOnTable = This.PutDependencyOnTable
    Logger.Log TRACE_LOG, "Exit FormulaParser.SendDependencyDataToSheet"

End Sub

Private Function GenerateDependencyStats(ByVal AllDependency As Collection) As DependenciesStat
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.GenerateDependencyStats"
    ' Generates statistics about the dependencies.

    ' Create a DependenciesStat object to store the statistics.
    Dim Stats As DependenciesStat
    Set Stats = New DependenciesStat

    ' Loop through each DependencyInfo in the AllDependency collection.
    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In AllDependency

        ' Update the statistics for input cell labels.
        Stats.UpdateInputCellStats CurrentDependencyInfo.IsLabelAsInputCell

        ' Update the dependency level statistics.
        Stats.UpdateLevel CurrentDependencyInfo.Level

    Next CurrentDependencyInfo

    ' Set the return value to the generated statistics and release the Stats object.
    Set GenerateDependencyStats = Stats
    Set Stats = Nothing
    Logger.Log TRACE_LOG, "Exit FormulaParser.GenerateDependencyStats"

End Function

Private Sub CreateLetAndLambdaStatements(ByVal DependencyOrderData As Variant _
                                         , ByVal FormulaStartCell As Range _
                                          , ByVal TypeOfOperation As OperationType)
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.CreateLetAndLambdaStatements"
    ' Creates Let and Lambda statements based on the given dependency order data.

    ' Get the collection of DependencyInfo objects from the DependencyOrderData.
    Dim DependencyObjects As Collection
    Set DependencyObjects = This.PrecedencyExtractor.GetDependencyObjectsFromData(DependencyOrderData)
    
    If TypeOfOperation = LET_STATEMENT_GENERATION Then
        PrecedencyExtractor.AddPrefixOnInputCells DependencyObjects
    End If
    
    ' Update input cells from the user's choice depending on the TypeOfOperation.
    UpdateInputCellsFromUserChoice DependencyObjects, TypeOfOperation
    ActivateFormulaStartCell This.PrecedencyExtractor
    If This.IsProcessTerminatedByUser Then Exit Sub
    
    Dim FormulaGen As LETORLAMBDAGenerator
    Set FormulaGen = New LETORLAMBDAGenerator
    FormulaGen.CreateLetAndLambdaStatements DependencyObjects, FormulaStartCell, TypeOfOperation, This.PrecedencyExtractor
    This.IsLetNeededInLambda = This.FormulaGen.IsLetNeededInLambda
    
    AddLambdaToNameManagerIfRequested FormulaGen.LambdaDef
    This.CurrentLambdaInfo.InvocationArgument = FormulaGen.LambdaInvocation

    This.CurrentLambdaInfo.LetFormula = FormatFormula(FormulaGen.LetFormula)

    Logger.Log TRACE_LOG, "Exit FormulaParser.CreateLetAndLambdaStatements"
    
End Sub

Private Sub ActivateFormulaStartCell(ByVal PrecedencyExtractor As FormulaPrecedencyExtractor)

    Logger.Log TRACE_LOG, "Enter FormulaParser.ActivateFormulaStartCell"
    ' Check if the FormulaStartCell exists.
    If IsNotNothing(PrecedencyExtractor.FormulaStartCell) Then
        ' Activate the worksheet containing the FormulaStartCell.
        PrecedencyExtractor.FormulaStartCell.Worksheet.Activate

        ' Select the FormulaStartCell.
        PrecedencyExtractor.FormulaStartCell.Select
    End If
    Logger.Log TRACE_LOG, "Exit FormulaParser.ActivateFormulaStartCell"

End Sub

Private Sub AddLambdaToNameManagerIfRequested(ByVal FullLambda As String)
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.AddLambdaToNameManagerIfRequested"
    ' Adds the Lambda formula to the Name Manager if requested and updates the LambdaFormula property.
    
    FullLambda = FormatFormula(FullLambda)
    ' If no need to add in the Named Range, then Exit Sub.
    If Not This.IsAddToNameManager Then
        ' Update the LambdaFormula property and exit the subroutine.
        This.CurrentLambdaInfo.LambdaFormula = FullLambda
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaParser.AddLambdaToNameManagerIfRequested"
        Exit Sub
    End If

    ' If need to add in the Named Range, then add and update LambdaFormula.
    Dim LambdaFormulaName As String
    LambdaFormulaName = This.PrecedencyExtractor.PossibleLambdaName
    If LambdaFormulaName = vbNullString Then
        ' If no valid name found, display a message and set the LambdaFormula property to the FullLambda.
        MsgBox "Unable to add Lambda formula to Name Manager. No valid name found." _
               , vbExclamation + vbOKOnly, APP_NAME
        This.CurrentLambdaInfo.LambdaFormula = FullLambda
    Else
        ' Add the Lambda formula to the Name Manager using the found LambdaFormulaName.
        Dim FormulaText As String
        FormulaText = ConvertDependencisToFullyQualifiedRef(ReplaceNewlineWithChar10(FullLambda) _
                                                            , This.PrecedencyExtractor.FormulaStartCellInSheet)
        This.PrecedencyExtractor.FormulaInWorkbook.Names.Add LambdaFormulaName, FormulaText

        ' Update the LambdaFormula property with the Lambda formula reference.
        This.CurrentLambdaInfo.LambdaFormula = EQUAL_SIGN & LambdaFormulaName
    End If
    Logger.Log TRACE_LOG, "Exit FormulaParser.AddLambdaToNameManagerIfRequested"

End Sub

Public Function GetLambdaPreview(ByVal DependencyObjects As Collection) As String
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.GetLambdaPreview"
    ' Returns the preview of the Lambda formula by creating the Let statements and the full Lambda expression.
    
    Dim Result As String
    With This.FormulaGen
        .CreateLetAndLambdaStatements DependencyObjects, This.PrecedencyExtractor.FormulaStartCell, LAMBDA_STATEMENT_GENERATION, This.PrecedencyExtractor
        ' Return the full Lambda expression along with the Lambda invocation argument.
        Result = FormatFormula(.LambdaDef & .LambdaInvocation)
        Result = TranslateEnUSFormulaToApplicationLanguage(Result)
        This.IsLetNeededInLambda = .IsLetNeededInLambda
    End With
    
    GetLambdaPreview = Result
    
    Logger.Log TRACE_LOG, "Exit FormulaParser.GetLambdaPreview"
    
End Function

Public Function GetLetPreview(ByVal DependencyObjects As Collection) As String
    
    ' Returns the preview of the LET formula.
    
    Dim Result As String
    With This.FormulaGen
        .CreateLetAndLambdaStatements DependencyObjects, This.PrecedencyExtractor.FormulaStartCell, LET_STATEMENT_GENERATION, This.PrecedencyExtractor
        ' Return the full Lambda expression along with the Lambda invocation argument.
        Result = FormatFormula(.LetFormula)
        Result = TranslateEnUSFormulaToApplicationLanguage(Result)
        This.IsLetNeededInLambda = .IsLetNeededInLambda
    End With
    
    GetLetPreview = Result
    

End Function

Private Sub UpdateInputCellsFromUserChoice(ByRef DependencyObjects As Collection _
                                           , ByVal TypeOfOperation As OperationType)
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.UpdateInputCellsFromUserChoice"
    ' Updates the input cells from the user's choice if the operation type is LAMBDA_STATEMENT_GENERATION.

    ' If the operation type is not LAMBDA_STATEMENT_GENERATION, exit the subroutine.
    If Not (TypeOfOperation = LAMBDA_STATEMENT_GENERATION Or TypeOfOperation = LET_STATEMENT_GENERATION) Then Exit Sub

    ' Update the DependencyObjects based on the user's choice for renaming, reordering, or demoting input cells.
    Set DependencyObjects = UpdateFromUserChooseForRenameOrReorderOrDemote(DependencyObjects, TypeOfOperation)

    ' Get the number of non-input dependencies from the updated DependencyObjects.
    This.PrecedencyExtractor.NumberOfNonInputDependency = GetNumberOfNonInputDependency(DependencyObjects)
    Logger.Log TRACE_LOG, "Exit FormulaParser.UpdateInputCellsFromUserChoice"

End Sub

Private Function UpdateFromUserChooseForRenameOrReorderOrDemote(ByVal DependencyObjects As Collection _
, ByVal TypeOfOperation As OperationType) As Collection
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.UpdateFromUserChooseForRenameOrReorderOrDemote"
    ' Updates the DependencyObjects based on the user's choice for renaming, reordering, or demoting input cells.

    ' Create a ParamSelector object.
    Dim UF As Object
    If TypeOfOperation = LAMBDA_STATEMENT_GENERATION Then
        Set UF = New ParamSelector
    Else
        Set UF = New LETManagerUI
    End If

    ' Set the DependencyObjects and Parser properties of the ParamSelector object.
    Set UF.DependencyObjects = DependencyObjects
    Set UF.Parser = Me

    ' Update the ListBox in the ParamSelector form with data from the DependencyObjects.
    UF.UpdateListBoxFromCollection

    ' Show the ParamSelector form to allow the user to make choices.
    UF.Show

    ' Return the updated DependencyObjects.
    Set UpdateFromUserChooseForRenameOrReorderOrDemote = UF.DependencyObjects
    Logger.Log TRACE_LOG, "Exit FormulaParser.UpdateFromUserChooseForRenameOrReorderOrDemote"
    
End Function

Private Function GetValueTwoIndexBack(ByVal CurrentIndex As Long, ByVal GivenArray As Variant) As String
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.GetValueTwoIndexBack"
    ' Returns the value from two indices back in the given array.

    If CurrentIndex > LBound(GivenArray) + 2 Then
        GetValueTwoIndexBack = GivenArray(CurrentIndex - 2, LBound(GivenArray, 2))
    End If
    Logger.Log TRACE_LOG, "Exit FormulaParser.GetValueTwoIndexBack"

End Function

Private Function ConvertFullLetToAFEScript(ByVal FullLetFormula As String) As Variant
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.ConvertFullLetToAFEScript"
    ' Converts the FullLetFormula to an array of AFE (Assignment, Formula, Equals) script lines.

    ' Remove the "Let" keyword and the ending parenthesis from FullLetFormula.
    FullLetFormula = modUtility.RemoveStartingLetAndEndParenthesis(FullLetFormula)

    ' Split the FullLetFormula into individual LetVarNameWithValue statements based on newline character.
    Dim SplittedLet As Variant
    SplittedLet = Split(FullLetFormula, Chr$(10))

    Dim LetVarNameWithValue As Variant
    Dim AllData As Collection
    Set AllData = New Collection

    Dim Counter As Long
    For Each LetVarNameWithValue In SplittedLet
        ' Check if the LetVarNameWithValue statement is not empty or just whitespace.
        If Application.WorksheetFunction.Trim(CStr(LetVarNameWithValue)) <> vbNullString Then
            ' Remove leading spaces, add "=" between variable name and value, and remove trailing characters like newline and comma.
            LetVarNameWithValue = VBA.LTrim$(LetVarNameWithValue)
            LetVarNameWithValue = Replace(LetVarNameWithValue, LIST_SEPARATOR, ONE_SPACE & EQUAL_SIGN, 1, 1, vbTextCompare)
            LetVarNameWithValue = Text.RemoveFromEndIfPresent(CStr(LetVarNameWithValue), Chr$(13), IGNORE_CASE)
            LetVarNameWithValue = Text.RemoveFromEndIfPresent(CStr(LetVarNameWithValue), LIST_SEPARATOR, IGNORE_CASE) & ";"

            ' Add the LetVarNameWithValue statement to the collection.
            AllData.Add LetVarNameWithValue
        End If

        Counter = Counter + 1

        ' Exit the loop if all LetVarNameWithValue statements have been processed.
        If Counter = UBound(SplittedLet) - 1 Then Exit For
    Next LetVarNameWithValue

    ' Convert the collection to a 2D array.
    ConvertFullLetToAFEScript = modUtility.CollectionToArray(AllData)

    ' Log the AFE script lines.
    Logger.Log DEBUG_LOG, THREE_NEW_LINE
    Logger.Log DEBUG_LOG, modUtility.ConcatenateOneColumnOf2DArray(ConvertFullLetToAFEScript, 1, vbNewLine)

    Logger.Log TRACE_LOG, "Exit FormulaParser.ConvertFullLetToAFEScript"
    
End Function

Private Sub Class_Initialize()
    ClearCatches
End Sub


