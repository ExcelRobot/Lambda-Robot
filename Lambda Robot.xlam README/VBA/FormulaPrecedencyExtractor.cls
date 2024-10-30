VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaPrecedencyExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed, ImplicitActiveWorkbookReference, SuperfluousAnnotationArgument, UnrecognizedAnnotation, ProcedureNotUsed
'@Folder "Lambda.PrecedencyExtractor"
Option Explicit

Public Enum SortByColumn
    Level = 3
    IsLabelAsInputCell = 7
    SheetName = 9
    ColumnNumber = 10
    RowNumber = 11
    OPTIONAL_ARGUMENT = 12
End Enum

Private Type TFormulaPrecedencyExtractor
    AllDependency As Collection
    FormulaStartCell As Range
    ActualFormulaStartCell As Range
    FormulaStartCellSheetName As String
    VarNameOfFormulaStartCell As String
    DependencySearchInRegion As Range
    PossibleLambdaName As String
    StructuredFormulaManager As StructuredRefManager
    NumberOfNonInputDependency As Long
    AllNamedRange As Dictionary
    DependencyData As Variant
    ModifiedDependencies As Collection
    FormulaStartCellInSheet As Worksheet
    FormulaInWorkbook As Workbook
    IsAskedUserForMoreThanThresholdSteps As Boolean
    IsStopOnMoreThanThresholdSteps As Boolean
    IsPopupShowedForLargerFormulaLength As Boolean
    TotalFormulaLength As Long
    IsProcessEnded As Boolean
    SheetNameVsDependencyRangeMap As Collection
End Type

Private This  As TFormulaPrecedencyExtractor

Private Const THRESHOLD_STEPS_COUNT As Long = 25
Private Const THRESHOLD_FORMULA_LENGTH As Long = MAX_LENGTH_OF_FORMULA + 1000   ' 1000 for tolerance.

Public Property Get IsProcessEnded() As Boolean
    IsProcessEnded = (This.IsStopOnMoreThanThresholdSteps Or This.IsPopupShowedForLargerFormulaLength)
End Property

Public Property Get FormulaInWorkbook() As Workbook
    Set FormulaInWorkbook = This.FormulaInWorkbook
End Property

Public Property Get FormulaStartCellInSheet() As Worksheet
    Set FormulaStartCellInSheet = This.FormulaStartCellInSheet
End Property

Public Property Get AllDependency() As Collection
    Set AllDependency = This.AllDependency
End Property

Public Property Set AllDependency(ByVal RHS As Collection)
    Set This.AllDependency = RHS
End Property

Public Property Get FormulaStartCell() As Range
    Set FormulaStartCell = This.FormulaStartCell
End Property

Public Property Set FormulaStartCell(ByVal RHS As Range)
    Set This.FormulaStartCell = RHS
End Property

Public Property Get ActualFormulaStartCell() As Range
    Set ActualFormulaStartCell = This.ActualFormulaStartCell
End Property

Public Property Set ActualFormulaStartCell(ByVal RHS As Range)
    Set This.ActualFormulaStartCell = RHS
End Property

Public Property Get VarNameOfFormulaStartCell() As String
    VarNameOfFormulaStartCell = This.VarNameOfFormulaStartCell
End Property

Public Property Let VarNameOfFormulaStartCell(ByVal RHS As String)
    This.VarNameOfFormulaStartCell = RHS
End Property

Public Property Get DependencySearchInRegion() As Range
    Set DependencySearchInRegion = This.DependencySearchInRegion
End Property

Public Property Set DependencySearchInRegion(ByVal RHS As Range)
    Set This.DependencySearchInRegion = RHS
End Property

Public Property Get PossibleLambdaName() As String
    PossibleLambdaName = This.PossibleLambdaName
End Property

Public Property Let PossibleLambdaName(ByVal RHS As String)
    This.PossibleLambdaName = RHS
End Property

Public Property Get NumberOfNonInputDependency() As Long
    NumberOfNonInputDependency = This.NumberOfNonInputDependency
End Property

Public Property Let NumberOfNonInputDependency(ByVal RHS As Long)
    This.NumberOfNonInputDependency = RHS
End Property

Public Property Get AllNamedRange() As Dictionary
    Set AllNamedRange = This.AllNamedRange
End Property

Public Property Set AllNamedRange(ByVal RHS As Dictionary)
    Set This.AllNamedRange = RHS
End Property

Public Property Get DependencyData() As Variant
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.DependencyData"
    If IsObject(This.DependencyData) Then
        Set DependencyData = This.DependencyData
    Else
        DependencyData = This.DependencyData
    End If
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.DependencyData"
    
End Property

Public Property Let DependencyData(ByVal RHS As Variant)
    This.DependencyData = RHS
End Property

Public Property Set DependencyData(ByVal RHS As Variant)
    Set This.DependencyData = RHS
End Property

Public Function DependencyDataForReset(TypeOfOperation As OperationType) As Collection
    Set DependencyDataForReset = GetDependencyObjectsFromData(This.DependencyData, TypeOfOperation)
End Function

'@EntryPoint
Public Sub ExtractAllPrecedencyInfo(ByVal FormulaStartCell As Range _
                                    , ByVal DependencySearchInRegion As Range _
                                     , ByVal UserModifiedDependencies As Collection)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.ExtractAllPrecedencyInfo"
    ' Extracts all the precedency information for the formula starting from the FormulaStartCell.
    ' FormulaStartCell: The cell containing the formula.
    ' DependencySearchInRegion: The range within which dependencies are searched.
    ' UserModifiedDependencies: Collection of user-modified dependencies.

    ' Check if the FormulaStartCell is provided.
    If IsNothing(FormulaStartCell) Then
        MsgBox "Unable to generate dependency info. Formula cell is not provided.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaPrecedencyExtractor.ExtractAllPrecedencyInfo"
        Exit Sub
    End If

    ' Clear previously stored data.
    ClearCatches

    ' Set DependencySearchInRegion and UserModifiedDependencies.
    Set This.DependencySearchInRegion = DependencySearchInRegion
    Set This.ModifiedDependencies = UserModifiedDependencies

    ' Extract basic information from the FormulaStartCell.
    InitializeAndExtractBaseInfo FormulaStartCell

    ' Generate dependency information starting from the FormulaStartCell.
    GenerateDependencyInfo FormulaStartCell
    ResolveRangeLabelConflict
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.ExtractAllPrecedencyInfo"

End Sub

Private Sub ResolveRangeLabelConflict()
    
    If This.IsProcessEnded Then Exit Sub
    
    Dim CurrentDependency As DependencyInfo
    For Each CurrentDependency In This.AllDependency
        With CurrentDependency
            If IsNotNothing(.RangeLabelCell) Then
                Dim SheetName As String
                SheetName = .SheetName
                
                Dim DependencyRegion As Range
                Set DependencyRegion = This.SheetNameVsDependencyRangeMap.Item(SheetName)
                
                ' Now we have a conflict. We need to retrace for the range label.
                If IsNotNothing(Intersect(DependencyRegion, .RangeLabelCell)) Then
                    ResolveRangeLabelConflictForCurrDependency CurrentDependency, DependencyRegion
                End If
                
            End If
        End With
    Next CurrentDependency
    
End Sub

Private Sub ResolveRangeLabelConflictForCurrDependency(ByRef CurrentDependency As DependencyInfo, ByVal DependencyRegion As Range)
    
    With CurrentDependency
        
        Logger.Log DEBUG_LOG, "Conflict found."
        Dim ResolvedRange As Range
        Set ResolvedRange = RangeResolver.GetRange(.RangeReference, Application.Workbooks(.BookName), Application.Worksheets(.SheetName))
        Set .RangeLabelCell = FindLetVarNameCell(ResolvedRange, DependencyRegion)
        If IsNotNothing(.RangeLabelCell) Then
            .RangeLabel = .RangeLabelCell.Value
            .ValidVarName = MakeValidLetVarName(.RangeLabel, GetNamingConv(.IsLabelAsInputCell))
            .ValidVarName = CorrectValidVarNameForNamedRange(.ValidVarName, .IsReferByNamedRange, .IsLabelAsInputCell)
        Else
            .RangeLabel = vbNullString
            .ValidVarName = vbNullString
        End If
    
    End With
                    
End Sub

Private Sub ClearCatches()
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.ClearCatches"
    ' Clears all the previously stored data.

    ' Clear all stored objects and variables.
    Set This.AllDependency = Nothing
    Set This.SheetNameVsDependencyRangeMap = Nothing
    This.TotalFormulaLength = 0
    Set This.FormulaStartCell = Nothing
    Set This.FormulaStartCellInSheet = Nothing
    This.FormulaStartCellSheetName = vbNullString
    Set This.FormulaInWorkbook = Nothing
    This.VarNameOfFormulaStartCell = vbNullString
    Set This.AllNamedRange = Nothing

    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.ClearCatches"
    
End Sub

Private Sub InitializeAndExtractBaseInfo(ByVal FormulaStartCell As Range)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.InitializeAndExtractBaseInfo"
    ' Creates basic information required for precedency extraction from the FormulaStartCell.
    ' FormulaStartCell: The cell containing the formula.

    ' Initialize collections and variables.
    Set This.AllDependency = New Collection
    Set This.SheetNameVsDependencyRangeMap = New Collection
    ' Store information related to the FormulaStartCell.
    Set This.FormulaStartCell = FormulaStartCell
    Set This.FormulaStartCellInSheet = This.FormulaStartCell.Worksheet
    This.FormulaStartCellSheetName = This.FormulaStartCellInSheet.name
    Set This.FormulaInWorkbook = This.FormulaStartCellInSheet.Parent

    ' Delete LET step named ranges having errors from the workbook.
    DeleteLETStepNamedRangesHavingError This.FormulaInWorkbook

    ' Get all named ranges and store them in a dictionary.
    Set This.AllNamedRange = GetNamedRangeToDictionary(This.FormulaInWorkbook)
    Set This.StructuredFormulaManager = New StructuredRefManager
    With This.StructuredFormulaManager
        Set .FormulaStartCellInSheet = This.FormulaStartCellInSheet
        Set .AllNamedRange = This.AllNamedRange
    End With

    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.InitializeAndExtractBaseInfo"
    
End Sub

Private Sub GenerateDependencyInfo(ByVal FormulaStartCell As Range)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.GenerateDependencyInfo"
    ' Generates the dependency information for the formula starting from the FormulaStartCell.
    ' FormulaStartCell: The cell containing the formula.

    On Error GoTo ErrorHandler

    ' Create dependency info for the FormulaStartCell.
    CreateDependencyInfoForStartCell FormulaStartCell

    ' Find dependencies recursively starting from the FormulaStartCell.
    FindDependency FormulaStartCell, PrecedencyLevel:=1
    If This.IsStopOnMoreThanThresholdSteps Then Exit Sub
    If This.IsPopupShowedForLargerFormulaLength Then Exit Sub
    ' Mark previous parameter cells as parameter cells in the new dependencies collection.
    MarkPreviousParameterCellsAsParameter This.AllDependency, This.ModifiedDependencies

    ' Maximize the level of dependencies.
    MaximizeLevel This.AllDependency

    ' Update the dependencies to mark optional parameters.
    UpdateForOptional This.AllDependency
    
    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaPrecedencyExtractor.GenerateDependencyInfo"
    Exit Sub
    Resume

ErrorHandler:
    
    Set This.AllDependency = Nothing
    ' Handle errors related to formula length exceeding Excel limit.
    If Err.Description = "Formula Lenth Cross Excel Limit" Then
        MsgBox "Unable to generate Lambda formula. Resulting formula too long.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaPrecedencyExtractor.GenerateDependencyInfo"
        Exit Sub
    End If

    ' Handle other errors and log them.
    If Err.Number <> 0 Then
        Logger.Log ERROR_LOG, Err.Number & "-" & Err.Description
        Err.Clear
    End If

    ' Close the log file.
    Logger.CloseLogFile
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.GenerateDependencyInfo"
    
End Sub

Private Sub CreateDependencyInfoForStartCell(ByRef FormulaStartCell As Range)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.CreateDependencyInfoForStartCell"
    ' Creates dependency information for the FormulaStartCell.
    ' FormulaStartCell: The cell containing the formula.

    ' Find the possible Lambda name from the FormulaStartCell.
    This.PossibleLambdaName = modUtility.FindLetVarName(FormulaStartCell.Cells(1, 1))
    This.PossibleLambdaName = modUtility.MakeValidLetVarName(This.PossibleLambdaName, GetNamingConv(False))

    ' Loop back to the actual cell (in case of a spill range).
    Set FormulaStartCell = LoopBackToActualCell(FormulaStartCell)
    Set This.ActualFormulaStartCell = FormulaStartCell

    Dim RangeReference As String
    ' Get the range reference of the FormulaStartCell.
    RangeReference = GetDependencyRangeReference(FormulaStartCell)

    ' Get the modified dependency info for the range reference.
    Dim ModifiedDependencyInfo As DependencyInfo
    Set ModifiedDependencyInfo = GetModifiedDependencyInfo(RangeReference)
    This.VarNameOfFormulaStartCell = FindVarNameOfFormulaStartCell(FormulaStartCell, ModifiedDependencyInfo)

    Dim FormulaText As String
    If IsNothing(ModifiedDependencyInfo) Then
        FormulaText = This.StructuredFormulaManager.GetStructuredFormula(FormulaStartCell)
    Else
        FormulaText = ModifiedDependencyInfo.FormulaText
    End If

    ' Create new dependency info for the FormulaStartCell.
    CreateNewDependency Level:=0, RangeReference:=RangeReference, _
                        FormulaText:=FormulaText, _
                        CurrentCell:=FormulaStartCell, _
                        HasAnyDependency:=True, _
                        NameInFormula:=FormulaStartCell.Address(False, False)

    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.CreateDependencyInfoForStartCell"
    
End Sub

Public Sub UpdateValidVarForUsedNamedRanges(ByRef AllDependency As Collection)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.UpdateValidVarForUsedNamedRanges"
    Dim UsedNamedRanges As Collection
    Set UsedNamedRanges = New Collection
    
    Dim CurrentDependency As DependencyInfo
    
    ' Find all the used named range ref. Now we need to check for other ref.
    Dim Index As Long
    For Each CurrentDependency In AllDependency
        Index = Index + 1
        With CurrentDependency
            If .IsReferByNamedRange Then
                UsedNamedRanges.Add Index, .NameInFormula
            End If
        End With
    Next CurrentDependency
    
    Index = 0
    For Each CurrentDependency In AllDependency
        Index = Index + 1
        With CurrentDependency
            If Not .IsReferByNamedRange And IsExistInCollection(UsedNamedRanges, .ValidVarName) Then
                .ValidVarName = FindUniqueNameByIncrementingNumber(UsedNamedRanges, .ValidVarName)
            End If
        End With
    Next CurrentDependency
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.UpdateValidVarForUsedNamedRanges"
    
End Sub

Private Sub MarkPreviousParameterCellsAsParameter(ByRef NewDependencies As Collection _
                                                  , ByVal OldDependencies As Collection)
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.MarkPreviousParameterCellsAsParameter"
    ' Marks the previous parameter cells as parameter cells in the new dependencies collection.
    ' NewDependencies: Collection of new dependencies.
    ' OldDependencies: Collection of old dependencies.

    If IsNothing(OldDependencies) Then Exit Sub
    If IsNothing(NewDependencies) Then Exit Sub

    Dim CurrentOldDependency As DependencyInfo
    For Each CurrentOldDependency In OldDependencies
        With CurrentOldDependency
            If .IsLabelAsInputCell Or .IsUserMarkAsParameterCell Then
                If IsExistInCollection(NewDependencies, .RangeReference) Then
                    Dim CurrentDependencyInfo As DependencyInfo
                    Set CurrentDependencyInfo = NewDependencies.Item(.RangeReference)
                    CurrentDependencyInfo.IsLabelAsInputCell = True
                End If
            End If
        End With
    Next CurrentOldDependency
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.MarkPreviousParameterCellsAsParameter"
    
End Sub

Private Function FindVarNameOfFormulaStartCell(ByVal FormulaStartCell As Range _
                                               , ByVal ModifiedDependencyInfo As DependencyInfo) As String
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.FindVarNameOfFormulaStartCell"
    ' Finds the variable name for the FormulaStartCell based on CurrentDependencyInfo if available,
    ' otherwise, generates a new variable name using the formula in the FormulaStartCell.
    ' FormulaStartCell: The cell containing the formula.
    ' CurrentDependencyInfo: DependencyInfo of the FormulaStartCell, if available.

    Dim VarName As String
    If IsNotNothing(ModifiedDependencyInfo) Then
        VarName = ModifiedDependencyInfo.ValidVarName
    Else
        ' Generate a new variable name from the formula in the FormulaStartCell.
        VarName = modUtility.FindLetVarName(FormulaStartCell.Cells(1, 1))
        VarName = modUtility.MakeValidLetVarName(VarName, GetNamingConv(False))
        If VarName = vbNullString Then VarName = GetLetVarPrefix() & MakeValidLetVarName(LCase$(LAST_STEP_NAME), GetNamingConv(False))
    End If

    FindVarNameOfFormulaStartCell = VarName
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.FindVarNameOfFormulaStartCell"
    
End Function


'@Recursive
Private Sub FindDependency(ByVal FormulaStartCell As Range, ByVal PrecedencyLevel As Long)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.FindDependency"
    ' Recursively finds dependencies for the FormulaStartCell and updates the dependency information.
    ' FormulaStartCell: The cell containing the formula.
    ' PrecedencyLevel: The level of the formula's precedents in the dependency hierarchy.
    Logger.Log INFO_LOG, vbNewLine
    Logger.Log TRACE_LOG, "Check dependency for " & FormulaStartCell.Address
    
    If IsStopIfMoreThanThresholdSteps() Then Exit Sub
    If IsFormulaLengthCrossedExcelCharacterLimit() Then Exit Sub
    
    ' Uncomment the line below for debugging purposes.
    ' Debug.Assert FormulaStartCell.Address(False, False) <> "I4"
     
    On Error GoTo ExitSub

    Dim DirectPrecedents As Collection
    Set DirectPrecedents = FindDirectPrecedency(FormulaStartCell)

    ' Base case to exit if there are no direct precedents.
    If DirectPrecedents.Count = 0 Then
        Logger.Log TRACE_LOG, "No Direct Precedent for " & FormulaStartCell.Address
        ' As no dependency would be found for this one. So level will be PrecedencyLevel-1 as this PrecedencyLevel is for it's precedents.
        UpdateDependencyIfNoMorePrecedents FormulaStartCell, PrecedencyLevel - 1
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaPrecedencyExtractor.FindDependency"
        Exit Sub
    End If

    Dim CurrentDirectPrecedent As DirectPrecedent
    For Each CurrentDirectPrecedent In DirectPrecedents
        
        If IsFormulaLengthCrossedExcelCharacterLimit() Then Exit Sub
        With CurrentDirectPrecedent
            Dim RangeReference As String
            RangeReference = .RangeRef

            ' Uncomment the line below for debugging purposes.
            ' Debug.Assert RangeReference <> "FILTERCOLUMNS!AF3#"

            Dim FormulaText As String
            FormulaText = FindFormulaText(This.FormulaInWorkbook, This.FormulaStartCellInSheet, RangeReference)

            If IsExistInCollection(This.AllDependency, RangeReference) Then
                UpdateLevel RangeReference, PrecedencyLevel
                This.AllDependency.Item(RangeReference).UpdateNameInFormula .NameInFormula
            Else
                Dim CurrentRange As Range
                Set CurrentRange = RangeResolver.FindRangeFromText(This.FormulaInWorkbook, This.FormulaStartCellInSheet, RangeReference)
                
                Dim IsAnyDep As Boolean
                IsAnyDep = True
                '                Debug.Assert IsNotNothing(CurrentRange)
                If IsNothing(CurrentRange) And IsLocalOrGlobalNamedRangeExist(.NameInFormula, This.FormulaStartCellSheetName) Then
                    FormulaText = GetNamedRangeByNameInFormula(.NameInFormula, This.FormulaStartCellSheetName).RefersTo
                    IsAnyDep = False
                Else
                    FormulaText = This.StructuredFormulaManager.FindStructuralFormulaText(CurrentRange, FormulaText)
                End If
            
                CreateNewDependency Level:=PrecedencyLevel, _
                                    RangeReference:=RangeReference, _
                                    FormulaText:=FormulaText, _
                                    CurrentCell:=CurrentRange, _
                                    HasAnyDependency:=IsAnyDep, _
                                    NameInFormula:=.NameInFormula
                
                ' If const named range then no need to trace back. In case of LETStep_ we have a cell for that so range will not be nothing.
                If Not IsNothing(CurrentRange) Then
                    FindDependency CurrentRange, PrecedencyLevel + 1
                End If
            
            End If
        End With
        
    Next CurrentDirectPrecedent

    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaPrecedencyExtractor.FindDependency"
    Exit Sub

ExitSub:
    If Err.Number = 1004 Then
        Logger.Log ERROR_LOG, FormulaStartCell.Address & " : " & Err.Description
        Err.Clear
        Resume Next
    ElseIf Err.Number <> 0 Then
        Logger.Log TRACE_LOG, Err.Number & ONE_SPACE & Err.Description
        Err.Raise Err.Number, Err.Source, Err.Description
        Resume
    End If
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.FindDependency"

End Sub

Private Function IsFormulaLengthCrossedExcelCharacterLimit() As Boolean
    
    If This.TotalFormulaLength >= THRESHOLD_FORMULA_LENGTH Then
        
        If Not This.IsPopupShowedForLargerFormulaLength Then
            MsgBox "Resulting formula will cross the excel formula character limit.", vbOKOnly + vbCritical + vbInformation, "Long Formula"
            This.IsPopupShowedForLargerFormulaLength = True
        End If

    End If
    
    IsFormulaLengthCrossedExcelCharacterLimit = This.IsPopupShowedForLargerFormulaLength
    
End Function

Private Function IsStopIfMoreThanThresholdSteps() As Boolean
    
    If This.AllDependency.Count > THRESHOLD_STEPS_COUNT Then
        
        If Not This.IsAskedUserForMoreThanThresholdSteps Then
            Dim Answer As VbMsgBoxResult
            Answer = MsgBox("The resulting formula will have more than " & THRESHOLD_STEPS_COUNT & " LET variables.  Continue anyway?", vbYesNo, "More Steps")
            This.IsStopOnMoreThanThresholdSteps = (Answer = vbNo)
            This.IsAskedUserForMoreThanThresholdSteps = True
        End If
        
    End If
    
    IsStopIfMoreThanThresholdSteps = This.IsStopOnMoreThanThresholdSteps
    
End Function

Private Sub UpdateDependencyIfNoMorePrecedents(ByVal FormulaStartCell As Range, ByRef PrecedencyLevel As Long)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.UpdateDependencyIfNoMorePrecedents"
    ' Updates the dependency information if there are no more precedents for the FormulaStartCell.
    ' FormulaStartCell: The cell containing the formula.
    ' PrecedencyLevel: The level of the formula's precedents in the dependency hierarchy.

    Dim NameInFormula As String

    ' If the formula is on a different sheet, include the sheet reference in NameInFormula.
    If FormulaStartCell.Worksheet.name <> This.FormulaStartCellSheetName Then
        NameInFormula = GetSheetRefForRangeReference(FormulaStartCell.Worksheet.name, False)
    End If

    Dim RangeReference As String
    RangeReference = GetDependencyRangeReference(FormulaStartCell)

    Dim ModifiedDependencyInfo As DependencyInfo
    Set ModifiedDependencyInfo = GetModifiedDependencyInfo(RangeReference)

    Dim FormulaText As String
    If IsNotNothing(ModifiedDependencyInfo) Then
        FormulaText = ModifiedDependencyInfo.FormulaText
        If Not ModifiedDependencyInfo.HasFormula Then
            NameInFormula = RemoveHashSignFromEnd(ModifiedDependencyInfo.NameInFormula)
        End If
    ElseIf HasDynamicFormula(FormulaStartCell) Then
        NameInFormula = NameInFormula & GetSpillParentAddress(FormulaStartCell) & DYNAMIC_CELL_REFERENCE_SIGN
        FormulaText = GetSpillParentCell(FormulaStartCell).Formula2
    Else
        NameInFormula = NameInFormula & GetNameInFormula(FormulaStartCell)
    End If

    If IsExistInCollection(This.AllDependency, RangeReference) Then
        UpdateLevel RangeReference, PrecedencyLevel
        This.AllDependency.Item(RangeReference).HasAnyDependency = False
    Else
        If IsNothing(ModifiedDependencyInfo) Then
            FormulaText = This.StructuredFormulaManager.FindStructuralFormulaText(FormulaStartCell, FormulaText)
        End If
        CreateNewDependency PrecedencyLevel, RangeReference, FormulaText, FormulaStartCell, False, NameInFormula
    End If

    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.UpdateDependencyIfNoMorePrecedents"
    
End Sub

Private Function GetDependencyRangeReference(ByVal ForCell As Range) As String
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.GetRangeReferenceIfAnyCellInsideSpillThenUseHash"
    ' Gets the range reference for a cell, and if it's inside a spill, appends a hash sign at the end.
    ' ForCell: The cell for which the range reference is needed.
    
    
    Dim SheetNamePrefix As String
    SheetNamePrefix = GetSheetRefForRangeReference(ForCell.Worksheet.name, False)
    
    Dim RangeReference As String
    RangeReference = SheetNamePrefix & ForCell.Address(False, False)
    
    Logger.Log DEBUG_LOG, "Range Ref for cell: " & GetRangeRefWithSheetName(ForCell)
    If ForCell.HasSpill Then
        If IsFullSpillingRangeIncluded(ForCell) Then
            RangeReference = SheetNamePrefix & GetSpillParentAddress(ForCell, False) & DYNAMIC_CELL_REFERENCE_SIGN
        End If
    End If

    GetDependencyRangeReference = RangeReference

    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.GetRangeReferenceIfAnyCellInsideSpillThenUseHash"
    
End Function

Private Function FindDirectPrecedency(ByVal ForCell As Range) As Collection
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.FindDirectPrecedency"
    ' Finds the direct precedents for the given cell.
    ' ForCell: The cell for which the direct precedents are to be found.
    ' FormulaApplicationHelperCell: The helper cell used for formula calculation.

    ' If there are no more precedents to find, return an empty collection.
    If IsNothing(ForCell) Or IsNoMorePrecedency(ForCell) Then
        Set FindDirectPrecedency = New Collection
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaPrecedencyExtractor.FindDirectPrecedency"
        Exit Function
    End If
    
    On Error GoTo ExitMethod
    Dim StartTime As Double
    StartTime = Timer()
    This.StructuredFormulaManager.ConvertRangeRefToStructuralRef ForCell
    Logger.Log DEBUG_LOG, "Time to convert formula to structured ref: " _
                         & Timer() - StartTime & " for cell: " & ForCell.Address
    
    Dim ForCellReference As String
    ForCellReference = GetRangeRefWithSheetName(ForCell, False)

    Dim LetStepLambdaRefs As Scripting.Dictionary
    Set LetStepLambdaRefs = This.StructuredFormulaManager.GetLetStepLambdaRefs(ForCell)

    Dim DirectPrecedentCells As Collection
    Set DirectPrecedentCells = New Collection
    
    Dim DirectPrecedents As Variant
    DirectPrecedents = GetDirectPrecedents(This.StructuredFormulaManager.GetStructuredFormula(ForCell) _
                                           , ForCell.Worksheet)
    
    Dim CurrentPrecendentReference As Variant
    For Each CurrentPrecendentReference In DirectPrecedents
        
        If CurrentPrecendentReference = vbNullString Then Exit For
        
        ' I don’t think we should trace precedence outside of active workbook even if it is open.
        ' Chat Date: 8 Sept 2024, Channel: Erik and Rob Teams group.
        
        If Not IsReferenceFromDifferentBook(CurrentPrecendentReference, This.FormulaInWorkbook) Then
            ' Add the direct precedent to the collection.
            AddDependencyToCollection CStr(CurrentPrecendentReference), DirectPrecedentCells _
                                                                       , ForCellReference _
                                                                        , ForCell, vbNullString
        End If
        
    Next CurrentPrecendentReference

    ' Add any Let-Step-Lambda references to the collection.
    Dim CurrentLetStepLambdaRef As Variant
    For Each CurrentLetStepLambdaRef In LetStepLambdaRefs.Keys
        AddDependencyToCollection CurrentLetStepLambdaRef, DirectPrecedentCells _
                                                          , ForCellReference _
                                                           , ForCell _
                                                            , LetStepLambdaRefs.Item(CurrentLetStepLambdaRef)
    Next CurrentLetStepLambdaRef

    Logger.Log DEBUG_LOG, "Total Precedent cells: " & DirectPrecedentCells.Count & " for cell: " & ForCell.Address
    Set FindDirectPrecedency = DirectPrecedentCells

    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaPrecedencyExtractor.FindDirectPrecedency"
    Exit Function

ExitMethod:
    Logger.Log ERROR_LOG, Err.Number & "-" & Err.Description
    Err.Clear
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.FindDirectPrecedency"
    
End Function

Private Sub AddDependencyToCollection(ByVal PrecendentRef As String _
                                      , ByVal DirectPrecedentCells As Collection _
                                       , ByVal ForCellReference As String _
                                       , ByVal DependencyForCell As Range _
                                        , Optional ByVal LetStepFXName As String = vbNullString)
                                        
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.AddDependencyToCollection"
    ' Adds the current precedent reference to the direct precedent cells collection, if it doesn't exist.
    ' CurrentPrecendentReference: The current precedent reference to be added.
    ' DirectPrecedentCells: The collection of direct precedent cells.
    ' ForCellReference: The reference of the current cell for which precedents are being found.

    Dim CurrentPrecedent As DirectPrecedent
    Set CurrentPrecedent = New DirectPrecedent
    
    ' As value will be used as name in formula
    If LetStepFXName = vbNullString Then
        ' Assign the current precedent reference to the Value property of the mapper.
        CurrentPrecedent.NameInFormula = PrecendentRef
    Else
        CurrentPrecedent.NameInFormula = LetStepFXName
    End If
    
    ' Extract the range reference from named range or table.
    PrecendentRef = RangeResolver.GetRangeRefFromNamedRangeOrTable(This.FormulaInWorkbook _
                                                                   , PrecendentRef _
                                                                    , DependencyForCell.Worksheet)

    ' Make the range reference valid with the sheet name.
    PrecendentRef = MakeValidRangeReference(PrecendentRef, DependencyForCell.Worksheet.name)

    ' Assign the current precedent reference to the Key property of the mapper.
    CurrentPrecedent.RangeRef = PrecendentRef

    ' Add the current precedent reference to the direct precedent cells collection if it doesn't exist and is not the same as the current cell reference.
    If Not IsExistInCollection(DirectPrecedentCells, CStr(PrecendentRef)) _
       And ForCellReference <> CStr(PrecendentRef) Then
        DirectPrecedentCells.Add CurrentPrecedent, CStr(PrecendentRef)
    End If
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.AddDependencyToCollection"
    
End Sub

'Logic to find precedency anymore or not.
Private Function IsNoMorePrecedency(ByVal ForCell As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.IsNoMorePrecedency"
    ' Determines whether there are no more precedents for the given cell.
    ' ForCell: The cell for which to check the precedents.

    Dim Result As Boolean
    Dim RangeReference As String
    RangeReference = GetDependencyRangeReference(ForCell)
    Logger.Log DEBUG_LOG, "Check For RangeRef: " & RangeReference
    Dim ModifiedDependencyInfo As DependencyInfo
    Set ModifiedDependencyInfo = GetModifiedDependencyInfo(RangeReference)

    If IsNotNothing(ModifiedDependencyInfo) Then
        With ModifiedDependencyInfo
            Result = ( _
                     .IsLabelAsInputCell _
                     Or Not .HasFormula _
                     Or (.IsInsideNamedRangeOrTable And Not IsStartsWithLetStepPrefix(.NameInFormula)) _
                     Or .IsUserMarkAsParameterCell _
                     Or .IsUserMarkAsValue _
                     Or (.IsDemotedFromParameterCellToLetStep And Not .IsExpandByUser) _
                     )
        End With
    ElseIf ForCell.Worksheet.Parent.name <> This.FormulaInWorkbook.name Then
        Result = True
    ElseIf Not ForCell.HasFormula Then
        Result = True
    ElseIf IsInputCell(ForCell, This.DependencySearchInRegion) Then
        Result = True
    ElseIf IsInsideTable(ForCell) Then
        Result = True
    ElseIf IsReferedByNamedRange(ForCell) Then
        Result = Not (Text.Contains(ForCell.name.name, LETSTEPREF_UNDERSCORE_PREFIX))
    Else
        Result = False
    End If
    
    IsNoMorePrecedency = Result
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.IsNoMorePrecedency"
    
End Function

'@Impure
Private Function IsReferedByNamedRange(ByVal ForCell As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.IsReferedByNamedRange"
    ' Checks whether the given cell is referred to by a named range.
    ' ForCell: The cell to check for named range references.

    IsReferedByNamedRange = False
    Dim RangeReference As String
    RangeReference = GetDependencyRangeReference(ForCell)
    Dim CurrentDependencyInfo As DependencyInfo

    If IsExistInCollection(This.AllDependency, RangeReference) Then
        Set CurrentDependencyInfo = This.AllDependency.Item(RangeReference)
        IsReferedByNamedRange = CurrentDependencyInfo.IsReferByNamedRange
    End If
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.IsReferedByNamedRange"
    
End Function


Private Function MakeValidRangeReference(ByVal RangeReference As String, ByVal SheetName As String) As String
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.MakeValidRangeReference"
    ' Makes the range reference valid by including the sheet name, if necessary.
    ' RangeReference: The range reference to validate.
    ' SheetName: The name of the sheet containing the range.

    ' If the range reference is already a named range, keep it as is.
    If IsLocalOrGlobalNamedRangeExist(RangeReference, SheetName) Then
        MakeValidRangeReference = RangeReference
    ElseIf Not Text.Contains(RangeReference, SHEET_NAME_SEPARATOR) Then
        ' If the range reference doesn't contain a sheet name, include the sheet reference.
        RangeReference = GetSheetRefForRangeReference(SheetName, False) & RangeReference
    ElseIf VBA.Left$(RangeReference, 1) <> SINGLE_QUOTE Then
        ' If the range reference contains a sheet name but doesn't start with a single quote, keep it as is.
        RangeReference = RangeReference
    End If

    ' Remove dollar signs from the range reference to make it valid.
    RangeReference = RemoveDollarSign(RangeReference)
    MakeValidRangeReference = RangeReference
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.MakeValidRangeReference"

End Function

'@Recursive
Private Sub UpdateLevel(ByVal RangeReference As String, ByVal NewLevel As Long)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.UpdateLevel"
    ' Updates the level of the given RangeReference.
    ' RangeReference: The range reference to update.
    ' NewLevel: The new level for the range reference.

    Dim CurrentDependency As DependencyInfo
    Set CurrentDependency = This.AllDependency.Item(RangeReference)

    ' Base case to exit if the new level is less than the current level.
    If NewLevel < CurrentDependency.Level Then Exit Sub

    ' Update the level of the range reference.
    CurrentDependency.Level = NewLevel

    Dim CurrentRange As Range
    Set CurrentRange = RangeResolver.FindRangeFromText(This.FormulaInWorkbook, This.FormulaStartCellInSheet, RangeReference)

    ' Find the direct precedent cells and update their levels recursively.
    Dim DirectPrecedentCells As Collection
    If IsNothing(CurrentRange) Then
        Set DirectPrecedentCells = New Collection
    Else
        Set DirectPrecedentCells = FindDirectPrecedency(CurrentRange)
    End If
    
    Dim CurrentDirectPrecedent As DirectPrecedent
    For Each CurrentDirectPrecedent In DirectPrecedentCells
        UpdateLevel CurrentDirectPrecedent.RangeRef, NewLevel + 1
    Next CurrentDirectPrecedent

    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.UpdateLevel"
    
End Sub

Private Sub UpdateForOptional(ByRef ObjectsCollection As Collection)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.UpdateForOptional"
    ' Updates the IsOptional property for the dependencies in the ObjectsCollection.
    ' ObjectsCollection: The collection of DependencyInfo objects to update.

    Dim CurrentDependency As DependencyInfo
    For Each CurrentDependency In ObjectsCollection
        ' Check if the current dependency is a label as an input cell and not optional.
        ' If it is, update its IsOptional property based on other dependencies in the collection.
        With CurrentDependency
            If .IsLabelAsInputCell And Not .IsOptional Then
                .IsOptional = IsCurrentDependencyOptional(CurrentDependency, ObjectsCollection)
            End If
        End With
    Next CurrentDependency
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.UpdateForOptional"

End Sub

Private Sub CreateNewDependency(ByVal Level As Long _
                                , ByVal RangeReference As String _
                                 , ByVal FormulaText As String _
                                  , ByVal CurrentCell As Range _
                                   , ByVal HasAnyDependency As Boolean _
                                    , ByVal NameInFormula As String)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.CreateNewDependency"
    ' Creates a new DependencyInfo object with the provided information and adds it to AllDependency collection.
    ' Level: The level of the dependency.
    ' RangeReference: The range reference for the dependency.
    ' FormulaText: The formula text for the dependency.
    ' CurrentCell: The cell for which the dependency is being created.
    ' HasAnyDependency: True if the dependency has any direct or indirect precedents, False otherwise.
    ' NameInFormula: The name of the dependency in the formula.

    Dim ModifiedDependencyInfo As DependencyInfo
    Set ModifiedDependencyInfo = GetModifiedDependencyInfo(RangeReference)
    
    '    Debug.Assert RangeReference <> "'All Test Result'!J3:J40"
    
    Dim CurrentDependency As DependencyInfo
    Set CurrentDependency = New DependencyInfo

    With CurrentDependency
        .Level = Level
        .RangeReference = RangeReference
        .NameInFormula = NameInFormula
        .HasAnyDependency = HasAnyDependency
        
        If IsNotNothing(CurrentCell) Then
            .RowNumber = CurrentCell.Row
            .ColumnNumber = CurrentCell.Column
            .SheetName = CurrentCell.Worksheet.name
            .BookName = CurrentCell.Worksheet.Parent.name
            UpdateSheetNameVsDependencyMap CurrentCell
        End If
        
        If IsNothing(ModifiedDependencyInfo) Then
            UpdateNewDependencyIfNoModified CurrentDependency, FormulaText, RangeReference, CurrentCell, NameInFormula
        Else
            UpdateNewDependencyFromModified CurrentDependency, ModifiedDependencyInfo, RangeReference, CurrentCell
        End If

        This.AllDependency.Add CurrentDependency, RangeReference
        
        This.TotalFormulaLength = This.TotalFormulaLength + Len(CurrentDependency.FormulaText)
        
        Logger.Log INFO_LOG, vbNewLine
'        Logger.Log TRACE_LOG, CurrentDependency.ToJSON
    End With

    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.CreateNewDependency"
    
End Sub

Private Sub UpdateSheetNameVsDependencyMap(ByVal DependencyCells As Range)
    
    If IsNothing(DependencyCells) Then Exit Sub
    
    Dim SheetName As String
    SheetName = DependencyCells.Worksheet.name
    
    If IsExistInCollection(This.SheetNameVsDependencyRangeMap, SheetName) Then
        
        Dim PrevRange As Range
        Set PrevRange = This.SheetNameVsDependencyRangeMap.Item(SheetName)
        Set PrevRange = UnionOfNonExistableRange(PrevRange, DependencyCells)
        This.SheetNameVsDependencyRangeMap.Remove SheetName
        This.SheetNameVsDependencyRangeMap.Add PrevRange, SheetName
    Else
        This.SheetNameVsDependencyRangeMap.Add DependencyCells, SheetName
    End If
    
End Sub

Private Sub UpdateNewDependencyIfNoModified(ByVal CurrentDependency As DependencyInfo _
                                            , ByVal FormulaText As String _
                                             , ByVal RangeReference As String _
                                              , ByVal CurrentCell As Range _
                                               , ByVal NameInFormula As String)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.UpdateNewDependencyIfNoModified"
    With CurrentDependency
    
        .FormulaText = FormulaText
        .HasFormula = HasFormula(CurrentCell)
            
        If IsNothing(CurrentCell) And IsLocalOrGlobalNamedRangeExist(NameInFormula, This.FormulaStartCellSheetName) Then
            ' For const named range we will have CurrentCell nothing.
            .RangeLabel = RemoveSheetQualifierIfPresent(NameInFormula)
            .IsLabelAsInputCell = False
            .IsReferByNamedRange = True
            .IsInsideNamedRangeOrTable = True
            .IsTableRef = False
            Set .ForName = GetNamedRangeByNameInFormula(NameInFormula, This.FormulaStartCellSheetName)
        Else
            '@TODO: Need to update it here.
            .RangeLabel = FindRangeLabel(RangeReference, CurrentCell, False)
            Set .RangeLabelCell = FindLetVarNameCell(CurrentCell)
            .IsLabelAsInputCell = IsInputCell(CurrentCell, This.DependencySearchInRegion)
            .IsReferByNamedRange = IsLocalOrGlobalNamedRangeExist(NameInFormula, CurrentCell.Worksheet.name)
            .IsInsideNamedRangeOrTable = IsInsideTableOrNamedRange(CurrentCell)
            
            If .IsReferByNamedRange Then
                Set .ForName = GetNamedRangeByNameInFormula(NameInFormula, CurrentCell.Worksheet.name)
            End If
            
            If .IsInsideNamedRangeOrTable Then
                .IsTableRef = Not (IsReferredByCellAddress(CurrentCell, NameInFormula) Or .IsReferByNamedRange)
            End If
            
        End If
     
        .ValidVarName = MakeValidLetVarName(.RangeLabel, GetNamingConv(.IsLabelAsInputCell))
        .ValidVarName = CorrectValidVarNameForNamedRange(.ValidVarName, .IsReferByNamedRange, .IsLabelAsInputCell)
        
    End With
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.UpdateNewDependencyIfNoModified"
    
End Sub

Private Function IsReferredByCellAddress(ByVal CurrentRange As Range, ByVal NameInFormula As String) As Boolean
    
    Dim UpdatedNameInFormula As String
    UpdatedNameInFormula = RemoveDollarSign(NameInFormula)
    
    Dim Result As Boolean
    If CurrentRange.Cells(1).Address(False, False) = UpdatedNameInFormula Then
        Result = True
    ElseIf GetRangeRefWithSheetName(CurrentRange.Cells(1)) = UpdatedNameInFormula Then
        Result = True
    ElseIf CurrentRange.Cells(1).Address(False, False) & DYNAMIC_CELL_REFERENCE_SIGN = UpdatedNameInFormula Then
        Result = True
    ElseIf GetRangeRefWithSheetName(CurrentRange.Cells(1)) & DYNAMIC_CELL_REFERENCE_SIGN = UpdatedNameInFormula Then
        Result = True
    Else
        Result = False
    End If
    
    IsReferredByCellAddress = Result
    
End Function

Private Sub UpdateNewDependencyFromModified(ByVal CurrentDependency As DependencyInfo _
                                            , ByVal ModifiedDependencyInfo As DependencyInfo _
                                             , ByVal RangeReference As String _
                                             , ByVal CurrentCell As Range)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.UpdateNewDependencyFromModified"
    With CurrentDependency
        
        .HasFormula = ModifiedDependencyInfo.HasFormula
        .IsUserSpecifiedName = ModifiedDependencyInfo.IsUserSpecifiedName
        If Not .HasFormula Then
            .NameInFormula = RemoveHashSignFromEnd(.NameInFormula)
        End If

        .FormulaText = ModifiedDependencyInfo.FormulaText

        If ModifiedDependencyInfo.IsExpandByUser Then
            '@TODO: Need to update here.
            .RangeLabel = FindRangeLabel(RangeReference, CurrentCell, True)
            Set .RangeLabelCell = FindLetVarNameCell(CurrentCell)
            .ValidVarName = MakeValidLetVarName(.RangeLabel, GetNamingConv(False))
        Else
            .RangeLabel = ModifiedDependencyInfo.RangeLabel
            .ValidVarName = ModifiedDependencyInfo.ValidVarName
        End If

        If ModifiedDependencyInfo.IsUserMarkAsParameterCell Then
            .IsLabelAsInputCell = True
        ElseIf ModifiedDependencyInfo.IsDemotedFromParameterCellToLetStep Then
            .IsLabelAsInputCell = False
        ElseIf ModifiedDependencyInfo.IsUserMarkAsValue Then
            .IsLabelAsInputCell = False
        Else
            .IsLabelAsInputCell = IsInputCell(CurrentCell, This.DependencySearchInRegion)
        End If

        .IsReferByNamedRange = ModifiedDependencyInfo.IsReferByNamedRange
        .IsInsideNamedRangeOrTable = ModifiedDependencyInfo.IsInsideNamedRangeOrTable
        .IsUserMarkAsParameterCell = ModifiedDependencyInfo.IsUserMarkAsParameterCell
        .IsMarkAsNotLetStatementByUser = ModifiedDependencyInfo.IsMarkAsNotLetStatementByUser
        .IsUserMarkAsValue = ModifiedDependencyInfo.IsUserMarkAsValue
        Set .ForName = ModifiedDependencyInfo.ForName
        .IsDemotedFromParameterCellToLetStep = ModifiedDependencyInfo.IsDemotedFromParameterCellToLetStep
        .IsInsideNamedRangeOrTable = ModifiedDependencyInfo.IsInsideNamedRangeOrTable
        .IsTableRef = ModifiedDependencyInfo.IsTableRef
        .IsOptional = ModifiedDependencyInfo.IsOptional
        
    End With
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.UpdateNewDependencyFromModified"
    
End Sub

Private Function IsLocalOrGlobalNamedRangeExist(ByVal NameOfNamedRange As String _
                                                , ByVal SheetName As String) As Boolean
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.IsLocalOrGlobalNamedRangeExist"
    ' Checks if the named range exists in the global named range collection or in the specified sheet's named range collection.
    ' NameOfNamedRange: The name of the named range to check.
    ' SheetName: The name of the sheet where the named range might exist.

    ' Check if the named range exists in the global named range collection.
    Dim Result As Boolean
    Result = This.AllNamedRange.Exists(NameOfNamedRange)

    ' If the named range doesn't exist in the global collection, check if it exists in the specified sheet's named range collection.
    If Not Result Then
        SheetName = GetSheetRefForRangeReference(SheetName, False)
        Result = This.AllNamedRange.Exists(SheetName & NameOfNamedRange)
    End If
    
    IsLocalOrGlobalNamedRangeExist = Result
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.IsLocalOrGlobalNamedRangeExist"
    
End Function

Private Function GetNamedRangeByNameInFormula(ByVal NameOfNamedRange As String _
                                              , ByVal SheetName As String) As name
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.GetNamedRangeByNameInFormula"
    ' Checks if the named range exists in the global named range collection or in the specified sheet's named range collection.
    ' NameOfNamedRange: The name of the named range to check.
    ' SheetName: The name of the sheet where the named range might exist.

    ' Check if the named range exists in the global named range collection.
    Dim Result As name
    SheetName = GetSheetRefForRangeReference(SheetName, False)
    If This.AllNamedRange.Exists(NameOfNamedRange) Then
        Set Result = This.AllNamedRange.Item(NameOfNamedRange)
    ElseIf This.AllNamedRange.Exists(SheetName & NameOfNamedRange) Then
        ' If the named range doesn't exist in the global collection, check if it exists in the specified sheet's named range collection.
        Set Result = This.AllNamedRange.Item(SheetName & NameOfNamedRange)
    End If
    
    Set GetNamedRangeByNameInFormula = Result
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.GetNamedRangeByNameInFormula"
    
End Function

Private Function GetModifiedDependencyInfo(ByVal RangeReference As String) As DependencyInfo
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.GetModifiedDependencyInfo"
    ' Gets the modified DependencyInfo object from the ModifiedDependencies collection if it exists.
    ' RangeReference: The range reference to check for modified dependency information.

    ' Check if the modified DependencyInfo object exists in the ModifiedDependencies collection.
    If IsExistInCollection(This.ModifiedDependencies, RangeReference) Then
        Set GetModifiedDependencyInfo = This.ModifiedDependencies.Item(RangeReference)
    End If
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.GetModifiedDependencyInfo"

End Function

Private Function CorrectValidVarNameForNamedRange(ByVal ValidVarName As String _
                                                  , ByVal IsReferedByNamedRange As Boolean _
                                                  , ByVal IsInputCell As Boolean) As String
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.CorrectValidVarNameForNamedRange"
    ' Corrects the valid variable name for a named range that is also referred to by another named range.
    ' ValidVarName: The valid variable name for the named range.
    ' IsReferedByNamedRange: True if the named range is referred to by another named range, False otherwise.

    ' If the named range is referred to by another named range and the current name exists in the named range collection,
    ' add an underscore to the beginning of the name.
    If IsReferedByNamedRange And IsLocalOrGlobalNamedRangeExist(ValidVarName, This.FormulaStartCellSheetName) Then
        ValidVarName = Text.PadIfNotPresent(ValidVarName, UNDER_SCORE, FROM_START)
    End If

    ' If the corrected valid variable name does not exist in the named range collection, return it.
    If Not IsLocalOrGlobalNamedRangeExist(ValidVarName, This.FormulaStartCellSheetName) Then
        CorrectValidVarNameForNamedRange = ValidVarName
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaPrecedencyExtractor.CorrectValidVarNameForNamedRange"
        Exit Function
    End If

    ' If the corrected valid variable name already exists in the named range collection, append a counter to the name until
    ' a unique name is found.
    
    If Not IsInputCell Then
        If Not IsLocalOrGlobalNamedRangeExist(Text.PadIfNotPresent(ValidVarName, UNDER_SCORE, FROM_START), This.FormulaStartCellSheetName) Then
            CorrectValidVarNameForNamedRange = ValidVarName
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaPrecedencyExtractor.CorrectValidVarNameForNamedRange"
            Exit Function
        End If
    End If
    
    Dim Counter As Long
    Counter = 2
    Do While True
        If IsLocalOrGlobalNamedRangeExist(ValidVarName & UNDER_SCORE & Counter, This.FormulaStartCellSheetName) Then
            Counter = Counter + 1
        Else
            Exit Do
        End If
    Loop

    ' Return the corrected valid variable name with the counter appended.
    CorrectValidVarNameForNamedRange = ValidVarName & UNDER_SCORE & Counter
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.CorrectValidVarNameForNamedRange"
    
End Function

Private Function GenerateDependencyStats(ByVal AllDependency As Collection) As DependenciesStat
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.GenerateDependencyStats"
    ' Generates dependency statistics from the provided DependencyInfo collection.
    ' AllDependency: The collection of DependencyInfo objects to analyze.

    Dim Stats As DependenciesStat
    Set Stats = New DependenciesStat

    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In AllDependency
        ' Update the input cell stats and the level in the dependency statistics.
        Stats.UpdateInputCellStats CurrentDependencyInfo.IsLabelAsInputCell
        Stats.UpdateLevel CurrentDependencyInfo.Level
    Next CurrentDependencyInfo

    Set GenerateDependencyStats = Stats
    Set Stats = Nothing
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.GenerateDependencyStats"

End Function

Private Sub MaximizeLevel(ByRef AllDependency As Collection)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.MaximizeLevel"
    ' Maximize the level of the dependencies to ensure they have the maximum level possible.

    ' Generate dependency statistics from the provided DependencyInfo collection.
    Dim Stats As DependenciesStat
    Set Stats = GenerateDependencyStats(AllDependency)

    ' Update the level of input cells to the maximum level found in the dependencies.
    UpdateInputCellsLevelToMax AllDependency, Stats

    ' Update the level of the dependencies based on their precedents.
    UpdateLevelBasedOnPrecedents AllDependency, Stats
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.MaximizeLevel"

End Sub

Private Sub UpdateInputCellsLevelToMax(ByRef AllDependency As Collection, ByVal Stats As DependenciesStat)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.UpdateInputCellsLevelToMax"
    ' Update the level of input cells to the maximum level found in the dependencies.
    ' AllDependency: The collection of DependencyInfo objects.
    ' Stats: The dependency statistics.

    Dim MaxLevel As Long
    MaxLevel = Stats.MaximumLevel

    ' Check if multi-cell ranges should be considered as input cells based on the statistics.
    Dim IsConsiderMultiCellAsInputCell As Boolean
    IsConsiderMultiCellAsInputCell = (Not Stats.HasAnyCellLabelAsInputCell)

    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In AllDependency
        With CurrentDependencyInfo
            ' If the dependency is an input cell, set its level to the maximum level.
            If .IsLabelAsInputCell Then
                
                If .Level < MaxLevel Then
                    .Level = MaxLevel
                    Logger.Log INFO_LOG, .RangeReference & " level has been upgrade to " & MaxLevel
                End If
                
            ElseIf .IsDemotedFromParameterCellToLetStep Then
                .IsLabelAsInputCell = False
            ElseIf .IsUserMarkAsValue Then
                .IsLabelAsInputCell = False
            ElseIf IsMultiCellNotDynamic(.RangeReference) And IsConsiderMultiCellAsInputCell _
                   And Not .IsReferByNamedRange Then
                ' If the dependency is a multi-cell range and should be considered as an input cell, set its level to the maximum level.
                If .Level < MaxLevel Then
                    .Level = MaxLevel
                    Logger.Log INFO_LOG, .RangeReference & " level has been upgrade to " & MaxLevel
                End If
                .IsLabelAsInputCell = True
            ElseIf Not (.HasAnyDependency Or Stats.HasAnyCellLabelAsInputCell Or .HasFormula) Then
                ' If the dependency has no dependencies, no input cells were found in the statistics, and it has no formula, set its level to the maximum level.
                If .Level < MaxLevel Then
                    .Level = MaxLevel
                    Logger.Log INFO_LOG, .RangeReference & " level has been upgrade to " & MaxLevel
                End If
                .IsLabelAsInputCell = True
            End If
        End With
    Next CurrentDependencyInfo
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.UpdateInputCellsLevelToMax"

End Sub

Private Function IsMultiCellNotDynamic(ByVal RangeReference As String) As Boolean
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.IsMultiCellNotDynamic"
    ' Checks if the given range reference represents a multi-cell range that is not dynamic.
    ' RangeReference: The range reference to check.

    ' If the range reference ends with the dynamic cell reference sign, it is considered dynamic and not a multi-cell range.
    If IsSpilledRangeRef(RangeReference) Then
        IsMultiCellNotDynamic = False
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaPrecedencyExtractor.IsMultiCellNotDynamic"
        Exit Function
    End If

    ' Get the range from the range reference.
    Dim RangeFromRef As Range
    Set RangeFromRef = RangeResolver.GetRange(RangeReference, This.FormulaInWorkbook, This.FormulaStartCellInSheet)
    
    Dim Result As Boolean
    ' Check if the range exists and if it contains more than one cell and is not a spill range.
    If IsNothing(RangeFromRef) Then
        Result = False
    ElseIf RangeFromRef.Cells.Count > 1 Then
        If IsNull(RangeFromRef.HasSpill) Then
            Result = True
        Else
            Result = Not RangeFromRef.HasSpill
        End If
    End If
    
    IsMultiCellNotDynamic = Result
    
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.IsMultiCellNotDynamic"

End Function

Private Sub UpdateLevelBasedOnPrecedents(ByRef AllDependency As Collection, ByVal Stats As DependenciesStat)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.UpdateLevelBasedOnPrecedents"
    ' Update the level of the dependencies based on their precedents.
    ' AllDependency: The collection of DependencyInfo objects.
    ' Stats: The dependency statistics.

    ' Check if multi-cell ranges should be considered as input cells based on the statistics.
    Dim IsConsiderMultiCellAsInputCell As Boolean
    IsConsiderMultiCellAsInputCell = (Not Stats.HasAnyCellLabelAsInputCell)

    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In AllDependency
        With CurrentDependencyInfo
            ' If the dependency is not an input cell based on certain criteria, find the minimum level from all its dependent cells and update its level.
            If IsNotInputCellInAnyCriteria(.IsLabelAsInputCell, .RangeReference, IsConsiderMultiCellAsInputCell) Then
                Dim MinLevel As Long
                MinLevel = FindMinDependencyLevelFromAllDependentCells(AllDependency, CurrentDependencyInfo)
                .Level = MinLevel
            End If
        End With
    Next CurrentDependencyInfo

    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.UpdateLevelBasedOnPrecedents"
    
End Sub

Private Function IsNotInputCellInAnyCriteria(ByVal IsCellLabelAsInput As Boolean _
                                             , ByVal RangeReference As String _
                                              , ByVal IsConsiderMultiCellAsInputCell As Boolean) As Boolean
   
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.IsNotInputCellInAnyCriteria"
    ' Check if cell is labelled as input, or if multi-cell is not dynamic and we are considering multi-cell as input
    If IsCellLabelAsInput Then
        IsNotInputCellInAnyCriteria = False
    ElseIf (IsMultiCellNotDynamic(RangeReference) And IsConsiderMultiCellAsInputCell) Then
        IsNotInputCellInAnyCriteria = False
    Else
        IsNotInputCellInAnyCriteria = True
    End If
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.IsNotInputCellInAnyCriteria"
    
End Function

Private Function FindMinDependencyLevelFromAllDependentCells(ByRef AllDependency As Collection _
                                                             , ByVal CurrentDependencyInfo As DependencyInfo) As Long
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.FindMinDependencyLevelFromAllDependentCells"
    Dim DirectPrecedentsCell As Collection
    Dim ForCell As Range
    ' Get the range of cells for current dependency
    Set ForCell = RangeResolver.GetRange(CurrentDependencyInfo.RangeReference, This.FormulaInWorkbook, This.FormulaStartCellInSheet)
    ' Find the direct precedents of the current cell
    If IsNothing(ForCell) Then
        Set DirectPrecedentsCell = New Collection
    Else
        Set DirectPrecedentsCell = FindDirectPrecedency(ForCell)
    End If
    
    Dim CurrentDirectPrecedent As DirectPrecedent
    Dim MinLevel As Long
    MinLevel = MAXIMUM_ALLOWABLE_DEPENDENCY_LEVEL
    
    ' Iterate over all precedents to find minimum dependency level
    For Each CurrentDirectPrecedent In DirectPrecedentsCell
        If AllDependency.Item(CurrentDirectPrecedent.RangeRef).Level < MinLevel Then
            MinLevel = AllDependency.Item(CurrentDirectPrecedent.RangeRef).Level
        End If
    Next CurrentDirectPrecedent
    
    ' We are adding one because value of FindMinDependencyLevelFromAllDependentCells is 1 less than MinLevel
    If MinLevel = MAXIMUM_ALLOWABLE_DEPENDENCY_LEVEL Then MinLevel = CurrentDependencyInfo.Level + 1
    ' Return minimum dependency level minus one
    FindMinDependencyLevelFromAllDependentCells = MinLevel - 1
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.FindMinDependencyLevelFromAllDependentCells"
    
End Function

'@Exposed
Public Function GetDependencyObjectsFromData(ByVal DependencyOrderData As Variant _
                                             , ByVal TypeOfOperation As OperationType) As Collection
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.GetDependencyObjectsFromData"
    Dim DependencyObjects As Collection
    Const RANGE_REFERENCE_COL_INDEX As Long = 2
    ' Create DependencyInfo objects from the input array and add them to the collection
    Set DependencyObjects = modUtility.CreateObjectsFromArray(DependencyOrderData _
                                                              , New DependencyInfo _
                                                               , "CreateMe", RANGE_REFERENCE_COL_INDEX)
    ' Update various properties of the DependencyInfo objects in the collection
    SetNamedRangeIfReferedByNamedRange DependencyObjects
    UpdateInputCellsForDefaultName DependencyObjects
    UpgradeDuplicateLetVarName DependencyObjects
    UpdateBlankRangeLabelForDefaultName DependencyObjects
    This.NumberOfNonInputDependency = GetNumberOfNonInputDependency(DependencyObjects)
    If TypeOfOperation <> LET_STATEMENT_GENERATION Then
        ' In case of LET Statement generation use the label instead of _Result
        UpdateLastStepName DependencyObjects
    End If
    
    ' Handle potential duplicates
    UpgradeDuplicateLetVarName DependencyObjects
    Set GetDependencyObjectsFromData = DependencyObjects
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.GetDependencyObjectsFromData"
    
End Function

Public Sub SetNamedRangeIfReferedByNamedRange(ByRef ObjectsCollection As Collection)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.SetNamedRangeIfReferedByNamedRange"
    Dim CurrentDependency As DependencyInfo
    ' Iterate over each DependencyInfo object in the collection and add an underscore to the ValidVarName property if it is a valid dependency
    For Each CurrentDependency In ObjectsCollection
        With CurrentDependency
            If .IsReferByNamedRange Then
                
                ' As we have decided to only trace precedents on FormulaInBook
                ' we can easily check in the FormulaInBook scope.
                If IsLocalOrGlobalNamedRangeExist(.NameInFormula, This.FormulaStartCellSheetName) Then
                    Set .ForName = GetNamedRangeByNameInFormula(.NameInFormula, This.FormulaStartCellSheetName)
                End If
                
            End If
        End With
    Next CurrentDependency
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.SetNamedRangeIfReferedByNamedRange"
    
End Sub

Public Sub AddPrefixOnInputCells(ByRef ObjectsCollection As Collection)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.AddPrefixOnInputCells"
    Dim LetVarPrefix As String
    LetVarPrefix = GetLetVarPrefix()
    
    Dim CurrentDependency As DependencyInfo
    ' Iterate over each DependencyInfo object in the collection and add an underscore to the ValidVarName property if it is a valid dependency
    For Each CurrentDependency In ObjectsCollection
        With CurrentDependency
            If .ValidVarName <> vbNullString And Not .IsUserSpecifiedName And .IsLabelAsInputCell Then
                .ValidVarName = Text.PadIfNotPresent(.ValidVarName, LetVarPrefix, FROM_START)
            End If
        End With
    Next CurrentDependency
    
    UpgradeDuplicateLetVarName ObjectsCollection, True
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.AddPrefixOnInputCells"
    
End Sub

Public Sub AddPrefixOnValidVarNameForNonInputCells(ByRef ObjectsCollection As Collection)
         
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.AddPrefixOnValidVarNameForNonInputCells"
    Dim LetVarPrefix As String
    LetVarPrefix = GetLetVarPrefix()
    
    Dim IsAddPrefixOnParam As Boolean
    IsAddPrefixOnParam = modDependencyLambdaResult.GetAddPrefixOnParamValue()
    
    Dim CurrentDependency As DependencyInfo
    ' Iterate over each DependencyInfo object in the collection and add an underscore to the ValidVarName property if it is a valid dependency
    For Each CurrentDependency In ObjectsCollection
        With CurrentDependency
            If .ValidVarName <> vbNullString And Not .IsUserSpecifiedName Then
                
                ' If param cell then use naming conv for param.
                If .IsLabelAsInputCell And IsAddPrefixOnParam Then
                    .ValidVarName = Text.PadIfNotPresent(MakeValidLetVarName(.RangeLabel, GetNamingConv(True)), LetVarPrefix)
                ElseIf .IsLabelAsInputCell Then
                    .ValidVarName = MakeValidLetVarName(.RangeLabel, GetNamingConv(True))
                Else
                    .ValidVarName = Text.PadIfNotPresent(.ValidVarName, LetVarPrefix, FROM_START)
                End If
                
            End If
        End With
    Next CurrentDependency
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.AddPrefixOnValidVarNameForNonInputCells"

End Sub

Private Sub UpdateLastStepName(ByRef DependencyObjects As Collection)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.UpdateLastStepName"
    ' Check if there is more than one non-input dependency
    If GetNumberOfNonInputDependency(DependencyObjects) > 1 Then
        Dim CurrentDependencyInfo  As DependencyInfo
        ' Get the last dependency in the collection
        Set CurrentDependencyInfo = DependencyObjects.Item(DependencyObjects.Count)
        
        ' Retrieve the ValidVarName of the second last dependency
        Dim PreviousStepName As String
        PreviousStepName = DependencyObjects.Item(DependencyObjects.Count - 1).ValidVarName
        ' Update the ValidVarName of the last dependency based on the casing of the second to last dependency
        Dim Prefix As String
        Prefix = GetLetVarPrefix()
        
        If Not CurrentDependencyInfo.IsUserSpecifiedName Then
            CurrentDependencyInfo.ValidVarName = Prefix & MakeValidLetVarName(LCase$(LAST_STEP_NAME), GetNamingConv(False))
        End If
        
        This.VarNameOfFormulaStartCell = CurrentDependencyInfo.ValidVarName
    End If
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.UpdateLastStepName"
    
End Sub

Private Sub UpdateInputCellsForDefaultName(ByRef DependencyObjects As Collection)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.UpdateInputCellsForDefaultName"
    ' Iterate over the dependencies and update the ValidVarName of input cells without a name
    Dim CurrentDependencyInfo As DependencyInfo
    Dim Counter As Long
    For Each CurrentDependencyInfo In DependencyObjects
        Counter = Counter + 1
        With CurrentDependencyInfo
            If .IsLabelAsInputCell And .ValidVarName = vbNullString Then
                .ValidVarName = "Input" & Counter
            End If
        End With
    Next CurrentDependencyInfo
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.UpdateInputCellsForDefaultName"
    
End Sub

Public Sub UpdateInputCellsValidVarNameNamingConvIfLetStepGen(ByRef DependencyObjects As Collection)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.UpdateInputCellsValidVarNameNamingConvIfLetStepGen"
    ' Iterate over the dependencies and update the ValidVarName of input cells without a name
    Dim CurrentDependencyInfo As DependencyInfo
    
    Dim Prefix As String
    Prefix = GetLetVarPrefix()
    
    Dim NamingConv As VarNamingStyle
    NamingConv = GetNamingConv(False)
    
    For Each CurrentDependencyInfo In DependencyObjects
        With CurrentDependencyInfo
            If .IsLabelAsInputCell And Not .IsUserSpecifiedName Then
                .ValidVarName = Prefix & MakeValidLetVarName(.ValidVarName, NamingConv)
            End If
        End With
    Next CurrentDependencyInfo
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.UpdateInputCellsValidVarNameNamingConvIfLetStepGen"
    
End Sub

Private Sub UpdateBlankRangeLabelForDefaultName(ByRef DependencyObjects As Collection)
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.UpdateBlankRangeLabelForDefaultName"
    ' Iterate over the dependencies and update the ValidVarName of non-input cells without a name
    Dim CurrentDependencyInfo As DependencyInfo
    Dim Counter As Long
    For Each CurrentDependencyInfo In DependencyObjects
        With CurrentDependencyInfo
            If Not .IsLabelAsInputCell Then Counter = Counter + 1
            If .ValidVarName = vbNullString Then
                .ValidVarName = "_Step" & Counter
            End If
        End With
    Next CurrentDependencyInfo
    
    ' Update the ValidVarName of the last dependency if necessary
    Set CurrentDependencyInfo = DependencyObjects.Item(DependencyObjects.Count)
    
    With CurrentDependencyInfo
        If Not .IsLabelAsInputCell And .ValidVarName = "_Step" & Counter And Not .IsUserSpecifiedName Then
            .ValidVarName = GetLetVarPrefix() & MakeValidLetVarName(LCase$(LAST_STEP_NAME), GetNamingConv(False))
        End If
    End With
    
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.UpdateBlankRangeLabelForDefaultName"
    
End Sub

Private Sub UpgradeDuplicateLetVarName(ByRef DependencyObjects As Collection _
                                       , Optional ByVal UniqueInBothParamAndStep As Boolean = False)

    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.UpgradeDuplicateLetVarName"
    ' Initialize a new collection to store the variable names encountered so far
    Dim SoFarVarNameCollection As Collection
    Set SoFarVarNameCollection = New Collection

    ' Loop through the Dependency Objects Collection and replace duplicate ValidVarName with unique names
    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In DependencyObjects
        Dim CurrentValidVarName As String
        CurrentValidVarName = CurrentDependencyInfo.ValidVarName
        If CurrentValidVarName <> vbNullString Then
            ' Find a unique name for the current dependency
            CurrentValidVarName = FindUniqueNameForCurrentDependency(CurrentDependencyInfo _
                                                                     , SoFarVarNameCollection _
                                                                     , UniqueInBothParamAndStep)
            CurrentDependencyInfo.ValidVarName = CurrentValidVarName
            
            ' So if we have same name in lambda param and in the LET Step then we want to store the LET Version.
            If IsExistInCollection(SoFarVarNameCollection, CurrentValidVarName) Then
                SoFarVarNameCollection.Remove CurrentValidVarName
            End If
            SoFarVarNameCollection.Add CurrentDependencyInfo, CurrentValidVarName
            
        End If
    Next CurrentDependencyInfo
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.UpgradeDuplicateLetVarName"

End Sub

Private Function FindUniqueNameForCurrentDependency(ByVal CurrentDependencyInfo As DependencyInfo _
                                                    , ByVal SoFarVarNameCollection As Collection _
                                                    , ByVal UniqueInBothParamAndStep As Boolean) As String
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.FindUniqueNameForCurrentDependency"
    ' Initialize a variable to store the ValidVarName
    Dim ValidVarName As String
    ValidVarName = CurrentDependencyInfo.ValidVarName
    
    ' If the current dependency is an input cell, find a unique name by incrementing a number
    If CurrentDependencyInfo.IsLabelAsInputCell Then
        FindUniqueNameForCurrentDependency = FindUniqueNameByIncrementingNumber(SoFarVarNameCollection _
                                                                                , ValidVarName)
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaPrecedencyExtractor.FindUniqueNameForCurrentDependency"
        Exit Function
    End If
    
    ' If no conflict then just exit
    If Not IsExistInCollection(SoFarVarNameCollection, ValidVarName) Then
        FindUniqueNameForCurrentDependency = ValidVarName
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FormulaPrecedencyExtractor.FindUniqueNameForCurrentDependency"
        Exit Function
    End If

    ' Resolve any conflicts by updating the ValidVarName
    Do While modUtility.IsExistInCollection(SoFarVarNameCollection, ValidVarName)
    
        Dim ConflictWithDependency As DependencyInfo
        Set ConflictWithDependency = SoFarVarNameCollection.Item(ValidVarName)
        
        If ConflictWithDependency.IsLabelAsInputCell And Not UniqueInBothParamAndStep Then
            Exit Do
        End If
        
        ' Update ValidVarName if there is a conflict
        If Text.ExtractNumberFromEnd(ValidVarName) = vbNullString Then
            ValidVarName = Text.PadIfNotPresent(ValidVarName, "_2", FROM_END)
        Else
            ValidVarName = Text.IncrementOrDecrementEndNumber(ValidVarName, 1, False)
        End If
        
    Loop
    
    FindUniqueNameForCurrentDependency = ValidVarName
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.FindUniqueNameForCurrentDependency"
    
End Function

Private Sub Class_Terminate()
    
    Logger.Log TRACE_LOG, "Enter FormulaPrecedencyExtractor.Class_Terminate"
    ' Check if the formula start cell exists and if so, clear the cache
    If IsNotNothing(This.FormulaStartCell) Then ClearCatches
    ' Nullify the mapping objects
    Set This.StructuredFormulaManager = Nothing
    Logger.Log TRACE_LOG, "Exit FormulaPrecedencyExtractor.Class_Terminate"
    
End Sub

Public Function FindStructuralFormulaText(ByVal ForCell As Range, ByVal FormulaText As String) As String
    FindStructuralFormulaText = This.StructuredFormulaManager.FindStructuralFormulaText(ForCell, FormulaText)
End Function


