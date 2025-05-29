Attribute VB_Name = "modDependencyFormulaReplacer"
Option Explicit
Option Private Module

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Include Lambda Dependencies
' Description:            Include lambda dependencies.
' Macro Expression:       modDependencyFormulaReplacer.IncludeLambdaDependencies([Selection.Cells(1)])
' Generated:              10/19/2022 02:47 PM
'----------------------------------------------------------------------------------------------------
Public Sub IncludeLambdaDependencies(ByVal LambdaInCell As Range _
                                     , Optional ByVal IsUndo As Boolean = False _
                                      , Optional ByVal IsOnlyLetStepOnes As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter modDependencyFormulaReplacer.IncludeLambdaDependencies"
    Const METHOD_NAME As String = "IncludeLambdaDependencies"
    Context.ExtractContextFromCell LambdaInCell, METHOD_NAME
    ' Static variables to hold Undo information
    Static PutFormulaOnUndo As Range
    Static OldFormula As String

    ' If it is Undo operation, restore old formula and exit subroutine
    If IsUndo Then
        If IsNotNothing(PutFormulaOnUndo) Then
            PutFormulaOnUndo.Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
            AutofitFormulaBar PutFormulaOnUndo
        End If
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modDependencyFormulaReplacer.IncludeLambdaDependencies"
        GoTo ExitMethod
    Else
        ' Otherwise, store the current formula to be used for undo operation in future
        Set PutFormulaOnUndo = LambdaInCell
        If IsNotNothing(PutFormulaOnUndo) Then OldFormula = LambdaInCell.Formula2
    End If

    ' If the current command is invalid, exit subroutine
    If modUtility.IsWorkbookOrWorksheetProtected(LambdaInCell, "Include Lambda Dependencies") Then GoTo ExitMethod

    ' Create a new DependencyFormulaReplacer and include Lambda dependencies
    Dim DependencyReplacer As DependencyFormulaReplacer
    Set DependencyReplacer = New DependencyFormulaReplacer
    DependencyReplacer.IncludeLambdaDependencies LambdaInCell, Nothing, UPDATE_FORMULA_IN_CELL, IsOnlyLetStepOnes
    AutofitFormulaBar LambdaInCell
    
    If Not IsUndo Then AssingOnUndo "IncludeLambdaDependencies"
    Logger.Log TRACE_LOG, "Exit modDependencyFormulaReplacer.IncludeLambdaDependencies"
    
ExitMethod:
    Context.ClearContext METHOD_NAME
    Exit Sub
    
End Sub

Private Sub IncludeLambdaDependencies_Undo()
    IncludeLambdaDependencies Nothing, True
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Generate Lambda Formula Dependency
' Description:            Generate lambda formula dependency.
' Macro Expression:       modDependencyFormulaReplacer.GenerateLambdaFormulaDependency([Selection.Cells(1)],[NewTableTargetToRight])
' Generated:              10/19/2022 03:25 PM
'----------------------------------------------------------------------------------------------------
Public Sub GenerateLambdaFormulaDependency(ByVal LambdaInCell As Range _
                                           , ByVal PutDependencyInRange As Range _
                                            , Optional ByVal IsUndo As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter modDependencyFormulaReplacer.GenerateLambdaFormulaDependency"
    Const METHOD_NAME As String = "GenerateLambdaFormulaDependency"
    Context.ExtractContextFromCell LambdaInCell, METHOD_NAME
    ' Static variables to handle undo operations
    Static Table As ListObject
    Static PutFormulaOnUndo As Range

    ' If undo operation, delete the table and select the original formula, then exit
    If IsUndo Then
        If IsNotNothing(Table) Then Table.Delete
        If IsNotNothing(Table) Then PutFormulaOnUndo.Select
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modDependencyFormulaReplacer.GenerateLambdaFormulaDependency"
        GoTo ExitMethod
    End If

    ' Check for command validity, exit if invalid
    If modUtility.IsWorkbookOrWorksheetProtected(LambdaInCell, "Generate Lambda Formula Dependency") Then GoTo ExitMethod

    ' Create DependencyFormulaReplacer and include Lambda dependencies
    Dim DependencyReplacer As DependencyFormulaReplacer
    Set DependencyReplacer = New DependencyFormulaReplacer
    DependencyReplacer.IncludeLambdaDependencies LambdaInCell, PutDependencyInRange, SEND_RESULT_TO_SHEET, False

    ' Handle non-undo operation, set table for undo and assign undo operation
    If Not IsUndo Then
        Set Table = DependencyReplacer.PutDependencyOnTable
        Set PutFormulaOnUndo = LambdaInCell
        AssingOnUndo "GenerateLambdaFormulaDependency"
    End If
    Logger.Log TRACE_LOG, "Exit modDependencyFormulaReplacer.GenerateLambdaFormulaDependency"
   
ExitMethod:
   Context.ClearContext METHOD_NAME
   Exit Sub
   
End Sub

Private Sub GenerateLambdaFormulaDependency_Undo()
    GenerateLambdaFormulaDependency Nothing, Nothing, True
End Sub


