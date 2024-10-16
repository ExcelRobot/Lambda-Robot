Attribute VB_Name = "modUnUsedLambdas"
Option Explicit
Option Private Module

Private Sub TestRemoveUnusedLambdas()
    RemoveUnusedLambdas ActiveWorkbook
End Sub

Public Sub RemoveUnusedLambdas(Optional ByVal FromBook As Workbook)
    
    Logger.Log TRACE_LOG, "Enter modUnUsedLambdas.RemoveUnusedLambdas"
    If FromBook Is Nothing Then Set FromBook = ActiveWorkbook
    
    Dim AllLambdas As Collection
    Set AllLambdas = New Collection
    
    Dim AllUniqueFormulas As Collection
    Set AllUniqueFormulas = New Collection
    
    ' Extract all the formulas from named formulas.
    ' At the same time Collect all lambdas.
    Dim CurrentName As name
    For Each CurrentName In FromBook.Names
        
        ' If it is built in name like _xlpm or _xlfn then ignore them.
        If Not IsBuiltInName(CurrentName) Then
            With CurrentName
                If IsLambdaFunction(.RefersTo) Then
                    ' Add the name to the collection of lambda.
                    AllLambdas.Add CurrentName, CStr(.name)
                Else
                    AddToCollectionIfNotExist AllUniqueFormulas _
                                              , .RefersToR1C1 _
                                               , FormulaInfo.Create(.RefersToR1C1, .name & ".RefersToR1C1", True)
                End If
            End With
        End If
    
    Next CurrentName
    
    ' Loop through each sheet and find all the formulas from cell and conditional formatting.
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In FromBook.Worksheets
        
        UpdateFormulaCollFromCellFormulas AllUniqueFormulas, CurrentSheet
        UpdateFormulaCollFromCF AllUniqueFormulas, CurrentSheet
        
    Next CurrentSheet
    
    ' Keep only those formulas where we have at least one lambda present by checking string contains.
    KeepFormulasIfLambdaIsUsedByTextParsing AllUniqueFormulas, AllLambdas
    
    ' Extract all the used lambdas in all of those formulas.
    Dim AllUsedLambdas As Collection
    Set AllUsedLambdas = GetAllUsedLambdas(AllUniqueFormulas, AllLambdas)
    
    ' No update used lambdas Collection for lambdas dependency on another lambdas.
    UpdateUsedLambdasForDependencies AllUsedLambdas, AllLambdas
    
    ' No extract all the unused lambdas.
    Dim UnusedLambdas As Collection
    Set UnusedLambdas = GetUnusedLambdas(AllLambdas, AllUsedLambdas)
    
    ' Delete all the unused lambdas.
    Dim CurrentUnusedLambda As name
    For Each CurrentUnusedLambda In UnusedLambdas
        Logger.Log DEBUG_LOG, CurrentUnusedLambda.name & " is not used anywhere."
        CurrentUnusedLambda.Delete
    Next CurrentUnusedLambda
    Logger.Log TRACE_LOG, "Exit modUnUsedLambdas.RemoveUnusedLambdas"
    
End Sub

Private Function GetUnusedLambdas(ByVal AllLambdas As Collection _
                                  , ByVal AllUsedLambdas As Collection) As Collection
    
    Logger.Log TRACE_LOG, "Enter modUnUsedLambdas.GetUnusedLambdas"
    Dim UnusedLambdas As Collection
    Set UnusedLambdas = New Collection
    
    Dim CurrentLambda As name
    For Each CurrentLambda In AllLambdas
        If Not IsExistInCollection(AllUsedLambdas, CurrentLambda.name) Then
            UnusedLambdas.Add CurrentLambda, CurrentLambda.name
        End If
    Next CurrentLambda
    
    Set GetUnusedLambdas = UnusedLambdas
    Logger.Log TRACE_LOG, "Exit modUnUsedLambdas.GetUnusedLambdas"
    
End Function

Private Sub UpdateUsedLambdasForDependencies(ByRef AllUsedLambdas As Collection _
                                             , ByVal AllLambdas As Collection)
    
    Logger.Log TRACE_LOG, "Enter modUnUsedLambdas.UpdateUsedLambdasForDependencies"
    
    ' Look for dependency lambdas.
    Dim IsAtleastOneNewLambdaAdded As Boolean
    IsAtleastOneNewLambdaAdded = True
    
    Dim AlreadyScannedLambdas As Collection
    Set AlreadyScannedLambdas = New Collection
    
    Do While IsAtleastOneNewLambdaAdded
        
        IsAtleastOneNewLambdaAdded = False
        
        Dim CurrentUsedLambdaName As Variant
        For Each CurrentUsedLambdaName In AllUsedLambdas
            
            ' If already scanned then try next lambda name
            If IsExistInCollection(AlreadyScannedLambdas, CStr(CurrentUsedLambdaName)) Then GoTo NextUsedLambdaName
            
            Dim CurrentName As name
            Set CurrentName = AllLambdas.Item(CStr(CurrentUsedLambdaName))
            Dim UsedLambdasInCurrentFormula As Variant
            UsedLambdasInCurrentFormula = GetUsedLambdas(CurrentName.RefersTo, AllLambdas)
            
            AlreadyScannedLambdas.Add CurrentUsedLambdaName, CStr(CurrentUsedLambdaName)
            
            ' If no dependent lambdas then try next one.
            If Not IsArray(UsedLambdasInCurrentFormula) Then GoTo NextUsedLambdaName
            
            Dim CurrentLambda As Variant
            For Each CurrentLambda In UsedLambdasInCurrentFormula
                        
                If Not IsExistInCollection(AllUsedLambdas, CStr(CurrentLambda)) Then
                    Logger.Log DEBUG_LOG, CurrentUsedLambdaName & " is dependent on " & CurrentLambda
                    AllUsedLambdas.Add CurrentLambda, CStr(CurrentLambda)
                    IsAtleastOneNewLambdaAdded = True
                End If
                    
            Next CurrentLambda

NextUsedLambdaName:

        Next CurrentUsedLambdaName
        
    Loop
    Logger.Log TRACE_LOG, "Exit modUnUsedLambdas.UpdateUsedLambdasForDependencies"
    
End Sub

Private Function GetAllUsedLambdas(ByVal AllUniqueFormulas As Collection _
                                   , ByVal AllLambdas As Collection) As Collection
    
    Logger.Log TRACE_LOG, "Enter modUnUsedLambdas.GetAllUsedLambdas"
    Dim AllUsedLambdas As Collection
    Set AllUsedLambdas = New Collection
    Dim CurrentFormula As FormulaInfo
    
    For Each CurrentFormula In AllUniqueFormulas
        Dim UsedLambdasInCurrentFormula As Variant
        UsedLambdasInCurrentFormula = GetUsedLambdas(CurrentFormula.FormulaText, AllLambdas, CurrentFormula.IsR1C1)
        If IsArray(UsedLambdasInCurrentFormula) Then
            Dim CurrentLambda As Variant
            For Each CurrentLambda In UsedLambdasInCurrentFormula
                AddToCollectionIfNotExist AllUsedLambdas, CStr(CurrentLambda) _
                                                         , CStr(CurrentLambda)
            Next CurrentLambda
        End If
    Next CurrentFormula
    
    Set GetAllUsedLambdas = AllUsedLambdas
    Logger.Log TRACE_LOG, "Exit modUnUsedLambdas.GetAllUsedLambdas"
    
End Function

Private Sub KeepFormulasIfLambdaIsUsedByTextParsing(ByRef AllUniqueFormulas As Collection _
                                                    , ByVal AllLambdas As Collection)
    
    Logger.Log TRACE_LOG, "Enter modUnUsedLambdas.KeepFormulasIfLambdaIsUsedByTextParsing"
    Dim CurrentFormula As FormulaInfo
    For Each CurrentFormula In AllUniqueFormulas
        If Not IsAnyLambdaBeingUsed(CurrentFormula.FormulaText, AllLambdas) Then
            AllUniqueFormulas.Remove CurrentFormula.FormulaText
        End If
    Next CurrentFormula
    Logger.Log TRACE_LOG, "Exit modUnUsedLambdas.KeepFormulasIfLambdaIsUsedByTextParsing"
    
End Sub

Private Function IsAnyLambdaBeingUsed(ByVal Formula As String _
                                      , ByVal AllLambdas As Collection)
    
    Logger.Log TRACE_LOG, "Enter modUnUsedLambdas.IsAnyLambdaBeingUsed"
    Dim IsPresent As Boolean
    IsPresent = False
    Dim CurrentName As name
    For Each CurrentName In AllLambdas
        If Text.Contains(Formula, CurrentName.name) Then
            IsPresent = True
            Exit For
        End If
    Next CurrentName
    
    IsAnyLambdaBeingUsed = IsPresent
    Logger.Log TRACE_LOG, "Exit modUnUsedLambdas.IsAnyLambdaBeingUsed"
    
End Function

Private Sub UpdateFormulaCollFromCellFormulas(ByRef AllUniqueFormulas As Collection _
                                              , ByVal CurrentSheet As Worksheet)
    
    Logger.Log TRACE_LOG, "Enter modUnUsedLambdas.UpdateFormulaCollFromCellFormulas"
    Dim FormulaCells As Range
    Set FormulaCells = GetSpecialCells(CurrentSheet.UsedRange, xlCellTypeFormulas)
    
    If FormulaCells Is Nothing Then Exit Sub
    
    Dim CurrentCell As Range
    For Each CurrentCell In FormulaCells.Cells
        ' Using R1C1 as key is important because it will prevent adding similar formula.
        ' Using A1 (Local) important because we are using parser to parse A1C1 formula.
        With CurrentCell
            Dim CurrentFormulaInfo As FormulaInfo
            Set CurrentFormulaInfo = FormulaInfo.Create(.Formula2R1C1 _
                                                        , "Range(" & GetRangeRefWithSheetName(CurrentCell) & ").Formula2R1C1" _
                                                         , True)
            AddToCollectionIfNotExist AllUniqueFormulas, .Formula2R1C1, CurrentFormulaInfo
        End With
    Next CurrentCell
    Logger.Log TRACE_LOG, "Exit modUnUsedLambdas.UpdateFormulaCollFromCellFormulas"
    
End Sub

Private Sub UpdateFormulaCollFromCF(ByRef AllUniqueFormulas As Collection _
                                    , ByVal CurrentSheet As Worksheet)
    
    Logger.Log TRACE_LOG, "Enter modUnUsedLambdas.UpdateFormulaCollFromCF"
    '@INFO: I am being clever here. If we use SpecialCells on a one cell address then it expand.
    '       As I am providing only first cell it will expand and find all the format conditions cells.
    Dim CFCells As Range
    Set CFCells = GetSpecialCells(CurrentSheet.Cells(1), xlCellTypeAllFormatConditions)
        
    If CFCells Is Nothing Then Exit Sub
    
    Dim CurrentAreaRange As Range
    For Each CurrentAreaRange In CFCells.Areas
        Dim CurrentFormat As Object
        ' We have multiple options (Check all the Add* method name https://learn.microsoft.com/en-us/office/vba/api/excel.formatconditions.addaboveaverage)
        ' But some has only Formula and some no Formula property.
        ' https://learn.microsoft.com/en-us/office/vba/api/excel.xlformatconditiontype Here is the exhaustive list.
        
        For Each CurrentFormat In CurrentAreaRange.FormatConditions
            UpdateFormulaCollForCurrentCF AllUniqueFormulas, CurrentFormat
        Next CurrentFormat
        
    Next CurrentAreaRange
    Logger.Log TRACE_LOG, "Exit modUnUsedLambdas.UpdateFormulaCollFromCF"
    
End Sub

Private Sub UpdateFormulaCollForCurrentCF(ByRef AllUniqueFormulas As Collection _
                                          , ByVal CurrentFormat As Object)
     
    On Error Resume Next
    
    ' We may have object which has only Formula property, Some has Formula1 only and some both Formula1 and Formula2
    With CurrentFormat
        AddToCollectionIfNotExist AllUniqueFormulas, .Formula, FormulaInfo.Create(.Formula, "FormatCondition.Formula", False)
        AddToCollectionIfNotExist AllUniqueFormulas, .Formula1, FormulaInfo.Create(.Formula1, "FormatCondition.Formula1", False)
        AddToCollectionIfNotExist AllUniqueFormulas, .Formula2, FormulaInfo.Create(.Formula2, "FormatCondition.Formula2", False)
    End With
    
    On Error GoTo 0
            
End Sub

Private Function GetSpecialCells(ByVal FromRange As Range _
                                 , ByVal CellType As XlCellType) As Range
    
    Logger.Log TRACE_LOG, "Enter modUnUsedLambdas.GetSpecialCells"
    On Error Resume Next
    Set GetSpecialCells = FromRange.SpecialCells(CellType)
    On Error GoTo 0
    Logger.Log TRACE_LOG, "Exit modUnUsedLambdas.GetSpecialCells"
    
End Function


