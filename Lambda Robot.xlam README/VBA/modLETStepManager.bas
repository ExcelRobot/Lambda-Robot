Attribute VB_Name = "modLETStepManager"
'@IgnoreModule UndeclaredVariable
'@Folder "Step.Manager"
Option Explicit
Option Private Module

Private Sub Test()
    AddLetStep ActiveCell, "Sum By Row", "=BYROW([[PreviousStep]],LAMBDA(x,SUM(x)))"
End Sub

Public Sub AddLetStep(ByVal FormulaCell As Range _
                      , ByVal StepName As String _
                       , ByVal StepFormula As String _
                        , Optional ByVal TargetCell As Range = Nothing)
    LetStepManager.AddLetStep FormulaCell, StepName, StepFormula, TargetCell
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Remove Last LET Step
' Description:            Remove last LET step.
' Macro Expression:       modLETStepManager.RemoveLastLETStep([ActiveCell])
' Generated:              04/08/2023 05:50 PM
'----------------------------------------------------------------------------------------------------
Public Sub RemoveLastLETStep(ByVal FormulaCell As Range, Optional ByVal TargetCell As Range = Nothing)
    LetStepManager.RemoveLastLETStep FormulaCell, TargetCell
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Cycle LET Steps
' Description:            Cyclically change last steps of the let so that we can see different steps calculated value.
' Macro Expression:       modLETStepManager.CycleLETSteps([ActiveCell],[ActiveCell])
' Generated:              08/07/2023 08:09 PM
'----------------------------------------------------------------------------------------------------
Public Sub CycleLETSteps(ByVal FormulaCell As Range _
                         , Optional ByVal TargetCell As Range = Nothing _
                          , Optional ByVal IsReset As Boolean = False)
    
    Const METHOD_NAME As String = "CycleLETSteps"
    Context.ExtractContextFromCell FormulaCell, METHOD_NAME
    If IsCellHasSavedLambdaFormula(FormulaCell) Then
        EditLambda FormulaCell
    End If
    
    LetStepManager.CycleLETSteps FormulaCell, TargetCell, IsReset
    Context.ClearContext METHOD_NAME
    
End Sub

Public Sub DebugLETSteps(ByVal FormulaCell As Range _
                         , Optional ByVal Spaced As Boolean = False _
                          , Optional ByVal IsUndo As Boolean = False)
    
    Static PutFormulaOnUndo As Range
    Static Formula As String
    Static IsDeleteComment As Boolean
    
    If IsUndo Then
        If IsNotNothing(PutFormulaOnUndo) Then PutFormulaOnUndo.Formula2 = Formula
        DeleteComment PutFormulaOnUndo
        Exit Sub
    Else
        ' If not undo, store the formula range for future use
        Set PutFormulaOnUndo = FormulaCell
        Formula = FormulaCell.Formula2
    End If
    
    Const METHOD_NAME As String = "DebugLETSteps"
    Context.ExtractContextFromCell FormulaCell, METHOD_NAME
    
    On Error GoTo HandleError
    If IsCellHasSavedLambdaFormula(FormulaCell) Then
        EditLambda FormulaCell
        IsDeleteComment = True
    Else
        IsDeleteComment = False
    End If
    
    LetStepManager.DebugLETSteps FormulaCell, Spaced
    
    If Not IsUndo Then AssingOnUndo "DebugLETSteps"
    
HandleError:
    Context.ClearContext METHOD_NAME
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbOKOnly + vbExclamation, "Stack LET Variables"
    End If
    
    
End Sub

Public Sub DebugLETSteps_Undo()
    DebugLETSteps Nothing, False, True
End Sub

