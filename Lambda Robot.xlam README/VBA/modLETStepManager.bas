Attribute VB_Name = "modLETStepManager"
'@IgnoreModule UndeclaredVariable
'@Folder "Step.Manager"
Option Explicit

Private Sub Test()
    AddLetStep ActiveCell, "Sum By Row", "=BYROW([[PreviousStep]],LAMBDA(x,SUM(x)))"
End Sub

Public Sub AddLetStep(ByVal FormulaCell As Range, ByVal StepName As String _
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
Public Sub CycleLETSteps(ByVal FormulaCell As Range, Optional ByVal TargetCell As Range = Nothing _
                                                    , Optional ByVal IsReset As Boolean = False)
    
    LetStepManager.CycleLETSteps FormulaCell, TargetCell, IsReset
    
End Sub



