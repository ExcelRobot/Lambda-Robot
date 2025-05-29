Attribute VB_Name = "modBOMode"
Option Explicit


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Enable Bo Mode
' Description:            Ensures generated LET and LAMBDA statements are as short as possible.
' Macro Expression:       modBOMode.EnableBoMode()
' Generated:              11/19/2024 10:44 AM
'----------------------------------------------------------------------------------------------------
Public Sub EnableBoMode()
    SetDefaultParameterValueByName "FormulaFormat_BoMode", "true"
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Disable Bo Mode
' Description:            Restores standard formatting rules for generated LET and LAMBDA statements.
' Macro Expression:       modBOMode.DisableBoMode()
' Generated:              11/19/2024 10:50 AM
'----------------------------------------------------------------------------------------------------
Public Sub DisableBoMode()
    SetDefaultParameterValueByName "FormulaFormat_BoMode", "false"
End Sub

Private Sub SetDefaultParameterValueByName(ByVal ParameterName As String, ByVal ParameterValue As String)
 
    Dim oXLL As Object
    Set oXLL = CreateObject("OARobot.ExcelAddin")
    oXLL.SetDefaultParamValueByName ParameterName, ParameterValue
    Set oXLL = Nothing
 
End Sub
