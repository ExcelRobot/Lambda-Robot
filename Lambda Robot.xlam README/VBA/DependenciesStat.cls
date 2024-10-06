VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DependenciesStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "Lambda.Editor.DependencyInfo"
Option Explicit

Private Type TDependencyStats
    HasAnyCellLabelAsInputCell As Boolean
    MaximumLevel As Long
    MinimumLevel As Long
End Type

Private This As TDependencyStats

Public Property Get HasAnyCellLabelAsInputCell() As Boolean
    HasAnyCellLabelAsInputCell = This.HasAnyCellLabelAsInputCell
End Property

Public Property Get MaximumLevel() As Long
    MaximumLevel = This.MaximumLevel
End Property

Public Property Get MinimumLevel() As Long
    MinimumLevel = This.MinimumLevel
End Property

Public Sub UpdateLevel(ByVal NewLevel As Long)
    
    Logger.Log TRACE_LOG, "Enter DependenciesStat.UpdateLevel"
    If This.MaximumLevel < NewLevel Then
        This.MaximumLevel = NewLevel
    End If
    
    If This.MinimumLevel > NewLevel Then
        This.MinimumLevel = NewLevel
    End If
    Logger.Log TRACE_LOG, "Exit DependenciesStat.UpdateLevel"
    
End Sub

Public Sub UpdateInputCellStats(ByVal IsLabelAsInputCell As Boolean)
    
    Logger.Log TRACE_LOG, "Enter DependenciesStat.UpdateInputCellStats"
    If This.HasAnyCellLabelAsInputCell Then Exit Sub
    This.HasAnyCellLabelAsInputCell = IsLabelAsInputCell
    Logger.Log TRACE_LOG, "Exit DependenciesStat.UpdateInputCellStats"
    
End Sub

