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

Private this As TDependencyStats

Public Property Get HasAnyCellLabelAsInputCell() As Boolean
    HasAnyCellLabelAsInputCell = this.HasAnyCellLabelAsInputCell
End Property

Public Property Get MaximumLevel() As Long
    MaximumLevel = this.MaximumLevel
End Property

Public Property Get MinimumLevel() As Long
    MinimumLevel = this.MinimumLevel
End Property

Public Sub UpdateLevel(ByVal NewLevel As Long)
    
    Logger.Log TRACE_LOG, "Enter DependenciesStat.UpdateLevel"
    If this.MaximumLevel < NewLevel Then
        this.MaximumLevel = NewLevel
    End If
    
    If this.MinimumLevel > NewLevel Then
        this.MinimumLevel = NewLevel
    End If
    Logger.Log TRACE_LOG, "Exit DependenciesStat.UpdateLevel"
    
End Sub

Public Sub UpdateInputCellStats(ByVal IsLabelAsInputCell As Boolean)
    
    Logger.Log TRACE_LOG, "Enter DependenciesStat.UpdateInputCellStats"
    If this.HasAnyCellLabelAsInputCell Then Exit Sub
    this.HasAnyCellLabelAsInputCell = IsLabelAsInputCell
    Logger.Log TRACE_LOG, "Exit DependenciesStat.UpdateInputCellStats"
    
End Sub

