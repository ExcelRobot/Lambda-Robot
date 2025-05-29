VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LetOrLambdaInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, ProcedureNotUsed
'@Folder "Lambda.Editor.DependencyInfo"
Option Explicit

Private Type TLetOrLambdaInfo
    LetFormula As String
    LambdaFormula As String
    AFEScriptCode As Variant
    InvocationArgument As String
    PutDependencyOnTable As ListObject
    UserMarkedInputCells As Collection
    RangeReqFXList As Collection
End Type

Private this  As TLetOrLambdaInfo

Public Property Get RangeReqFXList() As Collection
    Set RangeReqFXList = this.RangeReqFXList
End Property

Public Property Set RangeReqFXList(ByVal RHS As Collection)
    Set this.RangeReqFXList = RHS
End Property

Public Property Get UserMarkedInputCells() As Collection
    Set UserMarkedInputCells = this.UserMarkedInputCells
End Property

Public Property Set UserMarkedInputCells(ByVal RHS As Collection)
    Set this.UserMarkedInputCells = RHS
End Property

Public Property Get PutDependencyOnTable() As ListObject
    Set PutDependencyOnTable = this.PutDependencyOnTable
End Property

Public Property Set PutDependencyOnTable(ByVal RHS As ListObject)
    Set this.PutDependencyOnTable = RHS
End Property

Public Property Get InvocationArgument() As String
    InvocationArgument = this.InvocationArgument
End Property

Public Property Let InvocationArgument(ByVal RHS As String)
    this.InvocationArgument = RHS
End Property

Public Property Get AFEScriptCode() As Variant
    If IsObject(this.AFEScriptCode) Then
        Set AFEScriptCode = this.AFEScriptCode
    Else
        AFEScriptCode = this.AFEScriptCode
    End If
End Property

Public Property Let AFEScriptCode(ByVal RHS As Variant)
    this.AFEScriptCode = RHS
End Property

Public Property Set AFEScriptCode(ByVal RHS As Variant)
    Set this.AFEScriptCode = RHS
End Property

Public Property Get LetFormula() As String
    LetFormula = this.LetFormula
End Property

Public Property Let LetFormula(ByVal RHS As String)
    this.LetFormula = RHS
End Property

Public Property Get LambdaFormula() As String
    LambdaFormula = this.LambdaFormula
End Property

Public Property Let LambdaFormula(ByVal RHS As String)
    this.LambdaFormula = RHS
End Property

