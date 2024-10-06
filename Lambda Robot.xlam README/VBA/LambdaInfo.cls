VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LambdaInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, ProcedureNotUsed
'@Folder "Lambda.Editor.DependencyInfo"
Option Explicit

Private Type TLambdaInfo
    LetFormula As String
    LambdaFormula As String
    AFEScriptCode As Variant
    InvocationArgument As String
    PutDependencyOnTable As ListObject
End Type

Private This  As TLambdaInfo

Public Property Get PutDependencyOnTable() As ListObject
    Set PutDependencyOnTable = This.PutDependencyOnTable
End Property

Public Property Set PutDependencyOnTable(ByVal RHS As ListObject)
    Set This.PutDependencyOnTable = RHS
End Property

Public Property Get InvocationArgument() As String
    InvocationArgument = This.InvocationArgument
End Property

Public Property Let InvocationArgument(ByVal RHS As String)
    This.InvocationArgument = RHS
End Property

Public Property Get AFEScriptCode() As Variant
    If IsObject(This.AFEScriptCode) Then
        Set AFEScriptCode = This.AFEScriptCode
    Else
        AFEScriptCode = This.AFEScriptCode
    End If
End Property

Public Property Let AFEScriptCode(ByVal RHS As Variant)
    This.AFEScriptCode = RHS
End Property

Public Property Set AFEScriptCode(ByVal RHS As Variant)
    Set This.AFEScriptCode = RHS
End Property

Public Property Get LetFormula() As String
    LetFormula = This.LetFormula
End Property

Public Property Let LetFormula(ByVal RHS As String)
    This.LetFormula = RHS
End Property

Public Property Get LambdaFormula() As String
    LambdaFormula = This.LambdaFormula
End Property

Public Property Let LambdaFormula(ByVal RHS As String)
    This.LambdaFormula = RHS
End Property

