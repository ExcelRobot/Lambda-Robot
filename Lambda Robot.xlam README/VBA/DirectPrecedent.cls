VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DirectPrecedent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit


Private Type TDirectPrecedent
    RangeRef As String
    NameInFormula As String
End Type

Private This As TDirectPrecedent

Public Property Get RangeRef() As String
    RangeRef = This.RangeRef
End Property

Public Property Let RangeRef(ByVal RHS As String)
    This.RangeRef = RHS
End Property

Public Property Get NameInFormula() As String
    NameInFormula = This.NameInFormula
End Property

Public Property Let NameInFormula(ByVal RHS As String)
    This.NameInFormula = RHS
End Property

Public Property Get Self() As DirectPrecedent
    Set Self = Me
End Property

Public Function Create(ByVal RangeRef As String _
                       , ByVal NameInFormula As String) As DirectPrecedent

    Logger.Log TRACE_LOG, "Enter DirectPrecedent.Create"
    With New DirectPrecedent
        .RangeRef = RangeRef
        .NameInFormula = NameInFormula
        Set Create = .Self
    End With
    Logger.Log TRACE_LOG, "Exit DirectPrecedent.Create"

End Function



