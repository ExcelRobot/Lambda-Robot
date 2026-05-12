VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Private Type TFormulaInfo
    FormulaText As String
    A1FormulaText As String
    Source As String
    IsR1C1 As Boolean
End Type

Private this As TFormulaInfo

Public Property Get FormulaText() As String
    FormulaText = this.FormulaText
End Property

Public Property Let FormulaText(ByVal RHS As String)
    this.FormulaText = RHS
End Property

Public Property Get Source() As String
    Source = this.Source
End Property

Public Property Let Source(ByVal RHS As String)
    this.Source = RHS
End Property

Public Property Get IsR1C1() As Boolean
    IsR1C1 = this.IsR1C1
End Property

Public Property Let IsR1C1(ByVal RHS As Boolean)
    this.IsR1C1 = RHS
End Property

Public Property Get A1FormulaText() As String
    A1FormulaText = this.A1FormulaText
End Property

Public Property Let A1FormulaText(ByVal RHS As String)
    this.A1FormulaText = RHS
End Property

Public Property Get Self() As FormulaInfo
    Set Self = Me
End Property

' A1FormulaText is optional; defaults to FormulaText when omitted (correct for IsR1C1=False callers).
Public Function Create(ByVal FormulaText As String _
                       , ByVal Source As String _
                       , ByVal IsR1C1 As Boolean _
                       , Optional ByVal A1FormulaText As String = vbNullString) As FormulaInfo

    With New FormulaInfo
        .FormulaText = FormulaText
        .Source = Source
        .IsR1C1 = IsR1C1
        .A1FormulaText = IIf(A1FormulaText = vbNullString, FormulaText, A1FormulaText)
        Set Create = .Self
    End With

End Function



