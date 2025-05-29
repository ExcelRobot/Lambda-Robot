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

Public Property Get Self() As FormulaInfo
    Set Self = Me
End Property

Public Function Create(ByVal FormulaText As String, ByVal Source As String, ByVal IsR1C1 As Boolean) As FormulaInfo

    With New FormulaInfo
        .FormulaText = FormulaText
        .Source = Source
        .IsR1C1 = IsR1C1
        Set Create = .Self
    End With

End Function

