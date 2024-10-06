VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChartInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TChartInfo
    OnSheet As String
    ChartObjectName As String
    Formula As String
    VBASyntax As String
    IsWholeSearchRangeBeingUsed As Boolean
    RangeAddress As String
    PropertyName As String
    FoundOnObject As Object
End Type

Private This As TChartInfo

Public Property Get FoundOnObject() As Object
    Set FoundOnObject = This.FoundOnObject
End Property

Public Property Set FoundOnObject(ByVal RHS As Object)
    Set This.FoundOnObject = RHS
End Property

Public Property Get PropertyName() As String
    PropertyName = This.PropertyName
End Property

Public Property Let PropertyName(ByVal RHS As String)
    This.PropertyName = RHS
End Property

Public Property Get RangeAddress() As String
    RangeAddress = This.RangeAddress
End Property

Public Property Let RangeAddress(ByVal RHS As String)
    This.RangeAddress = RHS
End Property

Public Property Get IsWholeSearchRangeBeingUsed() As Boolean
    IsWholeSearchRangeBeingUsed = This.IsWholeSearchRangeBeingUsed
End Property

Public Property Let IsWholeSearchRangeBeingUsed(ByVal RHS As Boolean)
    This.IsWholeSearchRangeBeingUsed = RHS
End Property

Public Property Get VBASyntax() As String
    VBASyntax = This.VBASyntax
End Property

Public Property Let VBASyntax(ByVal RHS As String)
    This.VBASyntax = RHS
End Property

Public Property Get OnSheet() As String
    OnSheet = This.OnSheet
End Property

Public Property Let OnSheet(ByVal RHS As String)
    This.OnSheet = RHS
End Property

Public Property Get ChartObjectName() As String
    ChartObjectName = This.ChartObjectName
End Property

Public Property Let ChartObjectName(ByVal RHS As String)
    This.ChartObjectName = RHS
End Property

Public Property Get Formula() As String
    Formula = This.Formula
End Property

Public Property Let Formula(ByVal RHS As String)
    This.Formula = RHS
End Property

