VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DependencyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, ProcedureNotUsed
'@Folder "Lambda.Editor.DependencyInfo"
Option Explicit

Private Type TDependencyInfo
    RangeLabel As String
    RangeReference As String
    Level As Long
    FormulaText As String
    HasAnyDependency As Boolean
    NameInFormula As String
    IsLabelAsInputCell As Boolean
    ValidVarName As String
    NewValidVarName As String
    ColumnNumber As Long
    RowNumber As Long
    LetVarReference As String
    IsOptional As Boolean
    HasFormula As Boolean
    IsInsideNamedRangeOrTable As Boolean
    IsReferByNamedRange As Boolean
    ForName As name
    IsMarkAsNotLetStatementByUser As Boolean
    IsDemotedFromParameterCellToLetStep As Boolean
    IsUserMarkAsParameterCell As Boolean
    IsUserMarkAsValue As Boolean
    IsExpandByUser As Boolean
    IsUserSpecifiedName As Boolean
    IsLocalScopeNamedRange As Boolean
    SheetName As String
    BookName As String
    IsTableRef As Boolean
End Type
Private This As TDependencyInfo

Public Property Get IsTableRef() As Boolean
    IsTableRef = This.IsTableRef
End Property

Public Property Let IsTableRef(ByVal RHS As Boolean)
    This.IsTableRef = RHS
End Property

Public Property Get SheetName() As String
    SheetName = This.SheetName
End Property

Public Property Let SheetName(ByVal RHS As String)
    This.SheetName = RHS
End Property

Public Property Get BookName() As String
    BookName = This.BookName
End Property

Public Property Let BookName(ByVal RHS As String)
    This.BookName = RHS
End Property

Public Property Get IsLocalScopeNamedRange() As Boolean
    IsLocalScopeNamedRange = This.IsLocalScopeNamedRange
End Property

Public Property Let IsLocalScopeNamedRange(RHS As Boolean)
    This.IsLocalScopeNamedRange = RHS
End Property

Public Property Get IsUserSpecifiedName() As Boolean
    IsUserSpecifiedName = This.IsUserSpecifiedName
End Property

Public Property Let IsUserSpecifiedName(ByVal RHS As Boolean)
    This.IsUserSpecifiedName = RHS
End Property

Public Property Get NewValidVarName() As String
    NewValidVarName = This.NewValidVarName
End Property

Public Property Let NewValidVarName(ByVal RHS As String)
    This.NewValidVarName = RHS
End Property

Public Property Get IsExpandByUser() As Boolean
    IsExpandByUser = This.IsExpandByUser
End Property

Public Property Let IsExpandByUser(ByVal RHS As Boolean)
    This.IsExpandByUser = RHS
End Property

Public Property Get IsUserMarkAsValue() As Boolean
    IsUserMarkAsValue = This.IsUserMarkAsValue
End Property

Public Property Let IsUserMarkAsValue(ByVal RHS As Boolean)
    This.IsUserMarkAsValue = RHS
End Property

Public Property Get IsUserMarkAsParameterCell() As Boolean
    IsUserMarkAsParameterCell = This.IsUserMarkAsParameterCell
End Property

Public Property Let IsUserMarkAsParameterCell(ByVal RHS As Boolean)
    This.IsUserMarkAsParameterCell = RHS
End Property

Public Property Get IsDemotedFromParameterCellToLetStep() As Boolean
    IsDemotedFromParameterCellToLetStep = This.IsDemotedFromParameterCellToLetStep
End Property

Public Property Let IsDemotedFromParameterCellToLetStep(RHS As Boolean)
    This.IsDemotedFromParameterCellToLetStep = RHS
End Property

Public Property Get IsMarkAsNotLetStatementByUser() As Boolean
    IsMarkAsNotLetStatementByUser = This.IsMarkAsNotLetStatementByUser
End Property

Public Property Let IsMarkAsNotLetStatementByUser(ByVal RHS As Boolean)
    This.IsMarkAsNotLetStatementByUser = RHS
End Property

Public Property Get ForName() As name
    Set ForName = This.ForName
End Property

Public Property Set ForName(ByVal RHS As name)
    
    Logger.Log TRACE_LOG, "Enter DependencyInfo.ForName"
    Set This.ForName = RHS
    If Not RHS Is Nothing Then
        This.IsLocalScopeNamedRange = modUtility.IsLocalScopeNamedRange(RHS.name)
    End If
    Logger.Log TRACE_LOG, "Exit DependencyInfo.ForName"
    
End Property

Public Property Get IsReferByNamedRange() As Boolean
    IsReferByNamedRange = This.IsReferByNamedRange
End Property

Public Property Let IsReferByNamedRange(ByVal RHS As Boolean)
    This.IsReferByNamedRange = RHS
End Property

Public Property Get IsInsideNamedRangeOrTable() As Boolean
    IsInsideNamedRangeOrTable = This.IsInsideNamedRangeOrTable
End Property

Public Property Let IsInsideNamedRangeOrTable(RHS As Boolean)
    This.IsInsideNamedRangeOrTable = RHS
End Property

Public Property Get HasFormula() As Boolean
    HasFormula = This.HasFormula
End Property

Public Property Let HasFormula(ByVal RHS As Boolean)
    This.HasFormula = RHS
End Property

Public Property Get IsOptional() As Boolean
    IsOptional = This.IsOptional
End Property

Public Property Let IsOptional(ByVal RHS As Boolean)
    This.IsOptional = RHS
End Property

Public Property Get LetVarReference() As String
    LetVarReference = This.LetVarReference
End Property

Public Property Let LetVarReference(ByVal RHS As String)
    This.LetVarReference = RHS
End Property

Public Property Get RowNumber() As Long
    RowNumber = This.RowNumber
End Property

Public Property Let RowNumber(RHS As Long)
    This.RowNumber = RHS
End Property

Public Property Get NameInFormulaLength() As Long
    NameInFormulaLength = Len(This.NameInFormula)
End Property

Public Property Get ColumnNumber() As Long
    ColumnNumber = This.ColumnNumber
End Property

Public Property Let ColumnNumber(RHS As Long)
    This.ColumnNumber = RHS
End Property

Public Property Get ValidVarName() As String
    ValidVarName = This.ValidVarName
End Property

Public Property Let ValidVarName(RHS As String)
    This.ValidVarName = RHS
End Property

Public Property Get IsLabelAsInputCell() As Boolean
    IsLabelAsInputCell = This.IsLabelAsInputCell
End Property

Public Property Let IsLabelAsInputCell(ByVal RHS As Boolean)
    This.IsLabelAsInputCell = RHS
End Property

Public Property Get NameInFormula() As String
    NameInFormula = This.NameInFormula
End Property

Public Property Let NameInFormula(ByVal RHS As String)
    This.NameInFormula = RHS
End Property

Public Property Get HasAnyDependency() As Boolean
    HasAnyDependency = This.HasAnyDependency
End Property

Public Property Let HasAnyDependency(ByVal RHS As Boolean)
    This.HasAnyDependency = RHS
End Property

Public Property Get RangeLabel() As String
    RangeLabel = This.RangeLabel
End Property

Public Property Let RangeLabel(ByVal RHS As String)
    This.RangeLabel = RHS
End Property

Public Property Get RangeReference() As String
    RangeReference = This.RangeReference
End Property

Public Property Let RangeReference(ByVal RHS As String)
    This.RangeReference = RHS
End Property

Public Property Get Level() As Long
    Level = This.Level
End Property

Public Property Let Level(ByVal RHS As Long)
    If RHS > This.Level Then This.Level = RHS
End Property

Public Property Get FormulaText() As String
    
    Logger.Log TRACE_LOG, "Enter DependencyInfo.FormulaText"
    If Not This.HasFormula And Not IsNumeric(This.FormulaText) And This.FormulaText <> vbNullString Then
        FormulaText = Text.PadIfNotPresent(This.FormulaText, QUOTES, FROM_BOTH)
    Else
        FormulaText = This.FormulaText
    End If
    Logger.Log TRACE_LOG, "Exit DependencyInfo.FormulaText"
    
End Property

Public Property Let FormulaText(ByVal RHS As String)
    This.FormulaText = RHS
End Property


Private Function IsTableRefUsedInFormula(ByVal NameInFormula As String) As Boolean
    
    Logger.Log TRACE_LOG, "Enter DependencyInfo.IsTableRefUsedInFormula"
    Dim Result As Boolean
    Result = False
    
    If Text.IsEndsWith(NameInFormula, RIGHT_BRACKET) Then
        If Text.IsStartsWith(NameInFormula, LEFT_BRACKET & "@") Then
            Result = True
        ElseIf Text.Contains(NameInFormula, LEFT_BRACKET) Then
            Result = True
        End If
    End If
    
    IsTableRefUsedInFormula = Result
    Logger.Log TRACE_LOG, "Exit DependencyInfo.IsTableRefUsedInFormula"
    
End Function

Public Function StepOrArgumentCalcFunction(ByVal ContextCell As Range) As String
    
    Logger.Log TRACE_LOG, "Enter DependencyInfo.StepOrArgumentCalcFunction"
    ' This will find the correct calculation step with respect to ContextCell.
    ' As we are not finding any precedency when the reference is from any
    ' other workbook than Context Workbook it will return the NameInFormula if other
    ' workbook.
    
    Dim ContextSheetName As String
    ContextSheetName = ContextCell.Worksheet.name
    
    Dim ContextBookName As String
    ContextBookName = ContextCell.Worksheet.Parent.name
    
    Dim Result As String
    
    If Not This.IsLabelAsInputCell Then
        Result = vbNullString
    ElseIf This.BookName <> ContextBookName Then
        Result = This.NameInFormula
    ElseIf This.IsReferByNamedRange Then
        
        If This.IsLocalScopeNamedRange Then
            If This.SheetName <> ContextSheetName Then
                Result = GetSheetRefForRangeReference(This.SheetName) _
                       & RemoveSheetQualifierIfPresent(This.NameInFormula)
            Else
                Result = RemoveSheetQualifierIfPresent(This.NameInFormula)
            End If
        Else
            Result = This.NameInFormula
        End If
        
    ElseIf IsTableRefUsedInFormula(This.NameInFormula) Then
        Result = This.NameInFormula
    ElseIf This.SheetName <> ContextSheetName Then
        Result = GetSheetRefForRangeReference(This.SheetName) _
               & RemoveSheetQualifierIfPresent(This.NameInFormula)
    ElseIf This.SheetName = ContextSheetName Then
        Result = RemoveSheetQualifierIfPresent(This.NameInFormula)
    Else
        Result = This.NameInFormula
    End If
    
    StepOrArgumentCalcFunction = Result
    Logger.Log TRACE_LOG, "Exit DependencyInfo.StepOrArgumentCalcFunction"
        
End Function

Public Function ToJSON() As String
    
    Logger.Log TRACE_LOG, "Enter DependencyInfo.ToJSON"
    ' This function converts the current DependencyInfo object to a JSON-formatted string.
    
    Dim Map As New Scripting.Dictionary
    With Map
        .Add "Level", This.Level
        .Add "FormulaText", This.FormulaText
        .Add "RangeReference", This.RangeReference
        .Add "RangeLabel", This.RangeLabel
        .Add "NameInFormula", This.NameInFormula
    End With
    
    Const SPACE_COUNT As Long = 2
    ToJSON = JsonConverter.ConvertToJson(Map, SPACE_COUNT)
    Logger.Log TRACE_LOG, "Exit DependencyInfo.ToJSON"
    
End Function

Public Sub UpdateNameInFormula(ByVal UpdatedName As String)
    
    Logger.Log TRACE_LOG, "Enter DependencyInfo.UpdateNameInFormula"
    ' Updates the "NameInFormula" property with the given "UpdatedName."
    
    If This.NameInFormula = vbNullString Then
        This.NameInFormula = UpdatedName
    Else
        If RemoveDollarSign(This.NameInFormula) = RemoveDollarSign(UpdatedName) Then
            If Text.Contains(UpdatedName, DOLLAR_SIGN) Then
                This.NameInFormula = UpdatedName
            End If
        Else
            This.NameInFormula = UpdatedName
        End If
    End If
    Logger.Log TRACE_LOG, "Exit DependencyInfo.UpdateNameInFormula"
    
End Sub

Public Function CreateMe() As DependencyInfo
    ' Creates and returns a new instance of the DependencyInfo class.
    Set CreateMe = New DependencyInfo
End Function


