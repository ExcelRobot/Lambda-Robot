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
    ForName As Name
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
    RangeLabelCell As Range
    Is3DRangeRef As Boolean
End Type
Private this As TDependencyInfo

Public Property Get Is3DRangeRef() As Boolean
    Is3DRangeRef = this.Is3DRangeRef
End Property

Public Property Let Is3DRangeRef(ByVal RHS As Boolean)
    this.Is3DRangeRef = RHS
End Property

Public Property Get RangeLabelCell() As Range
    Set RangeLabelCell = this.RangeLabelCell
End Property

Public Property Set RangeLabelCell(ByVal RHS As Range)
    Set this.RangeLabelCell = RHS
End Property

Public Property Get IsTableRef() As Boolean
    IsTableRef = this.IsTableRef
End Property

Public Property Let IsTableRef(ByVal RHS As Boolean)
    this.IsTableRef = RHS
End Property

Public Property Get SheetName() As String
    SheetName = this.SheetName
End Property

Public Property Let SheetName(ByVal RHS As String)
    this.SheetName = RHS
End Property

Public Property Get BookName() As String
    BookName = this.BookName
End Property

Public Property Let BookName(ByVal RHS As String)
    this.BookName = RHS
End Property

Public Property Get IsLocalScopeNamedRange() As Boolean
    IsLocalScopeNamedRange = this.IsLocalScopeNamedRange
End Property

Public Property Let IsLocalScopeNamedRange(RHS As Boolean)
    this.IsLocalScopeNamedRange = RHS
End Property

Public Property Get IsUserSpecifiedName() As Boolean
    IsUserSpecifiedName = this.IsUserSpecifiedName
End Property

Public Property Let IsUserSpecifiedName(ByVal RHS As Boolean)
    this.IsUserSpecifiedName = RHS
End Property

Public Property Get NewValidVarName() As String
    NewValidVarName = this.NewValidVarName
End Property

Public Property Let NewValidVarName(ByVal RHS As String)
    this.NewValidVarName = RHS
End Property

Public Property Get IsExpandByUser() As Boolean
    IsExpandByUser = this.IsExpandByUser
End Property

Public Property Let IsExpandByUser(ByVal RHS As Boolean)
    this.IsExpandByUser = RHS
End Property

Public Property Get IsUserMarkAsValue() As Boolean
    IsUserMarkAsValue = this.IsUserMarkAsValue
End Property

Public Property Let IsUserMarkAsValue(ByVal RHS As Boolean)
    this.IsUserMarkAsValue = RHS
End Property

Public Property Get IsUserMarkAsParameterCell() As Boolean
    IsUserMarkAsParameterCell = this.IsUserMarkAsParameterCell
End Property

Public Property Let IsUserMarkAsParameterCell(ByVal RHS As Boolean)
    this.IsUserMarkAsParameterCell = RHS
End Property

Public Property Get IsDemotedFromParameterCellToLetStep() As Boolean
    IsDemotedFromParameterCellToLetStep = this.IsDemotedFromParameterCellToLetStep
End Property

Public Property Let IsDemotedFromParameterCellToLetStep(RHS As Boolean)
    this.IsDemotedFromParameterCellToLetStep = RHS
End Property

Public Property Get IsMarkAsNotLetStatementByUser() As Boolean
    IsMarkAsNotLetStatementByUser = this.IsMarkAsNotLetStatementByUser
End Property

Public Property Let IsMarkAsNotLetStatementByUser(ByVal RHS As Boolean)
    this.IsMarkAsNotLetStatementByUser = RHS
End Property

Public Property Get ForName() As Name
    Set ForName = this.ForName
End Property

Public Property Set ForName(ByVal RHS As Name)
    
    Logger.Log TRACE_LOG, "Enter DependencyInfo.ForName"
    Set this.ForName = RHS
    If Not RHS Is Nothing Then
        this.IsLocalScopeNamedRange = modUtility.IsLocalScopeNamedRange(RHS.Name)
    End If
    Logger.Log TRACE_LOG, "Exit DependencyInfo.ForName"
    
End Property

Public Property Get IsReferByNamedRange() As Boolean
    IsReferByNamedRange = this.IsReferByNamedRange
End Property

Public Property Let IsReferByNamedRange(ByVal RHS As Boolean)
    this.IsReferByNamedRange = RHS
End Property

Public Property Get IsInsideNamedRangeOrTable() As Boolean
    IsInsideNamedRangeOrTable = this.IsInsideNamedRangeOrTable
End Property

Public Property Let IsInsideNamedRangeOrTable(RHS As Boolean)
    this.IsInsideNamedRangeOrTable = RHS
End Property

Public Property Get HasFormula() As Boolean
    HasFormula = this.HasFormula
End Property

Public Property Let HasFormula(ByVal RHS As Boolean)
    this.HasFormula = RHS
End Property

Public Property Get IsOptional() As Boolean
    IsOptional = this.IsOptional
End Property

Public Property Let IsOptional(ByVal RHS As Boolean)
    this.IsOptional = RHS
End Property

Public Property Get LetVarReference() As String
    LetVarReference = this.LetVarReference
End Property

Public Property Let LetVarReference(ByVal RHS As String)
    this.LetVarReference = RHS
End Property

Public Property Get RowNumber() As Long
    RowNumber = this.RowNumber
End Property

Public Property Let RowNumber(RHS As Long)
    this.RowNumber = RHS
End Property

Public Property Get NameInFormulaLength() As Long
    NameInFormulaLength = Len(this.NameInFormula)
End Property

Public Property Get ColumnNumber() As Long
    ColumnNumber = this.ColumnNumber
End Property

Public Property Let ColumnNumber(RHS As Long)
    this.ColumnNumber = RHS
End Property

Public Property Get ValidVarName() As String
    ValidVarName = this.ValidVarName
End Property

Public Property Let ValidVarName(RHS As String)
    this.ValidVarName = RHS
End Property

Public Property Get IsLabelAsInputCell() As Boolean
    IsLabelAsInputCell = this.IsLabelAsInputCell
End Property

Public Property Let IsLabelAsInputCell(ByVal RHS As Boolean)
    this.IsLabelAsInputCell = RHS
End Property

Public Property Get NameInFormula() As String
    NameInFormula = this.NameInFormula
End Property

Public Property Let NameInFormula(ByVal RHS As String)
    this.NameInFormula = RHS
End Property

Public Property Get HasAnyDependency() As Boolean
    HasAnyDependency = this.HasAnyDependency
End Property

Public Property Let HasAnyDependency(ByVal RHS As Boolean)
    this.HasAnyDependency = RHS
End Property

Public Property Get RangeLabel() As String
    RangeLabel = this.RangeLabel
End Property

Public Property Let RangeLabel(ByVal RHS As String)
    this.RangeLabel = RHS
End Property

Public Property Get RangeReference() As String
    RangeReference = this.RangeReference
End Property

Public Property Let RangeReference(ByVal RHS As String)
    this.RangeReference = RHS
End Property

Public Property Get Level() As Long
    Level = this.Level
End Property

Public Property Let Level(ByVal RHS As Long)
    If RHS > this.Level Then this.Level = RHS
End Property

Public Property Get FormulaText() As String
    
    Logger.Log TRACE_LOG, "Enter DependencyInfo.FormulaText"
    
    If Not this.HasFormula _
       And Not IsNumeric(this.FormulaText) _
       And this.FormulaText <> vbNullString _
       And Not this.IsReferByNamedRange Then
        FormulaText = Text.PadIfNotPresent(this.FormulaText, QUOTES, FROM_BOTH)
    Else
        FormulaText = this.FormulaText
    End If
    
    Logger.Log TRACE_LOG, "Exit DependencyInfo.FormulaText"
    
End Property

Public Property Let FormulaText(ByVal RHS As String)
    this.FormulaText = RHS
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
    ContextSheetName = ContextCell.Worksheet.Name
    
    Dim ContextBookName As String
    ContextBookName = ContextCell.Worksheet.Parent.Name
    
    Dim Result As String
    
    If Not this.IsLabelAsInputCell Then
        Result = vbNullString
    ElseIf this.IsUserMarkAsValue Then
        Result = Text.RemoveFromStartIfPresent(this.FormulaText, EQUAL_SIGN)
    ElseIf this.BookName <> ContextBookName Then
        Result = this.NameInFormula
    ElseIf this.IsReferByNamedRange Then
        
        If this.IsLocalScopeNamedRange Then
            If this.SheetName <> ContextSheetName Then
                Result = GetSheetRefForRangeReference(this.SheetName) _
                         & RemoveSheetQualifierIfPresent(this.NameInFormula)
            Else
                Result = RemoveSheetQualifierIfPresent(this.NameInFormula)
            End If
        Else
            Result = this.NameInFormula
        End If
        
    ElseIf IsTableRefUsedInFormula(this.NameInFormula) Then
        Result = this.NameInFormula
    ElseIf this.SheetName <> ContextSheetName Then
        Result = GetSheetRefForRangeReference(this.SheetName) _
                 & RemoveSheetQualifierIfPresent(this.NameInFormula)
    ElseIf this.SheetName = ContextSheetName Then
        Result = RemoveSheetQualifierIfPresent(this.NameInFormula)
    Else
        Result = this.NameInFormula
    End If
    
    StepOrArgumentCalcFunction = Result
    Logger.Log TRACE_LOG, "Exit DependencyInfo.StepOrArgumentCalcFunction"
        
End Function

Public Sub UpdateNameInFormula(ByVal UpdatedName As String)
    
    Logger.Log TRACE_LOG, "Enter DependencyInfo.UpdateNameInFormula"
    ' Updates the "NameInFormula" property with the given "UpdatedName."
    
    If this.NameInFormula = vbNullString Then
        this.NameInFormula = UpdatedName
    Else
        If RemoveDollarSign(this.NameInFormula) = RemoveDollarSign(UpdatedName) Then
            If Text.Contains(UpdatedName, DOLLAR_SIGN) Then
                this.NameInFormula = UpdatedName
            End If
        Else
            this.NameInFormula = UpdatedName
        End If
    End If
    Logger.Log TRACE_LOG, "Exit DependencyInfo.UpdateNameInFormula"
    
End Sub

Public Function CreateMe() As DependencyInfo
    ' Creates and returns a new instance of the DependencyInfo class.
    Set CreateMe = New DependencyInfo
End Function


