VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MetadataParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Lambda.Editor.Metadata"
'@IgnoreModule UndeclaredVariable, IndexedDefaultMemberAccess, UnrecognizedAnnotation, ProcedureNotUsed
'@PredeclaredId

Option Explicit

Private Type TMetadataParser
    ForCell As Range
    StartFormulaText As String
    IsCellHasLambdaFormula As Boolean
    IsNeedToUpdateNameManager As Boolean
    OldLambdaName As String
    AllNameInNameManager As Collection
    LambdaParts As Variant
    LetParts As Variant
End Type

Private This  As TMetadataParser

Public Property Get LetParts() As Variant
    Logger.Log TRACE_LOG, "Enter MetadataParser.LetParts"
    If IsObject(This.LetParts) Then
        Set LetParts = This.LetParts
    Else
        LetParts = This.LetParts
    End If
    Logger.Log TRACE_LOG, "Exit MetadataParser.LetParts"
End Property

Public Property Get ForCell() As Range
    Set ForCell = This.ForCell
End Property

Public Property Set ForCell(ByVal RHS As Range)
    Set This.ForCell = RHS
End Property

Public Property Get IsCellHasLambdaFormula() As Boolean
    IsCellHasLambdaFormula = This.IsCellHasLambdaFormula
End Property

Public Property Let IsCellHasLambdaFormula(ByVal RHS As Boolean)
    This.IsCellHasLambdaFormula = RHS
End Property

Public Property Get StartFormulaText() As String
    StartFormulaText = This.StartFormulaText
End Property

Public Property Let StartFormulaText(ByVal RHS As String)
    
    Logger.Log TRACE_LOG, "Enter MetadataParser.StartFormulaText"
    This.StartFormulaText = RHS
    With This
        .LetParts = GetDependencyFunctionResult(.StartFormulaText, LET_PARTS)
        ' Retrieve lambda parts for processing
        .LambdaParts = GetDependencyFunctionResult(.StartFormulaText, LAMBDA_PARTS)
    End With
    Logger.Log TRACE_LOG, "Exit MetadataParser.StartFormulaText"
    
End Property

Public Property Get IsNeedToUpdateNameManager() As Boolean
    IsNeedToUpdateNameManager = This.IsNeedToUpdateNameManager
End Property

Public Property Let IsNeedToUpdateNameManager(ByVal RHS As Boolean)
    This.IsNeedToUpdateNameManager = RHS
End Property

Public Property Get OldLambdaName() As String
    OldLambdaName = This.OldLambdaName
End Property

Public Property Let OldLambdaName(ByVal RHS As String)
    This.OldLambdaName = RHS
End Property

Public Property Get AllNameInNameManager() As Collection
    Set AllNameInNameManager = This.AllNameInNameManager
End Property

Public Property Set AllNameInNameManager(ByVal RHS As Collection)
    Set This.AllNameInNameManager = RHS
End Property

Public Property Get LambdaParts() As Variant
    Logger.Log TRACE_LOG, "Enter MetadataParser.LambdaParts"
    If IsObject(This.LambdaParts) Then
        Set LambdaParts = This.LambdaParts
    Else
        LambdaParts = This.LambdaParts
    End If
    Logger.Log TRACE_LOG, "Exit MetadataParser.LambdaParts"
End Property

Public Property Let LambdaParts(ByVal RHS As Variant)
    This.LambdaParts = RHS
End Property

Public Property Set LambdaParts(ByVal RHS As Variant)
    Set This.LambdaParts = RHS
End Property

Public Property Get Self() As MetadataParser
    Set Self = Me
End Property

'@EntryPoint
Public Function Create(ByVal LambdaFormulaCell As Range) As MetadataParser
    
    Logger.Log TRACE_LOG, "Enter MetadataParser.Create"

    Dim CurrentMetadataFormulaParser As MetadataParser
    Set CurrentMetadataFormulaParser = New MetadataParser

    With CurrentMetadataFormulaParser
        Set .ForCell = LambdaFormulaCell
        
        ' Filling the NameCollection with names from workbook
        FillNameCollection LambdaFormulaCell, CurrentMetadataFormulaParser
        
        ' Update old formula name if cell contains lambda and lambda is in NameManager
        UpdateInCellOrNamedRangeRelatedLambdaProperties LambdaFormulaCell, CurrentMetadataFormulaParser
        
        .StartFormulaText = GetFormulaTextFromCellOrNamedRange(LambdaFormulaCell, .IsNeedToUpdateNameManager)
        
        ' Returning the created object
        Set Create = CurrentMetadataFormulaParser
    End With
    
    Logger.Log TRACE_LOG, "Exit MetadataParser.Create"
    
End Function

Private Sub FillNameCollection(ByVal LambdaFormulaCell As Range _
                               , ByVal CurrentMetadataFormulaParser As MetadataParser)
    
    Logger.Log TRACE_LOG, "Enter MetadataParser.FillNameCollection"
    With CurrentMetadataFormulaParser
        Dim LambdaInBook As Workbook
        Set LambdaInBook = LambdaFormulaCell.Worksheet.Parent
        
        Dim CurrentName As name
        ' Looping through all Names in Workbook
        For Each CurrentName In LambdaInBook.Names
            
            If CurrentName.Visible Then
                .AllNameInNameManager.Add CurrentName, CStr(CurrentName.name)
            End If
        
        Next CurrentName

    End With
    Logger.Log TRACE_LOG, "Exit MetadataParser.FillNameCollection"
    
End Sub

Private Sub UpdateInCellOrNamedRangeRelatedLambdaProperties(ByVal ForCell As Range _
                                                            , ByVal CurrentMetadataFormulaParser As MetadataParser)
  
    Logger.Log TRACE_LOG, "Enter MetadataParser.UpdateInCellOrNamedRangeRelatedLambdaProperties"
    
    Dim FormulaText As String
    FormulaText = ForCell.Formula2
    
    ' Set IsCellHasLambdaFormula flag to True
    CurrentMetadataFormulaParser.IsCellHasLambdaFormula = True
    
    ' Check if FormulaText starts with lambda. If yes, no need to update NameManager
    If IsLambdaFunction(FormulaText) Then
        CurrentMetadataFormulaParser.IsNeedToUpdateNameManager = False
    ElseIf modUtility.IsCellHasSavedLambdaFormula(ForCell) Then
        ' If cell has only lambda formula, update NameManager and store OldLambdaName
        CurrentMetadataFormulaParser.IsNeedToUpdateNameManager = True
        CurrentMetadataFormulaParser.OldLambdaName = modUtility.ExtractStartFormulaName(ForCell.Formula2)
    Else
        CurrentMetadataFormulaParser.IsCellHasLambdaFormula = False
    End If
    
    Logger.Log TRACE_LOG, "Exit MetadataParser.UpdateInCellOrNamedRangeRelatedLambdaProperties"
  
End Sub

Private Function GetFormulaTextFromCellOrNamedRange(ByVal LambdaFormulaCell As Range _
                                                    , ByVal IsNeedToUpdateNameManager As Boolean) As Variant
    
    Logger.Log TRACE_LOG, "Enter MetadataParser.GetFormulaTextFromCellOrNamedRange"
    
    Dim LambdaInBook As Workbook
    Set LambdaInBook = LambdaFormulaCell.Worksheet.Parent
    
    Dim FormulaText As String
    ' Retrieve the formula text from either the NameManager or directly from the cell
    If IsNeedToUpdateNameManager Then
        Dim LambdaName As String
        LambdaName = modUtility.ExtractStartFormulaName(LambdaFormulaCell.Formula2)
        FormulaText = LambdaInBook.Names(LambdaName).RefersTo
    Else
        FormulaText = LambdaFormulaCell.Formula2
    End If
    
    GetFormulaTextFromCellOrNamedRange = FormulaText
    
    Logger.Log TRACE_LOG, "Exit MetadataParser.GetFormulaTextFromCellOrNamedRange"
    
End Function

Private Sub Class_Initialize()
    Logger.Log TRACE_LOG, "Enter MetadataParser.Class_Initialize"
    ' Initialization of Collection objects at class initialization
    Set This.AllNameInNameManager = New Collection
    Logger.Log TRACE_LOG, "Exit MetadataParser.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    Logger.Log TRACE_LOG, "Enter MetadataParser.Class_Terminate"
    ' Cleanup of Collection objects at class termination
    Set This.AllNameInNameManager = Nothing
    Logger.Log TRACE_LOG, "Exit MetadataParser.Class_Terminate"
End Sub


