VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Presenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed, ConstantNotUsed, EmptyMethod, UnrecognizedAnnotation, ProcedureNotUsed
'@Folder "Lambda.Editor.Metadata.Presenter.Impl"
Option Explicit

Implements IPresenter

Private Type TPresenter
    LambdaMetadata As Metadata
    LambdaFormulaCell As Range
    TokenizedFormula As Variant
    View As LambdaMetadataView
    MetaParser As MetadataParser
    DefaultName As String
    IsCallFromSaveAs As Boolean
    IsProcessCancelled As Boolean
    GistText As String
End Type

Private This  As TPresenter

Public Property Get GistText() As String
    GistText = This.GistText
End Property

Public Property Get LambdaMetadata() As Metadata
    Set LambdaMetadata = This.LambdaMetadata
End Property

Public Property Get IsCallFromSaveAs() As Boolean
    IsCallFromSaveAs = This.IsCallFromSaveAs
End Property

Public Property Get IsProcessCancelled() As Boolean
    IsProcessCancelled = This.IsProcessCancelled
End Property

'@EntryPoint
Public Sub ShowView(ByVal LambdaFormulaCell As Range, ByVal IsCallFromSaveAs As Boolean _
                                                     , ByVal DefaultName As String _
                                                      , Optional ByVal IsShowUI As Boolean = True)
    
    Logger.Log TRACE_LOG, "Enter Presenter.ShowView"
    ' Create MetaParser and check if the cell contains a LAMBDA formula
    Set This.MetaParser = MetadataParser.Create(LambdaFormulaCell)
    If Not This.MetaParser.IsCellHasLambdaFormula Then
        MsgBox "Cell doesn't contain a LAMBDA formula.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Presenter.ShowView"
        Exit Sub
    End If
    
    ' Initialize properties and set LambdaMetadata
    Set This.LambdaFormulaCell = LambdaFormulaCell
    This.DefaultName = DefaultName
    Set This.LambdaMetadata = Metadata.CreateLambdaMetadata(This.MetaParser.LambdaParts _
                                                            , This.MetaParser.LetParts _
                                                             , This.DefaultName _
                                                              , This.MetaParser.OldLambdaName)
    This.IsCallFromSaveAs = IsCallFromSaveAs
    
    ' Check if any metadata is found and if not, exit the sub (unless it's called from Save As)
    If Not IsAnyMetadataFound(This.LambdaMetadata) And Not IsCallFromSaveAs Then Exit Sub
    
    ' Show the LambdaMetadataView (if IsShowUI is True)
    If IsShowUI Then
        Set This.View = LambdaMetadataView.Create(This.LambdaMetadata, Me)
        This.View.UpdateView
        This.View.ExpandOrCollapse IsPropertyBelongToMoreInTheUIPresent(LambdaMetadata)
        
        If IsCallFromSaveAs Then UpdateUIForSaveAs This.View
        
        If This.View.NameTextBox.Enabled Then
            This.View.NameTextBox.SetFocus
        Else
            This.View.CommandNameTextBox.SetFocus
        End If
        
        This.View.Show
        On Error Resume Next
        Unload This.View
        Set This.View = Nothing
        On Error GoTo 0
    End If
    
    ' Log function exit
    Logger.Log TRACE_LOG, "Exit Presenter.ShowView"
    
End Sub


Private Sub UpdateUIForSaveAs(ByVal View As LambdaMetadataView)
    
    Logger.Log TRACE_LOG, "Enter Presenter.UpdateUIForSaveAs"
    ' Enable the NameTextBox and change button caption and caption text for Save As operation
    View.NameTextBox.Enabled = True
    View.OkButton.Caption = "Save"
    View.Caption = " Save Lambda As"
    Logger.Log TRACE_LOG, "Exit Presenter.UpdateUIForSaveAs"
    
End Sub

'@PureFunction
Private Function IsAnyMetadataFound(ByVal LambdaMetadata As Metadata) As Boolean
    
    Logger.Log TRACE_LOG, "Enter Presenter.IsAnyMetadataFound"
    ' Check if any essential metadata (CommandName, LambdaName, Description) is present
    Dim PropertiesName As Variant
    PropertiesName = Array("CommandName", "LambdaName", "Description")
    Dim PropertyName As Variant
    For Each PropertyName In PropertiesName
        If CallByName(LambdaMetadata, CStr(PropertyName), VbGet) <> vbNullString Then
            IsAnyMetadataFound = True
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Presenter.IsAnyMetadataFound"
            Exit Function
        End If
    Next PropertyName
    
    ' Check if any additional metadata or properties are present in the LambdaMetadata
    IsAnyMetadataFound = IsPropertyBelongToMoreInTheUIPresent(LambdaMetadata)
    Logger.Log TRACE_LOG, "Exit Presenter.IsAnyMetadataFound"
   
End Function

Private Function IsPropertyBelongToMoreInTheUIPresent(ByVal LambdaMetadata As Metadata) As Boolean
    
    Logger.Log TRACE_LOG, "Enter Presenter.IsPropertyBelongToMoreInTheUIPresent"
    ' Check if additional properties (GistURL, SourceName) are present in the LambdaMetadata
    Dim PropertiesName As Variant
    PropertiesName = Array("GistURL", "SourceName")
    Dim PropertyName As Variant
    For Each PropertyName In PropertiesName
        If CallByName(LambdaMetadata, CStr(PropertyName), VbGet) <> vbNullString Then
            IsPropertyBelongToMoreInTheUIPresent = True
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Presenter.IsPropertyBelongToMoreInTheUIPresent"
            Exit Function
        End If
    Next PropertyName
    
    ' Check if Parameters, CustomProperties, or Dependencies are present in the LambdaMetadata
    If LambdaMetadata.GetParameters(, False) <> vbNullString Then
        IsPropertyBelongToMoreInTheUIPresent = True
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Presenter.IsPropertyBelongToMoreInTheUIPresent"
        Exit Function
    End If
    
    PropertiesName = Array("GetCustomProperties", "GetDependencies")
    
    For Each PropertyName In PropertiesName
        If CallByName(LambdaMetadata, CStr(PropertyName), VbMethod) <> vbNullString Then
            IsPropertyBelongToMoreInTheUIPresent = True
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Presenter.IsPropertyBelongToMoreInTheUIPresent"
            Exit Function
        End If
    Next PropertyName
    Logger.Log TRACE_LOG, "Exit Presenter.IsPropertyBelongToMoreInTheUIPresent"
    
End Function

Private Sub Class_Terminate()
    
    Logger.Log TRACE_LOG, "Enter Presenter.Class_Terminate"
    ClearCatches
    Logger.Log TRACE_LOG, "Exit Presenter.Class_Terminate"
    
End Sub

Public Sub ClearCatches()
    
    Logger.Log TRACE_LOG, "Enter Presenter.ClearCatches"
    Set This.MetaParser = Nothing
    Set This.LambdaMetadata = Nothing
    Set This.View = Nothing
    Logger.Log TRACE_LOG, "Exit Presenter.ClearCatches"
    
End Sub

Private Property Get IPresenter_CommandName() As String
    IPresenter_CommandName = This.LambdaMetadata.CommandName
End Property

Private Property Let IPresenter_CommandName(ByVal RHS As String)
    This.LambdaMetadata.CommandName = RHS
End Property

Private Property Get IPresenter_Description() As String
    IPresenter_Description = This.LambdaMetadata.Description
End Property

Private Sub IPresenter_ExpandCollapse(ByVal IsExpand As Boolean)
    Logger.Log TRACE_LOG, "Exit Presenter.IPresenter_ExpandCollapse"
    Logger.Log TRACE_LOG, "Enter Presenter.IPresenter_ExpandCollapse"

End Sub

Private Sub IPresenter_GenerateGist()
        
    Logger.Log TRACE_LOG, "Enter Presenter.IPresenter_GenerateGist"
    ' Make the LambdaName a valid defined name by removing invalid characters
    This.LambdaMetadata.LambdaName = MakeValidDefinedName(This.LambdaMetadata.LambdaName, False)
    
    ' Get the Gist text from the exported Lambda formula
    Dim GistText As String
    GistText = ReplaceNewlineWithChar10(GetExportText())
    This.GistText = GistText
    
    ' Log the generated Gist text
    Logger.Log DEBUG_LOG, "Gist Text : " & NEW_LINE & GistText
    
    ' Copy the Gist text to the clipboard
    CreateObject("htmlfile").parentWindow.clipboardData.SetData "text", GistText
    
    ' Log function exit
    Logger.Log TRACE_LOG, "Exit Presenter.IPresenter_GenerateGist"
    
End Sub

Private Function GetExportText() As String
    
    Logger.Log TRACE_LOG, "Enter Presenter.GetExportText"
    ' Generate the new Lambda formula based on LambdaMetadata and MetaParser
    Dim NewLambdaFormula As String
    NewLambdaFormula = GenerateNewLambda(This.LambdaMetadata, This.MetaParser)
    
    NewLambdaFormula = GetLambdaDefPart(NewLambdaFormula) & ";"
    
    ' Get the comments to append to the Lambda formula based on LambdaMetadata
    Dim Comments As String
    Comments = GetComments(This.LambdaMetadata)
    
    ' Concatenate the comments with the LambdaName and NewLambdaFormula
    GetExportText = IIf(Comments = vbNullString, vbNullString, Comments & NEW_LINE) & _
                    This.LambdaMetadata.LambdaName & ONE_SPACE & NewLambdaFormula
    Logger.Log TRACE_LOG, "Exit Presenter.GetExportText"
                
End Function

'@PureFunction
Private Function GetComments(ByVal LambdaMetadata As Metadata) As String
    
    Logger.Log TRACE_LOG, "Enter Presenter.GetComments"
    ' Generate the comments block based on LambdaMetadata properties
    GetComments = "/*" & NEW_LINE & _
                  GetNameCommentBlock(This.LambdaMetadata.LambdaName, This.LambdaMetadata.CommandName) & NEW_LINE & _
                  GetCommentToAppend("Description", LambdaMetadata.Description, NEW_LINE) & _
                  GetCommentToAppend("Parameters:", LambdaMetadata.GetParameters(, False), NEW_LINE, NEW_LINE) & _
                  GetCommentToAppend("Source", LambdaMetadata.SourceName, NEW_LINE) & _
                  GetCommentToAppend("Gist URL", LambdaMetadata.GistURL, NEW_LINE) & _
                  GetCommentToAppend("Dependencies:", LambdaMetadata.GetDependencies(, False), NEW_LINE, NEW_LINE) & _
                  "*/"
                  
    ' Check if comments block is empty and set to vbNullString if so
    If GetComments = "/*" & NEW_LINE & LambdaMetadata.LambdaName & NEW_LINE & "*/" Then
        GetComments = vbNullString
    End If
    
    ' Log function exit
    Logger.Log TRACE_LOG, "Exit Presenter.GetComments"
    
End Function

Private Function GetNameCommentBlock(ByVal LambdaName As String, ByVal CommandName As String) As String
    
    Logger.Log TRACE_LOG, "Enter Presenter.GetNameCommentBlock"
    ' Generate the Name comment block based on LambdaName and CommandName
    If CommandName <> vbNullString And LambdaName = vbNullString Then
        GetNameCommentBlock = CommandName
    ElseIf CommandName <> vbNullString And LambdaName <> vbNullString Then
        GetNameCommentBlock = CommandName & ONE_SPACE & FIRST_PARENTHESIS_OPEN & LambdaName & FIRST_PARENTHESIS_CLOSE
    ElseIf CommandName = vbNullString And LambdaName <> vbNullString Then
        GetNameCommentBlock = LambdaName
    Else
        GetNameCommentBlock = vbNullString
    End If
    
    ' Add "Name: " prefix to the comment block if not empty
    If GetNameCommentBlock <> vbNullString Then
        GetNameCommentBlock = "Name:  " & GetNameCommentBlock
    End If
    Logger.Log TRACE_LOG, "Exit Presenter.GetNameCommentBlock"
    
End Function

Private Function GetCommentToAppend(ByVal CommentKey As String _
                                    , ByVal CommentValue As String _
                                    , ByVal Suffix As String _
                                    , Optional ByVal KeyValueSeparator As String = ":  ") As String
    
    Logger.Log TRACE_LOG, "Enter Presenter.GetCommentToAppend"
    ' Check if CommentValue is not empty, and append the CommentKey, KeyValueSeparator, and CommentValue with the Suffix
    If CommentValue <> vbNullString Then
        GetCommentToAppend = CommentKey & KeyValueSeparator & CommentValue & Suffix
    Else
        GetCommentToAppend = vbNullString
    End If
    
    ' Log function exit
    Logger.Log TRACE_LOG, "Exit Presenter.GetCommentToAppend"
    
End Function

Private Function IPresenter_GetExportPreview() As String
    
    Logger.Log TRACE_LOG, "Enter Presenter.IPresenter_GetExportPreview"
    ' Make the LambdaName a valid defined name by removing invalid characters
    This.LambdaMetadata.LambdaName = MakeValidDefinedName(This.LambdaMetadata.LambdaName, False)
    
    ' Get the export preview text by calling GetExportText function
    IPresenter_GetExportPreview = GetExportText()
    Logger.Log TRACE_LOG, "Exit Presenter.IPresenter_GetExportPreview"
    
End Function

Private Function IPresenter_GetOverview() As String
       
    Logger.Log TRACE_LOG, "Enter Presenter.IPresenter_GetOverview"
    ' Concatenate the overview text with information about Sources, Parameters, Dependencies, and CustomProperties from LambdaMetadata
    IPresenter_GetOverview = "Source:" & NEW_LINE & _
                             This.LambdaMetadata.GetSources() & NEW_LINE & NEW_LINE & _
                             "Parameters:" & NEW_LINE & This.LambdaMetadata.GetParameters & NEW_LINE & _
                             NEW_LINE & "Dependencies: " & NEW_LINE & This.LambdaMetadata.GetDependencies & NEW_LINE & _
                             NEW_LINE & This.LambdaMetadata.GetCustomProperties
    
    ' Log function exit
    Logger.Log TRACE_LOG, "Exit Presenter.IPresenter_GetOverview"
    
End Function

Private Property Let IPresenter_IsCancelled(ByVal RHS As Boolean)
    This.IsProcessCancelled = RHS
End Property

Private Property Get IPresenter_IsCancelled() As Boolean
    IPresenter_IsCancelled = This.IsProcessCancelled
End Property

Private Property Let IPresenter_SourceName(ByVal RHS As String)
    This.LambdaMetadata.SourceName = RHS
End Property

Private Property Get IPresenter_SourceName() As String
    IPresenter_SourceName = This.LambdaMetadata.SourceName
End Property

Private Property Let IPresenter_GistURL(ByVal RHS As String)
    This.LambdaMetadata.GistURL = RHS
End Property

Private Property Get IPresenter_GistURL() As String
    IPresenter_GistURL = This.LambdaMetadata.GistURL
End Property

Private Property Get IPresenter_LambdaName() As String
    IPresenter_LambdaName = This.LambdaMetadata.LambdaName
End Property

Private Property Let IPresenter_LambdaName(ByVal RHS As String)
    This.LambdaMetadata.LambdaName = RHS
End Property

Private Property Let IPresenter_Description(ByVal RHS As String)
    This.LambdaMetadata.Description = RHS
End Property

Private Sub IPresenter_UpdateMetadataInFormula()
    
    Logger.Log TRACE_LOG, "Enter Presenter.IPresenter_UpdateMetadataInFormula"
    ' Make the LambdaName a valid defined name by removing invalid characters
    If Not IsValidDefinedName(This.LambdaMetadata.LambdaName) Then
        This.LambdaMetadata.LambdaName = MakeValidDefinedName(This.LambdaMetadata.LambdaName, False)
    End If
    
    ' Check if the LambdaName already exists and if it's a SaveAs operation, prompt for overwriting
    If modUtility.IsExistInCollection(This.MetaParser.AllNameInNameManager, This.LambdaMetadata.LambdaName) _
       And This.LambdaMetadata.LambdaName <> vbNullString And This.IsCallFromSaveAs Then
        
        If IsLambdaCreatedInExcelLabs(This.LambdaFormulaCell.Worksheet.Parent, This.LambdaMetadata.LambdaName) Then
            MsgBox "A lambda by this name already exists in Excel Labs' Advanced Formula Editor. Please choose a different name.", vbOKOnly + vbExclamation
            This.View.NameTextBox.SetFocus
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Presenter.IPresenter_UpdateMetadataInFormula"
            Exit Sub
        End If
        
        Dim Answer As VbMsgBoxResult
        Answer = MsgBox("The name " & This.LambdaMetadata.LambdaName & " already exists. Do you want to over-write it?" _
                        , vbYesNo + vbQuestion, APP_NAME)
        If Answer = vbNo Then
            This.View.NameTextBox.SetFocus
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Presenter.IPresenter_UpdateMetadataInFormula"
            Exit Sub
        End If
        
    ElseIf This.LambdaMetadata.LambdaName = vbNullString And IsCallFromSaveAs Then
        MsgBox "Lambda Name is blank. Please provide a valid name.", vbExclamation + vbOKOnly, APP_NAME
        This.View.NameTextBox.SetFocus
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Presenter.IPresenter_UpdateMetadataInFormula"
        Exit Sub
    End If
    
    ' Generate the new lambda formula and update the cell or name manager accordingly
    Dim NewLambdaFormula As String
    NewLambdaFormula = GenerateNewLambda(This.LambdaMetadata, This.MetaParser)
    If This.MetaParser.IsNeedToUpdateNameManager Then
        Dim OldName As name
        Set OldName = This.LambdaFormulaCell.Worksheet.Parent.Names(This.MetaParser.OldLambdaName)
        Logger.Log DEBUG_LOG, "Final Lambda to update in name manager : " & NEW_LINE & NewLambdaFormula
        OldName.RefersTo = ReplaceNewlineWithChar10(NewLambdaFormula)
        modUtility.UpdateNameComment OldName, This.LambdaMetadata.NameManagerComment
    Else
        AssignFormulaIfErrorPrintIntoDebugWindow This.LambdaFormulaCell, NewLambdaFormula, "Updated Formula : "
        This.LambdaFormulaCell.Calculate
    End If
    Logger.Log DEBUG_LOG, "New Lambda Formula : " & NewLambdaFormula
    Logger.Log DEBUG_LOG, THREE_NEW_LINE
    Logger.Log DEBUG_LOG, NewLambdaFormula
    
    ' Log function exit
    Logger.Log TRACE_LOG, "Exit Presenter.IPresenter_UpdateMetadataInFormula"
    
    ' Hide the view and release the event handler
    This.View.Hide
    Set This.View.EventHandler = Nothing
    Logger.Log TRACE_LOG, "Exit Presenter.IPresenter_UpdateMetadataInFormula"
    
End Sub

'@ImpureFunction.
Private Function GenerateNewLambda(ByVal LambdaMetadata As Metadata, ByVal MetaParser As MetadataParser) As String
        
    Logger.Log TRACE_LOG, "Enter Presenter.GenerateNewLambda"
    ' Remove metadata from the tokenized formula and get a copy of the new tokenized formula
    Dim FormulaText As String
    FormulaText = MetaParser.StartFormulaText
    Logger.Log DEBUG_LOG, "Before Removing Metadata : " & FormulaText
    FormulaText = RemoveMetadataFromFormula(FormulaText)
    Logger.Log DEBUG_LOG, "After Removing Metadata : " & FormulaText
    FormulaText = AddNewMetadata(FormulaText, LambdaMetadata)
    FormulaText = FormatFormula(FormulaText)
    Logger.Log DEBUG_LOG, "After adding Metadata : " & FormulaText
    GenerateNewLambda = FormulaText
    
    ' Log function exit
    Logger.Log TRACE_LOG, "Exit Presenter.GenerateNewLambda"
    
End Function

'@PureFunction
Private Function AddNewMetadata(ByVal FormulaText As String, ByVal LambdaMetadata As Metadata) As String
    
    Logger.Log TRACE_LOG, "Enter Presenter.AddNewMetadata"
    With LambdaMetadata
        ' As we are adding every step at index 1 we have to do this from the down side. So the end one(Command Name) will go in the index 1
        FormulaText = AddMetadataIfNotBlank(FormulaText, CUSTOMPROPERTIES_LET_VAR, .GetArrayConstant(CUSTOM_PROPERTIES, False))
        FormulaText = AddMetadataIfNotBlank(FormulaText, DEPENDENCIES_LET_VAR, .GetArrayConstant(LAMBDA_Dependencies, False))
        FormulaText = AddMetadataIfNotBlank(FormulaText, GIST_URL_LET_VAR, EscapeQuotes(.GistURL), QUOTES, QUOTES)
        FormulaText = AddMetadataIfNotBlank(FormulaText, SOURCE_NAME_LET_VAR, EscapeQuotes(.SourceName), QUOTES, QUOTES)
        FormulaText = AddMetadataIfNotBlank(FormulaText, PARAMETERS_LET_VAR, .GetArrayConstant(LAMBDA_PARAMETERS, False))
        FormulaText = AddMetadataIfNotBlank(FormulaText, DESCRIPTION_LET_VAR, EscapeQuotes(.Description), QUOTES, QUOTES)
        FormulaText = AddMetadataIfNotBlank(FormulaText, COMMAND_NAME_LET_VAR, EscapeQuotes(.CommandName), QUOTES, QUOTES)
        FormulaText = AddMetadataIfNotBlank(FormulaText, LAMBDA_NAME_LET_VAR, EscapeQuotes(.LambdaName), QUOTES, QUOTES)
    End With
    
    AddNewMetadata = FormulaText
    Logger.Log TRACE_LOG, "Exit Presenter.AddNewMetadata"
    
End Function

Private Function AddMetadataIfNotBlank(ByVal FormulaText As String, ByVal StepName As String _
                                                             , ByVal StepValue As String _
                                                              , Optional ByVal ValuePrefix As String = vbNullString _
                                                               , Optional ByVal ValueSuffix As String = vbNullString) As String
    
    Logger.Log TRACE_LOG, "Enter Presenter.AddMetadataIfNotBlank"
    If StepValue <> vbNullString Then
        AddMetadataIfNotBlank = InsertLetStep(FormulaText, 1, StepName, ValuePrefix & StepValue & ValueSuffix)
    Else
        AddMetadataIfNotBlank = FormulaText
    End If
    Logger.Log TRACE_LOG, "Exit Presenter.AddMetadataIfNotBlank"
    
End Function
