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

Public Enum InvokedFrom
    SAVE_LAMBDA
    SAVE_LAMBDA_AS
    EDIT_METADATA
    OTHER
End Enum

Implements IPresenter

Private Type TPresenter
    LambdaMetadata As Metadata
    LambdaFormulaCell As Range
    TokenizedFormula As Variant
    View As LambdaMetadataView
    MetaParser As MetadataParser
    DefaultName As String
    IsProcessCancelled As Boolean
    GistText As String
    CallFrom As InvokedFrom
    IsLambdaSaveOperation As Boolean
End Type

Private this  As TPresenter

Public Property Get GistText() As String
    GistText = this.GistText
End Property

Public Property Get LambdaMetadata() As Metadata
    Set LambdaMetadata = this.LambdaMetadata
End Property

Public Property Get IsProcessCancelled() As Boolean
    IsProcessCancelled = this.IsProcessCancelled
End Property

'@EntryPoint
Public Sub ShowView(ByVal LambdaFormulaCell As Range _
                    , ByVal CallFrom As InvokedFrom _
                     , ByVal DefaultName As String _
                      , Optional ByVal IsShowUI As Boolean = True)
    
    Logger.Log TRACE_LOG, "Enter Presenter.ShowView"
    ' Create MetaParser and check if the cell contains a LAMBDA formula
    Set this.MetaParser = MetadataParser.Create(LambdaFormulaCell)
    If Not this.MetaParser.IsCellHasLambdaFormula Then
        MsgBox "Cell doesn't contain a LAMBDA formula.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Presenter.ShowView"
        Exit Sub
    End If
    
    ' Initialize properties and set LambdaMetadata
    Set this.LambdaFormulaCell = LambdaFormulaCell
    this.DefaultName = DefaultName
    Set this.LambdaMetadata = Metadata.CreateLambdaMetadata(this.MetaParser.LambdaParts _
                                                            , this.MetaParser.LetParts _
                                                             , this.DefaultName _
                                                              , this.MetaParser.OldLambdaName)
    this.CallFrom = CallFrom
    this.IsLambdaSaveOperation = (CallFrom = SAVE_LAMBDA Or CallFrom = SAVE_LAMBDA_AS)
    ' Check if any metadata is found and if not, exit the sub (unless it's called from Save As)
    If Not IsAnyMetadataFound(this.LambdaMetadata) _
       And Not this.IsLambdaSaveOperation _
       And Not (CallFrom = EDIT_METADATA And DefaultName = vbNullString) Then
        Exit Sub
    End If
    
    ' Show the LambdaMetadataView (if IsShowUI is True)
    If IsShowUI Then
        Set this.View = LambdaMetadataView.Create(this.LambdaMetadata, Me)
        this.View.UpdateView
        this.View.ExpandOrCollapse IsPropertyBelongToMoreInTheUIPresent(LambdaMetadata)
        
        If this.View.NameTextBox.Value = vbNullString Then this.View.NameTextBox.Enabled = True
        If this.IsLambdaSaveOperation Then UpdateUIForSaveAs this.View
        
        If this.View.NameTextBox.Enabled Then
            this.View.NameTextBox.SetFocus
        Else
            this.View.CommandNameTextBox.SetFocus
        End If
        
        this.View.Show
        On Error Resume Next
        Unload this.View
        Set this.View = Nothing
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
    
    If this.CallFrom = SAVE_LAMBDA Then
        View.Caption = " Save Lambda"
    ElseIf this.CallFrom = SAVE_LAMBDA_AS Then
        View.Caption = " Save Lambda As"
    End If
    
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
    Set this.MetaParser = Nothing
    Set this.LambdaMetadata = Nothing
    Set this.View = Nothing
    Logger.Log TRACE_LOG, "Exit Presenter.ClearCatches"
    
End Sub

Private Property Get IPresenter_CommandName() As String
    IPresenter_CommandName = this.LambdaMetadata.CommandName
End Property

Private Property Let IPresenter_CommandName(ByVal RHS As String)
    this.LambdaMetadata.CommandName = RHS
End Property

Private Property Get IPresenter_Description() As String
    IPresenter_Description = this.LambdaMetadata.Description
End Property

Private Sub IPresenter_ExpandCollapse(ByVal IsExpand As Boolean)
    Logger.Log TRACE_LOG, "Exit Presenter.IPresenter_ExpandCollapse"
    Logger.Log TRACE_LOG, "Enter Presenter.IPresenter_ExpandCollapse"

End Sub

Private Sub IPresenter_GenerateGist()
        
    Logger.Log TRACE_LOG, "Enter Presenter.IPresenter_GenerateGist"
    ' Make the LambdaName a valid defined name by removing invalid characters
    this.LambdaMetadata.LambdaName = MakeValidDefinedName(this.LambdaMetadata.LambdaName, False)
    
    ' Get the Gist text from the exported Lambda formula
    Dim GistText As String
    GistText = ReplaceNewlineWithChar10(GetExportText())
    this.GistText = GistText
    
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
    NewLambdaFormula = GenerateNewLambda(this.LambdaMetadata, this.MetaParser)
    
    NewLambdaFormula = GetLambdaDefPart(NewLambdaFormula) & ";"
    
    ' Get the comments to append to the Lambda formula based on LambdaMetadata
    Dim Comments As String
    Comments = GetComments(this.LambdaMetadata)
    
    ' Concatenate the comments with the LambdaName and NewLambdaFormula
    GetExportText = IIf(Comments = vbNullString, vbNullString, Comments & NEW_LINE) & _
                    this.LambdaMetadata.LambdaName & ONE_SPACE & NewLambdaFormula
    Logger.Log TRACE_LOG, "Exit Presenter.GetExportText"
                
End Function

'@PureFunction
Private Function GetComments(ByVal LambdaMetadata As Metadata) As String
    
    Logger.Log TRACE_LOG, "Enter Presenter.GetComments"
    ' Generate the comments block based on LambdaMetadata properties
    GetComments = "/*" & NEW_LINE & _
                  GetNameCommentBlock(this.LambdaMetadata.LambdaName, this.LambdaMetadata.CommandName) & NEW_LINE & _
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
    this.LambdaMetadata.LambdaName = MakeValidDefinedName(this.LambdaMetadata.LambdaName, False)
    
    ' Get the export preview text by calling GetExportText function
    IPresenter_GetExportPreview = GetExportText()
    Logger.Log TRACE_LOG, "Exit Presenter.IPresenter_GetExportPreview"
    
End Function

Private Function IPresenter_GetOverview() As String
       
    Logger.Log TRACE_LOG, "Enter Presenter.IPresenter_GetOverview"
    ' Concatenate the overview text with information about Sources, Parameters, Dependencies, and CustomProperties from LambdaMetadata
    IPresenter_GetOverview = "Source:" & NEW_LINE & _
                             this.LambdaMetadata.GetSources() & NEW_LINE & NEW_LINE & _
                             "Parameters:" & NEW_LINE & this.LambdaMetadata.GetParameters & NEW_LINE & _
                             NEW_LINE & "Dependencies: " & NEW_LINE & this.LambdaMetadata.GetDependencies & NEW_LINE & _
                             NEW_LINE & this.LambdaMetadata.GetCustomProperties
    
    ' Log function exit
    Logger.Log TRACE_LOG, "Exit Presenter.IPresenter_GetOverview"
    
End Function

Private Property Let IPresenter_IsCancelled(ByVal RHS As Boolean)
    this.IsProcessCancelled = RHS
End Property

Private Property Get IPresenter_IsCancelled() As Boolean
    IPresenter_IsCancelled = this.IsProcessCancelled
End Property

Private Property Let IPresenter_SourceName(ByVal RHS As String)
    this.LambdaMetadata.SourceName = RHS
End Property

Private Property Get IPresenter_SourceName() As String
    IPresenter_SourceName = this.LambdaMetadata.SourceName
End Property

Private Property Let IPresenter_GistURL(ByVal RHS As String)
    this.LambdaMetadata.GistURL = RHS
End Property

Private Property Get IPresenter_GistURL() As String
    IPresenter_GistURL = this.LambdaMetadata.GistURL
End Property

Private Property Get IPresenter_LambdaName() As String
    IPresenter_LambdaName = this.LambdaMetadata.LambdaName
End Property

Private Property Let IPresenter_LambdaName(ByVal RHS As String)
    this.LambdaMetadata.LambdaName = RHS
End Property

Private Property Let IPresenter_Description(ByVal RHS As String)
    this.LambdaMetadata.Description = RHS
End Property

Private Sub IPresenter_UpdateMetadataInFormula()
    
    ' Check if the LambdaName already exists and if it's a SaveAs operation, prompt for overwriting
    Dim IsValidToSave As Boolean
    CheckLambdaNameValidity IsValidToSave
    If Not IsValidToSave Then Exit Sub
    
    ' Generate the new lambda formula and update the cell or name manager accordingly
    Dim NewLambdaFormula As String
    NewLambdaFormula = GenerateNewLambda(this.LambdaMetadata, this.MetaParser)
    If this.MetaParser.IsNeedToUpdateNameManager Then
        Dim OldName As Name
        Set OldName = this.LambdaFormulaCell.Worksheet.Parent.Names(this.MetaParser.OldLambdaName)
        Logger.Log DEBUG_LOG, "Final Lambda to update in name manager : " & NEW_LINE & NewLambdaFormula
        OldName.RefersTo = ReplaceNewlineWithChar10(NewLambdaFormula)
        modUtility.UpdateNameComment OldName, this.LambdaMetadata.NameManagerComment
    Else
        AssignFormulaIfErrorPrintIntoDebugWindow False, this.LambdaFormulaCell, NewLambdaFormula, "Updated Formula : "
        this.LambdaFormulaCell.Calculate
    End If
    Logger.Log DEBUG_LOG, "New Lambda Formula : " & NewLambdaFormula
    Logger.Log DEBUG_LOG, THREE_NEW_LINE
    Logger.Log DEBUG_LOG, NewLambdaFormula
    
    ' Log function exit
    Logger.Log TRACE_LOG, "Exit Presenter.IPresenter_UpdateMetadataInFormula"
    
    ' Hide the view and release the event handler
    this.View.Hide
    Set this.View.EventHandler = Nothing
    Logger.Log TRACE_LOG, "Exit Presenter.IPresenter_UpdateMetadataInFormula"
    
End Sub

Private Sub CheckLambdaNameValidity(ByRef IsValidToSave As Boolean)
    
    IsValidToSave = True
    If Not this.IsLambdaSaveOperation Then Exit Sub
    
    Dim Reason As String
    Reason = modUtility.GetInvalidLambdaNameReason(this.LambdaMetadata.LambdaName)
    
    If Reason = vbNullString Then
        
        If IsExistInCollection(this.MetaParser.AllNameInNameManager, this.LambdaMetadata.LambdaName) Then
            If IsLambdaCreatedInExcelLabs(this.LambdaFormulaCell.Worksheet.Parent, this.LambdaMetadata.LambdaName) Then
                Reason = "A lambda by this name already exists in Excel Labs' Advanced Formula Editor. Please choose a different name."
            Else
                Dim Answer As VbMsgBoxResult
                Answer = MsgBox("The name " & this.LambdaMetadata.LambdaName & " already exists. Do you want to over-write it?" _
                                , vbYesNo + vbQuestion, APP_NAME)
                
                Reason = vbNullString
                If Answer = vbYes Then
                    IsValidToSave = True
                Else
                    IsValidToSave = False
                End If
                
            End If
        End If
                        
    End If
    
    If Reason <> vbNullString Then
        MsgBox Reason, vbExclamation + vbOKOnly, APP_NAME
        IsValidToSave = False
    End If
    
    If Not IsValidToSave Then
        this.View.NameTextBox.SetFocus
        Exit Sub
    End If
    
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


