VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LambdaMetadataView 
   Caption         =   "Lambda Properties"
   ClientHeight    =   11100
   ClientLeft      =   -270
   ClientTop       =   -1290
   ClientWidth     =   13485
   OleObjectBlob   =   "LambdaMetadataView.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "LambdaMetadataView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, ProcedureNotUsed
'@Folder "Lambda.Editor.Metadata.View"

Option Explicit

Private Const METADATA_PAGE_CAPTION_TEXTBOX_NAME_PREFIX As String = "CaptionTextBoxFor"
Private Const METADATA_PAGE_VALUE_TEXTBOX_NAME_PREFIX As String = "ValueTextBoxFor"
Private Const METADATA_PAGE_NEW_BUTTON_NAME_PREFIX As String = "NewKeyValuePairButtonFor"
Private Const METADATA_PAGE_UPDATE_BUTTON_NAME_PREFIX As String = "UpdateKeyValuPairButtonFor"
Private Const METADATA_PAGE_DELETE_BUTTON_NAME_PREFIX As String = "DeleteKeyValuPairButtonFor"
Private Const METADATA_PAGE_LIST_BOX_NAME_PREFIX As String = "KeyValueMapHolderFor"
Private Const METADATA_PAGE_UP_BUTTON_NAME_PREFIX As String = "UpButtonFor"
Private Const METADATA_PAGE_DOWN_BUTTON_NAME_PREFIX As String = "DownButtonFor"

Private Type TLambdaMetadata
    CrudOperator As ICRUD
    GivenPresenter As IPresenter
    IsChangeByCode As Boolean
    NameTextBoxDefaultBackColor As Long
End Type

Private this  As TLambdaMetadata

Public EventHandler As Collection

Public Property Get CrudOperator() As ICRUD
    Set CrudOperator = this.CrudOperator
End Property

Public Property Set CrudOperator(ByVal RHS As ICRUD)
    Set this.CrudOperator = RHS
End Property

Public Property Get CurrentPresenter() As IPresenter
    Set CurrentPresenter = this.GivenPresenter
End Property

Public Property Set CurrentPresenter(ByVal RHS As IPresenter)
    Set this.GivenPresenter = RHS
End Property

Private Sub CancelButton_Click()
    Me.Hide
    this.GivenPresenter.IsCancelled = True
    Set EventHandler = Nothing
End Sub

Private Sub CommandNameTextBox_Change()
    
    this.GivenPresenter.CommandName = Me.CommandNameTextBox.Value
    
End Sub

Private Sub DescriptionTextBox_Change()
    
    this.GivenPresenter.Description = Me.DescriptionTextBox.Value
    
End Sub

Private Sub DynamicDataHolder_Change()
    
    ' Handle the change event for DynamicDataHolder (ListBox) in the LambdaMetadataView form
    Logger.Log TRACE_LOG, "Inside LambdaMetadataView.DynamicDataHolder_Change"
    Select Case Me.DynamicDataHolder.Value
        Case 0
            ' Update the overview of the metadata view
            UpdateOverView
        
        Case 1
            ' No CRUD operation required for this option
            Logger.Log DEBUG_LOG, "No crud operation"
        
        Case 2
            ' Apply CRUD operation on the Lambda Parameters group
            this.CrudOperator.ApplyOperationOnGroup LAMBDA_PARAMETERS
        
        Case 3
            ' Apply CRUD operation on the Lambda Dependencies group
            this.CrudOperator.ApplyOperationOnGroup LAMBDA_Dependencies
            
        Case 4
            ' Apply CRUD operation on the Custom Properties group
            this.CrudOperator.ApplyOperationOnGroup CUSTOM_PROPERTIES
        
        Case 5
            ' Show the export preview in the PreviewViewer TextBox
            Me.PreviewViewer.Value = this.GivenPresenter.GetExportPreview()
            Me.PreviewViewer.SelStart = 0
            Me.PreviewViewer.SetFocus
    End Select
    
End Sub

Public Property Get Self() As LambdaMetadataView
    ' Get reference to the current instance of LambdaMetadataView
    Set Self = Me
End Property

Public Function Create(ByVal CrudOperator As ICRUD, ByVal CurrentPresenter As IPresenter) As LambdaMetadataView
    
    ' Create a new instance of LambdaMetadataView and initialize it
    Logger.Log TRACE_LOG, "Inside LambdaMetadataView.Create"
    Dim UI    As LambdaMetadataView
    Set UI = New LambdaMetadataView
    With UI
        Set .CrudOperator = CrudOperator
        Set .CurrentPresenter = CurrentPresenter
        .DynamicDataHolder.Value = 0
        InitializeView UI
        Set Create = UI
    End With

End Function

Private Sub AddAllEventHandler()
    
    ' Add event handlers for various controls in the LambdaMetadataView form
    Logger.Log TRACE_LOG, "Inside LambdaMetadataView.AddAllEventHandler"
    Set EventHandler = New Collection
    Dim Groups As Variant
    Groups = Array(Group.LAMBDA_PARAMETERS, Group.LAMBDA_Dependencies, Group.CUSTOM_PROPERTIES)
    Dim Suffixs As Variant
    Suffixs = modSharedConstant.GetMetadataGroups()
    
    Dim CurrentSuffix As Variant
    Dim CurrentPageHandler As PageHandler
    Dim Counter As Long
    Dim CurrentGroup As Group
    For Each CurrentSuffix In Suffixs
        CurrentGroup = Groups(Counter)
        
        ' Create a new instance of PageHandler and set references to various controls
        Set CurrentPageHandler = PageHandler.Create(CurrentGroup _
                                                    , Me.Controls(METADATA_PAGE_CAPTION_TEXTBOX_NAME_PREFIX & CurrentSuffix) _
                                                     , Me.Controls(METADATA_PAGE_VALUE_TEXTBOX_NAME_PREFIX & CurrentSuffix))
     
        Set CurrentPageHandler.NewButton = Me.Controls(METADATA_PAGE_NEW_BUTTON_NAME_PREFIX & CurrentSuffix)
        Set CurrentPageHandler.UpdateButton = Me.Controls(METADATA_PAGE_UPDATE_BUTTON_NAME_PREFIX & CurrentSuffix)
        Set CurrentPageHandler.DeleteButton = Me.Controls(METADATA_PAGE_DELETE_BUTTON_NAME_PREFIX & CurrentSuffix)
        
        Set CurrentPageHandler.ListViewer = Me.Controls(METADATA_PAGE_LIST_BOX_NAME_PREFIX & CurrentSuffix)
        
        Set CurrentPageHandler.UpArrowButton = Me.Controls(METADATA_PAGE_UP_BUTTON_NAME_PREFIX & CurrentSuffix)
        Set CurrentPageHandler.DownArrowButton = Me.Controls(METADATA_PAGE_DOWN_BUTTON_NAME_PREFIX & CurrentSuffix)
                
        Set CurrentPageHandler.CrudOperator = this.CrudOperator
        
        ' Add the PageHandler to the EventHandler collection with a unique key
        EventHandler.Add CurrentPageHandler, METADATA_PAGE_NEW_BUTTON_NAME_PREFIX & CurrentSuffix
        Counter = Counter + 1
        
    Next CurrentSuffix
    
End Sub

Public Sub UpdateView()
    
    ' Update the LambdaMetadataView with the current metadata values and overview
    Logger.Log TRACE_LOG, "Inside LambdaMetadataView.UpdateView"
    
    ' If event handlers are not set, add them for various controls
    If IsNothing(EventHandler) Then AddAllEventHandler
    
    ' Set values in text boxes with the current metadata information
    Me.NameTextBox.Value = this.GivenPresenter.LambdaName
    Me.CommandNameTextBox.Value = this.GivenPresenter.CommandName
    Me.DescriptionTextBox.Value = this.GivenPresenter.Description
    Me.SourceNameTextBox.Value = this.GivenPresenter.SourceName
    Me.GistURLTextBox.Value = this.GivenPresenter.GistURL
    
    ' Update the overview and all list boxes in the view
    UpdateOverView
    UpdateAllListBox
    
End Sub

Private Sub UpdateAllListBox()
    
    ' Update all the list boxes (KeyValueMapHolders) with metadata group data
    Logger.Log TRACE_LOG, "Inside LambdaMetadataView.UpdateAllListBox"
    Dim Groups As Variant
    Groups = Array(Group.LAMBDA_PARAMETERS, Group.LAMBDA_Dependencies, Group.CUSTOM_PROPERTIES)
    Dim Suffixs As Variant
    Suffixs = modSharedConstant.GetMetadataGroups()
    
    Dim CurrentSuffix As Variant
    Dim Counter As Long
    Dim CurrentGroup As Group
    For Each CurrentSuffix In Suffixs
        
        CurrentGroup = Groups(Counter)
        ' Apply CRUD operation on the current metadata group
        this.CrudOperator.ApplyOperationOnGroup CurrentGroup
        
        ' Clear the list box to avoid duplicate entries
        Me.Controls(METADATA_PAGE_LIST_BOX_NAME_PREFIX & CurrentSuffix).Clear
        
        ' Read all the data for the current group and populate the list box
        Dim AllData As Variant
        AllData = this.CrudOperator.ReadAll
        If IsArray(AllData) Then
            Me.Controls(METADATA_PAGE_LIST_BOX_NAME_PREFIX & CurrentSuffix).List = AllData
            TryAdaptingScrollBarHeight Me.Controls(METADATA_PAGE_LIST_BOX_NAME_PREFIX & CurrentSuffix)
        End If
        Counter = Counter + 1
        
    Next CurrentSuffix
    
End Sub

Private Sub UpdateOverView()
    Me.OverViewTextBox.Value = this.GivenPresenter.GetOverview
End Sub

Private Sub ExpandCollapseToggle_Click()
    ExpandOrCollapseWithoutChangingStatus Me.ExpandCollapseToggle.Value
End Sub

Public Sub ExpandOrCollapse(ByVal IsExpand As Boolean)
    Me.ExpandCollapseToggle.Value = IsExpand
End Sub

Private Sub ExpandOrCollapseWithoutChangingStatus(ByVal IsExpand As Boolean)
    
    ' Debug print the value of IsExpand (for debugging purposes)
    Logger.Log DEBUG_LOG, "Is Expand : " & IsExpand
    
    ' Debug print the width of the form before expand or collapse
    Logger.Log DEBUG_LOG, "Before Expand Or Collapse : " & Me.Width
    
    ' If the form needs to be expanded
    If IsExpand Then
        ' Set the positions and visibility of controls for expanded view
        Me.OkButton.Top = 510
        Me.CancelButton.Top = 510
        Me.Height = 582.75
        Me.DynamicDataHolder.Visible = True
        Me.DynamicDataHolder.Enabled = True
        Me.DynamicDataHolder.Top = 186
        Me.GistButton.Visible = True
        Me.GistButton.Enabled = True
    Else
        ' Set the positions and visibility of controls for collapsed view
        Me.OkButton.Top = 153
        Me.CancelButton.Top = 153
        Me.Height = 218
        Me.DynamicDataHolder.Top = 190
        Me.DynamicDataHolder.Visible = False
        Me.DynamicDataHolder.Enabled = False
        Me.GistButton.Visible = False
        Me.GistButton.Enabled = False
    End If
    
    ' Adjust the width of the form based on the position of CancelButton and some extra space
    Me.Width = Me.CancelButton.Left + Me.CancelButton.Width + 25
    
    ' Debug print the width of the form after expand or collapse
    Logger.Log DEBUG_LOG, "After Expand Or Collapse : " & Me.Width
    
End Sub

Private Sub GistButton_Click()
    
    ' Generate the Gist for the given presenter
    Logger.Log TRACE_LOG, "Enter LambdaMetadataView.GistButton_Click"
    this.GivenPresenter.GenerateGist
    
    ' Display a message indicating that the Gist has been copied to clipboard
    Me.ExportProgressLabel.Caption = "Gist has been copied to clipboard"
    DoEvents
    
    ' Wait for 2 seconds (2000 milliseconds) before clearing the message
    Application.Wait (Now + TimeValue("00:00:02"))
    DoEvents
    
    ' Clear the progress label
    Me.ExportProgressLabel.Caption = vbNullString
    
End Sub

Private Sub GistURLTextBox_Change()
    this.GivenPresenter.GistURL = Me.GistURLTextBox.Value
End Sub

Private Sub NameTextBox_Change()
    
    If this.IsChangeByCode Then Exit Sub
    this.IsChangeByCode = True
    this.GivenPresenter.LambdaName = Me.NameTextBox.Value
    this.IsChangeByCode = False
    
End Sub

Private Sub OkButton_Click()

    Logger.Log TRACE_LOG, "Inside LambdaMetadataView.OkButton_Click"
    this.GivenPresenter.UpdateMetadataInFormula
    
End Sub

Private Sub SourceNameTextBox_Change()
    this.GivenPresenter.SourceName = Me.SourceNameTextBox.Value
End Sub

Private Sub InitializeView(ByVal View As LambdaMetadataView)
    
    ' Log the entry of the InitializeView procedure
    Logger.Log TRACE_LOG, "Inside LambdaMetadataView.UserForm_Initialize"
    
    ' Initialize the ListBox controls for each metadata group
    Dim Suffixs As Variant
    Suffixs = modSharedConstant.GetMetadataGroups()
    
    Dim CurrentSuffix As Variant
    Dim CurrentControl As MSForms.ListBox
    
    For Each CurrentSuffix In Suffixs
        
        Set CurrentControl = View.Controls(METADATA_PAGE_LIST_BOX_NAME_PREFIX & CurrentSuffix)
        
        With CurrentControl
            .ColumnCount = 2
            .ColumnWidths = "147 pt;379.45 pt"
            .TextAlign = fmTextAlignLeft
        End With
        
    Next CurrentSuffix
    
    ' Set the startup position and position the form at the center of the Excel application window
    View.StartUpPosition = 0
    View.Left = Application.Left + (0.5 * Application.Width) - (0.5 * View.Width)
    View.Top = Application.Top + (0.5 * Application.Height) - (0.5 * View.Height)
    
End Sub

Private Sub UserForm_Initialize()
    Me.NewKeyValuePairButtonForParameters.Visible = False
    Me.DeleteKeyValuPairButtonForParameters.Visible = False
    this.NameTextBoxDefaultBackColor = Me.NameTextBox.BackColor
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    ' If the user attempts to close the form, handle the Close event
    If CloseMode = CloseBy.User Then
        CancelButton_Click
    End If
    
End Sub

Private Sub UserForm_Activate()
    
    ' Set the startup position and position the form at the center of the Excel application window
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) * 0.5
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub


