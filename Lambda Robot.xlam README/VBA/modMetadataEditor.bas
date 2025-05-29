Attribute VB_Name = "modMetadataEditor"
'@IgnoreModule UndeclaredVariable, UnrecognizedAnnotation, ProcedureNotUsed
'@Folder "Lambda.Editor.Metadata.Driver"

Option Explicit
Option Private Module

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Lambda Properties
' Description:            Edit Lambda properties.
' Macro Expression:       modMetadataEditor.EditLambdaProperties([ActiveCell])
' Generated:              05/27/2022 10:34 PM
'----------------------------------------------------------------------------------------------------
Public Sub EditLambdaProperties(ByVal OpenUIForCell As Range)
    
    Logger.Log TRACE_LOG, "Enter modMetadataEditor.EditLambdaProperties"
    ' Verify whether it's valid to run the command for the given cell
    ' If it's invalid, the subroutine is exited
    If modUtility.IsWorkbookOrWorksheetProtected(OpenUIForCell, "Edit Lambda Properties") Then Exit Sub
    
    On Error GoTo ErrorHandler
    Const METHOD_NAME As String = "EditLambdaProperties"
    Context.ExtractContextFromCell OpenUIForCell, METHOD_NAME
    ' Call a subroutine to edit lambda metadata
    EditLambdaMetadata OpenUIForCell
    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modMetadataEditor.EditLambdaProperties"
    Context.ClearContext METHOD_NAME
    Exit Sub

ErrorHandler:
    
    Context.ClearContext METHOD_NAME
    ' Collect error information
    Dim ErrorNumber As Long
    ErrorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description

    ' If there was an error, re-raise it for higher-level handling
    If ErrorNumber <> 0 Then
        Err.Raise ErrorNumber, Err.Source, ErrorDescription
        ' This line is only for debugging purposes.
        Resume
    End If
    Logger.Log TRACE_LOG, "Exit modMetadataEditor.EditLambdaProperties"
       
End Sub

Public Sub EditLambdaMetadata(ByVal ForCell As Range)

    Logger.Log TRACE_LOG, "Enter modMetadataEditor.EditLambdaMetadata"
    ' If the formula is a lambda and not in edit mode, save it considering available resources and exit the subroutine
    If IsLambdaFunction(GetCellFormula(ForCell)) And Not modUtility.IsLambdaInEditMode(ForCell, LAMBDA_NAME_NOTE_PREFIX) Then
        modLambdaEditor.SaveLambdaAsAfterTakingUserInput ForCell, InvokedFrom.EDIT_METADATA, GetCellFormula(ForCell)
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modMetadataEditor.EditLambdaMetadata"
        Exit Sub
    Else
        ' If the formula is not a lambda or is already in edit mode, extract the old name from comment
        Dim DefaultName As String
        DefaultName = GetOldNameFromComment(ForCell, LAMBDA_NAME_NOTE_PREFIX)

        ' Edit metadata for the cell
        Dim CurrentPresenter As Presenter
        Set CurrentPresenter = EditMetadataForCell(ForCell, InvokedFrom.EDIT_METADATA, DefaultName)
        Set CurrentPresenter = Nothing
    End If
    Logger.Log TRACE_LOG, "Exit modMetadataEditor.EditLambdaMetadata"
    
End Sub

Public Function EditMetadataForCell(ByVal ForCell As Range _
                                    , ByVal CallFrom As InvokedFrom _
                                     , Optional ByVal DefaultName As String = vbNullString _
                                      , Optional ByVal IsShowUI As Boolean = True) As Presenter
    
    Logger.Log TRACE_LOG, "Enter modMetadataEditor.EditMetadataForCell"
    ' Ensure that ForCell is not null and contains only a single cell
    If IsNothing(ForCell) Then
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modMetadataEditor.EditMetadataForCell"
        Exit Function
    ElseIf ForCell.Cells.Count > 1 Then
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modMetadataEditor.EditMetadataForCell"
        Exit Function
    ElseIf Not IsLambdaFunction(GetCellFormula(ForCell)) And Not modUtility.IsCellHasSavedLambdaFormula(ForCell) Then
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modMetadataEditor.EditMetadataForCell"
        Exit Function
    End If
    
    ' Create a new presenter object and show the view for editing the metadata of the lambda formula in ForCell
    Dim CurrentPresenter As Presenter
    Set CurrentPresenter = New Presenter
    CurrentPresenter.ShowView ForCell, CallFrom, DefaultName, IsShowUI
    Set EditMetadataForCell = CurrentPresenter
    Logger.Log TRACE_LOG, "Exit modMetadataEditor.EditMetadataForCell"
    
End Function


