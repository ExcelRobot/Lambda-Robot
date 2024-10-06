VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Exporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, ImplicitActiveWorkbookReference
'@Folder "Lambda.Editor.Exporter"
'@PredeclaredId
Option Explicit

Private Enum ModuleType
    GENERAL_MODULE = 1
    CLASS_MODULE = 2
    USER_FORM_MODULE = 3
    DOCUMENT_MODULE = 100
End Enum

Private Type TExporter
    WorkbookName As String
    GivenWorkbook As Workbook
    LambdaDefinitionOnCell As Range
    WorkbookPath As String
    IsWorkbookSaved As Boolean
    VBASubFolderPath As String
    PowerQuerySubFolderPath As String
    LambdaSubFolderPath As String
End Type

Private This As TExporter

Public Property Get WorkbookName() As String
    WorkbookName = This.WorkbookName
End Property

Public Property Let WorkbookName(ByVal RHS As String)
    This.WorkbookName = RHS
    UpdateBasicStat RHS
End Property

Private Sub UpdateBasicStat(ByVal WorkbookName As String)

    Logger.Log TRACE_LOG, "Enter Exporter.UpdateBasicStat"
    ' Updates the basic statistics and paths related to the given workbook.

    ' Set the reference to the specified workbook
    Set This.GivenWorkbook = GetWorkbook(WorkbookName)

    ' Find the cell that contains Lambda Definitions in the workbook
    Set This.LambdaDefinitionOnCell = FindHelperCell()

    ' Determine the path of the workbook
    This.WorkbookPath = This.GivenWorkbook.Path & Application.PathSeparator

    ' Check if the workbook has been saved (path is not empty)
    This.IsWorkbookSaved = (This.WorkbookPath <> vbNullString)

    ' Find the folder path for VBA sub-folder in the workbook
    This.VBASubFolderPath = FindFolderPath(This.GivenWorkbook, VBA_SUB_FOLDER_NAME)

    ' Find the folder path for Power Query sub-folder in the workbook
    This.PowerQuerySubFolderPath = FindFolderPath(This.GivenWorkbook, POWER_QUERY_SUB_FOLDER_NAME)

    ' Find the folder path for Lambda sub-folder in the workbook
    This.LambdaSubFolderPath = FindFolderPath(This.GivenWorkbook, LAMBDA_SUB_FOLDER_NAME)

    ' Log entry for exiting the function
    Logger.Log TRACE_LOG, "Exit Exporter.UpdateBasicStat"
    
End Sub

Public Function Create(Optional ByVal GivenWorkbookName As String = vbNullString) As Exporter
    
    Logger.Log TRACE_LOG, "Enter Exporter.Create"
    ' Factory function to create an instance of the Exporter class.

    ' Create a new Exporter object
    Dim CurrentExporter As Exporter
    Set CurrentExporter = New Exporter

    ' Assign the given workbook name to the exporter object
    With CurrentExporter
        .WorkbookName = GivenWorkbookName
    End With

    ' Set the function return value to the created Exporter object
    Set Create = CurrentExporter

    ' Log entry for exiting the function
    Logger.Log TRACE_LOG, "Exit Exporter.Create"
    
End Function

Private Function FindFolderPath(ByVal GivenBook As Workbook, ByVal SubFolderName As String) As String
    
    Logger.Log TRACE_LOG, "Enter Exporter.FindFolderPath"
    ' Finds the folder path for the specified subfolder name in the given workbook.

    ' If the given workbook has no path (e.g., a new unsaved workbook), return an empty string.
    ' Otherwise, return the full path of the subfolder name concatenated with the workbook's full name.
    If GivenBook.Path = vbNullString Then
        FindFolderPath = vbNullString
    Else
        FindFolderPath = GivenBook.FullName & ONE_SPACE & SubFolderName & Application.PathSeparator
    End If
    Logger.Log TRACE_LOG, "Exit Exporter.FindFolderPath"
    
End Function

Private Sub DeleteEmptyFolder(ByVal FolderPath As String)
    
    Logger.Log TRACE_LOG, "Enter Exporter.DeleteEmptyFolder"
    ' Deletes the empty folder at the specified path, if it exists and contains no files.

    ' Check if the folder exists using the FileSystemObject.
    ' If the folder exists and contains no files, delete the folder.
    Dim FolderManager As FileSystemObject
    Set FolderManager = New FileSystemObject
    
    If FolderManager.FolderExists(FolderPath) Then
        Dim CurrentFolder As Folder
        Set CurrentFolder = FolderManager.GetFolder(FolderPath)
        If CurrentFolder.Files.Count = 0 Then
            CurrentFolder.Delete
        End If
    End If
    
    Set FolderManager = Nothing
    Logger.Log TRACE_LOG, "Exit Exporter.DeleteEmptyFolder"
    
End Sub

Private Function GetWorkbook(Optional ByVal GivenWorkbookName As String = vbNullString) As Workbook
    
    Logger.Log TRACE_LOG, "Enter Exporter.GetWorkbook"
    ' Retrieves the workbook specified by the given workbook name or returns the ActiveWorkbook.

    ' Check if a specific workbook is given
    If GivenWorkbookName = vbNullString Then
        ' If not, return the ActiveWorkbook
        Set GetWorkbook = ActiveWorkbook
    Else
        ' If a workbook name is given, retrieve the corresponding workbook from the Workbooks collection
        Set GetWorkbook = Application.Workbooks.Item(GivenWorkbookName)
    End If

    ' Log exiting the function
    Logger.Log TRACE_LOG, "Exit Exporter.GetWorkbook"
    
End Function

Private Sub CreateFolderIfNotExist(ByVal FolderPath As String)
    
    Logger.Log TRACE_LOG, "Enter Exporter.CreateFolderIfNotExist"
    ' Creates the specified folder if it does not exist.

    ' Create a FileSystemObject to manage folders and files
    Dim FolderManager As FileSystemObject
    Set FolderManager = New FileSystemObject

    ' Check if the folder already exists
    If Not FolderManager.FolderExists(FolderPath) Then
        ' If the folder does not exist, create it
        FolderManager.Createfolder FolderPath
    End If

    ' Release the reference to the FileSystemObject
    Set FolderManager = Nothing

    ' Log exiting the function
    Logger.Log TRACE_LOG, "Exit Exporter.CreateFolderIfNotExist"
    
End Sub

Private Sub SetWorkbookName(ByVal GivenWorkbookName As String)
    
    Logger.Log TRACE_LOG, "Enter Exporter.SetWorkbookName"
    ' This Me. is important. Don't use This. as we are inside of that class and it will just set that
    ' value without using Let Property
    ' Me.WorkbookName = GivenWorkbookName
    
    ' Or if you want to use This then use in this way.
    If This.WorkbookName <> GivenWorkbookName Then
        This.WorkbookName = GivenWorkbookName
        UpdateBasicStat GivenWorkbookName
    End If
    Logger.Log TRACE_LOG, "Exit Exporter.SetWorkbookName"
    
End Sub

Public Sub ExportLambdas(Optional ByVal GivenWorkbookName As String = vbNullString)
    
    Logger.Log TRACE_LOG, "Enter Exporter.ExportLambdas"
    ' Exports lambdas from the given workbook to separate files.

    ' Set the workbook name and initialize the exporter
    SetWorkbookName GivenWorkbookName

    ' Check if the workbook has any named ranges
    If This.GivenWorkbook.Names.Count = 0 Then
        Logger.Log DEBUG_LOG, "No Named range found"
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Exporter.ExportLambdas"
        Exit Sub
    End If

    ' Find all lambdas in the workbook
    Dim AllLambda As Collection
    Set AllLambda = FindLambdas(This.GivenWorkbook)

    ' Check if there are any lambdas to export
    If AllLambda.Count = 0 Then
        Logger.Log DEBUG_LOG, "No Lambd found to export"
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Exporter.ExportLambdas"
        Exit Sub
    End If

    ' Check if the workbook is saved
    If This.LambdaSubFolderPath = vbNullString Then
        Logger.Log DEBUG_LOG, "File is not saved : "
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Exporter.ExportLambdas"
        Exit Sub
    End If

    ' Log the path where lambdas will be exported
    Logger.Log INFO_LOG, "Lambda will be exported to :" & This.LambdaSubFolderPath

    ' Create the lambda subfolder if it does not exist
    CreateFolderIfNotExist This.LambdaSubFolderPath

    On Error GoTo ErrorHandler
    ' Create a FileSystemObject to manage files and folders
    Dim FileManager As FileSystemObject
    Set FileManager = New FileSystemObject

    ' Loop through each lambda and export it to a separate file
    Dim CurrentName As name
    Dim CurrentFile As TextStream
    For Each CurrentName In AllLambda
        Logger.Log DEBUG_LOG, "Currently Processing Lambda : " & CurrentName.name
        Set CurrentFile = CreateTextFile(This.LambdaSubFolderPath, CurrentName.name _
                                                                  & LAMBDA_FILE_EXTENSION, FileManager)
        CurrentFile.Write GenerateGist(CurrentName)
        CurrentFile.Close
        Set CurrentFile = Nothing
    Next CurrentName
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Exporter.ExportLambdas"
    Exit Sub

ErrorHandler:
    ' Handle errors and release resources
    Dim errorNumber As Long
    errorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description
    
    ' Raise the error again if it occurred
    If errorNumber <> 0 Then
        Err.Raise errorNumber, Err.Source, ErrorDescription
        ' This is only for debugging purposes.
        Resume
    End If

    ' Log exiting the function
    Logger.Log TRACE_LOG, "Exit Exporter.ExportLambdas"
    
End Sub

Private Function GenerateGist(ByVal CurrentName As name) As String
    
    Logger.Log TRACE_LOG, "Enter Exporter.GenerateGist"
    ' Generate a code snippet (gist) for the given lambda (CurrentName).

    ' Set the lambda definition on the cell to the formula of the current named range
    This.LambdaDefinitionOnCell.Formula2 = CurrentName.RefersTo

    ' Create a presenter to show the lambda definition
    Dim CurrentPresenter As Presenter
    Set CurrentPresenter = New Presenter

    ' Show the lambda definition in the presenter view
    DoEvents
    CurrentPresenter.ShowView This.LambdaDefinitionOnCell, False, CurrentName.name, False

    ' Get the presenter interface to retrieve the export preview (gist)
    Dim PresenterInterface As IPresenter
    Set PresenterInterface = CurrentPresenter
    DoEvents
    GenerateGist = PresenterInterface.GetExportPreview()

    ' Log the generated gist for debugging purposes
    Logger.Log DEBUG_LOG, NEW_LINE
    Logger.Log DEBUG_LOG, GenerateGist

    ' Clean up and release resources
    This.LambdaDefinitionOnCell.Formula2 = vbNullString
    Set PresenterInterface = Nothing
    Set CurrentPresenter = Nothing
    Logger.Log TRACE_LOG, "Exit Exporter.GenerateGist"
    
End Function

Private Function FindHelperCell() As Range
    
    Logger.Log TRACE_LOG, "Enter Exporter.FindHelperCell"
    ' Find and return the helper cell (first unused cell) in the given workbook or active sheet.

    ' Check if the active sheet is Nothing (no active sheet)
    If IsNothing(This.GivenWorkbook.ActiveSheet) Then
        ' If no active sheet, find the first not used cell in the first worksheet of the given workbook
        Set FindHelperCell = FindFirstNotUsedCell(This.GivenWorkbook.Worksheets(1))
    Else
        ' If there is an active sheet, find the first not used cell in the active sheet
        Set FindHelperCell = FindFirstNotUsedCell(This.GivenWorkbook.ActiveSheet)
    End If
    Logger.Log TRACE_LOG, "Exit Exporter.FindHelperCell"
    
End Function

Private Function CreateTextFile(ByVal FolderPath As String, ByVal FileName As String _
                                                           , ByVal FileManager As FileSystemObject) As TextStream
    
    Logger.Log TRACE_LOG, "Enter Exporter.CreateTextFile"
    ' Create a text file in the specified folder path with the given file name using the provided FileSystemObject.

    ' Try creating the text file using the original folder path and file name
    On Error GoTo EncodePath
    Set CreateTextFile = FileManager.CreateTextFile(FolderPath & FileName, True, True)

    ' Log exiting the function due to the "Exit Function" statement
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Exporter.CreateTextFile"
    Exit Function

EncodePath:
    ' If an error occurs during file creation, encode the folder path and file name to handle special characters
    Dim EncodedPath As String
    EncodedPath = FolderPath & Application.WorksheetFunction.EncodeURL(FileName)
    Set CreateTextFile = FileManager.CreateTextFile(EncodedPath, True, True)
    Err.Clear

    ' Log exiting the function
    Logger.Log TRACE_LOG, "Exit Exporter.CreateTextFile"
    
End Function

Private Sub Class_Terminate()
    
    Logger.Log TRACE_LOG, "Enter Exporter.Class_Terminate"
    ' This sub is automatically called when the instance of the class is about to be destroyed.

    ' Check if the LambdaDefinitionOnCell object is not nothing (i.e., has been initialized)
    If IsNotNothing(This.LambdaDefinitionOnCell) Then
        ' Set the LambdaDefinitionOnCell to vbNullString (empty string) to release the reference
        Set This.LambdaDefinitionOnCell = Nothing
    End If
    Logger.Log TRACE_LOG, "Exit Exporter.Class_Terminate"
    
End Sub


