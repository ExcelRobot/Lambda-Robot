VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed
'@GithubRawURL: https://raw.githubusercontent.com/1504168/All-Personal-VBA-Code/master/Reusable%20Code/Logger.cls
'@Folder "Utility"
'@PredeclaredId
Option Explicit

Const LEVEL_JSON_KEY As String = "Level"
Const TIME_JSON_KEY As String = "Log Time"
Const MESSAGE_JSON_KEY As String = "Message"
Const CALL_FROM_JSON_KEY As String = "Call From"
Const PARAMETERS_JSON_KEY As String = "Parameters"
Const LOG_FILE_OPEN_NUMBER As Long = 1
Const ALREADY_OPENED_FILE_ERROR_NUMBER As Long = 55
Private Const IS_PRODUCTION_MODE_ON As Boolean = True

Public IsLogFileOpen As Boolean

Public Enum Level
    INFO_LOG = 0
    TRACE_LOG = 1
    DEBUG_LOG = 2
    WARNING_LOG = 3
    ERROR_LOG = 4
End Enum

' Attention: This level description is linked with the Level Enum value
Private Function GetLogLevelDescription() As Variant
    GetLogLevelDescription = Array("INFO", "TRACE", "DEBUG", "DEBUG", "WARNING", "ERROR")
End Function

Public Sub Log(ByVal LogLevel As Level _
               , ByVal Message As String _
                , Optional ByVal CallFrom As String = vbNullString _
                 , Optional ByVal Parameters As Variant)
    
    If IS_PRODUCTION_MODE_ON Then Exit Sub
    
    Dim MapperCollection As Collection
    Set MapperCollection = New Collection
    
    Dim LogLevelDescription As String
    LogLevelDescription = GetLogLevelDescription(LogLevel)
    With MapperCollection
        .Add Mapper.Create(LEVEL_JSON_KEY, LogLevelDescription), LEVEL_JSON_KEY
        Dim CurrentTimeText As String
        CurrentTimeText = VBA.Format$(Now(), "DD MMM YYYY HH:MM:SS AM/PM")
        .Add Mapper.Create(TIME_JSON_KEY, CurrentTimeText), TIME_JSON_KEY
        .Add Mapper.Create(MESSAGE_JSON_KEY, Message), MESSAGE_JSON_KEY
        If CallFrom <> vbNullString Then
            .Add Mapper.Create(CALL_FROM_JSON_KEY, CallFrom), CALL_FROM_JSON_KEY
        End If
        If Not IsMissing(Parameters) Then
            .Add Mapper.Create(PARAMETERS_JSON_KEY, Parameters), PARAMETERS_JSON_KEY
        End If
    End With
    LogToDifferentLogContainer MapperCollection
End Sub

Private Sub LogToDifferentLogContainer(MapperCollection As Collection)
    
    Dim FileLogMessage As String
    FileLogMessage = PrepareJSON(MapperCollection)
    
    If Not IsLogFileOpen Then OpenLogFile
    On Error GoTo ExitLog
    Print #LOG_FILE_OPEN_NUMBER, FileLogMessage
    
    Dim ImmediateWindowLogMessage As String
    MapperCollection.Remove LEVEL_JSON_KEY
    MapperCollection.Remove TIME_JSON_KEY
    ImmediateWindowLogMessage = PrepareJSON(MapperCollection)
    ImmediateWindowLogMessage = VBA.Replace(ImmediateWindowLogMessage, "{", vbNullString, 1, 1)
    ImmediateWindowLogMessage = VBA.Left$(ImmediateWindowLogMessage, Len(ImmediateWindowLogMessage) - 1)
    Debug.Print ImmediateWindowLogMessage
    ' Application.StatusBar = ImmediateWindowLogMessage
    Exit Sub
    
ExitLog:
    Debug.Print Err.Description
    Debug.Print ImmediateWindowLogMessage
    
End Sub

Private Function PrepareJSON(ByVal KeyValuePair As Collection) As String
    Dim CurrentMapper As Mapper
    Dim JSONText As String
    For Each CurrentMapper In KeyValuePair
        JSONText = JSONText & "," & CurrentMapper.ToJSON
    Next CurrentMapper
    JSONText = VBA.Right$(JSONText, Len(JSONText) - 1)
    JSONText = "{" & JSONText & "}"
    PrepareJSON = JSONText
End Function

Public Sub OpenLogFile()
    
    Dim LogFilePath As String
    LogFilePath = GetLogFilePath
    On Error GoTo ErrorToOpenFile
    If Not IsLogFileOpen Then
        Open LogFilePath For Append As #LOG_FILE_OPEN_NUMBER
    End If
    
ErrorToOpenFile:
    If Err.Number = ALREADY_OPENED_FILE_ERROR_NUMBER Then
        IsLogFileOpen = True
        Log TRACE_LOG, "Log File is Opened for logging", "Logger.OpenLogFile"
    End If
    
End Sub

Private Function GetLogFilePath() As String
    Dim LogFolder As String
    LogFolder = ThisWorkbook.Path & Application.PathSeparator _
                & ThisWorkbook.name & ONE_SPACE & "Logs" & Application.PathSeparator
    CreateFolderIfNotExist LogFolder
    Dim LogFileName As String
    LogFileName = VBA.Format$(Date, "DD MMMM YYYY") & ".log"
    GetLogFilePath = LogFolder & LogFileName
End Function

Public Sub CloseLogFile()
    Log TRACE_LOG, "Log File is closing", "Logger.CloseLogFile"
    Close #LOG_FILE_OPEN_NUMBER
    IsLogFileOpen = False
End Sub

Private Sub CreateFolderIfNotExist(ByVal FolderPath As String)
    
    Dim FolderManager As Object
    Set FolderManager = CreateObject("Scripting.FileSystemObject")
    If Not FolderManager.FolderExists(FolderPath) Then
        FolderManager.Createfolder FolderPath
    End If
    Set FolderManager = Nothing
    
End Sub


