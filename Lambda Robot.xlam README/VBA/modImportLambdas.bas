Attribute VB_Name = "modImportLambdas"
Option Explicit
Option Private Module

Public Sub ImportAllLambdas(ByVal FileName As String _
                            , Optional ByVal ReplaceIfExists As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter modImportLambdas.ImportAllLambdas"
    Dim Calc As Integer
    Calc = Application.Calculation
    
    On Error GoTo ErrHandler
    
    Dim LambdaRobotPath As String
    LambdaRobotPath = GetLocalPathFromOneDrivePath(ThisWorkbook.Path) & Application.PathSeparator
    
    Dim Message As String
    
    If Text.Contains(FileName, ":\") Then
        
        If Not IsFileExist(FileName) Then
            Message = "The file specified, """ & FileName & """, was not found."
        ElseIf Not IsExcelFileNameOrPath(FileName) Then
            Message = "The file specified, """ & FileName & """, is not a valid excel file."
        End If
        
    ElseIf Not IsExcelFileNameOrPath(FileName) Then
        Message = "The file specified, """ & FileName & """, is not a valid excel file."
    
    ElseIf Not IsOpenWorkbookExists(FileName) And Not IsFileExist(LambdaRobotPath & FileName) Then
        Message = "The specified file, """ & FileName & """, was not found.  It must either be already open or located in the same folder as Lambda Robot."
        
    End If
    
    If Message <> vbNullString Then
        MsgBox Message, vbCritical + vbOKOnly, "Import All Lambdas"
        Exit Sub
    End If
    
    Dim AddToBook As Workbook
    Set AddToBook = ActiveWorkbook
    Dim BookName As String
    If IsFilePath(FileName) And Text.Contains(FileName, ":\") Then
        BookName = GetFileName(FileName)
    Else
        BookName = FileName
    End If
    
    Application.Calculation = xlCalculationManual
    
    Dim IsNeedToClose As Boolean
    Dim SourceBook As Workbook
    If IsWorkbookOpen(BookName) Then
        Set SourceBook = Application.Workbooks.Item(BookName)
    ElseIf IsFileExist(FileName) Then
        Set SourceBook = Application.Workbooks.Open(FileName)
        IsNeedToClose = True
    ElseIf IsFileExist(LambdaRobotPath & FileName) Then
        Set SourceBook = Application.Workbooks.Open(LambdaRobotPath & FileName)
        IsNeedToClose = True
    End If
    
    Const METHOD_NAME As String = "ImportAllLambdas"
    Context.ExtractContext SourceBook, METHOD_NAME
    
    Dim TablesName As Scripting.Dictionary
    Set TablesName = New Scripting.Dictionary
    
    Dim CurrentTable As ListObject
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In SourceBook.Worksheets
        For Each CurrentTable In CurrentSheet.ListObjects
            TablesName.Add CurrentTable.Name, CurrentTable.Name
        Next CurrentTable
    Next CurrentSheet
    
    Dim FirstSheet As Worksheet
    Set FirstSheet = SourceBook.Worksheets(1)
    
    Dim CurrentName As Name
    For Each CurrentName In SourceBook.Names
        
        If Not HasTableReferenceInRefersTo(TablesName, CurrentName.RefersTo, FirstSheet) Then
            AddNameIfValid CurrentName, AddToBook, ReplaceIfExists
        Else
            Logger.Log DEBUG_LOG, CurrentName.Name & " has table reference. skip this one..."
        End If
        
    Next CurrentName
    
    If IsNeedToClose Then
        SourceBook.Close False
    End If
    
    Application.Calculation = Calc
    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modImportLambdas.ImportAllLambdas"
    
ExitMethod:
    Context.ClearContext METHOD_NAME
    Exit Sub
    
ErrHandler:
    Application.Calculation = Calc
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume
    Logger.Log TRACE_LOG, "Exit modImportLambdas.ImportAllLambdas"
    
End Sub

Private Function HasTableReferenceInRefersTo(ByVal TablesName As Scripting.Dictionary _
                                             , ByVal RefersTo As String _
                                             , ByVal FromSht As Worksheet) As Boolean
    
    Dim Precedents As Variant
    On Error Resume Next
    Precedents = GetDirectPrecedents(RefersTo, FromSht)
    On Error GoTo 0
    
    Dim Result As Boolean
    
    If IsArray(Precedents) Then
        
        Dim CurrentPrecedent As Variant
        For Each CurrentPrecedent In Precedents
            
            If CurrentPrecedent <> vbNullString Then
                Dim TableName As String
                If Text.Contains(CStr(CurrentPrecedent), LEFT_BRACKET) And Text.IsEndsWith(CStr(CurrentPrecedent), RIGHT_BRACKET) Then
                    TableName = Text.BeforeDelimiter(CStr(CurrentPrecedent), LEFT_BRACKET)
                Else
                    TableName = CurrentPrecedent
                End If
                
                If TablesName.Exists(TableName) Then
                    Result = True
                    Exit For
                End If
            End If
            
        Next CurrentPrecedent
        
    End If
    
    HasTableReferenceInRefersTo = Result
    
End Function

Private Sub AddNameIfValid(ByVal CurrentName As Name _
                             , ByVal AddToBook As Workbook _
                              , ByVal ReplaceIfExists As Boolean)
    
    Logger.Log TRACE_LOG, "Enter modImportLambdas.AddLambdaIfValid"
    If IsBuiltInName(CurrentName) Then Exit Sub
    If IsLocalScopeNamedRange(CurrentName.Name) Then Exit Sub
    
    ' Ignore if it is referencing to any cell.
    If Not IsRefersToRangeIsNothing(CurrentName) Then Exit Sub
    
    With CurrentName
        
        ' Ignore this name if for some reason it can't be added.
        On Error Resume Next
        If Not Context.IsNamedRangeExists(AddToBook, .Name) Then
            AddToBook.Names.Add .Name, .RefersTo, .Visible
        ElseIf ReplaceIfExists Then
            AddToBook.Names(.Name).RefersTo = .RefersTo
        End If
        On Error GoTo 0
        
    End With
    Logger.Log TRACE_LOG, "Exit modImportLambdas.AddLambdaIfValid"
    
End Sub

Private Function IsWorkbookOpen(ByVal BookName As String) As Boolean
        
    Logger.Log TRACE_LOG, "Enter modImportLambdas.IsWorkbookOpen"
    On Error Resume Next
    Dim CurrentBook As Workbook
    Set CurrentBook = Application.Workbooks.Item(BookName)
    IsWorkbookOpen = (Not CurrentBook Is Nothing)
    On Error GoTo 0
    Logger.Log TRACE_LOG, "Exit modImportLambdas.IsWorkbookOpen"
    
End Function

Public Function IsExcelFileNameOrPath(ByVal FileNameOrPath As String) As Boolean
    
    IsExcelFileNameOrPath = ( _
                          FileNameOrPath Like "*.xl[a,s,t]" _
                          Or FileNameOrPath Like "*.xlam" _
                          Or FileNameOrPath Like "*.xls[b,m,x]" _
                          Or FileNameOrPath Like "*.xlt[m,x]" _
                          )
                          
End Function


