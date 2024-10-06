Attribute VB_Name = "modImportLambdas"
Option Explicit

Public Sub ImportAllLambdas(ByVal FileName As String _
                           , Optional ByVal ReplaceIfExists As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter modImportLambdas.ImportAllLambdas"
    Dim Calc As Integer
    Calc = Application.Calculation
    
    On Error GoTo ErrHandler
    
    Dim AddToBook As Workbook
    Set AddToBook = ActiveWorkbook
    Dim BookName As String
    If IsFilePath(FileName) Then
        BookName = GetFileName(FileName)
    Else
        BookName = FileName
    End If
    
    If Not IsExcelWorkbookName(BookName) Then
        MsgBox FileName & " seems invalid excel file name.", vbCritical + vbInformation, "Invalid Param"
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modImportLambdas.ImportAllLambdas"
        Exit Sub
    End If
    
    Application.Calculation = xlCalculationManual
    
    Dim IsNeedToClose As Boolean
    Dim SourceBook As Workbook
    If IsWorkbookOpen(BookName) Then
        Set SourceBook = Application.Workbooks.Item(BookName)
    Else
        Set SourceBook = Application.Workbooks.Open(FileName)
        IsNeedToClose = True
    End If
    
    Dim CurrentName As name
    For Each CurrentName In SourceBook.Names
        AddLambdaIfValid CurrentName, AddToBook, ReplaceIfExists
    Next CurrentName
    
    If IsNeedToClose Then
        SourceBook.Close False
    End If
    
    Application.Calculation = Calc
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modImportLambdas.ImportAllLambdas"
    Exit Sub
    
ErrHandler:
    Application.Calculation = Calc
    Err.Raise Err.Number, Err.Source, Err.Description
    Logger.Log TRACE_LOG, "Exit modImportLambdas.ImportAllLambdas"
    
End Sub

Private Sub AddLambdaIfValid(ByVal CurrentName As name _
                             , ByVal AddToBook As Workbook _
                              , ByVal ReplaceIfExists As Boolean)
    
    Logger.Log TRACE_LOG, "Enter modImportLambdas.AddLambdaIfValid"
    If IsBuiltInName(CurrentName) Then Exit Sub
    If Not IsLambdaFunction(CurrentName.RefersTo) Then Exit Sub
    
    If IsLocalScopeNamedRange(CurrentName.name) Then Exit Sub
    
    With CurrentName
        
        If Not IsNamedRangeExist(AddToBook, .name) Then
            AddToBook.Names.Add .name, .RefersTo, .Visible
        ElseIf ReplaceIfExists Then
            AddToBook.Names(.name).RefersTo = .RefersTo
        End If
        
    End With
    Logger.Log TRACE_LOG, "Exit modImportLambdas.AddLambdaIfValid"
    
End Sub

Private Function IsWorkbookOpen(ByVal BookName As String)
        
    Logger.Log TRACE_LOG, "Enter modImportLambdas.IsWorkbookOpen"
    On Error Resume Next
    Dim CurrentBook As Workbook
    Set CurrentBook = Application.Workbooks.Item(BookName)
    IsWorkbookOpen = (Not CurrentBook Is Nothing)
    On Error GoTo 0
    Logger.Log TRACE_LOG, "Exit modImportLambdas.IsWorkbookOpen"
    
End Function

Public Function IsExcelWorkbookName(ByVal BookName As String) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modImportLambdas.IsExcelWorkbookName"
    IsExcelWorkbookName = ( _
                          BookName Like "*.xl[a,s,t]" _
                          Or BookName Like "*.xlam" _
                          Or BookName Like "*.xls[b,m,x]" _
                          Or BookName Like "*.xlt[m,x]" _
                          )
    Logger.Log TRACE_LOG, "Exit modImportLambdas.IsExcelWorkbookName"
                          
End Function


