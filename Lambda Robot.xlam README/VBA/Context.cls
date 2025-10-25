VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Context"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredID
Option Explicit

#Const DEVELOPMENT_MODE = True
Private Const LOG_CONTEXT_TIMING As Boolean = False

Private DependentMethods As Scripting.Dictionary

#If DEVELOPMENT_MODE Then
    Public Scope As OARobot.FormulaScopeInfo
    Public DefinedNames As OARobot.XLDefinedNames
    Public Parser As OARobot.FormulaParser
    Public LocaleFactory As FormulaLocaleInfoFactory
    Public LocaleInfo As OARobot.FormulaLocaleInfo
    Public ExprProcessor As OARobot.ExprProcessing
    Public Formatter As OARobot.FormulaFormatter
    Public FormulaProcessor As OARobot.FormulaProcessing
    Public ExprReplacer As OARobot.ExpressionReplacer
    Public SheetNameRemover As OARobot.SheetNameRemover
    Public NamesOrFunctionsFilterer As OARobot.NamesOrFunctionsFilter
#Else
    Public Scope As Object
    Public DefinedNames As Object
    Public Parser As Object
    Public LocaleFactory As Object
    Public LocaleInfo As Object
    Public ExprProcessor As Object
    Public Formatter As Object
    Public FormulaProcessor As Object
    Public ExprReplacer As Object
    Public SheetNameRemover As Object
    Public NamesOrFunctionsFilterer As Object
#End If

Private AllNamedRangeDic As Scripting.Dictionary
Private AllNamedRangeColl As Collection
Public Lambdas As Collection
Public NonLambdas As Collection
Public RangeRefNamedRanges As Collection
Private CachedBookName As String
Private IsContextExtracted As Boolean
Private AllTablesDic As Scripting.Dictionary

Public Sub ExtractContextFromCell(ByVal FromCell As Range, ByVal CalledFromMethod As String)
    
    If IsNothing(FromCell) Then
        ExtractContext ActiveWorkbook, CalledFromMethod
    Else
        ExtractContext FromCell.Worksheet.Parent, CalledFromMethod
    End If
    
End Sub

Public Sub ExtractContext(ByVal ForBook As Workbook, ByVal CalledFromMethod As String)
    
    '    If IsContextExtracted Then Exit Sub
    '    Debug.Print "Context is called from: " & CalledFromMethod
    If IsNothing(ForBook) Then Exit Sub
    If IsNothing(DependentMethods) Then Set DependentMethods = New Scripting.Dictionary
    
    ' Increment method count or add if not present.
    DependentMethods.Item(CalledFromMethod) = DependentMethods.Item(CalledFromMethod) + 1
    
    ' Only Extract context for first method invocation.
    If IsContextExtracted Then Exit Sub
    
    If LOG_CONTEXT_TIMING Then
        Debug.Print
        Dim StartTime As Double
        StartTime = Timer
    End If
    
    #If DEVELOPMENT_MODE Then
        
        Dim ScopeFactory As New OARobot.FormulaScopeFactory
        Set Scope = ScopeFactory.CreateWorkbook(ForBook.Name)
        
        '        StartTime = Timer()
        Dim NamesFactory As New OARobot.DefinedNamesFactory
        Set DefinedNames = NamesFactory.Create(ForBook)
        '        Debug.Print "Total time to create Defined Names: " & Timer() - StartTime
        
        Set Parser = New OARobot.FormulaParser
        Set LocaleFactory = New FormulaLocaleInfoFactory
        Set LocaleInfo = LocaleFactory.CreateFromExcel(Application)
        Set ExprProcessor = New OARobot.ExprProcessing
        Set Formatter = New OARobot.FormulaFormatter
        Set FormulaProcessor = New OARobot.FormulaProcessing
        Set ExprReplacer = New OARobot.ExpressionReplacer
        Set SheetNameRemover = New OARobot.SheetNameRemover
        Set NamesOrFunctionsFilterer = New OARobot.NamesOrFunctionsFilter
    #Else
        Set Scope = CreateObject("OARobot.FormulaScopeFactory").CreateWorkbook(ForBook.Name)
        Set DefinedNames = CreateObject("OARobot.DefinedNamesFactory").Create(ForBook)
        Set Parser = CreateObject("OARobot.FormulaParser")
        Set LocaleFactory = CreateObject("FormulaLocaleInfoFactory")
        Set LocaleInfo = LocaleFactory.CreateFromExcel(Application)
        Set ExprProcessor = CreateObject("OARobot.ExprProcessing")
        Set Formatter = CreateObject("OARobot.FormulaFormatter")
        Set FormulaProcessor = CreateObject("OARobot.FormulaProcessing")
        Set ExprReplacer = CreateObject("OARobot.ExpressionReplacer")
        Set SheetNameRemover = CreateObject("OARobot.SheetNameRemover")
        Set NamesOrFunctionsFilterer = CreateObject("OARobot.NamesOrFunctionsFilter")
    #End If
    
    FormulaFormatConfig.Initialize
    ConfigFormulaFormatter
    
    ' This should be set to true before named range info Collection as it use IsLambdaFunction
    ' And if we don't set then it will pull again and again.
    
    UpdateNamedRangesCollection ForBook
    UpdateTableDictionary ForBook
    CachedBookName = ForBook.Name
    IsContextExtracted = True
    
    If LOG_CONTEXT_TIMING Then
        Debug.Print "Total Time to retrieve context: " & Timer - StartTime
    End If
    
End Sub

Public Function GetScope(ByVal ForBook As Workbook) As Object
    
    Dim Result As Object
    
    If ForBook Is Nothing Then
        Set Result = Scope
    ElseIf ForBook.Name = CachedBookName Then
        Set Result = Scope
    Else
        If LOG_CONTEXT_TIMING Then Debug.Print "This shouldn't happened too often: GetScope"
        
        #If DEVELOPMENT_MODE Then
            Dim ScopeFactory As New OARobot.FormulaScopeFactory
            Set Result = ScopeFactory.CreateWorkbook(ForBook.Name)
        #Else
            Set Result = CreateObject("OARobot.FormulaScopeFactory").CreateWorkbook(ForBook.Name)
        #End If
        
    End If
    
    Set GetScope = Result
    
End Function

Public Function GetDefinedNames(ByVal ForBook As Workbook) As Object
    
    Dim Result As Object
    
    If ForBook Is Nothing Then
        Set Result = DefinedNames
    ElseIf ForBook.Name = CachedBookName Then
        Set Result = DefinedNames
    Else
        If LOG_CONTEXT_TIMING Then Debug.Print "This shouldn't happened too often: GetDefinedNames"
        #If DEVELOPMENT_MODE Then
            Dim NamesFactory As New OARobot.DefinedNamesFactory
            Set Result = NamesFactory.Create(ForBook)
        #Else
            Set Result = CreateObject("OARobot.DefinedNamesFactory").Create(ForBook)
        #End If
        
    End If
    
    Set GetDefinedNames = Result
    
End Function

Private Sub ConfigFormulaFormatter()
        
    With Formatter
    
        ' Set configuration from user context.
        If FormulaFormatConfig.BoMode Then
            .CompactConfig
        Else
            .IndentChar = FormulaFormatConfig.IndentChar
            .IndentSize = FormulaFormatConfig.IndentSize
            .Multiline = FormulaFormatConfig.Multiline
            .OnlyWrapFunctionAfterNChars = FormulaFormatConfig.OnlyWrapFunctionAfterNChars
            .SpacesAfterArgumentSeparators = FormulaFormatConfig.SpacesAfterArgumentSeparators
            .SpacesAfterArrayColumnSeparators = FormulaFormatConfig.SpacesAfterArrayColumnSeparators
            .SpacesAfterArrayRowSeparators = FormulaFormatConfig.SpacesAfterArrayRowSeparators
            .SpacesAroundInfixOperators = FormulaFormatConfig.SpacesAroundInfixOperators
        End If
        
    End With
    
End Sub

Private Sub UpdateNamedRangesCollection(ByVal ForBook As Workbook)
    
    Dim StartTime As Double
    If LOG_CONTEXT_TIMING Then StartTime = Timer
    
    Set AllNamedRangeDic = New Scripting.Dictionary
    AllNamedRangeDic.CompareMode = vbTextCompare
    
    Set AllNamedRangeColl = New Collection
    Set Lambdas = New Collection
    Set NonLambdas = New Collection
    Set RangeRefNamedRanges = New Collection
    
    Dim CurrentName As Name
    For Each CurrentName In ForBook.Names
        
        ' Ignore built in and filter related named ranges
        If Not IsBuiltInName(CurrentName) Then
'            Debug.Print CurrentName.Name
            If Not Text.IsEndsWith(CurrentName.Name, "!_FilterDatabase") Then
            
                AllNamedRangeDic.Add CurrentName.Name, CurrentName
                AllNamedRangeColl.Add CurrentName, CurrentName.Name
                If IsLambdaName(CurrentName) Then
                    ' Add the name to the collection of lambda functions.
                    Lambdas.Add CurrentName, CurrentName.Name
                Else
                
                    NonLambdas.Add CurrentName, CurrentName.Name
                    If IsRangeRefNamedRange(CurrentName) Then
                        RangeRefNamedRanges.Add CurrentName, CurrentName.Name
                    End If
                
                End If
        
            End If
        
        End If
        
    Next CurrentName
    
    If LOG_CONTEXT_TIMING Then
        Debug.Print "Total Lambdas: " & Lambdas.Count
        Debug.Print "Total Non Lambdas: " & NonLambdas.Count
        Debug.Print "Total Range Ref Names: " & RangeRefNamedRanges.Count
        Debug.Print "Total time to retrieve named range info: " & Timer - StartTime
    End If
    
End Sub

Private Sub UpdateTableDictionary(ByVal ForBook As Workbook)
    
    Set AllTablesDic = New Scripting.Dictionary
    AllTablesDic.CompareMode = vbTextCompare
    
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In ForBook.Worksheets
        
        Dim CurrentTable As ListObject
        For Each CurrentTable In CurrentSheet.ListObjects
            AllTablesDic.Add CurrentTable.Name, CurrentSheet.Name
        Next CurrentTable
        
    Next CurrentSheet
    
End Sub

' Check if the outer function is LAMBDA or not and it is the entire function.
Private Function IsLambdaName(ByVal CurrentName As Name) As Boolean
    
'    Debug.Assert CurrentName.Name <> "TestLambda"
    If Not Text.Contains(CurrentName.RefersTo, "LAMBDA", IGNORE_CASE) Then
        IsLambdaName = False
        Exit Function
    End If
    
    Dim ParsedFormulaResult As Object
    
    Set ParsedFormulaResult = Parser.Parse(CurrentName.RefersTo, False, LocaleFactory.EN_US, Scope, DefinedNames)
    If Not ParsedFormulaResult.ParseSuccess Then
        IsLambdaName = False
        Exit Function
    End If
    
    Dim Result As Boolean
    If ParsedFormulaResult.Expr.IsLambda Then
        Result = True
    ElseIf ParsedFormulaResult.Expr.IsFunction Then
        Result = ParsedFormulaResult.Expr.AsFunction.FunctionName.IsLambda
    Else
        Result = False
    End If
    
    IsLambdaName = Result
    
End Function

Private Function IsRangeRefNamedRange(ByVal CurrentName As Name) As Boolean
    
    On Error GoTo NotRangeRef
    IsRangeRefNamedRange = IsNotNothing(CurrentName.RefersToRange)
    Exit Function
    
NotRangeRef:
    IsRangeRefNamedRange = False

End Function

' If same book that we already have cached then just returns the Collection
' Otherwise retrieve again and return. Most of the time we are working with single workbook.
' So this will definitely speed up the process.
Public Function GetAllNamedRangeCollectionFromBook(ByVal FromBook As Workbook) As Collection
    Set GetAllNamedRangeCollectionFromBook = GetAllNamedRangeFromBook(FromBook, True)
End Function

Public Function GetAllNamedRangeDictionaryFromBook(ByVal FromBook As Workbook) As Scripting.Dictionary
    Set GetAllNamedRangeDictionaryFromBook = GetAllNamedRangeFromBook(FromBook, False)
End Function

Private Function GetAllNamedRangeFromBook(ByVal FromBook As Workbook _
                                          , Optional ByVal IsAsCollection As Boolean = True) As Object
    
    Dim Result As Object
    
    If FromBook.Name = CachedBookName Then
        
        If IsAsCollection Then
            Set Result = AllNamedRangeColl
        Else
            Set Result = AllNamedRangeDic
        End If
        
    Else
        
        If LOG_CONTEXT_TIMING Then Debug.Print "This shouldn't happened too often: GetAllNamedRangeFromBook"
        If IsAsCollection Then
            Set Result = New Collection
        Else
            Set Result = New Scripting.Dictionary
            Result.CompareMode = vbTextCompare
        End If
        
        Dim CurrentName As Name
        For Each CurrentName In FromBook.Names
            If Not IsBuiltInName(CurrentName) And Not Text.IsEndsWith(CurrentName.Name, "!_FilterDatabase") Then
                If IsAsCollection Then
                    Result.Add CurrentName, CurrentName.Name
                Else
                    Result.Add CurrentName.Name, CurrentName
                End If
            End If
        Next CurrentName
        
    End If
    
    Set GetAllNamedRangeFromBook = Result
    
End Function

Public Function GetAllRangeRefNamedRangeFromBook(ByVal FromBook As Workbook) As Collection
    
    Dim Result As Collection
    
    If FromBook.Name = CachedBookName Then
        Set Result = RangeRefNamedRanges
    Else
        
        If LOG_CONTEXT_TIMING Then Debug.Print "This shouldn't happened too often: GetAllRangeRefNamedRangeFromBook"
        Set Result = New Collection
        
        Dim CurrentName As Name
        For Each CurrentName In FromBook.Names
            If Not IsBuiltInName(CurrentName) And IsRangeRefNamedRange(CurrentName) _
               And Not Text.IsEndsWith(CurrentName.Name, "!_FilterDatabase") Then
                Result.Add CurrentName, CurrentName.Name
            End If
        Next CurrentName
        
    End If
    
    Set GetAllRangeRefNamedRangeFromBook = Result
    
End Function

Public Function IsTableExists(ByVal SearchInBook As Workbook _
                              , ByVal TableName As String) As Boolean
    
    Dim IsExist As Boolean
    If SearchInBook.Name = CachedBookName Then
        IsExist = AllTablesDic.Exists(TableName)
    Else
        
        If LOG_CONTEXT_TIMING Then Debug.Print "This shouldn't happened too often: IsTableExists"
        
        Dim CurrentSheet As Worksheet
        For Each CurrentSheet In SearchInBook.Worksheets
        
            Dim CurrentTable As ListObject
            For Each CurrentTable In CurrentSheet.ListObjects
                If LCase$(CurrentTable.Name) = LCase$(TableName) Then
                    IsExist = True
                    Exit For
                End If
            Next CurrentTable
        
        Next CurrentSheet
        
    End If
    
    IsTableExists = IsExist
    
End Function

Public Function IsNamedRangeExists(ByVal SearchInBook As Workbook _
                                  , ByVal NameOfTheNamedRange As String) As Boolean
    
    ' Checks if a named range exists in the given workbook.
    Dim IsExist As Boolean
    If SearchInBook.Name = CachedBookName Then
        IsExist = AllNamedRangeDic.Exists(NameOfTheNamedRange)
    Else
        
        If LOG_CONTEXT_TIMING Then Debug.Print "This shouldn't happened too often: IsNamedRangeExist"
        Dim CurrentName As Name
        For Each CurrentName In SearchInBook.Names
            If LCase$(CurrentName.Name) = LCase$(NameOfTheNamedRange) Then
                IsExist = True
                Exit For
            End If
        Next CurrentName
    End If
    
    IsNamedRangeExists = IsExist
    
End Function

Public Function IsLocalScopedNamedRangeExists(ScopeSheet As Worksheet _
                                             , NamedRangeName As String) As Boolean
    
    Dim SheetQualifiedName As String
    SheetQualifiedName = NamedRangeName
    If Not Text.Contains(NamedRangeName, SHEET_NAME_SEPARATOR) Then
        SheetQualifiedName = GetSheetRefForRangeReference(ScopeSheet.Name, False) & NamedRangeName
    End If
    
    SheetQualifiedName = LCase$(SheetQualifiedName)
    
    Dim Result As Boolean
    If ScopeSheet.Parent.Name = CachedBookName Then
        Result = AllNamedRangeDic.Exists(SheetQualifiedName)
    Else
    
       If LOG_CONTEXT_TIMING Then Debug.Print "This shouldn't happened too often: IsLocalScopedNamedRangeExist"
        Dim CurrentName As Name
        For Each CurrentName In ScopeSheet.Names
            If LCase$(CurrentName.Name) = SheetQualifiedName Then
                Result = True
                Exit For
            End If
        Next CurrentName
    End If
    
    IsLocalScopedNamedRangeExists = Result
    
End Function

Public Function IsLambdaExists(ByVal LambdaName As String) As Boolean
    IsLambdaExists = IsExistInCollection(Lambdas, LambdaName)
End Function

Public Sub ClearContext(ByVal CalledFromMethod As String)
    
'    Debug.Print "Clear Context called from: " & CalledFromMethod
    
    If Not IsContextExtracted Then Exit Sub
    
    If IsNothing(DependentMethods) Then
        ReleaseContext
        Exit Sub
    End If
    
    If DependentMethods.Exists(CalledFromMethod) Then
        DependentMethods.Remove CalledFromMethod
    Else
        Err.Raise 13, , "Invalid source of clearing context."
'        Exit Sub
    End If
    
    If DependentMethods.Count = 0 Then
        ReleaseContext
    End If
    
End Sub

Private Sub ReleaseContext()
    
'    Debug.Print "Clearing context..."
    Set DependentMethods = Nothing
    
    Set Scope = Nothing
    Set DefinedNames = Nothing
    Set Parser = Nothing
    Set LocaleFactory = Nothing
    Set LocaleInfo = Nothing
    Set ExprProcessor = Nothing
    Set Formatter = Nothing
    Set FormulaProcessor = Nothing
    Set ExprReplacer = Nothing
    Set SheetNameRemover = Nothing
    Set NamesOrFunctionsFilterer = Nothing
    
    Set AllNamedRangeDic = Nothing
    Set AllNamedRangeColl = Nothing
    Set Lambdas = Nothing
    Set NonLambdas = Nothing
    Set RangeRefNamedRanges = Nothing
    CachedBookName = vbNullString
    IsContextExtracted = False
    
End Sub


