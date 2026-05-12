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
' [perf HOT-3232] Lazy classification state.
Private LambdasField As Collection
Private NonLambdasField As Collection
Private RangeRefNamedRangesField As Collection
Private LambdaCache As Scripting.Dictionary
Private RangeRefCache As Scripting.Dictionary
Private LambdasFullyClassified As Boolean
Private NonLambdasFullyClassified As Boolean
Private RangeRefFullyClassified As Boolean
Private CachedBookName As String
Private CachedBookRef As Workbook
Private IsContextExtracted As Boolean
Private AllTablesDic As Scripting.Dictionary
' [perf HOT-3232] Phase-time accumulators for IsLambdaName instrumentation.
' Reset at start of EnsureLambdasClassified, summary emitted at end.
Private LambdaTraceRefersToTime As Double
Private LambdaTraceContainsTime As Double
Private LambdaTraceParseTime As Double
Private LambdaTraceClassifyTime As Double
Private LambdaTraceTotalCalls As Long
Private LambdaTraceEarlyExits As Long
Private LambdaTraceParseAttempts As Long
Private LambdaTraceParseFailures As Long



' [perf HOT-3232] Lazy property getters � classification deferred until first read.
Public Property Get Lambdas() As Collection
    EnsureLambdasClassified
    Set Lambdas = LambdasField
End Property

Public Property Get NonLambdas() As Collection
    EnsureNonLambdasClassified
    Set NonLambdas = NonLambdasField
End Property

Public Property Get RangeRefNamedRanges() As Collection
    EnsureRangeRefClassified
    Set RangeRefNamedRanges = RangeRefNamedRangesField
End Property

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
    Set CachedBookRef = ForBook
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
    ' [perf HOT-3232] Defer iteration. Workbook.Names enumeration is itself
    ' ~1s per 1000 names in Excel COM. Build the dict only when an enumeration
    ' query (Lambdas / NonLambdas / RangeRefNamedRanges getters,
    ' GetAllNamedRange*FromBook) actually needs it. Direct single-name queries
    ' (IsLambdaExists, IsNamedRangeExists, IsLocalScopedNamedRangeExists) bypass
    ' this and use Workbook.Names(name) lookups.
    Set AllNamedRangeDic = Nothing
    Set AllNamedRangeColl = Nothing
    Set LambdasField = Nothing
    Set NonLambdasField = Nothing
    Set RangeRefNamedRangesField = Nothing
    LambdasFullyClassified = False
    NonLambdasFullyClassified = False
    RangeRefFullyClassified = False
    Set LambdaCache = New Scripting.Dictionary
    LambdaCache.CompareMode = vbTextCompare
    Set RangeRefCache = New Scripting.Dictionary
    RangeRefCache.CompareMode = vbTextCompare
    ' [perf HOT-3232] Reset per-name parse accumulators for the next operation.
    LambdaTraceRefersToTime = 0
    LambdaTraceContainsTime = 0
    LambdaTraceParseTime = 0
    LambdaTraceClassifyTime = 0
    LambdaTraceTotalCalls = 0
    LambdaTraceEarlyExits = 0
    LambdaTraceParseAttempts = 0
    LambdaTraceParseFailures = 0

End Sub

Private Sub EnsureNamedRangeDicBuilt()
    If Not AllNamedRangeDic Is Nothing Then Exit Sub
    If CachedBookRef Is Nothing Then Exit Sub

    Dim StartTime As Double
    If LOG_CONTEXT_TIMING Then StartTime = Timer

    Set AllNamedRangeDic = New Scripting.Dictionary
    AllNamedRangeDic.CompareMode = vbTextCompare
    Set AllNamedRangeColl = New Collection

    Dim CurrentName As Name
    For Each CurrentName In CachedBookRef.Names
        If Not IsBuiltInName(CurrentName) Then
            If Not Text.IsEndsWith(CurrentName.Name, "!_FilterDatabase") Then
                AllNamedRangeDic.Add CurrentName.Name, CurrentName
                AllNamedRangeColl.Add CurrentName, CurrentName.Name
            End If
        End If
    Next CurrentName

    If LOG_CONTEXT_TIMING Then
        Debug.Print "EnsureNamedRangeDicBuilt: " & Format$(Timer - StartTime, "0.000") & "s, " & AllNamedRangeDic.Count & " names"
    End If
End Sub

' [perf HOT-3232] Lazy classification helpers � called by Lambdas/NonLambdas/RangeRefNamedRanges
' getters. The whole-collection getters trigger a full iteration + classification pass.
' IsLambdaExists touches just one name and bypasses these via direct Workbook.Names lookup.

Private Sub EnsureLambdasClassified()
    If LambdasFullyClassified Then Exit Sub
    EnsureNamedRangeDicBuilt
    If AllNamedRangeDic Is Nothing Then Exit Sub
    Set LambdasField = New Collection
    Dim CurrentKey As Variant

    Dim CurrentName As Name
    For Each CurrentKey In AllNamedRangeDic.Keys
        Set CurrentName = AllNamedRangeDic(CurrentKey)
        If IsLambdaNameMemoized(CurrentName) Then
            LambdasField.Add CurrentName, CurrentName.Name
        End If
    Next CurrentKey

    LambdasFullyClassified = True
End Sub

Private Sub EnsureNonLambdasClassified()
    If NonLambdasFullyClassified Then Exit Sub
    EnsureNamedRangeDicBuilt
    If AllNamedRangeDic Is Nothing Then Exit Sub
    Set NonLambdasField = New Collection
    Dim CurrentKey As Variant
    Dim CurrentName As Name
    For Each CurrentKey In AllNamedRangeDic.Keys
        Set CurrentName = AllNamedRangeDic(CurrentKey)
        If Not IsLambdaNameMemoized(CurrentName) Then
            NonLambdasField.Add CurrentName, CurrentName.Name
        End If
    Next CurrentKey
    NonLambdasFullyClassified = True
End Sub

Private Sub EnsureRangeRefClassified()
    If RangeRefFullyClassified Then Exit Sub
    EnsureNamedRangeDicBuilt
    If AllNamedRangeDic Is Nothing Then Exit Sub
    Set RangeRefNamedRangesField = New Collection
    Dim CurrentKey As Variant
    Dim CurrentName As Name
    For Each CurrentKey In AllNamedRangeDic.Keys
        Set CurrentName = AllNamedRangeDic(CurrentKey)
        ' [perf HOT-3232] No explicit IsLambda check here � IsRangeRefMemoized
        ' returns False for lambdas (their RefersToRange errors), so an explicit
        ' filter would just force a full Parser.Parse pass over every name. This
        ' is the entire ~4s GenerateLambdaStatement hotspot.
        If IsRangeRefMemoized(CurrentName) Then
            RangeRefNamedRangesField.Add CurrentName, CurrentName.Name
        End If
    Next CurrentKey
    RangeRefFullyClassified = True
End Sub

Private Function IsLambdaNameMemoized(ByVal CurrentName As Name) As Boolean
    If LambdaCache Is Nothing Then
        IsLambdaNameMemoized = IsLambdaName(CurrentName)
        Exit Function
    End If
    If LambdaCache.Exists(CurrentName.Name) Then
        IsLambdaNameMemoized = LambdaCache(CurrentName.Name)
    Else
        IsLambdaNameMemoized = IsLambdaName(CurrentName)
        LambdaCache(CurrentName.Name) = IsLambdaNameMemoized
    End If
End Function

Private Function IsRangeRefMemoized(ByVal CurrentName As Name) As Boolean
    If RangeRefCache Is Nothing Then
        IsRangeRefMemoized = IsRangeRefNamedRange(CurrentName)
        Exit Function
    End If
    If RangeRefCache.Exists(CurrentName.Name) Then
        IsRangeRefMemoized = RangeRefCache(CurrentName.Name)
    Else
        IsRangeRefMemoized = IsRangeRefNamedRange(CurrentName)
        RangeRefCache(CurrentName.Name) = IsRangeRefMemoized
    End If
End Function

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
    LambdaTraceTotalCalls = LambdaTraceTotalCalls + 1

    Dim RefersToStr As String
    RefersToStr = CurrentName.RefersTo

    Dim HasLambdaText As Boolean
    HasLambdaText = Text.Contains(RefersToStr, "LAMBDA", IGNORE_CASE)

    If Not HasLambdaText Then
        LambdaTraceEarlyExits = LambdaTraceEarlyExits + 1
        IsLambdaName = False
        Exit Function
    End If

    Dim ParsedFormulaResult As Object
    Set ParsedFormulaResult = Parser.Parse(RefersToStr, False, LocaleFactory.EN_US, Scope, DefinedNames)
    LambdaTraceParseAttempts = LambdaTraceParseAttempts + 1

    If Not ParsedFormulaResult.ParseSuccess Then
        LambdaTraceParseFailures = LambdaTraceParseFailures + 1
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
    ' [perf HOT-3232] Clear the active error so it does not propagate to the caller's
    ' On Error handler. EnsureRangeRefClassified now calls this for lambdas (whose
    ' RefersToRange errors out), so this handler runs ~hundreds of times per pass.
    Err.Clear

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
        ' [perf HOT-3232] Build dict on first enumeration request, then reuse for the
        ' rest of the ExtractContext lifetime.
        EnsureNamedRangeDicBuilt
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
        ' Property accessor triggers EnsureRangeRefClassified � full pass + cache.
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
        ' [perf HOT-3232] Prefer dict if built; fall back to direct lookup.
        If Not AllNamedRangeDic Is Nothing Then
            IsExist = AllNamedRangeDic.Exists(NameOfTheNamedRange)
        Else
            Dim NamedRange As Name
            On Error Resume Next
            Set NamedRange = SearchInBook.Names(NameOfTheNamedRange)
            On Error GoTo 0
            IsExist = Not (NamedRange Is Nothing)
            If IsExist Then
                If IsBuiltInName(NamedRange) Or _
                   Text.IsEndsWith(NamedRange.Name, "!_FilterDatabase") Then
                    IsExist = False
                End If
            End If
        End If
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
        ' [perf HOT-3232] Prefer dict if built; fall back to direct lookup.
        ' Workbook.Names(...) accepts sheet-qualified names like "Sheet1!Foo".
        If Not AllNamedRangeDic Is Nothing Then
            Result = AllNamedRangeDic.Exists(SheetQualifiedName)
        Else
            Dim NamedRange As Name
            On Error Resume Next
            Set NamedRange = ScopeSheet.Parent.Names(SheetQualifiedName)
            On Error GoTo 0
            Result = Not (NamedRange Is Nothing)
            If Result Then
                If IsBuiltInName(NamedRange) Or _
                   Text.IsEndsWith(NamedRange.Name, "!_FilterDatabase") Then
                    Result = False
                End If
            End If
        End If
    Else
    
       If LOG_CONTEXT_TIMING Then Debug.Print "This shouldn't happened too often: IsLocalScopedNamedRangeExist"
        Dim FallbackName As Name
        For Each FallbackName In ScopeSheet.Names
            If LCase$(FallbackName.Name) = SheetQualifiedName Then
                Result = True
                Exit For
            End If
        Next FallbackName
    End If
    
    IsLocalScopedNamedRangeExists = Result
    
End Function

Public Function GetLambdaName(ByVal LambdaName As String) As Name
    ' [perf HOT-3232] Lazy single-name lambda lookup. Returns the Name object iff
    ' LambdaName resolves to a defined name AND it parses as a lambda. Returns Nothing
    ' otherwise. Memoizes the IsLambdaName parse via LambdaCache so each name is
    ' parsed at most once per ExtractContext lifetime � callers that ask "is X a lambda?"
    ' for a handful of names pay only that handful of parses, not the full N.
    If LambdaCache Is Nothing Then Exit Function

    Dim NamedRange As Name
    If Not AllNamedRangeDic Is Nothing Then
        If AllNamedRangeDic.Exists(LambdaName) Then
            Set NamedRange = AllNamedRangeDic(LambdaName)
        End If
    Else
        If CachedBookRef Is Nothing Then Exit Function
        On Error Resume Next
        Set NamedRange = CachedBookRef.Names(LambdaName)
        On Error GoTo 0
        If Not NamedRange Is Nothing Then
            If IsBuiltInName(NamedRange) Or _
               Text.IsEndsWith(NamedRange.Name, "!_FilterDatabase") Then
                Set NamedRange = Nothing
            End If
        End If
    End If

    If NamedRange Is Nothing Then Exit Function

    Dim Result As Boolean
    If LambdaCache.Exists(LambdaName) Then
        Result = LambdaCache(LambdaName)
    Else
        Result = IsLambdaName(NamedRange)
        LambdaCache(LambdaName) = Result
    End If

    If Result Then Set GetLambdaName = NamedRange
End Function

Public Function IsLambdaExists(ByVal LambdaName As String) As Boolean
    IsLambdaExists = Not (GetLambdaName(LambdaName) Is Nothing)
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
    Set LambdasField = Nothing
    Set NonLambdasField = Nothing
    Set RangeRefNamedRangesField = Nothing
    Set LambdaCache = Nothing
    Set RangeRefCache = Nothing
    LambdasFullyClassified = False
    NonLambdasFullyClassified = False
    RangeRefFullyClassified = False
    CachedBookName = vbNullString
    Set CachedBookRef = Nothing
    IsContextExtracted = False
    
End Sub


