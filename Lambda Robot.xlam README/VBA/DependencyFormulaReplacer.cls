VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DependencyFormulaReplacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed
Option Explicit

Public Enum DependencyGenerationType
    SEND_RESULT_TO_SHEET = 1
    UPDATE_FORMULA_IN_CELL = 2
End Enum

Private Type TDependencyFormulaReplacer
    LambdaInSheet As Worksheet
    LambdaInBook As Workbook
    AllLambdas As Collection
    DependentOnLambdas As Collection
    AllDependency As Collection
    PutDependencyOnTable As ListObject
End Type

Private This As TDependencyFormulaReplacer

Public Property Get PutDependencyOnTable() As ListObject
    Set PutDependencyOnTable = This.PutDependencyOnTable
End Property

Public Property Set PutDependencyOnTable(ByVal RHS As ListObject)
    Set This.PutDependencyOnTable = RHS
End Property

Private Sub RetriveBasicStat(ByVal LambdaInCell As Range)
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.RetriveBasicStat"
    ' Set LambdaInSheet property with the parent sheet of the lambda cell
    Set This.LambdaInSheet = LambdaInCell.Worksheet
    ' Set LambdaInBook property with the parent workbook of the lambda cell
    Set This.LambdaInBook = LambdaInCell.Worksheet.Parent
    ' Set AllLambdas property with all the lambdas in the lambda workbook
    Set This.AllLambdas = modUtility.FindLambdas(This.LambdaInBook)
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.RetriveBasicStat"
        
End Sub

' @Entry
Public Sub IncludeLambdaDependencies(ByVal LambdaInCell As Range _
                                     , ByVal PutDependencyInfoOnRange As Range _
                                      , ByVal TypeOfOperation As DependencyGenerationType _
                                       , ByVal IsOnlyLetStepOnes As Boolean)
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.IncludeLambdaDependencies"
    RetriveBasicStat LambdaInCell
    ' Find the dependent Lambdas of the given formula
    FindDependentLambdas LambdaInCell.Formula2, 1
    RemoveOtherIfOnlyLetStep IsOnlyLetStepOnes
    ' Check if there are any dependencies, if none then exit the subroutine
    If This.AllDependency.Count = 0 Then Exit Sub
    ' Check the type of operation to be performed
    If TypeOfOperation = SEND_RESULT_TO_SHEET Then
        ' If the type of operation is to send the result to the sheet, then execute the SendDependencyDataToSheet method
        SendDependencyDataToSheet This.AllDependency, PutDependencyInfoOnRange
    Else
        ' Find the parameters and the LET variables in the given formula
        Dim ParametersAndLetVars As Collection
        Dim ParamAndStepNames As Variant
        ParamAndStepNames = GetParametersAndStepsName(LambdaInCell.Formula2)
        If Not IsArray(ParamAndStepNames) Then
            Set ParametersAndLetVars = New Collection
        Else
            Set ParametersAndLetVars = ArrayToCollection(ParamAndStepNames, , LBound(ParamAndStepNames, 2))
        End If
        
        ' Delete the named ranges for the LET step
        DeleteLetStepNamedRanges This.AllDependency, IsOnlyLetStepOnes
        ' Sort and retrieve the dependency data
        Set This.AllDependency = SortAndRetriveDependencyData(This.AllDependency)
        ' For each DependencyInfo in the AllDependency collection
        Dim CurrentDependencyInfo As DependencyInfo
        For Each CurrentDependencyInfo In This.AllDependency
            Dim CurrentValidVarName As String
            ' Check if only the LET step ones are required
            CurrentValidVarName = CurrentDependencyInfo.ValidVarName
            CurrentValidVarName = Text.RemoveFromStartIfPresent(CurrentValidVarName, LETSTEP_UNDERSCORE_PREFIX)
            
            ' Join the parameters, LET variables and dependent Lambdas
            Dim Temp As Collection
            Set Temp = JoinParametersLetVarsAndDependencyLambdas(ParametersAndLetVars _
                                                                 , This.AllDependency, CurrentValidVarName)
            ' Find a unique name by incrementing a number
            CurrentValidVarName = FindUniqueNameByIncrementingNumber(Temp, CurrentValidVarName)
            CurrentDependencyInfo.NewValidVarName = CurrentValidVarName
            
        Next CurrentDependencyInfo
        ' Get the new formula text
        Dim NewFormulaText As String
        NewFormulaText = InsertAllDependents(This.AllDependency, LambdaInCell.Formula2)
        NewFormulaText = ReplaceAllDependentsWithNewVarName(NewFormulaText, This.AllDependency)
        NewFormulaText = FormatFormula(NewFormulaText)
        ' Assign the new formula, if there is an error, print into the debug window
        AssignFormulaIfErrorPrintIntoDebugWindow LambdaInCell, NewFormulaText, "Independent Lambda : "
    End If
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.IncludeLambdaDependencies"
    
End Sub

Private Sub RemoveOtherIfOnlyLetStep(IsOnlyLetStepOnes As Boolean)
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.RemoveOtherIfOnlyLetStep"
    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In This.AllDependency
        If IsOnlyLetStepOnes And Not IsStartsWithLetStepPrefix(CurrentDependencyInfo.ValidVarName) Then
            This.AllDependency.Remove CurrentDependencyInfo.ValidVarName
        End If
    Next CurrentDependencyInfo
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.RemoveOtherIfOnlyLetStep"
    
End Sub

Private Function ReplaceAllDependentsWithNewVarName(ByVal FormulaText As String _
                                                    , ByVal AllDependency As Collection) As String
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.ReplaceAllDependentsWithNewVarName"
    Dim UpdatedFormula As String
    UpdatedFormula = FormulaText
    
    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In AllDependency
        With CurrentDependencyInfo
            UpdatedFormula = modDependencyLambdaResult.ReplaceTokenWithNewToken(UpdatedFormula, .ValidVarName, .NewValidVarName)
        End With
    Next CurrentDependencyInfo
    ReplaceAllDependentsWithNewVarName = UpdatedFormula
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.ReplaceAllDependentsWithNewVarName"
    
End Function

Private Sub DeleteLetStepNamedRanges(ByVal AllDependency As Collection, ByVal IsOnlyLetStepOnes As Boolean)
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.DeleteLetStepNamedRanges"
    ' Only proceed if we are working with LetStep ones
    If Not IsOnlyLetStepOnes Then Exit Sub
    ' Resume next on error
    On Error Resume Next
    ' Iterate over all dependencies
    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In AllDependency
        ' Delete the name of the current dependency
        CurrentDependencyInfo.ForName.Delete
    Next CurrentDependencyInfo
    ' Reset error handling
    On Error GoTo 0
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.DeleteLetStepNamedRanges"
    
End Sub

Private Function InsertAllDependents(ByVal AllDependency As Collection, OnFormula As String) As String
      
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.InsertAllDependents"
    Dim FormulaText As String
    FormulaText = OnFormula
    
    Dim CurrentDependencyInfo As DependencyInfo
    Dim Counter As Long
    For Each CurrentDependencyInfo In AllDependency
        With CurrentDependencyInfo
            FormulaText = InsertLetStep(FormulaText, Counter + 1 _
                                                    , .NewValidVarName _
                                                     , Text.RemoveFromBothEndIfPresent(.FormulaText, QUOTES))
            Counter = Counter + 1
        End With
    Next CurrentDependencyInfo
    InsertAllDependents = FormulaText
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.InsertAllDependents"
    
End Function

Private Sub ReplaceVarName(ByRef TokenizedFormula As Variant, ByVal OldVarName As String _
                                                             , ByVal NewVarName As String)
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.ReplaceVarName"
    ' Get the index of the first column
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(TokenizedFormula, 2)
    ' Iterate over all rows in the tokenized formula
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(TokenizedFormula, 1) To UBound(TokenizedFormula, 1)
        
        Dim TokenName As String
        TokenName = TokenizedFormula(CurrentRowIndex, FirstColumnIndex)
        Dim TokenType As String
        TokenType = TokenizedFormula(CurrentRowIndex, FirstColumnIndex + 1)
        ' If old variable name and "parameter or variable token" found, replace old variable name with new variable name
        If IsEqual(TokenName, OldVarName) And _
           (IsEqual(TokenType, modSharedConstant.NAMED_RANGE_TOKEN) Or IsEqual(TokenType, "Function")) Then
            TokenizedFormula(CurrentRowIndex, FirstColumnIndex) = NewVarName
        End If
        
    Next CurrentRowIndex
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.ReplaceVarName"

End Sub

Private Function JoinParametersLetVarsAndDependencyLambdas(ByVal ParametersAndLetVars As Collection _
                                                           , ByVal AllDependency As Collection _
                                                            , ByVal ExcludeItem As String) As Collection
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.JoinParametersLetVarsAndDependencyLambdas"
    ' Initialize result collection with ParametersAndLetVars
    Dim Result As Collection
    Set Result = DuplicateCollection(ParametersAndLetVars, True, True)
    ' Iterate over all dependencies
    Dim CurrentItem As DependencyInfo
    For Each CurrentItem In AllDependency
        ' Check if the current item's ValidVarName is not the one to be excluded
        If CurrentItem.ValidVarName <> ExcludeItem Then
            ' Check if the current item's ValidVarName doesn't already exist in the result collection
            If Not IsExistInCollection(Result, CurrentItem.ValidVarName) Then
                Result.Add CurrentItem.ValidVarName, CurrentItem.ValidVarName
            End If
        End If
    Next CurrentItem
    ' Return the joined result collection
    Set JoinParametersLetVarsAndDependencyLambdas = Result
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.JoinParametersLetVarsAndDependencyLambdas"
    
End Function

Private Function SortAndRetriveDependencyData(ByVal AllDependency As Collection) As Collection
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.SortAndRetriveDependencyData"
    ' Get all dependency data
    Dim DependencyData As Variant
    DependencyData = GetAllLambdaDependencyData(AllDependency, False)
    ' Get the index of the first column
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(DependencyData, 2)
    ' Initialize result collection
    Dim CurrentRowIndex As Long
    Dim CurrentDependencyInfo As DependencyInfo
    Dim Result As Collection
    Set Result = New Collection
    ' Iterate over all rows in the DependencyData
    For CurrentRowIndex = LBound(DependencyData, 1) To UBound(DependencyData, 1)
        ' Create new DependencyInfo object and fill its properties
        Set CurrentDependencyInfo = New DependencyInfo
        With CurrentDependencyInfo
            .ValidVarName = DependencyData(CurrentRowIndex, FirstColumnIndex)
            .FormulaText = DependencyData(CurrentRowIndex, FirstColumnIndex + 1)
            .Level = DependencyData(CurrentRowIndex, FirstColumnIndex + 2)
        End With
        ' Add the DependencyInfo object to the result collection
        Result.Add CurrentDependencyInfo, CurrentDependencyInfo.ValidVarName
    Next CurrentRowIndex
    ' Return the sorted and retrieved dependency data
    Set SortAndRetriveDependencyData = Result
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.SortAndRetriveDependencyData"
    
End Function

Private Sub SendDependencyDataToSheet(ByVal AllDependency As Collection, ByVal FirstCellToPutDependencyData As Range)
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.SendDependencyDataToSheet"
    ' Exit Subroutine if no dependencies
    If AllDependency.Count = 0 Then Exit Sub
    ' Get all dependency data
    Dim DependencyData As Variant
    DependencyData = GetAllLambdaDependencyData(AllDependency, True)
    
    ' Prepare the range on sheet where data will be written
    Dim PutDependencyDataOnRange As Range
    Set PutDependencyDataOnRange = FirstCellToPutDependencyData.Resize(AllDependency.Count + 1 _
                                                                       , UBound(DependencyData, 2))
    ' If there is any non-blank cell in the range, throw error message and exit subroutine
    If Not IsAllCellBlank(PutDependencyDataOnRange) Then
        MsgBox "Unable to write dependency info to worksheet.  Existing data in range " _
               & PutDependencyDataOnRange.Address & _
               " would be overwritten. Please clear enough space and run again.", vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword DependencyFormulaReplacer.SendDependencyDataToSheet"
        Exit Sub
    End If
    
    ' Write the dependency data to the sheet
    PutDependencyDataOnRange.Value = DependencyData
    ' Create a table for the dependency data on the sheet
    Set This.PutDependencyOnTable = This.LambdaInSheet.ListObjects.Add(xlSrcRange, PutDependencyDataOnRange, , xlYes)
    ' Adjust the width of the second column of the table
    This.PutDependencyOnTable.ListColumns(2).Range.ColumnWidth = 50
    ' Ensure that the text in the table doesn't wrap
    This.PutDependencyOnTable.Range.WrapText = False
    ' Scroll to the range with dependency data
    ScrollToDependencyDataRange This.PutDependencyOnTable
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.SendDependencyDataToSheet"
    
End Sub

Private Function SortLambdasDependencyByLevel(ByVal InputData As Variant) As Variant
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.SortLambdasDependencyByLevel"
    ' Define sort type: -1 for descending, 1 for ascending
    Const SORT_TYPE As String = "-1"

    ' Define the column to sort by
    Dim LevelColIndex As Long
    LevelColIndex = UBound(InputData, 2)
    
    Dim ByColumns As Variant
    ByColumns = Array(LevelColIndex)

    ' Split sort type string to create an array
    Dim SortType As Variant
    SortType = Array(CLng(SORT_TYPE))

    ' Apply the sort function
    Dim Result As Variant
    Result = Application.WorksheetFunction.Sort(InputData, ByColumns, SortType)

    ' Calculate the number of rows in the input data
    Dim NumberOfRow As Long
    NumberOfRow = UBound(InputData, 1) - LBound(InputData, 1) + 1

    ' If there is only one row, convert the result into an array
    If NumberOfRow = 1 Then
        Result = modUtility.VectorToArray(Result, True)
    End If

    ' Return the sorted result
    SortLambdasDependencyByLevel = Result
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.SortLambdasDependencyByLevel"
    
End Function

Private Function GetAllLambdaDependencyData(ByVal AllDependency As Collection, ByVal IsWithHeader As Boolean) As Variant
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.GetAllLambdaDependencyData"
    ' Retrieve dependency data from collection of objects
    Dim DependencyData As Variant
    DependencyData = RetriveLambdasDependencyData(AllDependency, IsWithHeader)
    
    ' Sort the dependency data by dependency level
    DependencyData = SortLambdasDependencyByLevel(DependencyData)

    ' Return the sorted dependency data
    GetAllLambdaDependencyData = DependencyData
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.GetAllLambdaDependencyData"
    
End Function

Private Function RetriveLambdasDependencyData(ByVal AllDependency As Collection _
                                                           , ByVal IsWithHeader As Boolean) As Variant
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.RetriveLambdasDependencyData"
    ' Create a new collection to store property names
    Dim PropertiesName As Collection
    Set PropertiesName = New Collection

    ' Add the property names to the collection
    PropertiesName.Add "ValidVarName", "ValidVarName"
    PropertiesName.Add "FormulaText", "FormulaText"
    PropertiesName.Add "Level", "Level"

    ' Get the property values of the objects in the AllDependency collection
    RetriveLambdasDependencyData = modUtility.GetObjectsPropertyValue(AllDependency _
                                                                                   , PropertiesName, IsWithHeader)
    
    ' Log exit from function
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.RetriveLambdasDependencyData"
    
End Function

'@Recursive
Private Sub FindDependentLambdas(ByVal FormulaText As String, ByVal Level As Long)
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.FindDependentLambdas"
    ' Find direct dependent lambdas
    Dim DirectDependentLambdas As Collection
    Set DirectDependentLambdas = FindDirectDependentLambdas(FormulaText)
    If DirectDependentLambdas.Count = 0 Then Exit Sub
    
    Dim CurrentMapper As Mapper
    For Each CurrentMapper In DirectDependentLambdas
        ' If lambda already exists in dependencies, update its level, else create new dependency and find its dependents
        If IsExistInCollection(This.AllDependency, CurrentMapper.Key) Then
            UpdateLevel CurrentMapper.Key, Level
        Else
            CreateNewDependency CurrentMapper, Level
            FindDependentLambdas CurrentMapper.Value.RefersTo, Level + 1
        End If
    Next CurrentMapper
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.FindDependentLambdas"
End Sub

Private Sub CreateNewDependency(ByVal CurrentMapper As Mapper, ByVal Level As Long)
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.CreateNewDependency"
    ' Create new dependency info object and set its properties
    Dim CurrentDependencyInfo As DependencyInfo
    Set CurrentDependencyInfo = New DependencyInfo
    With CurrentDependencyInfo
        .FormulaText = CurrentMapper.Value.RefersTo
        .Level = Level
        .ValidVarName = CurrentMapper.Value.name
        Set .ForName = CurrentMapper.Value
        .HasFormula = True
    End With

    ' Add the newly created dependency to the main dependency collection
    This.AllDependency.Add CurrentDependencyInfo, CurrentDependencyInfo.ValidVarName
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.CreateNewDependency"
    
End Sub

'@Recursive
Private Sub UpdateLevel(ByVal NameOfNamedRange As String, ByVal NewLevel As Long)
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.UpdateLevel"
    Dim CurrentDependency As DependencyInfo
    Set CurrentDependency = This.AllDependency.Item(NameOfNamedRange)
    
    '@BaseCase to exit
    If NewLevel < CurrentDependency.Level Then Exit Sub
    CurrentDependency.Level = NewLevel
    
    Dim DirectDependentLambdas As Collection
    Set DirectDependentLambdas = FindDirectDependentLambdas(CurrentDependency.FormulaText)
    
    Dim CurrentMapper As Mapper
    For Each CurrentMapper In DirectDependentLambdas
        ' In case of recursive lambda it will come again and again and it will be in infinite recursion.
        If CurrentMapper.Key <> NameOfNamedRange Then
            UpdateLevel CurrentMapper.Key, NewLevel + 1
        End If
    Next CurrentMapper
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.UpdateLevel"
    
End Sub

Private Function FindDirectDependentLambdas(ByVal FormulaText As String) As Collection
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.FindDirectDependentLambdas"
    ' This function finds the direct dependent lambdas in the given FormulaText.
    ' It checks for lambdas used in the formula and excludes lambdas that are let variables.
    
    ' Create a new collection to store the dependent lambdas.
    Dim DependentLambdas As Collection
    Set DependentLambdas = New Collection
    
    Dim UsedFunctions As Collection
    Set UsedFunctions = GetNamesAndFunctions(FormulaText)
    
    Dim CurrentUsedFunction As Variant
    For Each CurrentUsedFunction In UsedFunctions
        Dim LambdaName As String
        LambdaName = CStr(CurrentUsedFunction)
        ' If not added already and if it is a lambda then add in the dependent list.
        If Not IsExistInCollection(DependentLambdas, LambdaName) _
           And IsExistInCollection(This.AllLambdas, LambdaName) Then
            Dim CurrentMapper As Mapper
            Set CurrentMapper = Mapper.Create(LambdaName, This.AllLambdas.Item(LambdaName))
            DependentLambdas.Add CurrentMapper, LambdaName
        End If
    Next CurrentUsedFunction
    
    ' Set the result as the collection of dependent lambdas.
    Set FindDirectDependentLambdas = DependentLambdas
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.FindDirectDependentLambdas"
    
End Function

Private Function GetMapperIfLambdaRef(ByVal TokenName As String, ByVal TokenType As String) As Mapper
    
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.GetMapperIfLambdaRef"
    ' If token type is not named range then no mapper
    If Not (TokenType = NAMED_RANGE_TOKEN Or TokenType = "Function") Then
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword DependencyFormulaReplacer.GetMapperIfLambdaRef"
        Exit Function
    End If
    ' If token type is named range but not lambdas ref then no mapper as well.
    If Not IsExistInCollection(This.AllLambdas, TokenName) Then Exit Function
    
    Set GetMapperIfLambdaRef = Mapper.Create(TokenName, This.AllLambdas.Item(TokenName))
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.GetMapperIfLambdaRef"
        
End Function

Private Sub Class_Initialize()
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.Class_Initialize"
    ' This sub is automatically called when a new instance of the class is created.
    
    ' Initialize the AllLambdas, DependentOnLambdas, and AllDependency collections.
    Set This.AllLambdas = New Collection
    Set This.DependentOnLambdas = New Collection
    Set This.AllDependency = New Collection
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.Class_Initialize"
    
End Sub

Private Sub Class_Terminate()
    Logger.Log TRACE_LOG, "Enter DependencyFormulaReplacer.Class_Terminate"
    ' This sub is automatically called when the class instance is about to be destroyed.
    
    ' Clear the references to the collections to release memory resources.
    Set This.AllLambdas = Nothing
    Set This.DependentOnLambdas = Nothing
    Set This.AllDependency = Nothing
    Logger.Log TRACE_LOG, "Exit DependencyFormulaReplacer.Class_Terminate"
    
End Sub


