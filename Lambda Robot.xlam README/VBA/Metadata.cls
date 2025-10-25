VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Metadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed, ProcedureNotUsed
'@Folder "Lambda.Editor.Metadata.Model"
Option Explicit

Implements ICRUD

Private Type TMetadata
    LambdaName As String
    Description As String
    Parameters As ICRUD
    Dependencies As ICRUD
    CustomProperties As ICRUD
    CommandName As String
    SourceName As String
    GistURL As String
    NameManagerComment As String
End Type

Private this  As TMetadata

Private CurrentKeyValuePair As ICRUD

Public Property Get NameManagerComment() As String
    NameManagerComment = GetNameManagerComment()
End Property

Public Property Let NameManagerComment(ByVal RHS As String)
    this.NameManagerComment = RHS
End Property

Public Property Get SourceName() As String
    SourceName = this.SourceName
End Property

Public Property Let SourceName(ByVal RHS As String)
    this.SourceName = RHS
End Property

Public Property Get GistURL() As String
    GistURL = this.GistURL
End Property

Public Property Let GistURL(ByVal RHS As String)
    this.GistURL = RHS
End Property

Public Property Get CommandName() As String
    CommandName = this.CommandName
End Property

Public Property Let CommandName(ByVal RHS As String)
    this.CommandName = RHS
End Property

Public Property Get LambdaName() As String
    LambdaName = this.LambdaName
End Property

Public Property Let LambdaName(ByVal RHS As String)
    this.LambdaName = RHS
End Property

Public Property Get Description() As String
    Description = this.Description
End Property

Public Property Let Description(ByVal RHS As String)
    this.Description = RHS
End Property

Private Sub Class_Initialize()
    
    Logger.Log TRACE_LOG, "Enter Metadata.Class_Initialize"
    Set this.Parameters = New KeyValuePairCollection
    Set this.Dependencies = New KeyValuePairCollection
    Set this.CustomProperties = New KeyValuePairCollection
    Logger.Log TRACE_LOG, "Exit Metadata.Class_Initialize"
    
End Sub

Private Sub Class_Terminate()
    
    Logger.Log TRACE_LOG, "Enter Metadata.Class_Terminate"
    Set this.Parameters = Nothing
    Set this.Dependencies = Nothing
    Set this.CustomProperties = Nothing
    Logger.Log TRACE_LOG, "Exit Metadata.Class_Terminate"

End Sub

Private Sub ICRUD_ApplyOperationOnGroup(ByVal ToGroup As Group)
    UpdateCurrentKeyValuePairReference ToGroup
End Sub

Private Function ICRUD_ReadAll() As Variant
    ICRUD_ReadAll = CurrentKeyValuePair.ReadAll
End Function

Private Sub ICRUD_Swap(ByVal FirstKey As String, ByVal SecondKey As String)
    
    CurrentKeyValuePair.Swap FirstKey, SecondKey
    
End Sub

Private Sub ICRUD_AddNew(ByVal Key As String, ByVal Value As String)

    CurrentKeyValuePair.AddNew Key, Value
    
End Sub

Private Sub ICRUD_Update(ByVal OldKey As String, ByVal NewKey As String, ByVal Value As String)
    
    CurrentKeyValuePair.Update OldKey, NewKey, Value
    
End Sub

Private Sub ICRUD_Delete(ByVal Key As String)
    
    CurrentKeyValuePair.Delete Key
    
End Sub

Private Sub UpdateCurrentKeyValuePairReference(ByVal ToGroup As Group)
    Set CurrentKeyValuePair = FindCollectionFromGroup(ToGroup)
End Sub

'@Entry
'@TODO: Need to remove IsForListing and update our Lambda Properties code. Discuss with Erik.
Public Function CreateLambdaMetadata(ByVal LambdaParts As Variant _
                                     , ByVal LetParts As Variant _
                                      , ByVal DefaultName As String _
                                       , ByVal OldLambdaName As String _
                                       , Optional ByVal IsForListing As Boolean = False) As Metadata

    Logger.Log TRACE_LOG, "Enter Metadata.CreateLambdaMetadata"
    Dim CurrentMetadata As Metadata
    Set CurrentMetadata = New Metadata
    
    Dim CurrentLetVarValue As String
    
    ' Lambda Name
    CurrentLetVarValue = FindMetadataValue(LetParts, LAMBDA_NAME_LET_VAR)
    If DefaultName <> vbNullString Then
        CurrentMetadata.LambdaName = DefaultName
    ElseIf OldLambdaName <> vbNullString Then
        CurrentMetadata.LambdaName = OldLambdaName
    Else
        CurrentMetadata.LambdaName = Sanitize(CurrentLetVarValue, QUOTES)
    End If
    
    ' Command Name
    CurrentLetVarValue = FindMetadataValue(LetParts, COMMAND_NAME_LET_VAR)
    CurrentMetadata.CommandName = Sanitize(CurrentLetVarValue, QUOTES)
    
    ' Description
    CurrentLetVarValue = FindMetadataValue(LetParts, DESCRIPTION_LET_VAR)
    CurrentMetadata.Description = Sanitize(CurrentLetVarValue, QUOTES)
    
    ' Source Name
    CurrentLetVarValue = FindMetadataValue(LetParts, SOURCE_NAME_LET_VAR)
    CurrentMetadata.SourceName = Sanitize(CurrentLetVarValue, QUOTES)
    
    ' Gist URL
    CurrentLetVarValue = FindMetadataValue(LetParts, GIST_URL_LET_VAR)
    CurrentMetadata.GistURL = Sanitize(CurrentLetVarValue, QUOTES)
    
    Dim CrudOperator As ICRUD
    Set CrudOperator = CurrentMetadata
    
    ' Lambda Parameters
    CrudOperator.ApplyOperationOnGroup LAMBDA_PARAMETERS
    CurrentLetVarValue = FindMetadataValue(LetParts, PARAMETERS_LET_VAR)
    If Text.Trim(CurrentLetVarValue) = vbNullString Or IsForListing Then
        CreateMapperForParametersFromSignature CrudOperator, LambdaParts, CurrentLetVarValue
    Else
        CreateMapperFromArrayConstant CrudOperator, CurrentLetVarValue
    End If
    
    ' Lambda Dependencies
    CrudOperator.ApplyOperationOnGroup LAMBDA_Dependencies
    CurrentLetVarValue = FindMetadataValue(LetParts, DEPENDENCIES_LET_VAR)
    CreateMapperFromArrayConstant CrudOperator, CurrentLetVarValue
    
    ' Custom Properties
    CrudOperator.ApplyOperationOnGroup CUSTOM_PROPERTIES
    CurrentLetVarValue = FindMetadataValue(LetParts, CUSTOMPROPERTIES_LET_VAR)
    CreateMapperFromArrayConstant CrudOperator, CurrentLetVarValue
    
    Set CreateLambdaMetadata = CurrentMetadata
    Logger.Log TRACE_LOG, "Exit Metadata.CreateLambdaMetadata"

End Function

Private Function Sanitize(ByVal CurrentLetVarName As String, ByVal RemoveTextFromBothEnd As String) As String

    Logger.Log TRACE_LOG, "Enter Metadata.Sanitize"
    CurrentLetVarName = VBA.LTrim$(CurrentLetVarName)
    CurrentLetVarName = VBA.RTrim$(CurrentLetVarName)
    Sanitize = Text.RemoveFromBothEndIfPresent(CurrentLetVarName, RemoveTextFromBothEnd, IGNORE_CASE)
    Sanitize = modUtility.RemoveEscapeQuotes(Sanitize)
    Logger.Log TRACE_LOG, "Exit Metadata.Sanitize"

End Function

'@ImpureFunction
Private Sub CreateMapperForParametersFromSignature(ByRef AddToCrudOperator As ICRUD, _
                                                   ByVal LambdaParts As Variant, _
                                                   ByVal ParamMetadataValue As String)
                                                   
    Logger.Log TRACE_LOG, "Enter Metadata.CreateMapperForParametersFromSignature"
    
    If Not IsArrayAllocated(LambdaParts) Then Exit Sub
    
    Dim Map As Collection
    If ParamMetadataValue <> vbNullString Then
        Set Map = ArrayToCollection(SplitArrayConstantTo2DArray(ParamMetadataValue))
    Else
        Set Map = New Collection
    End If
    
    Dim ParamName As String
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(LambdaParts, 2)
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(LambdaParts, 1) To UBound(LambdaParts, 1) - 1
        
        ParamName = LambdaParts(CurrentRowIndex, FirstColumnIndex)
        ParamName = RemoveEscapeQuotes(ParamName)
        Dim Value As Variant
        
        If IsExistInCollection(Map, ParamName) Then
            Value = Map.Item(ParamName)
        Else
            Value = vbNullString
        End If
        
        AddToCrudOperator.AddNew ParamName, Value
        
    Next CurrentRowIndex
    Logger.Log TRACE_LOG, "Exit Metadata.CreateMapperForParametersFromSignature"
    
End Sub

'@ImpureFunction
Private Sub CreateMapperFromArrayConstant(ByRef AddToCrudOperator As ICRUD, ByVal CurrentLetValue As String)
    
    Logger.Log TRACE_LOG, "Enter Metadata.CreateMapperFromArrayConstant"
    ' Check if the CurrentLetValue is not empty
    If Application.WorksheetFunction.Trim(CurrentLetValue) = vbNullString Then Exit Sub
    
    Dim SplittedArrayConstant As Variant
    ' Convert the CurrentLetValue into a 2D array using SplitArrayConstantTo2DArray function
    SplittedArrayConstant = SplitArrayConstantTo2DArray(CurrentLetValue)
    Dim FirstColumnIndex As Long
    FirstColumnIndex = LBound(SplittedArrayConstant, 2)
    Dim CurrentRowIndex As Long
    Dim Key As String
    Dim Value As String
    For CurrentRowIndex = LBound(SplittedArrayConstant, 1) To UBound(SplittedArrayConstant, 1)
        ' Extract Key and Value from the 2D array, and remove quotes if present
        Key = SplittedArrayConstant(CurrentRowIndex, FirstColumnIndex)
        Key = Text.RemoveFromBothEndIfPresent(Key, QUOTES, IGNORE_CASE)
        If FirstColumnIndex <> UBound(SplittedArrayConstant, 2) Then
            Value = SplittedArrayConstant(CurrentRowIndex, FirstColumnIndex + 1)
            Value = Text.RemoveFromBothEndIfPresent(Value, QUOTES, IGNORE_CASE)
        End If
        ' Add Key-Value pair to the CRUD operator
        AddToCrudOperator.AddNew Key, Value
    Next CurrentRowIndex
    
    ' Log function exit
    Logger.Log TRACE_LOG, "Exit Metadata.CreateMapperFromArrayConstant"

End Sub

Private Function FindCollectionFromGroup(ByVal ToGroup As Group) As ICRUD
    
    Logger.Log TRACE_LOG, "Enter Metadata.FindCollectionFromGroup"
    Select Case ToGroup                          ' Selecting group to perform operations on
        Case Group.LAMBDA_PARAMETERS
            Set FindCollectionFromGroup = this.Parameters
        Case Group.LAMBDA_Dependencies
            Set FindCollectionFromGroup = this.Dependencies
        Case Group.CUSTOM_PROPERTIES
            Set FindCollectionFromGroup = this.CustomProperties
        Case Else
            Set FindCollectionFromGroup = Nothing
    End Select
    Logger.Log TRACE_LOG, "Exit Metadata.FindCollectionFromGroup"
    
End Function

Public Function GetParameters(Optional ByVal InitialSpace As String = THREE_SPACE _
                              , Optional ByVal ShowBlankValuesKey As Boolean = True) As String
    
    GetParameters = GetDynamicData(this.Parameters, KEY_VALUE_SEPARATOR, ShowBlankValuesKey, InitialSpace) ' Retrieving parameters using dynamic data method
       
End Function

Public Function GetParametersForLambdaListing() As Variant
    
    Dim Data As Variant
    Data = this.Parameters.ReadAll()
    
    Dim Result As String
    If IsArray(Data) Then
        
        Dim FirstColumnIndex  As Long
        FirstColumnIndex = LBound(Data, 2)
        Dim RowIndex As Long
        For RowIndex = LBound(Data, 1) To UBound(Data, 1)
            
            Dim DescriptionPart As String
            
            DescriptionPart = Data(RowIndex, FirstColumnIndex + 1)
            If DescriptionPart <> vbNullString Then
                DescriptionPart = ONE_SPACE & FIRST_PARENTHESIS_OPEN & DescriptionPart & FIRST_PARENTHESIS_CLOSE
            End If
            
            Result = Result & Data(RowIndex, FirstColumnIndex) & DescriptionPart & COMMA & ONE_SPACE
                     
        Next RowIndex
        
        Result = Text.RemoveFromEndIfPresent(Result, COMMA & ONE_SPACE)
        
    End If
    
    GetParametersForLambdaListing = Result
    
End Function

Private Function GetNameManagerComment() As String
    
    Logger.Log TRACE_LOG, "Enter Metadata.GetNameManagerComment"
    Dim Result As String

    ' Create comment for NameManager based on metadata
    If this.Description <> vbNullString Then
        Result = Text.PadIfNotPresent(modUtility.RemoveEscapeQuotes(this.Description), ".", FROM_end) & Space(2)
    End If
    Dim ParameterData As String
    ParameterData = GetParameters(HASH_SIGN, False)
    ParameterData = VBA.Replace(ParameterData, NEW_LINE & HASH_SIGN, "; ")
    If ParameterData <> vbNullString Then
        ParameterData = FIRST_PARENTHESIS_OPEN & Text.RemoveFromStartIfPresent(ParameterData, HASH_SIGN) _
                        & FIRST_PARENTHESIS_CLOSE
    End If
    GetNameManagerComment = Result & ParameterData
    Logger.Log TRACE_LOG, "Exit Metadata.GetNameManagerComment"
    
End Function

Public Function GetDependencies(Optional ByVal InitialSpace As String = THREE_SPACE _
                                , Optional ByVal ShowBlankValuesKey As Boolean = True) As String
    
    GetDependencies = GetDynamicData(this.Dependencies, KEY_VALUE_SEPARATOR, ShowBlankValuesKey, InitialSpace) ' Retrieving dependencies using dynamic data method
       
End Function

Public Function GetSources() As String
    
    Logger.Log TRACE_LOG, "Enter Metadata.GetSources"
    GetSources = THREE_SPACE & "Source Name" & KEY_VALUE_SEPARATOR & this.SourceName & NEW_LINE & _
                 THREE_SPACE & "Gist URL" & KEY_VALUE_SEPARATOR & this.GistURL ' Getting source information
    Logger.Log TRACE_LOG, "Exit Metadata.GetSources"
       
End Function

Public Function GetCustomProperties() As String
    
    GetCustomProperties = GetDynamicData(this.CustomProperties, ":  ", True, vbNullString) ' Retrieving custom properties using dynamic data method
       
End Function

Private Function GetDynamicData(ByVal FromCrudOperator As ICRUD _
                                , ByVal Separator As String _
                                 , ByVal ShowBlankValuesKey As Boolean _
                                  , Optional ByVal InitialSpace As String = THREE_SPACE) As String
    
    GetDynamicData = ConcatenateKeyValuePair(FromCrudOperator.ReadAll, Separator, ShowBlankValuesKey, InitialSpace) ' Concatenate Key-Value pair for the CRUD operation data
   
End Function

Private Function ConcatenateKeyValuePair(ByVal GivenArray As Variant _
                                         , ByVal Separator As String _
                                          , ByVal ShowBlankValuesKey As Boolean _
                                           , Optional ByVal InitialSpace As String = THREE_SPACE) As String
    
    Logger.Log TRACE_LOG, "Enter Metadata.ConcatenateKeyValuePair"
    ' If array is not valid, return empty string
    If Not IsArray(GivenArray) Then
        ConcatenateKeyValuePair = vbNullString
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Metadata.ConcatenateKeyValuePair"
        Exit Function
    End If
    
    ' Creating string of key-value pair from given array
    Dim Result As String
    Dim FirstColumnIndex As Long
    FirstColumnIndex = LBound(GivenArray, 2)
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(GivenArray, 1) To UBound(GivenArray, 1)
        Dim CurrentValue As Variant
        CurrentValue = GivenArray(CurrentRowIndex, FirstColumnIndex + 1)
        If CurrentValue <> vbNullString Or ShowBlankValuesKey Then
            Result = Result & NEW_LINE & _
                     InitialSpace & GivenArray(CurrentRowIndex, FirstColumnIndex) & Separator & _
                     CurrentValue
        End If
    Next CurrentRowIndex
    ConcatenateKeyValuePair = Text.RemoveFromStartIfPresent(Result, NEW_LINE, IGNORE_CASE) ' Final key-value pair string after concatenation
    Logger.Log TRACE_LOG, "Exit Metadata.ConcatenateKeyValuePair"
   
End Function

Public Function GetArrayConstant(ByVal ForGroup As Group _
                                 , Optional ByVal ShowBlankValuesKey As Boolean = True) As String
    
    Logger.Log TRACE_LOG, "Enter Metadata.GetArrayConstant"
    Dim CurrentGroup As ICRUD
    Set CurrentGroup = FindCollectionFromGroup(ForGroup) ' Finding group for operation
    Dim KeyValues As Variant
    KeyValues = CurrentGroup.ReadAll
    
    ' If key values are not valid array, return empty string
    If Not IsArray(KeyValues) Then
        GetArrayConstant = vbNullString
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword Metadata.GetArrayConstant"
        Exit Function
    End If
    
    ' Creating array constant for the group
    Dim FirstColumnIndex As Long
    FirstColumnIndex = LBound(KeyValues, 2)
    Dim CurrentRowIndex As Long
    Dim ArrayConstant As String
    ArrayConstant = LEFT_BRACE & NEW_LINE
    For CurrentRowIndex = LBound(KeyValues, 1) To UBound(KeyValues, 1)
        
        Dim CurrentValue As Variant
        CurrentValue = KeyValues(CurrentRowIndex, FirstColumnIndex + 1)
        If Not IsNumeric(CurrentValue) And CurrentValue <> vbNullString Then
            CurrentValue = QUOTES & modUtility.EscapeQuotes(CStr(CurrentValue)) & QUOTES
        End If
        
        If CurrentValue <> vbNullString Or ShowBlankValuesKey Then
            ArrayConstant = ArrayConstant & THREE_SPACE & THREE_SPACE & QUOTES _
                            & KeyValues(CurrentRowIndex, FirstColumnIndex) & QUOTES & _
                            ARRAY_CONST_COLUMN_SEPARATOR & CurrentValue _
                            & ARRAY_CONST_ROW_SEPARATOR & NEW_LINE
        End If
    Next CurrentRowIndex
    ArrayConstant = Text.RemoveFromEndIfPresent(ArrayConstant, ARRAY_CONST_ROW_SEPARATOR & NEW_LINE)
    If ArrayConstant = LEFT_BRACE & NEW_LINE Then
        ArrayConstant = vbNullString
    Else
        ArrayConstant = ArrayConstant & NEW_LINE & THREE_SPACE & RIGHT_BRACE
    End If
    GetArrayConstant = ArrayConstant
    Logger.Log DEBUG_LOG, "Group : " & ForGroup & " Const : " & NEW_LINE & ArrayConstant
    Logger.Log TRACE_LOG, "Exit Metadata.GetArrayConstant"
   
End Function

