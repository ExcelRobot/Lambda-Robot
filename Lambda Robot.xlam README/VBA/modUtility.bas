Attribute VB_Name = "modUtility"
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed
'@Folder "Utility"
'@IgnoreModule ImplicitActiveSheetReference, SuperfluousAnnotationArgument, UnrecognizedAnnotation, ProcedureNotUsed, UnassignedVariableUsage
' @Folder "Lambda.Editor.Utility"
Option Explicit
Option Private Module

#If VBA7 Then                                    ' Excel 2010 or later
    
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
    
#Else                                            ' Excel 2007 or earlier
    
    Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
    
#End If

Public Function HasDynamicFormula(ByVal SelectionRange As Range) As Boolean
    
    ' Check if the selected range contains a dynamic formula (spill range).
    On Error Resume Next
    HasDynamicFormula = SelectionRange.Cells(1).HasSpill
    On Error GoTo 0
    
End Function

Public Function IsSubSet(ByVal ParentSet As Range, ByVal ChildSet As Range) As Boolean
    
    ' Check if ChildSet is a subset of ParentSet.
    Dim CommonSection As Range
    Set CommonSection = FindIntersection(ParentSet, ChildSet)
    If IsNothing(CommonSection) Then
        IsSubSet = False
    ElseIf IsNotNothing(ChildSet) Then
        IsSubSet = (CommonSection.Address = ChildSet.Address)
    End If
    
End Function

Public Function ExtractOneColumnOfA2DArray(ByVal GivenArray As Variant, ByVal ColumnIndex As Long) As Variant
    
    ' Extract a single column from a 2D array and return as a 1D array.
    Dim Result As Variant
    ReDim Result(LBound(GivenArray, 1) To UBound(GivenArray, 1), 1 To 1)
    Dim CurrentIndex As Long

    For CurrentIndex = LBound(GivenArray, 1) To UBound(GivenArray, 1)
        Result(CurrentIndex, 1) = GivenArray(CurrentIndex, ColumnIndex)
    Next CurrentIndex
    ExtractOneColumnOfA2DArray = Result
    
End Function

' This function checks if GivenCell is an input cell based on color values.
Public Function IsInputCell(ByVal GivenCell As Range, ByVal DependencySearchInRegion As Range) As Boolean
    
    ' Check if GivenCell is an input cell based on color values.
    Dim Result As Boolean
    If (GivenCell.Interior.Color = INPUT_CELL_BACKGROUND_COLOR) Then
        Result = True
    ElseIf GivenCell.Font.Color = INPUT_CELL_FONT_COLOR Then
        Result = True
    ElseIf IsNotNothing(DependencySearchInRegion) Then
        Result = IsNotInside(DependencySearchInRegion, GivenCell)
    End If
    
    IsInputCell = Result
    
End Function

Public Function IsExistInCollection(ByVal GivenCollection As Collection, ByVal Key As String) As Boolean
    
    ' Check if the given Key exists in the Collection.
    On Error GoTo NotExist
    Dim Item  As Variant
    If IsObject(GivenCollection.Item(Key)) Then
        Set Item = GivenCollection.Item(Key)
    Else
        Item = GivenCollection.Item(Key)
    End If
    IsExistInCollection = True
    Exit Function
    
NotExist:
    IsExistInCollection = False
    On Error GoTo 0
    
End Function

Public Function FindLetVarName(ByVal FromRange As Range) As String
    
    ' Find a suitable name for a variable based on the input FromRange.
    If FromRange.Cells.Count > 1 And FromRange.Cells(1).Address = "$A$1" Then
        FindLetVarName = FromRange.Worksheet.name
        Exit Function
    End If
    
    Dim CellAbove As Range
    Dim CellTwoAbove As Range
    Dim CellToLeft As Range

    If IsNotNothing(FromRange) Then Set FromRange = FromRange.Cells(1)
    If FromRange.Cells(1).Row > 1 Then Set CellAbove = FromRange.Offset(-1).Cells(1)
    If FromRange.Cells(1).Row > 2 Then Set CellTwoAbove = FromRange.Offset(-2).Cells(1)
    If FromRange.Cells(1).Column > 1 Then
        Set CellToLeft = FromRange.Offset(0, -1).Cells(1)
    Else
        Set CellToLeft = FromRange.Cells(1)
    End If

    ' Look above for a suitable name
    If IsNotNothing(CellAbove) Then
        If GetCellValueIfErrorNullString(CellAbove) = vbNullString Then
            If IsNotNothing(CellTwoAbove) Then
                If IsProbableLetVarName(CellTwoAbove) Then
                    FindLetVarName = CellTwoAbove.Value
                    Exit Function
                End If
            End If
        ElseIf IsProbableLetVarName(CellAbove) Then
            FindLetVarName = CellAbove.Value
            Exit Function
        End If
    End If

    ' Scan left until a non-blank cell is found or we reach column A
    Do While IsCellBlank(CellToLeft) And CellToLeft.Column > 1 And Not IsCellHidden(CellToLeft)
        Set CellToLeft = CellToLeft.Offset(0, -1).Cells(1)
    Loop

    ' Check for a suitable name
    If IsProbableLetVarName(CellToLeft) Then
        FindLetVarName = CellToLeft.Value
    End If
    
End Function

Private Function IsCellBlank(ByVal GivenCell As Range) As Boolean
    ' Check if the given cell is blank.
    IsCellBlank = (GetCellValueIfErrorNullString(GivenCell) = vbNullString)
End Function

Public Function GetCellValueIfErrorNullString(ByVal GivenCell As Range) As String
    
    ' Get the cell value or return an empty string if an error is encountered.
    Dim Result As String
    If IsError(GivenCell.Value) Then
        Result = vbNullString
    Else
        Result = GivenCell.Value
    End If
    
    GetCellValueIfErrorNullString = Result
    
End Function

Private Function IsProbableLetVarName(ByVal CurrentCell As Range) As Boolean
    
    ' Check if CurrentCell is a probable variable name.
    Dim Result As Boolean
    If Application.WorksheetFunction.Trim(CurrentCell.Value) = vbNullString Then
        Result = False
    ElseIf HasDynamicFormula(CurrentCell) Or CurrentCell.HasFormula Then
        Result = False
    ElseIf TypeName(CurrentCell.Value) <> "String" Then
        Result = False
    ElseIf modUtility.IsInsideTable(CurrentCell) Then
        Dim ActiveTable As ListObject
        Set ActiveTable = GetTableFromRange(CurrentCell)
        ' Condition for not in the header or not.
        Result = IsNotNothing(FindIntersection(ActiveTable.HeaderRowRange, CurrentCell))
    ElseIf IsInsideNamedRange(CurrentCell) Then
        Result = False
    ElseIf IsInputCell(CurrentCell, Nothing) Then
        Result = False
    Else
        Result = True
    End If
    
    IsProbableLetVarName = Result
    
End Function

Public Function IsCellHidden(ByVal CurrentCell As Range) As Boolean
    ' Check if the CurrentCell or its entire row/column is hidden.
    IsCellHidden = (CurrentCell.EntireColumn.Hidden Or CurrentCell.EntireRow.Hidden)
End Function

Public Function ConvertVarNameToSentence(VarName As String) As String
    
    Dim Sentence As String
    Sentence = Replace(VarName, DOT, ONE_SPACE)
    Sentence = Replace(Sentence, UNDER_SCORE, ONE_SPACE)
    Sentence = ReplaceLineBreak(Sentence, ONE_SPACE)
    Sentence = ConcatenateCollection(Text.SplitDigitAndNonDigit(Sentence), ONE_SPACE)
    Dim Words As Variant
    Words = Split(Trim$(Sentence), ONE_SPACE)
    Sentence = vbNullString
    
    Dim Word As Variant
    For Each Word In Words
        Sentence = Sentence & ONE_SPACE & PutSpaceOnLowerCaseToUpperCaseTransition(Word)
    Next Word
    
    Words = Split(Trim$(Sentence), ONE_SPACE)
    Sentence = vbNullString
    
    For Each Word In Words
        Sentence = Sentence & ONE_SPACE & PutSpaceBeforeLastCapsFromStart(Word)
    Next Word
    
    ConvertVarNameToSentence = Trim$(Sentence)
    
End Function

Public Function MakeValidLetVarName(ByVal GivenLetVarName As String _
                                    , NamingConv As VarNamingStyle) As String
    ' Make the given LetVarName a valid variable name by removing invalid characters.
    MakeValidLetVarName = MakeValidName(GivenLetVarName, NamingConv)
End Function

Public Function ConvertToValidLetVarName(ByVal GivenName As String) As String
    
    Dim ValidName As String
    ' Replace Newline with space.
    ValidName = ReplaceLineBreak(Trim$(GivenName), ONE_SPACE)
    
    ValidName = ReplacePlaceHolders(ValidName)
    
    ' Remove Invalid char but keep space.
    ValidName = RemoveInvalidCharcters(ValidName, True)
    
    ' Replace dots with underscores in the name.
    ValidName = VBA.Replace(ValidName, DOT, UNDER_SCORE)
    
    ' Convert To proper sentence form.
    ValidName = Replace(Text.Trim((ValidName)), ONE_SPACE, vbNullString)
    
    ' If the name is a range reference, split it and add underscores.
    If IsRangeReference(ValidName) Then
        Dim ColRefAndRowRef As Collection
        Set ColRefAndRowRef = Text.SplitDigitAndNonDigit(ValidName)
        ValidName = ColRefAndRowRef.Item(1) & UNDER_SCORE & ColRefAndRowRef.Item(2)
    End If
    
    ' Limit the length of the name to MAX_ALLOWED_LENGTH.
    If Len(ValidName) > modSharedConstant.MAX_ALLOWED_LET_STEP_NAME_LENGTH Then
        ValidName = Left$(ValidName, modSharedConstant.MAX_ALLOWED_LET_STEP_NAME_LENGTH)
    End If
    
    ConvertToValidLetVarName = ValidName
    
End Function

Public Function MakeValidName(ByVal GivenInvalidName As String _
                               , NamingConv As VarNamingStyle) As String
    
    Dim ValidName As String
    ' Replace Newline with space.
    ValidName = ReplaceLineBreak(Trim$(GivenInvalidName), ONE_SPACE)
    
    ValidName = ReplacePlaceHolders(ValidName)
    
    ' Remove Invalid char but keep space.
    ValidName = RemoveInvalidCharcters(ValidName, True)
    
    ' Replace dots with underscores in the name.
    ValidName = VBA.Replace(ValidName, DOT, UNDER_SCORE)
    
    ' Convert To proper sentence form.
    ValidName = Text.Trim(ConvertVarNameToSentence(ValidName))
    
    Select Case NamingConv
        Case VarNamingStyle.CAMEL_CASE
            ValidName = ConvertToCamelCase(ValidName)
            
        Case VarNamingStyle.PASCAL_CASE
            ValidName = ConvertToPascalCase(ValidName)
            
        Case VarNamingStyle.SNAKE_CASE
            ValidName = LCase$(Replace(ValidName, ONE_SPACE, UNDER_SCORE))
        
        Case Else
            Err.Raise 5, "Make Valid Name", "Invalid Naming Convention"
            
    End Select
    
    ' If the name is a range reference, split it and add underscores.
    If IsRangeReference(ValidName) Then
        Dim ColRefAndRowRef As Collection
        Set ColRefAndRowRef = Text.SplitDigitAndNonDigit(ValidName)
        ValidName = ColRefAndRowRef.Item(1) & UNDER_SCORE & ColRefAndRowRef.Item(2)
    End If
    
    ' Limit the length of the name to MAX_ALLOWED_LENGTH.
    If Len(ValidName) > modSharedConstant.MAX_ALLOWED_LET_STEP_NAME_LENGTH Then
        ValidName = Left$(ValidName, modSharedConstant.MAX_ALLOWED_LET_STEP_NAME_LENGTH)
    End If
    
    MakeValidName = ValidName
    
End Function

Public Function IsValidLetVarName(ByVal NameToCheck As String) As Boolean
    IsValidLetVarName = (NameToCheck = RemoveInvalidCharcters(NameToCheck, False))
End Function

Public Function IsValidDefinedName(ByVal NameToCheck As String) As Boolean
    IsValidDefinedName = (NameToCheck = RemoveInvalidCharcters(NameToCheck, False))
End Function

Private Function ReplaceLineBreak(ByVal Text As String, ReplaceWith As String) As String
    
    Dim ReplacedText As String
    ReplacedText = Replace(Text, vbNewLine, ReplaceWith)
    ReplacedText = Replace(ReplacedText, Chr$(10), ReplaceWith)
    ReplacedText = Replace(ReplacedText, Chr$(13), ReplaceWith)
    ReplaceLineBreak = ReplacedText
    
End Function

'  This just replace space with VBNullstring and convert first char of each word to upper case except first one.
Private Function ConvertToCamelCase(ByVal VarName As String) As String
    
    Dim ValidName As String
    ValidName = Text.Trim(CapitalizeFirstCharOfEachWord(VarName))
    If Text.Contains(ValidName, ONE_SPACE) Then
        
        If Not IsAllCaps(Text.BeforeDelimiter(ValidName, ONE_SPACE)) Then
            ValidName = LCase(Text.BeforeDelimiter(ValidName, ONE_SPACE)) & ONE_SPACE _
                        & ConvertToProperCaseOfEachWord( _
                        Text.AfterDelimiter(ValidName, ONE_SPACE))
        End If
        
    Else
        If Not IsAllCaps(ValidName) Then
            ValidName = LCase(ValidName)
        End If
    End If
    ValidName = Replace(ValidName, ONE_SPACE, vbNullString)
    
    ConvertToCamelCase = ValidName
    
End Function

'  This just replace space with VBNullstring and convert first char of each word to upper case
Private Function ConvertToPascalCase(ByVal VarName As String) As String
    
    Dim ValidName As String
    ValidName = Text.Trim(CapitalizeFirstCharOfEachWord(VarName))
    ValidName = ConvertToProperCaseOfEachWord(ValidName)
    ValidName = Replace(ValidName, ONE_SPACE, vbNullString)
    ConvertToPascalCase = ValidName
    
End Function

' Check if Upper case text and input text is equal or not.
Public Function IsAllCaps(Text As String) As Boolean
    IsAllCaps = (UCase$(Text) = Text)
End Function

' Check if the given reference is a valid range reference.
Public Function IsRangeReference(ByVal GivenRef As String) As Boolean

    ' Use ConvertFormula to try converting the reference to R1C1 notation.
    If IsError(Application.ConvertFormula(EQUAL_SIGN & GivenRef, xlA1, xlR1C1, , Range("A1"))) Then
        IsRangeReference = False
    Else
        ' Check if the converted R1C1 notation is different from the original reference.
        IsRangeReference = (UCase$(Application.ConvertFormula(EQUAL_SIGN & GivenRef _
                                                              , xlA1, xlR1C1 _
                                                                     , , Range("A1"))) <> UCase$(EQUAL_SIGN & GivenRef))
    End If

End Function

' Convert To proper case only if the entire word is not Upper Case
' Example ConvertToProperCaseOfEachWord("USA is a deveLoped Coutry") >> USA Is A Developed Coutry
Public Function ConvertToProperCaseOfEachWord(ByVal Sentence As String) As String
    
    Dim Words As Variant
    Words = Split(Sentence, ONE_SPACE)
    Dim CurrentIndex As Long
    For CurrentIndex = LBound(Words) To UBound(Words)
        Dim CurrentWord As String
        CurrentWord = Words(CurrentIndex)
        If IsAllCaps(CurrentWord) Then
            Words(CurrentIndex) = CurrentWord
        Else
            Words(CurrentIndex) = Text.Proper(CurrentWord)
        End If
    Next CurrentIndex
    ConvertToProperCaseOfEachWord = Join(Words, ONE_SPACE)
    
End Function

Public Function CapitalizeFirstCharOfEachWord(ByVal GivenName As String) As String

    Dim CurrentCharIndex As Long
    Dim CurrentChar As String

    ' Loop through each character in the given name.
    For CurrentCharIndex = 1 To Len(GivenName)
        CurrentChar = Mid$(GivenName, CurrentCharIndex, 1)

        ' If the current character is a space (ASCII code 32),
        ' capitalize the first char follows it.
        Const SPACE_ASCII_VALUE As Long = 32
        If Asc(CurrentChar) = SPACE_ASCII_VALUE Then
            If CurrentCharIndex < Len(GivenName) Then
                GivenName = CapitalizeNthCharacter(GivenName, CurrentCharIndex + 1)
            End If
        End If
    Next CurrentCharIndex
    
    CapitalizeFirstCharOfEachWord = CapitalizeNthCharacter(GivenName, 1)
    
End Function

' Capitalize first character of each word in the given name that follows a line break.
Public Function CapitalizeFirstCharOfEachWordAfterLineBreak(ByVal GivenName As String) As String

    Dim CurrentCharIndex As Long
    Dim CurrentChar As String

    ' Loop through each character in the given name.
    For CurrentCharIndex = 1 To Len(GivenName)
        CurrentChar = Mid$(GivenName, CurrentCharIndex, 1)

        ' If the current character is a line break (ASCII code 10),
        ' capitalize the first char that follows it.
        Const LINE_BREAK_ASC_CODE As Long = 10
        If Asc(CurrentChar) = LINE_BREAK_ASC_CODE Then
            If CurrentCharIndex < Len(GivenName) Then
                GivenName = CapitalizeNthCharacter(GivenName, CurrentCharIndex + 1)
            End If
        End If
    Next CurrentCharIndex

    ' Return the modified name.
    CapitalizeFirstCharOfEachWordAfterLineBreak = GivenName

End Function

' Capitalize the Nth character in the given text.
Public Function CapitalizeNthCharacter(ByRef GivenText As String, ByVal NthIndex As Long) As String

    Dim TextLength As Long
    TextLength = Len(GivenText)

    ' Check if the text is empty.
    If TextLength = 0 Then
        CapitalizeNthCharacter = GivenText
        Exit Function
    End If

    ' Check if the NthIndex is valid.
    If NthIndex > TextLength Then
        Err.Raise 13, "Type Mismatch", "NthIndex needs to be less than text length"
    End If

    ' Capitalize the Nth character based on its position.
    If NthIndex = TextLength Then
        CapitalizeNthCharacter = Left$(GivenText, TextLength - 1) & UCase$(Right$(GivenText, 1))
    ElseIf NthIndex = 1 Then
        CapitalizeNthCharacter = UCase$(Left$(GivenText, 1)) & Right$(GivenText, TextLength - 1)
    Else
        CapitalizeNthCharacter = Left$(GivenText, NthIndex - 1) _
                                 & UCase$(Mid$(GivenText, NthIndex, 1)) _
                                 & Text.SubString(GivenText, NthIndex + 1)
    End If

End Function

' Replace specific placeholders with their corresponding values.
Public Function ReplacePlaceHolders(ByVal GivenName As String) As String

    Dim PlaceHolders As Variant
    PlaceHolders = Array("%", HASH_SIGN, "&", "<", ">", EQUAL_SIGN)

    Dim ReplaceWiths As Variant
    ReplaceWiths = Array("Percent", "Number", "And", "LessThan", "GreaterThan", "Equals")

    Dim CurrentIndex As Long

    ' Loop through each placeholder and replace it with the corresponding value.
    For CurrentIndex = LBound(PlaceHolders) To UBound(PlaceHolders)
        GivenName = Replace(GivenName, PlaceHolders(CurrentIndex), ReplaceWiths(CurrentIndex))
    Next CurrentIndex

    ' Return the modified name.
    ReplacePlaceHolders = GivenName

End Function

' Remove invalid characters from the given name.
Public Function RemoveInvalidCharcters(ByVal GivenName As String, KeepSpace As Boolean) As String

    Dim Output As String
    Dim CurrentCharIndex As Long
    Dim CurrentChar As String

    ' Loop through each character in the given name.
    For CurrentCharIndex = 1 To Len(GivenName)
        CurrentChar = Mid$(GivenName, CurrentCharIndex, 1)
        ' Check if the current character is a valid first character for the name.
        If IsValidFirstChar(CurrentChar) Then
            Output = CurrentChar
            Exit For
        End If
    Next CurrentCharIndex

    ' If the given name is not empty and there are characters after the first valid character,
    ' update the name accordingly. Otherwise, set the name to an empty string.
    If Len(GivenName) <> CurrentCharIndex And Len(GivenName) > CurrentCharIndex Then
        GivenName = Right$(GivenName, Len(GivenName) - CurrentCharIndex)
    Else
        GivenName = vbNullString
    End If

    ' Return the updated name with valid characters.
    RemoveInvalidCharcters = Output & GetValidCharForSecondToOnward(GivenName, KeepSpace)

End Function

' Check if the given character is a valid first character for the name.
Public Function IsValidFirstChar(ByVal GivenChar As String) As Boolean
    
    Static InvalidFirstChars As Collection
    If InvalidFirstChars Is Nothing Then
        Set InvalidFirstChars = New Collection
        With InvalidFirstChars
            AddCharsToColl InvalidFirstChars, 1, 64
            .Add 91, CStr(91)
            AddCharsToColl InvalidFirstChars, 93, 94
            .Add 96, CStr(96)
            AddCharsToColl InvalidFirstChars, 123, 130
            .Add 132, CStr(132)
            .Add 136, CStr(136)
            .Add 139, CStr(139)
            .Add 141, CStr(141)
            AddCharsToColl InvalidFirstChars, 143, 144
            .Add 149, CStr(149)
            .Add 152, CStr(152)
            .Add 155, CStr(155)
            .Add 157, CStr(157)
            .Add 160, CStr(160)
            AddCharsToColl InvalidFirstChars, 162, 163
            AddCharsToColl InvalidFirstChars, 165, 166
            .Add 169, CStr(169)
            AddCharsToColl InvalidFirstChars, 171, 172
            .Add 174, CStr(174)
            .Add 187, CStr(187)
        End With
    End If
    
    IsValidFirstChar = (Not IsExistInCollection(InvalidFirstChars, CStr(Asc(GivenChar))))

End Function


' Get the valid characters from the given name starting from the second character.
Public Function GetValidCharForSecondToOnward(ByVal GivenName As String, KeepSpace As Boolean) As String

    Dim Result As String
    Dim CurrentCharIndex As Long
    Dim CurrentChar As String

    ' Loop through each character in the given name.
    For CurrentCharIndex = 1 To Len(GivenName)
        CurrentChar = Mid$(GivenName, CurrentCharIndex, 1)

        ' Check if the current character is a valid second character for the name.
        If IsValidSecondChar(CurrentChar) Or (KeepSpace And CurrentChar = ONE_SPACE) Then
            Result = Result & CurrentChar
        End If
    Next CurrentCharIndex

    ' Return the result containing valid characters.
    GetValidCharForSecondToOnward = Result

End Function

Public Function IsValidSecondChar(ByVal GivenChar As String) As Boolean
    
    Static InvalidChars As Collection
    If InvalidChars Is Nothing Then
        Set InvalidChars = New Collection
        AddCharsToColl InvalidChars, 1, 45
        InvalidChars.Add 47, CStr(47)
        AddCharsToColl InvalidChars, 58, 62
        With InvalidChars
            .Add 64, CStr(64)
            .Add 91, CStr(91)
            AddCharsToColl InvalidChars, 93, 94
            .Add 96, CStr(96)
            AddCharsToColl InvalidChars, 123, 127
            AddCharsToColl InvalidChars, 129, 130
            .Add 132, CStr(132)
            .Add 139, CStr(139)
            .Add 141, CStr(141)
            AddCharsToColl InvalidChars, 143, 144
            .Add 149, CStr(149)
            .Add 155, CStr(155)
            .Add 157, CStr(157)
            .Add 160, CStr(160)
            AddCharsToColl InvalidChars, 162, 163
            AddCharsToColl InvalidChars, 165, 166
            .Add 169, CStr(169)
            AddCharsToColl InvalidChars, 171, 172
            .Add 174, CStr(174)
            .Add 187, CStr(187)
        End With
    End If
    
    IsValidSecondChar = (Not IsExistInCollection(InvalidChars, CStr(Asc(GivenChar))))

End Function
Private Sub AddCharsToColl(ByRef ToColl As Collection, ByVal StartCodeIndex As Long, ByVal EndCodeIndex As Long)
    
    Dim CodeIndex As Long
    For CodeIndex = StartCodeIndex To EndCodeIndex
        ToColl.Add CodeIndex, CStr(CodeIndex)
    Next CodeIndex
    
End Sub

Public Function GetObjectsPropertyValue(ByVal ObjectCollection As Collection _
                                        , ByVal PropertiesName As Collection _
                                         , Optional ByVal IsHaveHeader As Boolean = True) As Variant
    ' Get the property values of objects in the ObjectCollection.

    Dim Output As Variant

    ' Check if the input arguments are valid.
    If Not IsValidObjectInput(ObjectCollection, PropertiesName) Then
        Err.Raise 13, "Wrong Input argument", "Check if you have given proper input arguments or not"
    End If

    Dim Counter As Long
    Dim PropertyName As Variant

    ' Set up the Output array based on whether there is a header row or not.
    If IsHaveHeader Then
        ReDim Output(1 To ObjectCollection.Count + 1, 1 To PropertiesName.Count)
        For Each PropertyName In PropertiesName
            Counter = Counter + 1
            Output(1, Counter) = PropertyName
        Next PropertyName
        Counter = 2
    Else
        ReDim Output(1 To ObjectCollection.Count, 1 To PropertiesName.Count)
        Counter = 1
    End If

    Dim TotalPropertyCount As Long
    TotalPropertyCount = PropertiesName.Count

    ' Retrieve the property values from each object in the collection.
    On Error GoTo HandleError
    Dim CurrentObject As Object
    For Each CurrentObject In ObjectCollection
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = 1 To TotalPropertyCount
            Output(Counter, CurrentColumnIndex) = CallByName(CurrentObject _
                                                             , PropertiesName.Item(CurrentColumnIndex), VbGet)
        Next CurrentColumnIndex
        Counter = Counter + 1
    Next CurrentObject

    GetObjectsPropertyValue = Output
    
Cleanup:
    Exit Function

HandleError:
    ' Handle specific errors that may occur during property access.
    Select Case Err.Number
        Case 450
            Err.Raise 450, "Property Access Problem", "Check if you have valid Property access or not."
        Case 438
            Err.Raise Err.Number, "Property doesn't exist", "Check if your given property exists or not. Also check spelling."
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description
    End Select
    Resume Cleanup
    ' This is only for debugging purposes.
    Resume

End Function

' @Description("This is helper function for GetObjectsPropertyValue function")
' @Dependency("No Dependency")
' @ExampleCall : IsValidObjectInput(ObjectCollection, PropertiesName)
Private Function IsValidObjectInput(ByVal ObjectCollection As Collection _
                                    , ByVal PropertiesName As Collection) As Boolean

    Dim Valid As Boolean
    If IsNothing(ObjectCollection) Then
        Valid = False
    ElseIf IsNothing(PropertiesName) Then
        Valid = False
    ElseIf ObjectCollection.Count = 0 Then
        Valid = False
    ElseIf PropertiesName.Count = 0 Then
        Valid = False
    ElseIf Not IsObject(ObjectCollection.Item(1)) Then
        Valid = False
    Else
        Valid = True
    End If
    IsValidObjectInput = Valid

End Function

' @Description("This is a function which take an array with property name in top row and value in rest of the row and create objects with those value. InstanceOfClassHavingEmptyConstructor is just new object of that class which has a method by CreatorMethodName to create a new object of the same type. possible function signature is below ")
' @Dependency("ReasonToBeInValidObjectInstance")
' @ExampleCall : CreateObjectsFromArray(PropertyNameWithValues,new ClassName,"CreateMe",3)

Public Function CreateObjectsFromArray(ByVal PropertyNameWithValues As Variant _
                                       , ByVal InstanceOfClassHavingEmptyConstructor As Object _
                                        , ByVal CreatorMethodName As String _
                                         , Optional ByVal KeyColumnIndex As Long = -1) As Collection
    
    ' Create objects from the data in PropertyNameWithValues array.

    Dim Reason As String
    ' Check if the InstanceOfClassHavingEmptyConstructor and CreatorMethodName are valid.
    Reason = ReasonToBeInValidObjectInstance(InstanceOfClassHavingEmptyConstructor, CreatorMethodName)
    If Reason <> vbNullString Then
        Err.Raise 13, "Invalid Call of CreateObjectsFromArray", Reason
        Exit Function
    ElseIf Not IsArray(PropertyNameWithValues) Then
        Err.Raise 13, "Invalid Array Data", "PropertyNameWithValues should be an array with property name at the top row"
        Exit Function
    End If

    On Error GoTo HandleError

    Dim KeyToObjectMap As Collection
    Set KeyToObjectMap = New Collection
    Dim CurrentRowIndex As Long
    Dim CurrentObject As Object
    For CurrentRowIndex = LBound(PropertyNameWithValues, 1) + 1 To UBound(PropertyNameWithValues, 1)
        ' Create a new object using the CreatorMethodName.
        Set CurrentObject = CallByName(InstanceOfClassHavingEmptyConstructor, CreatorMethodName, VbMethod)
        Dim FirstRowIndex As Long
        FirstRowIndex = LBound(PropertyNameWithValues, 1)
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = LBound(PropertyNameWithValues, 2) To UBound(PropertyNameWithValues, 2)
            Dim PropertyName As String
            PropertyName = CStr(PropertyNameWithValues(FirstRowIndex, CurrentColumnIndex))
            Dim PropertyValue As Variant
            PropertyValue = PropertyNameWithValues(CurrentRowIndex, CurrentColumnIndex)
            
            Dim ObjectPropTypeName As VbVarType
            ObjectPropTypeName = VarType(CallByName(CurrentObject, PropertyName, VbGet))
            
            
            Select Case ObjectPropTypeName
                Case VbVarType.vbString
                    CallByName CurrentObject, PropertyName, VbLet, CStr(PropertyValue)
                    
                Case VbVarType.vbBoolean
                    CallByName CurrentObject, PropertyName, VbLet, CBool(PropertyValue)
                    
                Case VbVarType.vbByte
                    CallByName CurrentObject, PropertyName, VbLet, CByte(PropertyValue)
                    
                Case VbVarType.vbCurrency
                    CallByName CurrentObject, PropertyName, VbLet, CCur(PropertyValue)
                    
                Case VbVarType.vbDate
                    CallByName CurrentObject, PropertyName, VbLet, CDate(PropertyValue)
                    
                Case VbVarType.vbDecimal
                    CallByName CurrentObject, PropertyName, VbLet, CDec(PropertyValue)
                    
                Case VbVarType.vbDouble
                    CallByName CurrentObject, PropertyName, VbLet, CDbl(PropertyValue)
                    
                Case VbVarType.vbInteger
                    CallByName CurrentObject, PropertyName, VbLet, CInt(PropertyValue)
                    
                Case VbVarType.vbLong
                    CallByName CurrentObject, PropertyName, VbLet, CLng(PropertyValue)
                    
                Case VbVarType.vbSingle
                    CallByName CurrentObject, PropertyName, VbLet, CDbl(PropertyValue)
            
                Case Else
                    CallByName CurrentObject, PropertyName, VbLet, CStr(PropertyValue)
                    
            End Select
        
        Next CurrentColumnIndex
        
        ' Add the object to the collection with or without a key, based on KeyColumnIndex.
        If KeyColumnIndex = -1 Then
            KeyToObjectMap.Add CurrentObject
        Else
            KeyToObjectMap.Add CurrentObject, CStr(PropertyNameWithValues(CurrentRowIndex, KeyColumnIndex))
        End If
        
    Next CurrentRowIndex

    Set CreateObjectsFromArray = KeyToObjectMap

Cleanup:
    Exit Function

HandleError:
    ' Handle specific errors that may occur during property access.
    Select Case Err.Number
        Case 450
            ' Logger.Log DEBUG_LOG, "Property Access Problem in " & PropertyName & ". Check If you have valid Property access or not."
            Resume Next
        Case 438
            Err.Raise Err.Number, "Property doesn't exist", "Check if your given property exists or not. Also check spelling."
        Case 451
            ' Logger.Log DEBUG_LOG, PropertyName & ONE_SPACE & Err.Description
            Resume Next
        Case Else
            ' Log the error and re-raise it.
            Logger.Log ERROR_LOG, Err.Description
            Err.Raise Err.Number, Err.Source, Err.Description
    End Select
    Resume Cleanup
    ' This is only for debugging purposes.
    Resume

End Function

Public Function IsDoubleNumber(ByVal NumberAsText As String) As Boolean
    
    If IsNumeric(NumberAsText) Then
        Dim Number As Double
        Number = CDbl(NumberAsText)
        IsDoubleNumber = True
    End If
    
End Function

Public Function IsLongNumber(ByVal NumberAsText As String) As Boolean
    
    On Error GoTo HandleError
    If IsNumeric(NumberAsText) Then
        Dim Number As Long
        Number = CLng(NumberAsText)
        IsLongNumber = (Number = NumberAsText)
    End If
    Exit Function
    
HandleError:
    Exit Function
    
End Function

' @Description("This is helper function for CreateObjectsFromArray function")
' @Dependency("No Dependency")
' @ExampleCall : ReasonToBeInValidObjectInstance(InstanceOfClassHavingEmptyConstructor, CreatorMethodName)
Private Function ReasonToBeInValidObjectInstance(ByVal InstanceOfClassHavingEmptyConstructor As Object _
                                                 , ByVal CreatorMethodName As String) As String
    
    ' Check if the InstanceOfClassHavingEmptyConstructor and CreatorMethodName are valid.

    On Error GoTo HandleError
    If IsNothing(InstanceOfClassHavingEmptyConstructor) Then
        ReasonToBeInValidObjectInstance = "InstanceOfClassHavingEmptyConstructor is nothing."
    Else
        Dim NewObject As Object
        Set NewObject = CallByName(InstanceOfClassHavingEmptyConstructor, CreatorMethodName, VbMethod)
        ' Check if the object returned by the CreatorMethodName has the same type as InstanceOfClassHavingEmptyConstructor.
        If TypeName(NewObject) <> TypeName(InstanceOfClassHavingEmptyConstructor) Then
            ReasonToBeInValidObjectInstance = "Creator method returned a different type of object than it should be."
        End If
    End If

Cleanup:
    Exit Function

HandleError:
    ' Handle specific errors that may occur during property access.
    Select Case Err.Number
        Case 450
            Err.Raise 450, "Property Access Problem", "Check If you have valid Property access or not."
        Case 438
            ReasonToBeInValidObjectInstance = "Property doesn't exist. Check if your given property exists or not. Also check spelling."
            GoTo Cleanup
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description
    End Select
    Resume Cleanup
    ' This is only for debugging purposes.
    Resume

End Function

Public Function CollectionToArray(ByVal GivenCollection As Collection) As Variant
    
    ' Convert a Collection into a 1D Variant Array.

    If GivenCollection.Count = 0 Then Exit Function
    Dim Result() As Variant
    ReDim Result(1 To GivenCollection.Count, 1 To 1)
    Dim CurrentElement As Variant
    Dim CurrentIndex As Long
    For Each CurrentElement In GivenCollection
        CurrentIndex = CurrentIndex + 1
        Result(CurrentIndex, 1) = CurrentElement
    Next CurrentElement
    CollectionToArray = Result

End Function

Public Function ConcatenateOneColumnOf2DArray(ByVal GivenArray As Variant _
                                              , ByVal ColumnIndex As Long _
                                               , Optional ByVal Delimiter As String = COMMA) As String
    
    ' Concatenate a specific column of a 2D array into a single string with the given delimiter.

    Dim StartTime As Double
    StartTime = Timer
    Dim CurrentRowIndex As Long
    Dim OutputText As String
    
    For CurrentRowIndex = LBound(GivenArray, 1) To UBound(GivenArray, 1)
        OutputText = OutputText & Delimiter & GivenArray(CurrentRowIndex, ColumnIndex)
    Next CurrentRowIndex
    
    ConcatenateOneColumnOf2DArray = Right$(OutputText, Len(OutputText) - Len(Delimiter))
    
    Logger.Log INFO_LOG, "Total Time To Concatenate : " & Timer() - StartTime _
                        & "  Total Row : " & UBound(GivenArray, 1) - LBound(GivenArray, 1) + 1

End Function

Private Sub AssignProperly(ByRef FromItem As Variant, ByRef ToItem As Variant)
    
    ' Properly assign the variant value to another variant.
    If IsObject(FromItem) Then
        Set ToItem = FromItem
    Else
        ToItem = FromItem
    End If

End Sub

Private Function IsValidArray(ByVal InputArray As Variant) As Boolean
    
    ' Check if the given variant is a valid array.
    Const ArrayIdentifier As String = "()"
    If IsEmpty(InputArray) Then
        IsValidArray = False
    ElseIf InStr(TypeName(InputArray), ArrayIdentifier) < 1 Then
        ' If the array is integer type, then TypeName returns "Integer()", so ArrayIdentifier should be present at least greater than 1st _
        ' place, otherwise, it's a broken array.
        IsValidArray = False
    Else
        IsValidArray = True
    End If

End Function

Public Function RemoveStartingSingleQuoteAndEqualSign(ByVal GivenText As String) As String
    ' Remove a single quote and an equal sign from the start of the given text.
    RemoveStartingSingleQuoteAndEqualSign = RemoveStartingEqualSign(RemoveStartingSingleQuote(GivenText))
End Function

Public Function RemoveStartingSingleQuote(ByVal GivenText As String) As String
    
    ' Remove a single quote from the start of the given text.
    Dim Result As String
    Result = RemoveStartingChar(GivenText, SINGLE_QUOTE)
    Result = RemoveStartingChar(Result, SINGLE_QUOTE)
    
    RemoveStartingSingleQuote = Result
    
End Function

Public Function RemoveStartingEqualSign(ByVal GivenText As String) As String
    ' Remove an equal sign from the start of the given text.
    RemoveStartingEqualSign = RemoveStartingChar(GivenText, EQUAL_SIGN)
End Function

Private Function RemoveStartingChar(ByVal GivenText As String, ByVal GivenChar As String) As String
    
    ' Remove the given character from the start of the given text.
    If Left$(GivenText, 1) = GivenChar Then
        RemoveStartingChar = Text.RemoveFromStart(GivenText, Len(GivenChar))
    Else
        RemoveStartingChar = GivenText
    End If
    
End Function

Public Function IsRangeAddress(ByVal GivenRangeAddress As String) As Boolean
    
    ' Check if the given address is a valid range address.
    On Error GoTo NotRangeAddress
    Dim CurrentRange As Range
    Set CurrentRange = RangeResolver.GetRange(GivenRangeAddress)
    IsRangeAddress = IsNotNothing(CurrentRange)

Cleanup:
    Exit Function

NotRangeAddress:
    ' Handle specific errors that may occur when checking the range address.
    Select Case Err.Number
        Case 0
            MsgBox "As it is error handling, it should not come here.." _
                   & "You are doing something bad to handle error." _
                   & " Check Error Handling code and also check if you use Exit Procedure on Cleanup." _
                   , vbCritical, "Error"
        Case 1004
            IsRangeAddress = False
            Resume Cleanup
        Case Else
            Err.Raise Err.Number, "ModuleName along with Procedure name", "Description"
    End Select
    Resume Cleanup
    ' This is only for debugging purposes.
    Resume

End Function

Public Function IsInsideTable(ByVal GivenRange As Range) As Boolean
    
    ' Check if the given range is inside a table.
    Dim ActiveTable As ListObject
    Set ActiveTable = GetTableFromRange(GivenRange)
    IsInsideTable = IsNotNothing(ActiveTable)

End Function

Public Function GetTableFromRange(ByVal GivenRange As Range) As ListObject
    ' Get the table object from the given range.
    Set GetTableFromRange = GivenRange.ListObject
End Function

Public Function IsInsideNamedRange(ByVal GivenRange As Range) As Boolean
    
    ' Check if the given range is inside a named range.
    Dim CurrentName As name
    Set CurrentName = FindNamedRangeFromSubCell(GivenRange)
    IsInsideNamedRange = IsNotNothing(CurrentName)
    
End Function

Public Function FindNamedRangeFromSubCell(ByVal GivenRange As Range) As name
    
    ' Find the named range containing the given range.
    Dim CurrentNameRange As name
    Dim NameOfCurrentNamedRange As String
    Dim ReferredRange As Range
    For Each CurrentNameRange In GivenRange.Worksheet.Parent.Names
        If CurrentNameRange.Visible Then
            NameOfCurrentNamedRange = Replace(CurrentNameRange.name, EQUAL_SIGN, vbNullString)
            On Error Resume Next
            Set ReferredRange = CurrentNameRange.RefersToRange
            On Error GoTo 0
            If IsNothing(ReferredRange) Then
                ' Logger.Log DEBUG_LOG, NameOfCurrentNamedRange & " not found"
                ' Debug.Assert NameOfCurrentNamedRange <> "_xlpm.side1"
            ElseIf GivenRange.Worksheet.name = ReferredRange.Worksheet.name Then
                If HasIntersection(ReferredRange, GivenRange) Then
                    Set FindNamedRangeFromSubCell = CurrentNameRange
                    Exit Function
                End If
            End If
        End If
    Next CurrentNameRange

    Set FindNamedRangeFromSubCell = Nothing

End Function

Public Function RemoveEndingText(ByVal FromText As String, ByVal RemoveText As String) As String
    
    ' Remove the given text from the end of the given text.
    If Right$(FromText, Len(RemoveText)) = RemoveText Then
        RemoveEndingText = VBA.Left$(FromText, Len(FromText) - Len(RemoveText))
    Else
        RemoveEndingText = FromText
    End If
    
End Function

Public Function IsNotInside(ByVal SearchInRange As Range, ByVal SearchForRange As Range) As Boolean
    ' Check if the search range is not inside the search for range.
    IsNotInside = IsNothing(FindIntersection(SearchInRange, SearchForRange))
End Function

Public Function IsOptionalArgument(ByVal LetOrLambdaFormula As String, ByVal NameInFormula As String) As Boolean
    
    ' Check if the given argument is an optional argument in the LET or LAMBDA formula.
    Dim IsOmittedText As String
    IsOmittedText = ISOMITTED_FX_NAME & FIRST_PARENTHESIS_OPEN _
                    & NameInFormula & FIRST_PARENTHESIS_CLOSE
    IsOptionalArgument = Text.Contains(LetOrLambdaFormula, IsOmittedText, CONSIDER_CASE)
    
End Function

Public Function RemoveStartingLetAndEndParenthesis(ByVal GiveLetFormula As String) As String
    
    ' Remove the starting LET and ending parenthesis from the given LET formula.
    GiveLetFormula = modUtility.RemoveStartingEqualSign(GiveLetFormula)
    GiveLetFormula = Replace(GiveLetFormula, LET_AND_OPEN_PAREN, vbNullString, 1, 1, vbTextCompare)
    GiveLetFormula = modUtility.RemoveEndingText(GiveLetFormula, FIRST_PARENTHESIS_CLOSE)
    
    RemoveStartingLetAndEndParenthesis = GiveLetFormula

End Function

' @Description("This will convert a list of item to a collection with having both key and value same.")
' @Dependency("No Dependency")
' @ExampleCall : VectorToCollection Array("key","value")
' @Date : 08 May 2022 06:09:29 PM
' @PossibleError :
Public Function VectorToCollection(ByVal GivenVector As Variant) As Collection
    
    ' Convert a one-dimensional array (vector) to a collection.
    Dim Result As Collection
    Set Result = New Collection
    Dim CurrentItem As Variant
    For Each CurrentItem In GivenVector
        Result.Add CurrentItem, CStr(CurrentItem)
    Next CurrentItem
    Set VectorToCollection = Result

End Function

Public Function FindFirstNotUsedCell(ByVal ForSheet As Worksheet) As Range

    ' Find the first not used cell in the specified worksheet.
    Dim UsedRanges As Range
    Set UsedRanges = ForSheet.UsedRange
    Dim LastCell As Range
    Set LastCell = UsedRanges.Cells(Application.Min(Cells.Rows.Count - 1000, UsedRanges.Rows.Count) _
                                    , Application.Min(Cells.Columns.Count - 1000, UsedRanges.Columns.Count))
    Set FindFirstNotUsedCell = LastCell.Offset(0, 1)

End Function

' @Description("This will find the first index of matching item in the array. If no match then it will return -1")
' @Dependency("IndexOf")
' @ExampleCall :FunctionCollection.FirstIndexOf(GivenArray, "Channel")
' @Date : 06 January 2022 10:02:13 PM
' @PossibleError :Type mismatch(13) if given input is not an array
Public Function FirstIndexOf(ByVal SearchInArray As Variant, ByVal SearchFor As Variant _
                                                            , Optional ByVal SearchInColumn As Long = -1 _
                                                             , Optional ByVal IsMustEqual As Boolean = False) As Long
    FirstIndexOf = IndexOf(SearchInArray, SearchFor, True, SearchInColumn, IsMustEqual)
End Function

' @Description("This will find the last index of matching item in the array. If no match then it will return -1")
' @Dependency("IndexOf")
' @ExampleCall :FunctionCollection.LastIndexOf(GivenArray, "Channel")
' @Date : 06 January 2022 10:02:13 PM
' @PossibleError :Type mismatch(13) if given input is not an array
Public Function LastIndexOf(ByVal SearchInArray As Variant, ByVal SearchFor As Variant _
                                                           , Optional ByVal SearchInColumn As Long = -1 _
                                                            , Optional ByVal IsMustEqual As Boolean = False) As Long
    LastIndexOf = IndexOf(SearchInArray, SearchFor, False, SearchInColumn, IsMustEqual)
End Function

' @Description("Private function used from two different function")
' @Dependency("Text.Contains")
' @ExampleCall : IndexOf(SearchInArray, SearchFor, False, SearchInColumn, IsMustEqual)
' @Date : 06 January 2022 10:02:13 PM
' @PossibleError :Type mismatch(13) if given input is not an array
Private Function IndexOf(ByVal SearchInArray As Variant, ByVal SearchFor As Variant _
                                                        , Optional ByVal IsSearchFromTop As Boolean = True _
                                                         , Optional ByVal SearchInColumn As Long = -1 _
                                                         , Optional ByVal IsMustEqual As Boolean = False) As Long
    
    If SearchInColumn = -1 Then
        SearchInColumn = LBound(SearchInArray, 2)
    End If

    Dim FromIndex As Long
    Dim ToIndex As Long
    Dim StepBy As Long
    If IsSearchFromTop Then
        FromIndex = LBound(SearchInArray, 1)
        ToIndex = UBound(SearchInArray, 1)
        StepBy = 1
    Else
        FromIndex = UBound(SearchInArray, 1)
        ToIndex = LBound(SearchInArray, 1)
        StepBy = -1
    End If

    Dim CurrentRowIndex As Long
    Dim SearchInText As String

    For CurrentRowIndex = FromIndex To ToIndex Step StepBy
        If IsMustEqual Then
            If SearchInArray(CurrentRowIndex, SearchInColumn) = SearchFor Then
                IndexOf = CurrentRowIndex
                Exit Function
            End If
        Else
            SearchInText = SearchInArray(CurrentRowIndex, SearchInColumn)
            If Text.Contains(SearchInText, CStr(SearchFor)) Then
                IndexOf = CurrentRowIndex
                Exit Function
            End If
        End If
    Next CurrentRowIndex
    IndexOf = -1

End Function

Public Function ConcatenateRangeOfRowOfA2DArray(ByVal OnArray As Variant _
                                                , ByVal FromRow As Long _
                                                 , ByVal ToRow As Long, ByVal OnColumn As Long) As String
    Dim Result As String
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = FromRow To ToRow
        Result = Result & OnArray(CurrentRowIndex, OnColumn)
    Next CurrentRowIndex
    ConcatenateRangeOfRowOfA2DArray = Result

End Function

Public Sub RemoveDataFromRangeOfRowsOfA2DArray(ByRef FromArray As Variant _
                                               , ByVal FromRowIndex As Long _
                                                , ByVal ToRowIndex As Long, ByVal FromColumn As Long)
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = FromRowIndex To ToRowIndex
        FromArray(CurrentRowIndex, FromColumn) = vbNullString
    Next CurrentRowIndex

End Sub

Public Function ExtractStartFormulaName(ByVal FormulaText As String) As String
    
    ' Extracts the formula name from the given formula text.
    ' If the formula contains parentheses (indicating a function call), it extracts the name before the first parenthesis.
    ' Otherwise, it considers the entire formula text as the name.
    
    Dim Result As String
    
    If Text.Contains(FormulaText, FIRST_PARENTHESIS_OPEN) Then
        Result = Text.BeforeDelimiter(FormulaText, FIRST_PARENTHESIS_OPEN)
    Else
        Result = FormulaText
    End If

    ' Remove the equal sign if present at the beginning and trim any leading/trailing spaces.
    Result = Text.RemoveFromStartIfPresent(Result, EQUAL_SIGN)
    
    ExtractStartFormulaName = Text.Trim(Result)

End Function

Public Function IsCellHasSavedLambdaFormula(ByVal FromCell As Range) As Boolean
    
    ' Checks if the given cell contains a Lambda formula.
    Dim LambdaFormulaName As String
    LambdaFormulaName = ExtractStartFormulaName(FromCell.Formula2)

    ' If the cell contains a valid formula name, retrieve the formula text from the defined name.
    If LambdaFormulaName <> vbNullString Then
        Dim FormulaText As String
        On Error GoTo ErrHandler
        FormulaText = FromCell.Worksheet.Parent.Names(LambdaFormulaName).RefersTo
        IsCellHasSavedLambdaFormula = IsLambdaFunction(FormulaText)
    End If
    Exit Function

ErrHandler:
    ' If an error occurs while retrieving the formula text, it indicates that the name is not defined or not a valid Lambda formula.
    IsCellHasSavedLambdaFormula = False

End Function

Public Function EscapeQuotes(ByVal InputText As String) As String
    EscapeQuotes = Replace(InputText, QUOTES, QUOTES & QUOTES)
End Function

Public Function RemoveEscapeQuotes(ByVal InputText As String) As String
    RemoveEscapeQuotes = Replace(InputText, QUOTES & QUOTES, QUOTES)
End Function

Private Function IsFirstCharEqualExceptWhiteSpace(ByVal GivenText As String) As Boolean
    GivenText = RemoveInitialSpaceAndNewLines(GivenText)
    IsFirstCharEqualExceptWhiteSpace = Text.IsStartsWith(GivenText, EQUAL_SIGN)
End Function

Public Function IsMetadataLetVarName(ByVal GivenText As String) As Boolean
    ' Checks if the given text is a metadata 'LET' variable name.
    IsMetadataLetVarName = Text.IsStartsWith(GivenText, METADATA_IDENTIFIER)
End Function

Public Function FindMetadataValue(ByVal LetParts As Variant, ByVal MetadataVarName As String) As String
    
    ' Finds and returns the value of the given metadata variable in the TokenizedFormula.
    MetadataVarName = Text.RemoveFromEndIfPresent(MetadataVarName, LIST_SEPARATOR)
    If Not IsArrayAllocated(LetParts) Then
        FindMetadataValue = vbNullString
        Exit Function
    End If
    
    Dim FoundAtIndex As Long
    FoundAtIndex = modUtility.FirstIndexOf(LetParts, MetadataVarName, LBound(LetParts, 2), True)
    If FoundAtIndex = -1 Then
        Logger.Log INFO_LOG, MetadataVarName & "Metadata var is not found."
        FindMetadataValue = vbNullString
        Exit Function
    End If
    
    FindMetadataValue = LetParts(FoundAtIndex, LBound(LetParts, 2) + LET_PARTS_VALUE_COL_INDEX - 1)
    
End Function

Private Function FindNextArgumentSeparatorOrEndOfListIndex(ByVal TokenizedFormula As Variant _
                                                           , ByVal StartFromRow As Long) As Long
    
    ' Finds the row index of the next argument separator or the end of the list in the TokenizedFormula.
    Dim FirstColumn As Long
    FirstColumn = LBound(TokenizedFormula, 2)
    Dim CurrentRow As Long
    For CurrentRow = StartFromRow To UBound(TokenizedFormula, 1)
        If TokenizedFormula(CurrentRow, FirstColumn + 1) = ARGUMENT_SEPARATOR Then
            FindNextArgumentSeparatorOrEndOfListIndex = CurrentRow
            Exit For
        End If
    Next CurrentRow

    If FindNextArgumentSeparatorOrEndOfListIndex = 0 Then
        FindNextArgumentSeparatorOrEndOfListIndex = UBound(TokenizedFormula, 1)
    End If
    
End Function

Public Function IsLambdaInEditMode(ByVal OfCell As Range, ByVal Prefix As String) As Boolean
    
    ' Checks if the lambda expression in the OfCell is in edit mode (modified from the original).
    Dim OldLambdaName As String
    OldLambdaName = GetOldNameFromComment(OfCell, Prefix)
    IsLambdaInEditMode = (OldLambdaName <> vbNullString)
    
End Function

Public Function GetOldNameFromComment(ByVal FromCell As Range, ByVal Prefix As String) As String
    
    ' Retrieves the old lambda name from the comment in the FromCell with the specified prefix.
    On Error GoTo NoComment
    Dim CurrentComment As Comment
    Set CurrentComment = FromCell.Comment
    If Text.IsStartsWith(CurrentComment.Text, Prefix) Then
        GetOldNameFromComment = Replace(CurrentComment.Text, Prefix, vbNullString)
    End If
    Exit Function

NoComment:
    GetOldNameFromComment = vbNullString
    
End Function

Public Sub UpdateNameComment(ByVal GivenName As name, ByVal NewComment As String)
    GivenName.Comment = NewComment
End Sub

Public Function ExtractNameFromLocalNameRange(ByVal LocalName As String) As String
    
    ' Extracts the name from a local named range.
    Dim Result As String
    If Text.Contains(LocalName, SHEET_NAME_SEPARATOR) Then
        Result = Text.AfterDelimiter(LocalName, SHEET_NAME_SEPARATOR, , FROM_END)
    Else
        Result = LocalName
    End If
    
    ExtractNameFromLocalNameRange = Result
    
End Function

Public Function CleanVarName(ByVal GivenName As String) As String
    ' Cleans the variable name by removing leading and trailing spaces and any non-printable characters.
    CleanVarName = VBA.LTrim$(Application.WorksheetFunction.Clean(Application.WorksheetFunction.Trim(GivenName)))
End Function

' @PureFunction
Public Function SplitArrayConstantTo2DArray(ByVal ArrayConstant As String) As Variant
    
    ' Splits the ArrayConstant string to a 2D array.
    Dim SplittedArrayConstant As Variant
    SplittedArrayConstant = Evaluate(EQUAL_SIGN & ArrayConstant)
    If Not IsArray(SplittedArrayConstant) Then SplittedArrayConstant = Array(SplittedArrayConstant)
    Dim Result As Variant
    If DimensionOfAnArray(SplittedArrayConstant) = 1 Then
    
        ReDim Result(1 To 1, LBound(SplittedArrayConstant) To UBound(SplittedArrayConstant))
        Dim ColIndex As Long
        For ColIndex = LBound(SplittedArrayConstant) To UBound(SplittedArrayConstant)
            Result(1, ColIndex) = SplittedArrayConstant(ColIndex)
        Next ColIndex
        
    Else
        Result = SplittedArrayConstant
    End If
    SplitArrayConstantTo2DArray = Result
    
End Function

' PureFunction
Public Function DimensionOfAnArray(ByVal GivenArray As Variant) As Long

    ' Purpose: get array dimension (MS)
    Dim Dimension As Long
    Dim ErrorCheck As Long
    On Error GoTo FinalDimension

    For Dimension = 1 To 60                      ' 60 being the absolute dimensions limitation
        ErrorCheck = LBound(GivenArray, Dimension)
    Next
    DimensionOfAnArray = Dimension - 1
    Exit Function
    
FinalDimension:
    DimensionOfAnArray = Dimension - 1

End Function

Public Function FindUniqueNameByIncrementingNumber(ByVal AllNameMap As Collection _
                                                   , ByVal StartName As String) As String

    If Not modUtility.IsExistInCollection(AllNameMap, StartName) Then
        FindUniqueNameByIncrementingNumber = StartName
        Exit Function
    End If

    If Text.ExtractNumberFromEnd(StartName) = vbNullString Then
        StartName = Text.PadIfNotPresent(StartName, "_2", FROM_END)
    End If

    Do While modUtility.IsExistInCollection(AllNameMap, StartName)
        StartName = Text.IncrementOrDecrementEndNumber(StartName, 1, False)
    Loop
    FindUniqueNameByIncrementingNumber = StartName

End Function

Public Sub AutoFitRange(ByVal ForRange As Range, ByVal MaximumColumnWidth As Long, ByVal MinimumColumnWidth As Long)
    
    ' Autofit columns in the given range and limit their width between MaximumColumnWidth and MinimumColumnWidth.
    Dim CurrentRange As Range
    For Each CurrentRange In ForRange.Areas
        ' Autofit columns in the current area.
        CurrentRange.Columns.AutoFit

        Dim Counter As Long
        For Counter = 1 To CurrentRange.Columns.Count
            ' Check and adjust column width if it exceeds the specified limits.
            Dim ColWidth As Double
            ColWidth = CurrentRange.Columns(Counter).ColumnWidth
            If ColWidth > MaximumColumnWidth Then
                CurrentRange.Columns(Counter).ColumnWidth = MaximumColumnWidth
            ElseIf ColWidth < MinimumColumnWidth Then
                CurrentRange.Columns(Counter).ColumnWidth = MinimumColumnWidth
            End If
        Next Counter
    Next CurrentRange
    
End Sub

Public Function ConvertToFullyQualifiedCellRef(ByVal ForCell As Range) As String
    
    ' Converts a cell reference to a fully qualified cell reference with book name and sheet names.
    ' Example output: '[Different Locale Functions Map.xlsm]Keywords Locale Map'!$H$6
    ConvertToFullyQualifiedCellRef = SINGLE_QUOTE & LEFT_BRACKET _
                                   & EscapeSingeQuote(ForCell.Worksheet.Parent.name) _
                                   & RIGHT_BRACKET & EscapeSingeQuote(ForCell.Worksheet.name) & SINGLE_QUOTE _
                                   & SHEET_NAME_SEPARATOR & ForCell.Address
                                     
End Function

Public Function GetRangeReference(ByVal GivenCells As Range _
                                  , Optional ByVal IsAbsolute As Boolean = True) As String
    
    ' Retrieves the reference of the given range as a string.

    GetRangeReference = GivenCells.Address(IsAbsolute, IsAbsolute)

    ' Check if the given range is part of a dynamic array formula.
    If GivenCells.Cells.Count > 1 And GivenCells.Cells(1, 1).HasSpill Then
        Dim TempRange As Range
        Set TempRange = GivenCells.Cells(1, 1)

        ' If it is a spill range, append the dynamic cell reference sign to the range reference.
        If TempRange.SpillParent.SpillingToRange.Address = GivenCells.Address Then
            GetRangeReference = TempRange.SpillParent.Address(IsAbsolute, IsAbsolute) & DYNAMIC_CELL_REFERENCE_SIGN
        End If
    End If
    
End Function

Public Function GetSheetRefForRangeReference(ByVal SheetName As String _
                                             , Optional ByVal IsSingleQuoteMandatory As Boolean = False) As String
    
    ' Returns the sheet reference for the range reference.
    Dim IsSingleQuoteNeeded As Boolean
    If IsSingleQuoteMandatory Then
        IsSingleQuoteNeeded = True
    Else
        IsSingleQuoteNeeded = IsAnyNonAlphanumeric(SheetName)
    End If
    
    Dim Result As String
    If IsSingleQuoteNeeded Then
        ' for single quote we need to escape with double single quote
        Result = SINGLE_QUOTE _
               & EscapeSingeQuote(SheetName) _
               & SINGLE_QUOTE & SHEET_NAME_SEPARATOR
    Else
        Result = SheetName & SHEET_NAME_SEPARATOR
    End If
    
    GetSheetRefForRangeReference = Result
    
End Function

Public Function IsAnyNonAlphanumeric(ByVal Text As String) As String

    Dim Result As Boolean
    Dim Index As Long
    Dim CurrentCharacter As String
    For Index = 1 To Len(Text)
        CurrentCharacter = Mid(Text, Index, 1)
        If Not CurrentCharacter Like "[A-Za-z0-9]" Then
            Result = True
            Exit For
        End If
    Next Index

    IsAnyNonAlphanumeric = Result

End Function

Public Function RemoveSheetQualifierIfPresent(ByVal RangeRef As String) As String
    
    ' e.g:  RemoveSheetQualifierFromSheetQualifiedRangeRef("'All Functions Name'!$C$5") >> $C$5
    
    Dim Result As String
    
    If Text.Contains(RangeRef, SHEET_NAME_SEPARATOR) Then
        Result = Text.AfterDelimiter(RangeRef, SHEET_NAME_SEPARATOR, , FROM_END)
    Else
        Result = RangeRef
    End If
    
    RemoveSheetQualifierIfPresent = Result
    
End Function

Public Function GetRangeRefWithSheetName(ByVal GivenRange As Range _
                                         , Optional ByVal IsSingleQuoteMandatory As Boolean = False _
                                          , Optional ByVal IsAbsolute As Boolean = True) As String
    
    ' Returns the reference of the given range with the sheet name.
    ' If IsAbsolute is True, the reference is absolute; otherwise, it's relative.
    Dim SheetRef As Worksheet
    Set SheetRef = GivenRange.Parent
    GetRangeRefWithSheetName = GetSheetRefForRangeReference(SheetRef.name, IsSingleQuoteMandatory) _
                               & GetRangeReference(GivenRange, IsAbsolute)
                               
End Function

Public Function FindLambdas(ByVal FromBook As Workbook) As Collection
    
    ' Finds all lambda functions in the given workbook and returns a collection of their names.
    Logger.Log TRACE_LOG, "Enter modUtility.FindLambdas"
    Dim CurrentName As name
    Dim AllLambda As Collection
    Set AllLambda = New Collection
    For Each CurrentName In FromBook.Names
        ' Check if the name refers to a lambda function.
        If IsLambdaFunction(CurrentName.RefersTo) Then
            ' Add the name to the collection of lambda functions.
            AllLambda.Add CurrentName, CStr(CurrentName.name)
        End If
    Next CurrentName
    Set FindLambdas = AllLambda
    Logger.Log TRACE_LOG, "Exit modUtility.FindLambdas"
    
End Function

Public Function IsAllCellBlank(ByVal NeededRange As Range) As Boolean
    ' Checks if all cells in the specified range are blank.
    IsAllCellBlank = (Application.WorksheetFunction.CountA(NeededRange) = 0)
End Function

' @Description : By default it will not through error for duplicate case
' @FullyQualifiedCase: It will use given column indexes
' @ExampleCall : Set ArrayToCollectionMapping = ArrayToCollection(SUTArray, 1, 2, IsSuppressDuplicateError:=False) >> It will through error if any duplicate key is present.
' @ExampleCall : Set ArrayToCollectionMapping = ArrayToCollection(SUTArray, 1, 2, IsSuppressDuplicateError:=True) >> It will skip duplicate item.

' @DefaultValueCase : first column is the key and second column is the item
' @ExampleCall : Set ArrayToCollectionMapping = ArrayToCollection(SUTArray, IsSuppressDuplicateError:=False) >> It will through error if any duplicate key is present.
' @ExampleCall : Set ArrayToCollectionMapping = ArrayToCollection(SUTArray, IsSuppressDuplicateError:=True) >> It will skip duplicate item.
Public Function ArrayToCollection(ByVal GivenArray As Variant _
                                  , Optional ByVal KeyColumnIndex As Long = -1 _
                                  , Optional ByVal ItemColumnIndex As Long = -1 _
                                  , Optional ByVal IsSuppressDuplicateError As Boolean = True) As Variant


    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(GivenArray, 2)

    If KeyColumnIndex = -1 Then KeyColumnIndex = FirstColumnIndex
    If ItemColumnIndex = -1 Then ItemColumnIndex = FirstColumnIndex + 1
    If IsSuppressDuplicateError Then On Error Resume Next

    Dim CurrentRowIndex As Long
    Dim KeyItemMapping As Collection
    Set KeyItemMapping = New Collection
    For CurrentRowIndex = LBound(GivenArray, 1) To UBound(GivenArray, 1)
        Dim Key As String
        Key = CStr(GivenArray(CurrentRowIndex, KeyColumnIndex))
        Dim Item As Variant
        Item = GivenArray(CurrentRowIndex, ItemColumnIndex)
        KeyItemMapping.Add Item, Key
    Next CurrentRowIndex

    Set ArrayToCollection = KeyItemMapping
    If IsSuppressDuplicateError Then On Error GoTo 0
    
End Function

Public Function IsLetForOuterLambda(ByVal NewTokenizedFormula As Variant _
                                    , Optional ByVal LetIndex As Long = -1) As Boolean
    
    If LetIndex = -1 Then LetIndex = modUtility.FirstIndexOf(NewTokenizedFormula, LET_FX_NAME, , True)
    Dim NumberOfLambda As Long
    Dim Counter As Long
    For Counter = LetIndex - 1 To LBound(NewTokenizedFormula, 1) Step -1
        If NewTokenizedFormula(Counter, LBound(NewTokenizedFormula, 2)) = LAMBDA_FX_NAME Then
            NumberOfLambda = NumberOfLambda + 1
        End If
    Next Counter
    IsLetForOuterLambda = (NumberOfLambda = 1)
    
End Function

Public Function VectorToArray(ByVal Vector As Variant, Optional ByVal IsInColumn As Boolean = False) As Variant
    
    Dim NumberOfRow As Long
    NumberOfRow = UBound(Vector) - LBound(Vector) + 1
    
    Dim Result As Variant
    If IsInColumn Then
        ReDim Result(1 To 1, 1 To NumberOfRow)
    Else
        ReDim Result(1 To NumberOfRow, 1 To 1)
    End If
    
    Dim Counter As Long
    Counter = LBound(Vector) - 1
    
    Dim CurrentValue As Variant
    For Each CurrentValue In Vector
        Counter = Counter + 1
        If IsInColumn Then
            Result(1, Counter) = CurrentValue
        Else
            Result(Counter, 1) = CurrentValue
        End If
    Next CurrentValue
    VectorToArray = Result
    
End Function

Public Function FindAllNamedRange(ByVal FromBook As Workbook, ByVal IsRefersToRange As Boolean) As Collection
    
    Dim Result As Collection
    Set Result = New Collection
    Dim CurrentName As name
    For Each CurrentName In FromBook.Names
        
        If CurrentName.Visible Then
            If IsRefersToRange Then
                If Not IsRefersToRangeIsNothing(CurrentName) Then
                    Result.Add CurrentName, CStr(CurrentName.name)
                End If
            Else
                Result.Add CurrentName, CStr(CurrentName.name)
            End If
        End If
        
    Next CurrentName
    Set FindAllNamedRange = Result
    
End Function

Public Function IsLocalScopeNamedRange(ByVal LocalName As String) As Boolean
    
    Dim FoundAt As Long
    FoundAt = InStr(1, LocalName, SHEET_NAME_SEPARATOR)
    IsLocalScopeNamedRange = (FoundAt <> 0)
    
End Function

Public Function IsNotLetVarValueAssignSection(ByVal CurrentIndex As Long, ByVal TokenizedFormula As Variant _
                                                                         , ByVal ValidVarName As String) As Boolean
    ' Checks if the current index position is not in a LET variable value assignment section.
    ' Returns True if it's not in a LET variable value assignment section; otherwise, False.

    Dim Counter As Long
    Dim Temp As String

    Dim Offset As Long
    For Counter = CurrentIndex - 1 To LBound(TokenizedFormula, 1) Step -1
        Offset = Offset + 1
        Temp = TokenizedFormula(Counter, LBound(TokenizedFormula, 2) + 1)
        If Trim$(Temp) = modSharedConstant.LET_STEP_NAME_TOKEN Then
            IsNotLetVarValueAssignSection = Not (TokenizedFormula(Counter, LBound(TokenizedFormula, 2)) = ValidVarName)
            Exit Function
        End If
    Next Counter
    
End Function

Public Sub ScrollToDependencyDataRange(ByVal Table As ListObject)
    
    ' Scrolls to the dependency data range in the specified table.
    Application.GoTo Table.Range, True
    Table.Range(1, 1).Select
    
End Sub

Public Sub AssingOnUndo(ByVal UndoForMethod As String)
    
    ' Assigns an Undo method for the specified action.
    Dim UndoSubName As String
    UndoSubName = SINGLE_QUOTE & EscapeSingeQuote(ThisWorkbook.name) _
                & SINGLE_QUOTE & EXCLAMATION_SIGN & UndoForMethod & "_Undo"
    Application.OnUndo "Undo " & UndoForMethod & " Action", UndoSubName
    
End Sub

Public Sub AssignFormulaIfErrorPrintIntoDebugWindow(ByVal PutFormulaOnCell As Range _
                                                    , ByVal FormulaText As String _
                                                     , Optional ByVal Message As String = vbNullString)
    
    ' Assigns a formula to the specified cell and prints the formula into the debug window if an error occurs.
    On Error GoTo PrintFormulaToDebugWindow
    PutFormulaOnCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(FormulaText)
    Exit Sub

PrintFormulaToDebugWindow:
    Debug.Print Message & FormulaText
    
End Sub

Public Function FindIntersection(ByVal FirstRange As Range, ByVal SecondRange As Range) As Range
    
    On Error Resume Next
    Set FindIntersection = Intersect(FirstRange, SecondRange)
    On Error GoTo 0
    
End Function

Public Function IsNoIntersection(ByVal FirstRange As Range, ByVal SecondRange As Range) As Boolean
    IsNoIntersection = IsNothing(FindIntersection(FirstRange, SecondRange))
End Function

Public Function HasIntersection(ByVal FirstRange As Range, ByVal SecondRange As Range) As Boolean
    HasIntersection = IsNotNothing(FindIntersection(FirstRange, SecondRange))
End Function

Public Function IsNothing(ByVal GivenObject As Object) As Boolean
    IsNothing = (GivenObject Is Nothing)
End Function

Public Function IsNotNothing(ByVal GivenObject As Object) As Boolean
    IsNotNothing = (Not GivenObject Is Nothing)
End Function

Public Function UpdateForIsOmitted(ByVal GivenFormula As String) As String
    
    ' Sample input:  =F3 + IF(ISOMITTED(G3), 1, G3)
    ' Sample output: =F3 + IF(OR(ISOMITTED(G3),AND(ISBLANK(G3))), 1, G3)
    
    ' Sample input:  =F3 + IF(OR(ISOMITTED(G3),AND(ISBLANK(G3))), 1, G3)
    ' Sample output: =F3 + IF(OR(ISOMITTED(G3),AND(ISBLANK(G3))), 1, G3)
    
    ' Updates the formula to replace occurrences of ISOMITTED with OR(ISOMITTED(VarName), AND(ISBLANK(VarName))).
    ' GivenFormula: The original formula to update.
    
     Const VAR_PLACE_HOLDER As String = "{VarName}"
    Dim IsOmittedWithBlank As String
    IsOmittedWithBlank = OR_FX_NAME & FIRST_PARENTHESIS_OPEN & ISOMITTED_FX_NAME _
                         & "(" & VAR_PLACE_HOLDER & ")" & LIST_SEPARATOR & AND_FX_NAME _
                         & FIRST_PARENTHESIS_OPEN & ISBLANK_FX_NAME & "(" & VAR_PLACE_HOLDER & ")))"
                         
    Dim IsOmittedWithBlankAndSpace As String
    
    IsOmittedWithBlankAndSpace = OR_FX_NAME & FIRST_PARENTHESIS_OPEN _
                                 & ISOMITTED_FX_NAME & "(" & VAR_PLACE_HOLDER & ")" _
                                 & LIST_SEPARATOR & ONE_SPACE _
                                 & AND_FX_NAME & FIRST_PARENTHESIS_OPEN _
                                 & ISBLANK_FX_NAME & "(" & VAR_PLACE_HOLDER & ")))"

    Dim IsOmittedWithParen As String
    IsOmittedWithParen = ISOMITTED_FX_NAME & FIRST_PARENTHESIS_OPEN

    Dim AllPosition As Collection
    Set AllPosition = Text.FindAllIndexOf(GivenFormula, IsOmittedWithParen, FROM_START, IGNORE_CASE)

    If AllPosition.Count = 0 Then
        UpdateForIsOmitted = GivenFormula
        Exit Function
    End If

    Dim CurrentIndex As Long
    Dim Counter As Long
    For Counter = AllPosition.Count To 1 Step -1
        CurrentIndex = AllPosition.Item(Counter)
        Dim RefAfterIsOmitted As String
        Dim CloseParenIndex As Long
        CloseParenIndex = InStr(CurrentIndex, GivenFormula, FIRST_PARENTHESIS_CLOSE)
        RefAfterIsOmitted = FromRange(GivenFormula, CurrentIndex + Len(IsOmittedWithParen), CloseParenIndex - 1)

        Dim SearchText As String
        SearchText = VBA.Replace(IsOmittedWithBlank, VAR_PLACE_HOLDER, RefAfterIsOmitted)

        Dim SearchTextWithSpace As String
        SearchTextWithSpace = VBA.Replace(IsOmittedWithBlankAndSpace, VAR_PLACE_HOLDER, RefAfterIsOmitted)
        
        Dim LengthOfORAndOpenParen As String
        LengthOfORAndOpenParen = Len(OR_FX_NAME & "(")
        
        If CurrentIndex - LengthOfORAndOpenParen <= 0 Then
            ' Update with OR one
            GivenFormula = Text.ReplaceRange(GivenFormula, CurrentIndex, Len(IsOmittedWithParen _
                                                                             & RefAfterIsOmitted _
                                                                             & FIRST_PARENTHESIS_CLOSE), SearchText)
        ElseIf Mid$(GivenFormula, CurrentIndex - LengthOfORAndOpenParen, Len(SearchText)) <> SearchText And _
               Mid$(GivenFormula, CurrentIndex - LengthOfORAndOpenParen, Len(SearchTextWithSpace)) <> SearchTextWithSpace Then
            ' Update with OR one
            GivenFormula = Text.ReplaceRange(GivenFormula, CurrentIndex _
                                                          , Len(IsOmittedWithParen _
                                                                & RefAfterIsOmitted _
                                                                & FIRST_PARENTHESIS_CLOSE), SearchText)
        End If
    Next Counter
    
    UpdateForIsOmitted = GivenFormula
    
End Function

Public Function GetMatchingVarNameDependency(ByVal VarName As String _
                                             , ByVal FromCollection As Collection) As DependencyInfo
    
    ' Returns the DependencyInfo object that matches the given VarName from the provided collection.
    ' VarName: The variable name to search for.
    ' FromCollection: The collection to search for the matching DependencyInfo.

    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In FromCollection
        If CurrentDependencyInfo.ValidVarName = VarName Then
            Set GetMatchingVarNameDependency = CurrentDependencyInfo
            Exit Function
        End If
    Next CurrentDependencyInfo
    
End Function

Public Function GetNamedRangeToDictionary(ByVal FromBook As Workbook) As Dictionary
    
    ' Returns a dictionary that maps named range names to their respective named ranges in the specified workbook.
    ' FromBook: The workbook to extract the named ranges from.

    Dim NamedRangeNameToNamedRangeMap As Dictionary
    Set NamedRangeNameToNamedRangeMap = New Scripting.Dictionary

    Dim CurrentName As name
    For Each CurrentName In FromBook.Names
        If CurrentName.Visible Then
            NamedRangeNameToNamedRangeMap.Add CStr(CurrentName.name), CurrentName
        End If
    Next CurrentName
    Set GetNamedRangeToDictionary = NamedRangeNameToNamedRangeMap
    
End Function

Public Function IsInsideTableOrNamedRange(ByVal CurrentCell As Range) As Boolean
    
    ' Checks if the given cell is inside a table or a named range.
    ' CurrentCell: The cell to check.

    If modUtility.IsInsideNamedRange(CurrentCell) Then
        IsInsideTableOrNamedRange = True
    ElseIf modUtility.IsInsideTable(CurrentCell) Then
        IsInsideTableOrNamedRange = True
    End If
    
End Function

Public Function HasFormula(ByVal GivenCells As Range) As Boolean
    
    Dim Result As Boolean
    If GivenCells Is Nothing Then
        Result = False
    Else
        Result = GivenCells.Cells(1, 1).HasFormula
    End If
    
    HasFormula = Result
    
End Function

Public Function IsBothRangeEqual(ByVal FirstRange As Range, ByVal SecondRange As Range) As Boolean
    IsBothRangeEqual = ( _
                       FirstRange.Address = SecondRange.Address _
                       And FirstRange.Worksheet.name = SecondRange.Worksheet.name _
                       )
End Function

Public Function ConvertToValueFormula(ByVal AllData As Variant _
                                       , Optional ByVal DefaultIfBlank As Variant = 0) As String
    
    ' Converts a 2D array to a string representation of an array constant.
    ' If single cell value then return the value.
    ' It handles error properly.
    
    If Not IsArray(AllData) Then
        ConvertToValueFormula = GetFormattedValueForArrayConst(AllData, DefaultIfBlank)
        Exit Function
    End If

    Dim Value As String
    Value = LEFT_BRACE

    Dim CurrentRowIndex As Long
    Dim CurrentColumnIndex As Long

    ' Loop through the 2D array and build the string representation of the array constant.
    For CurrentRowIndex = LBound(AllData, 1) To UBound(AllData, 1)
        
        For CurrentColumnIndex = LBound(AllData, 2) To UBound(AllData, 2)
            Value = Value & GetFormattedValueForArrayConst(AllData(CurrentRowIndex, CurrentColumnIndex))
            If CurrentColumnIndex <> UBound(AllData, 2) Then
                Value = Value & ARRAY_CONST_COLUMN_SEPARATOR
            End If
        Next CurrentColumnIndex
        
        Value = Value & ARRAY_CONST_ROW_SEPARATOR
    
    Next CurrentRowIndex

    ' Remove the trailing row separator and close the array constant.
    Value = Left$(Value, Len(Value) - Len(ARRAY_CONST_ROW_SEPARATOR))
    Value = Value & RIGHT_BRACE
    
    ConvertToValueFormula = Value               ' Return the final array constant string.
    
End Function

Private Function GetFormattedValueForArrayConst(ByVal Value As Variant _
                                                , Optional ByVal DefaultIfBlank As Variant = 0) As Variant
    
    Dim Result As Variant
    
    If IsError(Value) Then
        Result = GetErrorText(Value)
    ElseIf IsNumeric(Value) Then
        Result = Value                           ' Return numeric value as-is.
    ElseIf Value = vbNullString Then
        Result = DefaultIfBlank                  ' Return the default value for blank data.
    Else
        Result = QUOTES & Value & QUOTES        ' Wrap non-numeric values with quotes.
    End If
    
    GetFormattedValueForArrayConst = Result
    
End Function

Public Function GetErrorText(ByVal ErrData As Variant) As String
    
    If Not IsError(ErrData) Then Exit Function
    
    Dim ErrorText As String
    
    Select Case CVErr(ErrData)
        
        Case CVErr(xlErrValue), CVErr(xlErrCalc), CVErr(xlErrSpill)
            ErrorText = "#VALUE!"
        
        Case CVErr(xlErrDiv0)
            ErrorText = "#DIV/0!"
        
        Case CVErr(xlErrNA)
            ErrorText = "#N/A"
            
        Case CVErr(xlErrName)
            ErrorText = "#NAME?"
            
        Case CVErr(xlErrNull)
            ErrorText = "#NULL!"
            
        Case CVErr(xlErrNum)
            ErrorText = "#NUM!"
            
        Case CVErr(xlErrRef)
            ErrorText = "#REF!"
        
        Case CVErr(xlErrBlocked)
            ErrorText = "#BLOCKED!"
        
        Case CVErr(xlErrConnect)
            ErrorText = "#CONNECT!"
            
        Case CVErr(xlErrField)
            ErrorText = "#FIELD!"
        
        Case CVErr(xlErrGettingData)
            ErrorText = "#GETTINGDATA!"
            
        Case Else
            ErrorText = "Unknown error"
    End Select
    
    GetErrorText = ErrorText

End Function


Public Function FindFormulaText(ByVal FromBook As Workbook _
                                , ByVal StartFormulaInSheet As Worksheet _
                                 , ByVal RangeReference As String) As String
    
    ' Finds the formula text for the specified range reference.

    Dim CurrentRange As Range
    Set CurrentRange = RangeResolver.FindRangeFromText(FromBook, StartFormulaInSheet, RangeReference)

    If IsNothing(CurrentRange) Then
        FindFormulaText = vbNullString           ' Return an empty string if the range reference is not found.
    Else
        If CurrentRange.Cells.Count = 1 Then
            Dim Formula As String
            On Error Resume Next
            Formula = CurrentRange.Formula2
            Formula = GetLambdaDefIfLETStepRefCell(CurrentRange, Formula, StartFormulaInSheet)
            FindFormulaText = Formula            ' Return the formula if it's a single cell range.
            On Error GoTo 0
        End If
    End If
    
End Function

' Replace LETStepRef cell with LETStep_FX RefersTo
Public Function GetLambdaDefIfLETStepRefCell(ByVal ForCell As Range _
                                             , ByVal FormulaIfNotLETStepRefCell As String _
                                             , ByVal StartFormulaInSheet As Worksheet) As String
    
    Dim CurrentName As name
    Set CurrentName = FindNamedRangeFromSubCell(ForCell)
    
    If IsNothing(CurrentName) Then
        GetLambdaDefIfLETStepRefCell = FormulaIfNotLETStepRefCell
        Exit Function
    End If
    
    Dim FinalFormula As String
    FinalFormula = FormulaIfNotLETStepRefCell
    
    If Text.IsStartsWith(CurrentName.name, LETSTEPREF_UNDERSCORE_PREFIX) Then
        Dim name As String
        name = VBA.Replace(CurrentName.name, LETSTEPREF_UNDERSCORE_PREFIX, LETSTEP_UNDERSCORE_PREFIX)
        Set CurrentName = FindNamedRange(ForCell.Worksheet.Parent, name)
        If IsNotNothing(CurrentName) Then
            FinalFormula = CurrentName.RefersTo
            FinalFormula = RemoveSheetNameIfPresent(FinalFormula, StartFormulaInSheet.name)
        End If
    End If
    GetLambdaDefIfLETStepRefCell = FinalFormula
    
End Function

' Remove Sheet Name prefix if same sheet for LETStepRef cell
Private Function RemoveSheetNameIfPresent(ByVal FromFormula As String, ByVal SheetName As String) As String
     
    FromFormula = Text.PadIfNotPresent(FromFormula, EQUAL_SIGN)
    RemoveSheetNameIfPresent = RemoveSheetNameFromFormula(FromFormula, SheetName)
    
End Function

Public Function FindRangeLabel(ByVal RangeReference As String, ByVal CurrentCell As Range _
                                                              , ByVal IsJustCheckLabel As Boolean) As String
    
    ' Finds the label for the given range reference.

    Dim FindFromRange As Range
    Dim MatchToRange As Range

    If IsSpilledRangeRef(RangeReference) Then
        ' The range reference is dynamic, remove the trailing '$' sign to get the actual range.
        Set FindFromRange = CurrentCell.Parent.Range(Text.RemoveFromEnd(RangeReference, 1))
        Set MatchToRange = CurrentCell.Parent.Range(RangeReference)
    Else
        ' The range reference is not dynamic, use the current cell' s first cell as the base for finding the range.
        Set FindFromRange = CurrentCell.Cells(1, 1)
        Set MatchToRange = CurrentCell.Cells(1, 1)
    End If

    On Error GoTo LogErrorAndExit
    ' Call the helper function to get the label for the range.
    FindRangeLabel = GetRangeLabelFromNameOrLabel(FindFromRange, MatchToRange, IsJustCheckLabel, CurrentCell)
    Exit Function

LogErrorAndExit:
    FindRangeLabel = vbNullString                ' Return an empty string if an error occurs and log the error.
    Logger.Log ERROR_LOG, Err.Number & Err.Description
    Err.Clear
    
End Function

Public Function IsSpilledRangeRef(ByVal RangeReference As String) As Boolean
    IsSpilledRangeRef = Text.IsEndsWith(RangeReference, DYNAMIC_CELL_REFERENCE_SIGN)
End Function

Public Function GetRangeLabelFromNameOrLabel(ByVal FindFromRange As Range, ByVal MatchToRange As Range _
                                                                          , ByVal IsJustCheckLabel As Boolean _
                                                                           , ByVal CurrentCell As Range) As String
    
    ' Retrieves the label for the given range reference by examining named ranges and let variable names.

    Logger.Log TRACE_LOG, "Enter modUtility.GetRangeLabelFromNameOrLabel"

    ' Check if the current cell represents the header of a table and get the variable name if so.
    If CurrentCell.Cells.Count > 1 And CurrentCell.Cells(1).Address = "$A$1" Then
        GetRangeLabelFromNameOrLabel = modUtility.FindLetVarName(CurrentCell)
        Exit Function
    End If

    ' If the IsJustCheckLabel flag is set, only check and return the variable name without further processing.
    If IsJustCheckLabel Then
        GetRangeLabelFromNameOrLabel = modUtility.FindLetVarName(FindFromRange)
        Exit Function
    End If

    Dim CurrentName As name
    Set CurrentName = modUtility.FindNamedRangeFromSubCell(FindFromRange)

    If IsNotNothing(CurrentName) Then
        ' Use the named range label only if it matches the MatchToRange and the MatchToRange is a multi-cell range.
        If IsBothRangeEqual(CurrentName.RefersToRange, MatchToRange) And MatchToRange.Cells.Count > 1 Then
            GetRangeLabelFromNameOrLabel = modUtility.ExtractNameFromLocalNameRange(CurrentName.name)
        Else
            GetRangeLabelFromNameOrLabel = modUtility.FindLetVarName(FindFromRange) ' Otherwise, get the variable name.
        End If
    Else
        GetRangeLabelFromNameOrLabel = modUtility.FindLetVarName(FindFromRange) ' If not a named range, get the variable name.
    End If

    Logger.Log TRACE_LOG, "Exit modUtility.GetRangeLabelFromNameOrLabel"
    
End Function

Public Function GetNameInFormula(ByVal GivenCells As Range) As String
    
    ' Gets the name of the range in the formula if it is a structured reference in a table.

    If IsInsideTable(GivenCells) Then
        GetNameInFormula = ConvertToStructuredReference(GivenCells, GivenCells) ' Convert to a structured reference.
    Else
        GetNameInFormula = GivenCells.Address    ' Otherwise, get the cell address as the name in the formula.
    End If
    
End Function

Public Function GetNumberOfNonInputDependency(ByVal DependencyObjects As Collection) As Long
    
    ' Counts the number of dependencies that are not marked as input cells.

    Dim CurrentDependencyInfo As DependencyInfo
    Dim Counter As Long

    For Each CurrentDependencyInfo In DependencyObjects
        With CurrentDependencyInfo
            If Not .IsLabelAsInputCell Then
                Counter = Counter + 1
            End If
        End With
    Next CurrentDependencyInfo

    GetNumberOfNonInputDependency = Counter      ' Return the count of non-input dependencies.
    
End Function

Public Function GetNumberOfNonInputLetStepDependency(ByVal DependencyObjects As Collection) As Long
    
    ' Counts the number of dependencies that are not marked as input cells and not marked as not "Let" statements by the user.

    Dim CurrentDependencyInfo As DependencyInfo
    Dim Counter As Long

    For Each CurrentDependencyInfo In DependencyObjects
        With CurrentDependencyInfo
            If Not .IsLabelAsInputCell And Not .IsMarkAsNotLetStatementByUser Then
                Counter = Counter + 1
            End If
        End With
    Next CurrentDependencyInfo

    ' Get the last item from the collection.
    Set CurrentDependencyInfo = DependencyObjects.Item(DependencyObjects.Count)

    ' If the count is only for the last step, no need for the "_result" step.
    If Counter = 1 And Not CurrentDependencyInfo.IsLabelAsInputCell _
       And Not CurrentDependencyInfo.IsMarkAsNotLetStatementByUser Then
        Counter = 0
    End If

    GetNumberOfNonInputLetStepDependency = Counter ' Return the count of non-input "Let" step dependencies.
    
End Function

Public Function IsCurrentDependencyOptional(ByVal ForDependency As DependencyInfo _
                                            , ByVal ObjectsCollection As Collection) As Boolean
    
    ' Checks if the current dependency is optional based on the formula text and the name in the formula.

    Dim CurrentDependency As DependencyInfo
    For Each CurrentDependency In ObjectsCollection
        If modUtility.IsOptionalArgument(CurrentDependency.FormulaText, ForDependency.NameInFormula) Then
            IsCurrentDependencyOptional = True   ' Return True if the current dependency is optional.
            Exit Function
        End If
    Next CurrentDependency

    IsCurrentDependencyOptional = False          ' Return False if the current dependency is not optional.
    
End Function

Public Function IsRelativeFormulaEqual(ByVal CurrentText As String _
                                       , ByVal NameInFormula As String _
                                        , ByVal SheetName As String) As Boolean

    CurrentText = RTrim$(LTrim$(CurrentText))
    NameInFormula = RTrim$(LTrim$(NameInFormula))
    Dim SheetNameQualifiedRangeRef As String
    SheetNameQualifiedRangeRef = GetSheetRefForRangeReference(SheetName, False) & NameInFormula

    If RemoveDollarSign(CurrentText) = RemoveDollarSign(NameInFormula) Then
        IsRelativeFormulaEqual = True
    ElseIf RemoveDollarSign(CurrentText) = RemoveDollarSign(SheetNameQualifiedRangeRef) Then
        IsRelativeFormulaEqual = True
    Else
        IsRelativeFormulaEqual = False
    End If

End Function

Public Function RemoveInitialSpaceAndNewLines(ByVal OperationOnText As String) As String

    Dim Index As Long
    Dim CurrentChar As String
    For Index = 1 To VBA.Len(OperationOnText)
        CurrentChar = VBA.Mid$(OperationOnText, Index, 1)
        If Not (CurrentChar = ONE_SPACE Or CurrentChar = vbNewLine Or CurrentChar = VBA.Chr$(10)) Then
            Exit For
        End If
    Next Index
    RemoveInitialSpaceAndNewLines = VBA.Mid$(OperationOnText, Index)

End Function

Public Function FromRange(ByVal OperationOnText As String _
                          , ByVal StartIndex As Long, ByVal EndIndex As Long) As String
    
    If StartIndex <= 0 Or EndIndex <= 0 Or StartIndex > EndIndex _
       Or StartIndex > Len(OperationOnText) Or EndIndex > Len(OperationOnText) Then
        Err.Raise 13, , "Text.FromRange", "Invalid input argument"
    End If
    FromRange = Mid$(OperationOnText, StartIndex, EndIndex - StartIndex + 1)
    
End Function

Public Function IsEqual(ByVal FirstText As String _
                        , ByVal SecondText As String _
                         , Optional ByVal ComparisionType As Comparer = CONSIDER_CASE) As Boolean
    
    Dim ComparingOption As VbCompareMethod
    ComparingOption = IIf(ComparisionType = IGNORE_CASE, vbTextCompare, vbBinaryCompare)
    IsEqual = (StrComp(FirstText, SecondText, ComparingOption) = 0)
    
End Function

Public Function EncloseWithParenForMultiTerm(ByVal FormulaText As String) As String
    
    ' Encloses the formula text with parentheses if it contains any multi-term operator.

    If Text.IsStartsWith(LTrim$(FormulaText), FIRST_PARENTHESIS_OPEN) _
       And Text.IsEndsWith(RTrim$(FormulaText), FIRST_PARENTHESIS_CLOSE) Then
        ' If the formula text is already enclosed with parentheses, return it as it is.
        EncloseWithParenForMultiTerm = FormulaText
    Else
        If Text.IsAnyDelimiterExists(FormulaText, Array("+", "-", "/", "*")) Then
            ' If the formula contains any multi-term operator, enclose it with parentheses.
            EncloseWithParenForMultiTerm = FIRST_PARENTHESIS_OPEN & FormulaText & FIRST_PARENTHESIS_CLOSE
        Else
            ' If there are no multi-term operators, return the formula text as it is.
            EncloseWithParenForMultiTerm = FormulaText
        End If
    End If
    
End Function

Public Function GetAllDependencyToArrayAndSort(ByVal AllDependency As Collection) As Variant
    
    ' Retrieves all dependency data from the object collection, converts it to an array, and sorts it.

    Dim Result As Variant
    Result = GetDependencyDataFromObjects(AllDependency, True) ' Get all dependency data as an array.
    ' We are not separating header and number as "Level" will always come at the top before any number in level header.
    GetAllDependencyToArrayAndSort = SortDependencyDataUsingSortFunction(Result) ' Sort the dependency data array.
    
End Function

Private Function GetDependencyDataFromObjects(ByVal AllDependency As Collection _
                                              , ByVal IsWithHeader As Boolean) As Variant
    
    Dim PropertiesName As Collection
    Set PropertiesName = DependencyInfoObjectPropertiesName()
    GetDependencyDataFromObjects = modUtility.GetObjectsPropertyValue(AllDependency _
                                                                      , PropertiesName, IsWithHeader)
    
End Function

Private Function SortDependencyDataUsingSortFunction(ByVal InputData As Variant) As Variant
    
    Logger.Log TRACE_LOG, "Enter FormulaParser.SortDependencyDataUsingSortFunction"
    Dim ByColumns As Variant
    ReDim ByColumns(1 To 6)
    ByColumns(1) = SortByColumn.Level
    ByColumns(2) = SortByColumn.IsLabelAsInputCell
    ByColumns(3) = SortByColumn.OPTIONAL_ARGUMENT
    ByColumns(4) = SortByColumn.SheetName
    ByColumns(5) = SortByColumn.ColumnNumber
    ByColumns(6) = SortByColumn.RowNumber
    
    ' -1=Descending, 1=Ascending
    Const SORT_TYPE As String = "-1,-1,1,1,1,1"
    Dim SortType As Variant
    SortType = Split(SORT_TYPE, COMMA)
    SortDependencyDataUsingSortFunction = Application.WorksheetFunction.Sort(InputData, ByColumns, SortType)
    Logger.Log TRACE_LOG, "Exit FormulaParser.SortDependencyDataUsingSortFunction"
    
End Function

Public Function GetSpillParentAddress(ByVal ForCell As Range, Optional ByVal IsAbsolute As Boolean = True) As String
    
    If ForCell.Cells(1).HasSpill Then
        GetSpillParentAddress = GetRangeReference(GetSpillParentCell(ForCell), IsAbsolute)
    End If
    
End Function

Public Function GetSpillParentCell(ByVal ForCell As Range) As Range
    
    If ForCell.Cells(1).HasSpill Then
        Set GetSpillParentCell = ForCell.Cells(1).SpillParent
    End If
    
End Function

Public Sub UpdateFormulaAndCalculate(ByVal OnCell As Range, ByVal FormulaText As String)
    
    OnCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(FormulaText)
    OnCell.Calculate
    
End Sub

Public Function GetInputCellsVarNameAndRangeReference(ByVal DependencyObjects As Collection) As Variant
    
    ' Retrieves variable names and range references for input cells from the given dependency object collection.

    Dim InputCellsVarName As Collection
    Set InputCellsVarName = New Collection

    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In DependencyObjects
        ' Check if the current dependency info represents an input cell (user-defined label).
        If CurrentDependencyInfo.IsLabelAsInputCell Then
            ' Add the current dependency info to the collection of input cells.
            InputCellsVarName.Add CurrentDependencyInfo
        End If
    Next CurrentDependencyInfo

    ' Get the variable names and range references for the input cells.
    GetInputCellsVarNameAndRangeReference = GetVarNameAndRangeReference(InputCellsVarName)
    
End Function

Public Function GetNonInputLetStepsVarNameAndRangeReference(ByVal DependencyObjects As Collection) As Variant
    
    ' Retrieves variable names and range references for non-input let step cells from the given dependency object collection.

    Dim LetStepsVarName As Collection
    Set LetStepsVarName = New Collection

    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In DependencyObjects
        ' Check if the current dependency info represents a non-input let step cell.
        If Not CurrentDependencyInfo.IsLabelAsInputCell _
           And Not CurrentDependencyInfo.IsMarkAsNotLetStatementByUser Then
            ' Add the current dependency info to the collection of non-input let step cells.
            LetStepsVarName.Add CurrentDependencyInfo
        End If
    Next CurrentDependencyInfo

    ' Get the variable names and range references for the non-input let step cells.
    GetNonInputLetStepsVarNameAndRangeReference = GetVarNameAndRangeReference(LetStepsVarName)
    
End Function

Public Function GetLetStepsVarNameAndRangeReference(ByVal DependencyObjects As Collection) As Variant
    
    ' Retrieves variable names and range references for non-input let step cells from the given dependency object collection.

    Dim LetStepsVarName As Collection
    Set LetStepsVarName = New Collection

    Dim CurrentDependencyInfo As DependencyInfo
    For Each CurrentDependencyInfo In DependencyObjects
        ' Check if the current dependency info represents a non-input let step cell.
        If Not CurrentDependencyInfo.IsMarkAsNotLetStatementByUser Then
            ' Add the current dependency info to the collection of non-input let step cells.
            LetStepsVarName.Add CurrentDependencyInfo
        End If
    Next CurrentDependencyInfo

    ' Get the variable names and range references for the non-input let step cells.
    GetLetStepsVarNameAndRangeReference = GetVarNameAndRangeReference(LetStepsVarName)
    
End Function

Private Function GetVarNameAndRangeReference(ByVal FromDependencyColl As Collection) As Variant
    
    ' Retrieves variable names and range references from the given dependency object collection.

    If FromDependencyColl.Count = 0 Then Exit Function
    Dim Result As Variant
    ReDim Result(1 To FromDependencyColl.Count, 1 To 2)
    Dim CurrentDependencyInfo As DependencyInfo
    Dim Counter As Long
    For Each CurrentDependencyInfo In FromDependencyColl
        Counter = Counter + 1
        Result(Counter, 1) = CurrentDependencyInfo.ValidVarName
        Result(Counter, 2) = CurrentDependencyInfo.RangeReference
    Next CurrentDependencyInfo

    ' Return the array containing variable names and range references.
    GetVarNameAndRangeReference = Result
    
End Function

Public Sub SwapTwoRowsInPlace(ByRef InputArray As Variant, ByVal FirstRowIndex As Long, ByVal SecondRowIndex As Long)
    
    ' Swaps two rows in the 2D array in-place.

    Dim Temp As Variant
    Dim ColumnIndex As Long
    For ColumnIndex = LBound(InputArray, 2) To UBound(InputArray, 2)
        ' Swap the values in the two rows for the current column.
        Temp = InputArray(FirstRowIndex, ColumnIndex)
        InputArray(FirstRowIndex, ColumnIndex) = InputArray(SecondRowIndex, ColumnIndex)
        InputArray(SecondRowIndex, ColumnIndex) = Temp
    Next ColumnIndex
    
End Sub

Public Function RemoveDollarSign(ByVal RangeAddress As String) As String
    RemoveDollarSign = VBA.Replace(RangeAddress, DOLLAR_SIGN, vbNullString)
End Function

Public Function RemoveHashSignFromEnd(ByVal NameInFormulaOrRangeReference As String) As String
    RemoveHashSignFromEnd = Text.RemoveFromEndIfPresent(NameInFormulaOrRangeReference, HASH_SIGN)
End Function

Public Function IsWorksheetProtected(ByVal CheckForSheet As Worksheet) As Boolean
    
    With CheckForSheet
        IsWorksheetProtected = (.ProtectContents Or .ProtectDrawingObjects Or .ProtectScenarios)
    End With
    
End Function

Public Function IsWorkbookProtected(ByVal CheckForWorkbook As Workbook) As Boolean
    
    With CheckForWorkbook
        IsWorkbookProtected = (.ProtectWindows Or .ProtectStructure)
    End With
    
End Function

Public Function IsInvalidToRunCommand(ByVal ForCell As Range, ByVal CommandName As String) As Boolean
    
    ' Checks if it is invalid to run the specified command on the given cell.

    IsInvalidToRunCommand = True

    If IsWorkbookProtected(ForCell.Worksheet.Parent) Then
        ' Display a message if the workbook is protected.
        MsgBox "Unable to run '" & CommandName & "' command on a protected workbook. Unprotect the workbook and try again." _
               , vbExclamation + vbOKOnly, CommandName
    ElseIf IsWorksheetProtected(ForCell.Worksheet) Then
        ' Display a message if the worksheet is protected.
        MsgBox "Unable to run '" & CommandName & "' command on a protected worksheet. Unprotect the worksheet and try again." _
               , vbExclamation + vbOKOnly, CommandName
    Else
        ' If all conditions are met, the command can be executed.
        IsInvalidToRunCommand = False
    End If
    
End Function

Public Sub MoveColumnToRightOfScreen(ByVal StartCell As Range)
    
    ' Moves the column containing the StartCell to the right of the visible screen.

    If Not Intersect(ActiveWindow.VisibleRange, StartCell.Offset(0, 1)) Is Nothing Then
        ' If the column to the right of StartCell is already visible, select StartCell and exit the sub.
        StartCell.Select
        Exit Sub
    End If
    
    On Error GoTo LogErrorAndResetBack
    Dim PreviousStatus As Boolean
    PreviousStatus = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.GoTo StartCell, True
    Dim Count As Long
    Dim Temp As Range
    Set Temp = StartCell
    Do While Temp.Column > 1
        If Intersect(ActiveWindow.VisibleRange, StartCell) Is Nothing Then
            ' If StartCell is not in the visible range, exit the loop.
            ' -1 for StartCell is not in the visible range now.
            ' -1 for possibly partial of StartCell after subtracting -1
            Count = Count - 2
            Exit Do
        End If
        Count = Count + 1
        Set Temp = StartCell.Offset(0, -1 * Count)
        Application.GoTo Temp, True
    Loop

    If Count < StartCell.Column Then
        ' If the count is less than the StartCell column, move to the last visible column.
        Set Temp = StartCell.Offset(0, -1 * Count)
        Application.GoTo Temp, True
    End If
    If Temp.Row <> 1 Then Application.GoTo Temp.Offset(-1, 0), True
    StartCell.Select
    Application.ScreenUpdating = PreviousStatus
    Exit Sub
    
LogErrorAndResetBack:
    ' Log any error and reset the screen updating status.
    Logger.Log ERROR_LOG, Err.Number & "-" & Err.Description
    Application.ScreenUpdating = PreviousStatus
    
End Sub

Public Sub DeleteLETStepNamedRangesHavingError(Optional ByVal FromWorkbook As Workbook = Nothing)
    
    ' Deletes LET step named ranges that have #REF! errors.

    If IsNothing(FromWorkbook) Then
        Set FromWorkbook = ActiveWorkbook
    End If
    
    Dim CurrentName As name
    For Each CurrentName In FromWorkbook.Names
        If CurrentName.Visible Then
            If Text.Contains(CurrentName.RefersTo, REF_ERR_KEYWORD) _
               And Text.IsStartsWith(CurrentName.name, LETSTEPREF_PREFIX) Then
                On Error Resume Next
                Dim LetStep_FX_Name As String
                LetStep_FX_Name = LETSTEP_UNDERSCORE_PREFIX _
                                  & Text.AfterDelimiter(CurrentName.name, UNDER_SCORE)
                CurrentName.Delete
                Set CurrentName = FromWorkbook.Names(LetStep_FX_Name)
                CurrentName.Delete
                On Error GoTo 0
            End If
        End If
    Next CurrentName
    
End Sub

Public Function IsNamedRangeExist(ByVal SearchInBook As Workbook _
                                  , ByVal NameOfTheNamedRange As String) As Boolean
    
    ' Checks if a named range exists in the given workbook.

    Dim IsExist As Boolean
    Dim CurrentName As name
    For Each CurrentName In SearchInBook.Names
        If CurrentName.name = NameOfTheNamedRange Then
            IsExist = True
            Exit For
        End If
    Next CurrentName
    IsNamedRangeExist = IsExist
    
End Function

Public Function IsLocalScopedNamedRangeExist(ScopeSheet As Worksheet _
                                             , NamedRangeName As String) As Boolean
    
    Dim SheetQualifiedName As String
    SheetQualifiedName = NamedRangeName
    If Not Text.Contains(NamedRangeName, SHEET_NAME_SEPARATOR) Then
        SheetQualifiedName = GetSheetRefForRangeReference(ScopeSheet.name, False) & NamedRangeName
    End If
    
    Dim CurrentName As name
    For Each CurrentName In ScopeSheet.Names
        If CurrentName.name = SheetQualifiedName Then
            IsLocalScopedNamedRangeExist = True
            Exit Function
        End If
    Next CurrentName
    
    IsLocalScopedNamedRangeExist = False
    
End Function


Public Function AddNRowsTo2DArray(ByVal SourceArray As Variant, ByVal N As Long) As Variant
    
    Dim ResultArray() As Variant
    
    ReDim ResultArray(LBound(SourceArray, 1) To UBound(SourceArray, 1) + N, _
                      LBound(SourceArray, 2) To UBound(SourceArray, 2))
    
    Dim CurrentRow As Long, CurrentCol As Long
    For CurrentRow = LBound(SourceArray, 1) To UBound(SourceArray, 1)
        For CurrentCol = LBound(SourceArray, 2) To UBound(SourceArray, 2)
            ResultArray(CurrentRow, CurrentCol) = SourceArray(CurrentRow, CurrentCol)
        Next CurrentCol
    Next CurrentRow
    
    AddNRowsTo2DArray = ResultArray
    
End Function


Public Function ConvertDependencisToFullyQualifiedRef(ByVal FormulaText As String _
                                                      , ByVal FormulaInSheet As Worksheet) As String
    
    Dim Precedencies As Variant
    Precedencies = GetDirectPrecedents(FormulaText, FormulaInSheet)
    
    'If it is only one dependencies then it will return a 1D Array. And if error then it will have error.
    If LBound(Precedencies) = UBound(Precedencies) Then
        If IsError(Precedencies(LBound(Precedencies, 1), LBound(Precedencies, 2))) Then
            ConvertDependencisToFullyQualifiedRef = FormulaText
        End If
    End If
    
    Dim UpdatedFormula As String
    UpdatedFormula = FormulaText
    Dim CurrentPrecedency As Variant
    For Each CurrentPrecedency In Precedencies
        Dim TempRange As Range
        Set TempRange = RangeResolver.GetRange(CurrentPrecedency, FormulaInSheet.Parent, FormulaInSheet)
        If IsNotNothing(TempRange) Then
            If IsCellAddressUsed(TempRange, CurrentPrecedency) Then
                Dim FullyQualifiedRef As String
                FullyQualifiedRef = GetRangeRefWithSheetName(TempRange, True)
                UpdatedFormula = ReplaceCellRefWithStepName(UpdatedFormula, FullyQualifiedRef _
                                                                       , CStr(CurrentPrecedency), FormulaInSheet.name)
            End If
            
        End If
        Set TempRange = Nothing
        
    Next CurrentPrecedency
    ConvertDependencisToFullyQualifiedRef = UpdatedFormula
    
End Function

Private Function IsCellAddressUsed(ByVal ForRange As Range, ByVal CurrentPrecedency As String) As Boolean
    
    Dim CleanPrecedency As String
    CleanPrecedency = Replace(CurrentPrecedency, DOLLAR_SIGN, vbNullString)
    
    Dim SheetQualifiedRef As String
    SheetQualifiedRef = GetRangeRefWithSheetName(ForRange, False)
    
    IsCellAddressUsed = ( _
                        SheetQualifiedRef = CleanPrecedency _
                        Or ForRange.Address(False, False) = CleanPrecedency _
                        )
    
End Function

Private Sub FindAndReplaceInFirstCol(ByRef FromArray As Variant _
                                     , ByVal FindText As String _
                                      , ByVal ReplaceWith As String _
                                       , ByVal SheetName As String)
    
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(FromArray, 2)
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(FromArray, 1) To UBound(FromArray, 1)
        
        If IsRelativeFormulaEqual(CStr(FromArray(CurrentRowIndex, FirstColumnIndex)), FindText, SheetName) Then
            FromArray(CurrentRowIndex, FirstColumnIndex) = ReplaceWith
        End If
    Next CurrentRowIndex
    
End Sub

Public Function IsArrayAllocated(ByVal Arr As Variant) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsArrayAllocated
    ' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
    ' sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
    ' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
    ' allocated.
    '
    ' The VBA IsArray function indicates whether a variable is an array, but it does not
    ' distinguish between allocated and unallocated arrays. It will return TRUE for both
    ' allocated and unallocated arrays. This function tests whether the array has actually
    ' been allocated.
    '
    ' This function is just the reverse of IsArrayEmpty.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim N As Long
    On Error Resume Next

    ' if Arr is not an array, return FALSE and get out.
    If IsArray(Arr) = False Then
        IsArrayAllocated = False
        Exit Function
    End If

    ' Attempt to get the UBound of the array. If the array has not been allocated,
    ' an error will occur. Test Err.Number to see if an error occurred.
    N = UBound(Arr, 1)
    If (Err.Number = 0) Then
        ''''''''''''''''''''''''''''''''''''''
        ' Under some circumstances, if an array
        ' is not allocated, Err.Number will be
        ' 0. To acccomodate this case, we test
        ' whether LBound <= Ubound. If this
        ' is True, the array is allocated. Otherwise,
        ' the array is not allocated.
        '''''''''''''''''''''''''''''''''''''''
        If LBound(Arr) <= UBound(Arr) Then
            ' no error. array has been allocated.
            IsArrayAllocated = True
        Else
            IsArrayAllocated = False
        End If
    Else
        ' error. unallocated array
        IsArrayAllocated = False
    End If

End Function

Public Function RemoveTopRowHeader(ByVal InputArray As Variant) As Variant

    ' Check if the input is an array
    If Not IsArrayAllocated(InputArray) Then
        RemoveTopRowHeader = InputArray
        Exit Function
    End If

    ' Declare variable for number of rows
    Dim NumRows As Long
    Dim CurrentRow As Long
    Dim CurrentCol As Long

    ' Get the number of rows in the input array
    NumRows = UBound(InputArray, 1) - LBound(InputArray, 1) + 1

    ' Check if the input array has more than one row
    If NumRows <= 1 Then
        RemoveTopRowHeader = Empty
        Exit Function
    End If

    ' Declare a result array without the top row, using the same lower bounds as the input array
    Dim ResultArray() As Variant
    ReDim ResultArray(LBound(InputArray, 1) To UBound(InputArray, 1) - 1 _
                      , LBound(InputArray, 2) To UBound(InputArray, 2))

    ' Copy the content without the top row
    For CurrentRow = LBound(InputArray, 1) + 1 To UBound(InputArray, 1)
        For CurrentCol = LBound(InputArray, 2) To UBound(InputArray, 2)
            ResultArray(CurrentRow - 1, CurrentCol) = InputArray(CurrentRow, CurrentCol)
        Next CurrentCol
    Next CurrentRow

    ' Assign the result to the function's return value
    RemoveTopRowHeader = ResultArray

End Function

Public Function IsExpandAble(ByVal ForCell As Range) As Boolean

    ' Check if the input cell is not null
    If IsNothing(ForCell) Then Exit Function

    Dim CurrentName As name
    ' Find the named range that includes the input cell
    Set CurrentName = FindNamedRangeFromSubCell(ForCell)
    ' Check if the current named range is not null
    If IsNotNothing(CurrentName) Then
        ' Check if the address of the current named range is same as input cell address
        If CurrentName.RefersToRange.Address <> ForCell.Address Then Exit Function

        ' Check if the input cell has only one cell and that cell has a formula
    ElseIf ForCell.Cells.Count = 1 And ForCell.HasFormula Then
        ' If true, the cell is expandable
        IsExpandAble = True
        ' Check if the input cell is part of a spill range
    ElseIf ForCell.HasSpill Then
        ' If true, the cell is expandable
        IsExpandAble = True
    End If

End Function

Public Sub WriteStringToTextFile(Content As String, ToFilePath As String)
    
    Dim FileNo As Long
    FileNo = FreeFile()
    Open ToFilePath For Output As #FileNo
    Print #FileNo, Content
    Close #FileNo
    
End Sub

Public Function ConcatenateCollection(ByVal GivenCollection As Collection _
                                      , Optional ByVal Delimiter As String = ",") As String
    
    Dim Result As String
    Dim CurrentItem As Variant
    For Each CurrentItem In GivenCollection
        Result = Result & CStr(CurrentItem) & Delimiter
    Next CurrentItem
    
    If Result = vbNullString Then
        ConcatenateCollection = vbNullString
    Else
        ConcatenateCollection = Left$(Result, Len(Result) - Len(Delimiter))
    End If
    
End Function

Public Function PutSpaceOnLowerCaseToUpperCaseTransition(ByVal CurrentWord As String) As String
    
    Dim Result As String
    Dim Index As Long
    Dim CurrentCharacter As String
    Dim NextCharacter As String
    For Index = 1 To Len(CurrentWord) - 1
        CurrentCharacter = Mid$(CurrentWord, Index, 1)
        NextCharacter = Mid$(CurrentWord, Index + 1, 1)
        Result = Result & CurrentCharacter
        If Not IsCapitalLetter(CurrentCharacter) And IsAlphabet(CurrentCharacter) _
           And IsCapitalLetter(NextCharacter) Then
            Result = Result & ONE_SPACE
        End If
    Next Index
    If CurrentWord <> vbNullString Then Result = Result & Right$(CurrentWord, 1)
    PutSpaceOnLowerCaseToUpperCaseTransition = Result
    
End Function

'PutSpaceBeforeLastCapsFromStart("CASE%$Rules") >> "CASE%$ Rules"
Public Function PutSpaceBeforeLastCapsFromStart(ByVal CurrentWord As String) As String
    
    If CurrentWord = vbNullString Then Exit Function
    If IsAllCaps(CurrentWord) Then
        PutSpaceBeforeLastCapsFromStart = CurrentWord
        Exit Function
    End If
    
    
    Dim Index As Long
    Dim CurrentCharacter As String
    If Not IsCapitalLetter(Left$(CurrentWord, 1)) Then
        PutSpaceBeforeLastCapsFromStart = CurrentWord
        Exit Function
    End If
    
    Dim LowerCaseCharIndex As Long
    For Index = 2 To Len(CurrentWord)
        CurrentCharacter = Mid$(CurrentWord, Index, 1)
        If Not IsCapitalLetter(CurrentCharacter) And IsAlphabet(CurrentCharacter) Then
            LowerCaseCharIndex = Index
            Exit For
        End If
    Next Index
    
    Dim Result As String
    If LowerCaseCharIndex < 3 Then
        Result = CurrentWord
    Else
        Result = Left(CurrentWord, LowerCaseCharIndex - 2) _
                 & ONE_SPACE & Mid(CurrentWord, LowerCaseCharIndex - 1)
    End If
    
    PutSpaceBeforeLastCapsFromStart = Result
    
End Function

Public Function IsCapitalLetter(ByVal GivenLetter As String) As Boolean
    
    If Len(GivenLetter) > 1 Then
        Err.Raise 13, "IsCapitalLetter Function", "Given Letter need to be one character String"
    End If
    If GivenLetter = vbNullString Then
        Err.Raise 5, "IsCapitalLetter Function", "Given Letter can't be nullstring"
    End If

    Const ASCII_CODE_FOR_A As Integer = 65
    Const ASCII_CODE_FOR_Z As Integer = 90
    Dim ASCIICodeForGivenLetter As Integer
    ASCIICodeForGivenLetter = Asc(GivenLetter)
    IsCapitalLetter = (ASCIICodeForGivenLetter >= ASCII_CODE_FOR_A _
                       And ASCIICodeForGivenLetter <= ASCII_CODE_FOR_Z)

End Function

Public Function IsAlphabet(Char As String) As Boolean
    
    Dim CharCode As Long
    CharCode = Asc(LCase(Char))
    IsAlphabet = (CharCode >= Asc("a") And CharCode <= Asc("z"))
    
End Function

Public Function GetOAParamValue(ParamName As String, DefaultValue As Variant) As Variant
    
    Dim OARobotAddIn As Object
    Set OARobotAddIn = CreateObject("OARobot.ExcelAddin")
    On Error GoTo ReturnDefaultValue
    GetOAParamValue = OARobotAddIn.GetParamValueByName(ParamName)
    Set OARobotAddIn = Nothing
    Exit Function
    
ReturnDefaultValue:
    GetOAParamValue = DefaultValue
    
End Function

Public Function IsSpillParentIncluded(ByVal CheckOnCell As Range) As Boolean
    
    If Not CheckOnCell.HasSpill Then
        IsSpillParentIncluded = False
    ElseIf IsNotNothing(Intersect(CheckOnCell.Cells(1).SpillParent, CheckOnCell)) Then
        IsSpillParentIncluded = True
    End If
    
End Function

Public Function IsRefersToRangeIsNothing(ByVal CurrentName As name) As Boolean
    
    On Error GoTo ErrorHandler
    IsRefersToRangeIsNothing = IsNothing(CurrentName.RefersToRange)
    Exit Function
    
ErrorHandler:
    IsRefersToRangeIsNothing = True
    
End Function

Public Function FindNamedRange(ByVal FromWorkbook As Workbook _
                               , ByVal NameOfTheNamedRange As String) As name
    
    ' Logs entering the function
    Logger.Log TRACE_LOG, "Enter FindNamedRange"
    
    ' Iterates over all names in the workbook
    Dim CurrentNamedRange As name
    For Each CurrentNamedRange In FromWorkbook.Names
        Logger.Log DEBUG_LOG, CurrentNamedRange.name
        ' If the current name is a local scope name and matches the target name (case insensitive), return it
        If IsLocalScopeNamedRange(CurrentNamedRange.NameLocal) Then
            If VBA.UCase$(NameOfTheNamedRange) = VBA.UCase$(ExtractNameFromLocalNameRange(CurrentNamedRange.NameLocal)) Then
                Set FindNamedRange = CurrentNamedRange
                ' Logs exiting the function due to finding the target name
                Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FindNamedRange"
                Exit Function
            End If
            
            ' If the current name matches the target name (case insensitive), return it
        ElseIf VBA.UCase$(NameOfTheNamedRange) = VBA.UCase$(CurrentNamedRange.name) Then
            Set FindNamedRange = CurrentNamedRange
            ' Logs exiting the function due to finding the target name
            Logger.Log TRACE_LOG, "Exit Due to Exit Keyword FindNamedRange"
            Exit Function
        End If
    Next CurrentNamedRange
    
    ' Logs exiting the function without finding the target name
    Logger.Log TRACE_LOG, "Exit FindNamedRange"
    
End Function

Public Function MakeValidDefinedName(ByVal GivenDefinedName As String _
                                     , ByVal IsCapitalizeFirstCharOfWords As Boolean _
                                      , Optional ByVal IsFinal As Boolean) As String
    
    ' Logging entry into function
    Logger.Log TRACE_LOG, "Enter MakeValidDefinedName"

    ' Check if the GivenDefinedName is an empty or blank string
    If Trim$(GivenDefinedName) = vbNullString Then
        ' If IsFinal is True, assign "_Blank" to MakeValidDefinedName, this is used when the input GivenDefinedName is blank and we are finalizing the name
        If IsFinal Then MakeValidDefinedName = "_Blank"
        ' Log the exit due to the Exit Function statement
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword MakeValidDefinedName"
        ' Exit the function when GivenDefinedName is blank
        Exit Function
    End If
    
    ' If the GivenDefinedName is not blank, make it a valid name by calling the MakeValidName function from modUtility
    MakeValidDefinedName = MakeValidName(GivenDefinedName, GetNamingConv(False))

    ' Logging exit from function
    Logger.Log TRACE_LOG, "Exit MakeValidDefinedName"
    
End Function

Public Function DuplicateCollection(ByVal FromCollection As Collection _
                                    , ByVal IsShallowCopy As Boolean _
                                    , Optional ByVal IsKeyAndItemSame As Boolean) As Collection
        
    ' If both key and value are not same and need shallow copy then it will lost the key
    ' as key is not readable from Collection
        
    ' Early validation
    If Not IsShallowCopy Then
        Set DuplicateCollection = FromCollection
        Exit Function
    ElseIf FromCollection Is Nothing Then
        Exit Function
    ElseIf FromCollection.Count = 0 Then
        Set DuplicateCollection = New Collection
        Exit Function
    End If
    
    
    Dim Result As Collection
    Set Result = New Collection
    
    Dim CurrentItem As Variant
    For Each CurrentItem In FromCollection
        If IsKeyAndItemSame Then
            Result.Add CurrentItem, CStr(CurrentItem)
        Else
            Result.Add CurrentItem
        End If
    Next CurrentItem
    
    Set DuplicateCollection = Result
    Set Result = Nothing
    
End Function

Public Function FilterUsingSpecialCells(ByVal FromRange As Range _
                                        , ByVal CellType As XlCellType) As Range
    
    Set FilterUsingSpecialCells = Intersect(FromRange, FromRange.SpecialCells(CellType))
    
End Function

Public Sub AddToCollectionIfNotExist(ByRef ToCollection As Collection _
                                      , ByVal Key As String, ByVal Item As Variant)
    
    If Not IsExistInCollection(ToCollection, Key) Then
        ToCollection.Add Item, Key
    End If
    
End Sub

Public Function IsBuiltInName(ByVal CurrentName As name) As Boolean
    
    ' We need to use Name.MacroType to identify if it's built in or not.
    ' Checking visible or not is not ideal scenario. We may have a custom named range but hidden.
    
    Const XL_FUNCTION_MACRO_TYPE As Long = xlFunction ' Example name: _xlfn.HSTACK
    Const XL_PARAM_MACRO_TYPE As Long = xlCommand ' Example name: _xlpm.Curr
    IsBuiltInName = ( _
                    CurrentName.MacroType = XL_FUNCTION_MACRO_TYPE _
                    Or CurrentName.MacroType = XL_PARAM_MACRO_TYPE _
                    )
        
End Function

Public Function IsFilePath(ByVal GivenPath As String) As Boolean
    
    '@Description("Check if a path is File Path or not")
    '@Dependency("No Dependency")
    '@ExampleCall : IsFilePath("C:\Users\USER\Documents\Compare Folder - Copy.xlsm")
    '@Date : 23 November 2022 06:27:01 PM
    '@PossibleError:

    IsFilePath = (Dir(GivenPath, vbNormal) = Dir(GivenPath, vbDirectory) _
                  And Dir(GivenPath, vbNormal) <> vbNullString)

End Function

Public Function GetFileName(ByVal FilePath As String) As String

    Dim LastSeparatorIndex As Long
    LastSeparatorIndex = InStrRev(FilePath, Application.PathSeparator)
    If LastSeparatorIndex = 0 Then Exit Function
    GetFileName = Mid(FilePath, LastSeparatorIndex + 1)

End Function

Public Function ReplaceNewlineWithChar10(ByVal OnText As String) As String
    ReplaceNewlineWithChar10 = VBA.Replace(OnText, vbNewLine, Chr$(10))
End Function

Public Sub CopyDataToClipBoard(ByVal GivenText As String)
    
    Dim TotalWait As Long
    On Error GoTo HandleError
    CreateObject("htmlfile").parentWindow.clipboardData.SetData "text", GivenText
    Exit Sub
HandleError:
    If Err.Number = -2147352319 And Err.Description = "Automation error" Then
        Debug.Print "Wait for one sec."
        TotalWait = TotalWait + 1
        Sleep 1000
        If TotalWait > 5 Then Exit Sub
        DoEvents
        Resume
    End If
    
End Sub

Public Function IsSheetExist(ByVal SheetTabName As String _
                             , Optional ByVal GivenWorkbook As Workbook) As Boolean

    '@Description("This function will determine if a sheet is exist or not by using tab name")
    '@Dependency("No Dependency")
    '@ExampleCall : IsSheetExist("SheetTabName")
    '@Date : 14 October 2021 07:03:05 PM

    If GivenWorkbook Is Nothing Then Set GivenWorkbook = ThisWorkbook

    Dim TemporarySheet As Worksheet
    On Error Resume Next
    Set TemporarySheet = GivenWorkbook.Worksheets(SheetTabName)

    IsSheetExist = (Not TemporarySheet Is Nothing)
    On Error GoTo 0

End Function

Public Function IsOpenWorkbookExists(ByVal BookName As String) As Boolean
    
    
    Dim Result As Boolean
    Dim CurrentBook As Workbook
    For Each CurrentBook In Application.Workbooks
        If CurrentBook.name = BookName Then
            Result = True
            Exit For
        End If
    Next CurrentBook
    
    If Result Then
        IsOpenWorkbookExists = Result
        Exit Function
    End If
    
    Dim CurrentAddIn As AddIn
    For Each CurrentAddIn In Application.AddIns
        If CurrentAddIn.IsOpen And CurrentAddIn.name = BookName Then
            Result = True
            Exit For
        End If
    Next CurrentAddIn
    
    IsOpenWorkbookExists = Result
    
End Function

Public Function IsClosedWorkbookRef(ByVal RangeRef As String) As String
    
    Dim Result As Boolean
    ' One drive or share point location.
    ' example: 'https://d.docs.live.net/6edd704b8f8c537b/TextOffset lambda testing.xlsm'!TestName
    If Text.IsStartsWith(RangeRef, "'https://") Then
        Result = True
    ElseIf Text.Contains(RangeRef, ":\") Then
        ' local drive location:
        ' example: 'D:\Downloads\Email Manager V10.xlsm'!TemplateEmailFilePath
        Result = True
    Else
        Result = False
    End If
    
    IsClosedWorkbookRef = Result
    
End Function

Public Function IsReferenceFromDifferentBook(ByVal PrecedentsRef As String _
                                              , ByVal CheckAgainstBook As Workbook) As Boolean
    
    Dim Result As Boolean
    Result = False
    
    If IsClosedWorkbookRef(PrecedentsRef) Then
        Result = True
    ElseIf Text.Contains(PrecedentsRef, SHEET_NAME_SEPARATOR) Then
        
        Dim SheetName As String
        SheetName = Text.BeforeDelimiter(PrecedentsRef, SHEET_NAME_SEPARATOR, , FROM_END)
        SheetName = Text.RemoveFromBothEndIfPresent(SheetName, SINGLE_QUOTE)
        SheetName = UnEscapeSingleQuote(SheetName)
        
        Dim ResolvedRange As Range
        Set ResolvedRange = RangeResolver.GetRange(PrecedentsRef, CheckAgainstBook)
        If ResolvedRange Is Nothing Then
            If IsNamedRangeExist(CheckAgainstBook, PrecedentsRef) Then
                Result = False
            Else
                Err.Raise 13, "Range Resolver", "Can't find range from PrecedentsRef"
            End If
        Else
            Result = (ResolvedRange.Worksheet.Parent.name <> CheckAgainstBook.name)
        End If
        
    End If
    
    IsReferenceFromDifferentBook = Result
    
End Function

Public Function Max(ByVal FirstNumber As Variant, ByVal SecondNumber As Variant) As Variant
    Max = Application.WorksheetFunction.Max(FirstNumber, SecondNumber)
End Function

Public Function IsSubRange(ByVal ParentRange As Range _
                           , ByVal ChildRange As Range) As Boolean

    If ChildRange Is Nothing Then Exit Function
    If ParentRange Is Nothing Then Exit Function

    Dim InterSectionRange As Range
    Set InterSectionRange = Intersect(ParentRange, ChildRange)

    If InterSectionRange Is Nothing Then Exit Function
    IsSubRange = (ChildRange.Address = InterSectionRange.Address)

End Function

Public Function EscapeSingeQuote(ByVal OnText As String) As String
    EscapeSingeQuote = Replace(OnText, SINGLE_QUOTE, SINGLE_QUOTE & SINGLE_QUOTE)
End Function

Public Function UnEscapeSingleQuote(ByVal OnText As String) As String
    UnEscapeSingleQuote = Replace(OnText, SINGLE_QUOTE & SINGLE_QUOTE, SINGLE_QUOTE)
End Function

Public Function ReplaceInvalidCharFromFormulaWithValid(ByVal Formula As String) As String
    
    Dim Result As String
    Result = Replace(Formula, vbCrLf, vbLf)
    Result = Replace(Result, Chr(160), Chr(32))
    
    ReplaceInvalidCharFromFormulaWithValid = Result
    
End Function
