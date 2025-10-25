VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@GithubRawURL: https://raw.githubusercontent.com/1504168/Text-Utility-VBA/master/Text%20Utility.xlsm%20Definitions/VBA/Text.cls
'@IgnoreModule ParameterNotUsed, AssignmentNotUsed, EmptyMethod, SuperfluousAnnotationArgument, UnrecognizedAnnotation, ProcedureNotUsed, MissingAnnotationArgument
'@PredeclaredId
'@Folder("Utilility.Text")

Option Explicit

Public Enum RelativePosition
    FROM_START = 1
    FROM_end = -1
    FROM_BOTH = 2
End Enum

Public Enum Comparer
    IGNORE_CASE = 0
    CONSIDER_CASE = 1
End Enum

Public Enum Occurrence
    ALL = 0
    LAST = 1
    FIRST = 2
End Enum

'@Description: This function will retrive OperationOnText after that given delimiter. Comparing is case In Sensitive
'@Return: It will return text after a certain delimiter based on Index and SearchDirection param.
'         This will return vbNullString if any argument is vbNullString or if that delimiter isn't present
'@Param:
'       OperationOnText - > A string param containing the string in which we need to find the delimiter.
'       Delimiter -> A string which we need to search in the OperationOnText
'       Index -> Optional parameter which indicate that if same delimiter is found multiple times then which one to consider.
'       SearchDirection - > Optional param which indicate from which side we need to search.
'@ExampleCall : Text.AfterDelimiter("1997ismail.hosen@gmail.com", "@")  >> gmail.com
'               Text.AfterDelimiter("1997ismail.hosen@gmail.com", "#")  >>

Public Function AfterDelimiter(ByVal OperationOnText As String _
                               , ByVal Delimiter As String _
                                , Optional Index As Long = 1 _
                                 , Optional ByVal SearchDirection As RelativePosition = FROM_START) As String
    
    If OperationOnText = vbNullString Or Delimiter = vbNullString Then
        AfterDelimiter = OperationOnText
        Exit Function
    End If
    
    Dim Result As String
    
    Dim AllIndex As Collection
    Set AllIndex = FindAllIndexOf(OperationOnText, Delimiter, SearchDirection)
    
    If AllIndex.Count = 0 Then
        Result = vbNullString
    ElseIf Index > AllIndex.Count Then
        Err.Raise 13, "Text.AfterDelimiter", "There is not that much of data"
    Else
        Dim ItemStartIndex As Long
        ItemStartIndex = AllIndex.Item(Index)
        If ItemStartIndex + Len(Delimiter) - 1 = Len(OperationOnText) Then
            Result = vbNullString
        Else
            Result = SubString(OperationOnText, AllIndex.Item(Index) + Len(Delimiter))
        End If
    End If
    
    AfterDelimiter = Result
    
End Function

Public Function IsEqual(ByVal Text1 As String _
                        , ByVal Text2 As String _
                         , Optional ByVal ComparisionType As Comparer = CONSIDER_CASE) As Boolean
    
    Dim Result As Boolean
    If ComparisionType = CONSIDER_CASE Then
        Result = (Text1 = Text2)
    Else
        Result = (LCase$(Text1) = LCase$(Text2))
    End If
    
    IsEqual = Result
    
End Function

'@Description: This function will retrive all the index of given delimiter OperationOnText. The indexes are always from start .
'              What change based on search direction is how the index is added in the collection. If search direction is FROM_END then collection will have [8,4].
'              So basically we just need to reverse the collection nothing else.
'              So it's better to use helper function to reverse out that instead of using InstrRev function.
'@Return: All the index of given delimiter as Collection.
'@Param:
'       OperationOnText - > A string param containing the string in which we need to find the delimiter.
'       Delimiter       -> A string which we need to search in the OperationOnText
'       SearchDirection -> Optional param which indicate from which side we need to search.
'       ComparisionType -> Optional param(indicate whether matching will be case sensitive or not.
'@ExampleCall : Text.FindAllIndexOf("1997ismail.hosen@gmail.com", "@")  >> gmail.com
'               Text.FindAllIndexOf("1997ismail.hosen@gmail.com", "#")  >>
Public Function FindAllIndexOf(ByVal OperationOnText As String _
                               , ByVal Delimiter As String _
                                , ByVal SearchDirection As RelativePosition _
                                 , Optional ByVal ComparisionType As Comparer = CONSIDER_CASE) As Collection
    
    Dim AllIndex As Collection
    Set AllIndex = New Collection
    
    Dim CurrentIndex As Long
    CurrentIndex = 1
    Dim TextLength As Long
    TextLength = Len(OperationOnText)
    
    Dim ComparingOption As VbCompareMethod
    ComparingOption = GetComparisionOption(ComparisionType)
    
    CurrentIndex = InStr(1, OperationOnText, Delimiter, ComparingOption)
    
    If CurrentIndex <> 0 Then
        AllIndex.Add CurrentIndex
    Else
        Set FindAllIndexOf = New Collection
        Exit Function
    End If
    
    Do While (CurrentIndex <= TextLength)
        CurrentIndex = InStr(CurrentIndex + Len(Delimiter), OperationOnText, Delimiter, ComparingOption)
        If CurrentIndex <> 0 Then
            AllIndex.Add CurrentIndex
        Else
            Exit Do
        End If
    Loop
    
    If SearchDirection = FROM_START Then
        Set FindAllIndexOf = AllIndex
    ElseIf SearchDirection = FROM_end Then
        Set FindAllIndexOf = ReverseIndexCollection(AllIndex)
    Else
        Err.Raise 13, "Text.FindAllIndexOf", "Invalid SearchDirection"
    End If
    
End Function

Private Function GetComparisionOption(ByVal ComparisionType As Comparer) As VbCompareMethod
    
    Dim Result As VbCompareMethod
    
    If ComparisionType = IGNORE_CASE Then
        Result = vbTextCompare
    Else
        Result = vbBinaryCompare
    End If
    
    GetComparisionOption = Result
    
End Function

'@Description: This function will return the character at the Index position of the OperationOnText value.
'              The first character in the OperationOnText is at position 1.
'@Return: It will return the character At the Index position provided as param.
'         This will return vbNullString if argument is vbNullString or out of range
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Index           -> parameter by which we need to specify the Index position of the characters of the OperationOnText that we want to get.

'@ExampleCall : Text.At("1997ismail.hosen@gmail.com", 5)  >> i

Public Function At(ByVal OperationOnText As String, ByVal Index As Long) As String
    At = Mid$(OperationOnText, Index, 1)
End Function

'@Description: This function will retrive the portion of OperationOnText before that given delimiter.
'@Return: It will return text after a certain delimiter based on Index and SearchDirection param.
'         This will return vbNullString if any argument is vbNullString or if that delimiter isn't present
'@Param:
'       OperationOnText - > A string param containing the string in which we need to find the delimiter.
'       Delimiter -> A string which we need to search in the OperationOnText
'       Index -> Optional parameter which indicate that if same delimiter is found multiple times then which one to consider.
'       SearchDirection - > Optional param which indicate from which side we need to search (from the start or end) of the input.
'@ExampleCall : Text.BeforeDelimiter("1997ismail@hosen@gmail.com", "@")  >> 1997ismail.hosen
'               Text.BeforeDelimiter("1997ismail@hosen@gmail.com", "@", 2)  >>1997ismail@hosen

Public Function BeforeDelimiter(ByVal OperationOnText As String _
                                , ByVal Delimiter As String _
                                 , Optional ByVal Index As Long = 1 _
                                  , Optional ByVal SearchDirection As RelativePosition = FROM_START) As String
    
    If OperationOnText = vbNullString Or Delimiter = vbNullString Then
        BeforeDelimiter = OperationOnText
        Exit Function
    ElseIf OperationOnText = Delimiter Then
        BeforeDelimiter = vbNullString
        Exit Function
    End If
    
    Dim AllIndex As Collection
    Set AllIndex = FindAllIndexOf(OperationOnText, Delimiter, SearchDirection)
    If AllIndex.Count = 0 Then
        BeforeDelimiter = vbNullString
        Exit Function
    ElseIf Index > AllIndex.Count Then
        Err.Raise 13, "Text.BeforeDelimiter", "There is not that much of data"
    End If
    
    BeforeDelimiter = SubString(OperationOnText, 1, AllIndex.Item(Index) - 1)
    
End Function

'@Description: This function will retrive the portion of OperationOnText between two given delimiter.
'@Return: It will return text between two certain delimiter based on StartDelimiter and EndDelimiter param.
'         This will return vbNullString if any of the two parameter is absent in OperationOnText.
'@Param:
'       OperationOnText -> A string param containing the string in which we need to find the delimiter.
'       StartDelimiter  -> A specific string from which we will start to retrieve text in the OperationOnText.
'       EndDelimiter    -> A specific string at which we will stop to retrieve text in the OperationOnText.
'       StartIndex      -> Optional param(indicate the starting delimiter count if it have multiple number of delimiters).
'       EndIndex        -> Optional param(indicate the ending delimiter count if it have multiple number of delimiters).
'       StartDelimiterSearchDirection - > Optional param( indicate starting delimiter will start searching from start or end).
'       EndDelimiterSearchDirection - > Optional param( indicate  from which side ending delimiter will search- from start or end).

'@ExampleCall : Text.BetweenDelimiter("1997ismail@hosen#gmail.com", "@","#")  >> hosen
'               Text.BetweenDelimiter("1997@ismail@hosen#gmail.com", "@","#") >> ismail@hosen
'               Text.BetweenDelimiter("1997@ismail@hosen#gmail.com", "@","#", 2)  >>hosen

Public Function BetweenDelimiter(ByVal OperationOnText As String _
                                 , ByVal StartDelimiter As String _
                                  , ByVal EndDelimiter As String _
                                   , Optional StartIndex As Long = 1 _
                                    , Optional EndIndex As Long = 1 _
                                     , Optional ByVal StartDelimiterSearchDirection As RelativePosition = FROM_START _
                                      , Optional ByVal EndDelimiterSearchDirection As RelativePosition = FROM_START) As String
    
    Dim Result As String
    Result = AfterDelimiter(OperationOnText, StartDelimiter, StartIndex, StartDelimiterSearchDirection)
    Result = BeforeDelimiter(Result, EndDelimiter, EndIndex, EndDelimiterSearchDirection)
    BetweenDelimiter = Result
    
End Function

'@Description: This function will provide cleaned code of OperationOnText.
'@ExcelDependency: Need to know all control character and then need to implement that.
'@Return:Returns a OperationOnText value with all control characters of OperationOnText removed.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@ExampleCall : Text.Clean("1997ismail@hosen#gmail.com")  >>

Public Function Clean(ByVal OperationOnText As String) As String
    Clean = Application.WorksheetFunction.Clean(OperationOnText)
End Function

'@Description: It will combine values/texts of array.
'@Return: It will return the result of combining the list of OperationOnText values, texts, into a single OperationOnText value.
'@Param:
'       Texts       -> An array of values/texts that we need to combine.
'       Separator   -> Optional param( use to specify character/string by which it will combine all the provided values).
'@ExampleCall : Text.Combine(Array("1997ismail","hosen","gmail.com"),"/")  >> 1997ismail/hosen/gmail.com
Public Function Combine(ByVal Texts As Variant _
                        , Optional ByVal Separator As String = vbNullString) As String
    
    Combine = ConcatenateArray(Texts, Separator)
    
End Function

'@Description: This function will Detects whether OperationOnText contains the value substring. This function doesn't support wildcards or regular expressions.
'The following built-in comparers are available in the formula language:
'@Return: It will Returns true if the value is found.
'@Param:
'       OperationOnText -> A string param containing the string in which we need to find the substring.
'       SubString       -> A string param containg values that we want to search in OperationOnText.
'       ComparisionType -> Optional param(indicate whether matching will be case sensitive or not.
'@ExampleCall : Text.Contains("1997ismail.hosen@gmail.com", "gmail")  >> true

Public Function Contains(ByVal OperationOnText As String _
                         , ByVal SubString As String _
                          , Optional ByVal ComparisionType As Comparer = IGNORE_CASE) As Boolean
    
    Dim Result As Boolean
    
    If ComparisionType = IGNORE_CASE Then
        Result = (InStr(1, OperationOnText, SubString, vbTextCompare) <> 0)
    ElseIf ComparisionType = CONSIDER_CASE Then
        Result = (InStr(1, OperationOnText, SubString, vbBinaryCompare) <> 0)
    End If
    
    Contains = Result
    
End Function

'@Description: This function will return the last characters from the end of the OperationOnText value.
'@Return: It will return the spcified number of characters from the end.
'         This will return vbNullString if argument is vbNullString or out of range
'@Param:
'       OperationOnText -> A string param containing the string from which we need to extract characters.
'       Count           -> parameter by which we need to specify how many characters we need to extract.
'@ExampleCall : Text.FromEnd("1997ismail.hosen@gmail.com", 3)  >> com

Public Function FromEnd(ByVal OperationOnText As String, ByVal Count As Long) As String
    
    FromEnd = Right$(OperationOnText, Count)

End Function

'@Description: This function will Test if OperationOnText endswith some partial text or not.
'@Return: It will return boolean value of either the OperationOnText is ends with TextToMatch or not.
'         This will return vbNullString if argument is vbNullString.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       TextToMatch     -> parameter that we want to find in the end.
'       ComparisionType -> Optional param which will specify whether it will be case sensitive or not.
'@ExampleCall : text.IsEndsWith("1997ismail.hosen@gmail.com","com")  >> true
'               text.IsEndsWith("1997ismail.hosen@gmail.com","abc")  >> false
         
Public Function IsEndsWith(ByVal OperationOnText As String _
                           , ByVal TextToMatch As String _
                            , Optional ByVal ComparisionType As Comparer = IGNORE_CASE) As Boolean
    
    Dim Result As Boolean
    
    If ComparisionType = CONSIDER_CASE Then
        Result = (Right$(OperationOnText, Len(TextToMatch)) = TextToMatch)
    ElseIf ComparisionType = IGNORE_CASE Then
        Result = (UCase$(Right$(OperationOnText, Len(TextToMatch))) = UCase$(TextToMatch))
    End If
    
    IsEndsWith = Result
    
End Function

'@Description: This funcition will convert any valid data into string.
'@Return: Returns Srting format data of input value.
'@Param:
'       Value -> Any type of data that we need to convert into string.
'@ExampleCall : text.From(#12/02/2022#)  >> 12/2/2022
Public Function From(ByVal Value As Variant) As String
    
    From = CStr(Value)
    
End Function

'@Description: This function will insert texts or values into OperationOnText.
'@Return: The final result after inserting a specific values in OperationOnText.
'@Param:
'       OperationOnText -> A string param containing the string in which we need to insert some text or characters.
'       Offset          -> Index at which we want to insert our value.
'       NewText         -> New string or value that we want to insert into OperationOnText.
'@ExampleCall : Text.Insert("1997ismail@hosen#gmail.com",2 ,"@@@") >> 19@@@97ismail@hosen#gmail.com

Public Function Insert(ByVal OperationOnText As String _
                       , ByVal Offset As Long _
                        , ByVal NewText As String) As String
    
    Dim Result As String
    
    If Offset >= Len(OperationOnText) Then
        Result = OperationOnText & NewText
    ElseIf Offset <= 1 Then
        Result = NewText & OperationOnText
    Else
        Result = SubString(OperationOnText, 1, Offset) _
                 & NewText & SubString(OperationOnText, Offset + 1)
    End If
    
    Insert = Result
    
End Function

'@Description: This funciton will detect the number of characters in OperationOnText.
'@Return:The number of characters in OperationOnText.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@ExampleCall : Text.Length("1997ismail@hosen#gmail.com")  >> 26
Public Function Length(OperationOnText As String) As Long
    Length = Len(OperationOnText)
End Function

'@Description: Will Return the lower case of OperarationOnText
'@Return: Lower case of OperarationOnText
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@ExampleCall : Text.Lower("BANGLADESH")  >> bangladesh
Public Function Lower(OperationOnText As String) As String
    Lower = VBA.LCase$(OperationOnText)
End Function

'@Description: This will retrieve specified number of characters from an specified starting index position.
'@Return: Will return specific number of Count(number of characters) values after starting index from OperationOnText.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Start           ->Starting index from which it will start retrieving.
'       Count           -> Optional param(indicates Number of characters that we want to retrieve)
'@ExampleCall : Text.Middle("12/22/2017 Robin Hossen Bangladesh", 7, 4)  >> 2017
Public Function Middle(ByVal OperationOnText As String _
                       , ByVal Start As Long _
                        , Optional ByVal Count As Long) As String
    
    Middle = VBA.Mid$(OperationOnText, Start, Count)
    
End Function

 
'@Description: Add extra characters to make a specific length of the OperationOnText.
'@Return: OperationOnText after adding extra characters as many as needed to fullfill the specific length(Count).
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Count           -> Specific length that we required (full length).
'       Character       -> Optional Param( repeating character using which we will increase the full length).
'@ExampleCall : Text.PadEnd("abedin@gmail.com", 5, " ")  >> abedin@gmail.com
Public Function PadEnd(ByVal OperationOnText As String _
                       , ByVal Count As Long _
                        , Optional ByVal Character As String = " ") As String
    
    Dim Result As String
    
    If Len(OperationOnText) >= Count Then
        Result = OperationOnText
    Else
        Result = OperationOnText & Repeat(Character, Count - Len(OperationOnText))
    End If
    
    PadEnd = Result
    
End Function

'@Description: Add extra characters at the beginning to make a specific length of the OperationOnText.
'@Return: OperationOnText after adding extra characters as many as needed in the beginning to fullfill the specific length(Count).
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Count           -> Specific length that we need (full length).
'       Character       -> Optional Param( repeating character using which we will increase the full length).

'@ExampleCall : Text.PadStart("0056", 6, 0)  >> 000056
Public Function PadStart(ByVal OperationOnText As String _
                         , ByVal Count As Long _
                          , Optional ByVal Character As String = " ") As String
    
    Dim Result As String
    
    If Len(OperationOnText) >= Count Then
        Result = OperationOnText
    Else
        Result = Repeat(Character, Count - Len(OperationOnText)) & OperationOnText
    End If
    
    PadStart = Result
    
End Function

'@Description: This will detect the index position of searched text(SubString).
'@Return: will return the index position of searched text as long
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       SubString       -> A string param containg values that we want to search in OperationOnText.
'       GivenOccurence  -> Optional param (Use to specify which findings will be considered first or last or all the findings).
'       ComparisionType -> Optional param which will specify whether it will be case sensitive or not.
'@ExampleCall : Text.PositionOf("1997ismail@hosen#gmail.com", "@")  >>
Public Function PositionOf(ByVal OperationOnText As String _
                           , ByVal SubString As String _
                            , Optional ByVal GivenOccurence As Occurrence = FIRST _
                             , Optional ByVal ComparisionType As Comparer = IGNORE_CASE) As Variant
    
    Dim ComparisionOpt As VbCompareMethod
    ComparisionOpt = GetComparisionOption(ComparisionType)
    
    Dim Result As Variant
    
    Select Case GivenOccurence
    
        Case Occurrence.FIRST
            Result = InStr(FROM_START, OperationOnText, SubString, ComparisionOpt)
            
        Case Occurrence.LAST
            Result = InStrRev(OperationOnText, SubString, FROM_end, ComparisionOpt)
            
        Case Occurrence.ALL
            Set Result = FindAllIndexOf(OperationOnText, SubString, FROM_START, ComparisionType)
            
        Case Else
            Err.Raise 13, "Text.PositionOf", "Wrong Occurence type"
        
    End Select
    
    If IsObject(Result) Then
        Set PositionOf = Result
    Else
        PositionOf = Result
    End If
    
End Function

'@Description: This will detect the index position of the first finding character(delimiter) among multiple delimiters that are searching for.
'@Return: will return the index postion of the delimiter that is matched or found first in the OperationOntext.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Delimiters      -> A variant param which will take multiple delimiters to search in OperationOnText.
'@ExampleCall : Text.PositionOfAny("1997ismail@hosen#gmail.com", Array("#","@","n")) >> 11

Public Function PositionOfAny(ByVal OperationOnText As String _
                              , ByVal Delimiters As Variant) As Long
       
    'Initially consider that index is more than text length
    Dim FirstFoundIndex As Long
    FirstFoundIndex = Len(OperationOnText) + 1
    
    Dim IsAnyDelimiterFound As Boolean
    Dim CurrentDelimter As Variant
    
    For Each CurrentDelimter In Delimiters
        
        Dim FoundInPosition As Long
        FoundInPosition = PositionOf(OperationOnText, CStr(CurrentDelimter))
        
        If FoundInPosition < FirstFoundIndex And FoundInPosition <> 0 Then
            IsAnyDelimiterFound = True
            FirstFoundIndex = FoundInPosition
        End If
        
    Next CurrentDelimter
    
    PositionOfAny = IIf(IsAnyDelimiterFound, FirstFoundIndex, -1)
    
End Function

'@Description: This will detect either any of the Delimiters are exist in OperationOntext or not.
'@Return: Boolean value of the search result. If it found any of the delimiters then it return true. Else it will be false.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Delimiters       -> A variant param which will take multiple delimiters to search in OperationOnText.
'@ExampleCall : Text.IsAnyDelimiterExists("1997ismail@hosen#gmail.com",Array("$","*"))  >> False
Public Function IsAnyDelimiterExists(ByVal OperationOnText As String _
                                     , ByVal Delimiters As Variant) As Boolean
       
    IsAnyDelimiterExists = (PositionOfAny(OperationOnText, Delimiters) <> -1)
    
End Function

'@Description: This will detect either all the delimiters are exist in OperationOntext or not.
'@Return: Boolean value of the search result. If it found all the delimiters then it return true. Else it will be false.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Delimiters      -> A variant param which will take multiple delimiters to search in OperationOnText.
'@ExampleCall : Text.IsAllDelimiterExists("1997ismail@hosen#gmail.com",Array("#","@"))  >> True
'               Text.IsAllDelimiterExists("1997ismail@hosen#gmail.com",Array("$","@"))  >> False
Public Function IsAllDelimiterExists(ByVal OperationOnText As String _
                                     , ByVal Delimiters As Variant) As Boolean
       
    Dim FirstFoundIndex As Long
    FirstFoundIndex = Len(OperationOnText) + 1
    
    Dim Result As Boolean
    Result = True
    
    Dim CurrentDelimiter As Variant
    
    For Each CurrentDelimiter In Delimiters
        
        Dim FoundInPosition As Long
        FoundInPosition = PositionOf(OperationOnText, CStr(CurrentDelimiter))
        
        If FoundInPosition = 0 Then
            Result = False
            Exit For
        End If
        
    Next CurrentDelimiter
    
    IsAllDelimiterExists = Result
    
End Function


'@Description: Will Return the Propercase of OperarationOnText
'@Return: Propercase of OperarationOnText
'@Param:
'       OperationOnText -> A string param that we want to convert into Propercase.
'@ExampleCall : Text.Proper("Dhaka university")  >> Dhaka University
Public Function Proper(ByVal OperationOnText As String) As String

    Proper = Application.WorksheetFunction.Proper(OperationOnText)
    
End Function

'@Description: Returns a number of characters from a text value starting at a zero-based offset and for count number of characters.
'@Return:
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Offset          -> Offset is the starting point of index postion after which we want to extract. It starts from 0.
'       Count           -> Optional param( use to indicate the number of characters we need to extract)
'@ExampleCall : Text.Range("1997ismail@hosen#gmail.com", "@")  >>
Public Function Range(ByVal OperationOnText As String _
                      , ByVal Offset As Long _
                       , Optional ByVal Count As Long) As String

    Range = Middle(OperationOnText, Offset + 1, Count)
    
End Function

'@Description: This will remove all the characters(RemoveChars) from OperationOnText.
'@Return: Will return the value of OperationOnText without characters to remove.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       RemoveChars     -> An array of characters those are needed to remove from OperationOnText.
'@ExampleCall : Text.RemoveCharacters("1997@ismail@hosen#gmail.com",Array("@","#"))  >> 1997ismailhosengmail.com
Public Function RemoveCharacters(ByVal OperationOnText As String _
                                 , ByVal RemoveChars As Variant) As String
     
    Dim ResultText As String
    ResultText = OperationOnText
    
    Dim CurrentChar As Variant
    
    For Each CurrentChar In RemoveChars
        ResultText = VBA.Replace(ResultText, CurrentChar, vbNullString)
    Next CurrentChar
    
    RemoveCharacters = ResultText
    
End Function

'@Description: This will remove number of characters from the beginning of OperationOnText.
'@Return: Will return the value of OperationOnText without characters to remove.
'@Param:
'       OperationOnText     -> A string param on which we want to conduct our operation.
'       NumberOfCharacter   -> Total number of characters we want to remove from the beginning.
'@ExampleCall : Text.RemoveFromStart("1997@ismail@hosen#gmail.com", 4)  >> @ismail@hosen#gmail.com
Public Function RemoveFromStart(ByVal OperationOnText As String _
                                , ByVal NumberOfCharacter As Long) As String
    
    Dim Result As String
    
    If OperationOnText = vbNullString Then
        Result = vbNullString
    ElseIf NumberOfCharacter < 1 Then
        Result = OperationOnText
    ElseIf NumberOfCharacter >= Len(OperationOnText) Then
        Result = vbNullString
    Else
        Result = SubString(OperationOnText, NumberOfCharacter + 1)
    End If
    
    RemoveFromStart = Result
    
End Function

'@Description: This will remove characters from the beginning if it matches the searched string(TextToFInd) at the beginning.
'@Return: Will return the value of OperationOnText after removing matched text in the beginning. If it does't match it will return the original text.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       TextToFind      -> Text that we are searching for.
'       ComparisionType -> Optional param which will specify whether it will be case sensitive or not.
'       @ExampleCall : Text.RemoveFromStartIfPresent("1997ismail@hosen#gmail.com","1997")  >> ismail@hosen#gmail.com
Public Function RemoveFromStartIfPresent(ByVal OperationOnText As String _
                                         , ByVal TextToFind As String _
                                          , Optional ByVal ComparisionType As Comparer = IGNORE_CASE) As String
    
    Dim Result As String
    
    If IsStartsWith(OperationOnText, TextToFind, ComparisionType) Then
        Result = RemoveFromStart(OperationOnText, Len(TextToFind))
    Else
        Result = OperationOnText
    End If
    
    RemoveFromStartIfPresent = Result
    
End Function

'@Description: This will remove number of characters from the end of OperationOnText.
'@Return: Will return the value of OperationOnText without characters to remove.
'@Param:
'       OperationOnText     -> A string param on which we want to conduct our operation.
'       NumberOfCharacter   -> Total number of characters we want to remove from the end.
'@ExampleCall : Text.RemoveFromEnd("1997@ismail@hosen#gmail.com", 4)  >> 1997@ismail@hosen#gmail
Public Function RemoveFromEnd(ByVal OperationOnText As String _
                              , ByVal NumberOfCharacter As Long) As String
    
    Dim Result As String
    
    If OperationOnText = vbNullString Then
        Result = vbNullString
    ElseIf NumberOfCharacter < 1 Then
        Result = OperationOnText
    ElseIf NumberOfCharacter >= Len(OperationOnText) Then
        Result = vbNullString
    Else
        Result = Mid$(OperationOnText, 1, Len(OperationOnText) - NumberOfCharacter)
    End If
    
    RemoveFromEnd = Result
    
End Function

'@Description: This will remove characters from the end if it matches the searched string(TextToFInd) at the end.
'@Return: Will return the value of OperationOnText after removing matched text in the end. If it does't match it will return the original text.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       TextToFind      -> Text that we are searching for.
'       ComparisionType -> Optional param which will specify whether it will be case sensitive or not.
'       @ExampleCall : Text.RemoveFromEndIfPresent("1997ismail@hosen#gmail.com","1997")  >> ismail@hosen#gmail.com
'       @ExampleCall : Text.RemoveFromEndIfPresent("1997ismail@hosen#gmail.com","1998")  >> 1997ismail@hosen#gmail.com
Public Function RemoveFromEndIfPresent(ByVal OperationOnText As String _
                                       , ByVal TextToFind As String _
                                        , Optional ByVal ComparisionType As Comparer = IGNORE_CASE) As String
    
    Dim Result As String
    
    If IsEndsWith(OperationOnText, TextToFind, ComparisionType) Then
        Result = RemoveFromEnd(OperationOnText, Len(TextToFind))
    Else
        Result = OperationOnText
    End If
    
    RemoveFromEndIfPresent = Result
    
End Function

'@Description: Remove specific range of characters from a text value starting at 1 based offset and for count number of characters.
'@Return:
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Offset          -> Offset is the starting point of index postion after which we want to remove. It starts from 1.
'       Count           -> Optional param( use to indicate the number of characters we need to remove).
'@ExampleCall : Text.RemoveRange("1997ismail@hosen#gmail.com",1,4)  >> ismail@hosen#gmail.com
Public Function RemoveRange(ByVal OperationOnText As String _
                            , ByVal Offset As Long _
                             , Optional ByVal Count As Long = -1) As String
    
    Dim Result As String
    
    If OperationOnText = vbNullString Then
        Result = vbNullString
    ElseIf Count = -1 Then
        Result = Left$(OperationOnText, Offset - 1)
    Else
        Result = Left$(OperationOnText, Offset - 1) & Mid$(OperationOnText, Offset + Count)
    End If
    
    RemoveRange = Result
    
End Function

'correctionNeeded
'@Description:
'@Return:
'@Param:
'       OperationOnText  -> A string param on which we want to conduct our operation.
'       PadText          ->
'       PadDirection     -> Optional param(
'       ComparisionType  -> Optional param(
'@ExampleCall : Text.Range("1997ismail@hosen#gmail.com", "@")  >>
Public Function PadIfNotPresent(ByVal OperationOnText As String _
                                , ByVal PadText As String _
                                 , Optional ByVal PadDirection As RelativePosition = RelativePosition.FROM_START _
                                  , Optional ByVal ComparisionType As Comparer = Comparer.IGNORE_CASE) As String
    
    Dim Result As String
    Result = OperationOnText
    Select Case PadDirection
        Case RelativePosition.FROM_BOTH
            OperationOnText = PadIfNotPresent(OperationOnText, PadText, FROM_START)
            Result = PadIfNotPresent(OperationOnText, PadText, FROM_end)
        Case RelativePosition.FROM_START
            If Not IsStartsWith(OperationOnText, PadText, ComparisionType) Then
                Result = PadText & OperationOnText
            End If
        Case RelativePosition.FROM_end
            If Not IsEndsWith(OperationOnText, PadText, ComparisionType) Then
                Result = OperationOnText & PadText
            End If
        Case Else
            Err.Raise 13, "Text.PadIfNotPresent", "PadDirection value is wrong."
    End Select
    
    PadIfNotPresent = Result
    
End Function

'@Description: This will repeat the OperationOnText for a specific number of time(Count)
'@Return:
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Count           -> A long for specifying how many times we want to repeat the text
'@ExampleCall : Text.Repeat("ismail@hosen",3)  >> ismail@hosenismail@hosenismail@hosen
Public Function Repeat(ByVal OperationOnText As String _
                       , ByVal Count As Long) As String
    
    Dim ResultText As String
    ResultText = OperationOnText
    
    Dim CurrentCount As Long
    For CurrentCount = 1 To Count - 1
        ResultText = ResultText & OperationOnText
    Next CurrentCount
    
    Repeat = ResultText
    
End Function

'@Description: This will replace the specific length of character/s with a specific string.
'@Return: New modified string after removing specific characters and replaced them with some characters.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Offset          -> Offset is the starting point of index postion from which we want to replace. It starts from 1.
'       Count           -> A Long param( use to indicate the number of characters we want to replace).
'       ReplaceWith     -> String or values by which we want to replace the specified text.
'@ExampleCall : Text.ReplaceRange("1997 Downtown Street",15,6,"City")  >> 1997 Downtown City
Public Function ReplaceRange(ByVal OperationOnText As String _
                             , ByVal Offset As Long _
                              , ByVal Count As Long _
                               , ByVal ReplaceWith As String) As String
    
    Dim Result As String
    
    If OperationOnText = vbNullString Then
        Result = vbNullString
    Else
        Result = Left$(OperationOnText, Offset - 1) _
                 & ReplaceWith _
                 & Mid$(OperationOnText, Offset + Count)
    End If
    
    ReplaceRange = Result
    
End Function

'@Description: Reverses the provided text.
'@Return: Return reverse text
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@ExampleCall : Text.Reverse("ismail@hosen#gmail.com")  >> moc.liamg#nesoh@liamsi
Public Function Reverse(ByVal OperationOnText As String) As String
    Reverse = StrReverse(OperationOnText)
End Function

'@Description: This function will extract only provided characters(SelectChars) from OperationOnText
'@Return: Return extracted all the characters that are matched in OperationOnText
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       SelectChars     -> All the characters that we want to extract from OperationOnText as a string.
'@ExampleCall : Text.Keep("1997@ismail@hosen#gmail.com", "1234567890")  >> 1997
'               Text.Keep("1997@ismail@hosen#gmail.com", "@#") >> @@#
Public Function Keep(ByVal OperationOnText As String _
                     , ByVal SelectChars As String) As String
    
    Dim Result As String
    
    If OperationOnText = vbNullString Or SelectChars = vbNullString Then
        Result = OperationOnText
    Else
        
        Dim CurrentenIndex As Long
        Dim CurrentCharacter As String
        
        For CurrentenIndex = 1 To Len(OperationOnText)
            CurrentCharacter = Mid$(OperationOnText, CurrentenIndex, 1)
            If Contains(SelectChars, CurrentCharacter, CONSIDER_CASE) Then
                Result = Result & CurrentCharacter
            End If
        Next CurrentenIndex
        
    End If
    
    Keep = Result
    
End Function

'@Description: This will returns a list containing parts of a text value that are delimited by a separator text value.
'@Return: Will return a variant type data that contains the list of sperated part by seperator in OperationOntext.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Separator       -> A specific seperator( Function will split OperationOnText into another part when it found that seperator)
'@ExampleCall : Text.SplitText("C:\Users\Ismail\Desktop\hasan's VBA practices\Github collaboration\","\")  >>
Public Function SplitText(ByVal OperationOnText As String _
                          , ByVal Delimiter As String _
                           , Optional ByVal IgnoreInsideQuote As Boolean = False) As Variant
    
    Dim Result As Variant
    
    If Not IgnoreInsideQuote Then
        Result = Application.WorksheetFunction.Transpose(VBA.Split(OperationOnText, Delimiter))
    Else
        Result = SplitTextIgnoreDelInsideQuote(OperationOnText, Delimiter)
    End If

    SplitText = Result
    
End Function

Private Function SplitTextIgnoreDelInsideQuote(ByVal OperationOnText As String _
                                                 , ByVal Delimiter As String) As Variant
    Dim InQuotes As Boolean
    InQuotes = False
    Dim StartIndex As Long
    StartIndex = 1
    
    Dim Result As Collection
    Set Result = New Collection
    
    Dim CurrentCharIndex As Long
    For CurrentCharIndex = 1 To Len(OperationOnText)
        
        Dim CurrentCharacter As String
        CurrentCharacter = Mid$(OperationOnText, CurrentCharIndex, 1)
        Const QUOTE As String = """"
        If CurrentCharacter = QUOTE Then
            InQuotes = Not InQuotes
        ElseIf Not InQuotes And Mid$(OperationOnText, CurrentCharIndex, Len(Delimiter)) = Delimiter Then
            Dim EndIndex As Long
            EndIndex = CurrentCharIndex - 1
            Result.Add Mid$(OperationOnText, StartIndex, EndIndex - StartIndex + 1)
            StartIndex = CurrentCharIndex + Len(Delimiter)
        End If
        
    Next CurrentCharIndex

    Result.Add Mid$(OperationOnText, StartIndex)
    
    SplitTextIgnoreDelInsideQuote = CollectionToArray(Result)
    
End Function


'@Description: This will help to extract a specific number of characters from the beginning of the OperationOnText.
'@Return: Specific number of characters from the beginning.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Count           -> A long param( use to specify how many characters we want to extract)
'@ExampleCall : Text.Start("1997ismail@hosen#gmail.com", 4)  >> 1997
Public Function Start(ByVal OperationOnText As String, ByVal Count As Long) As String
    Start = Left$(OperationOnText, Count)
End Function

'@Description: This will extract part of a string like java.If You don't mention EndIndex then it will extract upto last part.
'@Return: Will return extracted part of OperationOnText based on StartIndex and EndIndex.
'@Param:
'       GivenText   -> A string param on which we want to conduct our operation.
'       StartIndex  -> A long param( use to specify the staring index to
'       EndIndex    -> Optional param( use to specify Ending point as a long. Function will extract upto last if this remain empty)
'@ExampleCall : Text.SubString("Ismail Hosen",8) >> Hosen
Public Function SubString(ByVal GivenText As String _
                          , ByVal StartIndex As Long _
                           , Optional ByVal EndIndex As Long = -1) As String
    
    If EndIndex = -1 Then EndIndex = Len(GivenText)
    
    If StartIndex > EndIndex Then
        Err.Raise 13, "SubString", "StartIndex need to be less or equal to EndIndex"
    End If
    
    SubString = Mid$(GivenText, StartIndex, EndIndex - StartIndex + 1)
    
End Function

'@Description: Returns a logical value indicating whether a text value substring was found at the beginning of a string or not.
'@Return:
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       TextToMatch     -> Text that we want to check whether existed at the beginning.
'       ComparisionType -> Optional param which will specify whether it will be case sensitive or not.
'@ExampleCall : Text.IsStartsWith("1997ismail@hosen#gmail.com", "1997")  >> True
Public Function IsStartsWith(ByVal OperationOnText As String _
                             , ByVal TextToMatch As String _
                              , Optional ByVal ComparisionType As Comparer = IGNORE_CASE) As Boolean
    
    Dim Result As Boolean
    
    If ComparisionType = CONSIDER_CASE Then
        Result = (Left$(OperationOnText, Len(TextToMatch)) = TextToMatch)
    ElseIf ComparisionType = IGNORE_CASE Then
        Result = (UCase$(Left$(OperationOnText, Len(TextToMatch))) = UCase$(TextToMatch))
    End If
    
    IsStartsWith = Result
    
End Function

'@Description: This funciton will return all the characters as an array. Every single character will be a single element.
'@Return: An array containing all the characters
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@ExampleCall : Text.ToCharacters("1997ismail@hosen#gmail.com")  >>
Public Function ToCharacters(ByVal OperationOnText As String) As Variant
    
    Dim Result As Variant
    
    If OperationOnText = vbNullString Then
        Result = vbEmpty
    Else
        
        ReDim Result(1 To Len(OperationOnText), 1 To 1)
        Dim CurrentIndex As Long
        For CurrentIndex = 1 To Len(OperationOnText)
            Result(CurrentIndex, 1) = Mid$(OperationOnText, CurrentIndex, 1)
        Next CurrentIndex

    End If
    
    ToCharacters = Result
  
End Function


'@Description:This will remove all the extra spaces from OperationOnString( srart, end, middle)
'@Return: Will return OperationOnText after removing all the extra spaces.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@DependencyOnExcel:
'@ExampleCall : Text.Trim("     1997ismail@ho       sen#gmail.com    ")  >> 1997ismail@ho sen#gmail.com
Public Function Trim(ByVal OperationOnText As String) As String
    Trim = Application.WorksheetFunction.Trim(OperationOnText)
End Function

'@Description:This will remove all the extra spaces from the end of OperationOnString( srart, end, middle)
'@Return: Will return OperationOnText after removing all the extra spaces from the end.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@ExampleCall : Text.TrimEnd("1997ismail@hosen#gmail.com      ")  >>
Public Function TrimEnd(ByVal OperationOnText As String) As String
    TrimEnd = VBA.RTrim$(OperationOnText)
End Function

'@Description:This will remove all the extra spaces from the beginning of OperationOnString( srart, end, middle)
'@Return: Will return OperationOnText after removing all the extra spaces from the beginning.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@ExampleCall : Text.TrimStart("           1997ismail@hosen#gmail.com")  >>
Public Function TrimStart(ByVal OperationOnText As String) As String
    TrimStart = VBA.LTrim$(OperationOnText)
End Function

'@Description: Will Return the Uppercase of OperarationOnText
'@Return: Uppercase of OperarationOnText
'@Param:
'       OperationOnText -> A string param that we want to convert into Uppercase.
'@ExampleCall : Text.Upper("bangladesh")  >> BANGLADESH
Public Function Upper(ByVal OperationOnText As String) As String
    Upper = VBA.UCase$(OperationOnText)
End Function

'@Description:This will extract all the numbers from the end.
'@Return: All the numbers from the end
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@ExampleCall : Text.ExtractNumberFromEnd("ismailhosen1997")  >> 1997
Public Function ExtractNumberFromEnd(ByVal OperationOnText As String) As String
    ExtractNumberFromEnd = ExtractNumber(OperationOnText, FROM_end)
End Function

'@Description: This will extract all the numbers from the beginning.
'@Return: All the numbers from the beginning.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@ExampleCall : Text.ExtractNumberFromStart("1997ismailhosen")  >>
Public Function ExtractNumberFromStart(ByVal OperationOnText As String) As String
    ExtractNumberFromStart = ExtractNumber(OperationOnText, FROM_START)
End Function

'@Description: This funciton will extract number from input string.
'@Return:Extracted numbers only
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       SearchDirection -> Optional param which indicate from which side we need to search.
'@ExampleCall : Text.ExtractNumber("1997ismailhosen286")  >> 1997
Private Function ExtractNumber(ByVal OperationOnText As String _
                               , Optional ByVal SearchDirection As RelativePosition = FROM_START) As String
    
    Dim ExtractedNumbers As Collection
    Set ExtractedNumbers = ExtractAllNumbers(OperationOnText)
    
    Dim Result As String
    If ExtractedNumbers.Count = 0 Then
        Result = vbNullString
    ElseIf SearchDirection = FROM_START Or SearchDirection = FROM_BOTH Then
        Result = ExtractedNumbers.Item(1)
    Else
        Result = ExtractedNumbers.Item(ExtractedNumbers.Count)
    End If
    
    ExtractNumber = Result
    
End Function

'@Description: This funciton will extract all number from input string.
'@Return:A Collection containing splitted part of combined numbers of Input string.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@ExampleCall : Text.ExtractAllNumbers("1997hasan47imam1998")  >> ExtractAllNumbers(Item 1 ="1997",Item 2 ="47", Item 3 = "1998")
Public Function ExtractAllNumbers(ByVal OperationOnText As String) As Collection
    
    Dim SplittedDigitAndAlphabet As Collection
    Set SplittedDigitAndAlphabet = SplitDigitAndNonDigit(OperationOnText)
    
    Dim Numbers As Collection
    Set Numbers = New Collection

    Dim Index As Long
    For Index = 1 To SplittedDigitAndAlphabet.Count
        
        Dim FirstCharacter As String
        FirstCharacter = Mid$(SplittedDigitAndAlphabet(Index), 1, 1)
        If IsDigit(FirstCharacter) Then
            Numbers.Add SplittedDigitAndAlphabet(Index)
        End If
        
    Next Index
    
    Set ExtractAllNumbers = Numbers
    
End Function

'@Description: This funciton will extract all texts from input string.
'@Return:A Collection containing splitted part of combined alphabets of Input string.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@ExampleCall : Text.ExtractAllTexts("1997hasan47imam1998")  >> ExtractAllTexts(Item 1 ="hasan",Item 2 ="imam")
Public Function ExtractAllTexts(ByVal OperationOnText As String) As Variant
    
    Dim SplittedDigitAndAlphabet As Collection
    Set SplittedDigitAndAlphabet = SplitDigitAndNonDigit(OperationOnText)
    
    Dim Texts As Collection
    Set Texts = New Collection

    Dim Index As Long
    For Index = 1 To SplittedDigitAndAlphabet.Count
        
        Dim FirstCharacter As String
        FirstCharacter = Mid$(SplittedDigitAndAlphabet(Index), 1, 1)
        If Not IsDigit(FirstCharacter) Then
            Texts.Add SplittedDigitAndAlphabet(Index)
        End If
        
    Next Index
    
    Set ExtractAllTexts = Texts
    
End Function

'@Description: This function will increase / decrease end number by Amount.
'@Return: Full Input text with Increased or decreased end number of OperationOntext by amount.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Amount          -> A long param( indicate how much we want to increase or decrease
'       IsNoOperationIfNoNumberPresent -> Optional param (use to specify exit function if no number is found)
'@ExampleCall : Text.IncrementOrDecrementEndNumber("1997ismailhosen286",4)  >> 1997ismailhosen290
'               Text.IncrementOrDecrementEndNumber("1997ismailhosen286", -4) >> 1997ismailhosen282
Public Function IncrementOrDecrementEndNumber(ByVal OperationOnText As String _
                                              , ByVal Amount As Long _
                                               , Optional ByVal IsNoOperationIfNoNumberPresent As Boolean = True) As String

    Dim NumberAtEnd As String
    NumberAtEnd = ExtractNumberFromEnd(OperationOnText)
    
    Dim Result As String
    
    If NumberAtEnd = vbNullString Then
        
        If IsNoOperationIfNoNumberPresent Then
            Result = OperationOnText
        ElseIf Amount = 1 Then
            Result = OperationOnText & (Amount + 1)
        Else
            Result = OperationOnText & Amount
        End If
        
    Else
        Result = RemoveFromEndIfPresent(OperationOnText, NumberAtEnd, IGNORE_CASE) & _
                 (CLng(NumberAtEnd) + Amount)
    End If
    
    IncrementOrDecrementEndNumber = Result
    
End Function

'@Description: This function will increase / decrease starting number by Amount.
'@Return: Full Input text with Increased or decreased starting number of OperationOntext by amount.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       Amount          -> A long param( indicate how much we want to increase or decrease
'       IsNoOperationIfNoNumberPresent -> Optional param (use to specify exit function if no number is found)
'@ExampleCall : Text.IncrementOrDecrementStartNumber("1997ismailhosen286",3)  >> 2000ismailhosen286
'               Text.IncrementOrDecrementStartNumber("1997ismailhosen286", -4) >> 1993ismailhosen286
Public Function IncrementOrDecrementStartNumber(ByVal OperationOnText As String _
                                                , ByVal Amount As Long _
                                                 , Optional IsNoOperationIfNoNumberPresent As Boolean = True) As String
    
    Dim NumberAtStart As String
    NumberAtStart = ExtractNumberFromStart(OperationOnText)
    
    Dim Result As String
    
    If NumberAtStart = vbNullString Then
        
        If IsNoOperationIfNoNumberPresent Then
            Result = OperationOnText
        Else
            Result = OperationOnText & Amount
        End If

    Else
        Result = (CLng(NumberAtStart) + Amount) & _
                 RemoveFromStartIfPresent(OperationOnText, NumberAtStart, IGNORE_CASE)
    End If
    
    IncrementOrDecrementStartNumber = Result
    
End Function

'@Description: This function just reverse the collection. It also doesn't add key to the output  collection
'@Return: Collection with reversed index.
'@Param:
'       InputCollection ->
'@ExampleCall : Text.ReverseIndexCollection()  >>
Private Function ReverseIndexCollection(ByVal InputCollection As Collection) As Collection
    
    Dim ReversedIndex As Collection
    Set ReversedIndex = New Collection
    
    Dim CurrentItemIndex As Long
    For CurrentItemIndex = InputCollection.Count To 1 Step -1
        ReversedIndex.Add InputCollection.Item(CurrentItemIndex)
    Next CurrentItemIndex
    
    Set ReverseIndexCollection = ReversedIndex
        
End Function

'@Description: This function will concatenate all the array with the delimiter.
'@Return: A string after combining all the elements of GivenArray with Delimiter.
'@Param:
'       GivenArray -> Array contains values that we want to combine.
'       Delimiter  -> String with which we want to combine each of the element of the array.
'@ExampleCall : Text.ReverseIndexCollection()  >>
Private Function ConcatenateArray(ByVal GivenArray As Variant _
                                  , ByVal Delimiter As String) As String
   
    Dim ConcatenatedText As String
    
    Dim Element As Variant
    For Each Element In GivenArray
        ConcatenatedText = ConcatenatedText & CStr(Element) & Delimiter
    Next Element
    
    ConcatenatedText = RemoveFromEndIfPresent(ConcatenatedText, Delimiter, IGNORE_CASE)
    
    ConcatenateArray = ConcatenatedText
    
End Function

'@Description: This will remove characters from both end if it matches the searched string(TextToFInd) at any side.
'@Return: Will return the value of OperationOnText after removing matched text. If it does't match it will return the original text.
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'       TextToFind      -> Text that we are searching for.
'       ComparisionType -> Optional param which will specify whether it will be case sensitive or not.
'       @ExampleCall : Text.RemoveFromBothEndIfPresent("1997ismail@hosen#gmail.com1998","1997")  >> ismail@hosen#gmail.com
'       @ExampleCall : Text.RemoveFromBothEndIfPresent("1997ismail@hosen#gmail.com","com")  >> 1997ismail@hosen#gmail.
Public Function RemoveFromBothEndIfPresent(ByVal OperationOnText As String _
                                           , ByVal TextToFind As String _
                                            , Optional ByVal ComparisionType As Comparer = IGNORE_CASE) As String

    OperationOnText = RemoveFromEndIfPresent(OperationOnText, TextToFind, ComparisionType)
    RemoveFromBothEndIfPresent = RemoveFromStartIfPresent(OperationOnText, TextToFind, ComparisionType)
    
End Function

'@Description: This will seperate numbers and alphabets with seperator and provide trimmed(white space removed) OperationOnText.
'@Return:
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@ExampleCall : Text.ConvertToMeaningfulText("1997ismail@h   1252osen#gmail.com    ","\")  >> 1997\ismail@h \1252\osen#gmail.com
Public Function ConvertToMeaningfulText(ByVal OperationOnText As String _
                                        , ByVal Separator As String) As String
    
    Dim SplittedNumberAndText As Collection
    Set SplittedNumberAndText = SplitDigitAndNonDigit(OperationOnText)
    
    Dim MeaningFulName As String
    
    Dim CurrentItem As Variant
    
    For Each CurrentItem In SplittedNumberAndText
        If IsNumeric(CurrentItem) Then
            MeaningFulName = MeaningFulName & CurrentItem & Separator
        Else
            MeaningFulName = MeaningFulName _
                             & ConvertTextToUseableName(CStr(CurrentItem), Separator) _
                             & Separator
        End If
    Next CurrentItem
    
    MeaningFulName = Text.RemoveFromEnd(MeaningFulName, Len(Separator))
    MeaningFulName = Text.Trim(MeaningFulName)
    
    ConvertToMeaningfulText = MeaningFulName
    
End Function

'@Description: This function will check whether the character is capital letter or not.
'@Return: Return boolean value of the result.
'@Param:
'       GivenChar -> Character that we want to check.
'@ExampleCall : Text.IsCapitalLetter("H")  >> True
'               Text.IsCapitalLetter("h")  >> False
Public Function IsCapitalLetter(ByVal GivenChar As String) As Boolean
    
    If GivenChar = vbNullString Then
        IsCapitalLetter = False
        Exit Function
    End If
    
    Const CHARACTER_CODE_OF_CAPITAL_A As Long = 65
    Const CHARACTER_CODE_OF_CAPITAL_Z As Long = 90
    Dim CurrentCharCode As Long
    CurrentCharCode = VBA.Asc(GivenChar)
    IsCapitalLetter = (CurrentCharCode >= CHARACTER_CODE_OF_CAPITAL_A _
                       And CurrentCharCode <= CHARACTER_CODE_OF_CAPITAL_Z)
                       
End Function

'@Description: This function will check whether the character is english alphabet or not.
'@Return: Return boolean value of the result.
'@Param:
'       GivenChar -> Character that we want to check.
'@ExampleCall : Text.IsEnglishAlphabet("H")  >> True
'               Text.IsEnglishAlphabet("5")  >> False
Public Function IsEnglishAlphabet(ByVal GivenChar As String) As String
    IsEnglishAlphabet = (IsSmallLetter(GivenChar) Or IsCapitalLetter(GivenChar))
End Function

'@Description:
'@Return:
'@Param:
'       OperationOnText -> A string param on which we want to conduct our operation.
'@ExampleCall : Text.ConvertTextToUseableName("1997ismail@hosen#gmail.com")  >>
'@TODO: Need refactoring
Private Function ConvertTextToUseableName(ByVal OperationOnText As String _
                                          , Optional ByVal Separator As String = " ") As String

    Dim Result As Collection
    Set Result = New Collection
    
    Dim IsNewCapitalLetterStart As Boolean
    Dim IsNewSmallLetterStart As Boolean
    Dim CurrentCharIndex As Long
    Dim LastSmallLetterSeenAtIndex As Long
    Dim LastCapitalLetterSeenAtIndex As Long
    
    For CurrentCharIndex = 1 To Len(OperationOnText)
        If IsCapitalLetter(Mid$(OperationOnText, CurrentCharIndex, 1)) Then
            
            If IsNewSmallLetterStart Then
                IsNewCapitalLetterStart = True
                LastCapitalLetterSeenAtIndex = CurrentCharIndex
                Result.Add Mid$(OperationOnText, LastSmallLetterSeenAtIndex, CurrentCharIndex - LastSmallLetterSeenAtIndex)
                IsNewSmallLetterStart = False
            ElseIf Not IsNewCapitalLetterStart Then
                IsNewCapitalLetterStart = True
                LastCapitalLetterSeenAtIndex = CurrentCharIndex
            End If
            
        Else
            
            If IsNewCapitalLetterStart Then
                IsNewSmallLetterStart = True
                LastSmallLetterSeenAtIndex = CurrentCharIndex
                Result.Add Mid$(OperationOnText, LastCapitalLetterSeenAtIndex, CurrentCharIndex - LastCapitalLetterSeenAtIndex)
                IsNewCapitalLetterStart = False
            ElseIf Not IsNewSmallLetterStart Then
                IsNewSmallLetterStart = True
                LastSmallLetterSeenAtIndex = CurrentCharIndex
            End If
            
        End If
    Next CurrentCharIndex
    
    If IsNewSmallLetterStart Then
        Result.Add Mid$(OperationOnText, LastSmallLetterSeenAtIndex, CurrentCharIndex - LastSmallLetterSeenAtIndex)
    End If
    
    If IsNewCapitalLetterStart Then
        Result.Add Mid$(OperationOnText, LastCapitalLetterSeenAtIndex, CurrentCharIndex - LastCapitalLetterSeenAtIndex)
    End If
    
    Dim MeaningFullName As String
    Dim Counter As Long
    Do While Counter < Result.Count
        Counter = Counter + 1
        Dim CurrentWord As String
        CurrentWord = Result.Item(Counter)
        If VBA.Len(CurrentWord) > 1 And IsCapitalLetter(Right$(CurrentWord, 1)) Then
            If Counter < Result.Count Then
                MeaningFullName = MeaningFullName & VBA.Left$(CurrentWord, Len(CurrentWord) - 1) & Separator _
                                  & VBA.Right$(CurrentWord, 1) _
                                  & Result.Item(Counter + 1) & Separator
            Else
                MeaningFullName = MeaningFullName & CurrentWord
            End If
            Counter = Counter + 1
        ElseIf VBA.Len(CurrentWord) = 1 And IsCapitalLetter(CurrentWord) Then
            If Counter < Result.Count Then
                MeaningFullName = MeaningFullName & CurrentWord & Result.Item(Counter + 1) & Separator
            Else
                MeaningFullName = MeaningFullName & CurrentWord
            End If
            Counter = Counter + 1
        Else
            MeaningFullName = MeaningFullName & CurrentWord
            If Counter < Result.Count Then MeaningFullName = MeaningFullName & Separator
            Counter = Counter + 1
        End If
    Loop
    
    ConvertTextToUseableName = MeaningFullName
    
End Function

'@Description: By default it will not through error for duplicate case
'@Return:Array after coverting from Collection.
'@Param:
'       GivenCollection -> Collection that we want to convert into an array
'@ExampleCall : Text.CollectionToArray() >>
Private Function CollectionToArray(ByVal GivenCollection As Collection) As Variant

    If GivenCollection.Count = 0 Then
        Exit Function
    End If
    
    Dim Result() As Variant
    ReDim Result(1 To GivenCollection.Count, 1 To 1)
    
    Dim CurrentElement As Variant
    Dim CurrentIndex As Long
    
    For Each CurrentElement In GivenCollection
        CurrentIndex = CurrentIndex + 1
        If IsObject(CurrentElement) Then
            Set Result(CurrentIndex, 1) = CurrentElement
        Else
            Result(CurrentIndex, 1) = CurrentElement
        End If
    Next CurrentElement
    
    CollectionToArray = Result

End Function

'@Description: This function will check whether the character is Number or not.
'@Return: Return boolean value of the result.
'@Param:
'       GivenChar -> Character that we want to check.
'@ExampleCall : Text.IsDigit("H")  >> False
'               Text.IsDigit("5")  >> True
Public Function IsDigit(ByVal GivenChar As String) As Boolean
    
    If GivenChar = vbNullString Then
        IsDigit = False
        Exit Function
    End If
    
    Const CHARACTER_CODE_OF_0 As Long = 48
    Const CHARACTER_CODE_OF_9 As Long = 57
    Dim CurrentCharCode As Long
    CurrentCharCode = VBA.Asc(GivenChar)
    
    IsDigit = (CurrentCharCode >= CHARACTER_CODE_OF_0 _
               And CurrentCharCode <= CHARACTER_CODE_OF_9)
               
End Function

'@Description: This function will check whether the character is small letter or not.
'@Return: Return boolean value of the result.
'@Param:
'       GivenChar -> Character that we want to check.
'@ExampleCall : Text.IsSmallLetter("H")  >> False
'               Text.IsSmallLetter("h")  >> True
Public Function IsSmallLetter(ByVal GivenChar As String) As Boolean
    
    If GivenChar = vbNullString Then
        IsSmallLetter = False
        Exit Function
    End If
    
    Const CHARACTER_CODE_OF_a As Long = 97
    Const CHARACTER_CODE_OF_z As Long = 122
    
    Dim CurrentCharCode As Long
    CurrentCharCode = VBA.Asc(GivenChar)
    
    IsSmallLetter = (CurrentCharCode >= CHARACTER_CODE_OF_a _
                     And CurrentCharCode <= CHARACTER_CODE_OF_z)
                       
End Function

'@Description: This function will split digit and non digit numbers and then return as a Collection.
'@Return: Collection.
'@Param:
'       GivenText -> Text that we want to split.
'@ExampleCall : Text.SplitDigitAndNonDigit("1997ismail@hosen#gmail.com")
'               Text.SplitDigitAndNonDigit("1997ismail4555lhosengmail.com") will return 4(1997, ismail,4555,hosengmail.com)
'@TODO: Need refactoring.
Public Function SplitDigitAndNonDigit(ByVal GivenText As String) As Collection

    Dim Result As Collection
    Set Result = New Collection
    Dim IsNewDigitStart As Boolean
    Dim IsNewNonDigitStart As Boolean
    Dim CurrentCharIndex As Long
    Dim LastNonDigitSeenAtIndex As Long
    Dim LastDigitSeenAtIndex As Long
    For CurrentCharIndex = 1 To Len(GivenText)
        If IsDigit(Mid$(GivenText, CurrentCharIndex, 1)) Then
            If IsNewNonDigitStart Then
                IsNewDigitStart = True
                LastDigitSeenAtIndex = CurrentCharIndex
                Result.Add Mid$(GivenText, LastNonDigitSeenAtIndex, CurrentCharIndex - LastNonDigitSeenAtIndex)
                IsNewNonDigitStart = False
            ElseIf Not IsNewDigitStart Then
                IsNewDigitStart = True
                LastDigitSeenAtIndex = CurrentCharIndex
            End If
        Else
            If IsNewDigitStart And Mid$(GivenText, CurrentCharIndex, 1) = "." Then
'                Debug.Print "Period Present"
            ElseIf IsNewDigitStart Then
                IsNewNonDigitStart = True
                LastNonDigitSeenAtIndex = CurrentCharIndex
                Result.Add Mid$(GivenText, LastDigitSeenAtIndex, CurrentCharIndex - LastDigitSeenAtIndex)
                IsNewDigitStart = False
            ElseIf Not IsNewNonDigitStart Then
                IsNewNonDigitStart = True
                LastNonDigitSeenAtIndex = CurrentCharIndex
            End If
        End If
    Next CurrentCharIndex

    If IsNewNonDigitStart Then
        Result.Add Mid$(GivenText, LastNonDigitSeenAtIndex, CurrentCharIndex - LastNonDigitSeenAtIndex)
    End If
    If IsNewDigitStart Then
        Result.Add Mid$(GivenText, LastDigitSeenAtIndex, CurrentCharIndex - LastDigitSeenAtIndex)
    End If

    Set SplitDigitAndNonDigit = Result

End Function


'@Description : By default it will not through error for duplicate case
'@FullyQualifiedCase: It will use given column indexes
'@ExampleCall : Set ArrayToCollectionMapping = ArrayToCollection(SUTArray, 1, 2, IsSuppressDuplicateError:=False) >> It will through error if any duplicate key is present.
'@ExampleCall : Set ArrayToCollectionMapping = ArrayToCollection(SUTArray, 1, 2, IsSuppressDuplicateError:=True) >> It will skip duplicate item.

'@DefaultValueCase : first column is the key and second column is the item
'@ExampleCall : Set ArrayToCollectionMapping = ArrayToCollection(SUTArray, IsSuppressDuplicateError:=False) >> It will through error if any duplicate key is present.
'@ExampleCall : Set ArrayToCollectionMapping = ArrayToCollection(SUTArray, IsSuppressDuplicateError:=True) >> It will skip duplicate item.


'@Return:Collection after coverting from Array.
'@Param:
'       GivenCollection -> Collection that we want to convert into an array
'       KeyColumnIndex  -> Optional param Index of the key column
'       ItemColumnIndex -> Optional param Index of the item column
'       IsSuppressDuplicateError -> Optional param for duplicate error.

'@ExampleCall : Text.CollectionToArray() >>
Private Function ArrayToCollection(ByVal GivenArray As Variant _
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

End Function

'@Description: This is use to check whether the key or value exists in given Collection or not.
'@Return:
'@Param:
'       GivenCollection -> Collection in which we want to search or value.
'       Key             -> String param( use to pass the value that we are searching for)
'@ExampleCall : Text.IsExistInCollection()  >>
Private Function IsExistInCollection(ByVal GivenCollection As Collection _
                                     , ByVal Key As String) As Boolean

    On Error GoTo NotExist
    Dim Item  As Variant
    If IsObject(GivenCollection.Item(1)) Then
        Set Item = GivenCollection.Item(Key)
    Else
        Item = GivenCollection.Item(Key)
    End If
    IsExistInCollection = True
    Exit Function
NotExist:
    IsExistInCollection = False

End Function

'@Description: This function will use to remove initial Spaces and New lines
'@Return:
'@Param:
'       OperationOnText -> Collection in which we want to search or value.
'@ExampleCall : Text.RemoveInitialSpaceAndNewLines("jdkfjdf  kdsfjdskfds    ")  >>
Public Function RemoveInitialSpaceAndNewLines(ByVal OperationOnText As String) As String

    Dim Index As Long
    Dim CurrentChar As String
    For Index = 1 To VBA.Len(OperationOnText)
        CurrentChar = VBA.Mid$(OperationOnText, Index, 1)
        If Not (CurrentChar = Space(1) _
                Or CurrentChar = vbNewLine _
                Or CurrentChar = VBA.Chr$(10)) Then
            Exit For
        End If
    Next Index
    
    RemoveInitialSpaceAndNewLines = VBA.Mid$(OperationOnText, Index)

End Function

'@Description: This function will retrive the portion of OperationOnText between two given index postion.
'@Return: It will return text between two  Parameter based on StartIndex and EndIndex param.
'@Param:
'       OperationOnText -> A string param containing the string in which we want to search.
'       StartIndex      -> param(indicate the starting index position).
'       EndIndex        -> param(indicate the End index position of operaitonOnText).

Public Function FromRange(ByVal OperationOnText As String _
                          , ByVal StartIndex As Long _
                           , ByVal EndIndex As Long) As String
    
    If StartIndex <= 0 Or EndIndex <= 0 _
       Or StartIndex > EndIndex Or StartIndex > Len(OperationOnText) _
       Or EndIndex > Len(OperationOnText) Then
        Err.Raise 13, , "Text.FromRange", "Invalid input argument"
    End If
    
    FromRange = Mid$(OperationOnText, StartIndex, EndIndex - StartIndex + 1)
    
End Function

'@Description("This will check if all the character of SearchForText is found in sequence in SearchInText.")
'@Dependency("No Dependency")
'@ExampleCall : IsCharacterOfSearchTextFoundInSequence("Dim CurrentCharacter As String", "DCuASt") > True
'@Date : 19 May 2023 07:57:46 PM
'@PossibleError :
Public Function IsCharacterOfSearchTextFoundInSequence(ByVal SearchInText As String _
                                                       , ByVal SearchForText As String) As Boolean

    Dim Index As Long
    Dim CurrentCharacterOfSearchForText As String
    Dim PreviousCharIndex As Long
    PreviousCharIndex = 0
    Dim CurrentCharIndex As Long
    
    Dim Result As Boolean
    Result = True
    
    For Index = 1 To Len(SearchForText)
        
        CurrentCharacterOfSearchForText = Mid(SearchForText, Index, 1)
        CurrentCharIndex = InStr(PreviousCharIndex + 1, SearchInText _
                                                       , CurrentCharacterOfSearchForText _
                                                        , vbTextCompare)
        If CurrentCharIndex = 0 Then
            Result = False
            Exit For
        End If
        PreviousCharIndex = CurrentCharIndex
        
    Next Index
    
    IsCharacterOfSearchTextFoundInSequence = Result
    
End Function

'@Description("This will try to match in each word.See example call. This function is not 100% perfect")
'@Dependency("No Dependency")
'@ExampleCall : IsCharacterOfSearchTextFoundInSequenceAndAtTheStartOfEachWord("Dim CurrentCharacter As String", "DCuAS")->True
'               because D is found on "Dim" and Cu is found on "CurrentCharacter" and "As" is found on "As" and finally S is found on "String

'@Date : 19 May 2023 07:54:27 PM
'@PossibleError :

'@TODO: Need refactoring.
Public Function IsCharacterOfSearchTextFoundInSequenceAndAtTheStartOfEachWord(ByVal SearchInText As String _
                                                                              , ByVal SearchForText As String) As Boolean

    SearchForText = Replace(SearchForText, Space(1), vbNullString)
    Dim Words As Variant
    Words = Split(SearchInText)
    Dim CurrentWord As Variant
    Dim Index As Long
    Index = 1
    
    For Each CurrentWord In Words
        
        Dim IndexInWord As Long
        IndexInWord = 1
        Do While IndexInWord <= Len(CurrentWord)
            Dim CurrentCharacterOfSearchForText As String
            CurrentCharacterOfSearchForText = LCase(Mid(SearchForText, Index, 1))
            Dim CurrentCharacterOfCurrentWord As String
            CurrentCharacterOfCurrentWord = LCase(Mid(CStr(CurrentWord), IndexInWord, 1))
            If CurrentCharacterOfCurrentWord = CurrentCharacterOfSearchForText Then
                Index = Index + 1
                IndexInWord = IndexInWord + 1
            Else
                Exit Do
            End If
        Loop
        
    Next CurrentWord
    
    'After Last Match we increment it by 1 . So subtract that.
    Index = Index - 1
    IsCharacterOfSearchTextFoundInSequenceAndAtTheStartOfEachWord = (Index = Len(SearchForText))
    
End Function

'@Description("This will check if all character is present no matter in which order.")
'@Dependency("No Dependency")
'@ExampleCall :IsAllCharacterPresent("present","rp")>True because we found r and p in "present"
'@Date : 19 May 2023 08:22:45 PM
'@PossibleError :
Public Function IsAllCharacterPresent(ByVal SearchInText As String _
                                      , ByVal SearchForText As String) As Boolean

    Dim Index As Long
    Dim CurrentCharacterOfSearchForText As String
    Dim CurrentCharIndex As Long
    
    Dim Result As Boolean
    Result = True
    
    For Index = 1 To Len(SearchForText)
        CurrentCharacterOfSearchForText = Mid(SearchForText, Index, 1)
        CurrentCharIndex = InStr(1, SearchInText, CurrentCharacterOfSearchForText, vbTextCompare)
        If CurrentCharIndex = 0 Then
            Result = False
            Exit For
        End If
    Next Index
    
    IsAllCharacterPresent = Result

End Function

Public Function IsStartsWithAndEndsWith(ByVal OperationText As String _
                                        , ByVal StartsWithWhat As String _
                                         , Optional ByVal EndsWithWhat As String = vbNullString) As Boolean
    
    ' This function will check whether the OperationText starts with StartsWithWhat and ends with EndsWithWhat
    ' If EndsWithWhat is not provided then it will check whether the OperationText starts with StartsWithWhat
    ' and ends with StartsWithWhat
    '@ExampleCall: Text.IsStartsWithAndEndsWith("1997ismail@hosen#gmail.com", "1997", "com") >> True
    '@Return: Return boolean value of the result.

    Dim StartsWith As Boolean
    Dim EndsWith As Boolean
    StartsWith = IsStartsWith(OperationText, StartsWithWhat)
    
    If EndsWithWhat = vbNullString Then
        EndsWith = IsEndsWith(OperationText, StartsWithWhat)
    Else
        EndsWith = IsEndsWith(OperationText, EndsWithWhat)
    End If

    IsStartsWithAndEndsWith = (StartsWith And EndsWith)

End Function

Public Function IsStartsWithOrEndsWith(ByVal OperationText As String _
                                        , ByVal StartsWithWhat As String _
                                         , Optional ByVal EndsWithWhat As String = vbNullString) As Boolean
    
    ' This function will check whether the OperationText starts with StartsWithWhat or ends with EndsWithWhat
    ' If EndsWithWhat is not provided then it will check whether the OperationText starts with StartsWithWhat
    ' or ends with StartsWithWhat
    '@ExampleCall: Text.IsStartsWithOrEndsWith("1997ismail@hosen#gmail.com", "1997", "com") >> True
    '@Return: Return boolean value of the result.

    Dim StartsWith As Boolean
    Dim EndsWith As Boolean
    StartsWith = IsStartsWith(OperationText, StartsWithWhat)
    
    If EndsWithWhat = vbNullString Then
        EndsWith = IsEndsWith(OperationText, StartsWithWhat)
    Else
        EndsWith = IsEndsWith(OperationText, EndsWithWhat)
    End If

    IsStartsWithOrEndsWith = (StartsWith Or EndsWith)

End Function


