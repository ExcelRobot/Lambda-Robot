VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, SuperfluousAnnotationArgument, ImplicitlyTypedConst, UnrecognizedAnnotation
'@Folder "Lambda.Editor.Logger"
'@PredeclaredId
Option Explicit

Const QUOTATION_MARK = """"

Const JSON_PATTERN = QUOTATION_MARK & "{0}" & QUOTATION_MARK & " : " & _
QUOTATION_MARK & "{1}" & QUOTATION_MARK
Const JAGGED_MAPPER_JSON_PATTERN = QUOTATION_MARK & "{0}" & QUOTATION_MARK & " : " & "[" & "{1}" & "]"

Private Type TMapper
    Key As String
    Value As Variant
    ToJSON As String
End Type

Private This As TMapper

Public Property Get ToJSON() As String
    
    Dim JSONText As String
    Dim JaggedMapperCollection As Collection
    Dim CurrentMapper As Mapper
    If TypeName(This.Value) = "Collection" Then
        Set JaggedMapperCollection = This.Value
        If TypeName(JaggedMapperCollection.Item(1)) = "Mapper" Then
            For Each CurrentMapper In JaggedMapperCollection
                JSONText = JSONText & COMMA & CurrentMapper.ToJSON
            Next CurrentMapper
        End If
        JSONText = BeautifyString(JAGGED_MAPPER_JSON_PATTERN, Array(This.Key, JSONText))
    Else
        JSONText = BeautifyString(JSON_PATTERN, Array(This.Key, This.Value))
    End If
    This.ToJSON = JSONText
    ToJSON = This.ToJSON
    
End Property

Public Property Get Key() As String
    Key = This.Key
End Property

Public Property Let Key(ByVal RHS As String)
    This.Key = RHS
End Property

Public Property Get Value() As Variant
    
    If IsObject(This.Value) Then
        Set Value = This.Value
    Else
        Value = This.Value
    End If
    
End Property

Public Property Let Value(ByVal RHS As Variant)
    This.Value = RHS
End Property

Public Property Set Value(ByVal RHS As Variant)
    Set This.Value = RHS
End Property

Public Function Create(GivenKey As String, GivenValue As Variant) As Mapper
    
    Dim CurrentMapper As Mapper
    Set CurrentMapper = New Mapper
    With CurrentMapper
        .Key = GivenKey
        If IsObject(GivenValue) Then
            Set .Value = GivenValue
        Else
            .Value = GivenValue
        End If
    End With
    Set Create = CurrentMapper
    
End Function

'@Author : Md.Ismail Hosen
'@Description("This function will replace placeholder text with appropriate values")
'@Dependency("No Dependency")
'@ExampleCall : BeautifyString("Your Name : {1}   Your Age: {2}",Array("Md.Ismail Hosen", 24),1) >> Your Name : Md.Ismail Hosen   Your Age: 24
'@ExampleCall : BeautifyString("Your Name : {0}","Md.Ismail Hosen") >> Your Name : Md.Ismail Hosen
'@ExampleCall : BeautifyString("Your Name : {}","Md.Ismail Hosen") >> Your Name : {} >> Because no place holder value..So if your text has {} this will help
Private Function BeautifyString(Pattern As String, PlaceHolderValues As Variant _
                                                  , Optional StartNumber As Long = 0 _
                                                   , Optional PlaceHolder As String = "{#}") As String
    
    Dim PlaceHolderValue As Variant
    Dim CurrentPlaceHolder As String
    If IsArray(PlaceHolderValues) Then
        For Each PlaceHolderValue In PlaceHolderValues
            CurrentPlaceHolder = Replace(PlaceHolder, HASH_SIGN, StartNumber)
            Pattern = Replace(Pattern, CurrentPlaceHolder, PlaceHolderValue)
            StartNumber = StartNumber + 1
        Next PlaceHolderValue
    Else
        CurrentPlaceHolder = Replace(PlaceHolder, HASH_SIGN, StartNumber)
        Pattern = Replace(Pattern, CurrentPlaceHolder, PlaceHolderValues)
    End If
    BeautifyString = Pattern
    
End Function




