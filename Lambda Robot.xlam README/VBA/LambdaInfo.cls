VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LAMBDAInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Private Type TLAMBDAInfo
    Lambda As Name
    Name As String
    Comment As String
    Parameters As String
    CommandName As String
    CommandDescription As String
    Definition As String
End Type

Private this As TLAMBDAInfo

Public Property Get Lambda() As Name
    Set Lambda = this.Lambda
End Property

Public Property Set Lambda(ByVal RHS As Name)
    
    Set this.Lambda = RHS
    
    With this
        .Name = .Lambda.Name
        .Comment = .Lambda.Comment
        Dim MetadataInfo As Metadata
        Set MetadataInfo = Metadata.CreateLambdaMetadata( _
                           GetDependencyFunctionResult(.Lambda.RefersTo, LAMBDA_PARTS) _
                           , GetDependencyFunctionResult(.Lambda.RefersTo, LET_PARTS) _
                            , .Name, .Name, True)
        .CommandName = MetadataInfo.CommandName
        .CommandDescription = MetadataInfo.Description
        .Parameters = MetadataInfo.GetParametersForLambdaListing()
        .Definition = SINGLE_QUOTE & RemoveMetadataFromFormula(.Lambda.RefersTo)
    End With
    
End Property

Public Property Get Self() As LAMBDAInfo
    Set Self = Me
End Property

Public Function Create(ByVal Lambda As Name) As LAMBDAInfo

    With New LAMBDAInfo
        Set .Lambda = Lambda
        Set Create = .Self
    End With

End Function

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Get Comment() As String
    Comment = this.Comment
End Property

Public Property Get Parameters() As String
    Parameters = this.Parameters
End Property

Public Property Get CommandName() As String
    CommandName = this.CommandName
End Property

Public Property Get CommandDescription() As String
    CommandDescription = this.CommandDescription
End Property

Public Property Get Definition() As String
    Definition = this.Definition
End Property

