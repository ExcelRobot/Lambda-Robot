VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Lambda.Editor.Metadata.Presenter.Interface"
Option Explicit

Public SourceName As String
Public GistURL As String
Public LambdaName As String
Public Description As String
Public CommandName As String
Public IsCancelled As Boolean

Public Function GetOverview() As String
    
End Function

Public Sub UpdateMetadataInFormula()
    
End Sub

Public Sub GenerateGist()

End Sub

Public Function GetExportPreview() As String
    
End Function

Public Sub ExpandCollapse(ByVal IsExpand As Boolean)
    
End Sub

