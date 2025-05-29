Attribute VB_Name = "modExcelLabs"
Option Explicit
Option Private Module

Private Sub Test()
    
    Debug.Print IsLambdaCreatedInExcelLabs(ActiveWorkbook, "EAP.AddTwoNumber")
   
End Sub

Public Function IsLambdaCreatedInExcelLabs(ByVal FromBook As Workbook _
                                           , ByVal LambdaName As String) As Boolean
    
    Dim ExcelLabsLambdas As Collection
    Set ExcelLabsLambdas = GetExcelLabsLambdas(FromBook)
    
    If ExcelLabsLambdas Is Nothing Then Exit Function
    
    If ExcelLabsLambdas.Count = 0 Then Exit Function
    
    Dim CurrentLambda As Variant
    For Each CurrentLambda In ExcelLabsLambdas
        If CStr(CurrentLambda) = LambdaName Then
            IsLambdaCreatedInExcelLabs = True
            Exit Function
        End If
    Next CurrentLambda
    
End Function

Public Sub WriteStringToTextFile(Content As String, ToFilePath As String)
    
    Dim FileNo As Long
    FileNo = FreeFile()
    Open ToFilePath For Output As #FileNo
    Print #FileNo, Content
    Close #FileNo
        
End Sub


