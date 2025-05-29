Attribute VB_Name = "modListAvailableLambdas"
Option Explicit

Public Sub ListAvailableLambdas(DestinationCell As Range)
    
    Const METHOD_NAME As String = "ListAvailableLambdas"
    Context.ExtractContextFromCell DestinationCell, METHOD_NAME
    Dim InfoColl As Collection
    Set InfoColl = New Collection
    
    Dim CurrentName As Name
    For Each CurrentName In Context.Lambdas
        InfoColl.Add LAMBDAInfo.Create(CurrentName), CurrentName.Name
    Next CurrentName
    
    If InfoColl.Count = 0 Then
        MsgBox "No lambdas found in '" & DestinationCell.Worksheet.Parent.Name & "'.", vbInformation + vbOKOnly, "List Available Lambdas"
        Exit Sub
    End If
    
    Dim ResultArr As Variant
    ReDim ResultArr(1 To InfoColl.Count + 1, 1 To 6)
        
    ResultArr(1, 1) = "Name"
    ResultArr(1, 2) = "Comment"
    ResultArr(1, 3) = "Parameters"
    ResultArr(1, 4) = "Command Name"
    ResultArr(1, 5) = "Command Description"
    ResultArr(1, 6) = "Definition"
    
    If InfoColl.Count > 0 Then
    
        Dim RowIndex As Long
        RowIndex = 2
        
        Dim CurrentLambdaInfo As LAMBDAInfo
        For Each CurrentLambdaInfo In InfoColl
            
            With CurrentLambdaInfo
                ResultArr(RowIndex, 1) = .Name
                ResultArr(RowIndex, 2) = .Comment
                ResultArr(RowIndex, 3) = .Parameters
                ResultArr(RowIndex, 4) = .CommandName
                ResultArr(RowIndex, 5) = .CommandDescription
                ResultArr(RowIndex, 6) = .Definition
            End With
            
            RowIndex = RowIndex + 1
            
        Next CurrentLambdaInfo
        
    End If
    
    Dim DumpRange As Range
    Set DumpRange = DestinationCell.Resize(InfoColl.Count + 1, UBound(ResultArr, 2) - LBound(ResultArr, 2) + 1)
    
    If Not IsAllCellBlank(DumpRange) Then
        MsgBox "Unable to list Lambdas without overriding current values in cells.  Please try again.", vbExclamation + vbOKOnly, "List Available Lambdas"
        GoTo ExitMethod
    End If
    
    DumpRange.Value = ResultArr
    DumpRange.Worksheet.ListObjects.Add xlSrcRange, DumpRange, , xlYes
    AutoFitRange DumpRange, MaximumColumnWidth:=50, MinimumColumnWidth:=8
    DumpRange.WrapText = False
    With DumpRange.ListObject
        .Range.VerticalAlignment = xlTop
        .ListColumns("Comment").DataBodyRange.WrapText = True
        .ListColumns("Command Description").DataBodyRange.WrapText = True
        .ListColumns("Definition").DataBodyRange.WrapText = True
    End With
    
ExitMethod:
    Context.ClearContext METHOD_NAME
    Exit Sub
    
End Sub

