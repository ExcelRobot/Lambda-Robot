VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextConcatenator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable
'@Folder "Utility"
'@PredeclaredId
Option Explicit

Private Const BUFFER_SIZE As Long = 32
Private Const INCREMENT_PERCENT As Long = 50

Private Type TTextConcatenator
    Text As String
    SoFarFillIndex As Long
    TotalLength As Long
End Type

Private This As TTextConcatenator

Public Property Get Text() As String
    Text = Mid$(This.Text, 1, This.SoFarFillIndex)
End Property

Public Property Get Self() As TextConcatenator
    Set Self = Me
End Property

Public Function Create(ByVal Text As String) As TextConcatenator

    With New TextConcatenator
        .Concatenate Text
        Set Create = .Self
    End With

End Function

Public Sub Concatenate(ByVal JoinText As String)
    
    Dim TextLength As Long
    TextLength = Len(JoinText)
    If TextLength = 0 Then Exit Sub
    IncreaseCapacityIfNeeded TextLength
    Mid(This.Text, This.SoFarFillIndex + 1, TextLength) = JoinText
    This.SoFarFillIndex = This.SoFarFillIndex + TextLength
    
End Sub

Private Sub IncreaseCapacityIfNeeded(ByVal TextLength As Long)
    
    If This.SoFarFillIndex + TextLength < This.TotalLength Then Exit Sub
    
    ' We are planning to increase 50% but if needed length is more than that then increase differently.
    If This.SoFarFillIndex + TextLength > (1 + (INCREMENT_PERCENT / 100)) * This.TotalLength Then
        
        Dim ExtraBufferSize As Long
        If This.TotalLength = 0 Then
            ExtraBufferSize = BUFFER_SIZE
        Else
            If This.TotalLength = This.SoFarFillIndex Then
                ExtraBufferSize = BUFFER_SIZE
            Else
                ExtraBufferSize = (INCREMENT_PERCENT / 100) * This.TotalLength
            End If
        End If
            
        This.Text = This.Text + Space(This.SoFarFillIndex + TextLength - This.TotalLength + ExtraBufferSize)
        This.TotalLength = This.SoFarFillIndex + TextLength + ExtraBufferSize
            
    Else
        This.Text = This.Text & Space(This.TotalLength * (INCREMENT_PERCENT / 100))
        This.TotalLength = This.TotalLength * (1 + (INCREMENT_PERCENT / 100))
    End If
    
    
End Sub
