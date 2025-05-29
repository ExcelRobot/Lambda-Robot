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

Private this As TTextConcatenator

Public Property Get Text() As String
    Text = Mid$(this.Text, 1, this.SoFarFillIndex)
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
    Mid(this.Text, this.SoFarFillIndex + 1, TextLength) = JoinText
    this.SoFarFillIndex = this.SoFarFillIndex + TextLength
    
End Sub

Private Sub IncreaseCapacityIfNeeded(ByVal TextLength As Long)
    
    If this.SoFarFillIndex + TextLength < this.TotalLength Then Exit Sub
    
    ' We are planning to increase 50% but if needed length is more than that then increase differently.
    If this.SoFarFillIndex + TextLength > (1 + (INCREMENT_PERCENT / 100)) * this.TotalLength Then
        
        Dim ExtraBufferSize As Long
        If this.TotalLength = 0 Then
            ExtraBufferSize = BUFFER_SIZE
        Else
            If this.TotalLength = this.SoFarFillIndex Then
                ExtraBufferSize = BUFFER_SIZE
            Else
                ExtraBufferSize = (INCREMENT_PERCENT / 100) * this.TotalLength
            End If
        End If
            
        this.Text = this.Text + Space(this.SoFarFillIndex + TextLength - this.TotalLength + ExtraBufferSize)
        this.TotalLength = this.SoFarFillIndex + TextLength + ExtraBufferSize
            
    Else
        this.Text = this.Text & Space(this.TotalLength * (INCREMENT_PERCENT / 100))
        this.TotalLength = this.TotalLength * (1 + (INCREMENT_PERCENT / 100))
    End If
    
    
End Sub

