VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ListItemPicker 
   Caption         =   "UserForm1"
   ClientHeight    =   2150
   ClientLeft      =   -315
   ClientTop       =   -1350
   ClientWidth     =   2550
   OleObjectBlob   =   "ListItemPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ListItemPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable


Option Explicit
Private Enum CloseBy
    User = 0
    Code = 1
    WindowsOS = 2
    TaskManager = 3
End Enum

Private Type TListItemPicker
    SelectedItem As Variant
    UnSelectedItems As Variant
End Type

Private This As TListItemPicker

Public Property Get UnSelectedItems() As Variant
    
    If IsObject(This.UnSelectedItems) Then
        Set UnSelectedItems = This.UnSelectedItems
    Else
        UnSelectedItems = This.UnSelectedItems
    End If
    
End Property

Public Property Let UnSelectedItems(ByVal RHS As Variant)
    This.UnSelectedItems = RHS
End Property

Public Property Set UnSelectedItems(ByVal RHS As Variant)
    Set This.UnSelectedItems = RHS
End Property

Public Property Get SelectedItem() As Variant
    
    If IsObject(This.SelectedItem) Then
        Set SelectedItem = This.SelectedItem
    Else
        SelectedItem = This.SelectedItem
    End If
    
End Property

Public Property Let SelectedItem(ByVal RHS As Variant)
    This.SelectedItem = RHS
End Property

Public Property Set SelectedItem(ByVal RHS As Variant)
    Set This.SelectedItem = RHS
End Property

Private Sub CancelButton_Click()
    SelectedItem = False
    UnSelectedItems = False
    Me.Hide
End Sub

Private Sub OkButton_Click()
    
    ' Check if an item is selected in the validation list
    If Me.ValidationListItems.ListIndex = -1 Then
        SelectedItem = vbNullString
    Else
        ' Get the selected item from the validation list
        SelectedItem = Me.ValidationListItems.List(Me.ValidationListItems.ListIndex, 0)
    End If
    
    ' Get all unselected items from the validation list
    UnSelectedItems = CustomizeListBox.GetAllUnSelectedItems(Me.ValidationListItems)
    
    ' Hide the form after processing
    Me.Hide
        
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    ' If the user attempts to close the form, handle the Close event
    If CloseMode = CloseBy.User Then
        ' Reset the SelectedItem and UnSelectedItems variables
        SelectedItem = False
        UnSelectedItems = False
        ' Hide the form and cancel the close event
        Me.Hide
        Cancel = True
    End If
    
End Sub

Private Sub UserForm_Activate()
    
    ' Set the startup position and position the form at the center of the Excel application window
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) * 0.5
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub

