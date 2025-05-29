VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyValuePairCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, ProcedureNotUsed
'@Folder "Lambda.Editor.Metadata.Presenter.Impl"

Option Explicit

Private Type TKeyValuePairCollection
    KeyValuePairCollection As Collection
End Type

Private this  As TKeyValuePairCollection

' Implements the ICRUD interface
Implements ICRUD

Public Property Get KeyValuePairCollection() As Collection
    ' Returns the KeyValuePairCollection property
    Set KeyValuePairCollection = this.KeyValuePairCollection
End Property

Private Sub ICRUD_ApplyOperationOnGroup(ByVal ToGroup As Group)
    
    ' Raising an error as group operations are not supported in this context
    Err.Raise 5
    
End Sub

Private Function ICRUD_ReadAll() As Variant
    
    ' Reads all key-value pairs in the KeyValuePairCollection, returns Empty if no pairs exist
    Logger.Log TRACE_LOG, "Enter KeyValuePairCollection.ICRUD_ReadAll"
    If this.KeyValuePairCollection.Count = 0 Then
        ICRUD_ReadAll = Empty
    Else
        ICRUD_ReadAll = modUtility.GetObjectsPropertyValue(this.KeyValuePairCollection _
                                                           , modUtility.VectorToCollection(Array("Key" _
                                                                                                 , "Value")), False)
    End If
    Logger.Log TRACE_LOG, "Exit KeyValuePairCollection.ICRUD_ReadAll"
    
End Function

Private Sub ICRUD_Swap(ByVal FirstKey As String, ByVal SecondKey As String)
    
    ' Swaps the position of two KeyValuePair objects in the KeyValuePairCollection based on their keys
    Logger.Log TRACE_LOG, "Enter KeyValuePairCollection.ICRUD_Swap"
    Dim CurrentItem As Mapper
    ' Temporary storage for the Mapper object being swapped
    Set CurrentItem = this.KeyValuePairCollection.Item(FirstKey)
    Dim FirstItemIndex As Long
    ' Get the index of the first item
    FirstItemIndex = FindIndexInCollection(this.KeyValuePairCollection, FirstKey)
    Dim SecondItemIndex As Long
    ' Get the index of the second item
    SecondItemIndex = FindIndexInCollection(this.KeyValuePairCollection, SecondKey)
    ' Remove the first item and add it to the position of the second item
    this.KeyValuePairCollection.Remove FirstKey
    this.KeyValuePairCollection.Add CurrentItem, FirstKey, , SecondItemIndex - 1
    ' Remove the second item and add it to the position of the first item
    Set CurrentItem = this.KeyValuePairCollection.Item(SecondKey)
    this.KeyValuePairCollection.Remove SecondKey
    this.KeyValuePairCollection.Add CurrentItem, SecondKey, FirstItemIndex
    Logger.Log TRACE_LOG, "Exit KeyValuePairCollection.ICRUD_Swap"
    
End Sub

Private Sub ICRUD_AddNew(ByVal Key As String, ByVal Value As String)
    
    ' Adds a new key-value pair to the KeyValuePairCollection
    Logger.Log TRACE_LOG, "Enter KeyValuePairCollection.ICRUD_AddNew"
    Dim CurrentMapper As Mapper
    ' Create a new Mapper object with the provided key and value
    Set CurrentMapper = Mapper.Create(Key, Value)
    this.KeyValuePairCollection.Add CurrentMapper, Key
    Logger.Log TRACE_LOG, "Exit KeyValuePairCollection.ICRUD_AddNew"
    
End Sub

Private Sub ICRUD_Update(ByVal OldKey As String, ByVal NewKey As String, ByVal Value As String)
    
    ' Updates the key and value of a KeyValuePair in the KeyValuePairCollection
    Logger.Log TRACE_LOG, "Enter KeyValuePairCollection.ICRUD_Update"
    Dim AddAtIndex As Long
    ' Get the index of the item to be updated
    AddAtIndex = FindIndexInCollection(this.KeyValuePairCollection, OldKey)
    Dim CurrentMapper As Mapper
    ' Get the Mapper object to be updated
    Set CurrentMapper = this.KeyValuePairCollection.Item(OldKey)
    ' Update the key and value
    CurrentMapper.Key = NewKey
    CurrentMapper.Value = Value
    
    If OldKey = NewKey Then Exit Sub

    ' Remove the old KeyValuePair and add the updated one at the same index
    this.KeyValuePairCollection.Add CurrentMapper, NewKey, , AddAtIndex
    this.KeyValuePairCollection.Remove OldKey
    Logger.Log TRACE_LOG, "Exit KeyValuePairCollection.ICRUD_Update"
    
End Sub

Private Function FindIndexInCollection(ByVal GivenCollection As Collection, ByVal Key As String) As Long
    
    ' Finds the index of a KeyValuePair in the GivenCollection based on the key
    Logger.Log TRACE_LOG, "Enter KeyValuePairCollection.FindIndexInCollection"
    Dim FoundAtIndex As Long
    ' Default value if the key is not found
    FoundAtIndex = 0
    Dim CurrentItemIndex As Long
    ' Iterate through the GivenCollection to find the key
    For CurrentItemIndex = 1 To GivenCollection.Count
        If GivenCollection.Item(CurrentItemIndex).Key = Key Then
            FoundAtIndex = CurrentItemIndex
            Exit For
        End If
    Next CurrentItemIndex
    
    ' Return the index of the found key
    FindIndexInCollection = FoundAtIndex
    Logger.Log TRACE_LOG, "Exit KeyValuePairCollection.FindIndexInCollection"
    
End Function

Private Sub ICRUD_Delete(ByVal Key As String)
    
    ' Removes a KeyValuePair from the KeyValuePairCollection based on the key
    Logger.Log TRACE_LOG, "Enter KeyValuePairCollection.ICRUD_Delete"
    this.KeyValuePairCollection.Remove Key
    Logger.Log TRACE_LOG, "Exit KeyValuePairCollection.ICRUD_Delete"
    
End Sub

Private Sub Class_Initialize()
    
    ' Initializes a new KeyValuePairCollection when the class is instantiated
    Logger.Log TRACE_LOG, "Enter KeyValuePairCollection.Class_Initialize"
    Set this.KeyValuePairCollection = New Collection
    Logger.Log TRACE_LOG, "Exit KeyValuePairCollection.Class_Initialize"
    
End Sub

Private Sub Class_Terminate()
    
    ' Cleans up the KeyValuePairCollection when the class is terminated
    Logger.Log TRACE_LOG, "Enter KeyValuePairCollection.Class_Terminate"
    Set this.KeyValuePairCollection = Nothing
    Logger.Log TRACE_LOG, "Exit KeyValuePairCollection.Class_Terminate"
    
End Sub


