VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomizeListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed, MissingMemberAnnotation, SuperfluousAnnotationArgument, UnrecognizedAnnotation, ProcedureNotUsed, UnassignedVariableUsage, MissingAnnotationArgument
'@Folder("Customization.ListBox")
'@PredeclaredId
Option Explicit

'@Date : 01 March 2022 11:41:16 PM
'@ModuleDescription : This is a custom class for listbox. You can Listen for Listbox selection change Event.
'                                   It has some utility Sub and Function for easy use of listbox in different situation.
'                                   As it has predeclaredId set to true you can call sub or function without having instance of this class.
'                                   For Example : CustomizeListBox.InvertSelectAllOfListbox is a valid call.



Public WithEvents GivenListBox As MSForms.ListBox
Attribute GivenListBox.VB_VarHelpID = -1

Event IndexSelectionChange(OldIndex As Long, NewIndex As Long)

Private Type TCustomizeListBox
    SelectedIndexes As Collection
    PreviousSelectedIndex As Long
End Type

Private this As TCustomizeListBox

Private Sub Class_Initialize()
    Set this.SelectedIndexes = New Collection
    this.PreviousSelectedIndex = -1
End Sub

Public Property Get SelectedIndexes() As Collection
    Set SelectedIndexes = this.SelectedIndexes
End Property

Private Sub GivenListBox_Change()
    UpdateSelectionData
End Sub

Private Sub UpdateSelectionData()
    
    'Creating a new collection for newly selected items
    Dim NewSelectedItemsIndex As Collection
    Set NewSelectedItemsIndex = New Collection
    Dim Index As Long
    Dim NewSelectedIndex As Long
    NewSelectedIndex = -1
    Dim PreviouslySelectedItemCount As Long
    PreviouslySelectedItemCount = this.SelectedIndexes.Count
    'Looping through all list box items to check the selected state
    For Index = 0 To GivenListBox.ListCount - 1
        If GivenListBox.Selected(Index) Then
            'If an item is selected, it's added to the new collection
            NewSelectedItemsIndex.Add Index, CStr(Index)
            'If the same item was previously selected, it's removed from old collection
            If IsExistInCollection(this.SelectedIndexes, CStr(Index)) Then
                this.SelectedIndexes.Remove CStr(Index)
            Else
                NewSelectedIndex = Index
            End If
        End If
    Next Index
    'Handling the case of multi-selection list box where only one item is now selected
    If PreviouslySelectedItemCount > 1 And GivenListBox.MultiSelect = fmMultiSelectExtended _
       And NewSelectedItemsIndex.Count = 1 Then
        NewSelectedIndex = NewSelectedItemsIndex.Item(1)
    End If
    'Handling the case of single selection remaining constant
    If PreviouslySelectedItemCount = 1 And NewSelectedItemsIndex.Count = 1 _
       And this.SelectedIndexes.Count = 0 Then NewSelectedIndex = NewSelectedItemsIndex.Item(1)
    'Setting the updated selection indexes
    Set this.SelectedIndexes = NewSelectedItemsIndex
    'If there is a change in selection, an event is raised
    If this.PreviousSelectedIndex <> NewSelectedIndex Then RaiseEvent IndexSelectionChange(this.PreviousSelectedIndex _
                                                                                           , NewSelectedIndex)
    'Updating the previously selected index
    this.PreviousSelectedIndex = NewSelectedIndex
    
End Sub

'@Description("This is for testing if a key is present in a collection or not.")
'@Dependency("No Dependency")
'@ExampleCall : IsExistInCollection(InputCollection,"Key")
'@Date : 14 October 2021 07:04:18 PM
Public Function IsExistInCollection(ByVal GivenCollection As Collection, ByVal Key As String) As Boolean
    
    Const INVALID_PROCEDURE_CALL_OR_ARGUMENT_ERROR_NUMBER As Long = 5
    'Using an error trap to determine existence
    On Error GoTo NotExist
    Dim Item  As Variant
    Set Item = GivenCollection.Item(Key)
    
NotExist:
    'If no error, the item exists
    IsExistInCollection = Not (Err.Number = INVALID_PROCEDURE_CALL_OR_ARGUMENT_ERROR_NUMBER And Err.Number <> 0)
    
End Function

'@Description("This will select currently un-selected item.")
'@Dependency("No Dependency")
'@ExampleCall :InvertSelectAllOfListbox Me.QueryListBox
'@Date : 06 March 2022 07:28:56 AM
'@PossibleError :
'Function to invert all selections in a given listbox
Public Sub InvertSelectAllOfListbox(ByVal GivenListBox As MSForms.ListBox)
    With GivenListBox
        'Loop over all items in the listbox
        Dim CurrentRowIndex As Long
        For CurrentRowIndex = 0 To .ListCount - 1
            'Invert the selection status of the current item
            .Selected(CurrentRowIndex) = Not .Selected(CurrentRowIndex)
        Next CurrentRowIndex
    End With
End Sub

'@Description("This will either select or Un-Select all based on the second argument.")
'@Dependency("No Dependency")
'@ExampleCall : For Select All : SelectOptionAllOfListbox Me.QueryListBox, True
'@ExampleCall : For Un-Select All : SelectOptionAllOfListbox Me.QueryListBox, False
'@Date : 06 March 2022 07:29:38 AM
'@PossibleError :
'Function to select or deselect all options in a given listbox
Public Sub SelectOptionAllOfListbox(ByVal GivenListBox As MSForms.ListBox, Optional IsSelect As Boolean = True)
    With GivenListBox
        'Loop over all items in the listbox
        Dim CurrentRowIndex As Long
        For CurrentRowIndex = 0 To .ListCount - 1
            'Set the selection status of the current item based on IsSelect
            .Selected(CurrentRowIndex) = IsSelect
        Next CurrentRowIndex
    End With
End Sub

'@Description("This will filter listbox item based on the FilterText.If FilterText is vbNullString then it will just add item to the listbox and exit sub.This is zero based index as listbox is zero based index. So pass accordingly.")
'@Dependency("No Dependency")
'@ExampleCall : FilterListBox Me.QueryListBox, this.QueryList, WorkbookList.Value, 0
'@Date : 06 March 2022 07:31:35 AM
'@PossibleError :
'Procedure to filter items in a given ListBox based on the FilterText and FilterByColumn
Public Sub FilterListBox(ByVal GivenListBox As MSForms.ListBox, ByVal FilterText As String, _
                         ByVal FilterByColumn As Long, Optional HasHeader As Boolean = False)
    
    'If no FilterText provided, then exit the subroutine
    If FilterText = vbNullString Then Exit Sub
    
    'Determines the start index of the filter operation depending on if the ListBox has headers
    Dim ToIndex As Long
    ToIndex = IIf(HasHeader, 1, 0)
    Dim CurrentItemIndex As Long
    With GivenListBox
        'Loop through the ListBox from end to start
        For CurrentItemIndex = .ListCount - 1 To ToIndex Step -1
            'If the FilterText is not in the current ListBox item, then remove the item
            If InStr(1, .List(CurrentItemIndex, FilterByColumn), FilterText, vbTextCompare) = 0 Then
                .RemoveItem CurrentItemIndex
            End If
        Next CurrentItemIndex
    End With
    
End Sub

'@Description("This will filter listbox item based on the two column FilterText.This is zero based index as listbox is zero based index. So pass accordingly.")
'@Dependency("No Dependency")
'@ExampleCall : FilterListBoxBy2Column Me.QueryListBox, this.QueryList, Me.WorkbookList.Value, 0, Me.QueryTypeCombobox.Value, 2
'@Date : 06 March 2022 07:32:57 AM
'@PossibleError :

'Procedure to filter items in a given ListBox based on the FirstFilterText, FirstFilterByColumn, SecondFilterText, and SecondFilterByColumn
Public Sub FilterListBoxBy2Column(ByVal GivenListBox As MSForms.ListBox _
                                  , ByVal FirstFilterText As String _
                                   , ByVal FirstFilterByColumn As Long _
                                    , ByVal SecondFilterText As String _
                                     , ByVal SecondFilterByColumn As Long _
                                      , Optional HasHeader As Boolean = False)
    
    'Determines the start index of the filter operation depending on if the ListBox has headers
    Dim ToIndex As Long
    ToIndex = IIf(HasHeader, 1, 0)
    
    Dim CurrentItemIndex As Long
    With GivenListBox
        'Loop through the ListBox from end to start
        For CurrentItemIndex = .ListCount - 1 To ToIndex Step -1
            'If neither the FirstFilterText nor the SecondFilterText is in the current ListBox item, then remove the item
            If InStr(1, .List(CurrentItemIndex, FirstFilterByColumn), FirstFilterText, vbTextCompare) = 0 Or _
               InStr(1, .List(CurrentItemIndex, SecondFilterByColumn), SecondFilterText, vbTextCompare) = 0 Then
                
                .RemoveItem CurrentItemIndex
                    
            End If
        Next CurrentItemIndex
    End With
    
End Sub

'@Description("This will Filter Listbox by any column.")
'@Dependency("No Dependency")
'@ExampleCall : FilterListBoxByAny Me.QueryListBox,"CSVQuery"
'@Date : 06 March 2022 07:33:47 AM
'@PossibleError :
'Procedure to filter items in a given ListBox based on the GivenKeyword
Public Sub FilterListBoxByAny(ByVal GivenListBox As MSForms.ListBox _
                              , ByVal GivenKeyword As String, Optional HasHeader As Boolean = False)
    
    'Exit procedure if no keyword is given
    If GivenKeyword = vbNullString Then
        Exit Sub
    End If
    
    'Determines the start index of the filter operation depending on if the ListBox has headers
    Dim ToIndex As Long
    ToIndex = IIf(HasHeader, 1, 0)
    
    With GivenListBox
        Dim CurrentRow As Long
        Dim IsMatch As Boolean
        'Loop through the ListBox from end to start
        For CurrentRow = .ListCount - 1 To ToIndex Step -1
            Dim CurrentColumn As Long
            'Loop through each column in the current row
            For CurrentColumn = 0 To .ColumnCount - 1
                'If the GivenKeyword is found in the current ListBox cell, set IsMatch to True and exit the loop
                If InStr(1, .List(CurrentRow, CurrentColumn), GivenKeyword, vbTextCompare) <> 0 Then
                    IsMatch = True
                    Exit For
                End If
            Next CurrentColumn
            'If the GivenKeyword was not found in the current row, remove the row
            If Not IsMatch Then
                .RemoveItem CurrentRow
            End If
            IsMatch = False
        Next CurrentRow
    End With
    
End Sub

'@Description("It will only get the first selected item index")
'@Dependency("No Dependency")
'@ExampleCall : GetSelectedItemIndex Me.QueryListBox
'@Date : 06 March 2022 07:34:45 AM
'@PossibleError :
'Function to get the index of the selected item in a given ListBox
Public Function GetSelectedItemIndex(ByVal GivenListBox As MSForms.ListBox) As Long
    
    'Flag variable to check if at least one row is selected
    Dim IsAtleastOneRowSelected As Boolean

    'Iterate through all the items in the ListBox
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = 0 To GivenListBox.ListCount - 1

        'If the current item is selected, set the flag to True and exit the loop
        If GivenListBox.Selected(CurrentRowIndex) Then
            IsAtleastOneRowSelected = True
            Exit For
        End If
    Next CurrentRowIndex

    'If at least one row is selected, return the index of the selected item
    'If no item is selected, return -1
    If IsAtleastOneRowSelected Then
        GetSelectedItemIndex = CurrentRowIndex
    Else
        GetSelectedItemIndex = -1
    End If
    
End Function

'@Description("This will return all the selected Item Index")
'@Dependency("No Dependency")
'@ExampleCall : GetAllSelectedItemIndex Me.QueryListBox
'@Date : 06 March 2022 07:34:45 AM
'@PossibleError :
'Function to get the indices of all selected items in a given ListBox
Public Function GetAllSelectedItemIndex(ByVal GivenListBox As MSForms.ListBox) As Collection
    
    'Initialize a new collection to store the indices of all selected items
    Dim AllSelectedIndex As Collection
    Set AllSelectedIndex = New Collection
    
    'Iterate through all items in the ListBox
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = 0 To GivenListBox.ListCount - 1

        'If the current item is selected, add its index to the collection
        If GivenListBox.Selected(CurrentRowIndex) Then
            AllSelectedIndex.Add CurrentRowIndex
        End If
    Next CurrentRowIndex

    'Return the collection of indices
    Set GetAllSelectedItemIndex = AllSelectedIndex
    
End Function

'@Description("This will return all not selected Item Index")
'@Dependency("No Dependency")
'@ExampleCall : GetAllUnSelectedItemIndex Me.QueryListBox
'@Date : 06 March 2022 07:34:45 AM
'@PossibleError :
Public Function GetAllUnSelectedItemIndex(ByVal GivenListBox As MSForms.ListBox) As Collection
    
    Dim AllNotSelectedIndex As Collection
    Set AllNotSelectedIndex = New Collection
    
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = 0 To GivenListBox.ListCount - 1
        If Not GivenListBox.Selected(CurrentRowIndex) Then
            AllNotSelectedIndex.Add CurrentRowIndex
        End If
    Next CurrentRowIndex
    Set GetAllUnSelectedItemIndex = AllNotSelectedIndex
    
End Function

'@Description("This will remove selected Items from the listbox")
'@Dependency("No Dependency")
'@ExampleCall : RemoveSelectedFromListbox Me.QueryListBox
'@Date : 06 March 2022 07:37:44 AM
'@PossibleError :
Public Sub RemoveSelectedFromListbox(ByVal GivenListBox As MSForms.ListBox)
    
    'Iterate over all items in the ListBox from the last to the first
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = GivenListBox.ListCount - 1 To 0 Step -1
        'Check if the current item is selected
        If GivenListBox.Selected(CurrentRowIndex) Then
            'If the item is selected, remove it from the ListBox
            GivenListBox.RemoveItem CurrentRowIndex
        End If
    Next CurrentRowIndex
     
End Sub

'@Description("This will update selected Row data with. This is for single selection with 1 or More column")
'@Dependency("No Dependency")
'@ExampleCall : UpdateSelectedRow Me.QueryListBox,UpdatedData
'@Date : 06 March 2022 07:38:40 AM
'@PossibleError :
Public Sub UpdateSelectedRow(ByVal GivenListBox As MSForms.ListBox, GivenValues As Variant)
    
    'Get the index of the selected row
    Dim SelectedRowIndex As Long
    SelectedRowIndex = GetSelectedItemIndex(GivenListBox)
    
    'If the GivenValues is an array
    If IsArray(GivenValues) Then
        Dim Counter As Long
        Dim CurrentElement As Variant
        'Iterate over each element in the GivenValues array
        For Each CurrentElement In GivenValues
            'Update the selected row in the ListBox with the current element in the array
            With GivenListBox
                .List(SelectedRowIndex, Counter) = CurrentElement
                Counter = Counter + 1
            End With
        Next CurrentElement
    Else
        'If the GivenValues is not an array, simply update the selected row with the GivenValues
        GivenListBox.List(SelectedRowIndex, Counter) = GivenValues
    End If
    
End Sub

'@Description("This will retrive all the selected rows data into an array.")
'@Dependency("No Dependency")
'@ExampleCall : GetAllSelectedItems Me.SourceListViewer
'@Date : 11 March 2022 12:44:20 AM
'@PossibleError :
Public Function GetAllSelectedItems(ByVal GivenListBox As MSForms.ListBox) As Variant
    
    'Get all selected item indices from the listbox
    Dim AllSelectedItemsIndex As Collection
    Set AllSelectedItemsIndex = GetAllSelectedItemIndex(GivenListBox)
    
    'If no items are selected, then exit the function
    If AllSelectedItemsIndex.Count = 0 Then Exit Function

    'Prepare an array to store all selected items
    Dim Result As Variant
    ReDim Result(1 To AllSelectedItemsIndex.Count, 1 To GivenListBox.ColumnCount)
    Dim Counter As Long
    
    'Iterate over all selected item indices
    Dim CurrentRowIndex As Variant
    For Each CurrentRowIndex In AllSelectedItemsIndex
        Counter = Counter + 1
        'For each selected row, iterate over all its columns
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = 0 To GivenListBox.ColumnCount - 1
            'Store each selected item in the Result array
            Result(Counter, CurrentColumnIndex + 1) = GivenListBox.List(CurrentRowIndex, CurrentColumnIndex)
        Next CurrentColumnIndex
    Next CurrentRowIndex
    'Return the Result array containing all selected items
    GetAllSelectedItems = Result
    
End Function

'@Description("This will retrive all not selected rows data into an array.")
'@Dependency("No Dependency")
'@ExampleCall : GetAllUnSelectedItems Me.SourceListViewer
'@Date : 11 March 2022 12:44:20 AM
'@PossibleError :
Public Function GetAllUnSelectedItems(ByVal GivenListBox As MSForms.ListBox) As Variant
    
    Dim AllNotSelectedItemsIndex As Collection
    Set AllNotSelectedItemsIndex = GetAllUnSelectedItemIndex(GivenListBox)
    If AllNotSelectedItemsIndex.Count = 0 Then Exit Function

    Dim Result As Variant
    ReDim Result(1 To AllNotSelectedItemsIndex.Count, 1 To GivenListBox.ColumnCount)
    Dim Counter As Long
    
    Dim CurrentRowIndex As Variant
    For Each CurrentRowIndex In AllNotSelectedItemsIndex
        Counter = Counter + 1
        Dim CurrentColumnIndex As Long
        For CurrentColumnIndex = 0 To GivenListBox.ColumnCount - 1
            Result(Counter, CurrentColumnIndex + 1) = GivenListBox.List(CurrentRowIndex, CurrentColumnIndex)
        Next CurrentColumnIndex
    Next CurrentRowIndex
    GetAllUnSelectedItems = Result
    
End Function

'@Description("It will remove or clear the listbox.")
'@Dependency("No Dependency")
'@ExampleCall : RemoveAll Me.SourceDataViewerListBox
'@Date : 11 March 2022 09:17:40 PM
'@PossibleError :
Public Sub RemoveAll(GivenListBox As MSForms.ListBox)
    'Loop will continue as long as there are items in the listbox
    Do While GivenListBox.ListCount > 0
        'Removes the first item in the listbox
        GivenListBox.RemoveItem 0
    Loop
End Sub

'@Description("This will select or deselect the first item of the listbox")
'@Dependency("No Dependency")
'@ExampleCall : SelectOrDeselectFirstItemOfListbox Me.SourceDataViewerListBox
'@Date : 11 March 2022 09:17:40 PM
'@PossibleError :
Public Sub SelectOrDeselectFirstItemOfListbox(ByVal GivenListBox As MSForms.ListBox _
                                              , Optional IsSelect As Boolean = True)
    
    With GivenListBox
        'Check if the listbox has any items
        If .ListCount > 0 Then
            'Depending on the value of IsSelect, the first item of the listbox will be selected or deselected
            .Selected(0) = IsSelect
        End If
    End With
    
End Sub

'@Description("This will select or deselect the last item of the listbox")
'@Dependency("No Dependency")
'@ExampleCall : SelectOrDeselectFirstItemOfListbox Me.SourceDataViewerListBox
'@Date : 11 March 2022 09:17:40 PM
'@PossibleError :
Public Sub SelectOrDeselectLastItemOfListbox(ByVal GivenListBox As MSForms.ListBox _
                                             , Optional IsSelect As Boolean = True)
    
    With GivenListBox
        'Check if the listbox has any items
        If .ListCount > 0 Then
            'Depending on the value of IsSelect, the last item of the listbox will be selected or deselected
            .Selected(.ListCount - 1) = IsSelect
        End If
    End With
    
End Sub


