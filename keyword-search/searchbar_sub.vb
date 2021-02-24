'Example subroutine using keyWordFilter for search-as-you-type functionality.
'Triggered by the search bar textbox's OnChange event

Private Sub SearchBar_Change()
    Dim searchBarText As String
    Dim searchFields As String
    Dim ct As Control
    
    Set ct = Me!SearchBar
    searchFields = "Fields" '<----- Change this value to the fields being searched
    searchBarText = ct.Text
    
    Me.Filter = keyWordFilter(searchBarText, searchFields)
    Me.FilterOn = True
    
    ct = searchBarText
    ct.SetFocus
    ct.SelStart = Len(searchBarText)
End Sub
