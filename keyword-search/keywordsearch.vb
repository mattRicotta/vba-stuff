'This function will return a string that can be used in a form filter or query
'WHERE clause. The search input can be taken from a form control.

'searchTerms = search bar text/input
'searchFields = String() array of the query/table field names to be searched


Public Function keyWordFilter(searchTerms As String, ParamArray searchFields()) As String
    Dim searchField As String
    Dim keyWords() As String
    Dim keyWord As Variant
    Dim keyWordFilterStr As String
    Dim keyWordFilterSubStr As String
    Dim i As Long

    If Trim(searchTerms) & "" = "" Then
            keyWordFilterStr = vbNullString
    Else
        keyWords = Split(Trim(searchTerms), " ")
        For Each keyWord In keyWords()
            keyWordFilterSubStr = vbNullString
            For i = LBound(searchFields) To UBound(searchFields)
                searchField = Trim(searchFields(i))
                keyWordFilterSubStr = keyWordFilterSubStr & _
                    " OR [" & searchField & "] LIKE ""*" & keyWord & "*"""
                'Debug.Print "keyWordFilterSubStr: " & keyWordFilterStr
            Next
            keyWordFilterSubStr = Right(keyWordFilterSubStr, Len(keyWordFilterSubStr) - 4)
            keyWordFilterStr = keyWordFilterStr & " AND (" & keyWordFilterSubStr & ") "
            'Debug.Print "keyWordFilterStr: " & keyWordFilterStr
        Next
    End If

    If keyWordFilterStr & "" <> "" Then
        keyWordFilterStr = Right(keyWordFilterStr, Len(keyWordFilterStr) - 5)
    End If

    keyWordFilter = keyWordFilterStr
    Debug.Print "Filter String: " & keyWordFilter
End Function