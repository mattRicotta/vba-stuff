Public Function DConcat(Delimiter As String, Expression As String, Domain As String, OrderBy As String = "", Optional Criteria As Variant = "")
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    If Criteria<>"" Then
        Criteria = " WHERE " & Criteria
    End If

    If OrderBy="" Then
        OrderBy = "Expr1" 'Default order by expression
    End If
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT " & Expression & " AS Expr1 " _
      & "FROM " & Domain & Criteria _
      & "ORDER BY" OrderBy, _
        dbOpenSnapshot)
    With rs
        If Not .RecordCount > 0 Then
            GoTo ExitErr
        End If
        .MoveFirst
        Do While Not .EOF
            DConcat = DConcat & Delimiter & !Expr1
            .MoveNext
        Loop
    End With
    rs.Close

    'Trim leading delimiter from String
    DConcat = Right(DConcat, Len(DConcat) - Len(Delimiter))
    'Debug.Print DConcat

ExitErr:
    Set rs = Nothing
    Exit Function
ErrHandler:
    DConcat = ""
    Resume ExitErr
End Function

