sub phoneNumberFormat()
    Dim phoneNumbers as range
    set phoneNumbers = Selection
    For Each a in phoneNumbers
        If Len(a) = 10 Then
            a.Value = Left(a, 3) & "-" & Mid(a, 4, 3) & "-" Right(a, 4)
        Else    
            a.Vale = Replace(Replace(Replace(Replace(a.Value, "(", ""), ")", "-"), " ", ""), ".", "-")
        End If
        Debug.Print a 
        Next a
End Sub