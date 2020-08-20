Sub APScript()
' Create headers
Cells(1, 1).Value = "External ID"
Cells(1, 2).Value = "Date"
Cells(1, 3).Value = "Account"
Cells(1, 4).Value = "Account Internal ID"
Cells(1, 5).Value = "Vendor Payee"
Cells(1, 6).Value = "Vendor Payee Internal ID"
Cells(1, 7).Value = "Vendor Bill"
Cells(1, 8).Value = "Vendor Bill Internal ID"
Cells(1, 9).Value = "Vendor Payment bills: Payment"
Cells(1, 10).Value = "A/P Account"
Cells(1, 11).Value = "A/P Account Internal ID"
Cells(1, 12).Value = "Subsidiary"
Cells(1, 13).Value = "Subsidiary Internal ID"
Cells(1, 14).Value = "Memo"
Cells(1, 15).Value = "Cost Center"
Cells(1, 16).Value = "Cost Center Internal ID"
Cells(1, 17).Value = "Department"
Cells(1, 18).Value = "Department Internal ID"
Cells(1, 19).Value = "Services"
Cells(1, 20).Value = "Services Internal ID"
Cells(1, 21).Value = "Posting Period"
Cells(1, 22).Value = "Check Number"
Dim i As Integer
Dim Last_Row As Long
Dim Subsidiary As Integer
Dim Vendor_Payee_Internal_ID As Long
' Last row determined by 5th column as columns 1-4 have no value
Last_Row = Cells(Rows.Count, 5).End(xlUp).Row
Time_Double = CDbl(Now())
For i = 2 To Last_Row
    'Column 12 contains subsidiary data
    Subsidiary = Cells(i, 12).Value
    'Column 6 contains vendor payee internal ID data
    Vendor_Payee_Internal_ID = Cells(i, 6).Value
    ' Prints data for column 1
    Cells(i, 1).Value = "VP" & Subsidiary & Vendor_Payee_Internal_ID & Time_Double
    ' Prints data for column 2 (date)
    Cells(i, 2).Value = Format(Now(), "mm/dd/yyyy")
    ' Prints data for columns 3 and 4 (Checkbook, Checkbook Internal ID)
    If Subsidiary = "12" Then
    Cells(i, 3).Value = "Chase Operating MA-8009"
    Cells(i, 4).Value = "991"
    End If
    If Subsidiary = "88" Then
    Cells(i, 3).Value = "Chase Operating LOG-3021"
    Cells(i, 4).Value = "444"
    End If
    If Subsidiary = "10" Then
    Cells(i, 3).Value = "Chase Operating SCPAA - 7723"
    Cells(i, 4).Value = "290"
    End If
    If Subsidiary = "50" Then
    Cells(i, 3).Value = "Chase Operating I/O - 5030"
    Cells(i, 4).Value = "228"
    End If
    If Subsidiary = "92" Then
    Cells(i, 3).Value = "Chase Operating Dist - 2118"
    Cells(i, 4).Value = "447"
    End If
    If Subsidiary = "11" Then
    Cells(i, 3).Value = "Chase Operating LTL - 9667"
    Cells(i, 4).Value = "145"
    End If
    If Subsidiary = "86" Then
  Cells(i, 3).Value = "Chase Operating Europe-4241"
    Cells(i, 4).Value = "219"
    End If
Next i
End Sub
