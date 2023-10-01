Sub Summarize(sh As Worksheet, K As Long, Arr_D)

Dim i As Long, LastLine As Long, Arr_N()
    ' Find the last line of the sheet
    LastLine = sh.Range("B" & Rows.Count).End(xlUp).Row
    ' Arr_N size
    Arr_N = sh.Range("B6:D" & LastLine)
    
    ' Run from the first to the last line of the summaize sheet
    For i = 1 To UBound(Arr_N, 1)
        K = K + 1
        Arr_D(K, 2) = Arr_N(i, 2)
        Arr_D(K, 3) = Arr_N(i, 3)
        Arr_D(K, 4) = Arr_N(i, 4)
    Next
    
End Sub

Sub DataInjection()
Dim K As Long, sh As Worksheet, Arr_D()
    ' Detect total rows
    For Each sh In Worksheets
        If sh.Name <> "TONG HOP DOI CHIEU" And sh.Name <> "DANH SACH" Then
            LastLine = LastLine + sh.Range("B" & Rows.Count).End(xlUp).Row
        End If
    Next
    
    K = 0
    ReDim Arr_D(1 To LastLine, 2 To 4)
    For Each sh In Worksheets
        If sh.Name <> "TONG HOP DOI CHIEU" And sh.Name <> "DANH SACH" Then
            Call Summarize(Sheets(sh.Name), K, Arr_D)
        End If
    Next
    
    ' Clean data
    Sheet3.Range("B6:D1000000").Clear
    ' Insert
    Sheet3.Range("B6").Resize(K, 4) = Arr_D
    
End Sub
