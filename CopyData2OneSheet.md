Sub CopyDate2Workdata()
'
' CopyDate2Workdata
'

'
    Dim Column_no As String
    Dim temp As String
    
    
        
    Column_no = "b"
        
    For Each sh In Worksheets
        If sh.Name <> "cover" And sh.Name <> "work data" And sh.Range("a1").Value <> "" Then
        sh.Select
        Range("C2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Sheets("work data").Select
        Range(Column_no + "2").Select
        ActiveSheet.Paste
        Range(Column_no + "1") = sh.Name
        
'        ii = Asc(Column_no)
'        ii = ii + 1
'        Column_no = Chr(ii)
        
        temp = Range(Column_no + "1").Offset(0, 1).Address
        Column_no = Split(temp, "$")(1)
                
              
        
        End If
    Next
    
    
    
End Sub

