Private Sub Workbook_Open()

    Dim f1 As String
    Dim f2 As String
    Dim columnn As String
    Dim temp As String
    
            
    
    Sheets("cover").Select
    f1 = Range("k4").Value
    f2 = Range("l4").Value
    columnn = "b"
    
    
    Sheets("data").Select
    
    Do While (Range(columnn + "2") <> 0)
        Range(columnn + "1106") = "=SQRT(sum(" + columnn + f1 + ":" + columnn + f2 + ")/(cover!M4+1))"
        temp = Range(columnn + "1").Offset(0, 1).Address
        columnn = Split(temp, "$")(1)
        
    Loop
    

End Sub
