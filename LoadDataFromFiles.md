Private Sub CommandButton1_Click()

'
' copyFromLogfile Macro
'

Dim path
Dim filename
Dim line As Integer


path = "C:\Users\chache\Documents\logfile\"
filename = Dir(path & ".\*.m")
line = 0


Set fso = CreateObject("scripting.filesystemobject")
Set f = fso.opentextfile(path & filename, 1)

a = f.readall
Data = Split(a, vbCr)

'find the line location
For ii = 0 To UBound(Data)
    If InStr(Data(ii), "front.msg_dfs_off") <> 0 Then
        line = ii
        Exit For
    End If
    
Next

    

'copy the data
msg_d = Data(line)
l = InStr(msg_d, "[")
r = InStr(msg_d, "]")
msg_dfs_off_f = Split(Mid(msg_d, l + 1, r - 1), " ")


msg_d = Data(line + 1)
l = InStr(msg_d, "[")
r = InStr(msg_d, "]")
msg_dfs_off_r = Split(Mid(msg_d, l + 1, r - 1), " ")


'go next
Sheets("data").Select
ActiveSheet.Range("a1") = "msg_dfs_off data from log files"

l = InStr(filename, "_fblog")
curvename = Mid(filename, 1, l - 1)
ActiveSheet.Range("a2") = curvename




End Sub
