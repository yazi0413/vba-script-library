Private Sub CommandButton1_Click()

'
    ' copyFromLogfile Macro
    '

    Dim path
    Dim filename
    Dim line_f As Integer
    Dim line_r As Integer
    Dim fre(127) As Double
    Dim l_index_f As Integer
    Dim l_index_r As Integer
    Dim dual_or_not As Boolean
    
    Sheets("calibration").Visible = True
    Sheets("front").Visible = True
    Sheets("rear").Visible = True


    'interpolate the calibration value
    Call interpolate


    'path = "C:\Users\chache\Documents\logfile\"
    path = Sheets("import").Range("g18")
    path = path & "\"
    filename = Dir(path & ".\*.m")


    'the interped frequency
    freq_end = 7812
    If Sheets("import").Range("f16") = "Dooku" Then
        freq_end = 10417
    End If


    For ii = 1 To 128
        fre(ii - 1) = 1 + (ii - 1) * freq_end / 128
    Next

    Sheets("output").Select
    ActiveSheet.Cells.Select
    Selection.Delete Shift:=xlUp
'    make the header in the sheets of front & rear
    Sheets("front").Select
    ActiveSheet.Cells.Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("a1") = "msg_dfs_off data from log files"
    ActiveSheet.Range("a1").Select
    Selection.Font.Bold = True
    ActiveSheet.Range("a2") = "frequency"
    ActiveSheet.Range("b2").Resize(1, UBound(fre) + 1).Value = fre

    Sheets("rear").Select
    ActiveSheet.Cells.Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("a1") = "msg_dfs_off data from log files"
    ActiveSheet.Range("a1").Select
    Selection.Font.Bold = True
    ActiveSheet.Range("a2") = "frequency"
    ActiveSheet.Range("b2").Resize(1, UBound(fre) + 1).Value = fre


    l_index_f = 3
    l_index_r = 3



    Do While filename <> ""

        line_f = 0
        line_r = 0
        dual_or_not = True

        Set fso = CreateObject("scripting.filesystemobject")
        Set f = fso.opentextfile(path & filename, 1)

        a = f.readall
        Data = Split(a, vbCr)

'        find the front and rear line location
        For ii = 0 To UBound(Data)
            If InStr(Data(ii), "front.msg_dfs_off") <> 0 Then
                line_f = ii
                Exit For
            End If

        Next

        For ii = 0 To UBound(Data)
            If InStr(Data(ii), "rear.msg_dfs_off") <> 0 Then
                line_r = ii
                Exit For
            End If
        Next

    'judge dual mic or not
        If line_r = 0 Then
            dual_or_not = False
        End If

        If line_f <> 0 Then


            'copy the data
            msg_d = Data(line_f)
            l = InStr(msg_d, "[")
            r = InStr(msg_d, "]")
            msg_dfs_off_f = Split(Mid(msg_d, l + 1, r - l - 1), " ")

        '    dual microphone or not

            If dual_or_not Then
                msg_d = Data(line_r)
                l = InStr(msg_d, "[")
                r = InStr(msg_d, "]")
                msg_dfs_off_r = Split(Mid(msg_d, l + 1, r - l - 1), " ")
            End If

            l = InStr(filename, "_fblog")
            curvename = Mid(filename, 1, l - 1)

            Sheets("front").Select
            ActiveSheet.Range("a" & CStr(l_index_f)) = curvename & "_front"
            If dual_or_not Then
                Sheets("rear").Select
                ActiveSheet.Range("a" & CStr(l_index_r)) = curvename & "_rear"
            End If
            Sheets("front").Select
            ActiveSheet.Range("b" & CStr(l_index_f)).Resize(1, UBound(msg_dfs_off_f) + 1).Value = msg_dfs_off_f
            If dual_or_not Then
                Sheets("rear").Select
                ActiveSheet.Range("b" & CStr(l_index_r)).Resize(1, UBound(msg_dfs_off_r) + 1).Value = msg_dfs_off_r
            End If

            If dual_or_not Then
                l_index_f = l_index_f + 1
                l_index_r = l_index_r + 1
            Else
                l_index_f = l_index_f + 1
            End If

        End If

        filename = Dir
    Loop


    Sheets("import").Select

    'clear original data
    Rows("30:30").Select
    ActiveWindow.SmallScroll Down:=138
    Rows("30:186").Select
    Selection.Delete Shift:=xlUp

    ActiveSheet.Range("A30") = "include the calibration"
    ActiveSheet.Range("A30").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    ActiveSheet.Range("b31").Resize(1, UBound(fre) + 1).Value = fre

    Sheets("front").Select
    If l_index_f <> 3 Then

        ActiveSheet.Range("a2", ActiveSheet.Range("a2").End(xlDown)).Select
        Selection.Copy
        Sheets("import").Select
        ActiveSheet.Range("a31").Select
        ActiveSheet.Paste
        ActiveSheet.Range("a31") = "Front_Hz"
        ActiveSheet.Range("b" & CStr(30 + l_index_f)).Resize(1, UBound(fre) + 1).Value = fre
    End If


    If l_index_r <> 3 Then
        Sheets("rear").Select
        ActiveSheet.Range("a2", ActiveSheet.Range("a2").End(xlDown)).Select
        Selection.Copy
        Sheets("import").Select
        ActiveSheet.Range("a" & CStr(30 + l_index_f)).Select
        ActiveSheet.Paste
        ActiveSheet.Range("a" & CStr(30 + l_index_f)) = "Rear_Hz"
    End If



    'include the calibration
     If l_index_f <> 3 Then
        For ii = 3 To l_index_f - 1
            For jj = 3 To UBound(msg_dfs_off_f) + 3
                ActiveSheet.Cells(29 + ii, jj - 1) = Sheets("front").Cells(ii, jj - 1) - Sheets("calibration").Cells(11, jj)
            Next


        Next
    End If

    If l_index_r <> 3 Then

'        Sheets("rear").Select
        For ii = 3 To l_index_r - 1
            For jj = 3 To UBound(msg_dfs_off_f) + 3
                ActiveSheet.Cells(28 + l_index_f + ii, jj - 1) = Sheets("rear").Cells(ii, jj - 1) - Sheets("calibration").Cells(11, jj)
            Next


        Next
    End If

    Application.CutCopyMode = False

        '------------- plot the output:msg curves ------------
    '    For ii = 1 To ActiveSheet.ChartObjects.Count
    '        ActiveSheet.ChartObjects.Item(1).Activate
    '        ActiveChart.Parent.Delete
    '    Next
    '
    If l_index_f <> 3 Then
        Sheets("import").Select
        ActiveSheet.Range("A31").Select
        ActiveSheet.Range(Selection, Selection.End(xlDown)).Select
        ActiveSheet.Range(Selection, Selection.End(xlToRight)).Select
        ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
    '    ActiveChart.Parent.Cut

        ActiveChart.ChartArea.Select
        ActiveChart.Location Where:=xlLocationAsObject, Name:="output"

    '    Sheets("output").Select
    '    ActiveSheet.Paste

        ActiveSheet.ChartObjects.Item(1).Activate
        ActiveSheet.Shapes.Item(1).IncrementLeft -418.5
        ActiveSheet.Shapes.Item(1).IncrementTop -5000
        ActiveSheet.Shapes.Item(1).ScaleWidth 1.8666666667, msoFalse, _
            msoScaleFromTopLeft
        ActiveSheet.Shapes.Item(1).ScaleHeight 2.5833333333, msoFalse, _
            msoScaleFromTopLeft
        ActiveWindow.SmallScroll Down:=-3
        ActiveSheet.ChartObjects.Item(1).Activate
        ActiveChart.ChartTitle.Select
        ActiveChart.ChartTitle.Text = "MSG_front"
        Selection.Format.TextFrame2.TextRange.Characters.Text = "MSG_front"
        With Selection.Format.TextFrame2.TextRange.Characters(1, 3).ParagraphFormat
            .TextDirection = msoTextDirectionLeftToRight
            .Alignment = msoAlignCenter
        End With
        With Selection.Format.TextFrame2.TextRange.Characters(1, 3).Font
            .BaselineOffset = 0
            .Bold = msoFalse
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(89, 89, 89)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 14
            .Italic = msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = msoNoUnderline
            .Spacing = 0
            .Strike = msoNoStrike
        End With
        ActiveChart.Axes(xlValue).Select
        ActiveChart.Axes(xlValue).MinimumScale = 0
        ActiveChart.Axes(xlValue).MaximumScale = 150
        Application.CommandBars("Format Object").Visible = False

        ActiveSheet.ChartObjects.Item(1).Activate
        ActiveChart.PlotArea.Select
        'ActiveChart.FullSeriesCollection(18).Select
        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.FullSeriesCollection.Item(l_index_f - 2).Name = "=import!$A$8"
        If Sheets("import").Range("f16") = "Dooku" Then
            ActiveChart.FullSeriesCollection(l_index_f - 2).XValues = "=import!$B$8:$K$8"
            ActiveChart.FullSeriesCollection(l_index_f - 2).Values = "=import!$B$9:$K$9"
        Else
            ActiveChart.FullSeriesCollection(l_index_f - 2).XValues = "=import!$B$8:$j$8"
            ActiveChart.FullSeriesCollection(l_index_f - 2).Values = "=import!$B$9:$j$9"

        End If


        ActiveChart.FullSeriesCollection(l_index_f - 2).Select
        With Selection.Format.line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
        With Selection.Format.line
            .Visible = msoTrue
            .Weight = 2
        End With
        With Selection.Format.line
            .Visible = msoTrue
            .DashStyle = msoLineDash
        End With

        ActiveChart.Axes(xlCategory).Select
        ActiveChart.Axes(xlCategory).HasMinorGridlines = True
        ActiveChart.Axes(xlCategory).ScaleType = xlLogarithmic
        ActiveChart.Axes(xlCategory).MinimumScale = 100
        ActiveChart.Axes(xlCategory).MaximumScale = 10000
        Application.CommandBars("Format Object").Visible = False
    End If

'    if there is rear curve, plot it
    If l_index_r <> 3 Then

        Sheets("import").Select
        ActiveSheet.Range("A" & CStr(l_index_f + 30)).Select
        ActiveSheet.Range(Selection, Selection.End(xlDown)).Select
        ActiveSheet.Range(Selection, Selection.End(xlToRight)).Select
        ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
'        ActiveChart.Parent.Cut

        ActiveChart.ChartArea.Select
        ActiveChart.Location Where:=xlLocationAsObject, Name:="output"

'        Sheets("output").Select
'        ActiveSheet.Range("o2").Select
'        ActiveSheet.Paste

        ActiveSheet.ChartObjects.Item(2).Activate

        ActiveSheet.Shapes.Item(2).IncrementLeft 254.25
        ActiveSheet.Shapes.Item(2).IncrementTop -5000
        ActiveSheet.Shapes.Item(2).ScaleWidth 1.8666666667, msoFalse, _
            msoScaleFromTopLeft
        ActiveSheet.Shapes.Item(2).ScaleHeight 2.5833333333, msoFalse, _
            msoScaleFromTopLeft
        ActiveWindow.SmallScroll Down:=-3
        ActiveSheet.ChartObjects.Item(2).Activate
        ActiveChart.ChartTitle.Select
        ActiveChart.ChartTitle.Text = "MSG_rear"
        Selection.Format.TextFrame2.TextRange.Characters.Text = "MSG_rear"
        With Selection.Format.TextFrame2.TextRange.Characters(1, 3).ParagraphFormat
            .TextDirection = msoTextDirectionLeftToRight
            .Alignment = msoAlignCenter
        End With
        With Selection.Format.TextFrame2.TextRange.Characters(1, 3).Font
            .BaselineOffset = 0
            .Bold = msoFalse
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = RGB(89, 89, 89)
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 14
            .Italic = msoFalse
            .Kerning = 12
            .Name = "+mn-lt"
            .UnderlineStyle = msoNoUnderline
            .Spacing = 0
            .Strike = msoNoStrike
        End With
        ActiveChart.Axes(xlValue).Select
        ActiveChart.Axes(xlValue).MinimumScale = 0
        ActiveChart.Axes(xlValue).MaximumScale = 150
        Application.CommandBars("Format Object").Visible = False

        ActiveSheet.ChartObjects.Item(2).Activate
        ActiveChart.PlotArea.Select
        'ActiveChart.FullSeriesCollection(18).Select
        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.FullSeriesCollection.Item(l_index_r - 2).Name = "=import!$A$8"
        If Sheets("import").Range("f16") = "Dooku" Then
            ActiveChart.FullSeriesCollection(l_index_r - 2).XValues = "=import!$B$8:$K$8"
            ActiveChart.FullSeriesCollection(l_index_r - 2).Values = "=import!$B$9:$K$9"
        Else
            ActiveChart.FullSeriesCollection(l_index_r - 2).XValues = "=import!$B$8:$j$8"
            ActiveChart.FullSeriesCollection(l_index_r - 2).Values = "=import!$B$9:$j$9"

        End If


        ActiveChart.FullSeriesCollection(l_index_r - 2).Select
        With Selection.Format.line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
        With Selection.Format.line
            .Visible = msoTrue
            .Weight = 2
        End With
        With Selection.Format.line
            .Visible = msoTrue
            .DashStyle = msoLineDash
        End With

        ActiveChart.Axes(xlCategory).Select
        ActiveChart.Axes(xlCategory).HasMinorGridlines = True
        ActiveChart.Axes(xlCategory).ScaleType = xlLogarithmic
        ActiveChart.Axes(xlCategory).MinimumScale = 100
        ActiveChart.Axes(xlCategory).MaximumScale = 10000
        Application.CommandBars("Format Object").Visible = False

    End If


    Sheets("calibration").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("front").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("rear").Select
    ActiveWindow.SelectedSheets.Visible = False

    Sheets("output").Select


End Sub

Public Sub interpolate()


    Dim x As Double, y As Double, a As Integer, maxbase As Integer, i As Integer, valgt As Variant, targetfrekarray(1 To 1000) As Double, frekvensarray(1 To 1000) As Double, talarray(1 To 1000) As Double, cellearray(1 To 5) As Boolean, trange As Range


    Sheets("calibration").Select
    ActiveSheet.Range("c7", ActiveSheet.Range("c7").End(xlToRight)).Select

    For Each valgt In Selection
      i = i + 1
        frekvensarray(i) = valgt
        maxbase = i
    Next


    i = 0
    ActiveSheet.Range("c8", ActiveSheet.Range("c8").End(xlToRight)).Select
    For Each valgt In Selection
      i = i + 1
        talarray(i) = valgt
    Next


    i = 0
    ActiveSheet.Range("c10", ActiveSheet.Range("c10").End(xlToRight)).Select
    For Each valgt In Selection
      i = i + 1
        targetfrekarray(i) = valgt
    Next

    a = 0
    ActiveSheet.Range("a1").Select
    ActiveCell.Offset(1, 0).Range("c10", ActiveSheet.Range("c10").End(xlToRight)).Select
    For Each valgt In Selection
      a = a + 1
      i = 1
        While targetfrekarray(a) > frekvensarray(i) And Not i = maxbase
          i = i + 1
        Wend
      If i <> 1 Then
        i = i - 1
      End If
      valgt.Value = talarray(i) + (talarray(i + 1) - talarray(i)) * (targetfrekarray(a) - frekvensarray(i)) / (frekvensarray(i + 1) - frekvensarray(i))
    Next


    'Selection.Copy
    'Range("c5").Select
    'Windows("personal.xls").Visible = False


End Sub
