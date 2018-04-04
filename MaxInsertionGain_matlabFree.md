Private Sub CommandButton1_Click()

    Dim path
    Dim filename
    Dim line_f As Integer, row_no As Integer, devices1 As Integer, devices2 As Integer, bound As Integer, chart_number As Integer
    Dim line_r As Integer, ii As Integer, jj As Integer, kk As Integer, dd As Integer
    Dim fre(127) As Double
    Dim l_index_f As Integer
    Dim l_index_r As Integer, devices_index() As Integer, prod_offset() As Double, temp_offset As Double, devices_f_column() As Integer, devices_r_column() As Integer
    Dim dual_or_not As Boolean, devices_f() As String, devices_r() As String, devices_f_row() As Integer, devices_r_row() As Integer

    Sheets("calibration").Visible = True
    Sheets("front").Visible = True
    Sheets("rear").Visible = True

    devices1 = 0
    devices2 = 0
    bound = 0
    row_no = 0



    'interpolate the calibration value
    Call interpolate
    Call interpolate_IG(Sheets("import").Range("f16"))



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

    'make the header in the sheets of front & rear
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

        'find the front and rear line location
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
            ActiveSheet.Range("a" & CStr(l_index_f)) = "Front_" & curvename
            If dual_or_not Then
                Sheets("rear").Select
                ActiveSheet.Range("a" & CStr(l_index_r)) = "Rear_" & curvename
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


    'get the Margin = FOIG - MSIG
    Sheets("cover").Select
    ActiveSheet.Range("E11:F15").Select
    Selection.ClearContents
    ActiveSheet.Range("J11:K15").Select
    Selection.ClearContents


    'device number stored in devices_f & devices_r, now the max devices number is 10
    ReDim devices_f(10)
    ReDim devices_r(10)
    ReDim devices_f_row(10)
    ReDim devices_f_column(10)
    ReDim devices_r_row(10)
    ReDim devices_r_column(10)
    ReDim devices_index(10)

    ii = 11
    Do While Sheets("cover").Cells(ii, 4) <> ""
        ii = ii + 1
    Loop
    devices1 = ii - 11

    ii = 11
    Do While Sheets("cover").Cells(ii, 9) <> ""
        ii = ii + 1
    Loop
    devices2 = ii - 11
    bound = devices1 + devices2

    If bound <> 0 Then
        ReDim prod_offset(bound - 1)
        For jj = 0 To devices1 - 1
            prod_offset(jj) = Sheets("cover").Cells(11 + jj, 3)
        Next

        For jj = 0 To devices2 - 1
            prod_offset(jj + devices1) = Sheets("cover").Cells(11 + jj, 8)
        Next


        Sheets("front").Select
        dd = 0
        For kk = 3 To l_index_f
            For jj = 1 To bound
                If InStr(ActiveSheet.Range("a" & CStr(kk)), Sheets("cover").Cells(jj + 10, 4)) And Sheets("cover").Cells(jj + 10, 4) <> "" Then
                    devices_f(dd) = Sheets("cover").Cells(jj + 10, 4)
                    devices_f_row(dd) = jj + 10
                    devices_f_column(dd) = 4
                    dd = dd + 1
                    GoTo nextdevice_f
                ElseIf InStr(ActiveSheet.Range("a" & CStr(kk)), Sheets("cover").Cells(jj + 10, 9)) And Sheets("cover").Cells(jj + 10, 9) <> "" Then
                    devices_f(dd) = Sheets("cover").Cells(jj + 10, 9)
                    devices_f_row(dd) = jj + 10
                    devices_f_column(dd) = 9
                    dd = dd + 1
                    GoTo nextdevice_f
                End If
            Next


nextdevice_f:
        Next

        If dd = 0 Then
            MsgBox "no devices data in cover found"
            End
        End If
        ReDim Preserve devices_f(dd - 1)
        ReDim Preserve devices_f_row(dd - 1)
        ReDim Preserve devices_f_column(dd - 1)

        Sheets("rear").Select
        dd = 0
        For kk = 3 To l_index_r
            For jj = 1 To bound
                If InStr(ActiveSheet.Range("a" & CStr(kk)), Sheets("cover").Cells(jj + 10, 4)) And Sheets("cover").Cells(jj + 10, 4) <> "" Then
                    devices_r(dd) = Sheets("cover").Cells(jj + 10, 4)
                    devices_r_row(dd) = jj + 10
                    devices_r_column(dd) = 4
                    dd = dd + 1
                    GoTo nextdevice_r
                ElseIf InStr(ActiveSheet.Range("a" & CStr(kk)), Sheets("cover").Cells(jj + 10, 9)) And Sheets("cover").Cells(jj + 10, 9) <> "" Then
                    devices_r(dd) = Sheets("cover").Cells(jj + 10, 9)
                    devices_r_row(dd) = jj + 10
                    devices_r_column(dd) = 9
                    dd = dd + 1
                    GoTo nextdevice_r
                End If
            Next

nextdevice_r:
        Next
        If dd = 0 Then
            ReDim devices_r(0)
            ReDim devices_r_row(0)
            ReDim devices_r_column(0)
        Else
            ReDim Preserve devices_r(dd - 1)
            ReDim Preserve devices_r_row(dd - 1)
            ReDim Preserve devices_r_column(dd - 1)
        End If

    Else
        ReDim prod_offset(0)
    End If


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

        'include the calibration
        temp_offset = 0
        For ii = 3 To l_index_f - 1
            For dd = 0 To bound - 1
                If InStr(ActiveSheet.Cells(29 + ii, 1), devices_f(dd)) Then
                    temp_offset = prod_offset(dd)
                    GoTo nextline:
                End If
            Next
nextline:
            For jj = 3 To UBound(msg_dfs_off_f) + 3
                ActiveSheet.Cells(29 + ii, jj - 1) = Sheets("front").Cells(ii, jj - 1) - Sheets("calibration").Cells(11, jj) - temp_offset
            Next
        Next


        'get the max of the margin
        ActiveSheet.Cells(30 + l_index_f + l_index_r, 1) = "Hz"
        For jj = 2 To UBound(msg_dfs_off_f) + 2
            ActiveSheet.Cells(30 + l_index_f + l_index_r, jj) = ActiveSheet.Cells(30 + l_index_f, jj)
        Next
        ActiveSheet.Cells(31 + l_index_f + l_index_r, 1) = Sheets("calibration").Cells(12, 2) & "_plusMargin"
        ActiveSheet.Cells(32 + l_index_f + l_index_r, 1) = "MarginLimit"
        For jj = 2 To UBound(msg_dfs_off_f) + 2
            ActiveSheet.Cells(31 + l_index_f + l_index_r, jj) = Sheets("calibration").Cells(12, jj + 1) - ActiveSheet.Range("j16")
            ActiveSheet.Cells(32 + l_index_f + l_index_r, jj) = ActiveSheet.Range("j16")
        Next
        ActiveSheet.Range("A" & CStr(34 + l_index_f + l_index_r)) = "Get the margin"
        ActiveSheet.Range("A" & CStr(34 + l_index_f + l_index_r)).Select
        Selection.Font.Bold = True
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        ActiveSheet.Cells(35 + l_index_f + l_index_r, 1) = "Hz"
        For jj = 2 To UBound(msg_dfs_off_f) + 2
            ActiveSheet.Cells(35 + l_index_f + l_index_r, jj) = ActiveSheet.Cells(31, jj)
        Next
        ActiveSheet.Cells(35 + l_index_f + l_index_r, jj + 1) = "Max"


        For ii = 3 To l_index_f - 1
            ActiveSheet.Cells(33 + l_index_f + l_index_r + ii, 1) = ActiveSheet.Cells(29 + ii, 1) & "_Margin"

            For jj = 2 To UBound(msg_dfs_off_f) + 2
                ActiveSheet.Cells(33 + l_index_f + l_index_r + ii, jj) = ActiveSheet.Cells(31 + l_index_f + l_index_r, jj) - _
                ActiveSheet.Cells(29 + ii, jj) + ActiveSheet.Range("j16")
            Next
            ActiveSheet.Cells(33 + l_index_f + l_index_r + ii, jj + 1).FormulaR1C1 = "=MAX(RC[-" & CStr(jj - 1) & "]:RC[-2])"
        Next


        For kk = 0 To UBound(devices_f)

            ActiveSheet.Range("a" & CStr(l_index_f + l_index_r + 35), ActiveSheet.Range("a" & CStr(l_index_f + l_index_r + 35)).End(xlDown)).Select
            Selection.Find(What:="front*" & devices_f(kk), After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False).Activate
            row_no = ActiveCell.Row

            ' count how many curves in one device name
            ii = row_no
            jj = 0
            Do While ActiveSheet.Range("a" & CStr(ii)) <> ""
                If InStr(ActiveSheet.Range("a" & CStr(ii)), devices_f(kk)) Then
                    devices_index(jj) = ii
                    jj = jj + 1
                End If
                ii = ii + 1

            Loop

            If jj <> 0 Then
                Sheets("import").Cells(row_no, UBound(msg_dfs_off_f) + 5).Select
                path = "R[0]C[-1]"
                For ii = 2 To jj
                    path = path & ",R[" & CStr(devices_index(ii - 1) - row_no) & "]C[-1]"
                Next
                ActiveCell.FormulaR1C1 = "=MAX(" & path & ")"
                Sheets("cover").Cells(devices_f_row(kk), devices_f_column(kk) + 1) = ActiveCell
            Else
                MsgBox devices_f(kk) & " data no found"
                End
            End If

        Next

    End If

    If l_index_r <> 3 Then
        Sheets("rear").Select
        ActiveSheet.Range("a2", ActiveSheet.Range("a2").End(xlDown)).Select
        Selection.Copy
        Sheets("import").Select
        ActiveSheet.Range("a" & CStr(30 + l_index_f)).Select
        ActiveSheet.Paste
        ActiveSheet.Range("a" & CStr(30 + l_index_f)) = "Rear_Hz"

        'include the calibration
        temp_offset = 0
        For ii = 3 To l_index_r - 1
            For dd = 0 To bound - 1
                If InStr(ActiveSheet.Cells(28 + l_index_f + ii, 1), devices_r(dd)) Then
                    temp_offset = prod_offset(dd)
                    GoTo nextline_1:
                End If
            Next
nextline_1:
            For jj = 3 To UBound(msg_dfs_off_f) + 3
                ActiveSheet.Cells(28 + l_index_f + ii, jj - 1) = Sheets("rear").Cells(ii, jj - 1) - Sheets("calibration").Cells(11, jj) - temp_offset
            Next
        Next



        'get the max of the margin
        For ii = 3 To l_index_r - 1
            ActiveSheet.Cells(30 + 2 * l_index_f + l_index_r + ii, 1) = ActiveSheet.Cells(28 + l_index_f + ii, 1) & "_Margin"

            For jj = 2 To UBound(msg_dfs_off_f) + 2
                ActiveSheet.Cells(30 + 2 * l_index_f + l_index_r + ii, jj) = ActiveSheet.Cells(31 + l_index_f + l_index_r, jj) - _
                ActiveSheet.Cells(28 + l_index_f + ii, jj) + ActiveSheet.Range("j16")
            Next
            ActiveSheet.Cells(30 + 2 * l_index_f + l_index_r + ii, jj + 1).FormulaR1C1 = "=MAX(RC[-" & CStr(jj - 1) & "]:RC[-2])"

        Next


        For kk = 0 To UBound(devices_r)

            If l_index_r > 4 Then
                ActiveSheet.Range("a" & CStr(2 * l_index_f + l_index_r + 32), ActiveSheet.Range("a" & CStr(2 * l_index_f + l_index_r + 32)).End(xlDown)).Select
            ElseIf l_index_r = 4 Then
                ActiveSheet.Range("a" & CStr(2 * l_index_f + l_index_r + 33)).Select
            End If

            Selection.Find(What:="rear*" & devices_r(kk), After:=ActiveCell, LookIn:=xlFormulas, _
                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False).Activate
            row_no = ActiveCell.Row


            ' count how many curves in one device name
            ii = row_no
            jj = 0
            Do While ActiveSheet.Range("a" & CStr(ii)) <> ""
                If InStr(ActiveSheet.Range("a" & CStr(ii)), devices_r(kk)) Then
                    devices_index(jj) = ii
                    jj = jj + 1
                End If
                ii = ii + 1

            Loop

            If jj <> 0 Then
                Sheets("import").Cells(row_no, UBound(msg_dfs_off_r) + 5).Select
                path = "R[0]C[-1]"
                For ii = 2 To jj
                    path = path & ",R[" & CStr(devices_index(ii - 1) - row_no) & "]C[-1]"
                Next
                ActiveCell.FormulaR1C1 = "=MAX(" & path & ")"

                Sheets("cover").Cells(devices_r_row(kk), devices_r_column(kk) + 2) = ActiveCell
            Else
                MsgBox devices_f(kk) & " data no found"
                End
            End If

        Next

    End If


    Application.CutCopyMode = False



    '------------- plot the output:msg curves ------------
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
        ActiveChart.ChartTitle.Text = "MSIG_front"
        Selection.Format.TextFrame2.TextRange.Characters.Text = "MSIG_front"
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
        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.FullSeriesCollection.Item(l_index_f - 2).Name = "=import!$A$" & CStr(31 + l_index_f + l_index_r)
        ActiveChart.FullSeriesCollection(l_index_f - 2).XValues = "=import!$B$" & CStr(30 + l_index_f + l_index_r) & ":$DY$" & CStr(30 + l_index_f + l_index_r)
        ActiveChart.FullSeriesCollection(l_index_f - 2).Values = "=import!$B$" & CStr(31 + l_index_f + l_index_r) & ":$DY$" & CStr(31 + l_index_f + l_index_r)


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
        ActiveChart.ChartTitle.Text = "MSIG_rear"
        Selection.Format.TextFrame2.TextRange.Characters.Text = "MSIG_rear"
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
        ActiveChart.FullSeriesCollection.Item(l_index_r - 2).Name = "=import!$A$" & CStr(31 + l_index_f + l_index_r)
        ActiveChart.FullSeriesCollection(l_index_r - 2).XValues = "=import!$B$" & CStr(30 + l_index_f + l_index_r) & ":$DY$" & CStr(30 + l_index_f + l_index_r)
        ActiveChart.FullSeriesCollection(l_index_r - 2).Values = "=import!$B$" & CStr(31 + l_index_f + l_index_r) & ":$DY$" & CStr(31 + l_index_f + l_index_r)


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

    If l_index_f <> 3 Or l_index_r <> 3 Then

        If l_index_r = 3 Then
            chart_number = 2
        Else
            chart_number = 3
        End If

        Sheets("import").Select
        ActiveSheet.Range("A" & CStr(l_index_f + l_index_r + 35)).Select
        ActiveSheet.Range(Selection, Selection.End(xlDown)).Select
        ActiveSheet.Range(Selection, Selection.End(xlToRight)).Select
        ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
    '    ActiveChart.Parent.Cut

        ActiveChart.ChartArea.Select
        ActiveChart.Location Where:=xlLocationAsObject, Name:="output"

    '    Sheets("output").Select
    '    ActiveSheet.Paste

        ActiveSheet.ChartObjects.Item(chart_number).Activate


        ActiveWindow.SmallScroll Down:=24
        ActiveChart.ChartArea.Select
        ActiveSheet.Shapes.Item(chart_number).IncrementLeft -418.5
        ActiveSheet.Shapes.Item(chart_number).IncrementTop -5000
        ActiveSheet.Shapes.Item(chart_number).IncrementTop 600
        ActiveWindow.SmallScroll Down:=12
        ActiveSheet.Shapes.Item(chart_number).ScaleWidth 2.0020833333, msoFalse, _
            msoScaleFromTopLeft
        ActiveSheet.Shapes.Item(chart_number).ScaleHeight 2.3854166667, msoFalse, _
            msoScaleFromTopLeft

        ActiveChart.ChartTitle.Select
        ActiveChart.ChartTitle.Text = "Margin"
        Selection.Format.TextFrame2.TextRange.Characters.Text = "Margin"
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
        ActiveChart.Axes(xlValue).MinimumScale = -50
        ActiveChart.Axes(xlValue).MaximumScale = 0
        Application.CommandBars("Format Object").Visible = False

        ActiveSheet.ChartObjects.Item(chart_number).Activate
        ActiveChart.PlotArea.Select
        'ActiveChart.FullSeriesCollection(18).Select
        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.FullSeriesCollection.Item(l_index_r + l_index_f - 5).Name = "=import!$A$" & CStr(32 + l_index_f + l_index_r)
        ActiveChart.FullSeriesCollection(l_index_r + l_index_f - 5).XValues = "=import!$B$" & CStr(30 + l_index_f + l_index_r) & ":$DY$" & CStr(30 + l_index_f + l_index_r)
        ActiveChart.FullSeriesCollection(l_index_r + l_index_f - 5).Values = "=import!$B$" & CStr(32 + l_index_f + l_index_r) & ":$DY$" & CStr(32 + l_index_f + l_index_r)
'        If Sheets("import").Range("f16") = "Dooku" Then
'            ActiveChart.FullSeriesCollection(l_index_f - 2).XValues = "=import!$B$8:$K$8"
'            ActiveChart.FullSeriesCollection(l_index_f - 2).Values = "=import!$B$9:$K$9"
'        Else
'            ActiveChart.FullSeriesCollection(l_index_f - 2).XValues = "=import!$B$8:$j$8"
'            ActiveChart.FullSeriesCollection(l_index_f - 2).Values = "=import!$B$9:$j$9"
'
'        End If


        ActiveChart.FullSeriesCollection(l_index_r + l_index_f - 5).Select
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
    Sheets("import").Select
    ActiveSheet.Range("a1").Select

    Sheets("COVER").Select
    ActiveSheet.Range("a1").Select


End Sub

Public Sub interpolate()


    Dim x As Double, y As Double, a As Integer, maxbase As Integer, i As Integer, valgt As Variant, targetfrekarray(1 To 1000) As Double, frekvensarray(1 To 1000) As Double, talarray(1 To 1000) As Double, cellearray(1 To 5) As Boolean, trange As Range


    Sheets("calibration").Select
    i = 0
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
    ActiveCell.offset(1, 0).Range("c10", ActiveSheet.Range("c10").End(xlToRight)).Select
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

Public Sub interpolate_IG(firmware As String)


    Dim x As Double, y As Double, a As Integer, maxbase As Integer, i As Integer, valgt As Variant, targetfrekarray(1 To 1000) As Double, frekvensarray(1 To 1000) As Double, talarray(1 To 1000) As Double, cellearray(1 To 5) As Boolean, trange As Range


    Sheets("import").Select

    If firmware = "Dooku" Then
        ActiveSheet.Range("b8", ActiveSheet.Range("b8").End(xlToRight)).Select
    Else
        ActiveSheet.Range("b8:j8").Select
    End If

    For Each valgt In Selection
      i = i + 1
        frekvensarray(i) = valgt
        maxbase = i
    Next


    i = 0
    If firmware = "Dooku" Then
        ActiveSheet.Range("b9", ActiveSheet.Range("b9").End(xlToRight)).Select
    Else
        ActiveSheet.Range("b9:j9").Select
    End If

    For Each valgt In Selection
      i = i + 1
        talarray(i) = valgt
    Next


    i = 0
    Sheets("calibration").Select
    ActiveSheet.Range("c10", ActiveSheet.Range("c10").End(xlToRight)).Select
    For Each valgt In Selection
      i = i + 1
        targetfrekarray(i) = valgt
    Next

    a = 0
    ActiveSheet.Range("a1").Select
    ActiveCell.offset(2, 0).Range("c10", ActiveSheet.Range("c10").End(xlToRight)).Select
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



Private Sub CommandButton2_Click()


    Sheets("import").Select
    Sheets("MSG_OPL").Visible = True
    Sheets("MSG_OPL").Select
    Sheets("matlab vs this script").Visible = True


End Sub
