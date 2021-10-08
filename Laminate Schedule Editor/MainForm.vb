Imports Excel = Microsoft.Office.Interop.Excel

Public Class MainForm
    Dim xlWorkBook As Excel.Workbook
    Dim Excel As Object = Nothing

    Public CalcState As Long
    Public EventState As Boolean
    Public PageBreakState As Boolean


    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim vers As Version = My.Application.Info.Version
        Me.Text = "Laminate Schedule Update " & vers.Major & "." & vers.Minor & "." & vers.Build

    End Sub

    Private Sub Btn_Browse_Click(sender As Object, e As EventArgs) Handles Btn_Browse.Click
        If OpenFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            FilePath.Text = OpenFileDialog.FileName
        End If
    End Sub

    Private Sub Btn_Open_Click(sender As Object, e As EventArgs) Handles Btn_Open.Click
        If checkFilePath() Then
            Excel = CreateObject("Excel.Application")
            xlWorkBook = Excel.Workbooks.Open(FilePath.Text)

            If xlWorkBook.ReadOnly Then
                MsgBox("File is Read-Only, please correct this and re-open the file.",, "Error")
                Excel.Quit
                Exit Sub
            End If

            If xlWorkBook.Worksheets.Item(1).Name <> "Laminate Schedule" Or xlWorkBook.Worksheets.Item(2).Name <> "CATIA PlyBook" Or xlWorkBook.Worksheets.Item(3).Name <> "Standard" Then
                MsgBox("Tabs names do not align with Laminate Schedule Template, please make sure this is a laminate schedule.",, "Error")
                Excel.Quit
                Exit Sub
            End If

            Excel.Visible = True
        End If
    End Sub

    Private Sub MainForm_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Try
            Excel.Quit
        Catch ex As Exception
        End Try
    End Sub

    Function checkFilePath() As Boolean
        If System.IO.File.Exists(FilePath.Text) Then
            If System.IO.Path.GetExtension(FilePath.Text) = ".xls" Or System.IO.Path.GetExtension(FilePath.Text) = ".xlsx" Or System.IO.Path.GetExtension(FilePath.Text) = ".xlsm" Then
                Return True
            End If
        End If

        Return False
    End Function

    Function checkWorksheet() As Boolean
        If xlWorkBook Is Nothing Then
            MsgBox("Excel workbook not found.",, "Error")
            Excel.Quit
            Return True
        End If

        If xlWorkBook.Worksheets.Item(1).Name <> "Laminate Schedule" Or xlWorkBook.Worksheets.Item(2).Name <> "CATIA PlyBook" Or xlWorkBook.Worksheets.Item(3).Name <> "Standard" Then
            MsgBox("Tabs names do not align with Laminate Schedule Template, please make sure this is a laminate schedule.",, "Error")
            Excel.Quit
            Return True
        End If

        Return False
    End Function

    Private Sub Txt_DebulkConst_TextChanged(sender As Object, e As EventArgs) Handles Txt_DebulkConst.TextChanged
        If IsNumeric(Txt_DebulkConst.Text) Then
            Txt_DebulkConst.Text = CInt(Txt_DebulkConst.Text)
            Txt_DebulkConst.BackColor = Color.White
        Else
            Txt_DebulkConst.BackColor = Color.Red
        End If
    End Sub

    Private Sub Btn_wrkUpdate_Click(sender As Object, e As EventArgs) Handles Btn_wrkUpdate.Click
        If checkWorksheet() Then
            Exit Sub
        End If

        Call wrkshtUpdate()
    End Sub

    Private Sub Btn_PlyStdCreate_Click(sender As Object, e As EventArgs) Handles Btn_PlyStdCreate.Click
        If checkWorksheet() Then
            Exit Sub
        End If

        If Txt_DebulkConst.BackColor = Color.Red Then
            MsgBox("Make sure 'Debulk Constant' is a valid number.",, "Error")
            Exit Sub
        End If

        Call wrkshtUpdate()
        Call plyStandardCreate()
    End Sub

    Private Sub Btn_ReRunVals_Click(sender As Object, e As EventArgs) Handles Btn_ReRunVals.Click
        If checkWorksheet() Then
            Exit Sub
        End If

        Call wrkshtUpdate()
        Call addValues()
    End Sub

    Sub wrkshtUpdate()

        xlWorkBook.Worksheets.Item(1).Activate
        xlWorkBook.ActiveSheet.PageSetup.DifferentFirstPageHeaderFooter = True

        Dim docNum As String = xlWorkBook.ActiveSheet.Cells(3, 4).Value
        Dim docRev As String = xlWorkBook.ActiveSheet.Cells(3, 7).Value
        Dim docTitle As String = xlWorkBook.ActiveSheet.Cells(4, 4).Value
        Dim customerName As String = xlWorkBook.ActiveSheet.Cells(6, 4).Value
        Dim prodNum As String = xlWorkBook.ActiveSheet.Cells(8, 4).Value
        Dim prodNomenclature As String = xlWorkBook.ActiveSheet.Cells(9, 4).Value

        Dim leftFooter As String
        leftFooter = "&12&""Calibri""&B" & "Doc. No. " & docNum & "_" & docRev

        xlWorkBook.ActiveSheet.PageSetup.leftFooter = leftFooter
        xlWorkBook.ActiveSheet.PageSetup.FirstPage.leftFooter.Text = leftFooter

        Dim rightHeader As String
        rightHeader = "&18&""Calibri""&B" & vbCr &
            docNum & "_" & docRev & " | " & docTitle & vbCr &
            customerName & vbCr &
            prodNum & " | " & prodNomenclature

        xlWorkBook.ActiveSheet.PageSetup.rightHeader = rightHeader
        xlWorkBook.ActiveSheet.PageSetup.FirstPage.rightHeader.Text = rightHeader

    End Sub

    Sub plyStandardCreate()

        xlWorkBook.Worksheets.Item(1).Activate
        Dim debulkConst As Integer
        debulkConst = CInt(Txt_DebulkConst.Text)

        Dim i As Integer
        i = 1
        Dim keyLine As Integer
        Do
            If CStr(xlWorkBook.ActiveSheet.Cells(i, 1).Value) = "KEY" Then
                keyLine = i
            End If
            i = i + 1
        Loop Until keyLine > 0

        xlWorkBook.ActiveSheet.Range("A" & i & ":A9999").Clear

        Dim currentLine As Integer
        currentLine = keyLine + 1
        xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "PREP"
        currentLine = currentLine + 1
        xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "PLYHEAD"
        currentLine = currentLine + 1

        Dim debulkRate As Integer
        debulkRate = 0
        Dim plyBookNum As Integer
        plyBookNum = 1
        Dim firstDebulk As Boolean
        firstDebulk = False

        Dim sequenceName As String = xlWorkBook.Sheets.Item(2).Cells(plyBookNum, 2).Value

        Do
            plyBookNum = plyBookNum + 1
            If xlWorkBook.Sheets.Item(2).Cells(plyBookNum, 2).Value <> sequenceName Then
                debulkRate = debulkRate + 1
                sequenceName = xlWorkBook.Sheets.Item(2).Cells(plyBookNum, 2).Value

                If firstDebulk = False And debulkRate = 2 Then
                    xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "TECH"
                    currentLine = currentLine + 1
                    xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "BULK"
                    currentLine = currentLine + 1
                    xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "SECONDARY"
                    currentLine = currentLine + 1
                    xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "BLANK"
                    currentLine = currentLine + 1
                    xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "PLYHEAD"
                    currentLine = currentLine + 1
                    firstDebulk = True
                    debulkRate = debulkRate - 1
                End If
            End If

            If debulkRate = debulkConst + 1 Then
                xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "TECH"
                currentLine = currentLine + 1
                xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "BULK"
                currentLine = currentLine + 1
                xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "SECONDARY"
                currentLine = currentLine + 1
                xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "BLANK"
                currentLine = currentLine + 1
                xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "PLYHEAD"
                currentLine = currentLine + 1
                debulkRate = 1
            End If

            xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = xlWorkBook.Sheets.Item(2).Cells(plyBookNum, 1).Value
            currentLine = currentLine + 1

        Loop Until CStr(xlWorkBook.Sheets.Item(2).Cells(plyBookNum, 1).Value) = ""

        currentLine = currentLine - 1
        xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "TC"
        currentLine = currentLine + 1
        xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "FB"
        currentLine = currentLine + 1
        xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "LEAK"
        currentLine = currentLine + 1
        xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "BLANK"
        currentLine = currentLine + 1
        xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "LEAK-END"
        currentLine = currentLine + 1
        xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "BLANK"
        currentLine = currentLine + 1
        xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "LABEL"
        currentLine = currentLine + 1
        xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "SECONDARY"
        currentLine = currentLine + 1
        xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "BLANK"
        currentLine = currentLine + 1
        xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "QUALITY"
        currentLine = currentLine + 1
        xlWorkBook.ActiveSheet.Cells(currentLine, 1).Value = "BLANK"
        currentLine = currentLine + 1


    End Sub

    Sub addValues()
        Dim startTime As DateTime = Now()

        Call OptimizeCode_Begin()

        xlWorkBook.Worksheets.Item(1).Activate


        'Gets values from PlyBook tab
        Dim arrPlies(,) As String
        ReDim arrPlies(3, 0)

        Dim m As Integer = 1
        Do Until CStr(xlWorkBook.Sheets.Item(2).Cells(m + 1, 1).Value) = ""
            ReDim Preserve arrPlies(3, m)
            arrPlies(0, m) = xlWorkBook.Sheets.Item(2).Cells(m + 1, 1).Value
            arrPlies(1, m) = xlWorkBook.Sheets.Item(2).Cells(m + 1, 3).Value
            arrPlies(2, m) = xlWorkBook.Sheets.Item(2).Cells(m + 1, 5).Value
            arrPlies(3, m) = xlWorkBook.Sheets.Item(2).Cells(m + 1, 4).Value
            m = m + 1
        Loop


        'Gets values from standards tab
        Dim arrStandard(,) As Object
        ReDim arrStandard(15, 0)

        m = 1
        Do Until CStr(xlWorkBook.Sheets.Item(3).Cells(m + 1, 1).Value) = ""
            ReDim Preserve arrStandard(15, m)
            arrStandard(0, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 1).Value

            arrStandard(1, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 2).Value
            arrStandard(2, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 2).HorizontalAlignment
            arrStandard(3, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 2).Interior.Color
            arrStandard(4, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 2).Font.Bold
            arrStandard(5, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 2).Font.Italic

            arrStandard(6, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 3).Value
            arrStandard(7, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 3).HorizontalAlignment
            arrStandard(8, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 3).Interior.Color
            arrStandard(9, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 3).Font.Bold
            arrStandard(10, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 3).Font.Italic

            arrStandard(11, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 4).Value
            arrStandard(12, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 4).HorizontalAlignment
            arrStandard(13, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 4).Interior.Color
            arrStandard(14, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 4).Font.Bold
            arrStandard(15, m) = xlWorkBook.Sheets.Item(3).Cells(m + 1, 4).Font.Italic

            m = m + 1
        Loop

        'Find where the key line is
        Dim i As Integer
        i = 1
        Dim keyLine As Integer
        Do
            If xlWorkBook.ActiveSheet.Cells(i, 1).Value = "KEY" Then
                keyLine = i
            End If
            i = i + 1
        Loop Until keyLine > 0


        'Get the rows that the user would like to run
        Dim ExcelStartRow As Integer = CInt(Txt_ExcelStartRow.Text)
        Dim ExcelEndRow As Integer = CInt(Txt_ExcelEndRow.Text)

        If ExcelStartRow < i - 1 Then ExcelStartRow = i - 1
        If ExcelEndRow < i - 1 Then ExcelEndRow = 9999


        'Unmerge the rows that the user would like to update
        xlWorkBook.ActiveSheet.Range("B" & ExcelStartRow & ":K" & ExcelEndRow).UnMerge
        xlWorkBook.ActiveSheet.Range("B" & ExcelStartRow & ":K" & ExcelEndRow).Clear



        'Update the key row
        xlWorkBook.ActiveSheet.Range("B" & i - 1 & ":G" & i - 1).Merge
        xlWorkBook.ActiveSheet.Range("H" & i - 1 & ":I" & i - 1).Merge
        xlWorkBook.ActiveSheet.Range("J" & i - 1 & ":L" & i - 1).Merge

        xlWorkBook.ActiveSheet.Cells(i - 1, 2).Interior.Color = RGB(242, 242, 242)

        xlWorkBook.ActiveSheet.Cells(i - 1, 2).Value = "DESCRIPTION"
        xlWorkBook.ActiveSheet.Cells(i - 1, 2).HorizontalAlignment = -4108
        xlWorkBook.ActiveSheet.Cells(i - 1, 2).Interior.Color = RGB(242, 242, 242)
        xlWorkBook.ActiveSheet.Cells(i - 1, 2).Font.Bold = True

        xlWorkBook.ActiveSheet.Cells(i - 1, 8).Value = "TECH. VERIFICATION"
        xlWorkBook.ActiveSheet.Cells(i - 1, 8).HorizontalAlignment = -4108
        xlWorkBook.ActiveSheet.Cells(i - 1, 8).Interior.Color = RGB(242, 242, 242)
        xlWorkBook.ActiveSheet.Cells(i - 1, 8).Font.Bold = True

        xlWorkBook.ActiveSheet.Cells(i - 1, 10).Value = "TIME & DATE"
        xlWorkBook.ActiveSheet.Cells(i - 1, 10).HorizontalAlignment = -4108
        xlWorkBook.ActiveSheet.Cells(i - 1, 10).Interior.Color = RGB(242, 242, 242)
        xlWorkBook.ActiveSheet.Cells(i - 1, 10).Font.Bold = True



        'Find the end of the file
        Dim numericCnt As Integer = 0

        Dim loopValue As Integer = i
        Do Until CStr(xlWorkBook.ActiveSheet.Cells(loopValue, 1).Value) = ""
            If IsNumeric(xlWorkBook.ActiveSheet.Cells(loopValue, 1).Value) Then
                numericCnt = numericCnt + 1
            End If

            loopValue = loopValue + 1
        Loop
        loopValue = loopValue - 1

        Dim numTotalTime As TimeSpan = New TimeSpan(0, 0, 0, 0, 0)
        Dim txtTotalTime As TimeSpan = New TimeSpan(0, 0, 0, 0, 0)
        Dim rollNumCnt As Integer = 0
        Dim rollTxtCnt As Integer = 0
        Dim avgNumTime As Double = 0.018
        Dim avgTxtTime As Double = 0.733
        Dim timeToComplete As Double = 0

        Dim plyCount As Integer = 0

        For i = i To loopValue
            Dim currentKey As String = xlWorkBook.ActiveSheet.Cells(i, 1).Value

            If i >= ExcelStartRow And i <= ExcelEndRow Then

                xlWorkBook.ActiveSheet.Range("H" & i & ":I" & i).Merge
                xlWorkBook.ActiveSheet.Range("J" & i & ":L" & i).Merge
                If CStr(currentKey) = "PLYHEAD" Then
                    With xlWorkBook.ActiveSheet
                        .Range("C" & i & ":D" & i).Merge
                        .Range("E" & i & ":G" & i).Merge
                        .Cells(i, 2).Value = "PLY"
                        .Cells(i, 3).Value = "ORIENTATION"
                        .Cells(i, 5).Value = "MATERIAL"
                        .Cells(i, 8).Value = "TECH. VERIFICATION"
                        .Cells(i, 10).Value = "TIME & DATE"
                        .Range(xlWorkBook.ActiveSheet.Cells(i, 2), xlWorkBook.ActiveSheet.Cells(i, 10)).HorizontalAlignment = -4108
                        .Range(xlWorkBook.ActiveSheet.Cells(i, 2), xlWorkBook.ActiveSheet.Cells(i, 10)).Interior.Color = RGB(242, 242, 242)
                        .Range(xlWorkBook.ActiveSheet.Cells(i, 2), xlWorkBook.ActiveSheet.Cells(i, 10)).Font.Bold = True
                    End With

                ElseIf IsNumeric(currentKey) Then
                    plyCount += 1

                    rollNumCnt = rollNumCnt + 1

                    Dim testStartTime As DateTime = Now()

                    timeToComplete = numericCnt * avgNumTime + (loopValue - i - numericCnt) * avgTxtTime
                    ToolStripStatusLabel1.Text = "Time To Complete (s): " & Math.Round(timeToComplete, 2) & " | Current Key #: " & currentKey

                    xlWorkBook.ActiveSheet.Range("C" & i & ":D" & i).Merge
                    xlWorkBook.ActiveSheet.Range("E" & i & ":G" & i).Merge

                    Dim failedFind As Boolean = True

                    Dim z As Integer
                    For z = 1 To UBound(arrPlies, 2)
                        If CStr(arrPlies(0, z)) = CStr(currentKey) Then
                            xlWorkBook.ActiveSheet.Cells(i, 2).Value = arrPlies(1, z)
                            xlWorkBook.ActiveSheet.Cells(i, 3).Value = arrPlies(2, z)
                            xlWorkBook.ActiveSheet.Cells(i, 5).Value = arrPlies(3, z)
                            failedFind = False
                        End If
                    Next

                    xlWorkBook.ActiveSheet.Range(xlWorkBook.ActiveSheet.Cells(i, 2), xlWorkBook.ActiveSheet.Cells(i, 10)).HorizontalAlignment = -4108

                    Dim testDuration As TimeSpan = Now() - testStartTime
                    numTotalTime = numTotalTime + testDuration
                    avgNumTime = numTotalTime.TotalSeconds / rollNumCnt

                    If failedFind = True Then
                        OptimizeCode_End()
                        If MsgBox("Falied to file ply with key " & currentKey, vbOKCancel, "Error") = vbCancel Then
                            Exit Sub
                        End If
                        OptimizeCode_Begin()
                    End If

                    numericCnt = numericCnt - 1

                ElseIf currentKey = "CLEAR" Then
                    timeToComplete = numericCnt * avgNumTime + (loopValue - i - numericCnt) * avgTxtTime
                    ToolStripStatusLabel1.Text = "Time To Complete (s): " & Math.Round(timeToComplete, 2) & " | Current Key #: " & currentKey

                    xlWorkBook.ActiveSheet.Range("B" & i & ":L" & i).Merge

                Else
                    rollTxtCnt = rollTxtCnt + 1
                    Dim testStartTime As DateTime = Now()

                    timeToComplete = numericCnt * avgNumTime + (loopValue - i - numericCnt) * avgTxtTime
                    ToolStripStatusLabel1.Text = "Time To Complete (s): " & Math.Round(timeToComplete, 2) & " | Current Key #: " & currentKey

                    xlWorkBook.ActiveSheet.Range("B" & i & ":G" & i).Merge

                    Dim failedFind As Boolean = True

                    Dim y As Integer
                    For y = 1 To UBound(arrStandard, 2)
                        If CStr(arrStandard(0, y)) = CStr(currentKey) Then

                            With xlWorkBook.ActiveSheet

                                If currentKey = "BULK" Then
                                    .Cells(i, 2).Value = arrStandard(1, y) & " || PLY COUNT: " & plyCount
                                    plyCount = 0
                                Else
                                    .Cells(i, 2).Value = arrStandard(1, y)
                                End If

                                .Cells(i, 2).HorizontalAlignment = arrStandard(2, y)
                                .Cells(i, 2).Interior.Color = arrStandard(3, y)
                                .Cells(i, 2).Font.Bold = arrStandard(4, y)
                                .Cells(i, 2).Font.Italic = arrStandard(5, y)

                                .Cells(i, 8).Value = arrStandard(6, y)
                                .Cells(i, 8).HorizontalAlignment = arrStandard(7, y)
                                .Cells(i, 8).Interior.Color = arrStandard(8, y)
                                .Cells(i, 8).Font.Bold = arrStandard(9, y)
                                .Cells(i, 8).Font.Italic = arrStandard(10, y)

                                .Cells(i, 10).Value = arrStandard(11, y)
                                .Cells(i, 10).HorizontalAlignment = arrStandard(12, y)
                                .Cells(i, 10).Interior.Color = arrStandard(13, y)
                                .Cells(i, 10).Font.Bold = arrStandard(14, y)
                                .Cells(i, 10).Font.Italic = arrStandard(15, y)
                            End With

                            failedFind = False
                        End If
                    Next


                    Dim testDuration As TimeSpan = Now() - testStartTime
                    txtTotalTime = txtTotalTime + testDuration
                    avgTxtTime = txtTotalTime.TotalSeconds / rollTxtCnt

                    If failedFind = True Then
                        OptimizeCode_End()
                        If MsgBox("Falied to file standard with key " & currentKey, vbOKCancel, "Error") = vbCancel Then
                            Exit Sub
                        End If
                        OptimizeCode_Begin()
                    End If

                End If

            ElseIf IsNumeric(currentKey) Then
                plyCount += 1
            ElseIf currentKey = "BULK" Then
                plyCount = 0
            End If

        Next

        With xlWorkBook.ActiveSheet
            .Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 12)).Font.Size = 14
            .Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 12)).RowHeight = 36
            .Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 12)).VerticalAlignment = -4108
            .Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 12)).Borders.LineStyle = 1
            .Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 12)).Borders.Color = 0
            .Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 12)).Borders.Weight = 2
            .Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 12)).WrapText = True
        End With

        Call OptimizeCode_End()

        Dim duration As TimeSpan = Now() - startTime
        ToolStripStatusLabel1.Text = "Total Duration (s): " & Math.Round(duration.TotalSeconds, 2)
    End Sub

    Sub OptimizeCode_Begin()

        Excel.ScreenUpdating = False

        EventState = Excel.EnableEvents
        Excel.EnableEvents = False

        CalcState = Excel.Calculation
        Excel.Calculation = -4135

        PageBreakState = Excel.ActiveSheet.DisplayPageBreaks
        Excel.ActiveSheet.DisplayPageBreaks = False

    End Sub

    Sub OptimizeCode_End()
        Excel.ActiveSheet.DisplayPageBreaks = PageBreakState
        Excel.Calculation = CalcState
        Excel.EnableEvents = EventState
        Excel.ScreenUpdating = True
    End Sub



    Sub buildUpRoll()
        Dim totalRows As Integer = 2
        Dim cellVal As String = Excel.Cells(1, 2).Value

        Do While cellVal <> ""
            cellVal = Excel.Cells(totalRows, 2).Value
            totalRows = totalRows + 1
        Loop

        totalRows = totalRows - 2

        If Excel.Selection.Row < totalRows Then
            cellVal = Excel.Cells(Excel.Selection.Row, 2).Value
        Else
            MsgBox("Selected row is beyond the end of the table", vbOKOnly, "Error")
            Exit Sub
        End If

        Dim seqName As String = Strings.Right(cellVal, Len(cellVal) - InStr(1, cellVal, "."))

        Dim rowCnt As Integer = 0
        Dim rowStart As Integer = 0

        Dim lastRow As Integer = 0

        Dim i As Integer
        For i = 2 To totalRows
            cellVal = Excel.Cells(i, 2).Value

            Dim testVal As String = Strings.Right(cellVal, Len(cellVal) - InStr(1, cellVal, "."))

            If testVal = seqName Then
                If rowStart = 0 Then
                    rowStart = i
                    lastRow = i
                ElseIf i <> lastRow + 1 Then
                    MsgBox("The selected sequence is non-continuous; to roll up a sequence it must be continuous.", vbOKOnly, "Error")
                    Exit Sub
                End If

                lastRow = i
                rowCnt = rowCnt + 1
            End If

        Next

        If rowCnt = 1 Then
            MsgBox("The selected sequence only has 1 row.", vbOKOnly, "Error")
            Exit Sub
        End If

        Excel.Cells(rowStart, 3).Value = seqName & " (A-" & retLetter(rowCnt) & " / " & rowCnt & " PCS)"

        For i = 1 To rowCnt - 1
            Excel.Rows(rowStart + 1).EntireRow.Delete
        Next

        For i = 1 To totalRows - rowCnt
            Excel.Cells(i + 1, 1).Value = i
        Next

    End Sub

    Function retLetter(inNumber As Integer) As String
        Dim letterVal As String

        If inNumber > 26 Then
            letterVal = Chr(64 + Math.Ceiling(inNumber / 26) - 1) & Chr(64 + inNumber - ((Math.Ceiling(inNumber / 26) - 1) * 26))
        Else
            letterVal = Chr(64 + inNumber)
        End If

        Return letterVal
    End Function

    Private Sub Btn_buildUpRoll_Click(sender As Object, e As EventArgs) Handles Btn_buildUpRoll.Click
        Call buildUpRoll()
    End Sub
End Class
