Imports Excel = Microsoft.Office.Interop.Excel

Public Class MainForm
    Dim xlWorkBook As Excel.Workbook
    Dim Excel As Object

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

        Dim leftFooter As String
        leftFooter = "&12&""Calibri""&B" & "Doc. No. " & xlWorkBook.ActiveSheet.Cells(3, 4).Value & "_" & xlWorkBook.ActiveSheet.Cells(3, 7).Value

        xlWorkBook.ActiveSheet.PageSetup.leftFooter = leftFooter
        xlWorkBook.ActiveSheet.PageSetup.FirstPage.leftFooter.Text = leftFooter

        Dim rightHeader As String
        rightHeader = "&18&""Calibri""&B" & vbCr & xlWorkBook.ActiveSheet.Cells(3, 4).Value & "_" & xlWorkBook.ActiveSheet.Cells(3, 7).Value & " | " & xlWorkBook.ActiveSheet.Cells(4, 4).Value & vbCr & xlWorkBook.ActiveSheet.Cells(6, 4).Value & vbCr & xlWorkBook.ActiveSheet.Cells(8, 4).Value & " | " & xlWorkBook.ActiveSheet.Cells(9, 4).Value
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

        xlWorkBook.Worksheets.Item(1).Activate

        Dim i As Integer
        i = 1
        Dim keyLine As Integer
        Do
            If xlWorkBook.ActiveSheet.Cells(i, 1).Value = "KEY" Then
                keyLine = i
            End If
            i = i + 1
        Loop Until keyLine > 0

        xlWorkBook.ActiveSheet.Range("B" & i - 1 & ":K9999").UnMerge
        xlWorkBook.ActiveSheet.Range("B" & i - 1 & ":K9999").Clear

        xlWorkBook.ActiveSheet.Range("B" & i - 1 & ":G" & i - 1).Merge
        xlWorkBook.ActiveSheet.Range("H" & i - 1 & ":I" & i - 1).Merge
        xlWorkBook.ActiveSheet.Range("J" & i - 1 & ":K" & i - 1).Merge
        xlWorkBook.ActiveSheet.Cells(i - 1, 8).Value = "TECH. VERIFICATION"
        xlWorkBook.ActiveSheet.Cells(i - 1, 8).HorizontalAlignment = -4108
        xlWorkBook.ActiveSheet.Cells(i - 1, 8).Interior.Color = RGB(242, 242, 242)
        xlWorkBook.ActiveSheet.Cells(i - 1, 8).Font.Bold = True
        xlWorkBook.ActiveSheet.Cells(i - 1, 10).Value = "TIME & DATE"
        xlWorkBook.ActiveSheet.Cells(i - 1, 10).HorizontalAlignment = -4108
        xlWorkBook.ActiveSheet.Cells(i - 1, 10).Interior.Color = RGB(242, 242, 242)
        xlWorkBook.ActiveSheet.Cells(i - 1, 10).Font.Bold = True

        Do Until CStr(xlWorkBook.ActiveSheet.Cells(i, 1).Value) = ""
            xlWorkBook.ActiveSheet.Range("H" & i & ":I" & i).Merge
            xlWorkBook.ActiveSheet.Range("J" & i & ":K" & i).Merge
            If CStr(xlWorkBook.ActiveSheet.Cells(i, 1).Value) = "PLYHEAD" Then
                xlWorkBook.ActiveSheet.Range("C" & i & ":D" & i).Merge
                xlWorkBook.ActiveSheet.Range("E" & i & ":G" & i).Merge
                xlWorkBook.ActiveSheet.Cells(i, 2).Value = "PLY"
                xlWorkBook.ActiveSheet.Cells(i, 3).Value = "ORIENTATION"
                xlWorkBook.ActiveSheet.Cells(i, 5).Value = "MATERIAL"
                xlWorkBook.ActiveSheet.Cells(i, 8).Value = "TECH. VERIFICATION"
                xlWorkBook.ActiveSheet.Cells(i, 10).Value = "TIME & DATE"
                xlWorkBook.ActiveSheet.Range(xlWorkBook.ActiveSheet.Cells(i, 2), xlWorkBook.ActiveSheet.Cells(i, 10)).HorizontalAlignment = -4108
                xlWorkBook.ActiveSheet.Range(xlWorkBook.ActiveSheet.Cells(i, 2), xlWorkBook.ActiveSheet.Cells(i, 10)).Interior.Color = RGB(242, 242, 242)
                xlWorkBook.ActiveSheet.Range(xlWorkBook.ActiveSheet.Cells(i, 2), xlWorkBook.ActiveSheet.Cells(i, 10)).Font.Bold = True

            ElseIf IsNumeric(xlWorkBook.ActiveSheet.Cells(i, 1).Value) Then
                xlWorkBook.ActiveSheet.Range("C" & i & ":D" & i).Merge
                xlWorkBook.ActiveSheet.Range("E" & i & ":G" & i).Merge

                Dim matchLine As Integer
                matchLine = 0
                Dim z As Integer
                z = 1
                Do
                    If CStr(xlWorkBook.Sheets.Item(2).Cells(z, 1).Value) = CStr(xlWorkBook.ActiveSheet.Cells(i, 1).Value) Then
                        matchLine = z
                    End If
                    If z = 9999 Then
                        MsgBox("Failed to match ply.", , "Error")
                        Exit Sub
                    End If
                    z = z + 1
                Loop Until matchLine <> 0

                xlWorkBook.ActiveSheet.Cells(i, 2).Value = xlWorkBook.Sheets.Item(2).Cells(matchLine, 3).Value
                xlWorkBook.ActiveSheet.Cells(i, 3).Value = xlWorkBook.Sheets.Item(2).Cells(matchLine, 5).Value
                xlWorkBook.ActiveSheet.Cells(i, 5).Value = xlWorkBook.Sheets.Item(2).Cells(matchLine, 4).Value

                xlWorkBook.ActiveSheet.Range(xlWorkBook.ActiveSheet.Cells(i, 2), xlWorkBook.ActiveSheet.Cells(i, 10)).HorizontalAlignment = -4108
            Else
                xlWorkBook.ActiveSheet.Range("B" & i & ":G" & i).Merge

                Dim matchLine2 As Integer
                matchLine2 = 0
                Dim y As Integer
                y = 1
                Do
                    If xlWorkBook.Sheets.Item(3).Cells(y, 1).Value = xlWorkBook.ActiveSheet.Cells(i, 1).Value Then
                        matchLine2 = y
                    End If
                    If y = 9999 Then
                        MsgBox("Failed to match ply.", , "Error")
                        Exit Sub
                    End If
                    y = y + 1
                Loop Until matchLine2 <> 0

                xlWorkBook.ActiveSheet.Cells(i, 2).Value = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 2).Value
                xlWorkBook.ActiveSheet.Cells(i, 2).HorizontalAlignment = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 2).HorizontalAlignment
                xlWorkBook.ActiveSheet.Cells(i, 2).Interior.Color = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 2).Interior.Color
                xlWorkBook.ActiveSheet.Cells(i, 2).Font.Bold = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 2).Font.Bold
                xlWorkBook.ActiveSheet.Cells(i, 2).Font.Italic = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 2).Font.Italic

                xlWorkBook.ActiveSheet.Cells(i, 8).Value = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 3).Value
                xlWorkBook.ActiveSheet.Cells(i, 8).HorizontalAlignment = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 3).HorizontalAlignment
                xlWorkBook.ActiveSheet.Cells(i, 8).Interior.Color = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 3).Interior.Color
                xlWorkBook.ActiveSheet.Cells(i, 8).Font.Bold = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 3).Font.Bold
                xlWorkBook.ActiveSheet.Cells(i, 8).Font.Italic = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 3).Font.Italic

                xlWorkBook.ActiveSheet.Cells(i, 10).Value = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 4).Value
                xlWorkBook.ActiveSheet.Cells(i, 10).HorizontalAlignment = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 4).HorizontalAlignment
                xlWorkBook.ActiveSheet.Cells(i, 10).Interior.Color = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 4).Interior.Color
                xlWorkBook.ActiveSheet.Cells(i, 10).Font.Bold = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 4).Font.Bold
                xlWorkBook.ActiveSheet.Cells(i, 10).Font.Italic = xlWorkBook.Sheets.Item(3).Cells(matchLine2, 4).Font.Italic



            End If
            i = i + 1
        Loop

        xlWorkBook.ActiveSheet.Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 11)).Font.Size = 14
        xlWorkBook.ActiveSheet.Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 11)).RowHeight = 36
        xlWorkBook.ActiveSheet.Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 11)).VerticalAlignment = -4108
        xlWorkBook.ActiveSheet.Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 11)).Borders.LineStyle = 1
        xlWorkBook.ActiveSheet.Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 11)).Borders.Color = 0
        xlWorkBook.ActiveSheet.Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 11)).Borders.Weight = 2
        xlWorkBook.ActiveSheet.Range(xlWorkBook.ActiveSheet.Cells(keyLine, 2), xlWorkBook.ActiveSheet.Cells(i - 1, 11)).WrapText = True

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
