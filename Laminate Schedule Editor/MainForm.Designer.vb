﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.FilePath = New System.Windows.Forms.TextBox()
        Me.Btn_Browse = New System.Windows.Forms.Button()
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.Btn_Open = New System.Windows.Forms.Button()
        Me.Btn_PlyStdCreate = New System.Windows.Forms.Button()
        Me.Btn_ReRunVals = New System.Windows.Forms.Button()
        Me.Txt_DebulkConst = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Btn_wrkUpdate = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel8 = New System.Windows.Forms.TableLayoutPanel()
        Me.Chk_FirstPlyDebulk = New System.Windows.Forms.CheckBox()
        Me.TableLayoutPanel9 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel5 = New System.Windows.Forms.TableLayoutPanel()
        Me.Btn_buildUpRoll = New System.Windows.Forms.Button()
        Me.Btn_ShtHeaders = New System.Windows.Forms.Button()
        Me.TableLayoutPanel6 = New System.Windows.Forms.TableLayoutPanel()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel7 = New System.Windows.Forms.TableLayoutPanel()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.Txt_ExcelStartRow = New System.Windows.Forms.TextBox()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.Txt_ExcelEndRow = New System.Windows.Forms.TextBox()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.TableLayoutPanel4.SuspendLayout()
        Me.TableLayoutPanel8.SuspendLayout()
        Me.TableLayoutPanel9.SuspendLayout()
        Me.TableLayoutPanel5.SuspendLayout()
        Me.TableLayoutPanel6.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.TableLayoutPanel7.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'FilePath
        '
        Me.FilePath.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FilePath.Location = New System.Drawing.Point(3, 21)
        Me.FilePath.Name = "FilePath"
        Me.FilePath.Size = New System.Drawing.Size(557, 20)
        Me.FilePath.TabIndex = 0
        '
        'Btn_Browse
        '
        Me.Btn_Browse.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Btn_Browse.Location = New System.Drawing.Point(566, 3)
        Me.Btn_Browse.Name = "Btn_Browse"
        Me.Btn_Browse.Size = New System.Drawing.Size(94, 56)
        Me.Btn_Browse.TabIndex = 1
        Me.Btn_Browse.Text = "Browse"
        Me.Btn_Browse.UseVisualStyleBackColor = True
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.FileName = "OpenFileDialog"
        Me.OpenFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
        '
        'Btn_Open
        '
        Me.Btn_Open.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Btn_Open.Location = New System.Drawing.Point(566, 65)
        Me.Btn_Open.Name = "Btn_Open"
        Me.Btn_Open.Size = New System.Drawing.Size(94, 56)
        Me.Btn_Open.TabIndex = 2
        Me.Btn_Open.Text = "Open File"
        Me.Btn_Open.UseVisualStyleBackColor = True
        '
        'Btn_PlyStdCreate
        '
        Me.Btn_PlyStdCreate.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Btn_PlyStdCreate.Location = New System.Drawing.Point(3, 23)
        Me.Btn_PlyStdCreate.Name = "Btn_PlyStdCreate"
        Me.Btn_PlyStdCreate.Size = New System.Drawing.Size(149, 66)
        Me.Btn_PlyStdCreate.TabIndex = 0
        Me.Btn_PlyStdCreate.Text = "Ply Standard Create"
        Me.Btn_PlyStdCreate.UseVisualStyleBackColor = True
        '
        'Btn_ReRunVals
        '
        Me.Btn_ReRunVals.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Btn_ReRunVals.Location = New System.Drawing.Point(3, 3)
        Me.Btn_ReRunVals.Name = "Btn_ReRunVals"
        Me.Btn_ReRunVals.Size = New System.Drawing.Size(182, 62)
        Me.Btn_ReRunVals.TabIndex = 1
        Me.Btn_ReRunVals.Text = "Apply Values"
        Me.Btn_ReRunVals.UseVisualStyleBackColor = True
        '
        'Txt_DebulkConst
        '
        Me.Txt_DebulkConst.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_DebulkConst.Location = New System.Drawing.Point(3, 16)
        Me.Txt_DebulkConst.Name = "Txt_DebulkConst"
        Me.Txt_DebulkConst.Size = New System.Drawing.Size(115, 20)
        Me.Txt_DebulkConst.TabIndex = 0
        Me.Txt_DebulkConst.Text = "4"
        Me.Txt_DebulkConst.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Txt_DebulkConst)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(121, 50)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Layers Per De-Bulk"
        '
        'Btn_wrkUpdate
        '
        Me.Btn_wrkUpdate.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Btn_wrkUpdate.Location = New System.Drawing.Point(3, 3)
        Me.Btn_wrkUpdate.Name = "Btn_wrkUpdate"
        Me.Btn_wrkUpdate.Size = New System.Drawing.Size(157, 39)
        Me.Btn_wrkUpdate.TabIndex = 0
        Me.Btn_wrkUpdate.Text = "Header/Footer Update"
        Me.Btn_wrkUpdate.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.TableLayoutPanel1)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(669, 143)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "File To Edit"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.FilePath, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Btn_Browse, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Btn_Open, 1, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(3, 16)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(663, 124)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 1
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.GroupBox2, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel3, 0, 1)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 3
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(675, 319)
        Me.TableLayoutPanel2.TabIndex = 9
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 3
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 45.71918!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 28.93836!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.GroupBox3, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel5, 2, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel6, 1, 0)
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(3, 152)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 1
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(669, 143)
        Me.TableLayoutPanel3.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.TableLayoutPanel4)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(300, 137)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Create Col A Keys"
        '
        'TableLayoutPanel4
        '
        Me.TableLayoutPanel4.ColumnCount = 2
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 55.10204!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 44.89796!))
        Me.TableLayoutPanel4.Controls.Add(Me.TableLayoutPanel8, 1, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.TableLayoutPanel9, 0, 0)
        Me.TableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel4.Location = New System.Drawing.Point(3, 16)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        Me.TableLayoutPanel4.RowCount = 1
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(294, 118)
        Me.TableLayoutPanel4.TabIndex = 0
        '
        'TableLayoutPanel8
        '
        Me.TableLayoutPanel8.ColumnCount = 1
        Me.TableLayoutPanel8.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel8.Controls.Add(Me.GroupBox1, 0, 0)
        Me.TableLayoutPanel8.Controls.Add(Me.Chk_FirstPlyDebulk, 0, 1)
        Me.TableLayoutPanel8.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel8.Location = New System.Drawing.Point(164, 3)
        Me.TableLayoutPanel8.Name = "TableLayoutPanel8"
        Me.TableLayoutPanel8.RowCount = 2
        Me.TableLayoutPanel8.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel8.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel8.Size = New System.Drawing.Size(127, 112)
        Me.TableLayoutPanel8.TabIndex = 7
        '
        'Chk_FirstPlyDebulk
        '
        Me.Chk_FirstPlyDebulk.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Chk_FirstPlyDebulk.AutoSize = True
        Me.Chk_FirstPlyDebulk.Checked = True
        Me.Chk_FirstPlyDebulk.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Chk_FirstPlyDebulk.Location = New System.Drawing.Point(3, 75)
        Me.Chk_FirstPlyDebulk.Name = "Chk_FirstPlyDebulk"
        Me.Chk_FirstPlyDebulk.Size = New System.Drawing.Size(121, 17)
        Me.Chk_FirstPlyDebulk.TabIndex = 7
        Me.Chk_FirstPlyDebulk.Text = "First Ply Debulk"
        Me.Chk_FirstPlyDebulk.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel9
        '
        Me.TableLayoutPanel9.ColumnCount = 1
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel9.Controls.Add(Me.Btn_PlyStdCreate, 0, 1)
        Me.TableLayoutPanel9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel9.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel9.Name = "TableLayoutPanel9"
        Me.TableLayoutPanel9.RowCount = 3
        Me.TableLayoutPanel9.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel9.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel9.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel9.Size = New System.Drawing.Size(155, 112)
        Me.TableLayoutPanel9.TabIndex = 0
        '
        'TableLayoutPanel5
        '
        Me.TableLayoutPanel5.ColumnCount = 1
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel5.Controls.Add(Me.Btn_wrkUpdate, 0, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.Btn_buildUpRoll, 0, 1)
        Me.TableLayoutPanel5.Controls.Add(Me.Btn_ShtHeaders, 0, 2)
        Me.TableLayoutPanel5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel5.Location = New System.Drawing.Point(503, 3)
        Me.TableLayoutPanel5.Name = "TableLayoutPanel5"
        Me.TableLayoutPanel5.RowCount = 3
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.TableLayoutPanel5.Size = New System.Drawing.Size(163, 137)
        Me.TableLayoutPanel5.TabIndex = 3
        '
        'Btn_buildUpRoll
        '
        Me.Btn_buildUpRoll.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Btn_buildUpRoll.Location = New System.Drawing.Point(3, 48)
        Me.Btn_buildUpRoll.Name = "Btn_buildUpRoll"
        Me.Btn_buildUpRoll.Size = New System.Drawing.Size(157, 39)
        Me.Btn_buildUpRoll.TabIndex = 1
        Me.Btn_buildUpRoll.Text = "Build-up Roll"
        Me.Btn_buildUpRoll.UseVisualStyleBackColor = True
        '
        'Btn_ShtHeaders
        '
        Me.Btn_ShtHeaders.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Btn_ShtHeaders.Location = New System.Drawing.Point(3, 93)
        Me.Btn_ShtHeaders.Name = "Btn_ShtHeaders"
        Me.Btn_ShtHeaders.Size = New System.Drawing.Size(157, 41)
        Me.Btn_ShtHeaders.TabIndex = 2
        Me.Btn_ShtHeaders.Text = "PLYHEAD Per Sheet"
        Me.Btn_ShtHeaders.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel6
        '
        Me.TableLayoutPanel6.ColumnCount = 1
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel6.Controls.Add(Me.Btn_ReRunVals, 0, 0)
        Me.TableLayoutPanel6.Controls.Add(Me.GroupBox4, 0, 1)
        Me.TableLayoutPanel6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel6.Location = New System.Drawing.Point(309, 3)
        Me.TableLayoutPanel6.Name = "TableLayoutPanel6"
        Me.TableLayoutPanel6.RowCount = 2
        Me.TableLayoutPanel6.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel6.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel6.Size = New System.Drawing.Size(188, 137)
        Me.TableLayoutPanel6.TabIndex = 2
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.TableLayoutPanel7)
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox4.Location = New System.Drawing.Point(3, 71)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(182, 63)
        Me.GroupBox4.TabIndex = 5
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Excel Row Limits"
        '
        'TableLayoutPanel7
        '
        Me.TableLayoutPanel7.ColumnCount = 2
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel7.Controls.Add(Me.GroupBox5, 0, 0)
        Me.TableLayoutPanel7.Controls.Add(Me.GroupBox6, 1, 0)
        Me.TableLayoutPanel7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel7.Location = New System.Drawing.Point(3, 16)
        Me.TableLayoutPanel7.Name = "TableLayoutPanel7"
        Me.TableLayoutPanel7.RowCount = 1
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel7.Size = New System.Drawing.Size(176, 44)
        Me.TableLayoutPanel7.TabIndex = 0
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.Txt_ExcelStartRow)
        Me.GroupBox5.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(82, 38)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Start"
        '
        'Txt_ExcelStartRow
        '
        Me.Txt_ExcelStartRow.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_ExcelStartRow.Location = New System.Drawing.Point(3, 16)
        Me.Txt_ExcelStartRow.Name = "Txt_ExcelStartRow"
        Me.Txt_ExcelStartRow.Size = New System.Drawing.Size(76, 20)
        Me.Txt_ExcelStartRow.TabIndex = 0
        Me.Txt_ExcelStartRow.Text = "0"
        Me.Txt_ExcelStartRow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.Txt_ExcelEndRow)
        Me.GroupBox6.Location = New System.Drawing.Point(91, 3)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(82, 38)
        Me.GroupBox6.TabIndex = 3
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "End"
        '
        'Txt_ExcelEndRow
        '
        Me.Txt_ExcelEndRow.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_ExcelEndRow.Location = New System.Drawing.Point(3, 16)
        Me.Txt_ExcelEndRow.Name = "Txt_ExcelEndRow"
        Me.Txt_ExcelEndRow.Size = New System.Drawing.Size(76, 20)
        Me.Txt_ExcelEndRow.TabIndex = 1
        Me.Txt_ExcelEndRow.Text = "0"
        Me.Txt_ExcelEndRow.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2, Me.ToolStripProgressBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 297)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(675, 22)
        Me.StatusStrip1.TabIndex = 10
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(48, 17)
        Me.ToolStripStatusLabel1.Text = "Status..."
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 16)
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(479, 17)
        Me.ToolStripStatusLabel2.Spring = True
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(675, 319)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.TableLayoutPanel2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "MainForm"
        Me.Text = "Laminate Schedule Update"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.TableLayoutPanel8.ResumeLayout(False)
        Me.TableLayoutPanel8.PerformLayout()
        Me.TableLayoutPanel9.ResumeLayout(False)
        Me.TableLayoutPanel5.ResumeLayout(False)
        Me.TableLayoutPanel6.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.TableLayoutPanel7.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents FilePath As TextBox
    Friend WithEvents Btn_Browse As Button
    Friend WithEvents OpenFileDialog As OpenFileDialog
    Friend WithEvents Btn_Open As Button
    Friend WithEvents Btn_PlyStdCreate As Button
    Friend WithEvents Btn_ReRunVals As Button
    Friend WithEvents Txt_DebulkConst As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Btn_wrkUpdate As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents TableLayoutPanel3 As TableLayoutPanel
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents TableLayoutPanel4 As TableLayoutPanel
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents TableLayoutPanel5 As TableLayoutPanel
    Friend WithEvents Btn_buildUpRoll As Button
    Friend WithEvents TableLayoutPanel6 As TableLayoutPanel
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents TableLayoutPanel7 As TableLayoutPanel
    Friend WithEvents Txt_ExcelEndRow As TextBox
    Friend WithEvents Txt_ExcelStartRow As TextBox
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents GroupBox6 As GroupBox
    Friend WithEvents Btn_ShtHeaders As Button
    Friend WithEvents TableLayoutPanel8 As TableLayoutPanel
    Friend WithEvents Chk_FirstPlyDebulk As CheckBox
    Friend WithEvents TableLayoutPanel9 As TableLayoutPanel
    Friend WithEvents ToolStripProgressBar1 As ToolStripProgressBar
    Friend WithEvents ToolStripStatusLabel2 As ToolStripStatusLabel
End Class
