<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D02F2007
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim Style9 As C1.Win.C1List.Style = New C1.Win.C1List.Style()
        Dim Style10 As C1.Win.C1List.Style = New C1.Win.C1List.Style()
        Dim Style11 As C1.Win.C1List.Style = New C1.Win.C1List.Style()
        Dim Style12 As C1.Win.C1List.Style = New C1.Win.C1List.Style()
        Dim Style13 As C1.Win.C1List.Style = New C1.Win.C1List.Style()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D02F2007))
        Dim Style14 As C1.Win.C1List.Style = New C1.Win.C1List.Style()
        Dim Style15 As C1.Win.C1List.Style = New C1.Win.C1List.Style()
        Dim Style16 As C1.Win.C1List.Style = New C1.Win.C1List.Style()
        Me.c1dateChangeDate = New C1.Win.C1Input.C1DateEdit()
        Me.lblteChangeDate = New System.Windows.Forms.Label()
        Me.txtDecisionNo = New System.Windows.Forms.TextBox()
        Me.lblDecisionNo = New System.Windows.Forms.Label()
        Me.tdbcChangeNo = New C1.Win.C1List.C1Combo()
        Me.lblChangeNo = New System.Windows.Forms.Label()
        Me.txtChangeNoName = New System.Windows.Forms.TextBox()
        Me.txtNotes = New System.Windows.Forms.TextBox()
        Me.lblNotes = New System.Windows.Forms.Label()
        Me.txtNotes2 = New System.Windows.Forms.TextBox()
        Me.lblNotes2 = New System.Windows.Forms.Label()
        Me.txtNotes3 = New System.Windows.Forms.TextBox()
        Me.lblNotes3 = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.tdbg1 = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.tdbg2 = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        CType(Me.c1dateChangeDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tdbcChangeNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tdbg1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tdbg2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'c1dateChangeDate
        '
        Me.c1dateChangeDate.AutoSize = False
        Me.c1dateChangeDate.CustomFormat = "dd/MM/yyyy"
        Me.c1dateChangeDate.EmptyAsNull = True
        Me.c1dateChangeDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.c1dateChangeDate.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateChangeDate.Location = New System.Drawing.Point(87, 8)
        Me.c1dateChangeDate.Name = "c1dateChangeDate"
        Me.c1dateChangeDate.Size = New System.Drawing.Size(106, 22)
        Me.c1dateChangeDate.TabIndex = 1
        Me.c1dateChangeDate.Tag = Nothing
        Me.c1dateChangeDate.TrimStart = True
        Me.c1dateChangeDate.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'lblteChangeDate
        '
        Me.lblteChangeDate.AutoSize = True
        Me.lblteChangeDate.Location = New System.Drawing.Point(4, 14)
        Me.lblteChangeDate.Name = "lblteChangeDate"
        Me.lblteChangeDate.Size = New System.Drawing.Size(70, 13)
        Me.lblteChangeDate.TabIndex = 0
        Me.lblteChangeDate.Text = "Ngày chuyển"
        Me.lblteChangeDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDecisionNo
        '
        Me.txtDecisionNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDecisionNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtDecisionNo.Location = New System.Drawing.Point(277, 8)
        Me.txtDecisionNo.MaxLength = 20
        Me.txtDecisionNo.Name = "txtDecisionNo"
        Me.txtDecisionNo.Size = New System.Drawing.Size(121, 22)
        Me.txtDecisionNo.TabIndex = 3
        '
        'lblDecisionNo
        '
        Me.lblDecisionNo.AutoSize = True
        Me.lblDecisionNo.Location = New System.Drawing.Point(204, 12)
        Me.lblDecisionNo.Name = "lblDecisionNo"
        Me.lblDecisionNo.Size = New System.Drawing.Size(43, 13)
        Me.lblDecisionNo.TabIndex = 2
        Me.lblDecisionNo.Text = "Số hiệu"
        Me.lblDecisionNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tdbcChangeNo
        '
        Me.tdbcChangeNo.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.tdbcChangeNo.AllowColMove = False
        Me.tdbcChangeNo.AllowSort = False
        Me.tdbcChangeNo.AlternatingRows = True
        Me.tdbcChangeNo.AutoCompletion = True
        Me.tdbcChangeNo.AutoDropDown = True
        Me.tdbcChangeNo.Caption = ""
        Me.tdbcChangeNo.CaptionHeight = 17
        Me.tdbcChangeNo.CaptionStyle = Style9
        Me.tdbcChangeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.tdbcChangeNo.ColumnCaptionHeight = 17
        Me.tdbcChangeNo.ColumnFooterHeight = 17
        Me.tdbcChangeNo.ColumnWidth = 100
        Me.tdbcChangeNo.ContentHeight = 17
        Me.tdbcChangeNo.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.tdbcChangeNo.DisplayMember = "ChangeNo"
        Me.tdbcChangeNo.DropdownPosition = C1.Win.C1List.DropdownPositionEnum.LeftDown
        Me.tdbcChangeNo.DropDownWidth = 300
        Me.tdbcChangeNo.EditorBackColor = System.Drawing.SystemColors.Window
        Me.tdbcChangeNo.EditorFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcChangeNo.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.tdbcChangeNo.EditorHeight = 17
        Me.tdbcChangeNo.EmptyRows = True
        Me.tdbcChangeNo.EvenRowStyle = Style10
        Me.tdbcChangeNo.ExtendRightColumn = True
        Me.tdbcChangeNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcChangeNo.FooterStyle = Style11
        Me.tdbcChangeNo.HeadingStyle = Style12
        Me.tdbcChangeNo.HighLightRowStyle = Style13
        Me.tdbcChangeNo.Images.Add(CType(resources.GetObject("tdbcChangeNo.Images"), System.Drawing.Image))
        Me.tdbcChangeNo.ItemHeight = 15
        Me.tdbcChangeNo.Location = New System.Drawing.Point(484, 8)
        Me.tdbcChangeNo.MatchEntryTimeout = CType(2000, Long)
        Me.tdbcChangeNo.MaxDropDownItems = CType(8, Short)
        Me.tdbcChangeNo.MaxLength = 32767
        Me.tdbcChangeNo.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.tdbcChangeNo.Name = "tdbcChangeNo"
        Me.tdbcChangeNo.OddRowStyle = Style14
        Me.tdbcChangeNo.RowDivider.Color = System.Drawing.Color.DarkGray
        Me.tdbcChangeNo.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.tdbcChangeNo.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbcChangeNo.SelectedStyle = Style15
        Me.tdbcChangeNo.Size = New System.Drawing.Size(108, 23)
        Me.tdbcChangeNo.Style = Style16
        Me.tdbcChangeNo.TabIndex = 5
        Me.tdbcChangeNo.ValueMember = "ChangeNo"
        Me.tdbcChangeNo.PropBag = resources.GetString("tdbcChangeNo.PropBag")
        '
        'lblChangeNo
        '
        Me.lblChangeNo.AutoSize = True
        Me.lblChangeNo.Location = New System.Drawing.Point(404, 13)
        Me.lblChangeNo.Name = "lblChangeNo"
        Me.lblChangeNo.Size = New System.Drawing.Size(56, 13)
        Me.lblChangeNo.TabIndex = 4
        Me.lblChangeNo.Text = "Nghiệp vụ"
        Me.lblChangeNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtChangeNoName
        '
        Me.txtChangeNoName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.txtChangeNoName.Location = New System.Drawing.Point(598, 8)
        Me.txtChangeNoName.Name = "txtChangeNoName"
        Me.txtChangeNoName.ReadOnly = True
        Me.txtChangeNoName.Size = New System.Drawing.Size(204, 22)
        Me.txtChangeNoName.TabIndex = 6
        Me.txtChangeNoName.TabStop = False
        '
        'txtNotes
        '
        Me.txtNotes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtNotes.Location = New System.Drawing.Point(87, 36)
        Me.txtNotes.MaxLength = 250
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.Size = New System.Drawing.Size(715, 22)
        Me.txtNotes.TabIndex = 8
        '
        'lblNotes
        '
        Me.lblNotes.AutoSize = True
        Me.lblNotes.Location = New System.Drawing.Point(4, 41)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(53, 13)
        Me.lblNotes.TabIndex = 7
        Me.lblNotes.Text = "Ghi chú 1"
        Me.lblNotes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtNotes2
        '
        Me.txtNotes2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtNotes2.Location = New System.Drawing.Point(87, 64)
        Me.txtNotes2.MaxLength = 250
        Me.txtNotes2.Multiline = True
        Me.txtNotes2.Name = "txtNotes2"
        Me.txtNotes2.Size = New System.Drawing.Size(311, 71)
        Me.txtNotes2.TabIndex = 10
        '
        'lblNotes2
        '
        Me.lblNotes2.AutoSize = True
        Me.lblNotes2.Location = New System.Drawing.Point(4, 69)
        Me.lblNotes2.Name = "lblNotes2"
        Me.lblNotes2.Size = New System.Drawing.Size(53, 13)
        Me.lblNotes2.TabIndex = 9
        Me.lblNotes2.Text = "Ghi chú 2"
        Me.lblNotes2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtNotes3
        '
        Me.txtNotes3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtNotes3.Location = New System.Drawing.Point(484, 64)
        Me.txtNotes3.MaxLength = 250
        Me.txtNotes3.Multiline = True
        Me.txtNotes3.Name = "txtNotes3"
        Me.txtNotes3.Size = New System.Drawing.Size(318, 71)
        Me.txtNotes3.TabIndex = 12
        '
        'lblNotes3
        '
        Me.lblNotes3.AutoSize = True
        Me.lblNotes3.Location = New System.Drawing.Point(404, 69)
        Me.lblNotes3.Name = "lblNotes3"
        Me.lblNotes3.Size = New System.Drawing.Size(53, 13)
        Me.lblNotes3.TabIndex = 11
        Me.lblNotes3.Text = "Ghi chú 3"
        Me.lblNotes3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(645, 466)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(76, 22)
        Me.btnSave.TabIndex = 15
        Me.btnSave.Text = "&Lưu"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(726, 466)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 22)
        Me.btnClose.TabIndex = 16
        Me.btnClose.Text = "Đó&ng"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'tdbg1
        '
        Me.tdbg1.AllowColMove = False
        Me.tdbg1.AllowColSelect = False
        Me.tdbg1.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbg1.AllowSort = False
        Me.tdbg1.AllowUpdate = False
        Me.tdbg1.AlternatingRows = True
        Me.tdbg1.CaptionHeight = 17
        Me.tdbg1.EmptyRows = True
        Me.tdbg1.ExtendRightColumn = True
        Me.tdbg1.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbg1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbg1.GroupByCaption = "Drag a column header here to group by that column"
        Me.tdbg1.Images.Add(CType(resources.GetObject("tdbg1.Images"), System.Drawing.Image))
        Me.tdbg1.Location = New System.Drawing.Point(8, 158)
        Me.tdbg1.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbg1.Name = "tdbg1"
        Me.tdbg1.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbg1.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbg1.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbg1.PrintInfo.PageSettings = CType(resources.GetObject("tdbg1.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbg1.RowHeight = 15
        Me.tdbg1.Size = New System.Drawing.Size(794, 142)
        Me.tdbg1.TabAcrossSplits = True
        Me.tdbg1.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbg1.TabIndex = 13
        Me.tdbg1.Tag = "COL1"
        Me.tdbg1.PropBag = resources.GetString("tdbg1.PropBag")
        '
        'tdbg2
        '
        Me.tdbg2.AllowColMove = False
        Me.tdbg2.AllowColSelect = False
        Me.tdbg2.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbg2.AllowSort = False
        Me.tdbg2.AlternatingRows = True
        Me.tdbg2.CaptionHeight = 17
        Me.tdbg2.EmptyRows = True
        Me.tdbg2.ExtendRightColumn = True
        Me.tdbg2.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbg2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbg2.GroupByCaption = "Drag a column header here to group by that column"
        Me.tdbg2.Images.Add(CType(resources.GetObject("tdbg2.Images"), System.Drawing.Image))
        Me.tdbg2.Location = New System.Drawing.Point(9, 317)
        Me.tdbg2.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbg2.Name = "tdbg2"
        Me.tdbg2.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbg2.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbg2.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbg2.PrintInfo.PageSettings = CType(resources.GetObject("tdbg2.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbg2.RowHeight = 15
        Me.tdbg2.Size = New System.Drawing.Size(793, 142)
        Me.tdbg2.TabAcrossSplits = True
        Me.tdbg2.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbg2.TabIndex = 14
        Me.tdbg2.Tag = "COL2"
        Me.tdbg2.PropBag = resources.GetString("tdbg2.PropBag")
        '
        'D02F2007
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(809, 493)
        Me.Controls.Add(Me.tdbg2)
        Me.Controls.Add(Me.tdbg1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.txtNotes3)
        Me.Controls.Add(Me.txtNotes2)
        Me.Controls.Add(Me.txtNotes)
        Me.Controls.Add(Me.tdbcChangeNo)
        Me.Controls.Add(Me.txtDecisionNo)
        Me.Controls.Add(Me.c1dateChangeDate)
        Me.Controls.Add(Me.lblteChangeDate)
        Me.Controls.Add(Me.lblDecisionNo)
        Me.Controls.Add(Me.lblChangeNo)
        Me.Controls.Add(Me.txtChangeNoName)
        Me.Controls.Add(Me.lblNotes)
        Me.Controls.Add(Me.lblNotes2)
        Me.Controls.Add(Me.lblNotes3)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D02F2007"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CËp nhËt nghiÖp vó chuyÓn nguän - D02F2007"
        CType(Me.c1dateChangeDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tdbcChangeNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tdbg1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tdbg2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents c1dateChangeDate As C1.Win.C1Input.C1DateEdit
    Private WithEvents lblteChangeDate As System.Windows.Forms.Label
    Private WithEvents txtDecisionNo As System.Windows.Forms.TextBox
    Private WithEvents lblDecisionNo As System.Windows.Forms.Label
    Private WithEvents tdbcChangeNo As C1.Win.C1List.C1Combo
    Private WithEvents lblChangeNo As System.Windows.Forms.Label
    Private WithEvents txtChangeNoName As System.Windows.Forms.TextBox
    Private WithEvents txtNotes As System.Windows.Forms.TextBox
    Private WithEvents lblNotes As System.Windows.Forms.Label
    Private WithEvents txtNotes2 As System.Windows.Forms.TextBox
    Private WithEvents lblNotes2 As System.Windows.Forms.Label
    Private WithEvents txtNotes3 As System.Windows.Forms.TextBox
    Private WithEvents lblNotes3 As System.Windows.Forms.Label
    Private WithEvents btnSave As System.Windows.Forms.Button
    Private WithEvents btnClose As System.Windows.Forms.Button
    Private WithEvents tdbg1 As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Private WithEvents tdbg2 As C1.Win.C1TrueDBGrid.C1TrueDBGrid
End Class