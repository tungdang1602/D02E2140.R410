<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class D02F5002
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(D02F5002))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.chkCheckAssetName = New System.Windows.Forms.CheckBox()
        Me.chkCheckAssetID = New System.Windows.Forms.CheckBox()
        Me.txtNotes = New System.Windows.Forms.TextBox()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.txtVoucherNo = New System.Windows.Forms.TextBox()
        Me.lblVoucherNo = New System.Windows.Forms.Label()
        Me.c1dateVoucherDate = New C1.Win.C1Input.C1DateEdit()
        Me.tdbcVoucherTypeID = New C1.Win.C1List.C1Combo()
        Me.lblVoucherTypeID = New System.Windows.Forms.Label()
        Me.lblVoucherDate = New System.Windows.Forms.Label()
        Me.lblNotes = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.tdbg = New C1.Win.C1TrueDBGrid.C1TrueDBGrid()
        Me.chkUseBOM = New System.Windows.Forms.CheckBox()
        Me.chkReCalculate = New System.Windows.Forms.CheckBox()
        Me.btnCalculate = New System.Windows.Forms.Button()
        Me.pgb1 = New System.Windows.Forms.ProgressBar()
        Me.lblProcess = New System.Windows.Forms.Label()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.grp4 = New System.Windows.Forms.GroupBox()
        Me.tdbcToCCodeID = New C1.Win.C1List.C1Combo()
        Me.tdbcFromCCodeID = New C1.Win.C1List.C1Combo()
        Me.tdbcTypeCodeID = New C1.Win.C1List.C1Combo()
        Me.lblTypeCodeID = New System.Windows.Forms.Label()
        Me.lblFromCCodeID = New System.Windows.Forms.Label()
        Me.lblToCCodeID = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        CType(Me.c1dateVoucherDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tdbcVoucherTypeID, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.tdbg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grp4.SuspendLayout()
        CType(Me.tdbcToCCodeID, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tdbcFromCCodeID, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tdbcTypeCodeID, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkCheckAssetName)
        Me.GroupBox1.Controls.Add(Me.chkCheckAssetID)
        Me.GroupBox1.Controls.Add(Me.txtNotes)
        Me.GroupBox1.Controls.Add(Me.txtDescription)
        Me.GroupBox1.Controls.Add(Me.lblDescription)
        Me.GroupBox1.Controls.Add(Me.txtVoucherNo)
        Me.GroupBox1.Controls.Add(Me.lblVoucherNo)
        Me.GroupBox1.Controls.Add(Me.c1dateVoucherDate)
        Me.GroupBox1.Controls.Add(Me.tdbcVoucherTypeID)
        Me.GroupBox1.Controls.Add(Me.lblVoucherTypeID)
        Me.GroupBox1.Controls.Add(Me.lblVoucherDate)
        Me.GroupBox1.Controls.Add(Me.lblNotes)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(680, 158)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Bút toán khấu hao"
        '
        'chkCheckAssetName
        '
        Me.chkCheckAssetName.AutoSize = True
        Me.chkCheckAssetName.Checked = True
        Me.chkCheckAssetName.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCheckAssetName.Location = New System.Drawing.Point(25, 133)
        Me.chkCheckAssetName.Name = "chkCheckAssetName"
        Me.chkCheckAssetName.Size = New System.Drawing.Size(118, 17)
        Me.chkCheckAssetName.TabIndex = 11
        Me.chkCheckAssetName.Text = "Tên tài sản cố định"
        Me.chkCheckAssetName.UseVisualStyleBackColor = True
        '
        'chkCheckAssetID
        '
        Me.chkCheckAssetID.AutoSize = True
        Me.chkCheckAssetID.Checked = True
        Me.chkCheckAssetID.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCheckAssetID.Location = New System.Drawing.Point(25, 110)
        Me.chkCheckAssetID.Name = "chkCheckAssetID"
        Me.chkCheckAssetID.Size = New System.Drawing.Size(114, 17)
        Me.chkCheckAssetID.TabIndex = 10
        Me.chkCheckAssetID.Text = "Mã tài sản cố định"
        Me.chkCheckAssetID.UseVisualStyleBackColor = True
        '
        'txtNotes
        '
        Me.txtNotes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtNotes.Location = New System.Drawing.Point(120, 51)
        Me.txtNotes.MaxLength = 250
        Me.txtNotes.Name = "txtNotes"
        Me.txtNotes.Size = New System.Drawing.Size(548, 20)
        Me.txtNotes.TabIndex = 7
        '
        'txtDescription
        '
        Me.txtDescription.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtDescription.Location = New System.Drawing.Point(120, 79)
        Me.txtDescription.MaxLength = 250
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(547, 20)
        Me.txtDescription.TabIndex = 9
        '
        'lblDescription
        '
        Me.lblDescription.AutoSize = True
        Me.lblDescription.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescription.Location = New System.Drawing.Point(22, 84)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(82, 13)
        Me.lblDescription.TabIndex = 8
        Me.lblDescription.Text = "Diễn giải chi tiết"
        Me.lblDescription.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtVoucherNo
        '
        Me.txtVoucherNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.249999!)
        Me.txtVoucherNo.Location = New System.Drawing.Point(324, 23)
        Me.txtVoucherNo.MaxLength = 20
        Me.txtVoucherNo.Name = "txtVoucherNo"
        Me.txtVoucherNo.Size = New System.Drawing.Size(140, 20)
        Me.txtVoucherNo.TabIndex = 3
        '
        'lblVoucherNo
        '
        Me.lblVoucherNo.AutoSize = True
        Me.lblVoucherNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVoucherNo.Location = New System.Drawing.Point(255, 28)
        Me.lblVoucherNo.Name = "lblVoucherNo"
        Me.lblVoucherNo.Size = New System.Drawing.Size(49, 13)
        Me.lblVoucherNo.TabIndex = 2
        Me.lblVoucherNo.Text = "Số phiếu"
        Me.lblVoucherNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'c1dateVoucherDate
        '
        Me.c1dateVoucherDate.AutoSize = False
        Me.c1dateVoucherDate.CustomFormat = "dd/MM/yyyy"
        Me.c1dateVoucherDate.EmptyAsNull = True
        Me.c1dateVoucherDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.c1dateVoucherDate.FormatType = C1.Win.C1Input.FormatTypeEnum.CustomFormat
        Me.c1dateVoucherDate.Location = New System.Drawing.Point(556, 23)
        Me.c1dateVoucherDate.Name = "c1dateVoucherDate"
        Me.c1dateVoucherDate.Size = New System.Drawing.Size(112, 22)
        Me.c1dateVoucherDate.TabIndex = 5
        Me.c1dateVoucherDate.Tag = Nothing
        Me.c1dateVoucherDate.TrimStart = True
        Me.c1dateVoucherDate.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.DropDown
        '
        'tdbcVoucherTypeID
        '
        Me.tdbcVoucherTypeID.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.tdbcVoucherTypeID.AllowColMove = False
        Me.tdbcVoucherTypeID.AllowColSelect = True
        Me.tdbcVoucherTypeID.AllowSort = False
        Me.tdbcVoucherTypeID.AlternatingRows = True
        Me.tdbcVoucherTypeID.AutoCompletion = True
        Me.tdbcVoucherTypeID.AutoDropDown = True
        Me.tdbcVoucherTypeID.Caption = ""
        Me.tdbcVoucherTypeID.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.tdbcVoucherTypeID.ColumnWidth = 100
        Me.tdbcVoucherTypeID.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.tdbcVoucherTypeID.DisplayMember = "VoucherTypeID"
        Me.tdbcVoucherTypeID.DropdownPosition = C1.Win.C1List.DropdownPositionEnum.LeftDown
        Me.tdbcVoucherTypeID.DropDownWidth = 500
        Me.tdbcVoucherTypeID.EditorBackColor = System.Drawing.SystemColors.Window
        Me.tdbcVoucherTypeID.EditorFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcVoucherTypeID.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.tdbcVoucherTypeID.EmptyRows = True
        Me.tdbcVoucherTypeID.ExtendRightColumn = True
        Me.tdbcVoucherTypeID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcVoucherTypeID.Images.Add(CType(resources.GetObject("tdbcVoucherTypeID.Images"), System.Drawing.Image))
        Me.tdbcVoucherTypeID.Location = New System.Drawing.Point(121, 23)
        Me.tdbcVoucherTypeID.MatchEntryTimeout = CType(2000, Long)
        Me.tdbcVoucherTypeID.MaxDropDownItems = CType(8, Short)
        Me.tdbcVoucherTypeID.MaxLength = 32767
        Me.tdbcVoucherTypeID.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.tdbcVoucherTypeID.Name = "tdbcVoucherTypeID"
        Me.tdbcVoucherTypeID.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.tdbcVoucherTypeID.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbcVoucherTypeID.Size = New System.Drawing.Size(118, 21)
        Me.tdbcVoucherTypeID.TabIndex = 1
        Me.tdbcVoucherTypeID.ValueMember = "VoucherTypeID"
        Me.tdbcVoucherTypeID.PropBag = resources.GetString("tdbcVoucherTypeID.PropBag")
        '
        'lblVoucherTypeID
        '
        Me.lblVoucherTypeID.AutoSize = True
        Me.lblVoucherTypeID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVoucherTypeID.Location = New System.Drawing.Point(22, 28)
        Me.lblVoucherTypeID.Name = "lblVoucherTypeID"
        Me.lblVoucherTypeID.Size = New System.Drawing.Size(56, 13)
        Me.lblVoucherTypeID.TabIndex = 0
        Me.lblVoucherTypeID.Text = "Loại phiếu"
        Me.lblVoucherTypeID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblVoucherDate
        '
        Me.lblVoucherDate.AutoSize = True
        Me.lblVoucherDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVoucherDate.Location = New System.Drawing.Point(479, 28)
        Me.lblVoucherDate.Name = "lblVoucherDate"
        Me.lblVoucherDate.Size = New System.Drawing.Size(61, 13)
        Me.lblVoucherDate.TabIndex = 4
        Me.lblVoucherDate.Text = "Ngày phiếu"
        Me.lblVoucherDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblNotes
        '
        Me.lblNotes.AutoSize = True
        Me.lblNotes.Location = New System.Drawing.Point(22, 56)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(77, 13)
        Me.lblNotes.TabIndex = 6
        Me.lblNotes.Text = "Diễn giải phiếu"
        Me.lblNotes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.tdbg)
        Me.GroupBox2.Controls.Add(Me.chkUseBOM)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 244)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(680, 203)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        '
        'tdbg
        '
        Me.tdbg.AllowColMove = False
        Me.tdbg.AllowRowSizing = C1.Win.C1TrueDBGrid.RowSizingEnum.None
        Me.tdbg.AllowSort = False
        Me.tdbg.AlternatingRows = True
        Me.tdbg.EmptyRows = True
        Me.tdbg.Enabled = False
        Me.tdbg.ExtendRightColumn = True
        Me.tdbg.FlatStyle = C1.Win.C1TrueDBGrid.FlatModeEnum.Standard
        Me.tdbg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbg.Images.Add(CType(resources.GetObject("tdbg.Images"), System.Drawing.Image))
        Me.tdbg.Location = New System.Drawing.Point(25, 44)
        Me.tdbg.MultiSelect = C1.Win.C1TrueDBGrid.MultiSelectEnum.None
        Me.tdbg.Name = "tdbg"
        Me.tdbg.PreviewInfo.Location = New System.Drawing.Point(0, 0)
        Me.tdbg.PreviewInfo.Size = New System.Drawing.Size(0, 0)
        Me.tdbg.PreviewInfo.ZoomFactor = 75.0R
        Me.tdbg.PrintInfo.PageSettings = CType(resources.GetObject("tdbg.PrintInfo.PageSettings"), System.Drawing.Printing.PageSettings)
        Me.tdbg.PropBag = resources.GetString("tdbg.PropBag")
        Me.tdbg.RecordSelectors = False
        Me.tdbg.Size = New System.Drawing.Size(642, 147)
        Me.tdbg.TabAcrossSplits = True
        Me.tdbg.TabAction = C1.Win.C1TrueDBGrid.TabActionEnum.ColumnNavigation
        Me.tdbg.TabIndex = 1
        Me.tdbg.Tag = "COL"
        '
        'chkUseBOM
        '
        Me.chkUseBOM.AutoSize = True
        Me.chkUseBOM.Location = New System.Drawing.Point(25, 19)
        Me.chkUseBOM.Name = "chkUseBOM"
        Me.chkUseBOM.Size = New System.Drawing.Size(214, 17)
        Me.chkUseBOM.TabIndex = 0
        Me.chkUseBOM.Text = "Sử dụng bộ định mức để tính khấu hao"
        Me.chkUseBOM.UseVisualStyleBackColor = True
        '
        'chkReCalculate
        '
        Me.chkReCalculate.AutoSize = True
        Me.chkReCalculate.Checked = True
        Me.chkReCalculate.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkReCalculate.Location = New System.Drawing.Point(37, 461)
        Me.chkReCalculate.Name = "chkReCalculate"
        Me.chkReCalculate.Size = New System.Drawing.Size(181, 17)
        Me.chkReCalculate.TabIndex = 3
        Me.chkReCalculate.Text = "Tính lại các TSCĐ đã được tính"
        Me.chkReCalculate.UseVisualStyleBackColor = True
        '
        'btnCalculate
        '
        Me.btnCalculate.Location = New System.Drawing.Point(532, 456)
        Me.btnCalculate.Name = "btnCalculate"
        Me.btnCalculate.Size = New System.Drawing.Size(78, 22)
        Me.btnCalculate.TabIndex = 4
        Me.btnCalculate.Text = "&Tính"
        Me.btnCalculate.UseVisualStyleBackColor = True
        '
        'pgb1
        '
        Me.pgb1.Location = New System.Drawing.Point(91, 486)
        Me.pgb1.Name = "pgb1"
        Me.pgb1.Size = New System.Drawing.Size(601, 14)
        Me.pgb1.TabIndex = 5
        '
        'lblProcess
        '
        Me.lblProcess.AutoSize = True
        Me.lblProcess.Location = New System.Drawing.Point(34, 486)
        Me.lblProcess.Name = "lblProcess"
        Me.lblProcess.Size = New System.Drawing.Size(30, 13)
        Me.lblProcess.TabIndex = 4
        Me.lblProcess.Text = "Xử lý"
        Me.lblProcess.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(614, 456)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(78, 22)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Đó&ng"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'grp4
        '
        Me.grp4.Controls.Add(Me.tdbcToCCodeID)
        Me.grp4.Controls.Add(Me.tdbcFromCCodeID)
        Me.grp4.Controls.Add(Me.tdbcTypeCodeID)
        Me.grp4.Controls.Add(Me.lblTypeCodeID)
        Me.grp4.Controls.Add(Me.lblFromCCodeID)
        Me.grp4.Controls.Add(Me.lblToCCodeID)
        Me.grp4.Location = New System.Drawing.Point(12, 173)
        Me.grp4.Name = "grp4"
        Me.grp4.Size = New System.Drawing.Size(680, 64)
        Me.grp4.TabIndex = 1
        Me.grp4.TabStop = False
        Me.grp4.Text = "Tính khấu hao theo mã phân tích"
        '
        'tdbcToCCodeID
        '
        Me.tdbcToCCodeID.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.tdbcToCCodeID.AllowColMove = False
        Me.tdbcToCCodeID.AllowSort = False
        Me.tdbcToCCodeID.AlternatingRows = True
        Me.tdbcToCCodeID.AutoCompletion = True
        Me.tdbcToCCodeID.AutoDropDown = True
        Me.tdbcToCCodeID.AutoSelect = True
        Me.tdbcToCCodeID.Caption = ""
        Me.tdbcToCCodeID.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.tdbcToCCodeID.ColumnWidth = 100
        Me.tdbcToCCodeID.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.tdbcToCCodeID.DisplayMember = "ACodeID"
        Me.tdbcToCCodeID.DropdownPosition = C1.Win.C1List.DropdownPositionEnum.RightDown
        Me.tdbcToCCodeID.DropDownWidth = 500
        Me.tdbcToCCodeID.EditorBackColor = System.Drawing.SystemColors.Window
        Me.tdbcToCCodeID.EditorFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcToCCodeID.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.tdbcToCCodeID.EmptyRows = True
        Me.tdbcToCCodeID.ExtendRightColumn = True
        Me.tdbcToCCodeID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcToCCodeID.Images.Add(CType(resources.GetObject("tdbcToCCodeID.Images"), System.Drawing.Image))
        Me.tdbcToCCodeID.Location = New System.Drawing.Point(540, 22)
        Me.tdbcToCCodeID.MatchEntryTimeout = CType(2000, Long)
        Me.tdbcToCCodeID.MaxDropDownItems = CType(8, Short)
        Me.tdbcToCCodeID.MaxLength = 32767
        Me.tdbcToCCodeID.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.tdbcToCCodeID.Name = "tdbcToCCodeID"
        Me.tdbcToCCodeID.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.tdbcToCCodeID.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbcToCCodeID.Size = New System.Drawing.Size(128, 21)
        Me.tdbcToCCodeID.TabIndex = 4
        Me.tdbcToCCodeID.ValueMember = "ACodeID"
        Me.tdbcToCCodeID.PropBag = resources.GetString("tdbcToCCodeID.PropBag")
        '
        'tdbcFromCCodeID
        '
        Me.tdbcFromCCodeID.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.tdbcFromCCodeID.AllowColMove = False
        Me.tdbcFromCCodeID.AllowSort = False
        Me.tdbcFromCCodeID.AlternatingRows = True
        Me.tdbcFromCCodeID.AutoCompletion = True
        Me.tdbcFromCCodeID.AutoDropDown = True
        Me.tdbcFromCCodeID.AutoSelect = True
        Me.tdbcFromCCodeID.Caption = ""
        Me.tdbcFromCCodeID.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.tdbcFromCCodeID.ColumnWidth = 100
        Me.tdbcFromCCodeID.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.tdbcFromCCodeID.DisplayMember = "ACodeID"
        Me.tdbcFromCCodeID.DropdownPosition = C1.Win.C1List.DropdownPositionEnum.LeftDown
        Me.tdbcFromCCodeID.DropDownWidth = 500
        Me.tdbcFromCCodeID.EditorBackColor = System.Drawing.SystemColors.Window
        Me.tdbcFromCCodeID.EditorFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcFromCCodeID.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.tdbcFromCCodeID.EmptyRows = True
        Me.tdbcFromCCodeID.ExtendRightColumn = True
        Me.tdbcFromCCodeID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcFromCCodeID.Images.Add(CType(resources.GetObject("tdbcFromCCodeID.Images"), System.Drawing.Image))
        Me.tdbcFromCCodeID.Location = New System.Drawing.Point(362, 22)
        Me.tdbcFromCCodeID.MatchEntryTimeout = CType(2000, Long)
        Me.tdbcFromCCodeID.MaxDropDownItems = CType(8, Short)
        Me.tdbcFromCCodeID.MaxLength = 32767
        Me.tdbcFromCCodeID.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.tdbcFromCCodeID.Name = "tdbcFromCCodeID"
        Me.tdbcFromCCodeID.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.tdbcFromCCodeID.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbcFromCCodeID.Size = New System.Drawing.Size(128, 21)
        Me.tdbcFromCCodeID.TabIndex = 2
        Me.tdbcFromCCodeID.ValueMember = "ACodeID"
        Me.tdbcFromCCodeID.PropBag = resources.GetString("tdbcFromCCodeID.PropBag")
        '
        'tdbcTypeCodeID
        '
        Me.tdbcTypeCodeID.AddItemSeparator = Global.Microsoft.VisualBasic.ChrW(59)
        Me.tdbcTypeCodeID.AllowColMove = False
        Me.tdbcTypeCodeID.AllowSort = False
        Me.tdbcTypeCodeID.AlternatingRows = True
        Me.tdbcTypeCodeID.AutoCompletion = True
        Me.tdbcTypeCodeID.AutoDropDown = True
        Me.tdbcTypeCodeID.Caption = ""
        Me.tdbcTypeCodeID.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.tdbcTypeCodeID.ColumnWidth = 100
        Me.tdbcTypeCodeID.DeadAreaBackColor = System.Drawing.Color.Empty
        Me.tdbcTypeCodeID.DisplayMember = "TypeCodeID"
        Me.tdbcTypeCodeID.DropdownPosition = C1.Win.C1List.DropdownPositionEnum.LeftDown
        Me.tdbcTypeCodeID.DropDownWidth = 500
        Me.tdbcTypeCodeID.EditorBackColor = System.Drawing.SystemColors.Window
        Me.tdbcTypeCodeID.EditorFont = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcTypeCodeID.EditorForeColor = System.Drawing.SystemColors.WindowText
        Me.tdbcTypeCodeID.EmptyRows = True
        Me.tdbcTypeCodeID.ExtendRightColumn = True
        Me.tdbcTypeCodeID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.tdbcTypeCodeID.Images.Add(CType(resources.GetObject("tdbcTypeCodeID.Images"), System.Drawing.Image))
        Me.tdbcTypeCodeID.Location = New System.Drawing.Point(120, 22)
        Me.tdbcTypeCodeID.MatchEntryTimeout = CType(2000, Long)
        Me.tdbcTypeCodeID.MaxDropDownItems = CType(8, Short)
        Me.tdbcTypeCodeID.MaxLength = 32767
        Me.tdbcTypeCodeID.MouseCursor = System.Windows.Forms.Cursors.Default
        Me.tdbcTypeCodeID.Name = "tdbcTypeCodeID"
        Me.tdbcTypeCodeID.RowDivider.Style = C1.Win.C1List.LineStyleEnum.None
        Me.tdbcTypeCodeID.RowSubDividerColor = System.Drawing.Color.DarkGray
        Me.tdbcTypeCodeID.Size = New System.Drawing.Size(128, 21)
        Me.tdbcTypeCodeID.TabIndex = 1
        Me.tdbcTypeCodeID.ValueMember = "TypeCodeID"
        Me.tdbcTypeCodeID.PropBag = resources.GetString("tdbcTypeCodeID.PropBag")
        '
        'lblTypeCodeID
        '
        Me.lblTypeCodeID.AutoSize = True
        Me.lblTypeCodeID.Location = New System.Drawing.Point(22, 27)
        Me.lblTypeCodeID.Name = "lblTypeCodeID"
        Me.lblTypeCodeID.Size = New System.Drawing.Size(76, 13)
        Me.lblTypeCodeID.TabIndex = 0
        Me.lblTypeCodeID.Text = "Loại phân tích"
        Me.lblTypeCodeID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFromCCodeID
        '
        Me.lblFromCCodeID.AutoSize = True
        Me.lblFromCCodeID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFromCCodeID.Location = New System.Drawing.Point(273, 27)
        Me.lblFromCCodeID.Name = "lblFromCCodeID"
        Me.lblFromCCodeID.Size = New System.Drawing.Size(71, 13)
        Me.lblFromCCodeID.TabIndex = 3
        Me.lblFromCCodeID.Text = "Mã phân tích"
        Me.lblFromCCodeID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblToCCodeID
        '
        Me.lblToCCodeID.AutoSize = True
        Me.lblToCCodeID.Location = New System.Drawing.Point(508, 27)
        Me.lblToCCodeID.Name = "lblToCCodeID"
        Me.lblToCCodeID.Size = New System.Drawing.Size(10, 13)
        Me.lblToCCodeID.TabIndex = 5
        Me.lblToCCodeID.Text = "-"
        Me.lblToCCodeID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'D02F5002
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(704, 483)
        Me.Controls.Add(Me.grp4)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.chkReCalculate)
        Me.Controls.Add(Me.lblProcess)
        Me.Controls.Add(Me.pgb1)
        Me.Controls.Add(Me.btnCalculate)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "D02F5002"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "TÛnh khÊu hao - D02F5002"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.c1dateVoucherDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tdbcVoucherTypeID, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.tdbg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grp4.ResumeLayout(False)
        Me.grp4.PerformLayout()
        CType(Me.tdbcToCCodeID, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tdbcFromCCodeID, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tdbcTypeCodeID, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grp1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Private WithEvents txtDescription As System.Windows.Forms.TextBox
    Private WithEvents lblDescription As System.Windows.Forms.Label
    Private WithEvents txtVoucherNo As System.Windows.Forms.TextBox
    Private WithEvents lblVoucherNo As System.Windows.Forms.Label
    Private WithEvents c1dateVoucherDate As C1.Win.C1Input.C1DateEdit
    Private WithEvents tdbcVoucherTypeID As C1.Win.C1List.C1Combo
    Private WithEvents lblVoucherTypeID As System.Windows.Forms.Label
    Private WithEvents lblVoucherDate As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Private WithEvents chkUseBOM As System.Windows.Forms.CheckBox
    Private WithEvents tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid
    Private WithEvents chkReCalculate As System.Windows.Forms.CheckBox
    Private WithEvents btnCalculate As System.Windows.Forms.Button
    Private WithEvents pgb1 As System.Windows.Forms.ProgressBar
    Private WithEvents lblProcess As System.Windows.Forms.Label
    Private WithEvents btnClose As System.Windows.Forms.Button
    Private WithEvents chkCheckAssetName As System.Windows.Forms.CheckBox
    Private WithEvents chkCheckAssetID As System.Windows.Forms.CheckBox
    Private WithEvents txtNotes As System.Windows.Forms.TextBox
    Private WithEvents lblNotes As System.Windows.Forms.Label
    Friend WithEvents grp4 As System.Windows.Forms.GroupBox
    Private WithEvents tdbcTypeCodeID As C1.Win.C1List.C1Combo
    Private WithEvents lblTypeCodeID As System.Windows.Forms.Label
    Private WithEvents tdbcFromCCodeID As C1.Win.C1List.C1Combo
    Private WithEvents lblFromCCodeID As System.Windows.Forms.Label
    Private WithEvents tdbcToCCodeID As C1.Win.C1List.C1Combo
    Private WithEvents lblToCCodeID As System.Windows.Forms.Label
End Class