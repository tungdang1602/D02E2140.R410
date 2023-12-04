Imports System.Data
Imports System
Public Class D02F3005

#Region "Const of tdbg"
    Private Const COL_TransactionID As String = "TransactionID"     ' TransactionID
    Private Const COL_BatchID As String = "BatchID"                 ' BatchID
    Private Const COL_ModuleID As String = "ModuleID"               ' ModuleID
    Private Const COL_DivisionID As String = "DivisionID"           ' DivisionID
    Private Const COL_Choose As String = "Choose"                   ' Chọn
    Private Const COL_VoucherTypeID As String = "VoucherTypeID"     ' Loại phiếu
    Private Const COL_VoucherNo As String = "VoucherNo"             ' Số phiếu
    Private Const COL_VoucherDate As String = "VoucherDate"         ' Ngày phiếu
    Private Const COL_RefDate As String = "RefDate"                 ' Ngày hóa đơn
    Private Const COL_SeriNo As String = "SeriNo"                   ' Số Sêri
    Private Const COL_RefNo As String = "RefNo"                     ' Số hóa đơn
    Private Const COL_ObjectTypeID As String = "ObjectTypeID"       ' Loại đối tượng
    Private Const COL_ObjectID As String = "ObjectID"               ' Đối tượng
    Private Const COL_Description As String = "Description"         ' Diễn giải
    Private Const COL_DebitAccountID As String = "DebitAccountID"   ' Tài khoản nợ
    Private Const COL_CreditAccountID As String = "CreditAccountID" ' Tài khoản có
    Private Const COL_CurrencyID As String = "CurrencyID"           ' Loại tiền
    Private Const COL_ExchangeRate As String = "ExchangeRate"       ' Tỷ giá
    Private Const COL_OriginalAmount As String = "OriginalAmount"   ' Nguyên tệ
    Private Const COL_ConvertedAmount As String = "ConvertedAmount" ' Quy đổi
    Private Const COL_Ana01ID As String = "Ana01ID"                 ' Ana01ID
    Private Const COL_Ana02ID As String = "Ana02ID"                 ' Ana02ID
    Private Const COL_Ana03ID As String = "Ana03ID"                 ' Ana03ID
    Private Const COL_Ana04ID As String = "Ana04ID"                 ' Ana04ID
    Private Const COL_Ana05ID As String = "Ana05ID"                 ' Ana05ID
    Private Const COL_Ana06ID As String = "Ana06ID"                 ' Ana06ID
    Private Const COL_Ana07ID As String = "Ana07ID"                 ' Ana07ID
    Private Const COL_Ana08ID As String = "Ana08ID"                 ' Ana08ID
    Private Const COL_Ana09ID As String = "Ana09ID"                 ' Ana09ID
    Private Const COL_Ana10ID As String = "Ana10ID"                 ' Ana10ID
    Private Const COL_OriginalCipID As String = "OriginalCipID"     ' OriginalCipID
#End Region

    Private _mode As Integer = 0
    Public WriteOnly Property Mode() As Integer
        Set(ByVal Value As Integer)
            _mode = Value
        End Set
    End Property


    Private _bSavedOK As Boolean = False
    Public ReadOnly Property bSavedOK() As Boolean
        Get
            Return _bSavedOK
        End Get
    End Property

    Private _dtChose As DataTable = Nothing
    Public ReadOnly Property dtChose() As DataTable
        Get
            Return _dtChose
        End Get
    End Property

    Private Sub D02F3005_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        ElseIf e.KeyCode = Keys.F11 Then
            HotKeyF11(Me, tdbg)
        ElseIf e.KeyCode = Keys.F5 Then
            btnFilter_Click(sender, Nothing)
        End If
    End Sub

    Private Sub D02F3005_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        setShortCut()
        SetBackColorObligatory()
        tdbg_NumberFormat()
        LoadTDBCombo()

        Loadlanguage()
        ResetColorGrid(tdbg, SPLIT0, tdbg.Splits.Count - 1)
        tdbg.Splits(0).MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.DottedCellBorder
        tdbg.Splits(2).MarqueeStyle = C1.Win.C1TrueDBGrid.MarqueeEnum.FloatingEditor
        ResetSplitDividerSize(tdbg)
        InputDateInTrueDBGrid(tdbg, COL_VoucherDate, COL_RefDate)
        'LoadTDBGrid()
        'Gán dropdown cho cột khoản muc
        tdbg.Columns(COL_Ana01ID).DropDown = tdbdAna01ID
        tdbg.Columns(COL_Ana02ID).DropDown = tdbdAna02ID
        tdbg.Columns(COL_Ana03ID).DropDown = tdbdAna03ID
        tdbg.Columns(COL_Ana04ID).DropDown = tdbdAna04ID
        tdbg.Columns(COL_Ana05ID).DropDown = tdbdAna05ID
        tdbg.Columns(COL_Ana06ID).DropDown = tdbdAna06ID
        tdbg.Columns(COL_Ana07ID).DropDown = tdbdAna07ID
        tdbg.Columns(COL_Ana08ID).DropDown = tdbdAna08ID
        tdbg.Columns(COL_Ana09ID).DropDown = tdbdAna09ID
        tdbg.Columns(COL_Ana10ID).DropDown = tdbdAna10ID
        For i As Integer = 0 To 9
            tdbg.Splits(2).DisplayColumns(i + IndexOfColumn(tdbg, COL_Ana01ID)).AutoDropDown = True
            tdbg.Splits(2).DisplayColumns(i + IndexOfColumn(tdbg, COL_Ana01ID)).AutoComplete = True
        Next
        LoadTDBGridAnalysisCaption(D02, tdbg, IndexOfColumn(tdbg, COL_Ana01ID), 2, True, gbUnicode)
        LoadTDBDropDownAna(tdbdAna01ID, tdbdAna02ID, tdbdAna03ID, tdbdAna04ID, tdbdAna05ID, tdbdAna06ID, tdbdAna07ID, tdbdAna08ID, tdbdAna09ID, tdbdAna10ID, tdbg, IndexOfColumn(tdbg, COL_Ana01ID), gbUnicode)
        InputbyUnicode(Me, gbUnicode)
        SetResolutionForm(Me, ContextMenuStrip1)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Danh_sach_phieu_de_tong_hop_XDCB_-_D02F3005") & UnicodeCaption(gbUnicode) 'Danh sÀch phiÕu ¢Ó tång híp XDCB - D02F3005
        '================================================================ 
        lblPeriod.Text = rl3("Den_ky") 'Đến kỳ
        lblPeriodFrom.Text = rl3("Tu_ky") 'Từ kỳ
        lblAccountID.Text = rl3("Tai_khoan") 'Tài khoản
        lblCipID.Text = rl3("Ma_XDCB") 'Mã XDCB
        '================================================================ 
        btnFilter.Text = rl3("Loc") & " (F5)" 'Lọc
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        btnSave.Text = rl3("Don_g_y") 'Đồng ý
        '================================================================ 
        tdbcCipID.Columns("CipNo").Caption = rl3("Ma") 'Mã
        tdbcCipID.Columns("CipName").Caption = rl3("Ten") 'Tên
        tdbcAccountID.Columns("AccountID").Caption = rl3("Ma") 'Mã
        tdbcAccountID.Columns("AccountName").Caption = rl3("Ten") 'Tên
        '================================================================ 
        tdbdAna01ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna01ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna02ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna02ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna03ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna03ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna04ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna04ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna05ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna05ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna06ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna06ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna07ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna07ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna08ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna08ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna09ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna09ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna10ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna10ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        '================================================================ 
        tdbg.Columns("Choose").Caption = rl3("Chon") 'Chọn
        tdbg.Columns("VoucherTypeID").Caption = rl3("Loai_phieu") 'Loại phiếu
        tdbg.Columns("VoucherNo").Caption = rl3("So_phieu") 'Số phiếu
        tdbg.Columns("VoucherDate").Caption = rl3("Ngay_phieu") 'Ngày phiếu
        tdbg.Columns("RefDate").Caption = rl3("Ngay_hoa_don") 'Ngày hóa đơn
        tdbg.Columns("SeriNo").Caption = rl3("So_Seri") 'Số Sêri
        tdbg.Columns("RefNo").Caption = rl3("So_hoa_don") 'Số hóa đơn
        tdbg.Columns("ObjectTypeID").Caption = rl3("Loai_doi_tuong") 'Loại đối tượng
        tdbg.Columns("ObjectID").Caption = rl3("Doi_tuong") 'Đối tượng
        tdbg.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbg.Columns("DebitAccountID").Caption = rl3("Tai_khoan_no") 'Tài khoản nợ
        tdbg.Columns("CreditAccountID").Caption = rl3("Tai_khoan_co") 'Tài khoản có
        tdbg.Columns("CurrencyID").Caption = rl3("Loai_tien") 'Loại tiền
        tdbg.Columns("ExchangeRate").Caption = rl3("Ty_gia") 'Tỷ giá
        tdbg.Columns("OriginalAmount").Caption = rl3("Nguyen_te") 'Nguyên tệ
        tdbg.Columns("ConvertedAmount").Caption = rl3("Quy_doi") 'Quy đổi
    End Sub

    Private Sub setShortCut()
        mnsFind.Image = My.Resources.find
        mnsFind.Text = rl3("Tim__kiem") 'Tìm &kiếm
        mnsFind.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F), System.Windows.Forms.Keys)

        'mnuListAll
        mnsListAll.Image = My.Resources.ListAll
        mnsListAll.Text = rl3("_Liet_ke_tat_ca") '&Liệt kế tất cả
        mnsListAll.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.A), System.Windows.Forms.Keys)

    End Sub


    'Private Sub tdbg_NumberFormat()
    '    tdbg.Columns(COL_ExchangeRate).NumberFormat = D02CustomFormat.ExchangeRate
    '    tdbg.Columns(COL_OriginalAmount).NumberFormat = D02CustomFormat.D90_ConvertedDecimals
    '    tdbg.Columns(COL_ConvertedAmount).NumberFormat = D02CustomFormat.D90_ConvertedDecimals
    'End Sub

    Private Sub tdbg_NumberFormat()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbg.Columns(COL_ExchangeRate).DataField, DxxFormat.ExchangeRateDecimals, 28, 8)
        AddDecimalColumns(arr, tdbg.Columns(COL_OriginalAmount).DataField, DxxFormat.DecimalPlaces, 28, 8)
        AddDecimalColumns(arr, tdbg.Columns(COL_ConvertedAmount).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        InputNumber(tdbg, arr)
    End Sub



    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        'Load tdbcAccountID
        Dim sUnicode As String = ""
        If gbUnicode Then sUnicode = "U"
        sSQL = "Select AccountID, AccountName" & IIf(geLanguage = EnumLanguage.English, "01", "").ToString & sUnicode & "  as AccountName" & vbCrLf
        sSQL &= "From D90T0001 WITH(NOLOCK)" & vbCrLf
        sSQL &= "Where Disabled = 0 And AccountStatus = 0 AND OffAccount = 0 " & _
                " AND GroupID IN (" & IIf(_mode = 0, "'7', ", "").ToString & "'9') " & vbCrLf
        sSQL &= "Order by AccountID"
        LoadDataSource(tdbcAccountID, sSQL, gbUnicode)
        'Load tdbcCipID
        sSQL = "Select CipID, CipNo, CipName" & sUnicode & " as CipName From D02T0100 WITH(NOLOCK)" & vbCrLf
        sSQL &= "WHERE Disabled = 0 AND Status < 2"
        sSQL &= " AND DivisionID = " & SQLString(gsDivisionID) ' uppdate 31/5/2013 id 56796
        LoadDataSource(tdbcCipID, sSQL, gbUnicode)

        'Load tdbcMonthFrom
        LoadCboPeriodReport(tdbcMonthFrom, tdbcMonthTo, "D02")
        tdbcMonthFrom.SelectedValue = giTranMonth.ToString("00") & "/" & giTranYear.ToString
        tdbcMonthTo.SelectedValue = giTranMonth.ToString("00") & "/" & giTranYear.ToString
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P4002
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 11/11/2011 01:28:10
    '# Modified User: 
    '# Modified Date: 
    '# Description: Load lưới
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P4002() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P4002 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(ReturnValueC1Combo(tdbcMonthFrom, "TranMonth")) & COMMA 'FromMonth, int, NOT NULL
        sSQL &= SQLNumber(ReturnValueC1Combo(tdbcMonthFrom, "TranYear")) & COMMA 'FromYear, int, NOT NULL
        sSQL &= SQLNumber(ReturnValueC1Combo(tdbcMonthTo, "TranMonth")) & COMMA 'ToMonth, int, NOT NULL
        sSQL &= SQLNumber(ReturnValueC1Combo(tdbcMonthTo, "TranYear")) & COMMA 'ToYear, int, NOT NULL
        sSQL &= SQLString(tdbcAccountID.Text) & COMMA 'AccountID, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'strFind, varchar[8000], NOT NULL
        sSQL &= SQLNumber(_mode) & COMMA 'Mode, tinyint, NOT NULL
        sSQL &= SQLString(tdbcCipID.SelectedValue) & COMMA 'CipID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode)
        Return sSQL
    End Function

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

#Region "Events tdbcMonthFrom"

    Private Sub tdbcMonthFrom_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcMonthFrom.LostFocus
        If tdbcMonthFrom.FindStringExact(tdbcMonthFrom.Text) = -1 Then tdbcMonthFrom.Text = ""
    End Sub

#End Region

#Region "Events tdbcMonthTo"

    Private Sub tdbcMonthTo_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcMonthTo.LostFocus
        If tdbcMonthTo.FindStringExact(tdbcMonthTo.Text) = -1 Then tdbcMonthTo.Text = ""
    End Sub

#End Region

#Region "Events tdbcAccountID with txtAccountName"

    Private Sub tdbcAccountID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAccountID.SelectedValueChanged
        If tdbcAccountID.SelectedValue Is Nothing Then
            txtAccountName.Text = ""
        Else
            txtAccountName.Text = tdbcAccountID.Columns(1).Value.ToString
        End If
    End Sub

    Private Sub tdbcAccountID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAccountID.LostFocus
        If tdbcAccountID.FindStringExact(tdbcAccountID.Text) = -1 Then
            tdbcAccountID.Text = ""
        End If
    End Sub

#End Region

#Region "Events tdbcCipID with txtCipName"

    Private Sub tdbcCipID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCipID.SelectedValueChanged
        If tdbcCipID.SelectedValue Is Nothing Then
            txtCipName.Text = ""
        Else
            txtCipName.Text = tdbcCipID.Columns(2).Value.ToString
        End If
    End Sub

    Private Sub tdbcCipID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCipID.LostFocus
        If tdbcCipID.FindStringExact(tdbcCipID.Text) = -1 Then
            tdbcCipID.Text = ""
            txtCipName.Text = ""
        End If
    End Sub

#End Region

#Region "Active Find Client - List All "
    Private WithEvents Finder As New D99C1001
    Private sFind As String = ""
    Dim dtCaptionCols As DataTable

    Private Sub tsbFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnsFind.Click
        gbEnabledUseFind = True
        '*****************************************
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbg.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        Dim Arr As New ArrayList
        For i As Integer = 0 To tdbg.Splits.Count - 1
            AddColVisible(tdbg, i, Arr, , , , gbUnicode)
        Next
        'Tạo tableCaption: đưa tất cả các cột trên lưới có Visible = True vào table 
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
        'End If

        ShowFindDialogClient(Finder, dtCaptionCols, Me.Name, "0", gbUnicode)

    End Sub

    Private Sub Finder_FindClick(ByVal ResultWhereClause As Object) Handles Finder.FindClick
        If ResultWhereClause Is Nothing Or ResultWhereClause.ToString = "" Then Exit Sub
        sFind = ResultWhereClause.ToString()
        ReLoadTDBGrid()
    End Sub

    Private Sub tsbListAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnsListAll.Click
        sFind = ""
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        ReLoadTDBGrid()
    End Sub

    Private Sub ReLoadTDBGrid()
        Dim strFind As String = sFind
        If sFilter.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter.ToString

        dtGrid.DefaultView.RowFilter = strFind
        ResetGrid()
    End Sub

    Private Sub ResetGrid()
        mnsFind.Enabled = gbEnabledUseFind OrElse tdbg.RowCount > 0

        mnsListAll.Enabled = mnsFind.Enabled
        FooterTotalGrid(tdbg, COL_VoucherNo)
        FooterSumNew(tdbg, COL_OriginalAmount, COL_ConvertedAmount)
    End Sub

#End Region

    Dim sFilter As New System.Text.StringBuilder()
    Dim bRefreshFilter As Boolean = False 'Cờ bật set FilterText =""
    Private Sub tdbg_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.FilterChange
        Try
            If (dtGrid Is Nothing) Then Exit Sub
            If bRefreshFilter Then Exit Sub 'set FilterText ="" thì thoát
            'Filter the data 
            FilterChangeGrid(tdbg, sFilter)
            ReLoadTDBGrid()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

    Dim bSelected As Boolean = False
    Private Sub HeadClick(ByVal iCol As Integer)
        Select Case iCol
            Case IndexOfColumn(tdbg, COL_Choose)
                tdbg.AllowSort = False
                L3HeadClick(tdbg, iCol, bSelected)
            Case IndexOfColumn(tdbg, COL_Ana01ID) To IndexOfColumn(tdbg, COL_Ana10ID)
                tdbg.AllowSort = False
                CopyColumns(tdbg, iCol, tdbg.Columns(iCol).Text, tdbg.Row)
            Case Else
                tdbg.AllowSort = True
        End Select
    End Sub

    Private Sub tdbg_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.HeadClick
        HeadClick(e.ColIndex)
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        If e.Control And e.KeyCode = Keys.S Then
            HeadClick(tdbg.Col)
            Exit Sub
        End If
        HotKeyCtrlVOnGrid(tdbg, e) 'Nhấn Ctrl + V trên lưới 'có trong D99X0000
    End Sub

    Private Sub tdbg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        Select Case tdbg.Columns(tdbg.Col).DataField
            Case COL_Choose  'Chặn Ctrl + V trên cột Check
                e.Handled = CheckKeyPress(e.KeyChar)
            Case COL_ConvertedAmount, COL_OriginalAmount, COL_ExchangeRate
                'e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
            Case COL_Ana01ID, COL_Ana02ID, COL_Ana03ID, COL_Ana04ID, COL_Ana05ID, COL_Ana06ID, COL_Ana07ID, COL_Ana08ID, COL_Ana09ID, COL_Ana10ID
                If tdbg.FilterActive Then e.Handled = True
        End Select
    End Sub

    Private Sub tdbg_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg.BeforeColUpdate
        '--- Kiểm tra giá trị hợp lệ
        Select Case e.ColIndex
            Case IndexOfColumn(tdbg, COL_Ana01ID) To IndexOfColumn(tdbg, COL_Ana10ID)
                If tdbg.Columns(e.ColIndex).Text <> tdbg.Columns(e.ColIndex).DropDown.Columns(0).Text Then
                    tdbg.Columns(e.ColIndex).Text = ""
                End If
        End Select
    End Sub

    Dim dtGrid As DataTable

    Private Sub LoadTDBGrid()
        dtGrid = ReturnDataTable(SQLStoreD02P4002())
        'dtGrid.Columns.Add("Choose", Type.GetType("System.Boolean"))
        gbEnabledUseFind = dtGrid.Rows.Count > 0
        LoadDataSource(tdbg, dtGrid, gbUnicode)

        ResetGrid()
    End Sub

    Private Function AllowFilter() As Boolean
        If tdbcMonthFrom.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Ky"))
            tdbcMonthFrom.Focus()
            Return False
        End If
        If tdbcMonthTo.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Ky"))
            tdbcMonthTo.Focus()
            Return False
        End If
        If Not CheckValidPeriodFromTo(tdbcMonthFrom, tdbcMonthTo) Then Return False

        If tdbcAccountID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Tai_khoan"))
            tdbcAccountID.Focus()
            Return False
        End If
        'If tdbcCipID.Text.Trim = "" Then
        '    D99C0008.MsgNotYetChoose(rl3("Ma_XDCB"))
        '    tdbcCipID.Focus()
        '    Return False
        'End If
        Return True
    End Function

    Private Sub SetBackColorObligatory()
        tdbcMonthFrom.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcMonthTo.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        'tdbcCipID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcAccountID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        If Not AllowFilter() Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        btnFilter.Enabled = False
        sFind = ""
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        LoadTDBGrid()

        btnFilter.Enabled = True
        Me.Cursor = Cursors.Default
    End Sub

    Private Function AllowSave() As Boolean
        If dtGrid Is Nothing Then
            tdbg.Focus()
            D99C0008.MsgNoDataInGrid()
            Return False
        Else
            Dim dr() As DataRow = dtGrid.Select(COL_Choose & "=1")
            If dr.Length <= 0 Then
                tdbg.Focus()
                D99C0008.MsgNoDataInGrid()
                tdbg.SplitIndex = 0
                tdbg.Col = IndexOfColumn(tdbg, COL_Choose)
                tdbg.Row = 0
                Return False
            End If
        End If
        Return True
    End Function


    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Not AllowSave() Then Exit Sub
        _bSavedOK = True
        dtGrid.DefaultView.RowFilter = COL_Choose & "=1"
        _dtChose = dtGrid.DefaultView.ToTable
        Me.Close()
    End Sub

 
End Class