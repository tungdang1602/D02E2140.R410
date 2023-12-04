'#-------------------------------------------------------------------------------------
'# Created Date: 03/10/2007 1:56:42 PM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 03/10/2007 1:56:42 PM
'# Modify User: Trần Thị ÁiTrâm
'#-------------------------------------------------------------------------------------
Imports System.Text

Public Class D02F1005

#Region "Const of tdbg"
    Private Const COL_Selected As Integer = 0         ' Chọn
    Private Const COL_VoucherTypeID As Integer = 1    ' Loại phiếu
    Private Const COL_VoucherNo As Integer = 2        ' Số phiếu
    Private Const COL_VoucherDate As Integer = 3      ' Ngày phiếu
    Private Const COL_SeriNo As Integer = 4           ' Số Sêri
    Private Const COL_RefNo As Integer = 5            ' Số hóa đơn
    Private Const COL_RefDate As Integer = 6          ' Ngày hóa đơn
    Private Const COL_ObjectTypeID As Integer = 7     ' Loại hóa đơn
    Private Const COL_ObjectID As Integer = 8         ' Mã đối tượng
    Private Const COL_Description As Integer = 9      ' Diễn giải
    Private Const COL_DebitAccountID As Integer = 10  ' TK nợ
    Private Const COL_CreditAccountID As Integer = 11 ' TK có
    Private Const COL_CurrencyID As Integer = 12      ' Loại tiền
    Private Const COL_ExchangeRate As Integer = 13    ' Tỷ giá
    Private Const COL_OriginalAmount As Integer = 14  ' Nguyên tệ
    Private Const COL_ConvertedAmount As Integer = 15 ' Qui đổi
    Private Const COL_CipID As Integer = 16           ' CipID
    Private Const COL_TransactionID As Integer = 17   ' TransactionID
    Private Const COL_ModuleID As Integer = 18        ' ModuleID
    Private Const COL_Ana01ID As Integer = 19         ' Khoản mục 01
    Private Const COL_Ana02ID As Integer = 20         ' Khoản mục 02
    Private Const COL_Ana03ID As Integer = 21         ' Khoản mục 03
    Private Const COL_Ana04ID As Integer = 22         ' Khoản mục 04
    Private Const COL_Ana05ID As Integer = 23         ' Khoản mục 05
    Private Const COL_Ana06ID As Integer = 24         ' Khoản mục 06
    Private Const COL_Ana07ID As Integer = 25         ' Khoản mục 07
    Private Const COL_Ana08ID As Integer = 26         ' Khoản mục 08
    Private Const COL_Ana09ID As Integer = 27         ' Khoản mục 09
    Private Const COL_Ana10ID As Integer = 28         ' Khoản mục 10
#End Region

    Private _accountID As String
    Private _status As Integer
    Private _cipID As String
    Private dtGrid As DataTable
    Dim iLastcol As Integer
    Dim bHeadClick As Boolean = False

    '---Kiểm tra khoản mục theo chuẩn gồm 6 bước
    '--- Chuẩn Khoản mục b1: Khai báo biến

#Region "Biến khai báo cho khoản mục"

    Private Const SplitAna As Int16 = 2 ' Ghi nhận Khoản mục chứa ở Split nào
    Dim bUseAna As Boolean 'Kiểm tra có sử dụng Khoản mục không, để set thuộc tính Enabled nút Khoản mục 
    'Dim iDisplayAnaCol As Integer = 0 ' Cột Khoản mục đầu tiên được hiển thị, khi nhấn nút Khoản mục thì Focus đến cột đó
    'Dim xCheckAna(9) As Boolean 'Khởi động tại Form_load: Ghi lại việc kiểm tra lần đầu Lưu, khi nhấn Lưu lần thứ 2 thì không cần kiểm tra nữa

#End Region

    Public Property CipID() As String
        Get
            Return _cipID
        End Get
        Set(ByVal value As String)
            If CipID = value Then
                _cipID = ""
                Return
            End If
            _cipID = value
        End Set
    End Property

    Public Property AccountID() As String
        Get
            Return _accountID
        End Get
        Set(ByVal value As String)
            If AccountID = value Then
                _accountID = ""
                Return
            End If
            _accountID = value
        End Set
    End Property

    Public Property Status() As Integer
        Get
            Return _status
        End Get
        Set(ByVal value As Integer)
            _status = value
        End Set
    End Property

    Private Sub D02F1005_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
        If e.Control Then
            Select Case e.KeyCode
                Case Keys.F
                    If mnuFind.Enabled Then
                        mnuFind_Click(Nothing, Nothing)
                    End If
                Case Keys.A
                    If mnuListAll.Enabled Then
                        mnuListAll_Click(Nothing, Nothing)
                    End If
            End Select
        End If
    End Sub

    Private Sub D02F1005_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        SetShortcutPopupMenu(C1CommandHolder)
        '--- Chuẩn Khoản mục b2: Lấy caption cho 10 khoản mục
        bUseAna = LoadTDBGridAnalysisCaption(D02, tdbg, COL_Ana01ID, SplitAna, True, gbUnicode)
        'SetNewXaCheckAna()
        'D91 có sử dụng Khoản mục
        'If bUseAna Then iDisplayAnaCol = 1
        If Not bUseAna Then tdbg.Splits(SplitAna).SplitSize = 0

        '------------------------------------
        Loadlanguage()
        ResetSplitDividerSize(tdbg)
        'LoadTDBGridAnalysisCaption(D02, tdbg, COL_Ana01ID, SPLIT2, True)
        LoadTDBDropDown()
        gbEnabledUseFind = False
        ResetColorGrid(tdbg, 1)
        LoadTDBGrid()
        tdbg_LockedColumns()
        tdbg_NumberFormat()
        btnCollection.Enabled = ReturnPermission("D02F1003") > EnumPermission.View
        iLastcol = CountCol(tdbg, 1)
        
    SetResolutionForm(Me)
Me.Cursor = Cursors.Default
End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub LoadTDBDropDown()
        '--- Chuẩn Khoản mục b3: Load 10 khoản mục
        LoadTDBDropDownAna(tdbdAna01ID, tdbdAna02ID, tdbdAna03ID, tdbdAna04ID, tdbdAna05ID, tdbdAna06ID, tdbdAna07ID, tdbdAna08ID, tdbdAna09ID, tdbdAna10ID, tdbg, COL_Ana01ID, gbUnicode)
        '------------------------------------------
    End Sub

    Private Sub tdbg_LockedColumns()
        tdbg.Splits(SPLIT1).DisplayColumns(COL_Selected).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_VoucherTypeID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_VoucherNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_VoucherDate).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_RefDate).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_SeriNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_RefNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_ObjectTypeID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_ObjectID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_Description).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_DebitAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_CreditAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_CurrencyID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_ExchangeRate).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_OriginalAmount).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_ConvertedAmount).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_CipID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_TransactionID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT1).DisplayColumns(COL_ModuleID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
    End Sub

    'Private Sub tdbg_NumberFormat()
    '    tdbg.Columns(COL_ExchangeRate).NumberFormat = DxxFormat.ExchangeRateDecimals
    '    tdbg.Columns(COL_OriginalAmount).NumberFormat = DxxFormat.DecimalPlaces
    '    tdbg.Columns(COL_ConvertedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
    'End Sub

    Private Sub tdbg_NumberFormat()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbg.Columns(COL_ExchangeRate).DataField, DxxFormat.ExchangeRateDecimals, 28, 8)
        AddDecimalColumns(arr, tdbg.Columns(COL_OriginalAmount).DataField, DxxFormat.DecimalPlaces, 28, 8)
        AddDecimalColumns(arr, tdbg.Columns(COL_ConvertedAmount).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        InputNumber(tdbg, arr)
    End Sub



#Region "Active Find Client - List All "
    Private WithEvents Finder As New D99C1001
    Private sFind As String = ""
    Dim dtCaptionCols As DataTable
    Private Sub mnuFind_Click(ByVal sender As Object, ByVal e As C1.Win.C1Command.ClickEventArgs) Handles mnuFind.Click
        'If CallMenuFromGrid(tdbg, e) = False Then Exit Sub
        'Dim sSQL As String = ""
        'gbEnabledUseFind = True
        'sSQL = "Select * From D02V1234 "
        'sSQL &= "Where FormID = " & SQLString(Me.Name) & "And Language = " & SQLString(gsLanguage)
        'ShowFindDialogClient(Finder, sSQL)
        gbEnabledUseFind = True
        '*****************************************
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbg.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        Dim Arr As New ArrayList
        AddColVisible(tdbg, SPLIT0, Arr, , , , gbUnicode)
        AddColVisible(tdbg, SPLIT1, Arr, , , , gbUnicode)
        'Tạo tableCaption: đưa tất cả các cột trên lưới có Visible = True vào table 
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
        'End If

        ShowFindDialogClient(Finder, dtCaptionCols, Me.Name, "0", gbUnicode)
        '*****************************************
    End Sub

    Private Sub Finder_FindClick(ByVal ResultWhereClause As Object) Handles Finder.FindClick
        If ResultWhereClause Is Nothing Then Exit Sub
        sFind = ResultWhereClause.ToString()
        ReLoadTDBGrid()
    End Sub

    Private Sub mnuListAll_Click(ByVal sender As Object, ByVal e As C1.Win.C1Command.ClickEventArgs) Handles mnuListAll.Click
        If CallMenuFromGrid(tdbg, e) = False Then Exit Sub
        sFind = ""
        ReLoadTDBGrid()
    End Sub

    Private Sub ReLoadTDBGrid()
        Dim strFind As String = sFind
        '  If sFilter.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        '  strFind &= sFilter.ToString
        dtGrid.DefaultView.RowFilter = strFind
        CheckMenuOther(gbEnabledUseFind)
    End Sub

#End Region

    Private Sub LoadTDBGrid()
        Dim sSQL As String
        sSQL = SQLStoreD02P0400()
        dtGrid = ReturnDataTable(sSQL)
        LoadDataSource(tdbg, dtGrid, gbUnicode)
        CheckMenuOther(gbEnabledUseFind)
    End Sub
    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0400
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 09/10/2007 04:11:32
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0400() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0400 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(_accountID) & COMMA 'CollectAccountID, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'strFind, varchar[8000], NOT NULL
        sSQL &= SQLNumber(gbUnicode)
        Return sSQL
    End Function

    Public Sub CheckMenuOther(ByVal UsedFind As Boolean)
        mnuFind.Enabled = UsedFind Or tdbg.RowCount > 0
        mnuListAll.Enabled = UsedFind Or tdbg.RowCount > 0
    End Sub

    Private Sub btnCollection_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCollection.Click
        tdbg.UpdateData()
        If Not AllowSave() Then Exit Sub
        btnCollection.Enabled = False
        btnClose.Enabled = False
        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        sSQL.Append(SQLUpdateD02T0012s)
        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default
        gbSavedOK = False
        If bRunSQL Then
            'SaveOK()
            gbSavedOK = True
            btnCollection.Enabled = True
            btnClose.Enabled = True
            btnClose.Focus()
            Me.Close()
        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnCollection.Enabled = True
        End If
        If _status = 0 Then
            sSQL.Append(SQLUpdateD02T0100())
            ExecuteSQL(sSQL.ToString)
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0012s
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 09/10/2007 04:19:28
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0012s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        For i As Integer = 0 To tdbg.RowCount - 1
            If Not IsDBNull(tdbg(i, COL_Selected)) And tdbg(i, COL_Selected).ToString <> "" Then
                If CBool(tdbg(i, COL_Selected)) = True Then
                    sSQL.Append("Update D02T0012 Set ")
                    '  sSQL.Append("Status = " & "1" & COMMA) 'tinyint, NOT NULL
                    sSQL.Append("Ana01ID = " & SQLString(tdbg(i, COL_Ana01ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana02ID = " & SQLString(tdbg(i, COL_Ana02ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana03ID = " & SQLString(tdbg(i, COL_Ana03ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana04ID = " & SQLString(tdbg(i, COL_Ana04ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana05ID = " & SQLString(tdbg(i, COL_Ana05ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana06ID = " & SQLString(tdbg(i, COL_Ana06ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana07ID = " & SQLString(tdbg(i, COL_Ana07ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana08ID = " & SQLString(tdbg(i, COL_Ana08ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana09ID = " & SQLString(tdbg(i, COL_Ana09ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana10ID = " & SQLString(tdbg(i, COL_Ana10ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("CipID = " & SQLString(_cipID)) 'varchar[20], NULL
                    sSQL.Append(" Where ")
                    sSQL.Append("TransactionID = " & SQLString(tdbg(i, COL_TransactionID)) & " And ")
                    sSQL.Append("DivisionID = " & SQLString(gsDivisionID))
                    sRet.Append(sSQL.ToString & vbCrLf)
                    sSQL.Remove(0, sSQL.Length)
                End If
            End If
            
        Next
        Return sRet
    End Function

    Private Function AllowSave() As Boolean
        Dim iCount As Integer = 0
        If tdbg.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tdbg.Focus()
            Return False
        End If
        For i As Integer = 0 To tdbg.RowCount - 1
            If Not IsDBNull(tdbg(i, COL_Selected)) And tdbg(i, COL_Selected).ToString <> "" Then
                If CBool(tdbg(i, COL_Selected)) = False Then
                    iCount += 1
                End If
            End If
        Next
        If iCount = tdbg.RowCount Then
            D99C0008.MsgL3(rl3("Ban_chua_chon_dong_nao_de_tap_hop"))
            tdbg.SplitIndex = SPLIT0
            tdbg.Col = COL_Selected
            tdbg.Bookmark = 0
            tdbg.Focus()
            Return False
        End If
        Return True
    End Function

    Private Sub tdbg_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.HeadClick
        Select Case tdbg.Col
            Case COL_Selected
                CheckedAll()
        End Select
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        If e.KeyCode = Keys.Enter Then
            If tdbg.Col = iLastcol Then
                HotKeyEnterGrid(tdbg, COL_Selected, e)
            End If
        End If
        If e.Control And e.KeyCode = Keys.S Then
            tdbg_HeadClick(Nothing, Nothing)
        End If
    End Sub

    Private Sub CheckedAll()
        bHeadClick = Not bHeadClick
        For i As Integer = 0 To tdbg.RowCount - 1
            tdbg(i, COL_Selected) = bHeadClick
        Next

    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Danh_sach_phieu_de_tap_hop_XDCB_-_D02F1005") & UnicodeCaption(gbUnicode) 'Danh sÀch phiÕu ¢Ó tËp híp XDCB - D02F1005
        '================================================================ 
        btnCollection.Text = rl3("Tap__hop") 'Tập &hợp
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        '================================================================ 
        tdbdAna01ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna01ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna02ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna02ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna03ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna03ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna04ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna04ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna05ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna05ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna06ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna06ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna07ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna07ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna08ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna08ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna09ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna09ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna10ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna10ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        '================================================================ 
        tdbg.Columns("Selected").Caption = rl3("Chon") 'Chọn
        tdbg.Columns("VoucherTypeID").Caption = rl3("Loai_phieu") 'Loại phiếu
        tdbg.Columns("VoucherNo").Caption = rl3("So_phieu") 'Số phiếu
        tdbg.Columns("VoucherDate").Caption = rl3("Ngay_phieu") 'Ngày phiếu
        tdbg.Columns("RefDate").Caption = rl3("Ngay_hoa_don") 'Ngày hóa đơn
        tdbg.Columns("SeriNo").Caption = rl3("So_Seri") 'Số Sêri
        tdbg.Columns("RefNo").Caption = rl3("So_hoa_don") 'Số hóa đơn
        tdbg.Columns("ObjectTypeID").Caption = rl3("Loai_hoa_don") 'Loại hóa đơn
        tdbg.Columns("ObjectID").Caption = rl3("Ma_doi_tuong") 'Mã đối tượng
        tdbg.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbg.Columns("DebitAccountID").Caption = rl3("TK_no") 'rl3("Tai_khoan_no") 'Tài khoản nợ
        tdbg.Columns("CreditAccountID").Caption = rl3("TK_co") 'rl3("Tai_khoan_co") 'Tài khoản có
        tdbg.Columns("CurrencyID").Caption = rl3("Loai_tien") 'Loại tiền
        tdbg.Columns("ExchangeRate").Caption = rl3("Ty_gia") 'Tỷ giá
        tdbg.Columns("OriginalAmount").Caption = rl3("Nguyen_te") 'Nguyên tệ
        tdbg.Columns("ConvertedAmount").Caption = rl3("Quy_doi") 'Qui đổi

        '================================================================ 
        mnuFind.Text = rl3("Tim__kiem") 'Tìm &kiếm
        mnuListAll.Text = rl3("_Liet_ke_tat_ca") '&Liệt kê tất cả
    End Sub

    Private Sub tdbg_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg.BeforeColUpdate
        Select Case e.ColIndex
            Case COL_Selected
                'Case COL_ExchangeRate
                '    If Not IsNumeric(tdbg.Columns(COL_ExchangeRate).Text) Then e.Cancel = True
                'Case COL_OriginalAmount
                '    If Not IsNumeric(tdbg.Columns(COL_OriginalAmount).Text) Then e.Cancel = True
                'Case COL_ConvertedAmount
                '    If Not IsNumeric(tdbg.Columns(COL_ConvertedAmount).Text) Then e.Cancel = True
            Case COL_CipID
            Case COL_TransactionID
            Case COL_ModuleID
                '--- Chuẩn Khoản mục b5: Kiểm tra Khoản mục lúc nhập liệu
                '---------------------------------------------
            Case COL_Ana01ID
                If tdbg.Columns(COL_Ana01ID).Text <> tdbdAna01ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(0) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana01ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana01ID).Text.Length > giArrAnaLength(0) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana01ID).Text = ""
                        End If
                    End If
                End If

            Case COL_Ana02ID
                If tdbg.Columns(COL_Ana02ID).Text <> tdbdAna02ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(1) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana02ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana02ID).Text.Length > giArrAnaLength(1) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana02ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana03ID
                If tdbg.Columns(COL_Ana03ID).Text <> tdbdAna03ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(2) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana03ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana03ID).Text.Length > giArrAnaLength(2) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana03ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana04ID
                If tdbg.Columns(COL_Ana04ID).Text <> tdbdAna04ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(3) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana04ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana04ID).Text.Length > giArrAnaLength(3) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana04ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana05ID
                If tdbg.Columns(COL_Ana05ID).Text <> tdbdAna05ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(4) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana05ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana05ID).Text.Length > giArrAnaLength(4) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana05ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana06ID
                If tdbg.Columns(COL_Ana06ID).Text <> tdbdAna06ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(5) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana06ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana06ID).Text.Length > giArrAnaLength(5) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana06ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana07ID
                If tdbg.Columns(COL_Ana07ID).Text <> tdbdAna07ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(6) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana07ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana07ID).Text.Length > giArrAnaLength(6) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana07ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana08ID
                If tdbg.Columns(COL_Ana08ID).Text <> tdbdAna08ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(7) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana08ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana08ID).Text.Length > giArrAnaLength(7) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana08ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana09ID
                If tdbg.Columns(COL_Ana09ID).Text <> tdbdAna09ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(8) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana09ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana09ID).Text.Length > giArrAnaLength(8) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana09ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana10ID
                If tdbg.Columns(COL_Ana10ID).Text <> tdbdAna10ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(9) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana10ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana10ID).Text.Length > giArrAnaLength(9) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana10ID).Text = ""
                        End If
                    End If
                End If
                '---------------------------------------------
        End Select
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0100
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 14/11/2006 04:04:24
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0100() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0100 Set ")
        sSQL.Append("Status = " & SQLNumber(1)) 'tinyint, NULL
        sSQL.Append(" Where ")
        sSQL.Append("CipID = " & SQLString(_cipID))
        Return sSQL
    End Function

End Class