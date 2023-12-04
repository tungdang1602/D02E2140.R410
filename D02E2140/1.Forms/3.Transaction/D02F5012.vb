Imports System.Drawing
Imports System
Public Class D02F5012

#Region "Const of tdbg1 - Total of Columns: 34"
    Private Const COL_TransactionID As Integer = 0           ' TransactionID
    Private Const COL_VoucherNo As Integer = 1               ' VoucherNo
    Private Const COL_BatchID As Integer = 2                 ' BatchID
    Private Const COL_AssetID As Integer = 3                 ' Mã tài sản
    Private Const COL_AssetName As Integer = 4               ' Tên tài sản
    Private Const COL_ToDepAccountID As Integer = 5          ' Tài khoản KH
    Private Const COL_DepRate As Integer = 6                 ' Tỉ lệ
    Private Const COL_DepAmount As Integer = 7               ' Mức khấu hao
    Private Const COL_CurrentCost As Integer = 8             ' Nguyên giá
    Private Const COL_LastDepAmount As Integer = 9           ' Mức KH kỳ trước
    Private Const COL_CurrentLTDDepreciation As Integer = 10 ' Lũy kế KH
    Private Const COL_DepreciatedPeriod As Integer = 11      ' Số kỳ đã KH
    Private Const COL_LastPeriod As Integer = 12             ' Số kỳ còn lại
    Private Const COL_DepreciationDayNum As Integer = 13     ' Số ngày tính khấu hao
    Private Const COL_MethodID As Integer = 14               ' MethodID
    Private Const COL_MethodName As Integer = 15             ' Phương pháp KH
    Private Const COL_Ana01ID As Integer = 16                ' KM 1
    Private Const COL_Ana02ID As Integer = 17                ' KM 2
    Private Const COL_Ana03ID As Integer = 18                ' KM 3
    Private Const COL_Ana04ID As Integer = 19                ' KM 4
    Private Const COL_Ana05ID As Integer = 20                ' KM 5
    Private Const COL_Ana06ID As Integer = 21                ' KM 6
    Private Const COL_Ana07ID As Integer = 22                ' KM 7
    Private Const COL_Ana08ID As Integer = 23                ' KM 8
    Private Const COL_Ana09ID As Integer = 24                ' KM 9
    Private Const COL_Ana10ID As Integer = 25                ' KM 10
    Private Const COL_ProjectID As Integer = 26              ' Dự án
    Private Const COL_ProjectName As Integer = 27            ' Tên dự án
    Private Const COL_TaskID As Integer = 28                 ' Hạng mục
    Private Const COL_TaskName As Integer = 29               ' Tên hạng mục
    Private Const COL_BudgetID As Integer = 30               ' Ngân sách
    Private Const COL_BudgetName As Integer = 31             ' Tên ngân sách
    Private Const COL_BudgetItemID As Integer = 32           ' Hạng mục ngân sách
    Private Const COL_BudgetItemName As Integer = 33         ' Tên hạng mục ngân sách
#End Region

#Region "Const of tdbg2"
    Private Const COL2_AssetID As Integer = 0   ' Mã tài sản
    Private Const COL2_AssetName As Integer = 1 ' Tên tài sản
#End Region

    Private _bSaveOk As Boolean = False
    Public WriteOnly Property  bSaveOk() As Boolean 
        Set(ByVal Value As Boolean )
            _bSaveOk = Value
        End Set
    End Property

    Private _voucherNo As String = ""
    Public WriteOnly Property  VoucherNo() As String 
        Set(ByVal Value As String )
            _voucherNo = Value
        End Set
    End Property

    Private _batchID As String = ""
    Public WriteOnly Property  BatchID() As String 
        Set(ByVal Value As String )
            _batchID = Value
        End Set
    End Property

    Private iColumns() As Integer = {COL_DepAmount, COL_CurrentCost, COL_LastDepAmount, COL_CurrentLTDDepreciation}
    Private dt, dtProjectID, dtTaskID, dtBudgetID, dtBudgetItemID As DataTable
    Private myTabRect As Rectangle
    Dim bUseAna As Boolean
    Private usrOption As New D99U1111()
    Dim dtF12 As DataTable

    Private Sub D02F5012_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If usrOption IsNot Nothing Then usrOption.Dispose()
        If Not _bSaveOk Then
            ExecuteSQL(SQLStoreD02P0013)
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0013
    '# Created User: HUỲNH KHANH
    '# Created Date: 21/12/2015 11:03:55
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0013() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Store xoa phieu rac neu khong luu" & vbCrlf)
        sSQL &= "Exec D02P0013 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostName, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(_voucherNo) & COMMA 'VoucherNo, varchar[20], NOT NULL
        sSQL &= SQLString(_batchID) 'VoucherIGE, varchar[20], NOT NULL
        Return sSQL
    End Function



    Private Sub btnF12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnF12.Click
        If usrOption Is Nothing Then Exit Sub 'TH lưới không có cột
        usrOption.Location = New Point(tdbg1.Left, btnF12.Top - (usrOption.Height + 7))
        Me.Controls.Add(usrOption)
        usrOption.BringToFront()
        usrOption.Visible = True
    End Sub

    Private Sub CallD99U1111()
        Dim arrColObligatory() As Object = {COL_AssetID}
        usrOption.AddColVisible(tdbg1, dtF12, arrColObligatory)
        If usrOption IsNot Nothing Then usrOption.Dispose()
        usrOption = New D99U1111(Me, tdbg1, dtF12)
    End Sub

    Private Sub D02F5012_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '12/10/2020, id 144622-Tài sản cố định_Lỗi chưa cảnh báo khi lưu
        If (tdbg1.RowCount > 0 Or tdbg2.RowCount > 0) Then
            If Not _bSaveOk Then
                If Not AskMsgBeforeClose() Then e.Cancel = True : Exit Sub
            End If
        End If
    End Sub

    Private Sub D02F5012_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
        If e.Alt Then
            Select Case e.KeyCode
                Case Keys.NumPad1, Keys.D1
                    tabMain.SelectedIndex = 0
                    tdbg1.Focus()
                    Exit Sub
                Case Keys.NumPad2, Keys.D2
                    tabMain.SelectedIndex = 1
                    tdbg2.Focus()
                    Exit Sub
                Case Keys.F12
                    btnF12_Click(Nothing, Nothing)
                Case Keys.Escape
                    usrOption.picClose_Click(Nothing, Nothing)
            End Select
        End If
    End Sub

    Private Sub D02F5012_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        SetShortcutPopupMenu(C1CommandHolder)
        ResetColorGrid(tdbg1)
        ResetColorGrid(tdbg2)
        gbEnabledUseFind = False
        tdbg1_NumberFormat()
        Loadlanguage()
        bUseAna = LoadTDBGridAnalysisCaption(D02, tdbg1, COL_Ana01ID, SPLIT0, True, gbUnicode)
        Load10TDBDropDownAna() '31/8/2021, Nguyễn Thị Mỹ Lài:id 180303-Cho phép sửa khoản mục Kcode khi tính KH TSCD (màn hình D2F5012) Hùng Vương
        LoadData()
        CheckMenu(Me.Name, C1CommandHolder, tdbg1.RowCount, gbEnabledUseFind, False)
        If tdbg2.RowCount < 1 Then
            tabMain.DrawMode = TabDrawMode.OwnerDrawFixed
            TabPage2.Enabled = False
            AddHandler tabMain.DrawItem, AddressOf OnDrawItem
            ' Sizes the tabs of tabControl1.
            Me.tabMain.ItemSize = New Size(145, 18)
            ' Makes the tab width definable. 
            Me.tabMain.SizeMode = TabSizeMode.Fixed
        End If

        dtF12 = Nothing
        CallD99U1111()
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub


    Private Sub OnDrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)
        ' Create pen.
        Dim blackPen As New Pen(tabMain.TabPages(0).BackColor, 3)
        'Get Location tabpage
        myTabRect = tabMain.GetTabRect(tabMain.SelectedIndex)
        ' Create coordinates of points that define line.
        Dim x1 As Integer = myTabRect.X
        Dim y1 As Integer = myTabRect.Bottom
        Dim x2 As Integer = myTabRect.X + myTabRect.Width
        ' Draw line to screen.
        e.Graphics.DrawLine(blackPen, x1, y1, x2, y1)
        '**************
        ' Set format of string.
        Dim drawFormat As New StringFormat
        drawFormat.LineAlignment = StringAlignment.Center
        Dim page As TabPage = tabMain.TabPages(e.Index)
        If Not page.Enabled Then
            Dim brush As New SolidBrush(SystemColors.GrayText)
            e.Graphics.DrawString(page.Text, page.Font, brush, e.Bounds, drawFormat)
        Else
            Dim brush As New SolidBrush(page.ForeColor)
            e.Graphics.DrawString(page.Text, page.Font, brush, e.Bounds, drawFormat)
        End If
    End Sub

    Private Sub tabMain_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles tabMain.Selecting
        If e.TabPage.Enabled = False Then
            e.Cancel = True
        Else
            e.Cancel = False
        End If
    End Sub

    'Private Sub tdbg1_NumberFormat()
    '    tdbg1.Columns(COL_DepRate).NumberFormat = DxxFormat.DefaultNumber2
    '    tdbg1.Columns(COL_DepAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
    '    tdbg1.Columns(COL_CurrentCost).NumberFormat = DxxFormat.D90_ConvertedDecimals
    '    tdbg1.Columns(COL_LastDepAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
    '    tdbg1.Columns(COL_CurrentLTDDepreciation).NumberFormat = DxxFormat.D90_ConvertedDecimals
    'End Sub

    Private Sub tdbg1_NumberFormat()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbg1.Columns(COL_DepRate).DataField, DxxFormat.DefaultNumber2, 28, 8)
        AddDecimalColumns(arr, tdbg1.Columns(COL_DepAmount).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        AddDecimalColumns(arr, tdbg1.Columns(COL_CurrentCost).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        AddDecimalColumns(arr, tdbg1.Columns(COL_LastDepAmount).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        AddDecimalColumns(arr, tdbg1.Columns(COL_CurrentLTDDepreciation).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        AddDecimalColumns(arr, tdbg1.Columns(COL_DepreciationDayNum).DataField, DxxFormat.DefaultNumber0, 28, 8) '18/12/2019, Lê Thị Thu Thảo:id 126368-PAN - Phát triển phương pháp tính khấu hao theo ngày thực tế sử dụng
        InputNumber(tdbg1, arr)
    End Sub


    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Ket_qua_tinh_khau_hao_-_D02F5012") & UnicodeCaption(gbUnicode) 'KÕt qu¶ tÛnh khÊu hao - D02F5012
        '================================================================ 
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        btnSave.Text = rl3("_Luu") '&Lưu
        btnChangeEntry.Text = rL3("_Chuyen_but_toan") '&Chuyển bút toán
        btnF12.Text = rL3("Hien_thi") & Space(1) & "(F12)" 'Hiển thị
        '================================================================ 
        TabPage1.Text = "1. " & rl3("Danh_sach_tai_san_tinh_KH") '1. Danh sách tài sản tính KH
        TabPage2.Text = "2. " & rL3("Tai_san_het_KH") & " (0)" '2. Tài sản hết KH
        '================================================================ 
        tdbdAna01ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna01ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna02ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna02ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna03ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna03ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna04ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna04ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna05ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna05ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna06ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna06ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna07ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna07ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna08ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna08ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna09ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna09ID.Columns("AnaName").Caption = rL3("Ten") 'Tên
        tdbdAna10ID.Columns("AnaID").Caption = rL3("Ma") 'Mã
        tdbdAna10ID.Columns("AnaName").Caption = rL3("Ten") 'Tên

        '================================================================ 
        tdbg1.Columns("AssetID").Caption = rl3("Ma_tai_san") 'Mã tài sản
        tdbg1.Columns("AssetName").Caption = rl3("Ten_tai_san") 'Tên tài sản
        tdbg1.Columns("DepAmount").Caption = rl3("Muc_khau_hao") 'Mức khấu hao
        tdbg1.Columns("CurrentCost").Caption = rl3("Nguyen_gia") 'Nguyên giá
        tdbg1.Columns("LastDepAmount").Caption = rl3("Muc_KH_ky_truoc") 'Mức KH kỳ trước
        tdbg1.Columns("CurrentLTDDepreciation").Caption = rl3("Luy_ke_KH") 'Lũy kế KH
        tdbg1.Columns("DepreciatedPeriod").Caption = rl3("So_ky_da_KH") 'Số kỳ đã KH
        tdbg1.Columns("LastPeriod").Caption = rL3("So_ky_con_lai") 'Số kỳ còn lại
        tdbg1.Columns("DepreciationDayNum").Caption = "Số ngày tính khấu hao"  ' Số ngày tính khấu hao
        tdbg1.Columns(COL_ProjectID).Caption = rL3("Cong_trinh") 'Dự án
        tdbg1.Columns(COL_ProjectName).Caption = rL3("Ten_cong_trinh") 'Tên dự án
        tdbg1.Columns(COL_TaskID).Caption = rL3("Hang_muc") 'Hạng mục
        tdbg1.Columns(COL_TaskName).Caption = rL3("Ten_hang_muc") 'Tên hạng mục
        tdbg1.Columns(COL_BudgetID).Caption = rL3("Ngan_sach") 'Ngân sách
        tdbg1.Columns(COL_BudgetName).Caption = rL3("Ten_ngan_sach") 'Tên ngân sách
        tdbg1.Columns(COL_BudgetItemID).Caption = rL3("Hang_muc_ngan_sach") 'Hạng mục ngân sách
        tdbg1.Columns(COL_BudgetItemName).Caption = rL3("Ten_hang_muc_NS") 'Tên hạng mục NS
        tdbg1.Columns(COL_MethodName).Caption = rL3("Phuong_phap_KHU")    'ID-131710

        tdbg2.Columns("AssetID").Caption = rl3("Ma_tai_san") 'Mã tài sản
        tdbg2.Columns("AssetName").Caption = rl3("Ten_tai_san") 'Tên tài sản
        '================================================================ 
        mnuFind.Text = rl3("Tim__kiem") 'Tìm &kiếm
        mnuListAll.Text = rl3("_Liet_ke_tat_ca") '&Liệt kê tất cả
    End Sub

    Private Sub Load10TDBDropDownAna()
        '31/8/2021, Nguyễn Thị Mỹ Lài:id 180303-Cho phép sửa khoản mục Kcode khi tính KH TSCD (màn hình D2F5012) Hùng Vương
        If D02Systems.IsEditAnaID = True Then
            LoadTDBDropDownAnaForDivision(tdbdAna01ID, tdbdAna02ID, tdbdAna03ID, tdbdAna04ID, tdbdAna05ID, tdbdAna06ID, tdbdAna07ID, tdbdAna08ID, tdbdAna09ID, tdbdAna10ID, tdbg1, COL_Ana01ID, gbUnicode)
        Else
            For i As Integer = COL_Ana01ID To COL_Ana10ID
                tdbg1.Splits(0).DisplayColumns(i).Locked = True
                tdbg1.Splits(0).DisplayColumns(i).Button = False
            Next
        End If

    End Sub

    Private Sub btnChangeEntry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChangeEntry.Click

        '======================
        '18/4/2017, id 96337-Đóng gói V4.1
        Dim dicParaIn As New Dictionary(Of String, Object) 'DS tham số đầu vào
        Lemon3.CallDxxMxx40("D02E2040", "D02F0502", dicParaIn)

    End Sub


#Region "Active Find Client - List All "
    Private WithEvents Finder As New D99C1001
    Private sFind As String = ""
    Private dtCaptionCols As DataTable

    Private Sub mnuFind_Click(ByVal sender As Object, ByVal e As C1.Win.C1Command.ClickEventArgs) Handles mnuFind.Click
        If Not CallMenuFromGrid(tdbg1, e) Then Exit Sub
        gbEnabledUseFind = True
        '*****************************************
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbg1.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        Dim Arr As New ArrayList
        AddColVisible(tdbg1, SPLIT0, Arr, , , , gbUnicode)
        'Tạo tableCaption: đưa tất cả các cột trên lưới có Visible = True vào table 
        dtCaptionCols = CreateTableForExcelOnly(tdbg1, Arr)
        'End If

        ShowFindDialogClient(Finder, dtCaptionCols, Me.Name, "0", gbUnicode)
        '*****************************************
    End Sub

    Private Sub Finder_FindClick(ByVal ResultWhereClause As Object) Handles Finder.FindClick
        If ResultWhereClause Is Nothing Or ResultWhereClause.ToString = "" Then Exit Sub
        sFind = ResultWhereClause.ToString()
        ReLoadTDBGrid()
    End Sub

    Private Sub mnuListAll_Click(ByVal sender As Object, ByVal e As C1.Win.C1Command.ClickEventArgs) Handles mnuListAll.Click
        sFind = ""
        ReLoadTDBGrid()
    End Sub

    Private Sub ReLoadTDBGrid()
        LoadGridFind(tdbg1, dt, sFind)
        'FooterBar
        FooterTotalGrid(tdbg1, COL_AssetID)
        FooterTotalGrid(tdbg2, COL2_AssetID)
        FooterSum(tdbg1, iColumns)

        CheckMenu(Me.Name, C1CommandHolder, tdbg1.RowCount, gbEnabledUseFind, True)
    End Sub
#End Region

    Private Sub LoadData()
        'Load Tab1
        Dim sSQL As String
        sSQL = SQLStoreD02P0029()
        dt = ReturnDataTable(sSQL)
        If dt.Rows.Count < 1 Then 'Không có dữ liệu
            D99C0008.MsgL3(rl3("Khong_co_but_toan_khau_hao"))
            gbEnabledUseFind = False
            LoadDataSource(tdbg1, dt, gbUnicode)
            CheckMenu(Me.Name, C1CommandHolder, tdbg1.RowCount, gbEnabledUseFind, False)
        Else 'Có dữ liệu
            If Not gbEnabledUseFind Then 'Chưa nhấn tìm kiếm
                LoadDataSource(tdbg1, dt, gbUnicode)
                CheckMenu(Me.Name, C1CommandHolder, tdbg1.RowCount, gbEnabledUseFind, False)
            Else 'Nhấn Tìm kiếm
                ReLoadTDBGrid()
            End If
        End If

        'Load Tab2
        sSQL = SQLStoreD02P1502()
        LoadDataSource(tdbg2, sSQL, gbUnicode)
        TabPage2.Text = TabPage2.Text.Replace("0", tdbg2.RowCount.ToString)

        'FooterBar  
        FooterTotalGrid(tdbg1, COL_AssetID)
        FooterTotalGrid(tdbg2, COL2_AssetID)
        FooterSum(tdbg1, iColumns)

    End Sub
    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0029
    '# Created User: Lê Sơn Long
    '# Created Date: 04/11/2010 11:30:53
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0029() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0029 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'gbUnicode, varchar[20], NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1502
    '# Created User: Lê Sơn Long
    '# Created Date: 04/11/2010 11:31:13
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1502() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P1502 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, tinyint, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'TranYear, int, NOT NULL
        Return sSQL
    End Function


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ' UPDATE 14/6/2013 ID 56700
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        'Chặn lỗi khi đang vi phạm trên lưới mà nhấn Alt + L
        btnSave.Focus()
        If btnSave.Focused = False Then Exit Sub
        '************************************
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        '   If Not AllowSave() Then Exit Sub

        'Kiểm tra Ngày phiếu có phù hợp với kỳ kế toán hiện tại không

        'If Not CheckVoucherDateInPeriod(c1dateVoucherDate.Text) Then c1dateVoucherDate.Focus() : Exit Sub
        btnSave.Enabled = False
        btnClose.Enabled = False

        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder ' = SQLStoreD02P1503()

        '31/8/2021, Nguyễn Thị Mỹ Lài:id 180303-Cho phép sửa khoản mục Kcode khi tính KH TSCD (màn hình D2F5012) Hùng Vương
        sSQL.Append(CreateTableTMP())
        sSQL.Append(InsertTableTMP())

        sSQL.Append(SQLStoreD02P1503())

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            _bSaveOk = True
            btnClose.Enabled = True
            btnChangeEntry.Enabled = True
            If D99C0008.MsgAsk(rl3("Ban_co_muon_thuc_hien_viec_chuyen_but_toan_khong")) = Windows.Forms.DialogResult.Yes Then
                btnChangeEntry_Click(Nothing, Nothing)
            End If
            '   btnSave.Enabled = True
            btnClose.Focus()
        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub

    Private Function CreateTableTMP() As StringBuilder
        '31/8/2021, Nguyễn Thị Mỹ Lài:id 180303-Cho phép sửa khoản mục Kcode khi tính KH TSCD (màn hình D2F5012) Hùng Vương
        Dim sSQL As New StringBuilder

        sSQL.Append("Create Table [#D02T5012_" & gsUserID & "] (" & vbCrLf)
        sSQL.Append("TransactionID varchar(20), " & vbCrLf)
        sSQL.Append("AssetID varchar(20), " & vbCrLf)
        sSQL.Append("VoucherNo varchar(20), " & vbCrLf)
        sSQL.Append("Ana01ID varchar(50), " & vbCrLf)
        sSQL.Append("Ana02ID varchar(50), " & vbCrLf)
        sSQL.Append("Ana03ID varchar(50), " & vbCrLf)
        sSQL.Append("Ana04ID varchar(50), " & vbCrLf)
        sSQL.Append("Ana05ID varchar(50), " & vbCrLf)
        sSQL.Append("Ana06ID varchar(50), " & vbCrLf)
        sSQL.Append("Ana07ID varchar(50), " & vbCrLf)
        sSQL.Append("Ana08ID varchar(50), " & vbCrLf)
        sSQL.Append("Ana09ID varchar(50), " & vbCrLf)
        sSQL.Append("Ana10ID varchar(50), " & vbCrLf)
        sSQL.Append("BatchID varchar(20) " & vbCrLf)
        sSQL.Append(") " & vbCrLf)

        Return sSQL
    End Function

    Private Function InsertTableTMP() As StringBuilder
        '31/8/2021, Nguyễn Thị Mỹ Lài:id 180303-Cho phép sửa khoản mục Kcode khi tính KH TSCD (màn hình D2F5012) Hùng Vương
        Dim sSQL As New StringBuilder
        Dim sRet As New StringBuilder

        For i As Integer = 0 To tdbg1.RowCount - 1
            sSQL.Append("Insert into [#D02T5012_" & gsUserID & "] (" & vbCrLf)
            sSQL.Append("TransactionID, AssetID, VoucherNo, " & vbCrLf)
            sSQL.Append("Ana01ID, Ana02ID, Ana03ID, Ana04ID, Ana05ID, " & vbCrLf)
            sSQL.Append("Ana06ID, Ana07ID, Ana08ID, Ana09ID, Ana10ID, BatchID" & vbCrLf)
            sSQL.Append(") Values ( " & vbCrLf)
            sSQL.Append(SQLString(tdbg1(i, COL_TransactionID)) & COMMA)
            sSQL.Append(SQLString(tdbg1(i, COL_AssetID)) & COMMA)
            sSQL.Append(SQLString(tdbg1(i, COL_VoucherNo)) & COMMA & vbCrLf)
            sSQL.Append(SQLString(tdbg1(i, COL_Ana01ID)) & COMMA)
            sSQL.Append(SQLString(tdbg1(i, COL_Ana02ID)) & COMMA)
            sSQL.Append(SQLString(tdbg1(i, COL_Ana03ID)) & COMMA)
            sSQL.Append(SQLString(tdbg1(i, COL_Ana04ID)) & COMMA)
            sSQL.Append(SQLString(tdbg1(i, COL_Ana05ID)) & COMMA & vbCrLf)
            sSQL.Append(SQLString(tdbg1(i, COL_Ana06ID)) & COMMA)
            sSQL.Append(SQLString(tdbg1(i, COL_Ana07ID)) & COMMA)
            sSQL.Append(SQLString(tdbg1(i, COL_Ana08ID)) & COMMA)
            sSQL.Append(SQLString(tdbg1(i, COL_Ana09ID)) & COMMA)
            sSQL.Append(SQLString(tdbg1(i, COL_Ana10ID)) & COMMA)
            sSQL.Append(SQLString(tdbg1(i, COL_BatchID)))
            sSQL.Append(") " & vbCrLf)

            sRet.Append(sSQL)
            sSQL.Remove(0, sSQL.Length)
        Next

        Return sRet
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1503
    '# Created User: Hoàng Nhân
    '# Created Date: 14/06/2013 08:50:11
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1503() As String
        Dim sSQL As String = ""
        sSQL &= ("-- Insert data to D02T0012 From temp table D05T5012" & vbCrlf)
        sSQL &= "Exec D02P1503 "
        sSQL &= SQLString(gsUserID) 'UserID, varchar[50], NOT NULL
        Return sSQL
    End Function

    
    Private Sub tdbg1_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg1.BeforeColUpdate
        '--- Kiểm tra giá trị hợp lệ
        Select Case e.ColIndex
            Case COL_Ana01ID, COL_Ana02ID, COL_Ana03ID, COL_Ana04ID, COL_Ana05ID, COL_Ana06ID, COL_Ana07ID, COL_Ana08ID, COL_Ana09ID, COL_Ana10ID
                If tdbg1.Columns(e.ColIndex).Text <> tdbg1.Columns(e.ColIndex).DropDown.Columns(tdbg1.Columns(e.ColIndex).DropDown.DisplayMember).Text Then
                    tdbg1.Columns(e.ColIndex).Text = ""
                End If
        End Select
    End Sub

    Private Sub tdbg1_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg1.ComboSelect
        tdbg1.UpdateData()
    End Sub

    Private Sub HeadClick(ByVal iCol As Integer)
        If tdbg1.RowCount <= 0 Then Exit Sub
        Select Case iCol
            Case COL_Ana01ID, COL_Ana02ID, COL_Ana03ID, COL_Ana04ID, COL_Ana05ID, COL_Ana06ID, COL_Ana07ID, COL_Ana08ID, COL_Ana09ID, COL_Ana10ID
                tdbg1.AllowSort = False
                'Copy 1 cột
                CopyColumns(tdbg1, iCol, tdbg1.Columns(iCol).Text, tdbg1.Bookmark)
               
            Case Else
                tdbg1.AllowSort = True
        End Select
    End Sub

    Private Sub tdbg1_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg1.HeadClick
        HeadClick(e.ColIndex)
    End Sub

    Private Sub tdbg1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg1.KeyDown
        If e.Control And e.KeyCode = Keys.S Then HeadClick(tdbg1.Col)
    End Sub



End Class