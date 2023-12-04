'#-------------------------------------------------------------------------------------
'# Created Date: 25/09/2007 5:05:37 PM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 25/09/2007 5:05:37 PM
'# Modify User: Trần Thị ÁiTrâm
'#-------------------------------------------------------------------------------------
Imports System.Text

Public Class D02F1003

#Region "Const of tdbg"
    Private Const COL_CipID As Integer = 0              ' CipID
    Private Const COL_VoucherTypeID As Integer = 1      ' Loại phiếu
    Private Const COL_VoucherNo As Integer = 2          ' Số phiếu
    Private Const COL_VoucherDate As Integer = 3        ' Ngày phiếu
    Private Const COL_SeriNo As Integer = 4             ' Số Sêri
    Private Const COL_RefNo As Integer = 5              ' Số hóa đơn
    Private Const COL_RefDate As Integer = 6            ' Ngày hóa đơn
    Private Const COL_Description As Integer = 7        ' Diễn giải
    Private Const COL_DebitAccountID As Integer = 8     ' TK nợ
    Private Const COL_CreditAccountID As Integer = 9    ' TK có
    Private Const COL_CurrencyID As Integer = 10        ' Loại tiền
    Private Const COL_ExchangeRate As Integer = 11      ' Tỷ giá
    Private Const COL_OriginalAmount As Integer = 12    ' Nguyên tệ
    Private Const COL_ConvertedAmount As Integer = 13   ' Qui đổi
    Private Const COL_ObjectTypeID As Integer = 14      ' Loại đối tượng
    Private Const COL_ObjectID As Integer = 15          ' Mã đối tượng
    Private Const COL_BatchID As Integer = 16           ' BatchID
    Private Const COL_TransactionID As Integer = 17     ' TransactionID
    Private Const COL_ModuleID As Integer = 18          ' ModuleID
    Private Const COL_CreateUserID As Integer = 19      ' CreateUserID
    Private Const COL_CreateDate As Integer = 20        ' CreateDate
    Private Const COL_LastModifyUserID As Integer = 21  ' LastModifyUserID
    Private Const COL_LastModifyDate As Integer = 22    ' LastModifyDate
    Private Const COL_Internal As Integer = 23          ' Internal
    Private Const COL_Status As Integer = 24            ' Status
    Private Const COL_TransactionTypeID As Integer = 25 ' TransactionTypeID
    Private Const COL_Ana01ID As Integer = 26         ' Khoản mục 01
    Private Const COL_Ana02ID As Integer = 27         ' Khoản mục 02
    Private Const COL_Ana03ID As Integer = 28        ' Khoản mục 03
    Private Const COL_Ana04ID As Integer = 29        ' Khoản mục 04
    Private Const COL_Ana05ID As Integer = 30      ' Khoản mục 05
    Private Const COL_Ana06ID As Integer = 31        ' Khoản mục 06
    Private Const COL_Ana07ID As Integer = 32         ' Khoản mục 07
    Private Const COL_Ana08ID As Integer = 33         ' Khoản mục 08
    Private Const COL_Ana09ID As Integer = 34         ' Khoản mục 09
    Private Const COL_Ana10ID As Integer = 35         ' Khoản mục 10

#End Region

#Region "Const of tdbg1"
    Private Const COL1_Selected As Integer = 0         ' Chọn
    Private Const COL1_VoucherTypeID As Integer = 1    ' Loại phiếu
    Private Const COL1_VoucherNo As Integer = 2        ' Số phiếu
    Private Const COL1_VoucherDate As Integer = 3      ' Ngày phiếu
    Private Const COL1_SeriNo As Integer = 4           ' Số Sêri
    Private Const COL1_RefNo As Integer = 5            ' Số hóa đơn
    Private Const COL1_RefDate As Integer = 6          ' Ngày hóa đơn
    Private Const COL1_ObjectTypeID As Integer = 7     ' Loại hóa đơn
    Private Const COL1_ObjectID As Integer = 8         ' Mã đối tượng
    Private Const COL1_Description As Integer = 9      ' Diễn giải
    Private Const COL1_DebitAccountID As Integer = 10  ' TK nợ
    Private Const COL1_CreditAccountID As Integer = 11 ' TK có
    Private Const COL1_CurrencyID As Integer = 12      ' Loại tiền
    Private Const COL1_ExchangeRate As Integer = 13    ' Tỷ giá
    Private Const COL1_OriginalAmount As Integer = 14  ' Nguyên tệ
    Private Const COL1_ConvertedAmount As Integer = 15 ' Qui đổi
    Private Const COL1_CipID As Integer = 16           ' CipID
    Private Const COL1_TransactionID As Integer = 17   ' TransactionID
    Private Const COL1_ModuleID As Integer = 18        ' ModuleID
    Private Const COL1_Ana01ID As Integer = 19         ' Khoản mục 01
    Private Const COL1_Ana02ID As Integer = 20         ' Khoản mục 02
    Private Const COL1_Ana03ID As Integer = 21         ' Khoản mục 03
    Private Const COL1_Ana04ID As Integer = 22         ' Khoản mục 04
    Private Const COL1_Ana05ID As Integer = 23         ' Khoản mục 05
    Private Const COL1_Ana06ID As Integer = 24         ' Khoản mục 06
    Private Const COL1_Ana07ID As Integer = 25         ' Khoản mục 07
    Private Const COL1_Ana08ID As Integer = 26         ' Khoản mục 08
    Private Const COL1_Ana09ID As Integer = 27         ' Khoản mục 09
    Private Const COL1_Ana10ID As Integer = 28         ' Khoản mục 10
#End Region

#Region "Const of tdbg2"
    Private Const COL2_CipNo As Integer = 0            ' Mã XDCB
    Private Const COL2_CipName As Integer = 1          ' Tên mã XDCB
    Private Const COL2_AccountID As Integer = 2        ' TK XDCB
    Private Const COL2_Desciption As Integer = 3       ' Diễn giải
    Private Const COL2_AccountName As Integer = 4      ' Tên
    Private Const COL2_Status As Integer = 5           ' 
    Private Const COL2_Disabled As Integer = 6         ' 
    Private Const COL2_CreateDate As Integer = 7       ' 
    Private Const COL2_CreateUserID As Integer = 8     ' 
    Private Const COL2_LastModifyUserID As Integer = 9 ' 
    Private Const COL2_LastModifyDate As Integer = 10  ' 
    Private Const COL2_CipID As Integer = 11           ' 
#End Region

    Private dtGrid, dtGrid2 As DataTable
    Private sAuditCode As String
    Private byAudit As Byte
    Dim dtCodeID As DataTable

#Region "Active Find Client - List All "

    Private WithEvents Finder As New D99C1001
    Private sFind As String = ""
    Dim dtCaptionCols As DataTable

    Private Sub tsbFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbFind.Click, tsmFind.Click, mnsFind.Click
        gbEnabledUseFind = True
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbg.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        Dim Arr As New ArrayList
        AddColVisible(tdbg, SPLIT0, Arr, , , , gbUnicode)
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
        'End If
        ShowFindDialogClient(Finder, dtCaptionCols, Me.Name, "0", gbUnicode)
    End Sub

    Private Sub Finder_FindClick(ByVal ResultWhereClause As Object) Handles Finder.FindClick
        If ResultWhereClause Is Nothing Then Exit Sub
        sFind = ResultWhereClause.ToString
        ReLoadTDBGrid()
    End Sub

    Private Sub tsbListAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbListAll.Click, tsmListAll.Click, mnsListAll.Click
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
#End Region

    Private Sub tsbSysInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, tsmSysInfo.Click, mnsSysInfo.Click
        ShowSysInfoDialog(tdbg.Columns(COL_CreateUserID).Text, tdbg.Columns(COL_CreateDate).Text, tdbg.Columns(COL_LastModifyUserID).Text, tdbg.Columns(COL_LastModifyDate).Text)
    End Sub


    Private Sub tsbClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P2050
    '# Created User: HUỲNH KHANH
    '# Created Date: 24/12/2014 02:26:18
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P2050(ByVal sCipID As String, ByVal sCipAccountID As String, Optional ByVal sReportID As String = "", Optional ByVal sFormID As String = "", Optional ByVal iMode As Integer = 0) As String
        Dim sSQL As String = ""
        sSQL &= ("-- Do nguon cho luoi 0" & vbCrLf)
        sSQL &= "Exec D02P2050 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[50], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(sCipID) & COMMA 'CipID, varchar[50], NOT NULL
        sSQL &= SQLString(sCipAccountID) & COMMA 'CipAccountID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, int, NOT NULL
        sSQL &= SQLString(sReportID) & COMMA
        sSQL &= SQLString(sFormID) & COMMA
        sSQL &= SQLNumber(iMode)
        Return sSQL
    End Function

    Private Sub LoadTDBGrid(ByVal sCipID As String, ByVal sAccountID As String, Optional ByVal bFlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        dtGrid = ReturnDataTable(SQLStoreD02P2050(sCipID, sAccountID))
        gbEnabledUseFind = dtGrid.Rows.Count > 0
        If bFlagAdd Then
            ResetFilter(tdbg, sFilter, bRefreshFilter)
            sFind = ""
        End If
        LoadDataSource(tdbg, dtGrid, gbUnicode)
        ReLoadTDBGrid()
        If sKey <> "" Then 'Khi Thêm mới hoặc Sửa đều thực thi
            Dim dt As DataTable = dtGrid.DefaultView.ToTable
            Dim dr() As DataRow = dt.Select("BatchID =" & SQLString(sKey), dt.DefaultView.Sort)
            If dr.Length > 0 Then tdbg.Row = dt.Rows.IndexOf(dr(0))
            If tdbg.Focused = False Then tdbg.Focus()
        End If
        ResetGrid()
    End Sub

    Private Sub LoadTDBGrid2(ByVal ID1 As String, ByVal ID2 As String, Optional ByVal bFlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        Dim sSQL As String = ""
        'Load tdbcCipNo
        sSQL = "--Do nguon cho luoi 2" & vbCrLf
        sSQL &= " Select  CipNo,CipName" & UnicodeJoin(gbUnicode) & " as  CipName, Description" & UnicodeJoin(gbUnicode) & " as Description, D02T0100.AccountID, Account.AccountName" & UnicodeJoin(gbUnicode) & " as AccountName,CipID," & vbCrLf
        sSQL &= " D02T0100.Status, D02T0100.Disabled, D02T0100.CreateDate, D02T0100.CreateUserID, D02T0100.LastModifyUserID, D02T0100.LastModifyDate " & vbCrLf
        sSQL &= " From D02T0100 WITH(NOLOCK) Inner Join Account WITH(NOLOCK) On Account.AccountID= D02T0100.AccountID"
        sSQL &= " Where Status <> 2 And D02T0100.Disabled=0 And D02T0100.DivisionID = " & SQLString(gsDivisionID)
        If ID1 <> "" And ID1 <> "%" And ID2 <> "" And ID2 <> "%" Then
            sSQL &= " And " & ID1 & "ID = " & SQLString(ID2)
        End If
        dtGrid2 = ReturnDataTable(sSQL)
        LoadDataSource(tdbg2, dtGrid2, gbUnicode)
        ReLoadTDBGrid2()
        If sKey <> "" Then 'Khi Thêm mới hoặc Sửa đều thực thi
            Dim dt As DataTable = dtGrid2.DefaultView.ToTable
            Dim dr() As DataRow = dt.Select("CipID =" & SQLString(sKey), dt.DefaultView.Sort)
            If dr.Length > 0 Then tdbg2.Row = dt.Rows.IndexOf(dr(0))
            If tdbg2.Focused = False Then tdbg2.Focus()
        End If

        'If tdbg2.RowCount > 0 Then
        '    LoadTDBGrid(tdbg2.Columns(COL2_CipID).Text, tdbg2.Columns(COL2_AccountID).Text)
        'Else
        '    If dtGrid IsNot Nothing Then dtGrid.Clear()

        '    ResetGrid()
        'End If

    End Sub

    Private Sub ResetGrid()
        CheckMenu(PARA_FormIDPermission, ToolStrip1, tdbg.RowCount, gbEnabledUseFind, True, ContextMenuStripp)
        tdbg_Footext()
        EnableMenu()
    End Sub

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""

        sSQL = "Select 0 as DisplayOrder,'%' As CodeID, " & AllName & " As Description, '%' As TypeCodeID Union" & vbCrLf
        sSQL &= "Select 1 as DisplayOrder,ACodeID As CodeID, Description" & UnicodeJoin(gbUnicode) & " as Description, TypeCodeID" & vbCrLf
        sSQL &= "FROM D02T0041 WITH(NOLOCK) "
        sSQL &= "Where Disabled = 0 And Type = 'X' Order By DisplayOrder,TypeCodeID, CodeID"
        dtCodeID = ReturnDataTable(sSQL)

        sSQL = "Select 0 as DisplayOrder,'%' As TypeCodeID, " & AllName & " As TypeCodeName Union" & vbCrLf
        sSQL &= "Select 1 as DisplayOrder,TypeCodeID, " & IIf(gsLanguage = "84", "VieTypeCodeName", "EngTypeCodeName").ToString & UnicodeJoin(gbUnicode) & " As TypeCodeName" & vbCrLf
        sSQL &= "FROM D02T0040 WITH(NOLOCK) "
        sSQL &= "Where Disabled = 0 And Type = 'X' Order By DisplayOrder,TypeCodeID"
        LoadDataSource(tdbcTypeCodeID, sSQL, gbUnicode)
        LoadtdbcCodeID("%")

        'ID 86798 30.05.2016
        c1dateDateFrom.Value = Date.Now
        c1dateDateTo.Value = Date.Now
        'Load tdbcPeriod
        LoadCboPeriodReport(tdbcPeriodFrom, tdbcPeriodTo, D02)
        If tdbcPeriodFrom IsNot Nothing And tdbcPeriodTo IsNot Nothing Then
            tdbcPeriodFrom.SelectedIndex = 0
            tdbcPeriodTo.SelectedIndex = 0
        End If
        '***************************************************************

        'Load dự án
        LoadProject(tdbcProjectID)
        'Load ngân sách
        LoadBudget(tdbcBudgetID, Me.Name)
        'Load tdbcAccountID
        LoadAccountID(tdbcAccountID, "AccountStatus = 0 And GroupID = '9'", gbUnicode)
    End Sub

    Private Sub LoadtdbcCodeID(ByVal ID As String)
        LoadDataSource(tdbcCodeID, ReturnTableFilter(dtCodeID, "TypeCodeID = " & SQLString(ID) & " Or TypeCodeID = '%'", True), gbUnicode)
    End Sub

    Private Sub D02F1003_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me, True)
        End If
        If e.KeyCode = Keys.F11 Then
            HotKeyF11(Me, tdbg)
        End If
        If e.KeyCode = Keys.Control Then
            If e.KeyCode = Keys.F Then
                tsbFind_Click(Nothing, Nothing)
            ElseIf e.KeyCode = Keys.A Then
                tsbListAll_Click(Nothing, Nothing)
            End If
        End If
    End Sub

    Private Sub D02F1003_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        LoadInfoGeneral()
        SetShortcutPopupMenu(Me, ToolStrip1, ContextMenuStripp)
        Loadlanguage()
        LoadTDBCombo()
        gbEnabledUseFind = False
        sAuditCode = "CIPCostColl"
        byAudit = PermissionAudit(sAuditCode)
        btnCollection.Enabled = False
        InputbyUnicode(Me, gbUnicode)
        InputDateInTrueDBGrid(tdbg, COL_RefDate, COL_VoucherDate)
        ResetColorGrid(tdbg, tdbg2)
        tdbg_NumberFormat()
        bUseAna = LoadTDBGridAnalysisCaption(D02, tdbg, COL_Ana01ID, SPLIT0, True, gbUnicode)

        If Not bUseAna Then
            For i As Integer = COL_Ana01ID To COL_Ana10ID
                tdbg.Splits(SPLIT0).DisplayColumns(i).Visible = Convert.ToBoolean(tdbg.Columns(i).Tag)
            Next
        End If

        LoadTDBDropDownAna(tdbdAna01ID_1, tdbdAna02ID_1, tdbdAna03ID_1, tdbdAna04ID_1, tdbdAna05ID_1, tdbdAna06ID_1, tdbdAna07ID_1, tdbdAna08ID_1, tdbdAna09ID_1, tdbdAna10ID_1, tdbg, COL_Ana01ID, gbUnicode)
        ResetGrid()
        EnableMenu()
        CheckOtherMenu2()
        SetBackColorObligatory()
        SetResolutionForm(Me, ContextMenuStripp)
        SplitContainer1.Panel2Collapsed = True
        'mnsCollection.Enabled = tdbg2.RowCount > 0
        'tsmCollection.Enabled = tdbg2.RowCount > 0
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        btnFilter.Focus()
        If btnFilter.Focused = False Then Exit Sub
        Me.Cursor = Cursors.WaitCursor

        ResetFilter(tdbg2, sFilter2, bRefreshFilter2)
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        'LoadTDBGrid2("%", "%")
        If ReturnValueC1Combo(tdbcTypeCodeID) <> "" And ReturnValueC1Combo(tdbcCodeID) <> "" Then
            LoadTDBGrid2(ReturnValueC1Combo(tdbcTypeCodeID), ReturnValueC1Combo(tdbcCodeID))
        Else
            LoadTDBGrid2("%", "%")
        End If
        EnableMenu()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub EnableMenu()
        'Status = 2: đã hình thành
        'Statys = 0; chưa tập hợp
        'Status = 1: đã tập hợp
        If tdbg2.RowCount > 0 Then
            If L3Int(tdbg2.Columns(COL2_Status).Text) = 0 Then
                tsmCancelCollection.Enabled = False
                mnsCancelCollection.Enabled = False
                mnsCollection.Enabled = True
                tsmCollection.Enabled = True
            Else
                tsmCancelCollection.Enabled = True
                mnsCancelCollection.Enabled = True
                mnsCollection.Enabled = True
                tsmCollection.Enabled = True
            End If
        Else
            tsmCancelCollection.Enabled = False
            mnsCancelCollection.Enabled = False
            mnsCollection.Enabled = False
            tsmCollection.Enabled = False
        End If
    End Sub

    ' Bỏ ngày 19/7/2012 theo incident 47959 của HOANGNAM người sửa VANVINH 
    'Private Sub tsbAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbAdd.Click, tsmAdd.Click, mnsAdd.Click
    '    Dim f As New D02F1004
    '    With f
    '        .BatchID = ""
    '        .CipID = tdbcCipNo.Columns("CipID").Text
    '        .CipNo = tdbcCipNo.Text
    '        .CipName = txtCipName.Text
    '        .AccountID = txtAccountID.Text
    '        .ByAudit = byAudit
    '        .sAuditCode = sAuditCode
    '        .FormState = EnumFormState.FormAdd
    '        .ShowDialog()
    '        If gbSavedOK = True Then LoadTDBGrid(tdbcCipNo.Columns("CipID").Text, txtAccountID.Text, True, .KeyID)
    '        .Dispose()
    '    End With
    'End Sub

    ' Bỏ ngày 19/7/2012 theo incident 47959 của HOANGNAM người sửa VANVINH 
    'Private Sub tsbEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbEdit.Click, tsmEdit.Click, mnsEdit.Click
    '    'Dim iBookmark As Integer
    '    'If Not IsDBNull(tdbg.Bookmark) Then iBookmark = tdbg.Bookmark
    '    'Kiểm tra điều kiện sửa
    '    If tdbg.Columns(COL_ModuleID).Text <> "02" Then
    '        D99C0008.MsgL3(rl3("Phieu_nay_tu_Module_khac_chuyen_sangBan_khong_duoc_phep_sua_phieu_nay"))
    '        Exit Sub
    '    Else
    '        Dim f As New D02F1004
    '        With f
    '            .BatchID = tdbg.Columns(COL_BatchID).Text
    '            .CipID = tdbcCipNo.Columns("CipID").Text
    '            .CipNo = tdbcCipNo.Text
    '            .CipName = txtCipName.Text
    '            .AccountID = txtAccountID.Text
    '            .ByAudit = byAudit
    '            .sAuditCode = sAuditCode
    '            .FormState = EnumFormState.FormEdit
    '            .ShowDialog()
    '            .Dispose()
    '            If gbSavedOK Then
    '                LoadTDBGrid(tdbcCipNo.Columns("CipID").Value.ToString, txtAccountID.Text, , .KeyID)
    '                'If Not IsDBNull(iBookmark) Then tdbg.Bookmark = iBookmark
    '            End If
    '        End With
    '    End If
    'End Sub

    Private Function AllowDelete() As Boolean
        '' Status = 2 'mº x¡y døng c¥ b¶n nªy ¢º hØnh thªnh xong tªi s¶n
        If L3Int(tdbg2.Columns(COL2_Status).Text) = 2 Then
            If D99C0008.MsgAsk("Hạng mục này đã hình thành TSCĐ." & Space(1) & "Bạn không thể xóa phiếu này được." & Space(1) & "Bạn có muốn xem không?", MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                tsbView_Click(Nothing, Nothing)
            End If
            Return False
        End If

        Return True
    End Function

    Private Sub mnsDeleteSum_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If AskDelete() = Windows.Forms.DialogResult.Yes Then
            If Not AllowDelete() Then Exit Sub
            ''Kiểm tra điều kiện xóa
            ''Nếu ModuleID <> 02
            Dim sSQL As String = ""
            If tdbg.Columns(COL_ModuleID).Text <> "02" Then
                sSQL = SQLUpdateD02T0012.ToString & vbCrLf
            Else
                'Sửa ngày 19/7/2012 theo incidetn 47959 của HOANGNAM bởi VANVINH
                If tdbg.Columns(COL_TransactionTypeID).Text = "SDXDCB" Then
                    D99C0008.MsgL3(rL3("Ban_khong_the_xoa_But_toan_Nhap_so_du_chi_phi_XDCB_Ban_can_sang_man_hinh_Nhap_so_du_chi_phi_XDCB_-_D02F1007"))
                    Exit Sub
                End If

                If tdbg.Columns(COL_Internal).Text = "0" Then
                    If (tdbg.Columns(COL_Status).Text = "0" Or tdbg.Columns(COL_Status).Text = "1") And tdbg.Columns(COL_TransactionTypeID).Text <> "SDXDCB" Then 'Nhập chứng từ, tách chi phí 
                        sSQL = SQLUpdateD02T0012.ToString & vbCrLf
                    Else 'Số dư
                        'If D99C0008.MsgAsk("Bạn không thể xóa bút toán nhập số dư XDCB." & Space(1) & "Bạn cần sang màn hình Nhập số dư XDCB - D02F1007." & Space(1) & "Bạn có muốn xem không?", MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                        '    tsbView_Click(Nothing, Nothing)
                        'End If
                        Exit Sub
                    End If
                Else
                    If tdbg.Columns(COL_Status).Text = "1" Then 'Nhập chứng từ, tách chi phí 
                        If D99C0008.MsgAsk("Phiếu này đã được xử lý." & Space(1) & "Bạn không thể xóa phiếu này được." & Space(1) & "Bạn có muốn xem không?", MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                            tsbView_Click(Nothing, Nothing)
                        End If
                        Exit Sub
                    Else 'Số dư
                        sSQL = SQLDeleteD02T0012()
                    End If
                End If
            End If
            sSQL &= "--Cập nhật lại trạng thái mã XDCB" & vbCrLf
            sSQL &= "If Not Exists (Select Top 1 1 From D02T0012 Where CipID=" & SQLString(tdbg2.Columns("CipID").Text) & ")" & vbCrLf
            sSQL &= "Begin" & vbCrLf
            sSQL &= "Update D02T0100 Set Status=0 Where CipID=" & SQLString(tdbg2.Columns("CipID").Text) & vbCrLf
            sSQL &= "End"

            Dim bResult As Boolean = ExecuteSQL(sSQL)
            If bResult = True Then
                'Kiểm tra Audit và thiết lập Auditlog
                'If byAudit = 1 Then
                '    ExecuteAuditLog(sAuditCode, "03", tdbcCipNo.Columns("CipID").Text, txtCipName.Text)
                'End If
                'ExecuteAuditLog(sAuditCode, "03", tdbg2.Columns("CipID").Text, tdbg.Columns(COL_VoucherNo).Text, tdbg.Columns(COL_VoucherDate).Text)
                Lemon3.D91.RunAuditLog("02", sAuditCode, "03", tdbg2.Columns("CipID").Text, tdbg.Columns(COL_VoucherNo).Text, tdbg.Columns(COL_VoucherDate).Text)
                DeleteOK()
                DeleteGridEvent(tdbg, dtGrid, gbEnabledUseFind)
                ResetGrid()
            Else
                DeleteNotOK()
            End If
        End If
    End Sub

    Private Sub tsbView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbView.Click, tsmView.Click, mnsView.Click
        Dim f As New D02F1004
        With f
            .BatchID = tdbg.Columns(COL_BatchID).Text
            .CipID = tdbg2.Columns("CipID").Text
            .CipNo = tdbg2.Columns(COL2_CipNo).Text
            .CipName = tdbg2.Columns(COL2_CipName).Text
            .AccountID = tdbg2.Columns(COL2_AccountID).Text
            .ByAudit = byAudit
            .sAuditCode = sAuditCode
            .FormState = EnumFormState.FormView
            .ShowDialog()
            .Dispose()
        End With
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T0012
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 03/10/2007 11:08:16
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T0012() As String
        Dim sSQL As String = ""
        sSQL &= "--Xóa phiếu nhập tại màn hình D02F1003" & vbCrLf
        sSQL &= "Delete From D02T0012"
        sSQL &= " Where "
        sSQL &= "TransactionID = " & SQLString(tdbg.Columns(COL_TransactionID).Text) & " And "
        sSQL &= "BatchID = " & SQLString(tdbg.Columns(COL_BatchID).Text) & " And "
        sSQL &= "ModuleID = " & SQLString("02") & " And "
        sSQL &= "CipID = " & SQLString(tdbg2.Columns("CipID").Text)
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0012
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 03/10/2007 11:10:45
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0012() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("--Cập nhật lại trạng thái và mã XDCB của phiếu" & vbCrLf)
        sSQL.Append("Update D02T0012 Set ")
        sSQL.Append("Status = " & SQLNumber("0") & COMMA) 'tinyint, NOT NULL
        sSQL.Append("CipID = " & SQLString("")) 'varchar[20], NULL
        sSQL.Append(" Where ")
        sSQL.Append("TransactionID = " & SQLString(tdbg.Columns(COL_TransactionID).Text) & " And ")
        sSQL.Append("ModuleID = " & SQLString(tdbg.Columns(COL_ModuleID).Text) & " And ")
        sSQL.Append("CipID = " & SQLString(tdbg2.Columns("CipID").Text))
        Return sSQL
    End Function

    Private Sub tdbg_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.DoubleClick
        If tdbg.RowCount = 0 Then Exit Sub
        If tdbg.FilterActive Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        If tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Dim sFilter As New System.Text.StringBuilder()
    Dim bRefreshFilter As Boolean = False

    Private Sub tdbg_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.FilterChange
        Try
            If (dtGrid Is Nothing) Then Exit Sub
            If bRefreshFilter Then Exit Sub
            'Filter the data 
            FilterChangeGrid(tdbg, sFilter)
            ReLoadTDBGrid()
        Catch ex As Exception
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte -> Không hiển thị thông báo
        End Try
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        Me.Cursor = Cursors.WaitCursor
        If e.KeyCode = Keys.Enter Then
            tdbg_DoubleClick(sender, e)
            'Else
            '    If e.Control Then
            '        CheckMenu(PARA_FormIDPermission, C1CommandHolder, tdbg.RowCount, gbEnabledUseFind, True)
            '    End If
        End If
        Me.Cursor = Cursors.Default
        HotKeyCtrlVOnGrid(tdbg, e) 'Đã bổ sung D99X0000
    End Sub

    Dim gbEnabledUseFind1 As Boolean
    Private Sub mnsCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnsCollection.Click, tsmCollection.Click
        Me.Cursor = Cursors.WaitCursor
        iMode = 0
        loadNewCollection()
        SetBackColorObligatory()
        '--- Chuẩn Khoản mục b2: Lấy caption cho 10 khoản mục
        bUseAna = LoadTDBGridAnalysisCaption(D02, tdbg1, COL1_Ana01ID, SplitAna, True, gbUnicode)
        'D91 có sử dụng Khoản mục
        'If bUseAna Then iDisplayAnaCol = 1

        If Not bUseAna Then
            tdbg1.Splits(SplitAna).SplitSize = 0
            tdbg1.Splits(SplitAna).HScrollBar.Style = C1.Win.C1TrueDBGrid.ScrollBarStyleEnum.None '19/11/2018, id 115915-Lỗi xem phiếu tập hợp chi phí xây dựng cơ bản
        Else
            tdbg1.Splits(SplitAna).HScrollBar.Style = C1.Win.C1TrueDBGrid.ScrollBarStyleEnum.Always '19/11/2018, id 115915-Lỗi xem phiếu tập hợp chi phí xây dựng cơ bản
        End If

        ResetSplitDividerSize(tdbg1)
        LoadTDBDropDown()
        gbEnabledUseFind1 = False
        ResetColorGrid(tdbg1, 1)
        If dtGrid1 IsNot Nothing Then dtGrid1.Clear()
        Dim dttem As DataTable = ReturnDataTable("Select * From D02T0000")

        If D02Systems.UseD54ForCIP = 0 AndAlso D02Systems.UseBudgetForCIP = 0 Then
            LoadTDBGrid1(tdbg2.Columns(COL2_AccountID).Text)
            pnlBudget.Visible = False
            pnlProject.Visible = False
            'btnFilter1.Visible = False
            tdbcAccountID.Location = pnlProject.Location
        Else
            If D02Systems.UseD54ForCIP = 1 AndAlso D02Systems.UseBudgetForCIP = 0 Then
                pnlBudget.Visible = False
                pnlProject.Visible = True
            ElseIf D02Systems.UseD54ForCIP = 0 AndAlso D02Systems.UseBudgetForCIP = 1 Then
                pnlProject.Visible = False
                pnlBudget.Visible = True
                pnlBudget.Location = pnlProject.Location
            End If
        End If

        tdbg1_LockedColumns()
        tdbg1_NumberFormat()
        btnCollection.Enabled = ReturnPermission("D02F1003") > EnumPermission.View
        iLastcol = CountCol(tdbg1, 1)
        SplitContainer1.Panel1Collapsed = True
        SplitContainer1.Panel2Collapsed = False
        lblAccountID.Visible = False
        tdbcAccountID.Visible = False

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub loadNewCollection()
        tdbcProjectID.Text = ""
        tdbcTaskID.Text = ""
        tdbcBudgetID.Text = ""
        tdbcBudgetItemID.Text = ""
        tdbcAccountID.Text = ""
        If dtGrid1 IsNot Nothing Then dtGrid1.Clear()
    End Sub

    Private Sub tdbg_NumberFormat()
        tdbg.Columns(COL_ExchangeRate).NumberFormat = DxxFormat.ExchangeRateDecimals
        tdbg.Columns(COL_OriginalAmount).NumberFormat = DxxFormat.DecimalPlaces
        tdbg.Columns(COL_ConvertedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
    End Sub

    Private Sub tdbg_Footext()
        FooterTotalGrid(tdbg, COL_VoucherNo)
        'FooterTotalGrid(tdbg2, COL2_CipNo)
        FooterSumNew(tdbg, COL_ConvertedAmount)
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rL3("Tap_hop_chi_phi_XDCB_-__D02F1003") & UnicodeCaption(gbUnicode) 'TËp híp chi phÛ XDCB -  D02F1003
        '================================================================ 

        lblInfoSet.Text = rL3("Thong_tin_tap_hop_chi_phi") 'Thông tin tập hợp chi phí
        lblTypeCodeID.Text = rL3("Ma_phan_tich")
        lblProjectID.Text = rL3("Cong_trinh") 'Dự án
        lblTaskID.Text = rL3("Hang_muc") 'Hạng mục
        lblBudgetID.Text = rL3("Ngan_sach") 'Ngân sách
        lblBudgetItemID.Text = rL3("Hang_muc_ngan_sach") 'Hạng mục ngân sách
        '================================================================ 
        btnCollection.Text = rL3("Tap__hop") 'Tập &hợp
        btnFilter.Text = rL3("Loc") & " (F5)" 'Lọc (F5)
        btnFilter1.Text = rL3("Lo_c")
        'btnAction.Text = rl3("_Thuc_hien_") '&Thực hiện...
        'btnClose.Text = rl3("Do_ng") 'Đó&ng

        tsmCancelCollection.Text = rL3("_Bo_tap_hop")
        mnsCancelCollection.Text = tsmCancelCollection.Text
        mnsAutoProject.Text = rL3("Tao_tu_dong_tu_du_an")
        mnsAutoBudget.Text = rL3("Tao_tu_dong_tu_ngan_sach")
        '================================================================ 

        tdbcCodeID.Columns("CodeID").Caption = rL3("Ma") 'Mã
        tdbcCodeID.Columns("Description").Caption = rL3("Dien_giai")
        '================================================================ 
        tdbg.Columns("VoucherTypeID").Caption = rL3("Loai_phieu") 'Loại phiếu
        tdbg.Columns("VoucherNo").Caption = rL3("So_phieu") 'Số phiếu
        tdbg.Columns("VoucherDate").Caption = rL3("Ngay_phieu") 'Ngày phiếu
        tdbg.Columns("RefDate").Caption = rL3("Ngay_hoa_don")  'Ngày hóa đơn
        tdbg.Columns("SeriNo").Caption = rL3("So_Seri") 'Số Sêri
        tdbg.Columns("RefNo").Caption = rL3("So_hoa_don") 'Số hóa đơn
        tdbg.Columns("Description").Caption = rL3("Dien_giai") 'Diễn giải
        tdbg.Columns("DebitAccountID").Caption = rL3("TK_no") 'rl3("Tai_khoan_no") 'Tài khoản nợ
        tdbg.Columns("CreditAccountID").Caption = rL3("TK_co") 'rl3("Tai_khoan_co") 'Tài khoản có
        tdbg.Columns("CurrencyID").Caption = rL3("Loai_tien") 'Loại tiền
        tdbg.Columns("ExchangeRate").Caption = rL3("Ty_gia") 'Tỷ giá
        tdbg.Columns("OriginalAmount").Caption = rL3("Nguyen_te") 'Nguyên tệ
        tdbg.Columns("ConvertedAmount").Caption = rL3("Quy_doi") 'Qui đổi
        tdbg.Columns("ObjectTypeID").Caption = rL3("Loai_doi_tuong") 'rl3("Ma_loai_doi_tuong") 'Mã loại đối tượng
        tdbg.Columns("ObjectID").Caption = rL3("Ma_doi_tuong") 'Mã đối tượng

        '================================================================ 
        '================================================================ 
        tdbg2.Columns(COL2_CipNo).Caption = rL3("Ma_XDCB") 'Mã XDCB
        tdbg2.Columns(COL2_CipName).Caption = rL3("Ten_ma_XDCB") 'Tên mã XDCB
        tdbg2.Columns(COL2_AccountID).Caption = rL3("TK_XDCB") 'TK XDCB
        tdbg2.Columns(COL2_Desciption).Caption = rL3("Dien_giai") 'Diễn giải
        tdbg2.Columns(COL2_AccountName).Caption = rL3("Ten") 'Tên
        '================================================================ 
        tdbcProjectID.Columns("ProjectID").Caption = rL3("Ma") 'Mã
        tdbcProjectID.Columns("ProjectName").Caption = rL3("Ten") 'Tên
        tdbcTaskID.Columns("TaskID").Caption = rL3("Ma") 'Mã
        tdbcTaskID.Columns("TaskName").Caption = rL3("Ten") 'Tên
        tdbcBudgetID.Columns("BudgetID").Caption = rL3("Ma") 'Mã
        tdbcBudgetID.Columns("BudgetName").Caption = rL3("Ten") 'Tên
        tdbcBudgetItemID.Columns("BudgetItemID").Caption = rL3("Ma") 'Mã
        tdbcBudgetItemID.Columns("BudgetItemName").Caption = rL3("Ten") 'Tên
        '================================================================ 
        lblAccountID.Text = rL3("Ma_tai_khoan") 'Mã tài khoản
        '================================================================ 
        tdbcAccountID.Columns("AccountID").Caption = rL3("Ma") 'Mã
        tdbcAccountID.Columns("AccountName").Caption = rL3("Ten") 'Tên

        '================================================================ 
        tdbdAna01ID.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna01ID.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna02ID.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna02ID.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna03ID.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna03ID.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna04ID.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna04ID.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna05ID.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna05ID.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna06ID.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna06ID.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna07ID.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna07ID.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna08ID.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna08ID.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna09ID.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna09ID.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna10ID.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna10ID.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        '================================================================ 
        tdbg1.Columns("Selected").Caption = rL3("Chon") 'Chọn
        tdbg1.Columns("VoucherTypeID").Caption = rL3("Loai_phieu") 'Loại phiếu
        tdbg1.Columns("VoucherNo").Caption = rL3("So_phieu") 'Số phiếu
        tdbg1.Columns("VoucherDate").Caption = rL3("Ngay_phieu") 'Ngày phiếu
        tdbg1.Columns("RefDate").Caption = rL3("Ngay_hoa_don") 'Ngày hóa đơn
        tdbg1.Columns("SeriNo").Caption = rL3("So_Seri") 'Số Sêri
        tdbg1.Columns("RefNo").Caption = rL3("So_hoa_don") 'Số hóa đơn
        tdbg1.Columns("ObjectTypeID").Caption = rL3("Loai_hoa_don") 'Loại hóa đơn
        tdbg1.Columns("ObjectID").Caption = rL3("Ma_doi_tuong") 'Mã đối tượng
        tdbg1.Columns("Description").Caption = rL3("Dien_giai") 'Diễn giải
        tdbg1.Columns("DebitAccountID").Caption = rL3("TK_no") 'rl3("Tai_khoan_no") 'Tài khoản nợ
        tdbg1.Columns("CreditAccountID").Caption = rL3("TK_co") 'rl3("Tai_khoan_co") 'Tài khoản có
        tdbg1.Columns("CurrencyID").Caption = rL3("Loai_tien") 'Loại tiền
        tdbg1.Columns("ExchangeRate").Caption = rL3("Ty_gia") 'Tỷ giá
        tdbg1.Columns("OriginalAmount").Caption = rL3("Nguyen_te") 'Nguyên tệ
        tdbg1.Columns("ConvertedAmount").Caption = rL3("Quy_doi") 'Qui đổi

        '================================================================ 
        mnsFind1.Text = rL3("Tim__kiem") 'Tìm &kiếm
        mnsListAll1.Text = rL3("_Liet_ke_tat_ca") '&Liệt kê tất cả
        '================================================================ 
        optPrintTypePeriod.Text = rL3("Ky") 'Kỳ
        optPrintTypeDate.Text = rL3("Ngay") 'Ngày

        '================================================================ 
        tsmCollection.Text = rL3("Tap__hop") 'Tập &hợp
        mnsCollection.Text = tsmCollection.Text 'Tập &hợp

        tsmCancelCollection.Text = rL3("_Bo_tap_hop") '&Bỏ tập hợp
        mnsCancelCollection.Text = tsmCancelCollection.Text

        '================================================================ 
        tdbdAna01ID_1.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna01ID_1.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna02ID_1.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna02ID_1.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna03ID_1.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna03ID_1.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna04ID_1.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna04ID_1.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna05ID_1.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna05ID_1.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna06ID_1.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna06ID_1.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna07ID_1.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna07ID_1.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna08ID_1.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna08ID_1.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna09ID_1.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna09ID_1.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục
        tdbdAna10ID_1.Columns("AnaID").Caption = rL3("Ma") 'Mã khoản mục
        tdbdAna10ID_1.Columns("AnaName").Caption = rL3("Ten") 'Tên khoản mục

    End Sub

    Private Sub tdbg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        Select Case tdbg.Col
            Case COL_VoucherDate, COL_RefDate
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
                'Case COL_ConvertedAmount, COL_OriginalAmount, COL_ExchangeRate
                '    e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub

    Private Sub c1dateDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles c1dateDate.KeyDown
        'Fix: khi xóa giá trị sau đó nhấn TAB thì không giữ lại giá trị cũ
        Try
            If e.KeyCode = Keys.Tab Then
                'Chú ý: Nếu cột cuối cùng hiển thị là Date thì không cộng
                tdbg.Col = tdbg.Col + 1
                Exit Sub
            End If
        Catch ex As Exception
        End Try

    End Sub

#Region "Events tdbcTypeCodeID load tdbcCodeID with txtCodeName"

    Private Sub tdbcTypeCodeID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.GotFocus
        'Dùng phím Enter
        tdbcTypeCodeID.Tag = tdbcTypeCodeID.Text
    End Sub

    Private Sub tdbcTypeCodeID_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbcTypeCodeID.MouseDown
        'Di chuyển chuột
        tdbcTypeCodeID.Tag = tdbcTypeCodeID.Text
    End Sub

    Private Sub tdbcTypeCodeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.SelectedValueChanged
        tdbcCodeID.Text = "%"
    End Sub

    Private Sub tdbcTypeCodeID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.LostFocus
        If tdbcTypeCodeID.Tag.ToString = "" And tdbcTypeCodeID.Text = "" Then Exit Sub
        If tdbcTypeCodeID.Tag.ToString = tdbcTypeCodeID.Text And tdbcTypeCodeID.SelectedValue IsNot Nothing Then Exit Sub
        If tdbcTypeCodeID.FindStringExact(tdbcTypeCodeID.Text) = -1 OrElse tdbcTypeCodeID.SelectedValue Is Nothing Then
            tdbcTypeCodeID.Text = ""
            LoadtdbcCodeID("-1")
            tdbcCodeID.Text = ""
            Exit Sub
        End If
        LoadtdbcCodeID(tdbcTypeCodeID.SelectedValue.ToString())
        tdbcCodeID.Text = "%"
    End Sub

    Private Sub tdbcTypeCodeID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcTypeCodeID.KeyDown
        Dim tdbcName As C1.Win.C1List.C1Combo = CType(sender, C1.Win.C1List.C1Combo)
        Select Case e.KeyCode
            Case Keys.A, Keys.D, Keys.E, Keys.I, Keys.O, Keys.U, Keys.Y, Keys.Back
                tdbcName.AutoCompletion = False
            Case Else
                tdbcName.AutoCompletion = True
        End Select
    End Sub

    Private Sub tdbcTypeCodeID_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.Leave
        Dim tdbcName As C1.Win.C1List.C1Combo = CType(sender, C1.Win.C1List.C1Combo)
        If tdbcName.SelectedIndex <> -1 Then
            tdbcName.Text = tdbcName.Columns("TypeCodeName").Text
        End If

    End Sub

    Private Sub tdbcCodeID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcCodeID.GotFocus
        'Dùng phím Enter
        tdbcCodeID.Tag = tdbcCodeID.Text
    End Sub

    Private Sub tdbcCodeID_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbcCodeID.MouseDown
        'Di chuyển chuột
        tdbcCodeID.Tag = tdbcCodeID.Text
    End Sub

    Private Sub tdbcCodeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcCodeID.SelectedValueChanged
        If tdbcCodeID.SelectedValue Is Nothing Then
            txtCodeName.Text = ""
        Else
            txtCodeName.Text = tdbcCodeID.Columns(1).Value.ToString
        End If
    End Sub

    Private Sub tdbcCodeID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcCodeID.LostFocus
        If tdbcCodeID.Tag.ToString = "" And tdbcCodeID.Text = "" Then Exit Sub
        If tdbcCodeID.Tag.ToString = tdbcCodeID.Text And tdbcCodeID.SelectedValue IsNot Nothing Then Exit Sub
    End Sub

#End Region

#Region "Phần cho lưới 1"

    Private dtGrid1 As DataTable
    Dim iLastcol As Integer
    Dim bHeadClick As Boolean = False

#Region "Biến khai báo cho khoản mục"

    Private Const SplitAna As Int16 = 2 ' Ghi nhận Khoản mục chứa ở Split nào
    Dim bUseAna As Boolean 'Kiểm tra có sử dụng Khoản mục không, để set thuộc tính Enabled nút Khoản mục 
#End Region

    Private Sub LoadTDBDropDown()
        '--- Chuẩn Khoản mục b3: Load 10 khoản mục
        LoadTDBDropDownAna(tdbdAna01ID, tdbdAna02ID, tdbdAna03ID, tdbdAna04ID, tdbdAna05ID, tdbdAna06ID, tdbdAna07ID, tdbdAna08ID, tdbdAna09ID, tdbdAna10ID, tdbg1, COL1_Ana01ID, gbUnicode)
        '------------------------------------------

    End Sub

    Private Sub tdbg1_LockedColumns()
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_Selected).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_VoucherTypeID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_VoucherNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_VoucherDate).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_RefDate).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_SeriNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_RefNo).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_ObjectTypeID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_ObjectID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_Description).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_DebitAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_CreditAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_CurrencyID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_ExchangeRate).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_OriginalAmount).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_ConvertedAmount).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_CipID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_TransactionID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT1).DisplayColumns(COL1_ModuleID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
    End Sub

    Private Sub tdbg1_NumberFormat()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbg1.Columns(COL1_ExchangeRate).DataField, DxxFormat.ExchangeRateDecimals, 28, 8)
        AddDecimalColumns(arr, tdbg1.Columns(COL1_OriginalAmount).DataField, DxxFormat.DecimalPlaces, 28, 8)
        AddDecimalColumns(arr, tdbg1.Columns(COL1_ConvertedAmount).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        InputNumber(tdbg1, arr)
    End Sub

#Region "Active Find Client - List All "
    Private sFind1 As String = ""

    Private Sub ReLoadTDBGrid1()
        Dim strFind1 As String = sFind1
        dtGrid1.DefaultView.RowFilter = strFind1
        CheckMenuOther1(gbEnabledUseFind1)
    End Sub

#End Region

    Private Sub LoadTDBGrid1(Optional ByVal sAccountID As String = "")
        Dim sSQL As String
        sSQL = SQLStoreD02P0400(sAccountID)
        dtGrid1 = ReturnDataTable(sSQL)
        LoadDataSource(tdbg1, dtGrid1, gbUnicode)
        ReLoadTDBGrid1()
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0400
    '# Created User: KIM LONG
    '# Created Date: 30/05/2016 02:20:59
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0400(Optional ByVal sAccountID As String = "") As String
        Dim sSQL As String = ""
        sSQL &= ("-- -- Do nguon cho luoi tap hop 1" & vbCrLf)
        sSQL &= "Exec D02P0400 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(sAccountID) & COMMA 'CollectAccountID, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'strFind, varchar[8000], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcProjectID)) & COMMA 'ProjectID, nvarchar[100], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcTaskID)) & COMMA 'TaskID, nvarchar[100], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcBudgetID)) & COMMA 'BudgetID, nvarchar[100], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcBudgetItemID)) & COMMA 'BudgetItemID, nvarchar[100], NOT NULL
        sSQL &= SQLNumber(IIf(optPrintTypeDate.Checked, 0, 1)) & COMMA 'IsTime, tinyint, NOT NULL
        sSQL &= SQLNumber(ReturnValueC1Combo(tdbcPeriodFrom, "TranMonth")) & COMMA 'FromMonth, int, NOT NULL
        sSQL &= SQLNumber(ReturnValueC1Combo(tdbcPeriodFrom, "TranYear")) & COMMA 'FromYear, int, NOT NULL
        sSQL &= SQLNumber(ReturnValueC1Combo(tdbcPeriodTo, "TranMonth")) & COMMA 'ToMonth, int, NOT NULL
        sSQL &= SQLNumber(ReturnValueC1Combo(tdbcPeriodTo, "TranYear")) & COMMA 'ToYear, int, NOT NULL
        sSQL &= SQLDateSave(c1dateDateFrom.Value) & COMMA 'FromDate, datetime, NOT NULL
        sSQL &= SQLDateSave(c1dateDateTo.Value) 'ToDate, datetime, NOT NULL
        Return sSQL
    End Function

    'Tập hợp
    Private Sub btnCollection_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCollection.Click
        tdbg.UpdateData()
        If Not AllowSave() Then Exit Sub
        'Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        If iMode <> 0 Then
            If IsExistKey("D02T0100", "CipNo", ReturnValueC1Combo(tdbcProjectID)) Then
                D99C0008.MsgDuplicatePKey()
                Exit Sub
            Else
                sSQL.Append(SQLInsertD02T0100(iMode).ToString & vbCrLf)
            End If
        End If
        sSQL.Append(SQLUpdateD02T0012s(iMode))
        btnCollection.Enabled = False
        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        'Me.Cursor = Cursors.Default
        gbSavedOK = False
        If bRunSQL Then
            SaveOK()
            gbSavedOK = True
            btnCollection.Enabled = True
            SplitContainer1.Panel2Collapsed = True
            SplitContainer1.Panel1Collapsed = False
            If iMode <> 0 Then
                LoadTDBGrid2(ReturnValueC1Combo(tdbcTypeCodeID), ReturnValueC1Combo(tdbcCodeID), , sCipID)
            Else
                LoadTDBGrid(tdbg2.Columns(COL2_CipID).Text, tdbg2.Columns(COL2_AccountID).Text)
            End If

        Else
            SaveNotOK()
            btnCollection.Enabled = True
        End If
        If tdbg2.Columns("Status").Text <> "" Then
            If CInt(tdbg2.Columns("Status").Text) = 0 Then
                sSQL.Append(SQLUpdateD02T0100(iMode))
                ExecuteSQL(sSQL.ToString)
            End If
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
    Private Function SQLUpdateD02T0012s(ByVal iMode As Integer) As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        For i As Integer = 0 To tdbg1.RowCount - 1
            If Not IsDBNull(tdbg1(i, COL1_Selected)) And tdbg1(i, COL1_Selected).ToString <> "" Then
                If CBool(tdbg1(i, COL1_Selected)) = True Then
                    sSQL.Append("Update D02T0012 Set ")
                    '  sSQL.Append("Status = " & "1" & COMMA) 'tinyint, NOT NULL
                    sSQL.Append("Ana01ID = " & SQLString(tdbg1(i, COL1_Ana01ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana02ID = " & SQLString(tdbg1(i, COL1_Ana02ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana03ID = " & SQLString(tdbg1(i, COL1_Ana03ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana04ID = " & SQLString(tdbg1(i, COL1_Ana04ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana05ID = " & SQLString(tdbg1(i, COL1_Ana05ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana06ID = " & SQLString(tdbg1(i, COL1_Ana06ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana07ID = " & SQLString(tdbg1(i, COL1_Ana07ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana08ID = " & SQLString(tdbg1(i, COL1_Ana08ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana09ID = " & SQLString(tdbg1(i, COL1_Ana09ID)) & COMMA) 'varchar[20], NULL
                    sSQL.Append("Ana10ID = " & SQLString(tdbg1(i, COL1_Ana10ID)) & COMMA) 'varchar[20], NULL
                    If iMode <> 0 Then
                        sSQL.Append("CipID = " & SQLString(sCipID)) 'varchar[20], NULL
                    Else
                        sSQL.Append("CipID = " & SQLString(tdbg2.Columns("CipID").Text)) 'varchar[20], NULL
                    End If

                    sSQL.Append(" Where ")
                    sSQL.Append("TransactionID = " & SQLString(tdbg1(i, COL1_TransactionID)) & " And ")
                    sSQL.Append("DivisionID = " & SQLString(gsDivisionID))
                    sRet.Append(sSQL.ToString & vbCrLf)
                    sSQL.Remove(0, sSQL.Length)
                End If
            End If

        Next
        Return sRet
    End Function

    Private Function AllowFilter1() As Boolean
        If optPrintTypeDate.Checked Then
            If Not CheckValidDateFromTo(c1dateDateFrom, c1dateDateTo) Then
                Return False
            End If
        End If
        If optPrintTypePeriod.Checked Then
            If Not CheckValidPeriodFromTo(tdbcPeriodFrom, tdbcPeriodTo) Then
                Return False
            End If
        End If
        If iMode = 1 Then
            If tdbcProjectID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblProjectID.Text)
                tdbcProjectID.Focus()
                Return False
            End If
            If tdbcAccountID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblAccountID.Text)
                tdbcAccountID.Focus()
                Return False
            End If
        ElseIf iMode = 2 Then
            'If tdbcAccountID.Text.Trim = "" Then
            '    D99C0008.MsgNotYetChoose(lblAccountID.Text)
            '    tdbcAccountID.Focus()
            '    Return False
            'End If
            If tdbcBudgetID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblBudgetID.Text)
                tdbcBudgetID.Focus()
                Return False
            End If
        End If
        Return True
    End Function

    Private Function AllowSave() As Boolean
        Dim iCount As Integer = 0
        If iMode = 1 Then
            If tdbcProjectID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblProjectID.Text)
                tdbcProjectID.Focus()
                Return False
            End If
            If tdbcAccountID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblAccountID.Text)
                tdbcAccountID.Focus()
                Return False
            End If
        ElseIf iMode = 2 Then
            'If tdbcAccountID.Text.Trim = "" Then
            '    D99C0008.MsgNotYetChoose(lblAccountID.Text)
            '    tdbcAccountID.Focus()
            '    Return False
            'End If
            If tdbcBudgetID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(lblBudgetID.Text)
                tdbcBudgetID.Focus()
                Return False
            End If
        End If

        If tdbg1.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tdbg1.Focus()
            Return False
        End If
        For i As Integer = 0 To tdbg1.RowCount - 1
            If Not IsDBNull(tdbg1(i, COL1_Selected)) And tdbg1(i, COL1_Selected).ToString <> "" Then
                If CBool(tdbg1(i, COL1_Selected)) = False Then
                    iCount += 1
                End If
            End If
        Next
        If iCount = tdbg1.RowCount Then
            D99C0008.MsgL3(rL3("Ban_chua_chon_dong_nao_de_tap_hop"))
            tdbg1.SplitIndex = SPLIT0
            tdbg1.Col = COL1_Selected
            tdbg1.Bookmark = 0
            tdbg1.Focus()
            Return False
        End If
       
        Return True
    End Function

    Private Sub SetBackColorObligatory()
        If iMode <> 0 Then
            tdbcProjectID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
            tdbcBudgetID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
            tdbcAccountID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        Else
            tdbcProjectID.EditorBackColor = System.Drawing.SystemColors.Window
            tdbcBudgetID.EditorBackColor = System.Drawing.SystemColors.Window
            tdbcAccountID.EditorBackColor = System.Drawing.SystemColors.Window
        End If
    End Sub

    Private Sub tdbg1_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg1.HeadClick
        Select Case tdbg.Col
            Case COL1_Selected
                CheckedAll()
        End Select
    End Sub

    Private Sub tdbg1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg1.KeyDown
        If e.KeyCode = Keys.Enter Then
            If tdbg1.Col = iLastcol Then
                HotKeyEnterGrid(tdbg1, COL1_Selected, e)
            End If
        End If
        If e.Control And e.KeyCode = Keys.S Then
            tdbg1_HeadClick(Nothing, Nothing)
        End If
    End Sub

    Private Sub CheckedAll()
        bHeadClick = Not bHeadClick
        For i As Integer = 0 To tdbg1.RowCount - 1
            tdbg(i, COL1_Selected) = bHeadClick
        Next
    End Sub

    Private Sub tdbg1_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg1.BeforeColUpdate
        Select Case e.ColIndex
            Case COL1_ModuleID
                '--- Chuẩn Khoản mục b5: Kiểm tra Khoản mục lúc nhập liệu
                '---------------------------------------------
            Case COL1_Ana01ID
                If tdbg1.Columns(COL1_Ana01ID).Text <> tdbdAna01ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(0) Then 'Kiểm tra nhập trong danh sách
                        tdbg1.Columns(COL1_Ana01ID).Text = ""
                    Else
                        If tdbg1.Columns(COL1_Ana01ID).Text.Length > giArrAnaLength(0) Then ' Kiểm tra chiều dài nhập vào
                            tdbg1.Columns(COL1_Ana01ID).Text = ""
                        End If
                    End If
                End If

            Case COL1_Ana02ID
                If tdbg1.Columns(COL1_Ana02ID).Text <> tdbdAna02ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(1) Then 'Kiểm tra nhập trong danh sách
                        tdbg1.Columns(COL1_Ana02ID).Text = ""
                    Else
                        If tdbg1.Columns(COL1_Ana02ID).Text.Length > giArrAnaLength(1) Then ' Kiểm tra chiều dài nhập vào
                            tdbg1.Columns(COL1_Ana02ID).Text = ""
                        End If
                    End If
                End If
            Case COL1_Ana03ID
                If tdbg1.Columns(COL1_Ana03ID).Text <> tdbdAna03ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(2) Then 'Kiểm tra nhập trong danh sách
                        tdbg1.Columns(COL1_Ana03ID).Text = ""
                    Else
                        If tdbg1.Columns(COL1_Ana03ID).Text.Length > giArrAnaLength(2) Then ' Kiểm tra chiều dài nhập vào
                            tdbg1.Columns(COL1_Ana03ID).Text = ""
                        End If
                    End If
                End If
            Case COL1_Ana04ID
                If tdbg1.Columns(COL1_Ana04ID).Text <> tdbdAna04ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(3) Then 'Kiểm tra nhập trong danh sách
                        tdbg1.Columns(COL1_Ana04ID).Text = ""
                    Else
                        If tdbg1.Columns(COL1_Ana04ID).Text.Length > giArrAnaLength(3) Then ' Kiểm tra chiều dài nhập vào
                            tdbg1.Columns(COL1_Ana04ID).Text = ""
                        End If
                    End If
                End If
            Case COL1_Ana05ID
                If tdbg1.Columns(COL1_Ana05ID).Text <> tdbdAna05ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(4) Then 'Kiểm tra nhập trong danh sách
                        tdbg1.Columns(COL1_Ana05ID).Text = ""
                    Else
                        If tdbg1.Columns(COL1_Ana05ID).Text.Length > giArrAnaLength(4) Then ' Kiểm tra chiều dài nhập vào
                            tdbg1.Columns(COL1_Ana05ID).Text = ""
                        End If
                    End If
                End If
            Case COL1_Ana06ID
                If tdbg1.Columns(COL1_Ana06ID).Text <> tdbdAna06ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(5) Then 'Kiểm tra nhập trong danh sách
                        tdbg1.Columns(COL1_Ana06ID).Text = ""
                    Else
                        If tdbg1.Columns(COL1_Ana06ID).Text.Length > giArrAnaLength(5) Then ' Kiểm tra chiều dài nhập vào
                            tdbg1.Columns(COL1_Ana06ID).Text = ""
                        End If
                    End If
                End If
            Case COL1_Ana07ID
                If tdbg1.Columns(COL1_Ana07ID).Text <> tdbdAna07ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(6) Then 'Kiểm tra nhập trong danh sách
                        tdbg1.Columns(COL1_Ana07ID).Text = ""
                    Else
                        If tdbg1.Columns(COL1_Ana07ID).Text.Length > giArrAnaLength(6) Then ' Kiểm tra chiều dài nhập vào
                            tdbg1.Columns(COL1_Ana07ID).Text = ""
                        End If
                    End If
                End If
            Case COL1_Ana08ID
                If tdbg1.Columns(COL1_Ana08ID).Text <> tdbdAna08ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(7) Then 'Kiểm tra nhập trong danh sách
                        tdbg1.Columns(COL1_Ana08ID).Text = ""
                    Else
                        If tdbg1.Columns(COL1_Ana08ID).Text.Length > giArrAnaLength(7) Then ' Kiểm tra chiều dài nhập vào
                            tdbg1.Columns(COL1_Ana08ID).Text = ""
                        End If
                    End If
                End If
            Case COL1_Ana09ID
                If tdbg1.Columns(COL1_Ana09ID).Text <> tdbdAna09ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(8) Then 'Kiểm tra nhập trong danh sách
                        tdbg1.Columns(COL1_Ana09ID).Text = ""
                    Else
                        If tdbg1.Columns(COL1_Ana09ID).Text.Length > giArrAnaLength(8) Then ' Kiểm tra chiều dài nhập vào
                            tdbg1.Columns(COL1_Ana09ID).Text = ""
                        End If
                    End If
                End If
            Case COL1_Ana10ID
                If tdbg1.Columns(COL1_Ana10ID).Text <> tdbdAna10ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(9) Then 'Kiểm tra nhập trong danh sách
                        tdbg1.Columns(COL1_Ana10ID).Text = ""
                    Else
                        If tdbg1.Columns(COL1_Ana10ID).Text.Length > giArrAnaLength(9) Then ' Kiểm tra chiều dài nhập vào
                            tdbg1.Columns(COL1_Ana10ID).Text = ""
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
    Private Function SQLUpdateD02T0100(ByVal iMode As Integer) As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0100 Set ")
        sSQL.Append("Status = " & SQLNumber(1)) 'tinyint, NULL
        sSQL.Append(" Where ")
        If iMode <> 0 Then
            sSQL.Append("CipID = " & SQLString(sCipID))
        Else
            sSQL.Append("CipID = " & SQLString(tdbg2.Columns("CipID").Text))
        End If

        Return sSQL
    End Function

#End Region

#Region "Phần cho lưới 2"

    Dim sFilter2 As New System.Text.StringBuilder()
    Dim bRefreshFilter2 As Boolean = False
    Private Sub tdbg2_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg2.FilterChange
        Try
            If (dtGrid2 Is Nothing) Then Exit Sub
            If bRefreshFilter2 Then Exit Sub
            FilterChangeGrid(tdbg2, sFilter2) 'Nếu có Lọc khi In
            ReLoadTDBGrid2()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

    Private Sub ReLoadTDBGrid2()
        Dim strFind As String = ""
        If sFilter2.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter2.ToString
        dtGrid2.DefaultView.RowFilter = strFind
        FooterTotalGrid(tdbg2, COL2_CipNo)
        'ResetGrid()
        CheckOtherMenu2()
        '17/8/2019, id 122976-FSC - Lỗi không refresh lại chứng từ theo mã khi dán mã vào cột filter hoặc tìm mã tương ứng
        If tdbg2.RowCount > 0 Then
            LoadTDBGrid(tdbg2.Columns(COL2_CipID).Text, tdbg2.Columns(COL2_AccountID).Text)
        Else
            If dtGrid IsNot Nothing Then dtGrid.Clear()
            ResetGrid()
        End If
    End Sub

    Private Sub tdbg2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg2.KeyDown
        Me.Cursor = Cursors.WaitCursor
        HotKeyCtrlVOnGrid(tdbg2, e)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tdbg2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg2.KeyPress
        If tdbg2.Columns(tdbg2.Col).ValueItems.Presentation = C1.Win.C1TrueDBGrid.PresentationEnum.CheckBox Then
            e.Handled = CheckKeyPress(e.KeyChar)
        ElseIf tdbg2.Splits(tdbg2.SplitIndex).DisplayColumns(tdbg2.Col).Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far Then
            'e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End If
    End Sub

#Region "Events tdbcBudgetID with txtBudgetName load tdbcBudgetItemID with txtBudgetItemName"

    Private Sub tdbcBudgetID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcBudgetID.SelectedValueChanged
        If tdbcBudgetID.SelectedValue Is Nothing OrElse tdbcBudgetID.Text = "" Then
            tdbcBudgetID.Text = ""
            LoadBudgetItem(tdbcBudgetItemID, Me.Name, , "-1")
            tdbcBudgetItemID.Text = ""

        Else
            LoadBudgetItem(tdbcBudgetItemID, Me.Name, , tdbcBudgetID.SelectedValue.ToString())
            tdbcBudgetItemID.Text = ""
        End If
    End Sub

    Private Sub tdbcBudgetID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcBudgetID.LostFocus
        If tdbcBudgetID.FindStringExact(tdbcBudgetID.Text) = -1 Then
            tdbcBudgetID.Text = ""
            LoadBudgetItem(tdbcBudgetItemID, Me.Name, , "-1")
            tdbcBudgetItemID.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub tdbcBudgetItemID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcBudgetItemID.LostFocus
        If tdbcBudgetItemID.FindStringExact(tdbcBudgetItemID.Text) = -1 Then tdbcBudgetItemID.Text = ""
    End Sub

#End Region

#Region "Events tdbcProjectID with txtProjectName load tdbcTaskID with txtTaskName"

    Private Sub tdbcProjectID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcProjectID.SelectedValueChanged
        If tdbcProjectID.SelectedValue Is Nothing OrElse tdbcProjectID.Text = "" Then
            tdbcProjectID.Text = ""
            LoadTask(tdbcTaskID, , "-1")
            tdbcTaskID.Text = ""
        Else
            LoadTask(tdbcTaskID, , tdbcProjectID.SelectedValue.ToString())
            tdbcTaskID.Text = ""
        End If
    End Sub

    Private Sub tdbcProjectID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcProjectID.LostFocus
        If tdbcProjectID.FindStringExact(tdbcProjectID.Text) = -1 Then
            tdbcProjectID.Text = ""
            LoadTask(tdbcTaskID, , "-1")
            tdbcTaskID.Text = ""
            Exit Sub
        End If
    End Sub

    Private Sub tdbcTaskID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcTaskID.LostFocus
        If tdbcTaskID.FindStringExact(tdbcTaskID.Text) = -1 Then tdbcTaskID.Text = ""
    End Sub

#End Region
#End Region

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        SplitContainer1.Panel2Collapsed = True
        SplitContainer1.Panel1Collapsed = False
    End Sub

    Private Sub tdbg2_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbg2.RowColChange
        If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
        'Neu luoi co 1 dong thi k can chay su kien nay
        If tdbg2.RowCount <= 1 Then Exit Sub
        If e.LastRow = tdbg2.Row Then Exit Sub
        LoadTDBGrid(tdbg2.Columns(COL2_CipID).Text, tdbg2.Columns(COL2_AccountID).Text)

    End Sub

    Private Sub CheckMenuOther1(ByVal UsedFind As Boolean)
        mnsFind1.Enabled = UsedFind Or tdbg1.RowCount > 0
        mnsListAll1.Enabled = UsedFind Or tdbg1.RowCount > 0
    End Sub

    Private Sub CheckOtherMenu2()
        mnsAutoProject.Enabled = (D02Systems.UseD54ForCIP = 1)
        mnsAutoBudget.Enabled = (D02Systems.UseBudgetForCIP = 1)
        mnsPrint.Enabled = tdbg2.RowCount > 0 AndAlso ReturnPermission("D02F1003") > 1
    End Sub

    Private Sub mnsCancelCollection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnsCancelCollection.Click, tsmCancelCollection.Click

        If AskDelete() = Windows.Forms.DialogResult.Yes Then
            If Not AllowDelete() Then Exit Sub
            ''Kiểm tra điều kiện xóa
            ''Nếu ModuleID <> 02
            Dim sSQL As String = ""
            If tdbg.Columns(COL_ModuleID).Text <> "02" Then
                sSQL = SQLUpdateD02T0012.ToString & vbCrLf
            Else
                'Sửa ngày 19/7/2012 theo incidetn 47959 của HOANGNAM bởi VANVINH
                If tdbg.Columns(COL_TransactionTypeID).Text = "SDXDCB" Then
                    D99C0008.MsgL3(rL3("Ban_khong_the_xoa_But_toan_Nhap_so_du_chi_phi_XDCB_Ban_can_sang_man_hinh_Nhap_so_du_chi_phi_XDCB_-_D02F1007"))
                    Exit Sub
                End If

                If tdbg.Columns(COL_Internal).Text = "0" Then
                    If (tdbg.Columns(COL_Status).Text = "0" Or tdbg.Columns(COL_Status).Text = "1") And tdbg.Columns(COL_TransactionTypeID).Text <> "SDXDCB" Then 'Nhập chứng từ, tách chi phí 
                        sSQL = SQLUpdateD02T0012.ToString & vbCrLf
                    Else 'Số dư
                        'If D99C0008.MsgAsk("Bạn không thể xóa bút toán nhập số dư XDCB." & Space(1) & "Bạn cần sang màn hình Nhập số dư XDCB - D02F1007." & Space(1) & "Bạn có muốn xem không?", MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                        '    tsbView_Click(Nothing, Nothing)
                        'End If
                        Exit Sub
                    End If
                Else
                    If tdbg.Columns(COL_Status).Text = "1" Then 'Nhập chứng từ, tách chi phí 
                        If D99C0008.MsgAsk("Phiếu này đã được xử lý." & Space(1) & "Bạn không thể xóa phiếu này được." & Space(1) & "Bạn có muốn xem không?", MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                            tsbView_Click(Nothing, Nothing)
                        End If
                        Exit Sub
                    Else 'Số dư
                        sSQL = SQLDeleteD02T0012()
                    End If
                End If
            End If
            sSQL &= "--Cap nhat trang thai ma XDCB" & vbCrLf
            sSQL &= "If Not Exists (Select Top 1 1 From D02T0012 Where CipID=" & SQLString(tdbg2.Columns("CipID").Text) & ")" & vbCrLf
            sSQL &= "Begin" & vbCrLf
            sSQL &= "Update D02T0100 Set Status=0 Where CipID=" & SQLString(tdbg2.Columns("CipID").Text) & vbCrLf 'Status = 0 là bỏ tập hợp
            sSQL &= "End"

            Dim bResult As Boolean = ExecuteSQL(sSQL)
            If bResult = True Then
                'Kiểm tra Audit và thiết lập Auditlog
                'If byAudit = 1 Then
                '    ExecuteAuditLog(sAuditCode, "03", tdbcCipNo.Columns("CipID").Text, txtCipName.Text)
                'End If
                'ExecuteAuditLog(sAuditCode, "03", tdbg2.Columns("CipID").Text, tdbg.Columns(COL_VoucherNo).Text, tdbg.Columns(COL_VoucherDate).Text)
                Lemon3.D91.RunAuditLog("02", sAuditCode, "03", tdbg2.Columns("CipID").Text, tdbg.Columns(COL_VoucherNo).Text, tdbg.Columns(COL_VoucherDate).Text)
                DeleteOK()
                'DeleteVoucherNoD91T9111(tdbg.Columns(COL_VoucherNo).Text, "D02T0012", "VoucherNo")
                DeleteGridEvent(tdbg, dtGrid, gbEnabledUseFind)
                LoadTDBGrid(tdbg2.Columns(COL2_CipID).Text, tdbg2.Columns(COL2_AccountID).Text)
            Else
                DeleteNotOK()
            End If
        End If
    End Sub

    Private Sub btnFilter1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter1.Click
        If Not AllowFilter1() Then Exit Sub
        btnFilter1.Focus()
        If btnFilter1.Focused = False Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        
        If tdbcAccountID.Visible Then
            If tdbcAccountID.SelectedValue Is Nothing Then
                LoadTDBGrid1("%")
            Else
                LoadTDBGrid1(tdbcAccountID.SelectedValue.ToString)
            End If
        Else
            LoadTDBGrid1(tdbg2.Columns(COL2_AccountID).Text)
        End If

        Me.Cursor = Cursors.Default
    End Sub
    Dim iMode As Integer = 0
    Private Sub mnsAutoProject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnsAutoProject.Click
        Me.Cursor = Cursors.WaitCursor
        iMode = 1
        loadNewCollection()
        SetBackColorObligatory()
        bUseAna = LoadTDBGridAnalysisCaption(D02, tdbg1, COL1_Ana01ID, SplitAna, True, gbUnicode)

        If Not bUseAna Then
            tdbg1.Splits(SplitAna).SplitSize = 0
            tdbg1.Splits(SplitAna).HScrollBar.Style = C1.Win.C1TrueDBGrid.ScrollBarStyleEnum.None '19/11/2018, id 115915-Lỗi xem phiếu tập hợp chi phí xây dựng cơ bản
        Else
            tdbg1.Splits(SplitAna).HScrollBar.Style = C1.Win.C1TrueDBGrid.ScrollBarStyleEnum.Always '19/11/2018, id 115915-Lỗi xem phiếu tập hợp chi phí xây dựng cơ bản
        End If

        ResetSplitDividerSize(tdbg1)
        LoadTDBDropDown()
        gbEnabledUseFind1 = False
        ResetColorGrid(tdbg1, 1)
        lblAccountID.Visible = True
        tdbcAccountID.Visible = True
        If D02Systems.UseD54ForCIP = 0 AndAlso D02Systems.UseBudgetForCIP = 0 Then
            pnlBudget.Visible = False
            pnlProject.Visible = False
            'btnFilter1.Visible = False
            'btnFilter1.Location = pnlProject.Location
        Else
            If D02Systems.UseD54ForCIP = 1 AndAlso D02Systems.UseBudgetForCIP = 0 Then
                pnlBudget.Visible = False
                pnlProject.Visible = True
                
            ElseIf D02Systems.UseD54ForCIP = 0 AndAlso D02Systems.UseBudgetForCIP = 1 Then
                pnlProject.Visible = False
                pnlBudget.Visible = True
                pnlBudget.Location = pnlProject.Location
            End If
        End If

        tdbg1_LockedColumns()
        tdbg1_NumberFormat()
        btnCollection.Enabled = ReturnPermission("D02F1003") > EnumPermission.View
        iLastcol = CountCol(tdbg1, 1)
        SplitContainer1.Panel1Collapsed = True
        SplitContainer1.Panel2Collapsed = False
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub mnsAutoBudget_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnsAutoBudget.Click
        Me.Cursor = Cursors.WaitCursor
        iMode = 2
        loadNewCollection()
        SetBackColorObligatory()
        bUseAna = LoadTDBGridAnalysisCaption(D02, tdbg1, COL1_Ana01ID, SplitAna, True, gbUnicode)

        If Not bUseAna Then
            tdbg1.Splits(SplitAna).SplitSize = 0
            tdbg1.Splits(SplitAna).HScrollBar.Style = C1.Win.C1TrueDBGrid.ScrollBarStyleEnum.None '19/11/2018, id 115915-Lỗi xem phiếu tập hợp chi phí xây dựng cơ bản
        Else
            tdbg1.Splits(SplitAna).HScrollBar.Style = C1.Win.C1TrueDBGrid.ScrollBarStyleEnum.Always '19/11/2018, id 115915-Lỗi xem phiếu tập hợp chi phí xây dựng cơ bản
        End If

        ResetSplitDividerSize(tdbg1)
        LoadTDBDropDown()
        gbEnabledUseFind1 = False
        ResetColorGrid(tdbg1, 1)
        lblAccountID.Visible = True
        tdbcAccountID.Visible = True
        If D02Systems.UseD54ForCIP = 0 AndAlso D02Systems.UseBudgetForCIP = 0 Then
            pnlBudget.Visible = False
            pnlProject.Visible = False
            btnFilter1.Visible = False
            btnFilter1.Location = pnlProject.Location
        Else
            If D02Systems.UseD54ForCIP = 1 AndAlso D02Systems.UseBudgetForCIP = 0 Then
                pnlBudget.Visible = False
                pnlProject.Visible = True
            ElseIf D02Systems.UseD54ForCIP = 0 AndAlso D02Systems.UseBudgetForCIP = 1 Then
                pnlProject.Visible = False
                pnlBudget.Visible = True
                pnlBudget.Location = pnlProject.Location
            End If
        End If

        tdbg1_LockedColumns()
        tdbg1_NumberFormat()
        btnCollection.Enabled = ReturnPermission("D02F1003") > EnumPermission.View
        iLastcol = CountCol(tdbg1, 1)
        SplitContainer1.Panel1Collapsed = True
        SplitContainer1.Panel2Collapsed = False
        Me.Cursor = Cursors.Default
    End Sub

    Dim sCipID As String
    Private Function SQLInsertD02T0100(ByVal iMode As Integer) As StringBuilder
        sCipID = CreateIGE("D02T0100", "CipID", "02", "CI", gsStringKey)
        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T0100(")
        sSQL.Append("CipID, CipNo, DescriptionU, CipNameU, ")
        sSQL.Append("AccountID, Disabled, Status, StartDate, ")
        sSQL.Append("CompletionDate, CreateDate, CreateUserID, LastModifyDate, ")
        sSQL.Append("LastModifyUserID, X01ID, X02ID, X03ID, X04ID, ")
        sSQL.Append("X05ID, X06ID, X07ID, X08ID, X09ID, ")
        sSQL.Append("X10ID, DivisionID, ContractorOTID, ContractorID, SupplierOTID, ")
        sSQL.Append("SupplierID, ExpStartDate, ExpEndDate, CipCost, ")
        sSQL.Append("CIPNum01, CIPNum02, CIPNum03, CIPNum04, ")
        sSQL.Append("CIPNum05, CIPNum06, CIPNum07, CIPNum08, CIPNum09, ")
        sSQL.Append("CIPNum10, CIPDate01, CIPDate02, CIPDate03, CIPDate04, ")
        sSQL.Append("CIPDate05, CIPDate06, CIPDate07, CIPDate08, CIPDate09, ")
        sSQL.Append("CIPDate10,")
        sSQL.Append("CIPString01U, CIPString02U, CIPString03U, CIPString04U,CIPString05U, CIPString06U, CIPString07U, CIPString08U, CIPString09U, ")
        sSQL.Append("CIPString10U,CIPObjectTypeID,CIPObjectID,CIPEmployeeID")

        sSQL.Append(") Values(")
        sSQL.Append(SQLString(sCipID) & COMMA) 'CipID [KEY], varchar[20], NOT NULL
        If iMode = 1 Then
            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcProjectID)) & COMMA) 'CipNo, varchar[20], NULL
            sSQL.Append(SQLStringUnicode(ReturnValueC1Combo(tdbcProjectID, "ProjectName"), gbUnicode, True) & COMMA) 'Description, varchar[250], NULL
            sSQL.Append(SQLStringUnicode(ReturnValueC1Combo(tdbcProjectID, "ProjectName"), gbUnicode, True) & COMMA) 'CipName, varchar[250], NULL
        ElseIf iMode = 2 Then
            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcBudgetID)) & COMMA) 'CipNo, varchar[20], NULL
            sSQL.Append(SQLStringUnicode(ReturnValueC1Combo(tdbcBudgetID, "BudgetName"), gbUnicode, True) & COMMA) 'Description, varchar[250], NULL
            sSQL.Append(SQLStringUnicode(ReturnValueC1Combo(tdbcBudgetID, "BudgetName"), gbUnicode, True) & COMMA) 'CipName, varchar[250], NULL
        End If
       
        sSQL.Append(SQLString(ReturnValueC1Combo(tdbcAccountID)) & COMMA) 'AccountID, varchar[20], NULL
        sSQL.Append(SQLNumber(0) & COMMA) 'Disabled, tinyint, NULL
        'sSQL.Append(SQLNumber(0) & COMMA) 'Status, tinyint, NULL 'theo màn hình cũ D02F1005 là status = 1 là đã tập hợp
        sSQL.Append(SQLNumber(1) & COMMA) 'theo màn hình cũ D02F1005 là status = 1 là đã tập hợp
        sSQL.Append(SQLDateSave("") & COMMA) 'StartDate, datetime, NULL
        sSQL.Append(SQLDateSave("") & COMMA) 'CompletionDate, datetime, NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'X01ID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'X02ID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'X03ID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'X04ID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'X05ID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'X06ID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'X07ID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'X08ID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'X09ID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'X10ID, varchar[20], NULL
        sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NULL
        sSQL.Append(SQLString("") & COMMA) 'ContractorOTID, varchar[20], NOT NULL
        sSQL.Append(SQLString("") & COMMA) 'ContractorID, varchar[20], NOT NULL
        sSQL.Append(SQLString("") & COMMA) 'SupplierOTID, varchar[20], NOT NULL
        sSQL.Append(SQLString("") & COMMA) 'SupplierID, varchar[20], NOT NULL
        sSQL.Append(SQLDateSave("") & COMMA) 'ExpStartDate, datetime, NULL
        sSQL.Append(SQLDateSave("") & COMMA) 'ExpEndDate, datetime, NULL
        sSQL.Append(SQLMoney(0) & COMMA) 'CipCost, decimal, NOT NULL
        sSQL.Append(SQLMoney("") & COMMA) 'CIPNum01, money, NOT NULL
        sSQL.Append(SQLMoney("") & COMMA) 'CIPNum02, money, NOT NULL
        sSQL.Append(SQLMoney("") & COMMA) 'CIPNum03, money, NOT NULL
        sSQL.Append(SQLMoney("") & COMMA) 'CIPNum04, money, NOT NULL
        sSQL.Append(SQLMoney("") & COMMA) 'CIPNum05, money, NOT NULL
        sSQL.Append(SQLMoney("") & COMMA) 'CIPNum06, money, NOT NULL
        sSQL.Append(SQLMoney("") & COMMA) 'CIPNum07, money, NOT NULL
        sSQL.Append(SQLMoney("") & COMMA) 'CIPNum08, money, NOT NULL
        sSQL.Append(SQLMoney("") & COMMA) 'CIPNum09, money, NOT NULL
        sSQL.Append(SQLMoney("") & COMMA) 'CIPNum10, money, NOT NULL
        sSQL.Append(SQLDateSave("") & COMMA) 'CIPDate01, datetime, NULL
        sSQL.Append(SQLDateSave("") & COMMA) 'CIPDate02, datetime, NULL
        sSQL.Append(SQLDateSave("") & COMMA) 'CIPDate03, datetime, NULL
        sSQL.Append(SQLDateSave("") & COMMA) 'CIPDate04, datetime, NULL
        sSQL.Append(SQLDateSave("") & COMMA) 'CIPDate05, datetime, NULL
        sSQL.Append(SQLDateSave("") & COMMA) 'CIPDate06, datetime, NULL
        sSQL.Append(SQLDateSave("") & COMMA) 'CIPDate07, datetime, NULL
        sSQL.Append(SQLDateSave("") & COMMA) 'CIPDate08, datetime, NULL
        sSQL.Append(SQLDateSave("") & COMMA) 'CIPDate09, datetime, NULL
        sSQL.Append(SQLDateSave("") & COMMA) 'CIPDate10, datetime, NULL
        sSQL.Append(SQLStringUnicode("", gbUnicode, True) & COMMA) 'CIPString01U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode("", gbUnicode, True) & COMMA) 'CIPString02U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode("", gbUnicode, True) & COMMA) 'CIPString03U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode("", gbUnicode, True) & COMMA) 'CIPString04U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode("", gbUnicode, True) & COMMA) 'CIPString05U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode("", gbUnicode, True) & COMMA) 'CIPString06U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode("", gbUnicode, True) & COMMA) 'CIPString07U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode("", gbUnicode, True) & COMMA) 'CIPString08U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode("", gbUnicode, True) & COMMA) 'CIPString09U, varchar[1000], NOT NULL
        sSQL.Append(SQLStringUnicode("", gbUnicode, True) & COMMA) 'CIPString10U, varchar[1000], NOT NULL
        sSQL.Append(SQLString("") & COMMA) 'SupplierID, varchar[20], NOT NULL
        sSQL.Append(SQLString("") & COMMA) 'SupplierID, varchar[20], NOT NULL
        sSQL.Append(SQLString("")) 'SupplierID, varchar[20], NOT NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    Private Sub UseEnterAsTab(ByVal frm As Form, Optional ByVal bForward As Boolean = True)
        Try
            Select Case frm.ActiveControl.GetType.Name
                Case "GridEditor", "C1TrueDBGrid" ' Không làm
                Case "SplitContainer"
                    Dim SplitCon As SplitContainer = CType(frm.ActiveControl, SplitContainer)
                    UseEnterAsTab(SplitCon, bForward)
                Case Else
                    frm.SelectNextControl(frm.ActiveControl, bForward, True, True, False)
            End Select
        Catch ex As Exception
            D99C0008.Msg("Lỗi UseEnterAsTab: " & ex.Message)
        End Try
    End Sub

    Private Sub UseEnterAsTab(ByVal SplitCon As SplitContainer, Optional ByVal bForward As Boolean = True)
        Try
            If (SplitCon.ActiveControl.GetType.Name = "GridEditor") Or (SplitCon.ActiveControl.GetType.Name = "C1TrueDBGrid") Then 'Khong phai luoi
                Exit Sub
            End If
            If SplitCon.ActiveControl.GetType.BaseType.Name = "UserControl" Then
                Dim uc As UserControl = CType(SplitCon.ActiveControl, UserControl)
                UseEnterAsTab(uc, bForward)
            Else
                SplitCon.SelectNextControl(SplitCon.ActiveControl, bForward, True, True, False)
            End If
        Catch ex As Exception
            D99C0008.Msg("Lỗi UseEnterAsTab: " & ex.Message)
        End Try
    End Sub

    Private Sub UseEnterAsTab(ByVal uc As UserControl, Optional ByVal bForward As Boolean = True)
        Try
            Select Case uc.ActiveControl.GetType.Name
                Case "GridEditor", "C1TrueDBGrid" ' Không làm
                Case "SplitContainer"
                    Dim SplitCon As SplitContainer = CType(uc.ActiveControl, SplitContainer)
                    UseEnterAsTab(SplitCon, bForward)
                Case Else
                    uc.SelectNextControl(uc.ActiveControl, bForward, True, True, False)
            End Select
        Catch ex As Exception
            D99C0008.Msg("Lỗi UseEnterAsTab: " & ex.Message)
        End Try
    End Sub

    Private Sub optPrintTypePeriod_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPrintTypePeriod.CheckedChanged
        If optPrintTypePeriod.Checked Then
            tdbcPeriodFrom.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
            tdbcPeriodTo.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        Else
            tdbcPeriodFrom.EditorBackColor = Color.White
            tdbcPeriodTo.EditorBackColor = Color.White
        End If
    End Sub

    Private Sub optPrintTypeDate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPrintTypeDate.CheckedChanged
        If optPrintTypeDate.Checked Then
            c1dateDateFrom.BackColor = COLOR_BACKCOLOROBLIGATORY
            c1dateDateTo.BackColor = COLOR_BACKCOLOROBLIGATORY
        Else
            c1dateDateFrom.BackColor = Color.White
            c1dateDateTo.BackColor = Color.White
        End If
    End Sub

    'Xuất Excel
    Private Sub tsbExportToExcel_Click(sender As Object, e As EventArgs) Handles tsbExportToExcel.Click, tsmExportToExcel.Click, mnsExportToExcel.Click
        '25/5/2018, id 108159-HỖ TRỢ XUẤT EXCEL MÀN HÌNH TẬP HỢP CHI PHÍ XDCB
        CreateTableCaption() 'Tạo table caption để xuất cột trên lưới
        CallShowD99F2222(Me, dtCaptionCols, dtGrid, gsGroupColumns)
    End Sub

    Private Sub CreateTableCaption()
        Dim Arr As New ArrayList
        For i As Integer = 0 To tdbg.Splits.Count - 1
            If tdbg.Splits(i).SplitSize = 0 Then Continue For
            AddColVisible(tdbg, i, Arr, , False, False, gbUnicode)
        Next
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
    End Sub

    Private Sub mnsPrint_Click(sender As Object, e As EventArgs) Handles mnsPrint.Click
        Print(Me, "D02F1003", "02")
    End Sub

    Private Sub Print(ByVal form As Form, Optional ByVal sReportTypeID As String = "", Optional ByVal ModuleID As String = "")
        ''If Not AllowNewD99C2003(report, Me) Then Exit Sub ''Mở rem khi form trong DLL
        Dim sReportName As String = "D02R1003"
        Dim sSubReportName As String = ""
        Dim sReportPath As String = ""
        Dim sReportTitle As String = "" ''Thêm biến
        Dim sCustomReport As String = ""
        Dim file As String = D99D0541.GetReportPathNew(ModuleID, sReportTypeID, sReportName, sCustomReport, sReportPath, sReportTitle)
        If sReportName = "" Then Exit Sub
        form.Cursor = Cursors.WaitCursor
        Dim sSQL As String = SQLStoreD02P2050("", "", "D02R1003", "D02F1003", 1)

        Select Case file.ToLower
            Case "rpt"
                printReport(sSubReportName, sReportPath, sReportTitle, sSQL)
            Case "xls", "xlsx"
                Dim sPathFile As String = D99D0541.GetObjectFile(sReportTypeID, sReportName, file, sReportPath)
                If sPathFile = "" Then Exit Select
                SetVariable()
                Dim excel As New Lemon3.Reports.L3XtraReportExcel()
                excel.ExcelType = Lemon3.Reports.PrintExcelType.Normal
                excel.MyExcel(sSQL, sPathFile, file)
                excel.OpenFileExcel(sPathFile)
        End Select
        form.Cursor = Cursors.Default
    End Sub

    Private Sub printReport(ByVal sSubReportName As String, ByVal sReportPath As String, ByVal sReportCaption As String, ByVal sSQL As String)
        Dim report As New D99C1003 ''Chỉ Sử dụng khi Form trong Exe
        Dim conn As New SqlConnection(gsConnectionString)
        Dim sSQLSub As String = ""
        UnicodeSubReport(sSubReportName, sSQLSub, gsDivisionID, gbUnicode)
        With report
            .OpenConnection(conn)
            .AddSub(sSQLSub, sSubReportName & ".rpt")
            .AddMain(sSQL)
            .PrintReport(sReportPath, sReportCaption)
        End With
    End Sub

End Class