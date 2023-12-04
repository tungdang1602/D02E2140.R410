Imports System
Public Class D02F3003

#Region "Const of tdbg"
    Private Const COL_TransactionID As String = "TransactionID"     ' TransactionID
    Private Const COL_BatchID As String = "BatchID"                 ' BatchID
    Private Const COL_ModuleID As String = "ModuleID"               ' ModuleID
    Private Const COL_DivisionID As String = "DivisionID"           ' DivisionID
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

#Region "Const of tdbg2"
    Private Const COL2_TransactionID As String = "TransactionID"     ' TransactionID
    Private Const COL2_BatchID As String = "BatchID"                 ' BatchID
    Private Const COL2_ As String = ""                               ' STT
    Private Const COL2_Description As String = "Description"         ' Diễn giải
    Private Const COL2_DebitAccountID As String = "DebitAccountID"   ' Tài khoản nợ
    Private Const COL2_CreditAccountID As String = "CreditAccountID" ' Tài khoản có
    Private Const COL2_ConvertedAmount As String = "ConvertedAmount" ' Số tiền
    Private Const COL2_ObjectTypeID As String = "ObjectTypeID"       ' Loại đối tượng
    Private Const COL2_ObjectID As String = "ObjectID"               ' Đối tượng
    Private Const COL2_CipID As String = "CipID"                     ' Mã chi phí XDCB
    Private Const COL2_Ana01ID As String = "Ana01ID"                 ' Ana01ID
    Private Const COL2_Ana02ID As String = "Ana02ID"                 ' Ana02ID
    Private Const COL2_Ana03ID As String = "Ana03ID"                 ' Ana03ID
    Private Const COL2_Ana04ID As String = "Ana04ID"                 ' Ana04ID
    Private Const COL2_Ana05ID As String = "Ana05ID"                 ' Ana05ID
    Private Const COL2_Ana06ID As String = "Ana06ID"                 ' Ana06ID
    Private Const COL2_Ana07ID As String = "Ana07ID"                 ' Ana07ID
    Private Const COL2_Ana08ID As String = "Ana08ID"                 ' Ana08ID
    Private Const COL2_Ana09ID As String = "Ana09ID"                 ' Ana09ID
    Private Const COL2_Ana10ID As String = "Ana10ID"                 ' Ana10ID
#End Region

    Private _batchID As String = ""
    Public WriteOnly Property batchID() As String
        Set(ByVal Value As String)
            _batchID = Value
        End Set
    End Property

    Private _keyID As String = ""
    Public ReadOnly Property KeyID() As String
        Get
            Return _keyID
        End Get
    End Property

    Private _bSavedOK As Boolean = False
    Public ReadOnly Property bSavedOK() As Boolean
        Get
            Return _bSavedOK
        End Get
    End Property

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Cap_nhat_phieu_tach_chi_phi_-_D02F3003") & UnicodeCaption(gbUnicode) 'CËp nhËt phiÕu tÀch chi phÛ - D02F3003
        '================================================================ 
        lblSplitMethodNo.Text = rl3("Phuong_phap_tach") 'Phương pháp tách
        lblVoucherTypesID.Text = rl3("Loai_phieu") 'Loại phiếu
        lblVoucherNo.Text = rl3("So_phieu") 'Số phiếu
        lblteVoucherDate.Text = rl3("Ngay_phieu") 'Ngày phiếu
        lblDescription.Text = rl3("Dien_giai") 'Diễn giải
        '================================================================ 
        btnChooseVoucher.Text = "&" & rl3("Chon_phieu") 'Chọn phiếu
        btnSplit.Text = "&" & rl3("Tach") 'Tách
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        btnNext.Text = rl3("Nhap__tiep") 'Nhập &tiếp
        btnSave.Text = rl3("_Luu") '&Lưu
        '================================================================ 
        chkSplitCipNo.Text = rl3("Tach_chi_phi_xay_dung_co_ban") 'Tách chi phí xây dựng cơ bản
        chkSplitCipNoFromVoucher.Text = rl3("Tach_chi_phi_XDCB_tu_phieu_nhap_so_du") 'Tách chi phí XDCB từ phiếu nhập số dư
        chkPosted.Text = rl3("Ket_chuyen_vao_chi_phi") 'Kết chuyển vào chi phí
        chkDisabled.Text = rl3("Khong_su_dung") 'Không sử dụng
        chkInherit.Text = rl3("Ke_thua_TK_no_TK_co_So_tien_Ma_loai_DT_Ma_DT") 'Kế thừa TK nợ, TK có, Số tiền, Mã loại ĐT, Mã ĐT
        '================================================================ 
        grpInfoVoucher.Text = rl3("Thong_tin_phieu_duoc_tach") 'Thông tin phiếu được tách
        '================================================================ 
        tdbcSplitMethodNo.Columns("SplitMethodNo").Caption = rl3("Ma") 'Mã
        tdbcSplitMethodNo.Columns("SplitMethodName").Caption = rl3("Ten") 'Tên
        tdbcVoucherTypeID.Columns("VoucherTypeID").Caption = rl3("Ma") 'Mã
        tdbcVoucherTypeID.Columns("VoucherTypeName").Caption = rl3("Ten") 'Tên
        '================================================================ 
        tdbdAna10ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna10ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna09ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna09ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna08ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna08ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna07ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna07ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna06ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna06ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna05ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna05ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna04ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna04ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna03ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna03ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna02ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna02ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdAna01ID.Columns("AnaID").Caption = rl3("Ma") 'Mã
        tdbdAna01ID.Columns("AnaName").Caption = rl3("Ten") 'Tên
        tdbdCipID.Columns("CipNo").Caption = rl3("Ma") 'Mã
        tdbdCipID.Columns("CipName").Caption = rl3("Ten") 'Tên
        tdbdObjectID.Columns("ObjectID").Caption = rl3("Ma") 'Mã
        tdbdObjectID.Columns("ObjectName").Caption = rl3("Ten") 'Tên
        tdbdObjectTypeID.Columns("ObjectTypeID").Caption = rl3("Ma") 'Mã
        tdbdObjectTypeID.Columns("ObjectTypeName").Caption = rl3("Ten") 'Tên
        tdbdCreditAccountID.Columns("AccountID").Caption = rl3("Ma") 'Mã
        tdbdCreditAccountID.Columns("AccountName").Caption = rl3("Ten") 'Tên
        tdbdDebitAccountID.Columns("AccountID").Caption = rl3("Ma") 'Mã
        tdbdDebitAccountID.Columns("AccountName").Caption = rl3("Ten") 'Tên
        '================================================================ 
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

        tdbg2.Columns("").Caption = rl3("STT") 'STT
        tdbg2.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbg2.Columns("DebitAccountID").Caption = rl3("Tai_khoan_no") 'Tài khoản nợ
        tdbg2.Columns("CreditAccountID").Caption = rl3("Tai_khoan_co") 'Tài khoản có
        tdbg2.Columns("ConvertedAmount").Caption = rl3("So_tien") 'Số tiền
        tdbg2.Columns("ObjectTypeID").Caption = rl3("Loai_doi_tuong") 'Loại đối tượng
        tdbg2.Columns("ObjectID").Caption = rl3("Doi_tuong") 'Đối tượng
        tdbg2.Columns("CipID").Caption = rl3("Ma_chi_phi_XDCB") 'Mã chi phí XDCB
    End Sub

    Private Sub D02F3003_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If sListBatchIDTemp <> "" Then ExecuteSQLNoTransaction("DELETE D02T0017 WHERE BatchIDTemp IN (" & sListBatchIDTemp & ")")
    End Sub

    Private Sub D02F3003_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        ElseIf e.KeyCode = Keys.F11 Then
            HotKeyF11(Me, tdbg)
        End If
    End Sub

    Private Sub D02F3003_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        iPer_F5558 = ReturnPermission("D02F5558")
        Dim bUseAna As Boolean = LoadTDBGridAnalysisCaption(D02, tdbg, IndexOfColumn(tdbg, COL_Ana01ID), 1, True, gbUnicode)
        If bUseAna Then
            For i As Integer = 0 To 9
                Dim iCol As Integer = IndexOfColumn(tdbg, COL_Ana01ID) + i 'Cột lưới 1
                Dim iCol2 As Integer = IndexOfColumn(tdbg2, COL2_Ana01ID) + i 'Cột lưới 2
                tdbg2.Columns(iCol2).Caption = tdbg.Columns(iCol).Caption
                tdbg2.Splits(1).DisplayColumns(iCol2).HeadingStyle.Font = FontUnicode(gbUnicode)
                tdbg2.Splits(1).DisplayColumns(iCol2).Visible = tdbg.Splits(1).DisplayColumns(iCol).Visible
                tdbg2.Columns(iCol2).Tag = tdbg2.Splits(1).DisplayColumns(iCol2).Visible 'để load dropdown
            Next
            LoadTDBDropDownAna(tdbdAna01ID, tdbdAna02ID, tdbdAna03ID, tdbdAna04ID, tdbdAna05ID, tdbdAna06ID, tdbdAna07ID, tdbdAna08ID, tdbdAna09ID, tdbdAna10ID, tdbg2, IndexOfColumn(tdbg2, COL2_Ana01ID), gbUnicode)
        Else
            tdbg.RemoveHorizontalSplit(1)
            tdbg2.RemoveHorizontalSplit(1)
        End If
        ResetSplitDividerSize(tdbg)
        ResetSplitDividerSize(tdbg2)
        ResetColorGrid(tdbg, 0, tdbg.Splits.Count - 1)
        ResetFooterGrid(tdbg2, 0, tdbg2.Splits.Count - 1)
        LoadTDBDropDown()
        InputDateInTrueDBGrid(tdbg, COL_RefDate, COL_VoucherDate)
        tdbg_NumberFormat()
        tdbg2_NumberFormat()
        Dim arr() As FormatColumn = Nothing
        AddNumberColumns(arr, SqlDbType.Money, COL2_ConvertedAmount, DxxFormat.D90_ConvertedDecimals)
        InputNumber(tdbg2, arr)
        tdbg2_LockedColumns()
        Loadlanguage()
        LoadTDBGrid()
        LoadTDBGrid2()
        InputbyUnicode(Me, gbUnicode)
        SetBackColorObligatory()
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
            _FormState = value
            LoadTDBCombo()
            Select Case _FormState
                Case EnumFormState.FormAdd
                    LoadVoucherTypeID(tdbcVoucherTypeID, D02, , gbUnicode)
                    LoadAddNew()
                Case EnumFormState.FormEdit
                    LoadEdit()
                Case EnumFormState.FormView
                    btnSave.Enabled = False
                    LoadEdit()
            End Select
        End Set
    End Property

    Private Sub LoadAddNew()
        btnSave.Enabled = False
        btnNext.Enabled = False
        btnSplit.Enabled = False
        btnChooseVoucher.Enabled = True

        c1dateVoucherDate.Value = Now.Date
    End Sub

    Private Sub ClearAllValue()

        'Xóa lưới
        If dtGrid IsNot Nothing Then
            dtGrid.Clear()
            LoadTDBGrid(False)
        End If
        If dtGrid2 IsNot Nothing Then
            dtGrid2.Clear()
            LoadTDBGrid2(False)
        End If

        '************
        'Master
        tdbcSplitMethodNo.SelectedValue = ""
        txtSplitMethodNoName.Text = ""
        chkDisabled.Checked = False
        tdbcVoucherTypeID.SelectedValue = ""
        txtVoucherNo.Text = ""
        txtDescription.Text = ""
        LoadAddNew()
    End Sub

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        'Load tdbcSplitMethodNo
        sSQL = "Select SplitMethodNo, SplitMethodName" & UnicodeJoin(gbUnicode) & " as SplitMethodName From D02T0014 WITH(NOLOCK) WHERE Disabled= 0 ORDER BY SplitMethodNo "
        LoadDataSource(tdbcSplitMethodNo, sSQL, gbUnicode)
    End Sub

    Private Sub LoadTDBDropDown()
        Dim sSQL As String = ""
        'Load tdbdDebitAccountID
        Dim dtAccountID As DataTable = ReturnTableAccountID("AccountStatus = 0", gbUnicode)
        LoadDataSource(tdbdDebitAccountID, dtAccountID, gbUnicode)
        LoadDataSource(tdbdCreditAccountID, dtAccountID.DefaultView.ToTable, gbUnicode)

        'Load tdbdObjectTypeID
        LoadObjectTypeID(tdbdObjectTypeID, gbUnicode)
        'Load tdbdCipID
        sSQL = "--Do nguon dropdown ma XDCB" & vbCrLf
        sSQL = "Select CipID, CipNo, CipName" & UnicodeJoin(gbUnicode) & " as CipName From D02T0100 WITH(NOLOCK) "
        sSQL &= "WHERE Disabled = 0 AND Status < 2 "
        sSQL &= "AND DivisionID = " & SQLString(gsDivisionID) ' uppdate 31/5/2013 id 56796
        LoadDataSource(tdbdCipID, sSQL, gbUnicode)
    End Sub

    Private Sub LoadtdbdObjectID(ByVal sObjectTypeID As String)
        Dim sSQL As String = ""
        sSQL = "SELECT ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " as ObjectName, VATNo "
        sSQL &= "FROM Object WITH(NOLOCK) WHERE ObjectTypeID = '" & sObjectTypeID & "' And Disabled=0 "
        sSQL &= " ORDER BY ObjectID"
        LoadDataSource(tdbdObjectID, sSQL, gbUnicode)
    End Sub

    Dim sEditVoucherTypeID As String = ""
    Private Sub LoadEdit()
        btnNext.Visible = False
        btnSave.Left = btnNext.Left
        btnSplit.Enabled = False
        ReadOnlyControl(tdbcVoucherTypeID, txtVoucherNo)
        chkPosted.Enabled = False
        chkSplitCipNo.Enabled = False
        chkSplitCipNoFromVoucher.Enabled = False

        'Load Master
        Dim strSQL As String = ""
        strSQL = "SELECT T1.VoucherTypeID, T1.VoucherNo,T1.VoucherDate, T1.Notes" & UnicodeJoin(gbUnicode) & " as Notes, T1.Disabled, T1.SplitCipNo, T1.SplitMethodNo" & vbCrLf
        strSQL &= "FROM D02T0016 T1 WITH(NOLOCK)  " & vbCrLf
        strSQL &= "INNER JOIN D02T0014 T2 WITH(NOLOCK) ON T1.SplitMethodNo = T2.SplitMethodNo " & vbCrLf
        strSQL &= "WHERE BatchID = " & SQLString(_batchID)
        Dim dtTemp As DataTable = ReturnDataTable(strSQL)
        If dtTemp.Rows.Count = 0 Then Exit Sub
        With dtTemp.Rows(0)
            sEditVoucherTypeID = .Item("VoucherTypeID").ToString
            LoadVoucherTypeID(tdbcVoucherTypeID, D02, sEditVoucherTypeID, gbUnicode)
            tdbcVoucherTypeID.SelectedValue = sEditVoucherTypeID
            tdbcSplitMethodNo.SelectedValue = .Item("SplitMethodNo").ToString
            txtVoucherNo.Text = .Item("VoucherNo").ToString
            c1dateVoucherDate.Value = .Item("VoucherDate")
            txtDescription.Text = .Item("Notes").ToString
            chkDisabled.Checked = L3Bool(.Item("Disabled"))
            Select Case L3Int(.Item("SplitCipNo"))
                Case 1
                    chkSplitCipNo.Checked = True
                    '       chkSplitCipNo_CheckedChanged(Nothing, Nothing)
                Case 2
                    chkPosted.Checked = True
                    '        chkPosted_CheckedChanged(Nothing, Nothing)
                Case 3
                    chkSplitCipNoFromVoucher.Checked = True
            End Select
            '            chkSplitCipNo_CheckedChanged(Nothing, Nothing)
            '            chkPosted_CheckedChanged(Nothing, Nothing)
            '            chkSplitCipNoFromVoucher_CheckedChanged(Nothing, Nothing)
            'Chỉ sáng khi cả 2 checkbox =False
            btnSplit.Enabled = Not (chkPosted.Checked OrElse chkSplitCipNo.Checked OrElse chkSplitCipNoFromVoucher.Checked)
            btnChooseVoucher.Enabled = btnSplit.Enabled
            ReadOnlyControl(chkPosted.Checked OrElse chkSplitCipNo.Checked OrElse chkSplitCipNoFromVoucher.Checked, tdbcSplitMethodNo)
            '**************
        End With
    End Sub

#Region "Events tdbcVoucherTypeID with txtVoucherNo"
    Private Sub tdbcVoucherTypeID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.LostFocus
        If tdbcVoucherTypeID.FindStringExact(tdbcVoucherTypeID.Text) = -1 Then
            tdbcVoucherTypeID.Text = ""
            txtVoucherNo.Text = ""
        End If
    End Sub

    Private Sub tdbcVoucherTypeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.SelectedValueChanged
        If tdbcVoucherTypeID.SelectedValue Is Nothing OrElse tdbcVoucherTypeID.Text = "" Then
            txtVoucherNo.Text = ""
            ReadOnlyControl(txtVoucherNo)
            Exit Sub
        End If
        If _FormState = EnumFormState.FormAdd Then
            If tdbcVoucherTypeID.Columns("Auto").Text = "1" Then 'Sinh tu dong
                txtVoucherNo.Text = CreateIGEVoucherNo(tdbcVoucherTypeID, False)
                ReadOnlyControl(txtVoucherNo)
                '   c1dateVoucherDate.Focus()
            Else
                txtVoucherNo.Text = ""
                UnReadOnlyControl(txtVoucherNo)
                '    txtVoucherNo.Focus()
            End If
        End If
    End Sub
#End Region

#Region "Events tdbcSplitMethodNo with txtSplitMethodNoName"

    Private Sub tdbcSplitMethodNo_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcSplitMethodNo.SelectedValueChanged
        If tdbcSplitMethodNo.SelectedValue Is Nothing Then
            txtSplitMethodNoName.Text = ""
        Else
            txtSplitMethodNoName.Text = tdbcSplitMethodNo.Columns(1).Value.ToString
            If dtGrid2 IsNot Nothing Then
                dtGrid2.Clear()
                LoadTDBGrid2(False)
            End If

        End If
    End Sub

    Private Sub tdbcSplitMethodNo_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcSplitMethodNo.LostFocus
        If tdbcSplitMethodNo.FindStringExact(tdbcSplitMethodNo.Text) = -1 Then
            tdbcSplitMethodNo.Text = ""
        End If
    End Sub

#End Region

    '************************
    Dim sOldVoucherNo As String = "" 'Lưu lại số phiếu cũ
    Dim bEditVoucherNo As Boolean = False '= True: có nhấn F2; = False: không 
    Dim bFirstF2 As Boolean = False 'Nhấn F2 lần đầu tiên 
    Dim iPer_F5558 As Integer = 0 'Phân quyền cho Sửa số phiếu
    '************************
    Private Sub txtVoucherNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVoucherNo.KeyDown
        If e.KeyCode = Keys.F2 Then
            'Loại phiếu hay Số phiếu = "" thì thoát
            If tdbcVoucherTypeID.Text = "" Or txtVoucherNo.Text = "" Then Exit Sub

            'Update 21/09/2010: Trường hợp Thêm mới phiếu và đã lưu Thành công thì không cho sửa Số phiếu
            If _FormState = EnumFormState.FormAdd And btnSave.Enabled = False Then Exit Sub
            'Kiểm tra quyền cho trường hợp Sửa
            If _FormState = EnumFormState.FormEdit And iPer_F5558 <= 2 Then Exit Sub

            'Cho sửa Số phiếu ở trạng thái Thêm mới hay Sửa
            If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormEdit Then
                'Trước khi gọi exe con thì nhớ lại Số phiếu cũ
                If bFirstF2 = False Then
                    sOldVoucherNo = txtVoucherNo.Text
                    bFirstF2 = True
                End If
                'Gọi exe con D91E0640
                'Dim frm As New D91F5558
                'With frm
                '    .FormName = "D91F5558"
                '    .FormPermission = "D02F5558" 'Màn hình phân quyền
                '    .ModuleID = D02 'Mã module hiện tại, VD: D22
                '    .TableName = "D02T0016" 'Tên bảng chứa số phiếu
                '    'Update 21/09/2010
                '    If _FormState = EnumFormState.FormAdd Then
                '        .VoucherID = "" 'Khóa sinh IGE là rỗng
                '    ElseIf _FormState = EnumFormState.FormEdit Then
                '        .VoucherID = _batchID   'Khóa sinh IGE
                '    End If
                '    .VoucherNo = txtVoucherNo.Text 'Số phiếu cần sửa
                '    .Mode = "0" ' Tùy theo Module, mặc định là 0
                '    .KeyID01 = ""
                '    .KeyID02 = ""
                '    .KeyID03 = ""
                '    .KeyID04 = ""
                '    .KeyID05 = ""
                '    .ShowDialog()
                '    Dim sVoucherNo As String
                '    sVoucherNo = .Output02
                '    .Dispose()
                '    If sVoucherNo <> "" Then
                '        txtVoucherNo.Text = sVoucherNo 'Giá trị trả về Số phiếu mới
                '        ReadOnlyControl(txtVoucherNo) 'Lock text Số phiếu
                '        bEditVoucherNo = True 'Đã nhấn F2
                '        _bSavedOK = True
                '    End If
                'End With

                Dim arrPro() As StructureProperties = Nothing
                SetProperties(arrPro, "FormIDPermission", "D02F5558")
                SetProperties(arrPro, "VoucherTypeID", ReturnValueC1Combo(tdbcVoucherTypeID))

                If _FormState = EnumFormState.FormAdd Then
                    SetProperties(arrPro, "VoucherID", "")
                ElseIf _FormState = EnumFormState.FormEdit Then
                    SetProperties(arrPro, "VoucherID", _batchID)
                End If
                SetProperties(arrPro, "Mode", 0)
                SetProperties(arrPro, "KeyID01", "")
                SetProperties(arrPro, "TableName", "D02T0016")
                SetProperties(arrPro, "ModuleID", D02)
                SetProperties(arrPro, "OldVoucherNo", txtVoucherNo.Text)
                SetProperties(arrPro, "KeyID02", "")
                SetProperties(arrPro, "KeyID03", "")
                SetProperties(arrPro, "KeyID04", "")
                SetProperties(arrPro, "KeyID05", "")
                Dim frm As Form = CallFormShowDialog("D91D0640", "D91F5558", arrPro)
                Dim sNew As String = GetProperties(frm, "NewVoucherNo").ToString
                If sNew <> "" Then
                    txtVoucherNo.Text = sNew 'Giá trị trả về Số phiếu mới
                    ReadOnlyControl(txtVoucherNo) 'Lock text Số phiếu
                    bEditVoucherNo = True 'Đã nhấn F2
                    _bSavedOK = True
                End If
            End If
        End If
    End Sub

#Region "Events of tdbg2"

    Private Sub HeadClick(ByVal iCol As Integer)
        Select Case iCol
            Case IndexOfColumn(tdbg2, COL2_CreditAccountID), IndexOfColumn(tdbg2, COL2_Ana01ID) To IndexOfColumn(tdbg2, COL2_Ana10ID)
                CopyColumns(tdbg2, iCol, tdbg2.Columns(iCol).Text, tdbg2.Row)
            Case IndexOfColumn(tdbg2, COL2_ObjectTypeID), IndexOfColumn(tdbg2, COL2_ObjectID)
                CopyColumnArr(tdbg2, iCol, New Integer() {IndexOfColumn(tdbg2, COL2_ObjectTypeID), IndexOfColumn(tdbg2, COL2_ObjectID)})
        End Select
    End Sub

    Private Sub tdbg2_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.HeadClick
        HeadClick(e.ColIndex)
    End Sub

    Private Sub tdbg2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg2.KeyDown
        If e.Control And e.KeyCode = Keys.S Then
            HeadClick(tdbg2.Col)
        End If
    End Sub

    Private Sub tdbg2_UnboundColumnFetch(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.UnboundColumnFetchEventArgs) Handles tdbg2.UnboundColumnFetch
        e.Value = (e.Row + 1).ToString
    End Sub

    Private Sub tdbg2_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.ComboSelect
        tdbg2.UpdateData()
    End Sub

    Dim bNotInList As Boolean = False

    Private Sub tdbg2_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg2.BeforeColUpdate
        '--- Kiểm tra giá trị hợp lệ
        Select Case e.Column.DataColumn.DataField
            Case COL2_CreditAccountID, COL2_ObjectTypeID, COL2_ObjectID, COL2_Ana01ID, COL2_Ana02ID, COL2_Ana03ID, COL2_Ana04ID, COL2_Ana05ID, COL2_Ana06ID, COL2_Ana07ID, COL2_Ana08ID, COL2_Ana09ID, COL2_Ana10ID
                If tdbg2.Columns(e.ColIndex).Value.ToString <> tdbg2.Columns(e.ColIndex).DropDown.Columns(0).Text Then
                    tdbg2.Columns(e.ColIndex).Text = ""
                End If
            Case COL2_CipID
                If tdbg2.Columns(e.ColIndex).Text <> tdbg2.Columns(e.ColIndex).DropDown.Columns(tdbg2.Columns(e.ColIndex).DropDown.DisplayMember).Text Then
                    tdbg2.Columns(e.ColIndex).Text = ""
                    bNotInList = True
                End If
        End Select
    End Sub

    Private Sub tdbg2_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbg2.RowColChange
  If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
        If tdbg2.RowCount = 0 Then Exit Sub
        '--- Đổ nguồn cho các Dropdown phụ thuộc
        Select Case tdbg2.Columns(tdbg2.Col).DataField
            Case COL2_ObjectID
                LoadtdbdObjectID(tdbg2(tdbg2.Row, COL_ObjectTypeID).ToString)
        End Select
    End Sub

    Private Sub tdbg2_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.AfterColUpdate
        '--- Gán giá trị cột sau khi tính toán
        Select Case e.Column.DataColumn.DataField
            Case COL2_ConvertedAmount
                FooterSumNew(tdbg2, e.ColIndex)
            Case COL2_CipID
                If tdbg2.Columns(e.ColIndex).Text = "" OrElse bNotInList Then
                    tdbg2.Columns(e.ColIndex).Text = ""
                    bNotInList = False
                    'Gắn rỗng các cột liên quan
                    Exit Select
                End If
        End Select
        bNotInList = False
    End Sub

#End Region

#Region "Events of tdbg"
    Dim sFilter As New System.Text.StringBuilder()
    Dim bRefreshFilter As Boolean = False 'Cờ bật set FilterText =""
    Private Sub tdbg_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.FilterChange
        Try
            'If (dtGrid Is Nothing) Then Exit Sub
            'If bRefreshFilter Then Exit Sub 'set FilterText ="" thì thoát
            ''Filter the data 
            'FilterChangeGrid(tdbg, sFilter)
            'ReLoadTDBGrid()
        Catch ex As Exception
            'Update 11/05/2011: Tạm thời có lỗi thì bỏ qua không hiện message
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte
        End Try
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        HotKeyCtrlVOnGrid(tdbg, e) 'Nhấn Ctrl + V trên lưới 'có trong D99X0000
    End Sub

    Private Sub tdbg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        Select Case tdbg.Columns(tdbg.Col).DataField
            Case COL_ConvertedAmount, COL_OriginalAmount, COL_ExchangeRate
                'e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub
#End Region

    Private Sub chkPosted_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPosted.CheckedChanged
        If chkPosted.Checked Then
            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_DebitAccountID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
            chkSplitCipNo.Checked = False
            chkSplitCipNoFromVoucher.Checked = False
            If _FormState = EnumFormState.FormAdd Then ClearAllValue()
        Else
            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_DebitAccountID).Style.ResetBackColor()
        End If
        
        LockbySplitCipNo()
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



    'Private Sub tdbg2_NumberFormat()
    '    tdbg2.Columns(COL2_).NumberFormat = D02Format.DefaultNumber0
    '    tdbg2.Columns(COL2_ConvertedAmount).NumberFormat = D02CustomFormat.D90_ConvertedDecimals
    'End Sub

    Private Sub tdbg2_NumberFormat()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbg2.Columns(COL2_ConvertedAmount).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        InputNumber(tdbg2, arr)
    End Sub



    Private Sub tdbg2_LockedColumns()
        tdbg2.Splits(SPLIT0).DisplayColumns(COL2_).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
    End Sub

    Private Sub LockbySplitCipNo()
        If Not chkSplitCipNo.Checked And Not chkSplitCipNoFromVoucher.Checked Then
            tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).AllowFocus = True
            tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).Locked = False
            tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).Style.ResetBackColor()

            tdbg2.Columns(COL2_DebitAccountID).DropDown = tdbdDebitAccountID
            tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).AutoDropDown = True
        Else
            If chkSplitCipNo.Checked Then
                tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).Locked = True
                tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
                tdbg2.Columns(COL2_DebitAccountID).DropDown = Nothing
                tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).AutoDropDown = False
                tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).AllowFocus = False
            End If
        End If
        tdbg2.Splits(0).DisplayColumns(COL2_CipID).Visible = chkSplitCipNo.Checked
        '        If chkSplitCipNo.Checked Then
        '            tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).Locked = True
        '            tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        '            tdbg2.Columns(COL2_DebitAccountID).DropDown = Nothing
        '            tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).AutoDropDown = False
        '            tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).AllowFocus = False
        '        Else
        '            tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).AllowFocus = True
        '            tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).Locked = False
        '            tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).Style.ResetBackColor()
        '
        '            tdbg2.Columns(COL2_DebitAccountID).DropDown = tdbdDebitAccountID
        '            tdbg2.Splits(0).DisplayColumns(COL2_DebitAccountID).AutoDropDown = True
        '        End If
        '        tdbg2.Splits(0).DisplayColumns(COL2_CipID).Visible = chkSplitCipNo.Checked
    End Sub

    Private Sub chkSplitCipNo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkSplitCipNo.CheckedChanged
        If Not chkSplitCipNo.Checked And Not chkSplitCipNoFromVoucher.Checked Then
            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectTypeID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
            LockbySplitCipNo()
            Exit Sub
        Else
            If chkSplitCipNo.Checked Then
                chkPosted.Checked = False
                chkSplitCipNoFromVoucher.Checked = False
                tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectTypeID).Style.ResetBackColor()
                tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectID).Style.ResetBackColor()
                If _FormState = EnumFormState.FormAdd Then ClearAllValue()
                LockbySplitCipNo()
            End If
        End If
        '        If chkSplitCipNo.Checked Then
        '            chkPosted.Checked = False
        '            chkSplitCipNoFromVoucher.Checked = False
        '            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectTypeID).Style.ResetBackColor()
        '            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectID).Style.ResetBackColor()
        '            If _FormState = EnumFormState.FormAdd Then ClearAllValue()
        '        Else
        '            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectTypeID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
        '            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
        '        End If
        '        LockbySplitCipNo()
    End Sub

    ' update 31/5/2013 id 56796
    Private Sub chkSplitCipNoFromVoucher_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSplitCipNoFromVoucher.CheckedChanged
        If Not chkSplitCipNo.Checked And Not chkSplitCipNoFromVoucher.Checked Then
            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectTypeID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
            LockbySplitCipNoFromVoucher()
            Exit Sub
        Else
            If chkSplitCipNoFromVoucher.Checked Then
                chkSplitCipNo.Checked = False
                chkPosted.Checked = False
                tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectTypeID).Style.ResetBackColor()
                tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectID).Style.ResetBackColor()
                If _FormState = EnumFormState.FormAdd Then ClearAllValue()
                LockbySplitCipNoFromVoucher()
            End If
        End If
        '        If chkSplitCipNoFromVoucher.Checked Then
        '            chkSplitCipNo.Checked = False
        '            chkPosted.Checked = False
        '            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectTypeID).Style.ResetBackColor()
        '            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectID).Style.ResetBackColor()
        '            If _FormState = EnumFormState.FormAdd Then ClearAllValue()
        '        Else
        '            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectTypeID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
        '            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
        '        End If
        '        LockbySplitCipNoFromVoucher()
    End Sub

    ' update 31/5/2013 id 56796
    Private Sub LockbySplitCipNoFromVoucher()
        If Not chkSplitCipNo.Checked And Not chkSplitCipNoFromVoucher.Checked Then
            tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).AllowFocus = True
            tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).Locked = False
            tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).Style.ResetBackColor()

            tdbg2.Columns(COL2_CreditAccountID).DropDown = tdbdDebitAccountID
            tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).AutoDropDown = True
        Else
            If chkSplitCipNoFromVoucher.Checked Then
                tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).Locked = True
                '  tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).Style.ResetBackColor()
                tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
                tdbg2.Columns(COL2_CreditAccountID).DropDown = Nothing
                tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).AutoDropDown = False
                tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).AllowFocus = False
            End If
        End If
        tdbg2.Splits(0).DisplayColumns(COL2_CipID).Visible = chkSplitCipNoFromVoucher.Checked
            '        If chkSplitCipNoFromVoucher.Checked Then
            '            tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).Locked = True
            '            tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
            '            tdbg2.Columns(COL2_CreditAccountID).DropDown = Nothing
            '            tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).AutoDropDown = False
            '            tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).AllowFocus = False
            '        Else
            '            tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).AllowFocus = True
            '            tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).Locked = False
            '            tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).Style.ResetBackColor()
            '
            '            tdbg2.Columns(COL2_CreditAccountID).DropDown = tdbdDebitAccountID
            '            tdbg2.Splits(0).DisplayColumns(COL2_CreditAccountID).AutoDropDown = True
            '        End If
            '        tdbg2.Splits(0).DisplayColumns(COL2_CipID).Visible = chkSplitCipNoFromVoucher.Checked
    End Sub

    Private Sub btnChooseVoucher_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChooseVoucher.Click
        Dim f As New D02F3005
        Dim iSplitCipNo As Integer = 0
        If chkSplitCipNo.Checked Then
            iSplitCipNo = 2 ' Truyền khác với mode khi luu - Đã hỏi PSD (Bao Tran)
        ElseIf chkPosted.Checked Then
            iSplitCipNo = 1
        ElseIf chkSplitCipNoFromVoucher.Checked Then
            iSplitCipNo = 3
        End If
        f.Mode = iSplitCipNo
        ' f.Mode = CInt(IIf(chkSplitCipNo.Checked, 2, IIf(chkPosted.Checked, 1, 0)))
        f.ShowDialog()
        If f.bSavedOK Then
            If dtGrid2 IsNot Nothing Then
                dtGrid2.Clear()
                LoadTDBGrid2(False)
            End If
            dtGrid = f.dtChose
            LoadTDBGrid(False)
        End If
        f.Dispose()
    End Sub

    Dim dtGrid As DataTable
    Private Sub LoadTDBGrid(Optional ByVal bLoadSQL As Boolean = True)
        If bLoadSQL Then
            Dim strSQL As String = "SELECT Description" & UnicodeJoin(gbUnicode) & " as Description, SignOriginalAmount as OriginalAmount , SignConvertedAmount as ConvertedAmount, CipID AS OriginalCipID, * FROM D02V1000 "
            strSQL &= "WHERE SplitBatchID = " & SQLString(_batchID) & " And DivisionID=" & SQLString(gsDivisionID)
            If _FormState = EnumFormState.FormAdd Then strSQL &= " And 1=0"
            dtGrid = ReturnDataTable(strSQL)
        End If
        LoadDataSource(tdbg, dtGrid, gbUnicode)
        FooterTotalGrid(tdbg, COL_VoucherNo)
        FooterSumNew(tdbg, COL_ConvertedAmount, COL_OriginalAmount)
        btnSplit.Enabled = tdbg.RowCount > 0
    End Sub

    Dim dtGrid2 As DataTable 'Lưới dưới

    Private Sub LoadTDBGrid2(Optional ByVal bRunSQL As Boolean = True)
        If bRunSQL Then
            Dim strSQL As String = ""
            strSQL = "SELECT T12.CipID, T12.TransactionID, T12.BatchID, T12.Description" & UnicodeJoin(gbUnicode) & " as Description, T12.DebitAccountID" & _
                       ", T12.CreditAccountID,T12.ConvertedAmount, T12.ObjectTypeID, T12.ObjectID" & _
                       ", T12.Ana01ID, T12.Ana02ID, T12.Ana03ID, T12.Ana04ID, T12.Ana05ID, T12.Ana06ID, T12.Ana07ID, T12.Ana08ID, T12.Ana09ID, T12.Ana10ID"
            strSQL &= vbCrLf
            strSQL &= " FROM D02T0012 T12 WITH(NOLOCK)" & vbCrLf
            strSQL &= "LEFT JOIN   D02T0100 T10 WITH(NOLOCK) ON T10.CipID = T12.CipID" & vbCrLf
            strSQL &= "WHERE BatchID = '" & _batchID & "' And T12.DivisionID=" & SQLString(gsDivisionID)
            dtGrid2 = ReturnDataTable(strSQL)
        End If
        LoadDataSource(tdbg2, dtGrid2, gbUnicode)
        FooterTotalGrid(tdbg2, COL2_Description)
        FooterSumNew(tdbg2, COL2_ConvertedAmount)
    End Sub


    Private Function AllowSplit() As Boolean
        If tdbcSplitMethodNo.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Phuong_phap_tach"))
            tdbcSplitMethodNo.Focus()
            Return False
        End If
        Return True
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P4003
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 16/11/2011 11:22:41
    '# Modified User: 
    '# Modified Date: 
    '# Description: Tách
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P4003(ByVal sCAmount As Object, Optional ByVal sBatchID As Object = "", _
                                                                Optional ByVal sTransactionID As Object = "") As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P4003 "
        sSQL &= SQLString(ReturnValueC1Combo(tdbcSplitMethodNo)) & COMMA 'SplitMethodNo, varchar[20], NOT NULL
        sSQL &= SQLMoney(sCAmount, DxxFormat.D90_ConvertedDecimals) & COMMA 'GeneralAmount, money, NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLNumber(chkInherit.Checked) & COMMA 'Mode, smallint, NOT NULL
        sSQL &= SQLString(sBatchID) & COMMA 'BatchID, varchar[20], NOT NULL
        sSQL &= SQLString(sTransactionID) & COMMA 'TransactionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA
        sSQL &= SQLString(gsDivisionID)
        Return sSQL
    End Function

    Dim sListBatchIDTemp As String = ""

    Private Sub btnSplit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSplit.Click
        'Xóa dữ liệu rác khi nhấn Tách nhiều lần
        ExecuteSQLNoTransaction("DELETE D02T0017 WHERE BatchIDTemp IN (" & IIf(sListBatchIDTemp = "", "''", sListBatchIDTemp).ToString & ")")
        sListBatchIDTemp = ""
        '***************
        If Not AllowSplit() Then Exit Sub

        Dim dtTemp As DataTable
        If chkInherit.Checked Then
            For i As Integer = 0 To tdbg.RowCount - 1
                dtTemp = ReturnDataTable(SQLStoreD02P4003(tdbg(i, COL_ConvertedAmount), tdbg(i, COL_BatchID), tdbg(i, COL_TransactionID)))
                If dtTemp.Rows.Count = 0 Then Continue For
                If sListBatchIDTemp <> "" Then sListBatchIDTemp &= ","
                sListBatchIDTemp &= SQLString(dtTemp.Rows(0).Item("NewKey"))
            Next
        Else
            dtTemp = ReturnDataTable(SQLStoreD02P4003(tdbg.Columns(COL_ConvertedAmount).FooterText))
            If dtTemp.Rows.Count > 0 Then sListBatchIDTemp = SQLString(dtTemp.Rows(0).Item("NewKey"))
        End If
        Dim strSQL As String = ""
        'Dieu chinh
        Dim iRow As Integer = 0 'Dòng của lưới trên
        strSQL = "SELECT SplitNo,Description" & UnicodeJoin(gbUnicode) & " as Description, Sum(ConvertedAmount) As ConvertedAmount, IsNull(DebitAccountID,'') As DebitAccountID, " & vbCrLf
        strSQL &= " IsNull(CreditAccountID,'') As CreditAccountID, IsNull(ObjectTypeID,'') As ObjectTypeID, IsNull(ObjectID,'') As ObjectID" & IIf(chkSplitCipNo.Checked OrElse chkSplitCipNoFromVoucher.Checked, ",CipIDTemp as CipID", "").ToString & vbCrLf
        strSQL &= " From D02T0017 WITH(NOLOCK) Where BatchIDTemp In " & "(" & sListBatchIDTemp & ")" & vbCrLf
        strSQL &= " Group by SplitNo,Description" & UnicodeJoin(gbUnicode) & ", DebitAccountID, CreditAccountID, ObjectTypeID, ObjectID" & IIf(chkSplitCipNo.Checked OrElse chkSplitCipNoFromVoucher.Checked, ",CipIDTemp", "").ToString
        Dim dt As DataTable = ReturnDataTable(strSQL)
        If dt.Rows.Count > 0 Then
            For i As Integer = 0 To dt.Rows.Count - 1
                If IsDBNull(dt.Rows(i).Item(COL2_DebitAccountID)) OrElse dt.Rows(i).Item(COL2_DebitAccountID).ToString = "" Then dt.Rows(i).Item(COL2_DebitAccountID) = tdbg(iRow, COL_DebitAccountID)
                If chkInherit.Checked Then
                    If IsDBNull(dt.Rows(i).Item(COL2_CreditAccountID)) OrElse dt.Rows(i).Item(COL2_CreditAccountID).ToString = "" Then dt.Rows(i).Item(COL2_CreditAccountID) = tdbg(iRow, COL_CreditAccountID)
                    If IsDBNull(dt.Rows(i).Item(COL2_ObjectTypeID)) OrElse dt.Rows(i).Item(COL2_ObjectTypeID).ToString = "" Then dt.Rows(i).Item(COL2_ObjectTypeID) = tdbg(iRow, COL_ObjectTypeID)
                    If IsDBNull(dt.Rows(i).Item(COL2_ObjectID)) OrElse dt.Rows(i).Item(COL2_ObjectID).ToString = "" Then dt.Rows(i).Item(COL2_ObjectID) = tdbg(iRow, COL_ObjectID)
                Else
                    dt.Rows(i).Item(COL2_CreditAccountID) = ""
                    dt.Rows(i).Item(COL2_ObjectTypeID) = ""
                    dt.Rows(i).Item(COL2_ObjectID) = ""
                End If
                If iRow < tdbg.RowCount - 1 Then iRow += 1
            Next
            dtGrid2.Clear() 'Xóa dữ liệu cũ
            dtGrid2.Merge(dt)
            FooterTotalGrid(tdbg2, COL2_Description)
            FooterSumNew(tdbg2, COL2_ConvertedAmount)
        End If

        'Tính lại chênh lệch Số tiền quy đổi
        If tdbg2.RowCount > 0 Then
            If Number(tdbg2.Columns(COL2_ConvertedAmount).FooterText) <> Number(tdbg.Columns(COL_ConvertedAmount).FooterText) Then
                tdbg2(tdbg2.RowCount - 1, COL2_ConvertedAmount) = Number(tdbg2(tdbg2.RowCount - 1, COL2_ConvertedAmount)) + (Number(tdbg.Columns(COL_ConvertedAmount).FooterText) - Number(tdbg2.Columns(COL2_ConvertedAmount).FooterText))
                FooterSumNew(tdbg2, COL2_ConvertedAmount)
            End If
        End If
        btnSave.Enabled = True
    End Sub

    Private Sub txtDescription_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.LostFocus
        If txtDescription.Text = "" Then Exit Sub
        'Chuyển ký đầu thành hoa
        txtDescription.Text = Strings.Left(txtDescription.Text, 1).ToUpper & txtDescription.Text.Substring(1)
    End Sub

    Dim TranArr As String = "" 'Danh sách TransactionID khi chọn phiếu

    Private Function AllowSave() As Boolean
        If _FormState = EnumFormState.FormAdd Then
            If tdbcSplitMethodNo.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(rl3("Phuong_phap_tach"))
                tdbcSplitMethodNo.Focus()
                Return False
            End If
            If tdbcVoucherTypeID.Text.Trim = "" Then
                D99C0008.MsgNotYetChoose(rl3("Loai_phieu"))
                tdbcVoucherTypeID.Focus()
                Return False
            End If
            If txtVoucherNo.Text.Trim = "" Then
                D99C0008.MsgNotYetEnter(rl3("So_phieu"))
                txtVoucherNo.Focus()
                Return False
            End If
            If c1dateVoucherDate.Value.ToString = "" Then
                D99C0008.MsgNotYetEnter(rl3("Ngay_phieu"))
                c1dateVoucherDate.Focus()
                Return False
            End If
            For i As Integer = 0 To tdbg.RowCount - 1
                If TranArr <> "" Then TranArr &= ","
                TranArr &= SQLString(tdbg(i, COL_TransactionID))
            Next
            If Not chkSplitCipNo.Checked Then
                Dim strSQL As String = ""
                strSQL &= "SELECT 1 FROM D02T0012 WITH(NOLOCK) "
                strSQL &= "WHERE TransactionID in (" & IIf(TranArr = "", "''", TranArr).ToString & ") "
                strSQL &= "AND Status=1 AND CipID ='' "
                Dim dtTemp As DataTable = ReturnDataTable(strSQL)
                If dtTemp.Rows.Count > 0 Then
                    D99C0008.MsgL3(rl3("Cac_phieu_duoc_chon_de_tach_da_duoc_xu_ly") & Space(1) & rl3("Yeu_cau_chon_lai_phieu_khac"))
                    btnNext_Click(Nothing, Nothing)
                    Return False
                End If
            End If
        End If
1:
        If tdbg2.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tdbg2.Focus()
            Return False
        End If
        For i As Integer = 0 To tdbg2.RowCount - 1
            ' update 31/5/2013 id 56796 - R nếu Mode = 3 Khi lưu không kiểm tra bắt buộc nhập

            If Not chkSplitCipNoFromVoucher.Checked Then
                If tdbg2(i, COL2_CreditAccountID).ToString = "" Then
                    D99C0008.MsgNotYetEnter(rl3("Tai_khoan_co"))
                    tdbg2.Focus()
                    tdbg2.SplitIndex = SPLIT0
                    tdbg2.Col = IndexOfColumn(tdbg2, COL2_CreditAccountID)
                    tdbg2.Row = i
                    Return False
                End If
            End If
            If Not chkSplitCipNoFromVoucher.Checked And Not chkSplitCipNo.Checked Then
                If tdbg2(i, COL2_ObjectTypeID).ToString = "" Then
                    D99C0008.MsgNotYetEnter(rl3("Loai_doi_tuong"))
                    tdbg2.Focus()
                    tdbg2.SplitIndex = SPLIT0
                    tdbg2.Col = IndexOfColumn(tdbg2, COL2_ObjectTypeID)
                    tdbg2.Row = i
                    Return False
                End If
                If tdbg2(i, COL2_ObjectID).ToString = "" Then
                    D99C0008.MsgNotYetEnter(rl3("Doi_tuong"))
                    tdbg2.Focus()
                    tdbg2.SplitIndex = SPLIT0
                    tdbg2.Col = IndexOfColumn(tdbg2, COL2_ObjectID)
                    tdbg2.Row = i
                    Return False
                End If
            End If
            If chkSplitCipNoFromVoucher.Checked Then
                If tdbg2(i, COL2_CipID).ToString = "" Then
                    D99C0008.MsgNotYetEnter(rl3("Ma_chi_phi_XDCB"))
                    tdbg2.Focus()
                    tdbg2.SplitIndex = SPLIT0
                    tdbg2.Col = IndexOfColumn(tdbg2, COL2_CipID)
                    tdbg2.Row = i
                    Return False
                End If
            End If
            If chkSplitCipNo.Checked Then
                If tdbg2(i, COL2_CipID).ToString = "" Then
                    D99C0008.MsgNotYetEnter(rl3("Ma_chi_phi_XDCB"))
                    tdbg2.Focus()
                    tdbg2.SplitIndex = SPLIT0
                    tdbg2.Col = IndexOfColumn(tdbg2, COL2_CipID)
                    tdbg2.Row = i
                    Return False
                End If
                '            Else
                '                If tdbg2(i, COL2_ObjectTypeID).ToString = "" Then
                '                    D99C0008.MsgNotYetEnter(rl3("Loai_doi_tuong"))
                '                    tdbg2.Focus()
                '                    tdbg2.SplitIndex = SPLIT0
                '                    tdbg2.Col = IndexOfColumn(tdbg2, COL2_ObjectTypeID)
                '                    tdbg2.Bookmark = i
                '                    Return False
                '                End If
                '                If tdbg2(i, COL2_ObjectID).ToString = "" Then
                '                    D99C0008.MsgNotYetEnter(rl3("Doi_tuong"))
                '                    tdbg2.Focus()
                '                    tdbg2.SplitIndex = SPLIT0
                '                    tdbg2.Col = IndexOfColumn(tdbg2, COL2_ObjectID)
                '                    tdbg2.Bookmark = i
                '                    Return False
                '                End If
            End If
            If chkPosted.Checked Then
                If tdbg2(i, COL2_DebitAccountID).ToString = "" Then
                    D99C0008.MsgNotYetEnter(rl3("Tai_khoan_no"))
                    tdbg2.Focus()
                    tdbg2.SplitIndex = SPLIT0
                    tdbg2.Col = IndexOfColumn(tdbg2, COL2_DebitAccountID)
                    tdbg2.Row = i
                    Return False
                End If
            End If
        Next
        If Format(tdbg.Columns(COL_ConvertedAmount).FooterText) <> Format(tdbg2.Columns(COL2_ConvertedAmount).FooterText) Then
            D99C0008.MsgL3(rl3("Tong_so_tien_tach_phai_bang_so_tien_tong")) ' 'Tång sç tiÒn tÀch ph¶i bÂng sç tiÒn tång.
            tdbg2.Focus()
            tdbg2.SplitIndex = SPLIT0
            tdbg2.Col = IndexOfColumn(tdbg2, COL2_ConvertedAmount)
            tdbg2.Row = 0
            Return False
        End If
        Return True
    End Function

    Private Sub SetBackColorObligatory()
        tdbcVoucherTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        txtVoucherNo.BackColor = COLOR_BACKCOLOROBLIGATORY
        c1dateVoucherDate.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcSplitMethodNo.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        If Not chkSplitCipNoFromVoucher.Checked Then
            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_CreditAccountID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
        End If
        tdbg2.Splits(SPLIT0).DisplayColumns(COL2_CipID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
        If Not chkSplitCipNo.Checked And Not chkSplitCipNoFromVoucher.Checked Then
            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectTypeID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
            tdbg2.Splits(SPLIT0).DisplayColumns(COL2_ObjectID).Style.BackColor = COLOR_BACKCOLOROBLIGATORY
        End If

    End Sub


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0016
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 16/11/2011 04:21:21
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0016() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T0016(")
        sSQL.Append("SplitMethodNo, BatchID, VoucherNo, VoucherTypeID, VoucherDate, ")
        sSQL.Append("NotesU, DivisionID, TranMonth, TranYear, ConvertedAmount, ")
        sSQL.Append("CreateDate, CreateUserID, LastmodifyDate, LastmodifyUserID, Disabled ")
        If chkPosted.Checked OrElse chkSplitCipNo.Checked OrElse chkSplitCipNoFromVoucher.Checked Then sSQL.Append(",SplitCipNo")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(tdbcSplitMethodNo.SelectedValue) & COMMA) 'SplitMethodNo [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString(_batchID) & COMMA) 'BatchID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString(txtVoucherNo.Text) & COMMA) 'VoucherNo, varchar[50], NULL
        sSQL.Append(SQLString(tdbcVoucherTypeID.SelectedValue) & COMMA) 'VoucherTypeID, varchar[20], NULL
        sSQL.Append(SQLDateSave(c1dateVoucherDate.Value) & COMMA) 'VoucherDate, datetime, NULL
        sSQL.Append(SQLStringUnicode(txtDescription.Text, gbUnicode, True) & COMMA) 'Notes, varchar[250], NULL
        sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NULL
        sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NULL
        sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, smallint, NULL
        sSQL.Append(SQLMoney(tdbg.Columns(COL_ConvertedAmount).FooterText, DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NOT NULL
        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'LastmodifyDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'LastmodifyUserID, varchar[20], NULL
        sSQL.Append(SQLNumber(chkDisabled.Checked)) 'Disabled, bit, NOT NULL
        If chkPosted.Checked OrElse chkSplitCipNo.Checked OrElse chkSplitCipNoFromVoucher.Checked Then
            Dim iSplitCipNo As Integer = 0
            If chkSplitCipNo.Checked Then
                iSplitCipNo = 1
            ElseIf chkPosted.Checked Then
                iSplitCipNo = 2
            ElseIf chkSplitCipNoFromVoucher.Checked Then
                iSplitCipNo = 3
            End If
            sSQL.Append(COMMA & SQLNumber(iSplitCipNo)) 'SplitCipNo, tinyint, NOT NULL
        End If
        sSQL.Append(")")

        Return sSQL
    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0012s
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 21/11/2011 08:27:50
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0012s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        Dim sTrans As String = ""
        Dim iCount As Integer = tdbg2.RowCount
        Dim iFirstTrans As Long = 0
        For i As Integer = 0 To tdbg2.RowCount - 1
            sTrans = CreateIGENewS("D02T0012", "TransactionID", "02", "TE", gsStringKey, sTrans, iCount, iFirstTrans)
            tdbg2(i, COL2_TransactionID) = sTrans

            sSQL.Append("Insert Into D02T0012(")
            sSQL.Append("TransactionID, DivisionID, ModuleID, VoucherTypeID, VoucherNo, ")
            sSQL.Append("VoucherDate, TranMonth, TranYear, CurrencyID, ")
            sSQL.Append("ExchangeRate, DebitAccountID,CreditAccountID, OriginalAmount, ConvertedAmount, ")
            sSQL.Append(" Status, CreateUserID, CreateDate,LastModifyUserID, LastModifyDate,")
            sSQL.Append(" ObjectTypeID, ObjectID, BatchID, ")
            sSQL.Append(" Ana01ID, Ana02ID, Ana03ID, Ana04ID, Ana05ID, Ana06ID, Ana07ID, Ana08ID,Ana09ID, Ana10ID,")
            sSQL.Append(" Posted, Internal, DescriptionU,  NotesU ")
            If chkSplitCipNo.Checked OrElse chkSplitCipNoFromVoucher.Checked Then sSQL.Append(", CipID ")
            If chkSplitCipNoFromVoucher.Checked Then sSQL.Append(", TransactionTypeID ")
            sSQL.Append(") Values(")
            sSQL.Append(SQLString(tdbg2(i, COL2_TransactionID)) & COMMA) 'TransactionID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString("02") & COMMA) 'ModuleID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbcVoucherTypeID.SelectedValue) & COMMA) 'VoucherTypeID, varchar[20], NULL
            sSQL.Append(SQLString(txtVoucherNo.Text) & COMMA) 'VoucherNo, varchar[50], NULL
            sSQL.Append(SQLDateSave(c1dateVoucherDate.Value) & COMMA) 'VoucherDate, datetime, NULL
            If chkSplitCipNoFromVoucher.Checked Then
                sSQL.Append(SQLNumber(giFirstTranMonth) & COMMA)
                sSQL.Append(SQLNumber(giFirstTranYear) & COMMA)
            Else
                sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NULL
                sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, smallint, NULL
            End If

            sSQL.Append(SQLString(DxxFormat.BaseCurrencyID) & COMMA) 'CurrencyID, varchar[20], NOT NULL
            sSQL.Append(1 & COMMA) 'ExchangeRate, money, NOT NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_DebitAccountID)) & COMMA) 'DebitAccountID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_CreditAccountID)) & COMMA) 'CreditAccountID, varchar[20], NULL
            sSQL.Append(SQLMoney(tdbg2(i, COL2_ConvertedAmount), DxxFormat.D90_ConvertedDecimals) & COMMA) 'OriginalAmount, money, NULL
            sSQL.Append(SQLMoney(tdbg2(i, COL2_ConvertedAmount), DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL
            sSQL.Append(SQLNumber(0) & COMMA) 'Status, tinyint, NOT NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
            sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
            sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_ObjectTypeID)) & COMMA) 'ObjectTypeID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_ObjectID)) & COMMA) 'ObjectID, varchar[20], NULL
            sSQL.Append(SQLString(_batchID) & COMMA) 'BatchID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_Ana01ID)) & COMMA) 'Ana01ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_Ana02ID)) & COMMA) 'Ana02ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_Ana03ID)) & COMMA) 'Ana03ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_Ana04ID)) & COMMA) 'Ana04ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_Ana05ID)) & COMMA) 'Ana05ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_Ana06ID)) & COMMA) 'Ana06ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_Ana07ID)) & COMMA) 'Ana07ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_Ana08ID)) & COMMA) 'Ana08ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_Ana09ID)) & COMMA) 'Ana09ID, varchar[50], NULL
            sSQL.Append(SQLString(tdbg2(i, COL2_Ana10ID)) & COMMA) 'Ana10ID, varchar[50], NULL
            sSQL.Append(SQLNumber(chkPosted.Checked) & COMMA) 'Posted, tinyint, NOT NULL
            sSQL.Append(SQLNumber(0) & COMMA) 'Internal, tinyint, NOT NULL
            sSQL.Append(SQLStringUnicode(tdbg2(i, COL2_Description), gbUnicode, True) & COMMA) 'Description, varchar[500], NULL
            sSQL.Append(SQLStringUnicode(txtDescription.Text, gbUnicode, True)) 'NotesU, nvarchar, NOT NULL
            If chkSplitCipNo.Checked OrElse chkSplitCipNoFromVoucher.Checked Then sSQL.Append(COMMA & SQLString(tdbg2(i, COL2_CipID))) 'CipID, varchar[20], NULL
            If chkSplitCipNoFromVoucher.Checked Then
                sSQL.Append(COMMA & SQLString("SDXDCB"))
            End If
            sSQL.Append(")")
            If chkPosted.Checked OrElse chkSplitCipNo.Checked OrElse chkSplitCipNoFromVoucher.Checked Then
                sSQL.Append(vbCrLf & "UPDATE D02T0100 SET Status = 1 WHERE CipID = " & SQLString(tdbg2(i, COL2_CipID)))
            End If
            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0016
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 16/11/2011 04:26:11
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0016() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0016 Set ")
        sSQL.Append("SplitMethodNo = " & SQLString(tdbcSplitMethodNo.SelectedValue) & COMMA) '[KEY], varchar[20], NOT NULL
        sSQL.Append("VoucherDate = " & SQLDateSave(c1dateVoucherDate.Value) & COMMA) 'datetime, NULL
        sSQL.Append("NotesU = " & SQLStringUnicode(txtDescription.Text, gbUnicode, True) & COMMA) 'varchar[250], NULL
        sSQL.Append("DivisionID = " & SQLString(gsDivisionID) & COMMA) 'varchar[20], NULL
        sSQL.Append("TranMonth = " & SQLNumber(giTranMonth) & COMMA) 'tinyint, NULL
        sSQL.Append("TranYear = " & SQLNumber(giTranYear) & COMMA) 'smallint, NULL
        sSQL.Append("ConvertedAmount = " & SQLMoney(tdbg.Columns(COL_ConvertedAmount).Text, DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NOT NULL
        sSQL.Append("LastmodifyDate = GetDate()" & COMMA) 'datetime, NULL
        sSQL.Append("LastmodifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NULL
        sSQL.Append("Disabled = " & SQLNumber(chkDisabled.Checked) & COMMA) 'bit, NOT NULL
        Dim iSplitCipNo As Integer = 0
        If chkSplitCipNo.Checked Then
            iSplitCipNo = 1
        ElseIf chkPosted.Checked Then
            iSplitCipNo = 2
        ElseIf chkSplitCipNoFromVoucher.Checked Then
            iSplitCipNo = 3
        End If
        sSQL.Append("SplitCipNo = " & SQLNumber(iSplitCipNo)) 'tinyint, NOT NULL
        sSQL.Append(" Where ")
        sSQL.Append("BatchID = " & SQLString(_batchID))

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0012s
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 21/11/2011 09:20:43
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0012s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        For i As Integer = 0 To tdbg2.RowCount - 1
            sSQL.Append("Update D02T0012 Set ")
            sSQL.Append("DebitAccountID = " & SQLString(tdbg2(i, COL2_DebitAccountID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("CreditAccountID = " & SQLString(tdbg2(i, COL2_CreditAccountID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("ObjectTypeID = " & SQLString(tdbg2(i, COL2_ObjectTypeID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("ObjectID = " & SQLString(tdbg2(i, COL2_ObjectID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("Ana01ID = " & SQLString(tdbg2(i, COL2_Ana01ID)) & COMMA) 'varchar[50], NULL
            sSQL.Append("Ana02ID = " & SQLString(tdbg2(i, COL2_Ana02ID)) & COMMA) 'varchar[50], NULL
            sSQL.Append("Ana03ID = " & SQLString(tdbg2(i, COL2_Ana03ID)) & COMMA) 'varchar[50], NULL
            sSQL.Append("Ana04ID = " & SQLString(tdbg2(i, COL2_Ana04ID)) & COMMA) 'varchar[50], NULL
            sSQL.Append("Ana05ID = " & SQLString(tdbg2(i, COL2_Ana05ID)) & COMMA) 'varchar[50], NULL
            sSQL.Append("Ana06ID = " & SQLString(tdbg2(i, COL2_Ana06ID)) & COMMA) 'varchar[50], NULL
            sSQL.Append("Ana07ID = " & SQLString(tdbg2(i, COL2_Ana07ID)) & COMMA) 'varchar[50], NULL
            sSQL.Append("Ana08ID = " & SQLString(tdbg2(i, COL2_Ana08ID)) & COMMA) 'varchar[50], NULL
            sSQL.Append("Ana09ID = " & SQLString(tdbg2(i, COL2_Ana09ID)) & COMMA) 'varchar[50], NULL
            sSQL.Append("Ana10ID = " & SQLString(tdbg2(i, COL2_Ana10ID)) & COMMA) 'varchar[50], NULL
            If chkSplitCipNo.Checked OrElse chkSplitCipNoFromVoucher.Checked Then sSQL.Append("CipID = " & SQLString(tdbg2(i, COL2_CipID)) & COMMA) 'varchar[20], NULL
            sSQL.Append("Posted = " & SQLNumber(chkPosted.Checked) & COMMA) 'tinyint, NOT NULL
            sSQL.Append("DescriptionU = " & SQLStringUnicode(tdbg2(i, COL2_Description), gbUnicode, True) & COMMA) 'nvarchar, NOT NULL
            sSQL.Append("NotesU = " & SQLStringUnicode(txtDescription.Text, gbUnicode, True)) 'nvarchar, NOT NULL
            sSQL.Append(" Where ")
            sSQL.Append("TransactionID = " & SQLString(tdbg2(i, COL2_TransactionID)))
            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function



    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        If Not AllowSave() Then Exit Sub

        'Kiểm tra Ngày phiếu có phù hợp với kỳ kế toán hiện tại không (gọi hàm CheckVoucherDateInPeriod)
        If Not CheckVoucherDateInPeriod(c1dateVoucherDate.Text) Then c1dateVoucherDate.Focus() : Exit Sub

        btnSave.Enabled = False
        btnClose.Enabled = False
        btnChooseVoucher.Enabled = False
        btnSplit.Enabled = False

        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        Select Case _FormState
            Case EnumFormState.FormAdd
                If _batchID = "" Then _batchID = CreateIGE("D02T0012", "BatchID", "02", "BE", gsStringKey)
                ''Kiểm tra phiếu 
                If tdbcVoucherTypeID.Columns("Auto").Text = "1" And bEditVoucherNo = False Then 'Tự động
                    txtVoucherNo.Text = CreateIGEVoucherNoNew(tdbcVoucherTypeID, "D02T0012", _batchID)
                Else 'Không sinh tự động hay có nhấn F2
                    If bEditVoucherNo = False Then
                        If CheckDuplicateVoucherNoNew("D02", "D02T0012", _batchID, txtVoucherNo.Text) Then
                            Me.Cursor = Cursors.Default
                            btnSave.Enabled = True
                            btnClose.Enabled = True
                            btnChooseVoucher.Enabled = True
                            btnSplit.Enabled = True
                            txtVoucherNo.Focus()
                            Exit Sub
                        End If
                    Else 'Có nhấn F2 để sửa số phiếu
                        'SQLInsertD02T5558(_batchID, sOldVoucherNo, txtVoucherNo.Text)
                        InsertD02T5558(_batchID, sOldVoucherNo, txtVoucherNo.Text)
                    End If
                    InsertVoucherNoD91T9111(txtVoucherNo.Text, "D02T0012", _batchID)
                End If
                bEditVoucherNo = False
                sOldVoucherNo = ""
                bFirstF2 = False
                ''****************************************
                sSQL.Append(SQLInsertD02T0016().ToString & vbCrLf)
                sSQL.Append(SQLInsertD02T0012s().ToString & vbCrLf)
                sSQL.Append(SQLUpdateD02T0100s.ToString)
                '  sSQL.Append(vbCrLf & "UPDATE D02T0100 SET Status = 2 WHERE CipID = ''")

                'Lưu LastKey của Số phiếu xuống Database (gọi hàm CreateIGEVoucherNo bật cờ True)
                'Kiểm tra trùng Số phiếu (gọi hàm CheckDuplicateVoucherNo)
                'Nếu tra trùng Số phiếu thì bật
                'btnSave.Enabled = True
                'btnClose.Enabled = True

            Case EnumFormState.FormEdit
                sSQL.Append(SQLUpdateD02T0016.ToString & vbCrLf)
                sSQL.Append(SQLUpdateD02T0012s)
        End Select

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            btnClose.Enabled = True
            _bSavedOK = True
            Select Case _FormState
                Case EnumFormState.FormAdd
                    _keyID = _batchID
                    Dim strSQL As String = ""
                    strSQL &= "UPDATE D02T0012 SET SplitBatchID = " & SQLString(_batchID) & ", Status = 1  WHERE isnull(TransactionID,'') in (" & IIf(TranArr = "", "''", TranArr).ToString & ") " & vbCrLf
                    strSQL &= " DELETE D02T0017 WHERE BatchIDTemp =" & SQLString(_batchID)
                    ExecuteSQL(strSQL)

                    btnNext.Enabled = True
                    ' btnChooseVoucher.Enabled = True
                    btnNext.Focus()
                Case EnumFormState.FormEdit
                    btnSave.Enabled = True ' chỉ sáng khi Chọn lại nút Tách
                    btnClose.Focus()
            End Select
        Else
            If _FormState = EnumFormState.FormAdd Then
                DeleteVoucherNoD91T9111_Transaction(txtVoucherNo.Text, "D02T0012", "VoucherNo", tdbcVoucherTypeID, bEditVoucherNo)
            End If
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
            btnChooseVoucher.Enabled = True
            btnSplit.Enabled = True
        End If
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        TranArr = ""
        _batchID = ""
        ClearAllValue()
        btnChooseVoucher.Focus()
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0100s
    '# Created User: Hoàng Nhân
    '# Created Date: 09/10/2013 10:24:51
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0100s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        dtGrid.AcceptChanges()
        Dim dr() As DataRow = dtGrid.Select("OriginalCipID <> ''")
        If dr.Length > 0 Then sSQL.Append("-- Cap nhat trang thai cho ma XDCB duoc tach" & vbCrLf)
        For i As Integer = 0 To dr.Length - 1
            sSQL.Append("Update D02T0100 Set ")
            sSQL.Append("Status = 2") 'tinyint, NULL
            sSQL.Append(" Where ")
            sSQL.Append("CipID = " & SQLString(dr(i).Item("OriginalCipID")))
            sSQL.Append(" AND NOT EXISTS (SELECT TOP 1 1  FROM 	D02T0012 T12 ")
            sSQL.Append(" INNER JOIN D02T0016 T16 ON T12.BatchID = T16.BatchID ")
            sSQL.Append(" WHERE T12.CipID = " & SQLString(dr(i).Item("OriginalCipID")) & " AND T12.VoucherNo = " & SQLString(txtVoucherNo.Text) & ")")
            sSQL.Append("")

            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function


End Class