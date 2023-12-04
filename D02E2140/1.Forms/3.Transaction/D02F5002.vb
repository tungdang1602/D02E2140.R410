Imports System
Imports System.Text
Imports System.Collections
Public Class D02F5002

#Region "Const of tdbg"
    Private Const COL_Choose As Integer = 0      ' Chọn
    Private Const COL_NormID As Integer = 1      ' NormID
    Private Const COL_NormNo As Integer = 2      ' Tên BĐM
    Private Const COL_Description As Integer = 3 ' Diễn giải
#End Region

    Private _batchID As String = ""
    Public Property BatchID() As String 
        Get
            Return _batchID
        End Get
        Set(ByVal Value As String )
            _batchID = Value
        End Set
    End Property

    Private dt As DataTable

    Private Sub D02F5002_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Them ngay 28/202013 theo ID 54631 cuả Bảo Trân bởi Văn Vinh
        Dim sSQL As String = "DELETE D91T2024 WHERE	UserID = " & SQLString(gsUserID) & " AND FormID = 'D02F5002'"
        ExecuteSQL(sSQL)
    End Sub


    Private Sub D02F5002_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
    End Sub

    Private Sub D02F5002_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        LoadTDBCombo()
        InputbyUnicode(Me, gbUnicode)
        LoadTDBGrid()
        c1dateVoucherDate.Value = Now
        SetBackColorObligatory()
        lblProcess.Visible = False
        tdbcFromCCodeID.Enabled = False
        tdbcToCCodeID.Enabled = False
        pgb1.Visible = False
        Loadlanguage()
        LoadDefault()
    SetResolutionForm(Me)
Me.Cursor = Cursors.Default
End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Tinh_khau_hao_-_D02F5002") & UnicodeCaption(gbUnicode) 'TÛnh khÊu hao - D02F5002
        '================================================================ 
        lblDescription.Text = rl3("Dien_giai_chi_tiet") 'Diễn giải chi tiết
        lblVoucherNo.Text = rl3("So_phieu") 'Số phiếu
        lblVoucherTypeID.Text = rl3("Loai_phieu") 'Loại phiếu
        lblVoucherDate.Text = rl3("Ngay_phieu") 'Ngày phiếu
        lblProcess.Text = rl3("Xu_ly") 'Xử lý
        lblNotes.Text = rl3("Dien_giai_phieu") 'Diễn giải phiếu
        '================================================================ 
        btnCalculate.Text = rl3("_Tinh") '&Tính
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        grp4.Text = rl3("Tinh_khau_hao_theo_ma_phan_tich")
        lblTypeCodeID.Text = rl3("Loai_phan_tich")
        lblFromCCodeID.Text = rl3("Ma_phan_tich")
        '================================================================ 
        chkUseBOM.Text = rl3("Su_dung_bo_dinh_muc_de_tinh_khau_hao") 'Sử dụng bộ định mức để tính khấu hao
        chkReCalculate.Text = rl3("Tinh_lai_cac_TSCD_da_duoc_tinh") 'Tính lại các TSCĐ đã được tính
        chkCheckAssetID.Text = rl3("Ma_tai_san_co_dinh") 'Mã tài sản cố định
        chkCheckAssetName.Text = rl3("Ten_tai_san_co_dinh") 'Tên tài sản cố định
        '================================================================ 
        GroupBox1.Text = rl3("But_toan_khau_hao") 'Bút toán khấu hao
        '================================================================ 
        tdbcVoucherTypeID.Columns("VoucherTypeID").Caption = rl3("Loai_phieu") 'Loại phiếu
        tdbcVoucherTypeID.Columns("VoucherTypeName").Caption = rl3("Dien_giai") 'Diễn giải
        '================================================================ 
        tdbg.Columns("Choose").Caption = rl3("Chon") 'Chọn
        tdbg.Columns("NormNo").Caption = rl3("Ten_BDM") 'Ten BDM
        tdbg.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        '================================================================ 
        If gbUnicode Then
            txtDescription.Text = rl3("Tinh_khau_hao_tai_san_co_dinhU") 'Tính khấu hao tài sản cố định
            txtNotes.Text = rl3("Tinh_khau_hao_tai_san_co_dinhU") 'Tính khấu hao tài sản cố định
        Else
            txtDescription.Text = ConvertUnicodeToVni(rl3("Tinh_khau_hao_tai_san_co_dinhU")) 'Tính khấu hao tài sản cố định
            txtNotes.Text = ConvertUnicodeToVni(rl3("Tinh_khau_hao_tai_san_co_dinhU")) 'Tính khấu hao tài sản cố định
        End If
    End Sub

    Private Sub SetBackColorObligatory()
        tdbcVoucherTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        txtVoucherNo.BackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub
    Private Sub LoadDefault()
        For i As Integer = 0 To dtVoucherTypeID.Rows.Count - 1
            If dtVoucherTypeID.Rows(i).Item("FormID").ToString = "D02F5002" Then
                Dim sFormID As String = dtVoucherTypeID.Rows(i).Item("VoucherTypeID").ToString
                tdbcVoucherTypeID.Text = sFormID
                Exit Sub
            End If
        Next
    End Sub
    Dim dtCodeID As DataTable
    Dim dtVoucherTypeID As DataTable
    Private Sub LoadTDBCombo()
        'Load(tdbcVoucherTypeID)
        'LoadVoucherTypeID(tdbcVoucherTypeID, D02, , gbUnicode)
        dtVoucherTypeID = ReturnDataTable(ReturnTableVoucherTypeID("D02", gsDivisionID, "", gbUnicode))
        LoadDataSource(tdbcVoucherTypeID, dtVoucherTypeID, gbUnicode)
        'Thêm ngày 16/10/2012 theo incident 51667 của Bảo Trân bởi Văn Vinh
        'Load tdbcTypeCodeID
        Dim sSQL As String = ""
        'Load tdbcFromCCodeID
        sSQL = "SELECT 	'%' AS AcodeID, " & AllName & " AS Description, '%' AS TypeCodeID, 0 AS DisplayOrder UNION"
        sSQL &= " SELECT  AcodeID, Description" & UnicodeJoin(gbUnicode) & " As Description, TypeCodeID,  1 AS DisplayOrder           "
        sSQL &= " FROM   D02T0041 WITH(NOLOCK) WHERE  Type = 'A' ORDER BY DisplayOrder, TypeCodeID, AcodeID"
        dtCodeID = ReturnDataTable(sSQL)
        LoadDataSource(tdbcFromCCodeID, dtCodeID, gbUnicode)
        'Load tdbcToCCodeID
        LoadDataSource(tdbcToCCodeID, dtCodeID, gbUnicode)
        sSQL = "SELECT 	TypeCodeID, VieTypeCodeName" & UnicodeJoin(gbUnicode) & " AS Description FROM D02T0040 WITH(NOLOCK) WHERE 	Type = 'A' AND Disabled = 0 ORDER BY TypeCodeID "
        LoadDataSource(tdbcTypeCodeID, sSQL, gbUnicode)
    End Sub
    Public Function ReturnTableVoucherTypeID(ByVal sModuleID As String, ByVal DivisionID As String, ByVal sEditTransTypeID As String, Optional ByVal bUseUnicode As Boolean = False) As String
        Dim sSQL As String = "--Do nguon cho combo loai phieu" & vbCrLf
        sSQL &= "Select T01.VoucherTypeID, " & IIf(bUseUnicode, "VoucherTypeNameU", "VoucherTypeName").ToString & " as VoucherTypeName, Auto, S1Type, S1, S2Type, S2, " & vbCrLf
        sSQL &= "S3, S3Type, OutputOrder, OutputLength, Separator, T40.FormID " & vbCrLf
        sSQL &= "From D91T0001 T01 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Left Join D02T0080 T40 WITH(NOLOCK) ON T01.VoucherTypeID = T40.VoucherTypeID" & vbCrLf
        sSQL &= "Where Use" & sModuleID & " = 1 And Disabled = 0 " & vbCrLf
        If DivisionID <> "" Then sSQL &= "AND( VoucherDivisionID='' Or VoucherDivisionID = " & SQLString(DivisionID) & ") " & vbCrLf
        'Load cho trường hợp Sửa, Xem
        If sEditTransTypeID <> "" Then
            sSQL &= "Or T01.VoucherTypeID = " & SQLString(sEditTransTypeID) & vbCrLf
        End If
        sSQL &= "Order By VoucherTypeID"
        Return sSQL
    End Function

#Region "Events tdbcVoucherTypeID"

    Private Sub tdbcVoucherTypeID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.LostFocus
        If tdbcVoucherTypeID.FindStringExact(tdbcVoucherTypeID.Text) = -1 Then tdbcVoucherTypeID.Text = ""
    End Sub

#End Region

#Region "Events tdbcTypeCodeID"

    Private Sub tdbcTypeCodeID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.LostFocus
        If tdbcTypeCodeID.FindStringExact(tdbcTypeCodeID.Text) = -1 Then tdbcTypeCodeID.Text = ""
        If tdbcTypeCodeID.Text = "" Then
            tdbcFromCCodeID.Enabled = False
            tdbcToCCodeID.Enabled = False
        Else
            tdbcFromCCodeID.Enabled = True
            tdbcToCCodeID.Enabled = True
        End If
    End Sub

    Private Sub tdbcTypeCodeID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcTypeCodeID.SelectedValueChanged
        If tdbcTypeCodeID.Text = "" Then
            tdbcFromCCodeID.Enabled = False
            tdbcToCCodeID.Enabled = False
        Else
            tdbcFromCCodeID.Enabled = True
            tdbcToCCodeID.Enabled = True
        End If
        LoadDataSource(tdbcFromCCodeID, ReturnTableFilter(dtCodeID, "TypeCodeID = " & SQLString(ReturnValueC1Combo(tdbcTypeCodeID)) & " or TypeCodeID = '%' ", True), gbUnicode)
        LoadDataSource(tdbcToCCodeID, ReturnTableFilter(dtCodeID, "TypeCodeID = " & SQLString(ReturnValueC1Combo(tdbcTypeCodeID)) & " or TypeCodeID = '%' ", True), gbUnicode)
    End Sub
#End Region


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        ExecuteSQL(SQLDeleteD91T9009)
        Me.Close()
    End Sub

    Private Sub LoadTDBGrid()
        Dim sSQL As String = ""
        sSQL = "SELECT 	Convert(Bit,0) as Choose, NormID, NormNo,"
        sSQL &= " CASE WHEN NormID <> 'BDMKHDT' "
        sSQL &= " THEN Description" & UnicodeJoin(gbUnicode)
        sSQL &= " ELSE " & IIf(gbUnicode, "N'" & rl3("Bo_dinh_muc_khau_hao_duong_thang"), "'" & rl3("Bo_dinh_muc_khao_hao_duong_thangV")).ToString & "' END AS Description, "
        sSQL &= " CASE WHEN NormID = 'BDMKHDT' "
        sSQL &= " THEN 0  ELSE 1 END AS DisplayOrder" & vbCrLf

        sSQL &= " FROM  D02T0054 WITH(NOLOCK)" & vbCrLf
        sSQL &= " WHERE Disabled = 0 OR NormID = 'BDMKHDT'" & vbCrLf
        sSQL &= " ORDER BY 	DisplayOrder,NormNo"
        dt = ReturnDataTable(sSQL)
        LoadDataSource(tdbg, dt, gbUnicode)
    End Sub

#Region "Events tdbcVoucherTypeID with txtVoucherNo"
    Private Sub tdbcVoucherTypeID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.Close
        If tdbcVoucherTypeID.FindStringExact(tdbcVoucherTypeID.Text) = -1 Then
            tdbcVoucherTypeID.Text = ""
            txtVoucherNo.Text = ""
        End If
    End Sub

    Private Sub tdbcVoucherTypeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.SelectedValueChanged
        bFirstF2 = False
        bEditVoucherNo = False

        If tdbcVoucherTypeID.SelectedValue Is Nothing OrElse tdbcVoucherTypeID.Text = "" Then
            txtVoucherNo.Text = ""
            txtDescription.Text = ""
            ReadOnlyControl(txtVoucherNo)
            Exit Sub
        Else
            If gbUnicode Then
                txtDescription.Text = rl3("Tinh_khau_hao_tai_san_co_dinhU")
            Else
                txtDescription.Text = rl3("Tinh_khau_hao_tai_san_co_dinh")
            End If

        End If
        If tdbcVoucherTypeID.Columns("Auto").Text = "1" Then 'Sinh tự động
            txtVoucherNo.Text = CreateIGEVoucherNo(tdbcVoucherTypeID, False)
            ReadOnlyControl(txtVoucherNo)
        Else 'Không sinh tự động
            txtVoucherNo.Text = ""
            UnReadOnlyControl(txtVoucherNo, True)
        End If
    End Sub

#End Region

    Private Sub btnCalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalculate.Click
        Dim sSQL As String
        lblProcess.Visible = True
        pgb1.Visible = True
        Me.Height = Me.Height + 30
        Application.DoEvents()
        pgb1.Minimum = 0
        pgb1.Maximum = 100
        If Not AllowCalculate() Then
            Me.Height = Me.Height - 30
            Exit Sub
        End If
        For i As Integer = 0 To 10
            lblProcess.Text = rl3("Xu_ly") & ":" & Space(1) & i & "%"
        Next
        pgb1.Value = 10

        If _batchID = "" Then _batchID = CreateIGE("D02T0012", "BatchID", "02", "KH", gsStringKey)

        If tdbcVoucherTypeID.Columns("Auto").Text = "1" And bEditVoucherNo = False Then 'Sinh tự động
            txtVoucherNo.Text = CreateIGEVoucherNoNew(tdbcVoucherTypeID, "D02T0012", _batchID)
        Else 'Không sinh tự động hay có nhấn F2
            If bEditVoucherNo = False Then
                'Kiểm tra trùng Số phiếu
                If CheckDuplicateVoucherNoNew("D02", "D02T0012", _batchID, txtVoucherNo.Text) = True Then btnCalculate.Enabled = True : btnClose.Enabled = True : Me.Cursor = Cursors.Default : Exit Sub
            Else 'Có nhấn F2 để sửa số phiếu
                'Insert Số phiếu vào bảng D02T5558
                InsertD02T5558(_batchID, sOldVoucherNo, txtVoucherNo.Text)
            End If
            'Insert VoucherNo vào bảng D91T9111
            InsertVoucherNoD91T9111(txtVoucherNo.Text, "D02T0012", _batchID)
        End If
        bEditVoucherNo = False
        sOldVoucherNo = ""
        bFirstF2 = False

        If chkUseBOM.Checked Then
            sSQL = SQLDeleteD91T9009() & vbCrLf
            sSQL &= SQLInsertD91T9009s().ToString
            ExecuteSQL(sSQL)
            For i As Integer = 11 To 40
                lblProcess.Text = rl3("Xu_ly") & ":" & Space(1) & i & "%"
            Next
            pgb1.Value = 40
            If CheckStore(SQLStoreD02P0008) Then
                'ExecuteSQL(SQLStoreD91P9106)
                Dim Desc4 As String = ""
                For i As Integer = 0 To tdbg.RowCount - 1
                    If L3Bool(tdbg(i, COL_Choose)) = True Then
                        Desc4 &= tdbg(i, COL_NormID).ToString & ","
                    End If
                Next
                Lemon3.D91.RunAuditLog("02", "DepCal", "01", IIf(gbUnicode = False, ConvertUnicodeToVni(rL3("Tinh_khau_hao_ky")), rL3("Tinh_khau_hao_ky")).ToString & Space(1) & giTranMonth.ToString("00") & "/" & giTranYear, _
                                                            txtVoucherNo.Text, L3String(c1dateVoucherDate.Value), Desc4.Substring(0, Desc4.Length - 1))
                For i As Integer = 41 To 100
                    lblProcess.Text = rl3("Xu_ly") & ":" & Space(1) & i & "%"
                Next
                pgb1.Value = 100
                Application.DoEvents()
                D99C0008.MsgL3(rl3("Tinh_khau_hao_thanh_cong"))
                Me.Height = Me.Height - 30
                Dim frm As New D02F5012
                frm.VoucherNo = txtVoucherNo.Text
                frm.BatchID = _batchID
                frm.ShowDialog()
                frm.Dispose()
                btnClose_Click(Nothing, Nothing)
            Else
                'Incident 	73369 cập nhật ngày 10/03/2015
                DeleteVoucherNoD91T9111_Transaction(txtVoucherNo.Text, "D02T0012", "VoucherNo", tdbcVoucherTypeID, bEditVoucherNo)
            End If
        Else
            If ExecuteSQL(SQLStoreD02P0009()) Then
                For i As Integer = 11 To 100
                    lblProcess.Text = rl3("Xu_ly") & ":" & Space(1) & i & "%"
                Next
                pgb1.Value = 100
                Application.DoEvents()
                D99C0008.MsgL3(rl3("Tinh_khau_hao_thanh_cong"))
                Me.Height = Me.Height - 30
                Dim frm As New D02F5012
                frm.VoucherNo = txtVoucherNo.Text
                frm.BatchID = _batchID
                frm.ShowDialog()
                frm.Dispose()
                btnClose_Click(Nothing, Nothing)
            Else
                'Incident 	73369 cập nhật ngày 10/03/2015
                DeleteVoucherNoD91T9111_Transaction(txtVoucherNo.Text, "D02T0012", "VoucherNo", tdbcVoucherTypeID, bEditVoucherNo)
            End If
        End If
    End Sub
    Private Function AllowCalculate() As Boolean
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
        If chkUseBOM.Checked And tdbg.RowCount > 0 Then
            Dim isChoose As Boolean = False
            For i As Integer = 0 To tdbg.RowCount - 1
                isChoose = L3Bool(tdbg(i, COL_Choose))
                If isChoose Then Exit For
            Next
            If Not isChoose Then
                D99C0008.MsgL3(rl3("Ban_phai_chon_bo_dinh_muc"))
                Return False
            End If
        End If
        If CheckDuplicateVoucherNo("D02", "D02T0012", "", txtVoucherNo.Text) Then
            Return False
        End If
        If Not CheckVoucherDateInPeriod(c1dateVoucherDate.Value.ToString) Then
            c1dateVoucherDate.Focus()
            Return False
        End If

        '  Return CheckStoreD02P0019(SQLStoreD02P0019)
        Return CheckStore(SQLStoreD02P0019)
        Return CheckStore(SQLStoreD02P0057)
        Return True
    End Function

    Private Function CheckStoreD02P0019(ByVal SQL As String) As Boolean
        Dim dt As DataTable = Nothing
        ' update 14/8/2013 id 57853 - Theo chuan CheckStore
        dt = ReturnDataTable(SQL)
        If dt.Rows.Count > 0 Then
            If dt.Rows(0).Item("Status").ToString <> "0" Then
                If dt.Rows(0).Item("Status").ToString = "2" Then
                    ' If MessageBox.Show(dt.Rows(0).Item("Message").ToString, MsgAnnouncement, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Yes Then
                    If D99C0008.MsgAsk(ConvertVietwareFToUnicode(dt.Rows(0).Item("Message").ToString)) = Windows.Forms.DialogResult.Yes Then
                        dt = Nothing
                        Return True
                    Else
                        dt = Nothing
                        Return False
                    End If
                ElseIf dt.Rows(0).Item("Status").ToString = "1" Then
                    'MessageBox.Show(dt.Rows(0).Item("Message").ToString, MsgAnnouncement, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    D99C0008.MsgL3(ConvertVietwareFToUnicode(dt.Rows(0).Item("Message").ToString))
                    dt = Nothing
                    Return False
                End If
            End If
            dt = Nothing
        Else
            D99C0008.MsgL3("Không có dòng nào trả ra từ Store")
            Return False
        End If
        Return True
    End Function

    Private Sub tdbg_HeadClick(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.HeadClick
        If e.ColIndex = COL_Choose And tdbg.RowCount > 0 Then
            Dim isChoose As Boolean = True
            For i As Integer = 0 To tdbg.RowCount - 1
                If L3Bool(tdbg(i, COL_Choose)) = False Then
                    isChoose = False
                    Exit For
                End If
            Next
            If Not isChoose Then
                For i As Integer = 0 To tdbg.RowCount - 1
                    tdbg(i, COL_Choose) = 1
                Next
            Else
                For i As Integer = 0 To tdbg.RowCount - 1
                    tdbg(i, COL_Choose) = 0
                Next
            End If
        End If
    End Sub

    Private Sub chkUseBOM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUseBOM.Click
        If chkUseBOM.Checked = True Then
            tdbg.Enabled = True
        Else
            If tdbg.RowCount > 0 Then
                For i As Integer = 0 To tdbg.RowCount - 1
                    tdbg(i, COL_Choose) = 0
                Next
            End If
            tdbg.Enabled = False
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0019
    '# Created User: Lê Sơn Long
    '# Created Date: 04/11/2010 03:34:29
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0019() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0019 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        If gsLanguage = "84" Then
            sSQL &= SQLNumber(0) & COMMA 'Language, tinyint, NOT NULL
        Else
            sSQL &= SQLNumber(1) & COMMA  'Language, tinyint, NOT NULL
        End If
        ' update 7/8/2013 id 57853
        sSQL &= SQLString(tdbcTypeCodeID.Text) & COMMA 'TypeCodeID, varchar[50], NOT NULL
        sSQL &= SQLString(tdbcFromCCodeID.Text) & COMMA 'ACodeIDFrom, varchar[50], NOT NULL
        sSQL &= SQLString(tdbcToCCodeID.Text) 'ACodeIDTo, varchar[50], NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0057
    '# Created User: Lê Sơn Long
    '# Created Date: 04/11/2010 03:41:46
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0057() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0057 "
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        If gsLanguage = "84" Then
            sSQL &= SQLNumber(0) 'Language, tinyint, NOT NULL
        Else
            sSQL &= SQLNumber(1) 'Language, tinyint, NOT NULL
        End If
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD91T9009
    '# Created User: Lê Sơn Long
    '# Created Date: 04/11/2010 03:51:05
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD91T9009() As String
        Dim sSQL As String = ""
        sSQL &= "Delete From D91T9009"
        sSQL &= " Where UserID = " & SQLString(gsUserID) & " And HostID = " & SQLString(My.Computer.Name)
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD91T9009s
    '# Created User: Lê Sơn Long
    '# Created Date: 04/11/2010 03:55:25
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD91T9009s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        Dim count As Integer = 1
        For i As Integer = 0 To tdbg.RowCount - 1
            If L3Bool(tdbg(i, COL_Choose)) Then
                sSQL.Append("Insert Into D91T9009(")
                sSQL.Append("UserID, HostID, Key01ID, Key02ID, Num01")
                sSQL.Append(") Values(")
                sSQL.Append(SQLString(gsUserID) & COMMA) 'UserID, varchar[20], NULL
                sSQL.Append(SQLString(My.Computer.Name) & COMMA) 'HostID, varchar[20], NULL
                sSQL.Append(SQLString("D02F5002") & COMMA) 'Key01ID, varchar[250], NULL
                sSQL.Append(SQLString(tdbg(i, COL_NormID)) & COMMA) 'Key02ID, varchar[250], NULL
                sSQL.Append(SQLNumber(count))
                sSQL.Append(")")
                sRet.Append(sSQL.ToString & vbCrLf)
                sSQL.Remove(0, sSQL.Length)
                count = count + 1
            End If
        Next
        Return sRet
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0008
    '# Created User: Lê Đình Thái
    '# Created Date: 09/01/2012 10:31:49
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------

    Private Function SQLStoreD02P0008() As String
        Dim sSQL As String = ""
        sSQL = "--Tinh khau hao theo BDM (V3.9)" & vbCrLf
        sSQL &= "Exec D02P0008 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsLanguage) & COMMA 'Language, varchar[20], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(tdbcVoucherTypeID.Text) & COMMA 'VoucherTypeID, varchar[20], NOT NULL
        sSQL &= SQLString(txtVoucherNo.Text) & COMMA 'VoucherNo, varchar[20], NOT NULL
        sSQL &= SQLDateSave(c1dateVoucherDate.Value) & COMMA 'VoucherDate, datetime, NOT NULL
        sSQL &= SQLString("") & COMMA 'Description, varchar[250], NOT NULL
        sSQL &= SQLNumber(chkReCalculate.Checked) & COMMA 'ReCalculate, tinyint, NOT NULL
        sSQL &= SQLString(_batchID) & COMMA 'BatchID, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'Notes, varchar[250], NOT NULL
        sSQL &= SQLNumber(chkCheckAssetID.Checked) & COMMA 'CheckAssetID, tinyint, NOT NULL
        sSQL &= SQLNumber(chkCheckAssetName.Checked) & COMMA 'CheckAssetName , tinyint, NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable , tinyint, NOT NULL
        sSQL &= SQLStringUnicode(txtDescription.Text, gbUnicode, True) & COMMA 'DescriptionU, nvarchar, NOT NULL
        sSQL &= SQLStringUnicode(txtNotes.Text, gbUnicode, True) & COMMA 'NotesU, nvarchar, NOT NULL
        'Thêm ngay f16/10/2012 theo incident 51667 của Bảo Trân bởi Văn Vinh
        sSQL &= SQLString(ReturnValueC1Combo(tdbcTypeCodeID)) & COMMA 'TypeAodeID, VARCHAR(50), NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcFromCCodeID)) & COMMA 'FromACodeID, varchar[50], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcToCCodeID)) 'ToACodeID, varchar[50], NOT NULL

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P9106
    '# Created User: Lê Sơn Long
    '# Created Date: 04/11/2010 04:12:45
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    'Private Function SQLStoreD91P9106() As String
    '    Dim Desc4 As String = ""
    '    For i As Integer = 0 To tdbg.RowCount - 1
    '        If L3Bool(tdbg(i, COL_Choose)) = True Then
    '            Desc4 &= tdbg(i, COL_NormID).ToString & ","
    '        End If
    '    Next
    '    Dim sSQL As String = ""
    '    sSQL &= "Exec D91P9106 "
    '    sSQL &= SQLDateSave(Now) & COMMA 'AuditDate, datetime, NOT NULL
    '    sSQL &= SQLString("DepCal") & COMMA 'AuditCode, varchar[20], NOT NULL
    '    sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
    '    sSQL &= SQLString("02") & COMMA 'ModuleID, varchar[2], NOT NULL
    '    sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
    '    sSQL &= SQLString("01") & COMMA 'EventID, varchar[20], NOT NULL
    '    sSQL &= SQLString(rl3("Tinh_khau_hao_ky") & Space(1) & giTranMonth.ToString("00") & "/" & giTranYear) & COMMA 'Desc1, varchar[250], NOT NULL
    '    sSQL &= SQLString(txtVoucherNo.Text) & COMMA 'Desc2, varchar[250], NOT NULL
    '    sSQL &= SQLDateSave(c1dateVoucherDate.Value) & COMMA 'Desc3, varchar[250], NOT NULL
    '    sSQL &= SQLString(Desc4.Substring(0, Desc4.Length - 1)) & COMMA 'Desc4, varchar[250], NOT NULL
    '    sSQL &= SQLString("") & COMMA 'Desc5, varchar[250], NOT NULL
    '    sSQL &= SQLNumber(0) & COMMA 'IsAuditDetail, tinyint, NOT NULL
    '    sSQL &= SQLString("") 'AuditItemID, varchar[50], NOT NULL
    '    Return sSQL
    'End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0009
    '# Created User: Lê Đình Thái
    '# Created Date: 09/01/2012 10:29:35
    '# Modified User: Văn Vinh
    '# Modified Date: 
    '# Description: Thêm một số trường vào Store
    '#---------------------------------------------------------------------------------------------------

    Private Function SQLStoreD02P0009() As String
        Dim sSQL As String = ""
        sSQL = "--Tinh khau hao khong dung bo dinh muc" & vbCrLf
        sSQL &= "Exec D02P0009 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(tdbcVoucherTypeID.Text) & COMMA 'VoucherTypeID, varchar[20], NOT NULL
        sSQL &= SQLDateSave(c1dateVoucherDate.Value) & COMMA 'VoucherDate, datetime, NOT NULL
        sSQL &= SQLString(txtVoucherNo.Text) & COMMA 'VoucherNo, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'Description, varchar[250], NOT NULL
        sSQL &= SQLString(_batchID) & COMMA  'BatchID, varchar[20], NOT NULL
        sSQL &= SQLString("") & COMMA 'Notes, varchar[250], NOT NULL
        sSQL &= SQLNumber(chkCheckAssetID.Checked) & COMMA 'CheckAssetID, tinyint, NOT NULL
        sSQL &= SQLNumber(chkCheckAssetName.Checked) & COMMA 'CheckAssetName , tinyint, NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable , tinyint, NOT NULL
        sSQL &= SQLStringUnicode(txtDescription.Text, gbUnicode, True) & COMMA 'DescriptionU, nvarchar, NOT NULL
        sSQL &= SQLStringUnicode(txtNotes.Text, gbUnicode, True) & COMMA 'NotesU, nvarchar, NOT NULL
        'Thêm ngay f16/10/2012 theo incident 51667 của Bảo Trân bởi Văn Vinh
        sSQL &= SQLString(ReturnValueC1Combo(tdbcTypeCodeID)) & COMMA 'TypeAodeID, VARCHAR(50), NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcFromCCodeID)) & COMMA 'FromACodeID, varchar[50], NOT NULL
        sSQL &= SQLString(ReturnValueC1Combo(tdbcToCCodeID)) 'ToACodeID, varchar[50], NOT NULL
        Return sSQL
    End Function


    Dim bFirstF2 As Boolean = False 'Nhấn F2 lần đầu tiên 
    Dim sOldVoucherNo As String = "" 'Lưu lại số phiếu cũ
    Dim bEditVoucherNo As Boolean = False '= True: có nhấn F2; = False: không 

    Private Sub txtVoucherNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVoucherNo.KeyDown
        If e.KeyCode = Keys.F2 Then
            'Loại phiếu hay Số phiếu = "" thì thoát
            If tdbcVoucherTypeID.Text = "" Or txtVoucherNo.Text = "" Then Exit Sub

            'Kiểm tra quyền cho trường hợp Sửa
            If ReturnPermission("D02F5558") <= 2 Then Exit Sub

            'Cho sửa Số phiếu ở trạng thái Thêm mới hay Sửa

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
            '    .TableName = "D02T0012" 'Tên bảng chứa số phiếu
            '    .VoucherNo = txtVoucherNo.Text 'Số phiếu cần sửa
            '    .VoucherID = _batchID
            '    .Mode = "0" ' Tùy theo Module, mặc định là 0
            '    .KeyID01 = "AL"
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
            '        gbSavedOK = True
            '    End If
            'End With

            Dim arrPro() As StructureProperties = Nothing
            SetProperties(arrPro, "FormIDPermission", "D02F5558")
            SetProperties(arrPro, "VoucherTypeID", ReturnValueC1Combo(tdbcVoucherTypeID))
            SetProperties(arrPro, "VoucherID", _batchID)
            SetProperties(arrPro, "Mode", 0)
            SetProperties(arrPro, "KeyID01", "AL")
            SetProperties(arrPro, "TableName", "D02T0012")
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
                gbSavedOK = True
            End If
        End If
    End Sub
#Region "Events tdbcFromCCodeID"

    Private Sub tdbcFromCCodeID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcFromCCodeID.LostFocus
        If tdbcFromCCodeID.FindStringExact(tdbcFromCCodeID.Text) = -1 Then tdbcFromCCodeID.Text = ""
    End Sub

    'Them ngay 16/10/2012 theo incident 51667 cua Bảo Trân bởi Văn Vinh
    'Private Sub tdbcFromCCodeID_KeyDown(ByVal iMode As Integer)
    '    Dim sKey As String = ""
    '    Dim f As New D91F6020
    '    With f
    '        .FormPermision = "D02F5002"
    '        .ModeSelect = "1"
    '        .SQLSelection = "SELECT TypeCodeID as SelectionGroup, ACodeID AS SelectionID, Description" & UnicodeJoin(gbUnicode) & " AS SelectionName FROM D02T0041 WHERE Type = 'A' AND TypeCodeID = " & SQLString(ReturnValueC1Combo(tdbcTypeCodeID)) & " ORDER BY ACodeID"
    '        .ShowDialog()
    '        sKey = .OutPut01
    '        .Dispose()
    '        If sKey = "" Then Exit Sub
    '        If sKey = "True" Then
    '            tdbcFromCCodeID.Text = "%"
    '            tdbcToCCodeID.Text = "%"
    '        Else
    '            sKey = sKey.Trim()
    '            Dim chuoi() As String = sKey.Split(";"c)
    '            If chuoi.Length > 0 Then
    '                If chuoi.Length = 2 Then
    '                    tdbcFromCCodeID.Text = chuoi(0)
    '                    tdbcToCCodeID.Text = chuoi(1)
    '                Else
    '                    If iMode = 1 Then
    '                        tdbcFromCCodeID.Text = chuoi(0)
    '                    ElseIf iMode = 2 Then
    '                        tdbcToCCodeID.Text = chuoi(0)
    '                    End If
    '                End If
    '            End If
    '        End If
    '    End With
    'End Sub

    Private Sub tdbcFromCCodeID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcFromCCodeID.KeyDown

        Dim StrCCodeID As String = ""
        If e.KeyCode = Keys.F2 Then
            'If tdbcFromCCodeID.Tag Is Nothing Then Exit Sub 'tdbc.Tag lưu câu SQL đổ nguồn cho combo
            Me.Cursor = Cursors.WaitCursor
            Dim sSQL As String = "SELECT TypeCodeID as SelectionGroup, ACodeID AS SelectionID, Description" & UnicodeJoin(gbUnicode) & " AS SelectionName, CONVERT(TINYINT,0) AS Choose FROM D02T0041 WITH(NOLOCK) WHERE 	Type = 'A' AND TypeCodeID = " & SQLString(ReturnValueC1Combo(tdbcTypeCodeID)) & " ORDER BY ACodeID"
            StrCCodeID = HotKeyF2D91F6020(sSQL, tdbcFromCCodeID, tdbcToCCodeID, 1) 'Gán giá trị sau khi tìm kiếm
            Me.Cursor = Cursors.Default
        End If
    End Sub
#End Region

#Region "Events tdbcToCCodeID"

    Private Sub tdbcToCCodeID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcToCCodeID.KeyDown
        Dim StrCCodeID As String = ""
        If e.KeyCode = Keys.F2 Then
            ' If tdbcToCCodeID.Tag Is Nothing Then Exit Sub 'tdbc.Tag lưu câu SQL đổ nguồn cho combo
            Me.Cursor = Cursors.WaitCursor
            Dim sSQL As String = "SELECT TypeCodeID as SelectionGroup, ACodeID AS SelectionID, Description" & UnicodeJoin(gbUnicode) & " AS SelectionName, CONVERT(TINYINT,0) AS Choose FROM D02T0041 WITH(NOLOCK) WHERE 	Type = 'A' AND TypeCodeID = " & SQLString(ReturnValueC1Combo(tdbcTypeCodeID)) & " ORDER BY ACodeID"
            StrCCodeID = HotKeyF2D91F6020(sSQL, tdbcFromCCodeID, tdbcToCCodeID, 1) 'Gán giá trị sau khi tìm kiếm
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub tdbcToCCodeID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcToCCodeID.LostFocus
        If tdbcToCCodeID.FindStringExact(tdbcToCCodeID.Text) = -1 Then tdbcToCCodeID.Text = ""
    End Sub

#End Region



End Class