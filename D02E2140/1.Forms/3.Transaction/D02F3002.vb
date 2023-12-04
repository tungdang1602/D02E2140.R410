Imports System
Public Class D02F3002

#Region "Const of tdbg"
    Private Const COL_BatchID As String = "BatchID"                   ' BatchID
    Private Const COL_VoucherTypeID As String = "VoucherTypeID"       ' Loại phiếu
    Private Const COL_VoucherNo As String = "VoucherNo"               ' Số phiếu
    Private Const COL_VoucherDate As String = "VoucherDate"           ' Ngày phiếu
    Private Const COL_SplitMethodNo As String = "SplitMethodNo"       ' Mã phương pháp
    Private Const COL_SplitMethodName As String = "SplitMethodName"   ' Tên phương pháp
    Private Const COL_ConvertedAmount As String = "ConvertedAmount"   ' Số tiền
    Private Const COL_Disabled As String = "Disabled"                 ' KSD
    Private Const COL_SplitCipNo As String = "SplitCipNo"             ' Tách CP XDCB
    Private Const COL_CreateUserID As String = "CreateUserID"         ' CreateUserID
    Private Const COL_CreateDate As String = "CreateDate"             ' CreateDate
    Private Const COL_LastModifyDate As String = "LastModifyDate"     ' LastModifyDate
    Private Const COL_LastModifyUserID As String = "LastModifyUserID" ' LastModifyUserID
#End Region

    Dim dtGrid As DataTable
    Dim sFilter As New System.Text.StringBuilder()
    Dim bRefreshFilter As Boolean = False 'Cờ bật set FilterText =""

    Private Sub D09F2250_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                UseEnterAsTab(Me)
                Exit Sub
        End Select

    End Sub

    Private Sub D09F2250_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        SetShortcutPopupMenu(Me, tbrTableToolStrip, ContextMenuStrip1)
        tdbg_NumberFormat()
        Loadlanguage()
        ResetColorGrid(tdbg, SPLIT0, tdbg.Splits.Count - 1)
        ResetSplitDividerSize(tdbg)
        InputDateInTrueDBGrid(tdbg, COL_VoucherDate)
        LoadTDBGrid()
        SetResolutionForm(Me, ContextMenuStrip1)
        Me.Cursor = Cursors.Default
    End Sub


    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Danh_sach_phieu_tach_chi_phi_-_D02F3002") & UnicodeCaption(gbUnicode) 'Danh sÀch phiÕu tÀch chi phÛ - D02F3002
        '================================================================ 
        tdbg.Columns("VoucherTypeID").Caption = rl3("Loai_phieu") 'Loại phiếu
        tdbg.Columns("VoucherNo").Caption = rl3("So_phieu") 'Số phiếu
        tdbg.Columns("VoucherDate").Caption = rl3("Ngay_phieu") 'Ngày phiếu
        tdbg.Columns("SplitMethodNo").Caption = rl3("Ma_phuong_phap") 'Mã phương pháp
        tdbg.Columns("SplitMethodName").Caption = rl3("Ten_phuong_phap") 'Tên phương pháp
        tdbg.Columns("ConvertedAmount").Caption = rl3("So_tien") 'Số tiền
        tdbg.Columns("Disabled").Caption = rl3("KSD") 'KSD
        tdbg.Columns("SplitCipNo").Caption = rl3("Tach_CP_XDCB") 'Tách CP XDCB

    End Sub



#Region "LoadTDBGrid"

    Private Sub LoadTDBGrid(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKeyID As String = "")
        Dim sSQL As String = ""
        sSQL = "SELECT A.BatchID, A.DivisionID, A.SplitMethodNo, B.SplitMethodName" & UnicodeJoin(gbUnicode) & " as SplitMethodName, A.Disabled, "
        sSQL &= "A.VoucherTypeID, A.VoucherNo, A.VoucherDate, A.ConvertedAmount, A.SplitCipNo, A.CreateDate, A.CreateUserID, A.LastModifyDate, A.LastModifyUserID" & vbCrLf
        sSQL &= "FROM D02T0016 A WITH(NOLOCK) " & vbCrLf
        sSQL &= " INNER JOIN D02T0014 B WITH(NOLOCK) ON A.SplitMethodNo = B.SplitMethodNo " & vbCrLf
        sSQL &= "WHERE DivisionID = " & SQLString(gsDivisionID) & " AND TranMonth = " & giTranMonth & " AND TranYear = " & giTranYear
        sSQL &= " ORDER BY A.SplitMethodNo"
        dtGrid = ReturnDataTable(sSQL)
        gbEnabledUseFind = dtGrid.Rows.Count > 0
        LoadDataSource(tdbg, dtGrid, gbUnicode)
        ReLoadTDBGrid()
        If sKeyID <> "" Then
            Dim dt1 As DataTable = dtGrid.DefaultView.ToTable
            Dim dr() As DataRow = dt1.Select(COL_BatchID & "=" & SQLString(sKeyID), dt1.DefaultView.Sort)
            If dr.Length > 0 Then tdbg.Row = dt1.Rows.IndexOf(dr(0)) 'dùng tdbg.Bookmark có thể không đúng
            If Not tdbg.Focused Then tdbg.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
        End If

    End Sub

    Private Sub tdbg_NumberFormat()
        tdbg.Columns(COL_ConvertedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
    End Sub

#End Region

#Region "Active Find Client - List All "
    Private WithEvents Finder As New D99C1001
    Private sFind As String = ""
    Dim dtCaptionCols As DataTable

    Private Sub tsbFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbFind.Click, mnsFind.Click, tsmFind.Click
        gbEnabledUseFind = True
        '*****************************************
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbg.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        Dim Arr As New ArrayList
        AddColVisible(tdbg, SPLIT0, Arr, , , , gbUnicode)
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

    Private Sub tsbListAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbListAll.Click, mnsListAll.Click, tsmListAll.Click
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
        CheckMenu(Me.Name, tbrTableToolStrip, tdbg.RowCount, gbEnabledUseFind, True, ContextMenuStrip1)
        FooterTotalGrid(tdbg, COL_VoucherDate)
        FooterSumNew(tdbg, COL_ConvertedAmount)
    End Sub

#End Region

#Region "Menu Bar"

    Private Sub tsbAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbAdd.Click, mnsAdd.Click, tsmAdd.Click
        Dim frm As New D02F3003

        With frm
            .FormState = EnumFormState.FormAdd
            .ShowDialog()

            If .bSavedOK Then
                Dim sKey As String = frm.KeyID
                LoadTDBGrid(True, sKey)
            End If
            .Dispose()
        End With
    End Sub

    Private Sub tsbEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbEdit.Click, mnsEdit.Click, tsmEdit.Click
        If Not CheckStore(SQLStoreD02P0700) Then Exit Sub
        If Not AllowEdit_Delete() Then Exit Sub
        Dim frm As New D02F3003
        With frm
            .batchID = tdbg.Columns(COL_BatchID).Text
            .FormState = EnumFormState.FormEdit
            .ShowDialog()
            .Dispose()

            If .bSavedOK Then LoadTDBGrid(False, tdbg.Columns(COL_BatchID).Text)
        End With
    End Sub

    Private Sub tsbView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbView.Click, tsmView.Click, mnsView.Click
        Dim frm As New D02F3003
        With frm
            .batchID = tdbg.Columns(COL_BatchID).Text
            .FormState = EnumFormState.FormView
            .ShowDialog()
            .Dispose()
        End With
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1401
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 16/11/2011 04:11:19
    '# Modified User: 
    '# Modified Date: 
    '# Description: Kiểm tra trước khi xóa chưa Tách XDCB
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1401() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P1401 "
        sSQL &= SQLString(tdbg.Columns(COL_BatchID).Text) & COMMA 'BatchID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gsLanguage) 'Language, tinyint, NOT NULL
        Return sSQL
    End Function

    Private Function AllowEdit_Delete() As Boolean
        If L3Bool(tdbg.Columns(COL_SplitCipNo).Text) Then
            Dim strSQL As String = ""
            strSQL &= "SELECT 1 FROM D02T0012 WITH(NOLOCK) " & vbCrLf
            strSQL &= "INNER JOIN D02T0100 WITH(NOLOCK) ON D02T0012.CipID = D02T0100.CipID" & vbCrLf
            strSQL &= "WHERE D02T0012.BatchID = " & SQLString(tdbg.Columns(COL_BatchID).Text) & _
                    " And (D02T0012.Status = 1 OR D02T0100.Status = 2) And D02T0012.DivisionID=" & SQLString(gsDivisionID)
            Dim dtTemp As DataTable = ReturnDataTable(strSQL)
            If dtTemp.Rows.Count > 0 Then
                D99C0008.MsgL3("Dữ liệu đã được xử lý." & Space(1) & rL3("MSG000023"))
                Return False
            End If
        Else
            If Not CheckStore(SQLStoreD02P1401()) Then Return False
        End If
        Return True
    End Function

    Private Function AllowDelete() As Boolean
        If Not CheckStore(SQLStoreD02P0700) Then Return False
        If Not AllowEdit_Delete() Then Return False
        Return True
    End Function

    Private Sub tsbDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDelete.Click, tsmDelete.Click, mnsDelete.Click
        If Not AllowDelete() Then Exit Sub

        If AskDelete() = Windows.Forms.DialogResult.No Then Exit Sub
        If D99C0008.MsgAsk(rL3("Canh_bao_Du_lieu_se_bi_xoa_vinh_vien_khoi_co_so_du_lieu") & Space(1) & rL3("MSG000027") & vbCrLf & rL3("Nhan_Yes_de_dong_y_nhan_No_de_huy_bo")) = Windows.Forms.DialogResult.No Then Exit Sub

        Dim sSQL As String = ""
        sSQL = SQLUpdateD02T0100().ToString & vbCrLf
        sSQL &= " UPDATE T00 SET Status = 0" & vbCrLf
        sSQL &= " FROM    D02T0100 T00 WITH(NOLOCK)" & vbCrLf
        sSQL &= " WHERE   Isnull(CipID,'') IN (SELECT CipID FROM D02T0012 WITH(NOLOCK) WHERE BatchID = " & SQLString(tdbg.Columns(COL_BatchID).Text) & " )"
        sSQL &= " AND NOT EXISTS(SELECT TOP 1 1 FROM D02T0012 T12 WITH(NOLOCK) WHERE Isnull(T12.CipID,'') = Isnull(T00.CipID,'') AND BatchID <> " & SQLString(tdbg.Columns(COL_BatchID).Text) & ")" & vbCrLf

        sSQL &= "Delete D02T0016  Where  BatchID = " & SQLString(tdbg.Columns(COL_BatchID).Text) & vbCrLf

        sSQL &= "Delete D02T0012  Where  BatchID = " & SQLString(tdbg.Columns(COL_BatchID).Text) & vbCrLf

        sSQL &= "UPDATE D02T0012 SET Status = 0, SplitBatchID = '' "
        sSQL &= "Where  SplitBatchID = " & SQLString(tdbg.Columns(COL_BatchID).Text) & " " & vbCrLf

        Dim bResult As Boolean = ExecuteSQL(sSQL)
        If bResult Then
            DeleteOK()
            DeleteVoucherNoD91T9111(tdbg.Columns(COL_VoucherNo).Text, "D02T0012", "VoucherNo")
            DeleteGridEvent(tdbg, dtGrid, gbEnabledUseFind)
            ResetGrid()
        Else
            DeleteNotOK()
        End If
    End Sub

    Private Sub tsbClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub

    Private Sub tsbSysInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, mnsSysInfo.Click, tsmSysInfo.Click
        ShowSysInfoDialog(tdbg.Columns(COL_CreateUserID).Text, tdbg.Columns(COL_CreateDate).Text, tdbg.Columns(COL_LastModifyUserID).Text, tdbg.Columns(COL_LastModifyDate).Text)
    End Sub
#End Region

#Region "Events tdbg"

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

    Private Sub tdbg_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.DoubleClick
        If tdbg.FilterActive Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        Me.Cursor = Cursors.WaitCursor
        If e.KeyCode = Keys.Enter Then
            If tdbg.FilterActive Then Me.Cursor = Cursors.Default : Exit Sub
            If tsbEdit.Enabled Then
                tsbEdit_Click(sender, Nothing)
            ElseIf tsbView.Enabled Then
                tsbView_Click(sender, Nothing)
            End If
        End If
        HotKeyCtrlVOnGrid(tdbg, e) 'Nhấn Ctrl + V trên lưới 'có trong D99X0000
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tdbg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        Select Case tdbg.Columns(tdbg.Col).DataField
            Case COL_Disabled, COL_SplitCipNo  'Chặn Ctrl + V trên cột Check
                e.Handled = CheckKeyPress(e.KeyChar)
            Case COL_ConvertedAmount
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub

#End Region

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P0700
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 11/11/2011 10:33:10
    '# Modified User: 
    '# Modified Date: 
    '# Description: Kiểm tra trước khi sửa/Xóa
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P0700() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P0700 "
        sSQL &= SQLString(tdbg.Columns(COL_BatchID).Text) & COMMA 'BatchID, varchar[20], NOT NULL
        sSQL &= SQLString(gsLanguage)
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0100
    '# Created User: Nguyễn Thị Ánh
    '# Created Date: 11/11/2011 10:36:31
    '# Modified User: 
    '# Modified Date: 
    '# Description: Xóa
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0100() As StringBuilder
        Dim sSQL As New StringBuilder
        sSQL.Append("Update D02T0100 Set ")
        sSQL.Append("Status = " & SQLNumber(0)) 'tinyint, NULL
        sSQL.Append(" Where ")
        sSQL.Append("CipID IN (SELECT cipID FROM D02V1000 WHERE SplitBatchId = " & SQLString(tdbg.Columns(COL_BatchID).Text) & ") ")
        Return sSQL
    End Function

End Class