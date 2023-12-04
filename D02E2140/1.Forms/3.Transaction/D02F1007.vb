'#-------------------------------------------------------------------------------------
'# Created Date: 26/09/2007 3:18:51 PM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 26/09/2007 3:18:51 PM
'# Modify User: Trần Thị ÁiTrâm
'#-------------------------------------------------------------------------------------
Imports System.Text
Imports System

Public Class D02F1007

#Region "Const of tdbg"
    Private Const COL_CipNo As Integer = 0             '  Mã XDCB
    Private Const COL_CipName As Integer = 1           ' Tên hạng mục
    Private Const COL_StatusCIP As Integer = 2         ' Tình trạng
    Private Const COL_AccountID As Integer = 3         ' Tài khoản tập hợp
    Private Const COL_VoucherTypeID As Integer = 4     ' Loại phiếu
    Private Const COL_VoucherNo As Integer = 5         ' Số phiếu
    Private Const COL_VoucherDate As Integer = 6       ' Ngày phiếu
    Private Const COL_RefDate As Integer = 7           ' Ngày hóa đơn
    Private Const COL_SeriNo As Integer = 8            ' Số Sêri
    Private Const COL_RefNo As Integer = 9             ' Số hóa đơn
    Private Const COL_ObjectTypeID As Integer = 10     ' Loại đối tượng
    Private Const COL_ObjectID As Integer = 11         ' Mã đối tượng
    Private Const COL_ObjectName As Integer = 12       ' Tên đối tượng
    Private Const COL_Description As Integer = 13      ' Diễn giải
    Private Const COL_DebitAccountID As Integer = 14   ' TK nợ
    Private Const COL_CurrencyID As Integer = 15       ' Loại tiền
    Private Const COL_ExchangeRate As Integer = 16     ' Tỷ giá
    Private Const COL_OriginalAmount As Integer = 17   ' Nguyên tệ
    Private Const COL_ConvertedAmount As Integer = 18  ' Qui đổi
    Private Const COL_BatchID As Integer = 19          ' BatchID
    Private Const COL_ModuleID As Integer = 20         ' ModuleID
    Private Const COL_CreateUserID As Integer = 21     ' CreateUserID
    Private Const COL_CreateDate As Integer = 22       ' CreateDate
    Private Const COL_LastModifyUserID As Integer = 23 ' LastModifyUserID
    Private Const COL_LastModifyDate As Integer = 24   ' LastModifyDate
    Private Const COL_Status As Integer = 25           ' Status
    Private Const COL_CipID As Integer = 26            ' CipID
#End Region

    Private dt As DataTable
    Private sBatchID As String
    Private sCipID As String
    Private iRPLang As Integer
    Dim sFilter As New StringBuilder()
    Dim iColumns() As Integer = {COL_ConvertedAmount, COL_OriginalAmount}
    Private iPerD02F5605 As Integer = -1

    Private Sub D02F1007_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
    End Sub

    Private Sub D02F1007_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetShortcutPopupMenu(Me, tbrTableToolStrip, ContextMenuStrip1)
        iPerD02F5605 = ReturnPermission("D02F5605")
        Loadlanguage()
        InputDateInTrueDBGrid(tdbg, COL_VoucherDate, COL_RefDate)
        InputbyUnicode(Me, gbUnicode)
        ResetColorGrid(tdbg)
        tdbg_NumberFormat()
        gbEnabledUseFind = False
        LoadTDBGrid()
        iRPLang = CInt(GetSetting("D02", "Options", "nRPLang", "0"))
    SetResolutionForm(Me)

SetResolutionForm(Me, ContextMenuStrip1)
Me.Cursor = Cursors.Default
End Sub

#Region "Active Find Client - List All "
    Private WithEvents Finder As New D99C1001
    Private sFind As String = ""
    Dim dtCaptionCols As DataTable

    Private Sub tsbFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbFind.Click, tsmFind.Click, mnsFind.Click
        'Chuẩn hóa D09U1111 : Tìm kiếm dùng table caption có sẵn
        tdbg.UpdateData()
        'If dtCaptionCols Is Nothing OrElse dtCaptionCols.Rows.Count < 1 Then 'Incident 72333
        'Những cột bắt buộc nhập
        Dim Arr As New ArrayList
        AddColVisible(tdbg, SPLIT0, Arr, , False, False, gbUnicode)
        'Tạo tableCaption: đưa tất cả các cột trên lưới có Visible = True vào table 
        dtCaptionCols = CreateTableForExcelOnly(tdbg, Arr)
        'End If
        ShowFindDialogClient(Finder, dtCaptionCols, Me.Name, "0", gbUnicode)
        'Dim sSQL As String = ""
        'gbEnabledUseFind = True
        'sSQL = "Select * From D02V1234 "
        'sSQL &= "Where FormID = " & SQLString(Me.Name) & "And Language = " & SQLString(gsLanguage)
        'ShowFindDialogClient(Finder, sSQL, gbUnicode)
    End Sub

    Private Sub Finder_FindClick(ByVal ResultWhereClause As Object) Handles Finder.FindClick
        If ResultWhereClause Is Nothing Then Exit Sub
        sFind = ResultWhereClause.ToString()
        ReLoadTDBGrid()
    End Sub

    Private Sub tsbListAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbListAll.Click, tsmListAll.Click, mnsListAll.Click
        sFind = ""
        ResetFilter(tdbg, sFilter, bRefreshFilter)
        ReLoadTDBGrid()
    End Sub

    'Import dữ liệu
    Private Sub tsbImportData_Click(sender As Object, e As EventArgs) Handles tsbImportData.Click, tsmImportData.Click, mnsImportData.Click
        '22/8/2018, 	Phạm Thị Mỹ Tiên: id 111968-Bổ sung import số dư XDCB
        Me.Cursor = Cursors.WaitCursor
        If CallShowDialogD80F2090(D02, "D02F5605", "D02F1007") Then
            LoadTDBGrid(True)
        End If
        Me.Cursor = Cursors.Default
    End Sub


    Private Sub ReLoadTDBGrid()
        Dim strFind As String = sFind
        If sFilter.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter.ToString
        dtGrid.DefaultView.RowFilter = strFind
        ResetGrid()
    End Sub

    Private Sub ResetGrid()
        '28/4/2017, id 96484-Lỗi khóa sổ vẫn sáng menu thêm sửa xóa D02
        'CheckMenu(Me.Name, tbrTableToolStrip, tdbg.RowCount, gbEnabledUseFind,False, ContextMenuStrip1)
        CheckMenu(Me.Name, tbrTableToolStrip, tdbg.RowCount, gbEnabledUseFind, True, ContextMenuStrip1)

        '22/8/2018, 	Phạm Thị Mỹ Tiên: id 111968-Bổ sung import số dư XDCB
        tsbImportData.Enabled = iPerD02F5605 >= EnumPermission.Add And Not gbClosed
        tsmImportData.Enabled = iPerD02F5605 >= EnumPermission.Add And Not gbClosed
        mnsImportData.Enabled = iPerD02F5605 >= EnumPermission.Add And Not gbClosed


        FooterTotalGrid(tdbg, COL_CipNo)
        FooterSum(tdbg, iColumns, , True)
        ' TotalFooter()
    End Sub

    'Private Sub LoadGridFind1(ByVal C1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dtRoot As DataTable, ByVal sClauseFind As String)
    '    Dim dtF As DataTable
    '    dtF = dtRoot.Copy
    '    dtF.DefaultView.RowFilter = sClauseFind.Replace("N'", "'")

    '    'Dim strFind As String
    '    'strFind = sFind
    '    'If sFilter.ToString() <> "" Then
    '    '    If strFind <> "" Then
    '    '        strFind &= " And " & sFilter.ToString
    '    '    Else
    '    '        strFind &= sFilter.ToString
    '    '    End If
    '    'End If

    '    LoadDataSource(C1Grid, dtF)

    'End Sub
#End Region
    Private Sub tdbg_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tdbg.MouseClick
        iHeight = e.Location.Y
    End Sub

    Private Sub tdbg_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.DoubleClick
        If iHeight <= tdbg.Splits(0).ColumnCaptionHeight Then Exit Sub
        If tdbg.FilterActive Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        If tsbEdit.Enabled Then
            tsbEdit_Click(sender, Nothing)
        ElseIf tsbView.Enabled Then
            tsbView_Click(sender, Nothing)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tdbg_FilterChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg.FilterChange
        'Try
        '    If (dt Is Nothing) Then Exit Sub
        '    sFilter = New StringBuilder("")
        '    Dim dc As C1.Win.C1TrueDBGrid.C1DataColumn
        '    For Each dc In Me.tdbg.Columns
        '        Select Case dc.DataType.Name
        '            Case "DateTime"
        '                If dc.FilterText.Length = 10 Then
        '                    If sFilter.Length > 0 Then sFilter.Append(" AND ")
        '                    Dim sClause As String = ""
        '                    sClause = "(" & dc.DataField & " >= #" & DateSave(CDate(dc.FilterText)) & "#"
        '                    sClause &= " And " & dc.DataField & " < #" & DateSave(CDate(dc.FilterText).AddDays(1)) & "# )"
        '                    sFilter.Append(sClause)
        '                End If

        '            Case "Boolean"
        '                If dc.FilterText.Length > 0 Then
        '                    If sFilter.Length > 0 Then sFilter.Append(" AND ")
        '                    sFilter.Append((dc.DataField + " = " + "'" + dc.FilterText + "'"))
        '                End If

        '            Case "String"
        '                If dc.FilterText.Length > 0 Then
        '                    If sFilter.Length > 0 Then sFilter.Append(" AND ")
        '                    sFilter.Append((dc.DataField + " like " + "'%" + dc.FilterText.Replace("'", "''") + "%'"))
        '                End If

        '            Case "Decimal", "Byte", "Integer"
        '                If dc.FilterText.Length > 0 Then
        '                    If sFilter.Length > 0 Then sFilter.Append(" AND ")
        '                    sFilter.Append((dc.DataField + " = " + "" + dc.FilterText + ""))
        '                End If
        '        End Select
        '    Next
        '    'Filter the data 
        '    If sFilter.ToString() <> "" And sFind <> "" Then
        '        dt.DefaultView.RowFilter = sFilter.ToString() & " AND " & sFind
        '    ElseIf sFind <> "" Then
        '        dt.DefaultView.RowFilter = sFind
        '    ElseIf sFind = "" Then
        '        dt.DefaultView.RowFilter = sFilter.ToString()
        '    End If

        '    CheckMenu(PARA_FormIDPermission, C1CommandHolder, tdbg.RowCount, gbEnabledUseFind, False)
        '    FooterTotalGrid(tdbg, COL_CipNo)
        '    FooterSum(tdbg, iColumns)
        '    Me.Cursor = Cursors.Default

        'Catch ex As Exception
        '    MessageBox.Show(ex.Message & " - " & ex.Source)
        'End Try
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

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        If e.KeyCode = Keys.Enter Then
            If tdbg.FilterActive Then Exit Sub
            If tsbEdit.Enabled Then
                tsbEdit_Click(sender, Nothing)
            ElseIf tsbView.Enabled Then
                tsbView_Click(sender, Nothing)
            End If
        End If
        HotKeyCtrlVOnGrid(tdbg, e)
    End Sub

    Private Sub tsbAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbAdd.Click, tsmAdd.Click, mnsAdd.Click
        Dim f As New D02F1008
        With f
            .BatchID = ""
            .TransactionTypeID = "SDXDCB"
            .CipID = ""
            .FormState = EnumFormState.FormAdd
            .ShowDialog()
            .Dispose()
            sBatchID = .BatchID
            sCipID = .CipID
        End With
        If gbSavedOK Then
            LoadTDBGrid(True, sCipID)
        End If
    End Sub

   Private Sub tsbEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbEdit.Click, tsmEdit.Click, mnsEdit.Click
        'Kiểm tra điều kiện sửa
        If Not CheckStore(SQLStoreD02P1070(0)) Then Exit Sub 'ID 89282 02.08.2016
        If tdbg.Columns(COL_StatusCIP).Text = "2" Then
            D99C0008.MsgL3(rl3("Phieu_nay_da_duoc_xu_ly_ban_khong_the_sua"))
            Exit Sub
        Else
            Dim f As New D02F1008
            With f
                .BatchID = tdbg.Columns(COL_BatchID).Text
                .TransactionTypeID = "SDXDCB"
                .CipID = tdbg.Columns(COL_CipID).Text
                .FormState = EnumFormState.FormEdit
                .ShowDialog()
                .Dispose()
                sCipID = .CipID
            End With
            If gbSavedOK Then
                LoadTDBGrid(False, tdbg.Columns(COL_CipID).Text)
            End If
        End If
    End Sub

    Private Sub tsbDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbDelete.Click, tsmDelete.Click, mnsDelete.Click
        Dim sSQL As String = ""
        Dim bResult As Boolean
        If AskDelete() = Windows.Forms.DialogResult.No Then Exit Sub
        If Not CheckStore(SQLStoreD02P1070(1)) Then Exit Sub 'ID 89282 02.08.2016
        sSQL &= "Delete From D02T0012" & vbCrLf
        sSQL &= "Where BatchID=" & SQLString(tdbg.Columns(COL_BatchID).Text) & vbCrLf
        sSQL &= "And TransactionTypeID='SDXDCB' And ModuleID='02' And VoucherNo=" & SQLString(tdbg.Columns(COL_VoucherNo).Text) & vbCrLf
        sSQL &= "If Not Exists (Select 1 From    D02T0012 WITH(NOLOCK) Where CipID =" & SQLString(tdbg.Columns(COL_CipID).Text) & ")" & vbCrLf
        sSQL &= "Begin " & vbCrLf
        sSQL &= "       Update D02T0100" & vbCrLf
        sSQL &= "       Set Status=0" & vbCrLf
        sSQL &= "       Where CipID=" & SQLString(tdbg.Columns(COL_CipID).Text) & vbCrLf
        sSQL &= "End"
        bResult = ExecuteSQL(sSQL)
        If bResult = True Then
            DeleteOK()
            DeleteVoucherNoD91T9111(tdbg.Columns(COL_VoucherNo).Text, "D02T0012", "VoucherNo")
            DeleteGridEvent(tdbg, dtGrid, gbEnabledUseFind)
        Else
            DeleteNotOK()
        End If

    End Sub


    Private Sub tsbView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbView.Click, tsmView.Click, mnsView.Click
        Dim f As New D02F1008
        With f
            .BatchID = tdbg.Columns(COL_BatchID).Text
            .TransactionTypeID = "SDXDCB"
            .CipID = tdbg.Columns(COL_CipID).Text
            .FormState = EnumFormState.FormView
            .ShowDialog()
            .Dispose()
        End With
    End Sub

    Private Sub tsbSysInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, tsmSysInfo.Click, mnsSysInfo.Click
        ShowSysInfoDialog(tdbg.Columns(COL_CreateUserID).Text, tdbg.Columns(COL_CreateDate).Text, tdbg.Columns(COL_LastModifyUserID).Text, tdbg.Columns(COL_LastModifyDate).Text)
    End Sub

    Private Sub tsbClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub


    Dim bRefreshFilter As Boolean = False 'Cờ bật set FilterText =""
    Private dtGrid, dtGrid1 As DataTable
    Dim iHeight As Integer = 0
    Private Sub LoadTDBGrid(Optional ByVal FlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        Dim sSQL As New StringBuilder("")
        sSQL.Append(" Select  A.BatchID, ModuleID, VoucherTypeID, VoucherDate, VoucherNo, RefDate, SeriNo, " & vbCrLf)
        sSQL.Append("  RefNo, A.ObjectTypeID,  A.ObjectID, Object.ObjectName" & UnicodeJoin(gbUnicode) & " as ObjectName, A.Description" & UnicodeJoin(gbUnicode) & " as Description, DebitAccountID, B. ")
        sSQL.Append(" AccountID, CurrencyID, ExchangeRate, sum(OriginalAmount) OriginalAmount,    sum(ConvertedAmount) ")
        sSQL.Append(" ConvertedAmount, A.CipID, B.CipNo, CipName" & UnicodeJoin(gbUnicode) & " as CipName,Convert(varchar(250), B.Status) as StatusCIP, A.Status, B.CreateUserID, B.CreateDate, B.LastModifyUserID, MAX(A.LastModifyDate) AS  LastModifyDate" & vbCrLf)
        sSQL.Append(" From D02T0012 A WITH(NOLOCK) Inner Join D02T0100 B WITH(NOLOCK) On A.CipID = B.CipID " & vbCrLf)
        sSQL.Append("  		Left Join Object WITH(NOLOCK) On A.ObjectTypeID = Object.ObjectTypeID " & vbCrLf)
        sSQL.Append(" And A.ObjectID = Object.ObjectID" & vbCrLf)
        sSQL.Append(" Where TransactionTypeID = 'SDXDCB' And B.DivisionID = " & SQLString(gsDivisionID) & vbCrLf)
        sSQL.Append(" Group By  A.BatchID, ModuleID, VoucherTypeID, VoucherDate, VoucherNo, RefDate, SeriNo, ")
        sSQL.Append(" Object.ObjectName" & UnicodeJoin(gbUnicode) & ", B. AccountID, RefNo, A.ObjectTypeID, A.ObjectID,  A.Description" & UnicodeJoin(gbUnicode) & ", DebitAccountID, ")
        sSQL.Append(" CurrencyID,  ExchangeRate, A.CipID, B.CipNo, CipName" & UnicodeJoin(gbUnicode) & ", B.Status, A.Status, B.CreateUserID, B.CreateDate, B.LastModifyUserID" & vbCrLf)
        dtGrid = ReturnDataTable(sSQL.ToString)
        gbEnabledUseFind = dtGrid.Rows.Count > 0
        If FlagAdd Then
            ' Thêm mới thì gán sFind ="" và gán FilterText =’’
            ResetFilter(tdbg, sFilter, bRefreshFilter)
            sFind = ""
        End If
        LoadDataSource(tdbg, dtGrid, gbUnicode)
        For i As Integer = 0 To dtGrid.Rows.Count - 1
            Select Case dtGrid.Rows(i).Item("StatusCIP").ToString
                Case "0"
                    tdbg(i, COL_StatusCIP) = rl3("Moi_thiet_lap")
                    If gbUnicode Then
                        tdbg(i, COL_StatusCIP) = ConvertVniToUnicode(rl3("Moi_thiet_lap"))
                    End If

                Case "1"
                    tdbg(i, COL_StatusCIP) = rl3("Dang_tap_hop")
                    If gbUnicode Then
                        tdbg(i, COL_StatusCIP) = ConvertVniToUnicode(rl3("Dang_tap_hop"))
                    End If

                Case "2"
                    tdbg(i, COL_StatusCIP) = rl3("Da_ket_thuc")
                    If gbUnicode Then
                        tdbg(i, COL_StatusCIP) = ConvertVniToUnicode(rl3("Da_ket_thuc"))
                    End If
            End Select
        Next
        ReLoadTDBGrid()
        If sKey <> "" Then
            Dim dt1 As DataTable = dtGrid.DefaultView.ToTable
            Dim dr() As DataRow = dt1.Select(tdbg.Columns(COL_CipID).DataField & "=" & SQLString(sKey), dt1.DefaultView.Sort)
            If dr.Length > 0 Then tdbg.Row = dt1.Rows.IndexOf(dr(0)) 'dùng tdbg.Bookmark có thể không đúng
            If Not tdbg.Focused Then tdbg.Focus() 'Nếu con trỏ chưa đứng trên lưới thì Focus về lưới
        End If
    End Sub

    Private Sub tdbg_NumberFormat()
        tdbg.Columns(COL_ExchangeRate).NumberFormat = DxxFormat.ExchangeRateDecimals
        tdbg.Columns(COL_OriginalAmount).NumberFormat = DxxFormat.DecimalPlaces
        tdbg.Columns(COL_ConvertedAmount).NumberFormat = DxxFormat.D90_ConvertedDecimals
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Phieu_nhap_so_du_chi_phi_XDCB_-_D02F1007") & UnicodeCaption(gbUnicode) 'PhiÕu nhËp sç d§ chi phÛ XDCB - D02F1007
        '================================================================ 
        '================================================================ 
        tdbg.Columns("CipNo").Caption = rl3("Ma_XDCB") ' Mã XDCB
        tdbg.Columns("CipName").Caption = rl3("Ten_hang_muc") 'Tên hạng mục
        tdbg.Columns("StatusCIP").Caption = rl3("Tinh_trang") 'Tình trạng
        tdbg.Columns("AccountID").Caption = rl3("Tai_khoan_tap_hop") 'Tài khoản tập hợp
        tdbg.Columns("VoucherTypeID").Caption = rl3("Loai_phieu") 'Loại phiếu
        tdbg.Columns("VoucherNo").Caption = rl3("So_phieu") 'Số phiếu
        tdbg.Columns("VoucherDate").Caption = rl3("Ngay_phieu") 'Ngày phiếu
        tdbg.Columns("RefDate").Caption = rl3("Ngay_hoa_don") 'Ngày hóa đơn
        tdbg.Columns("SeriNo").Caption = rl3("So_Seri") 'Số Sêri
        tdbg.Columns("RefNo").Caption = rl3("So_hoa_don") 'Số hóa đơn
        tdbg.Columns("ObjectTypeID").Caption = rl3("Loai_doi_tuong") 'rl3("Ma_loai_doi_tuong") 'Mã loại đối tượng
        tdbg.Columns("ObjectID").Caption = rl3("Ma_doi_tuong") 'Mã đối tượng
        tdbg.Columns("ObjectName").Caption = rl3("Ten_doi_tuong") 'Tên đối tượng
        tdbg.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbg.Columns("DebitAccountID").Caption = rl3("TK_no") 'rl3("Tai_khoan_no") 'Tài khoản nợ
        tdbg.Columns("CurrencyID").Caption = rl3("Loai_tien") 'Loại tiền
        tdbg.Columns("OriginalAmount").Caption = rl3("Nguyen_te") 'Nguyên tệ
        tdbg.Columns("ConvertedAmount").Caption = rl3("Quy_doi") 'Qui đổi
        tdbg.Columns("ExchangeRate").Caption = rl3("Ty_gia") 'Tỷ giá
        '================================================================ 
       
    End Sub

    Private Sub TotalFooter()
        Dim dTotalOrginalAmount As Double = 0
        Dim dTotalConvertedAmount As Double = 0

        If tdbg.RowCount <= 0 Then Exit Sub
        For i As Int32 = 0 To tdbg.RowCount - 1
            dTotalOrginalAmount += Number(SQLNumber(tdbg(i, COL_OriginalAmount).ToString, DxxFormat.DefaultNumber2))
            dTotalConvertedAmount += Number(SQLNumber(tdbg(i, COL_ConvertedAmount).ToString, DxxFormat.DefaultNumber2))
        Next
        tdbg.Columns(COL_OriginalAmount).FooterText = SQLNumber(dTotalOrginalAmount.ToString, DxxFormat.DefaultNumber2)
        tdbg.Columns(COL_ConvertedAmount).FooterText = SQLNumber(dTotalConvertedAmount.ToString, DxxFormat.DefaultNumber2)
        ResetFooterGrid(tdbg)
    End Sub

    Private Sub tsbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbPrint.Click, mnsPrint.Click, tsmPrint.Click
        If giAppMode = 1 Then
            'PrintDataWs()
            Exit Sub
        Else
            PrintData()
        End If
    End Sub

    Private Sub PrintData()
        Me.Cursor = Cursors.WaitCursor

        Dim report As New D99C1003
        Dim conn As New SqlConnection(gsConnectionString)
        Dim sReportName As String = "D02R1007"
        Dim sSubReportName As String = "D02R0000"
        Dim sReportCaption As String = ""
        Dim sPathReport As String = ""
        Dim sSQL As String = ""
        Dim sSQLSub As String = ""

        sSQLSub = "Select Top 1 * From D91T0025 WITH(NOLOCK)"
        sReportCaption = rl3("Phieu_nhap_so_du_chi_phi_XDCB") & " - " & sReportName
        sPathReport = UnicodeGetReportPath(gbUnicode, D02Options.ReportLanguage, "") & sReportName & ".rpt"
        UnicodeSubReport(sSubReportName, sSQLSub, gsDivisionID, gbUnicode)
        'If iRPLang = 0 Then
        '    sPathReport = Application.StartupPath & "\XReports\" & sReportName & ".rpt"
        'ElseIf iRPLang = 1 Then
        '    sPathReport = Application.StartupPath & "\XReports\VE-XReports\" & sReportName & ".rpt"
        'ElseIf iRPLang = 2 Then
        '    sPathReport = Application.StartupPath & "\XReports\E-XReports\" & sReportName & ".rpt"
        'End If
        sSQL = SQLStoreD02P1007()
        With report
            .OpenConnection(conn)
            .AddSub(sSQLSub, sSubReportName & ".rpt")
            .AddMain(sSQL)
            .PrintReport(sPathReport, sReportCaption)
        End With
        Me.Cursor = Cursors.Default
    End Sub

    'Private Sub PrintDataWs()
    '    Me.Cursor = Cursors.WaitCursor

    '    Dim report As New D99C0009
    '    CallWebService.Url = gsAppServer & "D91W0000.asmx"
    '    CallWebService.Timeout = nWSTimeOut

    '    Dim sReportName As String = "D02R1007"
    '    Dim sSubReportName As String = "D02R0000"
    '    Dim sReportCaption As String = ""
    '    Dim sPathReport As String = ""
    '    Dim sSQL As String = ""
    '    Dim sSQLSub As String = ""
    '    sSQLSub = "Select Top 1 * From D91T0025 WITH(NOLOCK)"
    '    UnicodeSubReport(sSubReportName, sSQLSub, gsDivisionID, gbUnicode)
    '    sReportCaption = rl3("Phieu_nhap_so_du_chi_phi_XDCB") & " - " & sReportName
    '    sSQL = SQLStoreD02P1007()
    '    sSQL = " Select * From D02V1007 "
    '    With report
    '        .OpenConnection(CallWebService.Url, gsUserID, gsCompanyID, gsWSSPara01, gsWSSPara02, gsWSSPara03, gsWSSPara04, gsWSSPara05)
    '        .AddMain(sSQL)
    '        .PrintReport(sReportName & ".rpt", sReportCaption & " - " & sReportName)
    '    End With
    '    Me.Cursor = Cursors.Default
    'End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1007
    '# Created User: 
    '# Created Date: 02/01/2008 04:15:14
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1007() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P1007 "
        sSQL &= SQLString(tdbg.Columns(COL_CipID).Text) & COMMA 'CipID, varchar[20], NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_BatchID).Text) & COMMA 'BatchID, varchar[20], NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'DivisionID, varchar[20], NOT NULL
        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1070
    '# Created User: KIM LONG
    '# Created Date: 02/08/2016 09:59:24
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1070(ByVal iMode As Integer) As String
        Dim sSQL As String = ""
        sSQL &= ("-- store kiem tra " & vbCrLf)
        sSQL &= "Exec D02P1070 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[50], NOT NULL
        sSQL &= SQLString(Me.Name) & COMMA 'FormID, varchar[50], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLNumber(gsLanguage) & COMMA 'Language, int, NOT NULL
        sSQL &= SQLNumber(gbUnicode) & COMMA 'CodeTable, tinyint, NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_CipID).Text) & COMMA 'CipID, varchar[50], NOT NULL
        sSQL &= SQLNumber(iMode) 'Mode, int, NOT NULL
        Return sSQL
    End Function



    Private Sub tdbg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        Select Case tdbg.Col
            Case COL_VoucherDate, COL_RefDate
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Number)
            Case COL_ConvertedAmount, COL_OriginalAmount, COL_ExchangeRate
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub

    Private Sub c1dateDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
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

   
End Class