'#-------------------------------------------------------------------------------------
'# Created Date: 06/12/2007 2:12:26 PM
'# Created User: 
'# Modify Date: 06/12/2007 2:12:26 PM
'# Modify User: 
'#-------------------------------------------------------------------------------------

Public Class D02F2006

#Region "Const of tdbg"
    Private Const COL_BatchID As Integer = 0           ' BatchID
    Private Const COL_AssetID As Integer = 1           ' Mã tài sản
    Private Const COL_ChangeNo As Integer = 2          ' Mã nghiệp vụ 
    Private Const COL_ChangeName As Integer = 3        ' Tên nghiệp vụ
    Private Const COL_ChangeDate As Integer = 4        ' Ngày
    Private Const COL_DecisionNo As Integer = 5        ' Số hiệu
    Private Const COL_Notes1 As Integer = 6            ' Ghi chú 1
    Private Const COL_Notes2 As Integer = 7            ' Ghi chú 2
    Private Const COL_Notes3 As Integer = 8            ' Ghi chú 3
    Private Const COL_CreateUserID As Integer = 9      ' CreateUserID
    Private Const COL_CreateDate As Integer = 10       ' CreateDate
    Private Const COL_LastModifyUserID As Integer = 11 ' LastModifyUserID
    Private Const COL_LastModifyDate As Integer = 12   ' LastModifyDate
#End Region

    Private dtGrid As DataTable
    'Private sBatchID As String

    'Private Sub btnAction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    C1ContextMenu.ShowContextMenu(Me, New Point(btnAction.Left, btnAction.Top))
    'End Sub

    Private Sub D02F2006_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                UseEnterAsTab(Me)
            Case Keys.F
                If tsbFind.Enabled Then
                    tsbFind_Click(Nothing, Nothing)
                End If
            Case Keys.A
                If tsbListAll.Enabled Then
                    tsbListAll_Click(Nothing, Nothing)
                End If
        End Select
    End Sub

    Private Sub D02F2006_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        SetShortcutPopupMenu(Me, ToolStrip1, ContextMenuStrip1)
        ResetColorGrid(tdbg)
        Loadlanguage()
        gbEnabledMenuFind = False
        LoadTDBGrid()
        SetResolutionForm(Me, ContextMenuStrip1)
        Me.Cursor = Cursors.Default
    End Sub

#Region "Active Find Client - List All "
    Private WithEvents Finder As New D99C1001
    Private sFind As String = ""

    'Private Sub mnuFind_Click(ByVal sender As Object, ByVal e As C1.Win.C1Command.ClickEventArgs)
    '    If CallMenuFromGrid(tdbg, e) = False Then Exit Sub
    '    Dim sSQL As String = ""
    '    gbEnabledUseFind = True
    '    sSQL = "Select * From D02V1234 "
    '    sSQL &= "Where FormID = " & SQLString(Me.Name) & "And Language = " & SQLString(gsLanguage)
    '    ShowFindDialogClient(Finder, sSQL)
    'End Sub

    'Private Sub Finder_FindClick(ByVal ResultWhereClause As Object) Handles Finder.FindClick
    '    If ResultWhereClause Is Nothing Then Exit Sub
    '    sFind = ResultWhereClause.ToString()
    '    ReLoadTDBGrid()
    'End Sub

    'Private Sub mnuListAll_Click(ByVal sender As Object, ByVal e As C1.Win.C1Command.ClickEventArgs)
    '    If CallMenuFromGrid(tdbg, e) = False Then Exit Sub
    '    sFind = ""
    '    ReLoadTDBGrid()
    'End Sub
    Dim dtCaptionCols As DataTable
    Private Sub tsbFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbFind.Click, tsmFind.Click, mnsFind.Click
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
        '*****************************************
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


    'Private Sub ReLoadTDBGrid()
    '    LoadGridFind(tdbg, dtGrid, sFind)
    '    CheckMenu(PARA_FormIDPermission, C1CommandHolder, tdbg.RowCount, gbEnabledUseFind, True)
    'End Sub
#End Region

    Private Sub ReLoadTDBGrid()
        Dim strFind As String = sFind
        If sFilter.ToString.Equals("") = False And strFind.Equals("") = False Then strFind &= " And "
        strFind &= sFilter.ToString
        dtGrid.DefaultView.RowFilter = strFind
        ResetGrid()
    End Sub

    Private Sub ResetGrid()
        CheckMenu(PARA_FormIDPermission, ToolStrip1, tdbg.RowCount, gbEnabledUseFind, True, ContextMenuStrip1)
        FooterTotalGrid(tdbg, COL_AssetID)
    End Sub
    'Private Sub LoadTDBGrid(Optional ByVal bFlagAdd As Boolean = False, Optional ByVal sKey As String = "")
    '    Dim sSQL As String = SQLStoreD02P4500()
    '    dt = ReturnDataTable(sSQL)
    '    LoadDataSource(tdbg, dt)
    '    CheckMenu(PARA_FormIDPermission, ToolStrip1, tdbg.RowCount, gbEnabledMenuFind, True, ContextMenuStrip1, False)
    '    If bFlagAdd Then
    '        dt.DefaultView.Sort = "BatchID"
    '        tdbg.Bookmark = dt.DefaultView.Find(sBatchID)
    '    End If
    'End Sub


    Private Sub LoadTDBGrid(Optional ByVal bFlagAdd As Boolean = False, Optional ByVal sKey As String = "")
        dtGrid = ReturnDataTable(SQLStoreD02P4500)
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
            'MessageBox.Show(ex.Message & " - " & ex.Source)
            ' tdbg.Columns(tdbg.Col).FilterText = ""
            WriteLogFile(ex.Message) 'Ghi file log TH nhập số >MaxInt cột Byte -> Không hiển thị thông báo
        End Try
    End Sub


    Private Sub tsbSysInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbSysInfo.Click, tsmSysInfo.Click, mnsSysInfo.Click
        ShowSysInfoDialog(tdbg.Columns(COL_CreateUserID).Text, tdbg.Columns(COL_CreateDate).Text, tdbg.Columns(COL_LastModifyUserID).Text, tdbg.Columns(COL_LastModifyDate).Text)
    End Sub


    Private Sub tsbClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbClose.Click
        Me.Close()
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P4500
    '# Created User: 
    '# Created Date: 05/12/2007 03:34:30
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P4500() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P4500 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLString("") & COMMA 'strFind, varchar[8000], NOT NULL
        sSQL &= SQLNumber(gbUnicode)
        Return sSQL
    End Function

    Private Sub tsbAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbAdd.Click, tsmAdd.Click, mnsAdd.Click
        Dim f As New D02F2010
        With f
            .BatchID = ""
            .FormState = EnumFormState.FormAdd
            .ShowDialog()
            If gbSavedOK Then LoadTDBGrid(True, .BatchID)
            .Dispose()
        End With
    End Sub

    Private Sub tsbDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDelete.Click, tsmDelete.Click, mnsDelete.Click
        Dim sSQL As String = ""
        'Dim iBookmark As Integer

        If AskDelete() = Windows.Forms.DialogResult.Yes Then
            'If Not IsDBNull(tdbg.Bookmark) Then iBookmark = tdbg.Bookmark
            sSQL = SQLStoreD02P4504()
            Dim bResult As Boolean = ExecuteSQL(sSQL)
            If bResult = True Then
                DeleteOK()
                DeleteGridEvent(tdbg, dtGrid, gbEnabledUseFind)
                ResetGrid()
                'LoadTDBGrid()
                'If Not IsDBNull(iBookmark) Then tdbg.Bookmark = iBookmark
            Else
                DeleteNotOK()
            End If
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P4504
    '# Created User: 
    '# Created Date: 05/12/2007 03:45:10
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P4504() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P4504 "
        sSQL &= SQLString(tdbg.Columns(COL_AssetID).Text) & COMMA 'AssetID, varchar[20], NOT NULL
        sSQL &= SQLString(tdbg.Columns(COL_BatchID).Text) & COMMA 'BatchID, varchar[20], NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, tinyint, NOT NULL
        sSQL &= SQLNumber(giTranYear) 'TranYear, int, NOT NULL
        Return sSQL
    End Function


    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Danh_sach_nghiep_vu_chuyen_nguon__-_D02F2006") & UnicodeCaption(gbUnicode) 'Danh sÀch nghiÖp vó chuyÓn nguän  - D02F2006
        '================================================================ 
        '================================================================ 
        tdbg.Columns("AssetID").Caption = rl3("Ma_tai_san") 'Mã tài sản
        tdbg.Columns("ChangeNo").Caption = rl3("Ma_nghiep_vu") 'Mã nghiệp vụ 
        tdbg.Columns("ChangeName").Caption = rl3("Ten_nghiep_vu") 'Tên nghiệp vụ
        tdbg.Columns("ChangeDate").Caption = rl3("Ngay") 'Ngày
        tdbg.Columns("DecisionNo").Caption = rl3("So_hieu") 'Số hiệu
        tdbg.Columns("Notes1").Caption = rl3("Ghi_chu") & " 1" ' rl3("Ghi_chu_1") 'Ghi chú 1
        tdbg.Columns("Notes2").Caption = rl3("Ghi_chu") & " 2" 'Ghi chú 2
        tdbg.Columns("Notes3").Caption = rl3("Ghi_chu") & " 3" 'Ghi chú 3

    End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        HotKeyCtrlVOnGrid(tdbg, e) 'Đã bổ sung D99X0000
    End Sub
End Class