
'#-------------------------------------------------------------------------------------
'# Created Date: 06/12/2007 2:12:16 PM
'# Created User: 
'# Modify Date: 06/12/2007 2:12:16 PM
'# Modify User: 
'#-------------------------------------------------------------------------------------
Imports System.Text

Public Class D02F2010

#Region "Const of tdbg1"
    Private Const COL_AssetID As Integer = 0   ' Mã tài sản
    Private Const COL_AssetName As Integer = 1 ' Tên tài sản
#End Region

    Private dtAsset As DataTable
    Dim dtSource As New DataTable
    Dim dtDes As New DataTable
    Private _batchID As String


    Public Property BatchID() As String
        Get
            Return _batchID
        End Get
        Set(ByVal value As String)
            If BatchID = value Then
                _batchID = ""
                Return
            End If
            _batchID = value
        End Set
    End Property

    Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
            _FormState = value
            LoadTDBCombo()
            Select Case _FormState
                Case EnumFormState.FormAdd

            End Select
        End Set
    End Property


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub D02F2010_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        Select Case e.KeyCode
            Case Keys.Enter
                UseEnterAsTab(Me)
            Case Keys.F11
                HotKeyF11(Me, tdbg1)
                'If tdbg1.Focus Then
                '    HotKeyF11(Me, tdbg1)
                'ElseIf tdbg2.Focus Then
                '    HotKeyF11(Me, tdbg2)
                'End If
        End Select
    End Sub

    Private Sub D02F2010_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Loadlanguage()
        'ResetColorGrid(tdbg1)
        'ResetColorGrid(tdbg2)
        CheckStatus()
        InitdtDes()
    SetResolutionForm(Me)
Me.Cursor = Cursors.Default
End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        Dim sSQL As String = ""
        Dim bResult As Boolean
        Try
            sSQL = " Delete From D02T2010 Where UserID=" & SQLString(gsUserID) & vbCrLf
            sSQL &= SQLInsertD02T2010s.ToString
            bResult = ExecuteSQL(sSQL)
            If bResult Then
                Dim f As New D02F2007
                With f
                    .BatchID = ""
                    .ShowDialog()
                    .Dispose()
                    _batchID = .BatchID
                End With
                Me.Close()
            End If
        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try
    End Sub

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        Dim sUnicode As String = UnicodeJoin(gbUnicode)
        'Load tdbcGroupTypeID
        sSQL = "Select GroupTypeID," & IIf(geLanguage = EnumLanguage.Vietnamese, "VieTypeCaption", "EngTypeCaption").ToString & sUnicode & " as GroupTypeCaption , TableName, WhereClause " & vbCrLf
        sSQL &= " From D02V3333 Order By GroupTypeID"

        Dim dtGroupTypeID1 As DataTable = ReturnDataTable(sSQL)
        LoadDataSource(tdbcGroupTypeID1, dtGroupTypeID1, gbUnicode)
        '  Dim dtGroupTypeID2 As DataTable = dtGroupTypeID1.Copy
        LoadDataSource(tdbcGroupTypeID2, dtGroupTypeID1.DefaultView.ToTable, gbUnicode)
        ' Dim dtGroupTypeID3 As DataTable = dtGroupTypeID1.Copy
        LoadDataSource(tdbcGroupTypeID3, dtGroupTypeID1.DefaultView.ToTable, gbUnicode)

        'Load tdbcAssetID
        sSQL = "Select 0 as DisplayOrder,'%' as ACodeID, " & AllName & " as Description, '%' as GroupTypeID " & vbCrLf
        sSQL &= "Union All " & vbCrLf
        sSQL &= "Select 1 as DisplayOrder,ACodeID,Description" & sUnicode & " as Description, TypeCodeID as GroupTypeID From D02V4444 " & vbCrLf
        sSQL &= "Order By DisplayOrder,ACodeID"
        dtAsset = ReturnDataTable(sSQL)

    End Sub

    Private Sub LoadtdbcAssetID(ByVal tdbcFrm As C1.Win.C1List.C1Combo, ByVal tdbcTo As C1.Win.C1List.C1Combo, ByVal ID As String)
        LoadDataSource(tdbcFrm, ReturnTableFilter(dtAsset, "GroupTypeID='%' or GroupTypeID  like " & SQLString(ID)), gbUnicode)
        LoadDataSource(tdbcTo, ReturnTableFilter(dtAsset, " GroupTypeID='%' or GroupTypeID  like " & SQLString(ID)), gbUnicode)
    End Sub

#Region "Events tdbcGroupTypeID1"

    Private Sub tdbcGroupTypeID1_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcGroupTypeID1.Close
        If tdbcGroupTypeID1.FindStringExact(tdbcGroupTypeID1.Text) = -1 Then tdbcGroupTypeID1.Text = ""
    End Sub

    Private Sub tdbcGroupTypeID1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcGroupTypeID1.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcGroupTypeID1.Text = ""
    End Sub

    Private Sub tdbcGroupTypeID1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcGroupTypeID1.SelectedValueChanged
        If Not (tdbcGroupTypeID1.Tag Is Nothing OrElse tdbcGroupTypeID1.Tag.ToString = "") Then
            tdbcGroupTypeID1.Tag = ""
            Exit Sub
        End If
        If tdbcGroupTypeID1.SelectedValue Is Nothing Then
            LoadtdbcAssetID(tdbcFromAssetID1, tdbcToAssetID1, "-1")
            Exit Sub
        End If
        LoadtdbcAssetID(tdbcFromAssetID1, tdbcToAssetID1, tdbcGroupTypeID1.SelectedValue.ToString)
        tdbcFromAssetID1.AutoSelect = True
        tdbcToAssetID1.AutoSelect = True

    End Sub
#End Region

#Region "Events tdbcGroupTypeID2"

    Private Sub tdbcGroupTypeID2_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcGroupTypeID2.Close
        If tdbcGroupTypeID2.FindStringExact(tdbcGroupTypeID2.Text) = -1 Then tdbcGroupTypeID2.Text = ""
    End Sub

    Private Sub tdbcGroupTypeID2_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcGroupTypeID2.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcGroupTypeID2.Text = ""
    End Sub

    Private Sub tdbcGroupTypeID2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcGroupTypeID2.SelectedValueChanged
        If Not (tdbcGroupTypeID2.Tag Is Nothing OrElse tdbcGroupTypeID2.Tag.ToString = "") Then
            tdbcGroupTypeID2.Tag = ""
            Exit Sub
        End If
        If tdbcGroupTypeID2.SelectedValue Is Nothing Then
            LoadtdbcAssetID(tdbcFromAssetID2, tdbcToAssetID2, "-1")
            Exit Sub
        End If
        LoadtdbcAssetID(tdbcFromAssetID2, tdbcToAssetID2, tdbcGroupTypeID2.SelectedValue.ToString)
        tdbcFromAssetID2.AutoSelect = True
        tdbcToAssetID2.AutoSelect = True

    End Sub
#End Region

#Region "Events tdbcGroupTypeID3 load tdbcFromAssetID1"

    Private Sub tdbcGroupTypeID3_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcGroupTypeID3.Close
        If tdbcGroupTypeID3.FindStringExact(tdbcGroupTypeID3.Text) = -1 Then tdbcGroupTypeID3.Text = ""
    End Sub

    Private Sub tdbcGroupTypeID3_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcGroupTypeID3.SelectedValueChanged
        If Not (tdbcGroupTypeID3.Tag Is Nothing OrElse tdbcGroupTypeID3.Tag.ToString = "") Then
            tdbcGroupTypeID3.Tag = ""
            Exit Sub
        End If
        If tdbcGroupTypeID3.SelectedValue Is Nothing Then
            LoadtdbcAssetID(tdbcFromAssetID3, tdbcToAssetID3, "-1")
            Exit Sub
        End If
        LoadtdbcAssetID(tdbcFromAssetID3, tdbcToAssetID3, tdbcGroupTypeID3.SelectedValue.ToString())
        tdbcFromAssetID3.AutoSelect = True
        tdbcToAssetID3.AutoSelect = True
    End Sub

    Private Sub tdbcGroupTypeID3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcGroupTypeID3.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcGroupTypeID3.Text = ""
    End Sub

    Private Sub tdbcFromAssetID1_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcFromAssetID1.Close
        If tdbcFromAssetID1.FindStringExact(tdbcFromAssetID1.Text) = -1 Then tdbcFromAssetID1.Text = ""
    End Sub

    Private Sub tdbcFromAssetID1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcFromAssetID1.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcFromAssetID1.Text = ""
    End Sub

#End Region

#Region "Events tdbcToAssetID1 load tdbcFromAssetID2"

    Private Sub tdbcToAssetID1_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcToAssetID1.Close
        If tdbcToAssetID1.FindStringExact(tdbcToAssetID1.Text) = -1 Then tdbcToAssetID1.Text = ""
    End Sub


    Private Sub tdbcToAssetID1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcToAssetID1.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcToAssetID1.Text = ""
    End Sub

    Private Sub tdbcFromAssetID2_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcFromAssetID2.Close
        If tdbcFromAssetID2.FindStringExact(tdbcFromAssetID2.Text) = -1 Then tdbcFromAssetID2.Text = ""
    End Sub

    Private Sub tdbcFromAssetID2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcFromAssetID2.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcFromAssetID2.Text = ""
    End Sub

#End Region

#Region "Events tdbcToAssetID2 load tdbcFromAssetID3"

    Private Sub tdbcToAssetID2_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcToAssetID2.Close
        If tdbcToAssetID2.FindStringExact(tdbcToAssetID2.Text) = -1 Then tdbcToAssetID2.Text = ""
    End Sub

    Private Sub tdbcToAssetID2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcToAssetID2.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcToAssetID2.Text = ""
    End Sub

    Private Sub tdbcFromAssetID3_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcFromAssetID3.Close
        If tdbcFromAssetID3.FindStringExact(tdbcFromAssetID3.Text) = -1 Then tdbcFromAssetID3.Text = ""
    End Sub

    Private Sub tdbcFromAssetID3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcFromAssetID3.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcFromAssetID3.Text = ""
    End Sub

#End Region

#Region "Events tdbcToAssetID3"

    Private Sub tdbcToAssetID3_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcToAssetID3.Close
        If tdbcToAssetID3.FindStringExact(tdbcToAssetID3.Text) = -1 Then tdbcToAssetID3.Text = ""
    End Sub

    Private Sub tdbcToAssetID3_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcToAssetID3.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcToAssetID3.Text = ""
    End Sub

#End Region
    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P4505
    '# Created User: 
    '# Created Date: 06/12/2006 11:14:24
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P4505() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P4505 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLString(tdbcGroupTypeID1.Text) & COMMA 'SelType1, varchar[20], NOT NULL
        sSQL &= SQLString(tdbcFromAssetID1.Text) & COMMA 'SelFrom1, varchar[20], NOT NULL
        sSQL &= SQLString(tdbcToAssetID1.Text) & COMMA 'SelTo1, varchar[20], NOT NULL
        sSQL &= SQLString(tdbcGroupTypeID2.Text) & COMMA 'SelType2, varchar[20], NOT NULL
        sSQL &= SQLString(tdbcFromAssetID2.Text) & COMMA 'SelFrom2, varchar[20], NOT NULL
        sSQL &= SQLString(tdbcToAssetID2.Text) & COMMA 'SelTo2, varchar[20], NOT NULL
        sSQL &= SQLString(tdbcGroupTypeID3.Text) & COMMA 'SelType3, varchar[20], NOT NULL
        sSQL &= SQLString(tdbcFromAssetID3.Text) & COMMA 'SelFrom3, varchar[20], NOT NULL
        sSQL &= SQLString(tdbcToAssetID3.Text) 'SelTo3, varchar[20], NOT NULL
        sSQL &= COMMA & SQLNumber(gbUnicode)
        Return sSQL
    End Function

    Private Sub btnShow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim sSQL As String = SQLStoreD02P4505()
        dtSource = ReturnDataTable(sSQL)
        Dim keys(0) As DataColumn
        keys(0) = dtSource.Columns("AssetID")
        dtSource.PrimaryKey = keys
        LoadDataSource(tdbg1, dtSource, gbUnicode)
        CheckStatus()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Dim bookmark As Integer
        Dim iCount As Integer = 0
        iCount = dtDes.Rows.Count
        If tdbg1.Bookmark < 0 Then Exit Sub

        bookmark = tdbg1.Bookmark
        If bookmark > dtSource.Rows.Count Then bookmark = dtSource.Rows.Count
        Dim aSelectRows As C1.Win.C1TrueDBGrid.SelectedRowCollection
        aSelectRows = tdbg1.SelectedRows
        If aSelectRows.Count > 0 Then 'chọn nhiều dòng
            Dim i As Integer
            'Thêm dữ liệu lưới phải
            For i = 0 To aSelectRows.Count - 1
                CopyItem(aSelectRows.Item(i), tdbg1, dtSource, tdbg2, dtDes)
                iCount += 1
            Next
            'Xoá dl lưới trái
            i = aSelectRows.Count - 1
            While (i >= 0)
                RemoveItem(aSelectRows.Item(i), tdbg1, dtSource, COL_AssetID)
                i -= 1
            End While
        Else 'dòng có con trỏ, không chọn
            CopyItem(tdbg1.Row, tdbg1, dtSource, tdbg2, dtDes)
            RemoveItem(tdbg1.Row, tdbg1, dtSource, COL_AssetID)
        End If
        'LoadDataSource(tdbg1, dtSource)
        tdbg1.Bookmark = bookmark
        LoadDataSource(tdbg2, dtDes, gbUnicode)
        CheckStatus()

    End Sub

    Private Sub btnAddMulti_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddMulti.Click
        Dim i As Integer = tdbg1.RowCount - 1
        While i >= 0
            CopyItem(i, tdbg1, dtSource, tdbg2, dtDes)
            RemoveItem(i, tdbg1, dtSource, COL_AssetID)
            i = tdbg1.RowCount - 1
        End While
        LoadDataSource(tdbg1, dtSource, gbUnicode)
        LoadDataSource(tdbg2, dtDes, gbUnicode)
        CheckStatus()
    End Sub

    Private Sub btnRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        If tdbg2.Bookmark < 0 Then Exit Sub
        Dim iBookmark As Integer = tdbg2.Bookmark
        If iBookmark > dtDes.Rows.Count Then iBookmark = dtDes.Rows.Count
        Dim aSelectRows As C1.Win.C1TrueDBGrid.SelectedRowCollection
        aSelectRows = tdbg2.SelectedRows
        If aSelectRows.Count > 0 Then 'chọn nhiều dòng
            Dim i As Integer
            'Thêm dữ liệu lưới phải
            For i = 0 To aSelectRows.Count - 1
                CopyItem(i, tdbg2, dtDes, tdbg1, dtSource)
            Next
            'Xoá dl lưới trái
            i = aSelectRows.Count - 1
            While (i >= 0)
                RemoveItem(i, tdbg2, dtDes, COL_AssetID)
                i -= 1
            End While
           
        Else 'dòng có con trỏ, không chọn
            CopyItem(tdbg2.Row, tdbg2, dtDes, tdbg1, dtSource)
            RemoveItem(tdbg2.Row, tdbg2, dtDes, COL_AssetID)

        End If
        LoadDataSource(tdbg1, dtSource, gbUnicode)
        'LoadDataSource(tdbg2, dtDes)
        tdbg2.Bookmark = iBookmark
        CheckStatus()
    End Sub

    Private Sub btnRemoveAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRemoveAll.Click
        Dim i As Integer = tdbg2.RowCount - 1
        While i >= 0
            CopyItem(i, tdbg2, dtDes, tdbg1, dtSource)
            RemoveItem(i, tdbg2, dtDes, COL_AssetID)
            i = tdbg2.RowCount - 1
        End While
        LoadDataSource(tdbg1, dtSource, gbUnicode)
        LoadDataSource(tdbg2, dtDes, gbUnicode)
        CheckStatus()
    End Sub

    Private Sub CheckStatus()
        btnAdd.Enabled = dtSource.Rows.Count > 0
        btnAddMulti.Enabled = dtSource.Rows.Count > 0
        btnRemove.Enabled = dtDes.Rows.Count > 0
        btnRemoveAll.Enabled = dtDes.Rows.Count > 0
        btnNext.Enabled = dtDes.Rows.Count > 0
    End Sub

    Private Sub CopyItem(ByVal iRow As Integer, ByVal tdbgSource As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dtSource As DataTable, ByVal tdbgDes As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByRef dtDes As DataTable)
        Dim Row As DataRow
        Row = dtDes.NewRow
        For i As Integer = 0 To tdbgSource.Columns.Count - 1
            Row(i) = tdbgSource(iRow, i)
        Next
        dtDes.Rows.Add(Row)
    End Sub

    Private Sub RemoveItem(ByVal iRow As Integer, ByVal tdbgSource As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByRef dtSource As DataTable, Optional ByVal iColPkey As Integer = 0)
        Dim myDataRow As DataRowCollection = dtSource.Rows
        If myDataRow.Contains(tdbgSource(iRow, iColPkey)) Then
            Dim row As DataRow = myDataRow.Find(tdbgSource(iRow, iColPkey))
            myDataRow.Remove(row)
        End If
    End Sub

    Private Sub InitdtDes()
        Dim sSQL As String = ""
        sSQL = " Select AssetID, AssetName "
        sSQL &= "FROM D02T0001 WITH(NOLOCK)" & vbCrLf
        sSQL &= "WHERE 0=1"
        dtDes = ReturnDataTable(sSQL)
        'Create Primary key
        Dim keys(0) As DataColumn
        keys(0) = dtDes.Columns("AssetID")
        dtDes.PrimaryKey = keys

    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T2010s
    '# Created User: 
    '# Created Date: 06/12/2007 02:11:52
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T2010s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        For i As Integer = 0 To tdbg2.RowCount - 1
            sSQL.Append("Insert Into D02T2010(")
            sSQL.Append("UserID, AssetID, AssetName")
            sSQL.Append(") Values(")
            sSQL.Append(SQLString(gsUserID) & COMMA) 'UserID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg2(i, COL_AssetID)) & COMMA) 'AssetID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg2(i, COL_AssetName))) 'AssetName, varchar[250], NULL
            sSQL.Append(")")
            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Chon_nhieu_tai_san_-_D02F2010") & UnicodeCaption(gbUnicode) 'Chãn nhiÒu tªi s¶n - D02F2010
        '================================================================ 
        lblGroupTypeID1.Text = rl3("Phan_loai") & " 1"  'Phân loại 1
        lblGroupTypeID2.Text = rl3("Phan_loai") & " 2" 'Phân loại 2
        lblGroupTypeID3.Text = rl3("Phan_loai") & " 3" 'Phân loại 3
        lblFromAssetID1.Text = rl3("Tu") 'Từ
        lblToAssetID1.Text = rl3("Den") 'Đến
        '================================================================ 
        btnShow.Text = "&" & rl3("Hien_thi") '&Hiển thị
        btnNext.Text = rl3("_Tiep_tuc") '&Tiếp tục
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        '================================================================ 
        grp1.Text = rl3("Ma_TSCD") 'Mã TSCĐ
        grp2.Text = rl3("Ma_TSCD_duoc_chon") 'Mã TSCĐ được chọn
        '================================================================ 
        tdbcGroupTypeID1.Columns("GroupTypeID").Caption = rl3("Ma") 'Mã
        tdbcGroupTypeID1.Columns("GroupTypeCaption").Caption = rl3("Ten") 'Tên
        tdbcGroupTypeID2.Columns("GroupTypeID").Caption = rl3("Ma") 'Mã
        tdbcGroupTypeID2.Columns("GroupTypeCaption").Caption = rl3("Ten") 'Tên
        tdbcGroupTypeID3.Columns("GroupTypeID").Caption = rl3("Ma") 'Mã
        tdbcGroupTypeID3.Columns("GroupTypeCaption").Caption = rl3("Ten") 'Tên
        tdbcFromAssetID1.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcFromAssetID1.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcToAssetID1.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcToAssetID1.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcFromAssetID2.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcFromAssetID2.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcToAssetID2.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcToAssetID2.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcFromAssetID3.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcFromAssetID3.Columns("Description").Caption = rl3("Ten") 'Tên
        tdbcToAssetID3.Columns("ACodeID").Caption = rl3("Ma") 'Mã
        tdbcToAssetID3.Columns("Description").Caption = rl3("Ten") 'Tên
        '================================================================ 
        tdbg1.Columns("AssetID").Caption = rl3("Ma_tai_san") 'Mã tài sản
        tdbg1.Columns("AssetName").Caption = rl3("Ten_tai_san") 'Tên tài sản
        tdbg2.Columns("AssetID").Caption = rl3("Ma_tai_san") 'Mã tài sản
        tdbg2.Columns("AssetName").Caption = rl3("Ten_tai_san") 'Tên tài sản
    End Sub

    Private Sub tdbg1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg1.DoubleClick
        btnAdd_Click(sender, Nothing)
    End Sub

    Private Sub tdbg2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbg2.DoubleClick
        btnRemove_Click(sender, Nothing)
    End Sub
End Class