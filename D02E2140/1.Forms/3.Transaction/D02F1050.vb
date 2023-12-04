'#-------------------------------------------------------------------------------------
'# Created Date: 21/11/2007 3:14:05 PM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 21/11/2007 3:14:05 PM
'# Modify User: Trần Thị ÁiTrâm
'#-------------------------------------------------------------------------------------
Imports System.Text

Public Class D02F1050

#Region "Const of tdbg"
    Private Const COL_AssetID As Integer = 0   ' Mã tài sản
    Private Const COL_AssetName As Integer = 1 ' Tên tài sản
    Private Const COL_IndexName As Integer = 2 ' Chỉ số
#End Region

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""

        'Load tdbcAssetS1ID
        sSQL = "Select 0 as DisplayOrder,'%' as AssetS1ID, " & AllName & " As AssetS1Name " & vbCrLf
        sSQL &= "Union" & vbCrLf
        sSQL &= "Select 1 as DisplayOrder,AssetS1ID, AssetS1Name" & UnicodeJoin(gbUnicode) & " as AssetS1Name From D02T1000 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where Disabled = 0 Order By DisplayOrder,AssetS1ID"
        LoadDataSource(tdbcAssetS1ID, sSQL, gbUnicode)

        'Load tdbcAssetS2ID
        sSQL = "Select 0 as DisplayOrder,'%' as AssetS2ID, " & AllName & " As AssetS2Name " & vbCrLf
        sSQL &= "Union" & vbCrLf
        sSQL &= "Select 1 as DisplayOrder,AssetS2ID, AssetS2Name" & UnicodeJoin(gbUnicode) & " as AssetS2Name From D02T2000 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where Disabled = 0 Order By DisplayOrder,AssetS2ID"
        LoadDataSource(tdbcAssetS2ID, sSQL, gbUnicode)

        'Load tdbcAssetS3ID
        sSQL = "Select  0 as DisplayOrder,'%' as AssetS3ID, " & AllName & " As AssetS3Name " & vbCrLf
        sSQL &= "Union" & vbCrLf
        sSQL &= "Select 1 as DisplayOrder,AssetS3ID, AssetS3Name" & UnicodeJoin(gbUnicode) & " as AssetS3Name From D02T3000 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where Disabled = 0 Order By DisplayOrder,AssetS3ID"
        LoadDataSource(tdbcAssetS3ID, sSQL, gbUnicode)
        'Load tdbcID
        sSQL = "Select ID," & IIf(geLanguage = EnumLanguage.Vietnamese, "VieCaption" & UnicodeJoin(gbUnicode), " EngCaption" & UnicodeJoin(gbUnicode)).ToString & " As Caption, FieldName " & vbCrLf
        sSQL &= "From D02T0039 WITH(NOLOCK) where Disabled = 0 Order By ID" & vbCrLf
        LoadDataSource(tdbcID, sSQL, gbUnicode)
    End Sub

#Region "Events tdbcAssetS1ID with txtAssetS1Name"

    Private Sub tdbcAssetS1ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetS1ID.Close
        If tdbcAssetS1ID.FindStringExact(tdbcAssetS1ID.Text) = -1 Then
            tdbcAssetS1ID.Text = ""
            txtAssetS1Name.Text = ""
        End If
    End Sub

    Private Sub tdbcAssetS1ID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetS1ID.SelectedValueChanged
        txtAssetS1Name.Text = tdbcAssetS1ID.Columns(1).Value.ToString
    End Sub

    'Private Sub tdbcAssetS1ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAssetS1ID.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '        tdbcAssetS1ID.Text = ""
    '        txtAssetS1Name.Text = ""
    '    End If
    'End Sub

#End Region

#Region "Events tdbcAssetS2ID with txtAssetS2Name"

    Private Sub tdbcAssetS2ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetS2ID.Close
        If tdbcAssetS2ID.FindStringExact(tdbcAssetS2ID.Text) = -1 Then
            tdbcAssetS2ID.Text = ""
            txtAssetS2Name.Text = ""
        End If
    End Sub

    Private Sub tdbcAssetS2ID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetS2ID.SelectedValueChanged
        txtAssetS2Name.Text = tdbcAssetS2ID.Columns(1).Value.ToString
    End Sub

    'Private Sub tdbcAssetS2ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAssetS2ID.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '        tdbcAssetS2ID.Text = ""
    '        txtAssetS2Name.Text = ""
    '    End If
    'End Sub

#End Region

#Region "Events tdbcAssetS3ID with txtAssetS3Name"

    Private Sub tdbcAssetS3ID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetS3ID.Close
        If tdbcAssetS3ID.FindStringExact(tdbcAssetS3ID.Text) = -1 Then
            tdbcAssetS3ID.Text = ""
            txtAssetS3Name.Text = ""
        End If
    End Sub

    Private Sub tdbcAssetS3ID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcAssetS3ID.SelectedValueChanged
        txtAssetS3Name.Text = tdbcAssetS3ID.Columns(1).Value.ToString
    End Sub

    'Private Sub tdbcAssetS3ID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcAssetS3ID.KeyDown
    '    'If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '    '    tdbcAssetS3ID.Text = ""
    '    '    txtAssetS3Name.Text = ""
    '    End If
    'End Sub

#End Region

#Region "Events tdbcID"

    Private Sub tdbcID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcID.Close
        If tdbcID.FindStringExact(tdbcID.Text) = -1 Then tdbcID.Text = ""
    End Sub

    'Private Sub tdbcID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcID.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcID.Text = ""
    'End Sub


    Private Sub tdbcID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcID.SelectedValueChanged
        LoadTDBGrid()
        tdbg.Columns(COL_IndexName).Caption = tdbcID.Columns("Caption").Text
        tdbg.Splits(0).DisplayColumns(COL_IndexName).HeadingStyle.Font = FontUnicode(gbUnicode)
    End Sub
#End Region

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub D02F1050_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                UseEnterAsTab(Me)
            Case Keys.F11
                HotKeyF11(Me, tdbg)
        End Select
    End Sub

    Private Sub D02F1050_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InputbyUnicode(Me, gbUnicode)
        Loadlanguage()
        LoadTDBCombo()
        EnablebControl()
        txtPeriod.Text = giTranMonth.ToString("00") & "/" & giTranYear.ToString
        ExecuteSQLNoTransaction(" exec D02P1051 " & SQLNumber(giTranMonth) & ", " & SQLNumber(giTranYear))     'Tra ra DL
        tdbg_NumberFormat()
    SetResolutionForm(Me)
Me.Cursor = Cursors.Default
End Sub

    Private Sub EnablebControl()
        Dim sSQL As String
        sSQL = "Select AssetS1Enabled, AssetS2Enabled, AssetS3Enabled" & vbCrLf
        sSQL &= "From D02T0000 WITH(NOLOCK)" & vbCrLf
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            With dt.Rows(0)
                tdbcAssetS1ID.Enabled = CBool(.Item("AssetS1Enabled"))
                tdbcAssetS2ID.Enabled = CBool(.Item("AssetS2Enabled"))
                tdbcAssetS3ID.Enabled = CBool(.Item("AssetS3Enabled"))
            End With
        End If
    End Sub


    Private Sub LoadTDBGrid()

        Dim sSQL As String = SQLStoreD02P1050()
        LoadDataSource(tdbg, sSQL, gbUnicode)
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1050
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 21/11/2007 03:54:57
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1050() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P1050 "
        sSQL &= SQLString(tdbcAssetS1ID.Text) & COMMA 'AssetS1ID, varchar[20], NOT NULL
        sSQL &= SQLString(tdbcAssetS2ID.Text) & COMMA 'AssetS2ID, varchar[20], NOT NULL
        sSQL &= SQLString(tdbcAssetS3ID.Text) & COMMA 'AssetS3ID, varchar[20], NOT NULL
        sSQL &= SQLNumber(tdbcID.Columns("ID").Value.ToString) & COMMA 'IndexID, tinyint, NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, tinyint, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, int, NOT NULL
        Return sSQL
    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        tdbg.UpdateData()

        If Not AllowSave() Then Exit Sub
        btnSave.Enabled = False
        btnClose.Enabled = False

        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        sSQL.Append(SQLUpdateD02T1050s)

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
            btnClose.Focus()
        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T1050s
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 21/11/2007 04:03:53
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T1050s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        For i As Integer = 0 To tdbg.RowCount - 1
            sSQL.Append("Update D02T1050 Set ")
            sSQL.Append(tdbcID.Columns("FieldName").Value.ToString & " = " & SQLMoney(tdbg(i, COL_IndexName))) 'money, NOT NULL
            sSQL.Append(" Where ")
            sSQL.Append("AssetID = " & SQLString(tdbg(i, COL_AssetID)) & " And ")
            sSQL.Append("TranMonth = " & SQLNumber(giTranMonth) & " And ")
            sSQL.Append("TranYear = " & SQLNumber(giTranYear))
            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        If e.KeyCode = Keys.Enter Then
            If tdbg.Col = COL_IndexName Then
                HotKeyEnterGrid(tdbg, COL_IndexName, e)
            End If
        End If
        If e.KeyCode = Keys.F7 Then
            HotKeyF7(tdbg)
        End If
    End Sub

    Private Sub tdbg_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        Select Case tdbg.Col
            Case COL_IndexName
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub

    Private Sub tdbg_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.AfterColUpdate
        Select Case e.ColIndex
            Case COL_IndexName
                tdbg.Columns(COL_IndexName).Text = SQLNumber(tdbg.Columns(COL_IndexName).Text, DxxFormat.DefaultNumber4)
        End Select
    End Sub

    Private Sub tdbg_NumberFormat()
        tdbg.Columns(COL_IndexName).NumberFormat = DxxFormat.DefaultNumber4
    End Sub

    Private Function AllowSave() As Boolean
        If tdbg.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tdbg.Focus()
            Return False
        End If
        For i As Integer = 0 To tdbg.RowCount - 1
            If tdbg(i, COL_IndexName).ToString <> "" Then
                If CDbl(tdbg(i, COL_IndexName)) > MaxMoney Then
                    D99C0008.MsgL3(rl3("So_qua_lon"))
                    tdbg.SplitIndex = SPLIT0
                    tdbg.Col = COL_IndexName
                    tdbg.Bookmark = i
                    tdbg.Focus()
                    Return False
                End If
            End If
        Next
        Return True
    End Function

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Cap_nhat_chi_so_-_D02F1050") & UnicodeCaption(gbUnicode) 'CËp nhËt chÙ sç - D02F1050
        '================================================================ 
        lblAssetS1ID.Text = rl3("Phan_loai") & " 1" 'Phân loại 1
        lblID.Text = rl3("Chi_so") 'Chỉ số
        lblPeriod.Text = rl3("Ky") 'Kỳ
        lblAssetS2ID.Text = rl3("Phan_loai") & " 2" 'Phân loại 2
        lblAssetS3ID.Text = rl3("Phan_loai") & " 3" 'Phân loại 3
        '================================================================ 
        btnSave.Text = rl3("_Luu") '&Lưu
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        '================================================================ 
        tdbcAssetS1ID.Columns("AssetS1ID").Caption = rl3("Ma") 'Mã
        tdbcAssetS1ID.Columns("AssetS1Name").Caption = rl3("Ten") 'Tên
        tdbcID.Columns("ID").Caption = rl3("Ma") 'Mã
        tdbcID.Columns("Caption").Caption = rl3("Ten") 'Tên
        tdbcAssetS2ID.Columns("AssetS2ID").Caption = rl3("Ma") 'Mã
        tdbcAssetS2ID.Columns("AssetS2Name").Caption = rl3("Ten") 'Tên
        tdbcAssetS3ID.Columns("AssetS3ID").Caption = rl3("Ma") 'Mã
        tdbcAssetS3ID.Columns("AssetS3Name").Caption = rl3("Ten") 'Tên
        '================================================================ 
        tdbg.Columns("AssetID").Caption = rl3("Ma_tai_san") 'Mã tài sản
        tdbg.Columns("AssetName").Caption = rl3("Ten_tai_san") 'Tên tài sản
        tdbg.Columns("IndexName").Caption = rl3("Chi_soV") 'rl3("Chi_so") 'Chỉ số 
    End Sub
End Class