'#-------------------------------------------------------------------------------------
'# Created Date: 06/12/2007 2:12:36 PM
'# Created User: 
'# Modify Date: 06/12/2007 2:12:36 PM
'# Modify User: 
'#-------------------------------------------------------------------------------------
Imports System.Text

Public Class D02F2007

#Region "Const of tdbg1"
    Private Const COL1_AssetID As Integer = 0   ' Mã tài sản
    Private Const COL1_AssetName As Integer = 1 ' Tên tài sản
    Private Const COL1_AccountID As Integer = 2 ' Tài khoản
    Private Const COL1_Total As Integer = 3     ' Tổng cộng
#End Region

#Region "Const of tdbg2"
    Private Const COL2_AssetID As Integer = 0   ' Mã tài sản
    Private Const COL2_AssetName As Integer = 1 ' Tên tài sản
    Private Const COL2_AccountID As Integer = 2 ' Tài khoản
    Private Const COL2_Total As Integer = 3     ' Tổng cộng
#End Region

    Private dtMain As DataTable
    Private dtGrid1 As DataTable
    Private dtGrid2 As DataTable
    Dim nTotalSourceID As Integer ' Tổng số SourceID
    Dim iCountCol As Int32 ' Số cột trên lưới 2
    Dim iLastCol1 As Integer
    Dim iLastCol2 As Integer

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

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub D02F2007_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                UseEnterAsTab(Me)
            Case Keys.F11
                HotKeyF11(Me, tdbg2)
        End Select
    End Sub

    Private Sub SetBackColorObligatory()
        tdbcChangeNo.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private Sub D02F2007_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Cursor = Cursors.WaitCursor
        Loadlanguage()
        SetBackColorObligatory()
        c1dateChangeDate.Value = Date.Today
        LoadTDBCombo()
        InputbyUnicode(Me, gbUnicode)
        DataFill()
        AddFieldTDBGrid1()
        AddFieldTDBGrid2()
        tdbg1_LockedColumns()
        tdbg2_LockedColumns()
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

#Region "Events tdbcChangeNo with txtChangeNoName"

    Private Sub tdbcChangeNo_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcChangeNo.Close
        If tdbcChangeNo.FindStringExact(tdbcChangeNo.Text) = -1 Then
            tdbcChangeNo.Text = ""
            txtChangeNoName.Text = ""
            txtNotes.Text = ""
            txtNotes2.Text = ""
            txtNotes3.Text = ""
        End If
    End Sub

    Private Sub tdbcChangeNo_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcChangeNo.SelectedValueChanged
        txtChangeNoName.Text = tdbcChangeNo.Columns(1).Value.ToString
        txtNotes.Text = tdbcChangeNo.Columns("Notes1").Value.ToString
        txtNotes2.Text = tdbcChangeNo.Columns("Notes2").Value.ToString
        txtNotes3.Text = tdbcChangeNo.Columns("Notes3").Value.ToString
    End Sub

    Private Sub tdbcChangeNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcChangeNo.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            tdbcChangeNo.Text = ""
            txtChangeNoName.Text = ""
            txtNotes.Text = ""
            txtNotes2.Text = ""
            txtNotes3.Text = ""
        End If
    End Sub

#End Region

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        Dim sUnicode As String = UnicodeJoin(gbUnicode)
        'Load tdbcChangeNo
        sSQL = "Select ChangeNo,ChangeName" & sUnicode & " as ChangeName, Notes1" & sUnicode & " as Notes1, Notes2" & sUnicode & " as Notes2, Notes3" & sUnicode & " as Notes3 From D02T0201 WITH(NOLOCK) Where Disabled=0 "
        LoadDataSource(tdbcChangeNo, sSQL, gbUnicode)
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P4501
    '# Created User: 
    '# Created Date: 06/12/2007 02:36:05
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P4501() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D02P4501 "
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[20], NOT NULL
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[20], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) 'TranYear, int, NOT NULL
        sSQL &= COMMA & SQLNumber(gbUnicode)
        Return sSQL
    End Function

    Private Sub DataFill()
        Dim sSQL As String = SQLStoreD02P4501()
        dtMain = ReturnDataTable(sSQL)
    End Sub

    Private Sub LoadTDBGrid(ByVal tdbg As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dt As DataTable)
        LoadDataSource(tdbg, dt, gbUnicode)
    End Sub

    Private Sub tdbg1_LockedColumns()
        tdbg1.Splits(SPLIT0).DisplayColumns(COL1_AssetID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT0).DisplayColumns(COL1_AssetName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT0).DisplayColumns(COL1_AccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg1.Splits(SPLIT0).DisplayColumns(COL1_Total).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
    End Sub

    Private Sub tdbg2_LockedColumns()
        tdbg2.Splits(SPLIT0).DisplayColumns(COL2_AssetID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(SPLIT0).DisplayColumns(COL2_AssetName).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(SPLIT0).DisplayColumns(COL2_AccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(SPLIT0).DisplayColumns(COL2_Total).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
    End Sub

    'Private Sub tdbg1_NumberFormat()
    '    tdbg1.Columns(COL1_Total).NumberFormat = DxxFormat.D90_ConvertedDecimals
    'End Sub

    Private Sub tdbg1_NumberFormat()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbg1.Columns(COL1_Total).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        InputNumber(tdbg1, arr)
    End Sub



    'Private Sub tdbg2_NumberFormat()
    '    tdbg2.Columns(COL2_Total).NumberFormat = DxxFormat.D90_ConvertedDecimals
    'End Sub

    Private Sub tdbg2_NumberFormat()
        Dim arr() As FormatColumn = Nothing
        AddDecimalColumns(arr, tdbg2.Columns(COL2_Total).DataField, DxxFormat.D90_ConvertedDecimals, 28, 8)
        InputNumber(tdbg2, arr)
    End Sub



    Private Sub AddFieldTDBGrid1()

        Dim i, j As Integer
        Dim col As C1.Win.C1TrueDBGrid.C1DataColumn
        dtGrid1 = dtMain.Copy

        nTotalSourceID = dtGrid1.Columns.Count - (COL1_Total + 2)

        With tdbg1
            .InsertHorizontalSplit(1)
            .Splits(1).CaptionHeight = 34
            .Splits(1).CaptionStyle.Font = FontUnicode(True)
            .SplitDividerSize = New Size(0, 0)
            .Splits(1).RecordSelectors = False
            .Splits(1).BorderStyle = Border3DStyle.Flat

            'Visible  cac cot tinh o split1 
            For j = 0 To COL1_Total
                .Splits(1).DisplayColumns(j).Visible = False
            Next j
        End With

        For i = 1 To nTotalSourceID
            col = New C1.Win.C1TrueDBGrid.C1DataColumn
            col.DataField = dtGrid1.Columns(4 + i).Caption.ToString
            col.Caption = dtGrid1.Columns(4 + i).Caption.ToString '& vbCrLf
            tdbg1.Columns.Add(col)
            tdbg1.Columns(col.DataField).NumberFormat = DxxFormat.D90_ConvertedDecimals

            col.DataWidth = 18
            tdbg1.Splits(1).DisplayColumns(COL1_Total + i).Visible = True
            tdbg1.Splits(1).DisplayColumns(COL1_Total + i).Width = 140 'nWidth
            tdbg1.Splits(1).DisplayColumns(COL1_Total + i).HeadingStyle.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
            tdbg1.Splits(1).DisplayColumns(COL1_Total + i).Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
            tdbg1.Splits(1).DisplayColumns(COL1_Total + i).Locked = True
        Next i
        ResetColorGrid(tdbg1, 0, 0)
        LoadTDBGrid(tdbg1, dtGrid1)
        tdbg1_NumberFormat()
        iLastCol1 = CountCol(tdbg1, SPLIT1)
    End Sub

    Private Sub AddFieldTDBGrid2()
        Dim i, j As Integer
        Dim col As C1.Win.C1TrueDBGrid.C1DataColumn

        dtGrid2 = dtMain.Copy

        nTotalSourceID = dtGrid2.Columns.Count - (COL1_Total + 2)
        
        'Add cột Tỷ trọng cho Lưới 2
        col = New C1.Win.C1TrueDBGrid.C1DataColumn
        dtGrid2.Columns.Add("Density", System.Type.GetType("System.Double"))
        col.DataField = "Density"
        col.Caption = rl3("Ty_trong") & vbCrLf
        tdbg2.Columns.Add(col)
        tdbg2.Columns(col.DataField).NumberFormat = "Percent"

        'tdbg2.Columns(col.DataField).NumberFormat = D02Format.ConvertedAmount
        tdbg2.Splits(0).DisplayColumns("Density").Visible = True
        tdbg2.Splits(0).DisplayColumns("Density").HeadingStyle.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
        tdbg2.Splits(0).DisplayColumns("Density").Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
        tdbg2.Splits(0).DisplayColumns("Density").Locked = True
        tdbg2.Splits(SPLIT0).DisplayColumns("Density").Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg2.Splits(SPLIT0).DisplayColumns("Density").Width = 65
        With tdbg2
            .InsertHorizontalSplit(1)
            .Splits(1).SplitSize = 1
            .Splits(1).CaptionHeight = 34
            .Splits(1).CaptionStyle.Font = FontUnicode(True)
            .SplitDividerSize = New Size(0, 0)
            .Splits(1).RecordSelectors = False
            .Splits(1).BorderStyle = Border3DStyle.Flat

            'Visible  cac cot tinh o split1 
            For j = 0 To COL2_Total + 1
                .Splits(1).DisplayColumns(j).Visible = False
            Next j
        End With

        iCountCol = tdbg2.Columns.Count

        For i = 1 To nTotalSourceID
            'Add cột Tỷ trọng thứ i
            col = New C1.Win.C1TrueDBGrid.C1DataColumn
            If i < 10 Then
                dtGrid2.Columns.Add("Density" & "0" & i.ToString, System.Type.GetType("System.Double"))
                col.DataField = "Density" & "0" & i.ToString
            Else
                dtGrid2.Columns.Add("Density" & i.ToString, System.Type.GetType("System.Double"))
                col.DataField = "Density" & i.ToString
            End If
            col.Caption = rl3("Ty_trong") & vbCrLf & "[" & dtGrid2.Columns(4 + i).Caption.ToString & "]"
            tdbg2.Columns.Add(col)
            tdbg2.Splits(1).DisplayColumns(col.DataField).Width = 65
            tdbg2.Columns(col.DataField).NumberFormat = "Percent"
            'tdbg2.Columns(col.DataField).NumberFormat = D02Format.Percentage

            'Add cột Tổng cộng thứ i
            col = New C1.Win.C1TrueDBGrid.C1DataColumn
            If i < 10 Then
                dtGrid2.Columns.Add("ConvertedAmount" & "0" & i.ToString, System.Type.GetType("System.Double"))
                col.DataField = "ConvertedAmount" & "0" & i.ToString
            Else
                dtGrid2.Columns.Add("ConvertedAmount" & i.ToString, System.Type.GetType("System.Double"))
                col.DataField = "ConvertedAmount" & i.ToString
            End If

            col.Caption = rl3("Tong_cong") & vbCrLf & "[" & dtGrid2.Columns(4 + i).Caption.ToString & "]"
            tdbg2.Columns.Add(col)
            tdbg2.Columns(col.DataField).NumberFormat = DxxFormat.D90_ConvertedDecimals
            tdbg2.Splits(1).DisplayColumns(col.DataField).Width = 140
            col.DataWidth = 24
            iCountCol += 2
        Next i
        For k As Integer = (COL2_Total + 2) To iCountCol - 1
            tdbg2.Splits(1).DisplayColumns(k).Visible = True
            'tdbg2.Splits(1).DisplayColumns(k).Width = nWidth
            tdbg2.Splits(1).DisplayColumns(k).HeadingStyle.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Center
            tdbg2.Splits(1).DisplayColumns(k).Style.HorizontalAlignment = C1.Win.C1TrueDBGrid.AlignHorzEnum.Far
        Next
        ResetColorGrid(tdbg2, 0, 0)
        LoadTDBGrid(tdbg2, dtGrid2)
        tdbg2_NumberFormat()
        iLastCol2 = CountCol(tdbg2, SPLIT1)
    End Sub

    Private Sub tdbg2_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg2.AfterColUpdate
        Dim iBalance As Integer = 0
        Dim iResult As Integer = 0

        Dim iBalance1 As Integer = 0
        Dim iResult1 As Integer = 0

        Dim dSum As Double = 0
        Dim dx As Integer = 0

        tdbg2.UpdateData()
        iResult = Math.DivRem(e.ColIndex, 2, iBalance)

        If e.ColIndex > 4 And iBalance = 1 Then ' iBalance = 1: là cột tỷ trọng
            tdbg2.Columns(e.ColIndex).Text = (Number(tdbg2.Columns(e.ColIndex).Value) / 100).ToString
            If tdbg2.Columns(e.ColIndex).Text = "" Then
                tdbg2.Columns(e.ColIndex + 1).Text = ""
            Else
                tdbg2.Columns(e.ColIndex + 1).Text = SQLNumber(((Number(tdbg2.Columns(e.ColIndex).Text) * Number(tdbg2.Columns(COL2_Total).Text)) / 100).ToString, DxxFormat.D90_ConvertedDecimals)
            End If
            Try
                For i As Integer = 5 To iCountCol - 1
                    dx = GetDx(i)
                    If dx = 0 Or dx = 2 Then
                        dSum = dSum + Number(tdbg2.Columns(i).Text)
                    End If
                Next
                tdbg2.Columns("Density").Text = (dSum / 100).ToString
            Catch ex As Exception
                D99C0008.Msg(ex.Message)
            End Try
        ElseIf e.ColIndex > 4 And iBalance = 0 Then
            If tdbg2.Columns(e.ColIndex - 1).Text = "" Then
                tdbg2.Columns(e.ColIndex).Text = ""
            End If
        End If
    End Sub

    'Lấy khoảng cách cộng cột
    Private Function GetDx(ByVal iCol As Integer) As Integer
        Dim iBalance As Integer = 0
        Dim iResult As Integer = 0

        Dim dx As Integer = 0
        iResult = Math.DivRem(iCol, 2, iBalance)
        If iCol = 5 Then
            dx = 0 ' cột Tỷ trọng đầu tiên 
        ElseIf iBalance = 0 Then ' Nếu là cột chẵn => Là cột tổng cộng
            dx = 1
        Else 'Nếu là cột lẻ => Là cột tỷ trọng
            dx = 2
        End If
        Return dx
    End Function


    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        tdbg2.UpdateData()
        If Not AllowSave() Then Exit Sub

        'Kiểm tra Ngày phiếu có phù hợp với kỳ kế toán hiện tại không (gọi hàm CheckVoucherDateInPeriod)
        btnSave.Enabled = False
        btnClose.Enabled = False
        gbSavedOK = False
        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder

        sSQL.Append(SQLInsertD02T0202)
        sSQL.Append(vbCrLf)
        sSQL.Append(SQLInsertD02T0012s_tdbg1)
        sSQL.Append(vbCrLf)
        sSQL.Append(SQLInsertD02T0012s_tdbg2)

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            gbSavedOK = True
            btnClose.Enabled = True
            btnSave.Enabled = False
            btnClose.Focus()
        Else
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
     
    End Sub

    Private Function AllowSave() As Boolean
        Dim dx As Integer = 0
        Dim iColMax As Integer = 0
        Dim dPercentMax As Double = 0
        Dim dDiff As Double = 0
        tdbg2.UpdateData()

        If tdbcChangeNo.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Nghiep_vu"))
            tdbcChangeNo.Focus()
            Return False
        End If
        If tdbg1.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tdbg1.Focus()
            Return False
        End If
        If tdbg2.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tdbg2.Focus()
            Return False
        End If
        For i As Integer = 0 To tdbg2.RowCount - 1
            iColMax = ReturnColPercentMax(i)
            If Number(tdbg2(i, "Density")) > 1 Then
                D99C0008.MsgL3(rl3("Tong_ty_trong_khong_duoc_vuot_qua_100%") & Space(1) & rl3("Ban_phai_nhap_lai_ty_trong"))
                tdbg2.Col = iColMax
                tdbg2.Row = i
                tdbg2.SplitIndex = SPLIT1
                tdbg2.Focus()
                Return False
            End If

            For j As Integer = 5 To tdbg2.Columns.Count - 1
                dx = GetDx(j)
                If dx = 0 Or dx = 2 Then
                    If Number(tdbg2(i, j)) > 100 Then
                        D99C0008.MsgL3(rl3("Ty_trong_khong_duoc_vuot_qua_100%")) '(rl3("Ty_trong_khong_duoc_vuot_qua_100%"))
                        tdbg2.Col = j
                        tdbg2.Row = i
                        tdbg2.SplitIndex = SPLIT1
                        tdbg2.Focus()
                        Return False
                    End If
                End If

                If Number(tdbg2(i, "Density")) <> 100 Then
                    dDiff = 1 - Number(tdbg2(i, "Density"))
                    tdbg2(i, iColMax) = SQLNumber((Number(tdbg2(i, iColMax)) + dDiff).ToString) ', D02Format.Percentage)
                    tdbg2(i, "Density") = 1
                End If
                ReturnValue(i, j)

                If Not IsDBNull(tdbg2(i, j)) And tdbg2(i, j).ToString <> "" Then
                    If (CDbl(tdbg2(i, j)) > MaxMoney) Then ' Or (CDbl(tdbg2(i, j)) < (-MaxMoney)) Then
                        D99C0008.MsgL3(rl3("So_khong_hop_le"))
                        tdbg2.Col = j
                        tdbg2.Row = i
                        tdbg2.SplitIndex = SPLIT1
                        tdbg2.Focus()
                        Return False
                    End If
                End If
            Next
        Next
        Return True
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0202
    '# Created User: 
    '# Created Date: 07/12/2007 04:25:39
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0202() As StringBuilder

        Dim sSQL As New StringBuilder

        _batchID = CreateIGE("D02T0012", "BatchID", "02", "BD", gsStringKey)

        sSQL.Append("Insert Into D02T0202(")
        sSQL.Append("BatchID, AssetID, ChangeNo, TranMonth, TranYear, ")
        sSQL.Append("DivisionID, ChangeDate, DecisionNo, Notes1U, Notes2U, Notes3U,")
        sSQL.Append(" CreateDate, CreateUserID, LastModifyDate, LastModifyUserID")

        sSQL.Append(") Values(")
        sSQL.Append(SQLString(_batchID) & COMMA) 'BatchID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString("...") & COMMA) 'AssetID, varchar[20], NULL
        sSQL.Append(SQLString(tdbcChangeNo.Text) & COMMA) 'ChangeNo, varchar[20], NULL
        sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NULL
        sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, smallint, NULL
        sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NULL
        sSQL.Append(SQLDateSave(c1dateChangeDate.Value) & COMMA) 'ChangeDate, datetime, NULL
        sSQL.Append(SQLString(txtDecisionNo.Text) & COMMA) 'DecisionNo, varchar[20], NULL

        sSQL.Append(SQLStringUnicode(txtNotes.Text, gbUnicode, True) & COMMA) 'Notes1, varchar[250], NULL
        sSQL.Append(SQLStringUnicode(txtNotes2.Text, gbUnicode, True) & COMMA) 'Notes2, varchar[250], NULL
        sSQL.Append(SQLStringUnicode(txtNotes3.Text, gbUnicode, True) & COMMA) 'Notes3, varchar[250], NULL

        sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NULL
        sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID)) 'LastModifyUserID, varchar[20], NULL

        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0012s
    '# Created User: 
    '# Created Date: 10/12/2007 11:06:33
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0012s_tdbg1() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        Dim sTransactionID As String
        sTransactionID = ""

        Dim iCountIGE As Int32 = 0

        For i As Integer = 0 To tdbg1.RowCount - 1
            If Number(tdbg1(i, COL1_Total).ToString) > 0 Then
                For j As Integer = 4 To tdbg1.Columns.Count - 1
                    If tdbg1(i, j).ToString <> "" And Number(tdbg1(i, j).ToString) > 0 Then
                        iCountIGE += 1
                    End If
                Next
            End If
        Next
        For i As Integer = 0 To tdbg1.RowCount - 1
            If Number(tdbg1(i, COL1_Total)) > 0 Then
                For j As Integer = 4 To tdbg1.Columns.Count - 1
                    If tdbg1(i, j).ToString <> "" Then
                        If CDbl(tdbg1(i, j).ToString) > 0 Then
                            sTransactionID = CreateIGEs("D02T0012", "TransactionID", "02", "TD", gsStringKey, sTransactionID, iCountIGE)

                            sSQL.Append("Insert Into D02T0012(")
                            sSQL.Append("TransactionID, DivisionID, ModuleID, SplitNo, AssetID, ")
                            sSQL.Append("VoucherTypeID, VoucherNo, VoucherDate, TranMonth, TranYear, ")
                            sSQL.Append("TransactionDate, Description, DescriptionU, CurrencyID, ExchangeRate, DebitAccountID, ")
                            sSQL.Append("CreditAccountID, OriginalAmount, ConvertedAmount, Status, TransactionTypeID, ")
                            sSQL.Append("RefNo, RefDate, Disabled, CreateUserID, CreateDate, ")
                            sSQL.Append("LastModifyUserID, LastModifyDate, SeriNo, ObjectTypeID, ObjectID, ")
                            sSQL.Append("BatchID, VATObjectTypeID, VATObjectID, ObjectName, VATNo, ")
                            sSQL.Append("VATGroupID, VATTypeID, Ana01ID, Ana02ID, Ana03ID, ")
                            sSQL.Append("Ana04ID, Ana05ID, Ana06ID, Ana07ID, Ana08ID, ")
                            sSQL.Append("Ana09ID, Ana10ID, CipID, Notes,NotesU, AssignmentID, ")
                            sSQL.Append("NormID, Posted, SourceID, DebitObjectTypeID, DebitObjectID, ")
                            sSQL.Append("CreditObjectTypeID, CreditObjectID, SplitBatchID, Internal, DeprTableID, ")
                            sSQL.Append("Str01, Str02, Str03, Str04, Str05,Str01U, Str02U, Str03U, Str04U, Str05U, ")
                            sSQL.Append("Num01, Num02, Num03, Num04, Num05, ")
                            sSQL.Append("Date01, Date02, Date03, Date04, Date05 ")
                            sSQL.Append(") Values(")
                            sSQL.Append(SQLString(sTransactionID) & COMMA) 'TransactionID [KEY], varchar[20], NOT NULL
                            sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID [KEY], varchar[20], NOT NULL
                            sSQL.Append(SQLString("02") & COMMA) 'ModuleID [KEY], varchar[20], NOT NULL
                            sSQL.Append("0" & COMMA) 'SplitNo [KEY], int, NOT NULL

                            sSQL.Append(SQLString(tdbg1(i, COL1_AssetID)) & COMMA) 'AssetID, varchar[20], NULL

                            sSQL.Append(SQLString("") & COMMA) 'VoucherTypeID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'VoucherNo, varchar[20], NULL
                            sSQL.Append("NULL" & COMMA) 'VoucherDate, datetime, NULL
                            sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NULL
                            sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, smallint, NULL
                            sSQL.Append("NULL" & COMMA) 'TransactionDate, datetime, NULL
                            sSQL.Append("'', N''" & COMMA) 'Description, varchar[250], NULL
                            sSQL.Append(SQLString(DxxFormat.BaseCurrencyID) & COMMA) 'CurrencyID, varchar[20], NOT NULL
                            sSQL.Append(SQLMoney("1", DxxFormat.ExchangeRateDecimals) & COMMA) 'ExchangeRate, money, NOT NULL
                            sSQL.Append(SQLString("") & COMMA) 'DebitAccountID, varchar[20], NULL
                            sSQL.Append(SQLString(tdbg1(i, COL1_AccountID)) & COMMA) 'CreditAccountID, varchar[20], NULL
                            If Not IsDBNull(tdbg1(i, j)) And tdbg1(i, j).ToString <> "" Then
                                sSQL.Append(SQLMoney(tdbg1(i, j), DxxFormat.DecimalPlaces) & COMMA) 'OriginalAmount, money, NULL
                                sSQL.Append(SQLMoney(tdbg1(i, j), DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL
                            Else
                                sSQL.Append(SQLMoney("") & COMMA) 'OriginalAmount, money, NULL
                                sSQL.Append(SQLMoney("") & COMMA) 'ConvertedAmount, money, NULL
                            End If
                            sSQL.Append("0" & COMMA) 'Status, tinyint, NOT NULL
                            sSQL.Append(SQLString("SC") & COMMA) 'TransactionTypeID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'RefNo, varchar[20], NULL
                            sSQL.Append("NULL" & COMMA) 'RefDate, datetime, NULL
                            sSQL.Append("0" & COMMA) 'Disabled, bit, NOT NULL
                            sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
                            sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
                            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
                            sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
                            sSQL.Append(SQLString("") & COMMA) 'SeriNo, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'ObjectTypeID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'ObjectID, varchar[20], NULL
                            sSQL.Append(SQLString(_batchID) & COMMA) 'BatchID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'VATObjectTypeID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'VATObjectID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'ObjectName, varchar[250], NULL
                            sSQL.Append(SQLString("") & COMMA) 'VATNo, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'VATGroupID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'VATTypeID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'Ana01ID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'Ana02ID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'Ana03ID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'Ana04ID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'Ana05ID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'Ana06ID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'Ana07ID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'Ana08ID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'Ana09ID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'Ana10ID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'CipID, varchar[20], NULL
                            sSQL.Append("'',N'',")
                            'sSQL.Append(SQLString("") & COMMA) 'Notes, varchar[250], NULL
                            sSQL.Append(SQLString("") & COMMA) 'AssignmentID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'NormID, varchar[20], NULL
                            sSQL.Append("0" & COMMA) 'Posted, tinyint, NOT NULL
                            sSQL.Append(SQLString(tdbg1.Columns(j).Caption) & COMMA) 'SourceID, varchar[20], NULL

                            sSQL.Append(SQLString("") & COMMA) 'DebitObjectTypeID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'DebitObjectID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'CreditObjectTypeID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'CreditObjectID, varchar[20], NULL
                            sSQL.Append(SQLString("") & COMMA) 'SplitBatchID, varchar[20], NULL
                            sSQL.Append("0" & COMMA) 'Internal, tinyint, NOT NULL
                            sSQL.Append(SQLString("") & COMMA) 'DeprTableID, varchar[20], NULL
                            sSQL.Append("'','','','','',N'',N'',N'',N'',N'',")
                            'sSQL.Append(SQLString("") & COMMA) 'Str01, varchar[250], NULL
                            'sSQL.Append(SQLString("") & COMMA) 'Str02, varchar[250], NULL
                            'sSQL.Append(SQLString("") & COMMA) 'Str03, varchar[250], NULL
                            'sSQL.Append(SQLString("") & COMMA) 'Str04, varchar[250], NULL
                            'sSQL.Append(SQLString("") & COMMA) 'Str05, varchar[250], NULL
                            sSQL.Append("0" & COMMA) 'Num01, money, NULL
                            sSQL.Append("0" & COMMA) 'Num02, money, NULL
                            sSQL.Append("0" & COMMA) 'Num03, money, NULL
                            sSQL.Append("0" & COMMA) 'Num04, money, NULL
                            sSQL.Append("0" & COMMA) 'Num05, money, NULL
                            sSQL.Append("NULL" & COMMA) 'Date01, datetime, NULL
                            sSQL.Append("NULL" & COMMA) 'Date02, datetime, NULL
                            sSQL.Append("NULL" & COMMA) 'Date03, datetime, NULL
                            sSQL.Append("NULL" & COMMA) 'Date04, datetime, NULL
                            sSQL.Append("NULL") 'Date05, datetime, NULL
                            sSQL.Append(")")
                            sRet.Append(sSQL.ToString & vbCrLf)
                            sSQL.Remove(0, sSQL.Length)
                        End If
                    End If
                Next
            End If
        Next
        Return sRet
    End Function

    Private Function GetSourceID(ByVal iCol As Integer) As String
        Dim s As String = ""
        Dim iIndexStart As Integer
        Dim iIndexEnd As Integer
        Dim sSub As String
        s = tdbg2.Columns(iCol).Caption
        iIndexStart = s.IndexOf("[") + 1 '+1 : vị trí của "["
        iIndexEnd = s.IndexOf("]")
        sSub = s.Substring(iIndexStart, (iIndexEnd - iIndexStart))
        Return sSub
    End Function


    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0012s
    '# Created User: 
    '# Created Date: 10/12/2007 11:06:33
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0012s_tdbg2() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        Dim sTransactionID As String
        Dim dx As Integer = 0
        sTransactionID = ""
        Dim sSourceID As String = ""

        Dim iCountIGE As Int32 = 0

        For i As Integer = 0 To tdbg2.RowCount - 1
            If Number(tdbg2(i, "Density")) > 0 Then
                For j As Integer = 5 To tdbg2.Columns.Count - 2
                    dx = GetDx(j)
                    If dx = 0 Or dx = 2 Then
                        If tdbg2(i, j).ToString <> "" And Number(tdbg2(i, j).ToString) > 0 Then
                            iCountIGE += 1
                        End If

                    End If
                Next
            End If

        Next
        Try
            For i As Integer = 0 To tdbg2.RowCount - 1
                If Number(tdbg2(i, "Density")) > 0 Then
                    For j As Integer = 5 To tdbg2.Columns.Count - 2
                        dx = GetDx(j)
                        If dx = 0 Or dx = 2 Then
                            If tdbg2(i, j).ToString <> "" And Number(tdbg2(i, j).ToString) > 0 Then
                                sTransactionID = CreateIGEs("D02T0012", "TransactionID", "02", "TD", gsStringKey, sTransactionID, iCountIGE)
                                sSQL.Append("Insert Into D02T0012(")
                                sSQL.Append("TransactionID, DivisionID, ModuleID, SplitNo, AssetID, ")
                                sSQL.Append("VoucherTypeID, VoucherNo, VoucherDate, TranMonth, TranYear, ")
                                sSQL.Append("TransactionDate, Description,DescriptionU, CurrencyID, ExchangeRate, DebitAccountID, ")
                                sSQL.Append("CreditAccountID, OriginalAmount, ConvertedAmount, Status, TransactionTypeID, ")
                                sSQL.Append("RefNo, RefDate, Disabled, CreateUserID, CreateDate, ")
                                sSQL.Append("LastModifyUserID, LastModifyDate, SeriNo, ObjectTypeID, ObjectID, ")
                                sSQL.Append("BatchID, VATObjectTypeID, VATObjectID, ObjectName,ObjectNameU, VATNo, ")
                                sSQL.Append("VATGroupID, VATTypeID, Ana01ID, Ana02ID, Ana03ID, ")
                                sSQL.Append("Ana04ID, Ana05ID, Ana06ID, Ana07ID, Ana08ID, ")
                                sSQL.Append("Ana09ID, Ana10ID, CipID, Notes,NotesU, AssignmentID, ")
                                sSQL.Append("NormID, Posted, SourceID, DebitObjectTypeID, DebitObjectID, ")
                                sSQL.Append("CreditObjectTypeID, CreditObjectID, SplitBatchID, Internal, DeprTableID, ")
                                sSQL.Append("Str01, Str02, Str03, Str04, Str05,Str01U, Str02U, Str03U, Str04U, Str05U, ")
                                sSQL.Append("Num01, Num02, Num03, Num04, Num05, ")
                                sSQL.Append("Date01, Date02, Date03, Date04, Date05 ")
                                sSQL.Append(") Values(")
                                sSQL.Append(SQLString(sTransactionID) & COMMA) 'TransactionID [KEY], varchar[20], NOT NULL
                                sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID [KEY], varchar[20], NOT NULL
                                sSQL.Append(SQLString("02") & COMMA) 'ModuleID [KEY], varchar[20], NOT NULL
                                sSQL.Append("0" & COMMA) 'SplitNo [KEY], int, NOT NULL

                                sSQL.Append(SQLString(tdbg2(i, COL2_AssetID)) & COMMA) 'AssetID, varchar[20], NULL

                                sSQL.Append(SQLString("") & COMMA) 'VoucherTypeID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'VoucherNo, varchar[20], NULL
                                sSQL.Append("NULL" & COMMA) 'VoucherDate, datetime, NULL
                                sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NULL
                                sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, smallint, NULL
                                sSQL.Append("NULL" & COMMA) 'TransactionDate, datetime, NULL
                                sSQL.Append("'',N'',") 'SQLString("") & COMMA) 'Description, varchar[250], NULL
                                sSQL.Append(SQLString(DxxFormat.BaseCurrencyID) & COMMA) 'CurrencyID, varchar[20], NOT NULL
                                sSQL.Append(SQLMoney("1", DxxFormat.ExchangeRateDecimals) & COMMA) 'ExchangeRate, money, NOT NULL

                                sSQL.Append(SQLString(tdbg2(i, COL2_AccountID)) & COMMA) 'DebitAccountID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'CreditAccountID, varchar[20], NULL
                                If Not IsDBNull(tdbg2(i, j + 1)) And tdbg2(i, j + 1).ToString <> "" Then
                                    sSQL.Append(SQLMoney(tdbg2(i, j + 1), DxxFormat.DecimalPlaces) & COMMA) 'OriginalAmount, money, NULL
                                    sSQL.Append(SQLMoney(tdbg2(i, j + 1), DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL
                                Else
                                    sSQL.Append(SQLMoney("") & COMMA) 'OriginalAmount, money, NULL
                                    sSQL.Append(SQLMoney("") & COMMA) 'ConvertedAmount, money, NULL
                                End If

                                sSQL.Append("0" & COMMA) 'Status, tinyint, NOT NULL
                                sSQL.Append(SQLString("SC") & COMMA) 'TransactionTypeID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'RefNo, varchar[20], NULL
                                sSQL.Append("NULL" & COMMA) 'RefDate, datetime, NULL
                                sSQL.Append("0" & COMMA) 'Disabled, bit, NOT NULL
                                sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
                                sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
                                sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
                                sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
                                sSQL.Append(SQLString("") & COMMA) 'SeriNo, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'ObjectTypeID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'ObjectID, varchar[20], NULL
                                sSQL.Append(SQLString(_batchID) & COMMA) 'BatchID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'VATObjectTypeID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'VATObjectID, varchar[20], NULL
                                sSQL.Append("'',N'',") 'SQLString("") & COMMA) 'ObjectName, varchar[250], NULL
                                sSQL.Append(SQLString("") & COMMA) 'VATNo, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'VATGroupID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'VATTypeID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'Ana01ID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'Ana02ID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'Ana03ID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'Ana04ID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'Ana05ID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'Ana06ID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'Ana07ID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'Ana08ID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'Ana09ID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'Ana10ID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'CipID, varchar[20], NULL
                                sSQL.Append("'',N'',") 'SQLString("") & COMMA) 'Notes, varchar[250], NULL
                                sSQL.Append(SQLString("") & COMMA) 'AssignmentID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'NormID, varchar[20], NULL
                                sSQL.Append("0" & COMMA) 'Posted, tinyint, NOT NULL
                                sSourceID = GetSourceID(j)
                                sSQL.Append(SQLString(sSourceID) & COMMA) 'SourceID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'DebitObjectTypeID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'DebitObjectID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'CreditObjectTypeID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'CreditObjectID, varchar[20], NULL
                                sSQL.Append(SQLString("") & COMMA) 'SplitBatchID, varchar[20], NULL
                                sSQL.Append("0" & COMMA) 'Internal, tinyint, NOT NULL
                                sSQL.Append(SQLString("") & COMMA) 'DeprTableID, varchar[20], NULL
                                'sSQL.Append(SQLString("") & COMMA) 'Str01, varchar[250], NULL
                                'sSQL.Append(SQLString("") & COMMA) 'Str02, varchar[250], NULL
                                'sSQL.Append(SQLString("") & COMMA) 'Str03, varchar[250], NULL
                                'sSQL.Append(SQLString("") & COMMA) 'Str04, varchar[250], NULL
                                'sSQL.Append(SQLString("") & COMMA) 'Str05, varchar[250], NULL
                                sSQL.Append("'','','','','',N'',N'',N'',N'',N'',")
                                sSQL.Append("0" & COMMA) 'Num01, money, NULL
                                sSQL.Append("0" & COMMA) 'Num02, money, NULL
                                sSQL.Append("0" & COMMA) 'Num03, money, NULL
                                sSQL.Append("0" & COMMA) 'Num04, money, NULL
                                sSQL.Append("0" & COMMA) 'Num05, money, NULL
                                sSQL.Append("NULL" & COMMA) 'Date01, datetime, NULL
                                sSQL.Append("NULL" & COMMA) 'Date02, datetime, NULL
                                sSQL.Append("NULL" & COMMA) 'Date03, datetime, NULL
                                sSQL.Append("NULL" & COMMA) 'Date04, datetime, NULL
                                sSQL.Append("NULL") 'Date05, datetime, NULL
                                sSQL.Append(")")

                                sRet.Append(sSQL.ToString & vbCrLf)
                                sSQL.Remove(0, sSQL.Length)
                            End If
                        End If

                    Next
                End If
            Next
        Catch ex As Exception
            D99C0008.MsgL3(ex.Message)
        End Try
        Return sRet
    End Function

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Cap_nhat_nghiep_vu_chuyen_nguon_-_D02F2007") & UnicodeCaption(gbUnicode) 'CËp nhËt nghiÖp vó chuyÓn nguän - D02F2007
        '================================================================ 
        lblteChangeDate.Text = rl3("Ngay_chuyen") 'Ngày chuyển
        lblDecisionNo.Text = rl3("So_hieu") 'Số hiệu
        lblChangeNo.Text = rl3("Nghiep_vu") 'Nghiệp vụ
        lblNotes.Text = rl3("Ghi_chu") & " 1" 'rl3("Ghi_chu_1") 'Ghi chú 1
        lblNotes2.Text = rl3("Ghi_chu") & " 2" 'Ghi chú 2
        lblNotes3.Text = rl3("Ghi_chu") & " 3" 'Ghi chú 3
        '================================================================ 
        btnSave.Text = rl3("_Luu") '&Lưu
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        '================================================================ 
        tdbcChangeNo.Columns("ChangeNo").Caption = rl3("Ma") 'Mã
        tdbcChangeNo.Columns("ChangeName").Caption = rl3("Ten") 'Tên

        '================================================================ 
        tdbg1.Columns("AssetID").Caption = rl3("Ma_tai_san") 'Mã tài sản
        tdbg1.Columns("AssetName").Caption = rl3("Ten_tai_san") 'Tên tài sản
        tdbg1.Columns("AccountID").Caption = rl3("Tai_khoan") 'Tài khoản
        tdbg1.Columns("Total").Caption = rl3("Tong_cong") 'Tổng cộng
        tdbg2.Columns("AssetID").Caption = rl3("Ma_tai_san") 'Mã tài sản
        tdbg2.Columns("AssetName").Caption = rl3("Ten_tai_san") 'Tên tài sản
        tdbg2.Columns("AccountID").Caption = rl3("Tai_khoan") 'Tài khoản
        tdbg2.Columns("Total").Caption = rl3("Tong_cong") 'Tổng cộng
    End Sub

    Private Sub tdbg2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg2.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If tdbg2.Col = iLastCol2 Then
                    HotKeyEnterGrid(tdbg2, COL2_AssetID, e)
                End If
            Case Keys.F7
                HotKeyF7(tdbg2)
        End Select
        HotKeyDownGrid(e, tdbg2, COL2_AssetID, SPLIT0, SPLIT1, True, True, True, -1, "")
    End Sub

    Private Sub tdbg2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg2.KeyPress
        Select Case tdbg2.Col
            Case COL2_Total
                'e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select

        For i As Integer = 5 To tdbg2.Columns.Count - 1
            If tdbg2.Col = i Then
                e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
            End If
        Next
    End Sub

    Private Sub tdbg1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg1.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                If tdbg1.Col = iLastCol1 Then
                    HotKeyEnterGrid(tdbg1, COL1_AssetID, e)
                End If
        End Select
    End Sub

    'Trả về cột có tỷ lệ lớn nhất
    Private Function ReturnColPercentMax(ByVal iRow As Integer) As Integer
        Dim dPercentMax As Double = 0
        Dim iReturnColPercentMax As Integer = -1
        Dim dx As Integer = 0

        For j As Integer = 5 To tdbg2.Columns.Count - 1
            dx = GetDx(j)
            If dx = 0 Or dx = 2 Then
                If Number(tdbg2(iRow, j)) > dPercentMax Then
                    dPercentMax = Number(tdbg2(iRow, j))
                    iReturnColPercentMax = j
                End If
            End If
        Next

        If iReturnColPercentMax = -1 Then iReturnColPercentMax = 5
        Return iReturnColPercentMax
    End Function

    Private Sub ReturnValue(ByVal iRow As Integer, ByVal iCol As Integer)
        Dim dx As Integer
        dx = GetDx(iCol)
        If dx = 1 Then
            'tdbg2(iRow, iCol - 1) = SQLNumber((Number(tdbg2(iRow, iCol - 1).ToString) * 100).ToString) ', D02Format.Percentage)
            tdbg2(iRow, iCol) = Number(tdbg2(iRow, iCol - 1)) * Number(tdbg2(iRow, COL2_Total))
            tdbg2(iRow, iCol) = SQLNumber(tdbg2(iRow, iCol).ToString, DxxFormat.D90_ConvertedDecimals)

            If Number(tdbg2(iRow, iCol)) = 0 And tdbg2(iRow, iCol - 1).ToString = "" Then
                tdbg2(iRow, iCol - 1) = 0
            End If
        End If
    End Sub

End Class