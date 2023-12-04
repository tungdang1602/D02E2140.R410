'#-------------------------------------------------------------------------------------
'# Created Date: 26/09/2007 4:37:14 PM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 26/09/2007 4:37:14 PM
'# Modify User: Trần Thị ÁiTrâm
'#-------------------------------------------------------------------------------------
Imports System.Text


Public Class D02F1008

#Region "Const of tdbg"
    Private Const COL_RefDate As Integer = 0            ' Ngày hóa đơn
    Private Const COL_SeriNo As Integer = 1             ' Số Sêri
    Private Const COL_RefNo As Integer = 2              ' Số hóa đơn
    Private Const COL_ObjectTypeID As Integer = 3       ' Mã loại đối tượng
    Private Const COL_ObjectID As Integer = 4           ' Mã đối tượng
    Private Const COL_Description As Integer = 5        ' Diễn giải
    Private Const COL_CurrencyID As Integer = 6         ' Loại tiền
    Private Const COL_ExchangeRate As Integer = 7       ' Tỷ giá
    Private Const COL_OriginalAmount As Integer = 8     ' Nguyên tệ
    Private Const COL_ConvertedAmount As Integer = 9    ' Qui đổi
    Private Const COL_VATTypeID As Integer = 10         ' Loại hóa đơn
    Private Const COL_VATNo As Integer = 11             ' Mã số thuế
    Private Const COL_VATGroupID As Integer = 12        ' Nhóm thuế
    Private Const COL_ObjectName As Integer = 13        ' Tên đối tượng GTGT
    Private Const COL_CipID As Integer = 14             ' CipID
    Private Const COL_TransactionTypeID As Integer = 15 ' TransactionTypeID
    Private Const COL_BatchID As Integer = 16           ' BatchID
    Private Const COL_ModuleID As Integer = 17          ' ModuleID
    Private Const COL_Status As Integer = 18            ' Status
    Private Const COL_Ana01ID As Integer = 19           ' Khoản mục 01
    Private Const COL_Ana02ID As Integer = 20           ' Khoản mục 02
    Private Const COL_Ana03ID As Integer = 21           ' Khoản mục 03
    Private Const COL_Ana04ID As Integer = 22           ' Khoản mục 04
    Private Const COL_Ana05ID As Integer = 23           ' Khoản mục 05
    Private Const COL_Ana06ID As Integer = 24           ' Khoản mục 06
    Private Const COL_Ana07ID As Integer = 25           ' Khoản mục 07
    Private Const COL_Ana08ID As Integer = 26           ' Khoản mục 08
    Private Const COL_Ana09ID As Integer = 27           ' Khoản mục 09
    Private Const COL_Ana10ID As Integer = 28           ' Khoản mục 10
    Private Const COL_TransactionID As Integer = 29     ' TransactionID
#End Region


    Private bInsertRow As Boolean = False

    Private _batchID As String
    Private _cipID As String
    Private _transactionTypeID As String
    Dim sTransactionID As String
    Private dtObject As DataTable
    Private dtMain As DataTable
    Private dtExchangeRate As DataTable
    Private iLastCol As Integer = 0
    Private sArrTransactionID(10000) As String
    Private sArrTransactionTypeID(10000) As String
    Private row As Integer = 0
    Private iLengthArr As Integer = 0
    Private bDelete As Boolean = False

    '---Kiểm tra khoản mục theo chuẩn gồm 6 bước
    '--- Chuẩn Khoản mục b1: Khai báo biến

#Region "Biến khai báo cho khoản mục"

    Private Const SplitAna As Int16 = 1 ' Ghi nhận Khoản mục chứa ở Split nào
    Dim bUseAna As Boolean 'Kiểm tra có sử dụng Khoản mục không, để set thuộc tính Enabled nút Khoản mục 
    Dim iDisplayAnaCol As Integer = 0 ' Cột Khoản mục đầu tiên được hiển thị, khi nhấn nút Khoản mục thì Focus đến cột đó
    Dim xCheckAna(9) As Boolean 'Khởi động tại Form_load: Ghi lại việc kiểm tra lần đầu Lưu, khi nhấn Lưu lần thứ 2 thì không cần kiểm tra nữa

#End Region

    'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b1:
    Dim sEditVoucherTypeID As String = ""
    Dim oFilterCombo As Lemon3.Controls.FilterCombo

    Dim clsFilterDropdown As Lemon3.Controls.FilterDropdown

    Public Property CipID() As String
        Get
            Return _cipID
        End Get
        Set(ByVal value As String)
            If CipID = value Then
                _cipID = ""
                Return
            End If
            _cipID = value
        End Set
    End Property

    Public Property TransactionTypeID() As String
        Get
            Return _transactionTypeID
        End Get
        Set(ByVal value As String)
            If TransactionTypeID = value Then
                _transactionTypeID = ""
                Return
            End If
            _transactionTypeID = value
        End Set
    End Property

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
            '--- Chuẩn Khoản mục b2: Lấy caption cho 10 khoản mục
            bUseAna = LoadTDBGridAnalysisCaption(D02, tdbg, COL_Ana01ID, SplitAna, gbUnicode)
            'SetNewXaCheckAna()
            'D91 có sử dụng Khoản mục
            'If bUseAna Then iDisplayAnaCol = 1
            If Not bUseAna Then tdbg.Splits(SplitAna).SplitSize = 0

            '18/7/2017, id 99844-Bổ sung điều kiện tìm kiếm Mã XDCB, Tên XDCB
            oFilterCombo = New Lemon3.Controls.FilterCombo
            oFilterCombo.CheckD91 = True
            oFilterCombo.UseFilterCombo(tdbcCipNo)

            clsFilterDropdown = New Lemon3.Controls.FilterDropdown()
            clsFilterDropdown.CheckD91 = True
            clsFilterDropdown.UseFilterDropdown(tdbg, COL_ObjectID)

            '------------------------------------
            'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b2:
            Select Case _FormState
                Case EnumFormState.FormAdd
                    btnSave.Enabled = True
                    btnNext.Enabled = False

                    LoadTDBCombo()
                    LoadTDBDropDown()
                    LoadAddNew()
                Case EnumFormState.FormEdit
                    btnSave.Enabled = True
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    LoadEdit()
                    LoadTDBDropDown()
                Case EnumFormState.FormView
                    btnSave.Enabled = False
                    btnNext.Visible = False
                    btnSave.Left = btnNext.Left
                    LoadEdit()

                    LoadTDBDropDown()
            End Select
        End Set
    End Property

    Private Sub D02F1008_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
        If e.Control And e.KeyCode = Keys.F1 Then
            btnHotKey_Click(Nothing, Nothing)
        End If
        If e.KeyCode = Keys.F11 Then
            HotKeyF11(Me, tdbg)
        End If
    End Sub

    Private Sub D02F1008_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Loadlanguage()
        ResetSplitDividerSize(tdbg)
        SetBackColorObligatory()
        InputDateInTrueDBGrid(tdbg, COL_RefDate)
        iPer_F5558 = ReturnPermission("D02F5558")
        InputbyUnicode(Me, gbUnicode)
        LoadTDBGridAnalysisCaption(D02, tdbg, COL_Ana01ID, SPLIT1, True, gbUnicode)
        tdbg_LockedColumns()
        tdbg_NumberFormat()
        iLastCol = CountCol(tdbg, SPLIT1)
    SetResolutionForm(Me)
Me.Cursor = Cursors.Default
End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Dim dtVoucherTypeID As DataTable
    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b5:
        'Load tdbcVoucherTypeID
        dtVoucherTypeID = ReturnDataTable(ReturnTableVoucherTypeID("D02", gsDivisionID, sEditVoucherTypeID, gbUnicode))
        LoadDataSource(tdbcVoucherTypeID, dtVoucherTypeID, gbUnicode)
        'LoadVoucherTypeID(tdbcVoucherTypeID, D02, sEditVoucherTypeID, gbUnicode)
        'Load tdbcCipNo
        'sSQL = "Select CipID, CipNo, CipName" & UnicodeJoin(gbUnicode) & " as CipName,Description" & UnicodeJoin(gbUnicode) & " as DescriptionMaster, D02T0100.AccountID, Account.AccountName" & UnicodeJoin(gbUnicode) & " as AccountName, D02T0100.Status, D02T0100.Disabled, D02T0100.CreateDate, D02T0100.CreateUserID, D02T0100.LastModifyUserID, D02T0100.LastModifyDate " & vbCrLf
        'sSQL &= " From D02T0100 D02T0100 WITH(NOLOCK) Inner Join Account WITH(NOLOCK) On Account.AccountID=D02T0100.AccountID " & vbCrLf
        'sSQL &= "Where Status <> 2 And D02T0100.Disabled=0 And DivisionID=" & SQLString(gsDivisionID)

        'ID 89282 Bo cau select,su dung store
        LoadDataSource(tdbcCipNo, SQLStoreD02P1008, gbUnicode)
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


#Region "Events tdbcVoucherTypeID with txtVoucherTypeName"

    'Private Sub tdbcVoucherTypeID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.Close
    '    If tdbcVoucherTypeID.FindStringExact(tdbcVoucherTypeID.Text) = -1 Then
    '        tdbcVoucherTypeID.Text = ""
    '        txtVoucherNo.Text = ""
    '    End If
    'End Sub

    'Private Sub tdbcVoucherTypeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.SelectedValueChanged

    '    If _FormState = EnumFormState.FormAdd Or _FormState = EnumFormState.FormOther Then
    '        If Not (tdbcVoucherTypeID.Tag Is Nothing OrElse tdbcVoucherTypeID.Tag.ToString = "") Then
    '            tdbcVoucherTypeID.Tag = ""
    '            Exit Sub
    '        End If
    '        GetVoucherNo(tdbcVoucherTypeID, txtVoucherNo, btnSetNewKey)
    '    End If
    'End Sub

#End Region

#Region "Events tdbcCipNo with txtCipNoName"
    Private Sub tdbcCipNo_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCipNo.SelectedValueChanged
        txtCipNoName.Text = tdbcCipNo.Columns("CipName").Value.ToString
        txtAccountID.Text = tdbcCipNo.Columns("AccountID").Value.ToString
        txtAccountName.Text = tdbcCipNo.Columns("AccountName").Value.ToString
        txtDescriptionMaster.Text = tdbcCipNo.Columns("DescriptionMaster").Value.ToString
    End Sub

    Private Sub tdbcCipNo_Validated(sender As Object, e As EventArgs) Handles tdbcCipNo.Validated
        '18/7/2017, id 99844-Bổ sung điều kiện tìm kiếm Mã XDCB, Tên XDCB
        oFilterCombo.FilterCombo(tdbcCipNo, e)
        If tdbcCipNo.FindStringExact(tdbcCipNo.Text) = -1 Then 'Code của sự kiện LostFocus
            tdbcCipNo.Text = ""
            txtCipNoName.Text = ""
            txtAccountID.Text = ""
            txtAccountName.Text = ""
        End If
    End Sub

#End Region

    Private Sub btnSetNewKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GetNewVoucherNo(tdbcVoucherTypeID, txtVoucherNo)
    End Sub

    Private Sub LoadTDBDropDown()
        Dim sSQL As String = ""
        'Load tdbdObjectTypeID
        sSQL = "Select ObjectTypeID," & IIf(geLanguage = EnumLanguage.Vietnamese, "ObjectTypeName" & UnicodeJoin(gbUnicode), "ObjectTypeName01" & UnicodeJoin(gbUnicode)).ToString & " as ObjectTypeName " & vbCrLf
        sSQL &= "From D91T0005 WITH(NOLOCK) Order By ObjectTypeID"
        LoadDataSource(tdbdObjectTypeID, sSQL, gbUnicode)
        'Load tdbdObjectID
        sSQL = "Select ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " as ObjectName, ObjectTypeID, VATNo From Object WITH(NOLOCK) Where Disabled=0 Order By ObjectID "
        dtObject = ReturnDataTable(sSQL)
        'Load tdbdCurrencyID
        sSQL = "Select D91T0010.CurrencyID, D91T0010.CurrencyName" & UnicodeJoin(gbUnicode) & " as CurrencyName, D91T0010.ExchangeRate, D91T0010.Operator," & vbCrLf
        sSQL &= "(Case When D91T0010.CurrencyID=A.BaseCurrencyID Then D90_ConvertedDecimals Else D91T0010.DecimalPlaces End) As DecimalPlaces" & vbCrLf
        sSQL &= " From D91T0010 WITH(NOLOCK), (Select Top 1 BaseCurrencyID, D90_ConvertedDecimals From D91T0025 WITH(NOLOCK)) As A " & vbCrLf
        sSQL &= "Order By CurrencyID"
        LoadDataSource(tdbdCurrencyID, sSQL, gbUnicode)
        'Load tdbdVATTypeID
        sSQL = "Select VATTypeID, Description" & UnicodeJoin(gbUnicode) & "  as Description From D91T9001 WITH(NOLOCK) " & vbCrLf
        If geLanguage = EnumLanguage.Vietnamese Then
            sSQL &= "Where Language='84'" & vbCrLf
        ElseIf geLanguage = EnumLanguage.English Then
            sSQL &= "Where Language='01'" & vbCrLf
        End If
        sSQL &= " Order By VATTypeID "
        LoadDataSource(tdbdVATTypeID, sSQL, gbUnicode)
        'Load tdbdVATGroupID
        sSQL = "Select VATGroupID, VATGroupName" & UnicodeJoin(gbUnicode) & " as VATGroupName, RateTax From D91T0040 WITH(NOLOCK) Where Disabled=0 Order By VATGroupID "
        LoadDataSource(tdbdVATGroupID, sSQL, gbUnicode)
        '--- Chuẩn Khoản mục b3: Load 10 khoản mục
        LoadTDBDropDownAna(tdbdAna01ID, tdbdAna02ID, tdbdAna03ID, tdbdAna04ID, tdbdAna05ID, tdbdAna06ID, tdbdAna07ID, tdbdAna08ID, tdbdAna09ID, tdbdAna10ID, tdbg, COL_Ana01ID, gbUnicode)
        '------------------------------------------
    End Sub

    Private Sub SetBackColorObligatory()
        tdbcVoucherTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        txtVoucherNo.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcCipNo.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private Sub LoadAddNew()
        c1dateVoucherDate.Value = Date.Today
        txtDescriptionMaster.Enabled = False
        chkPosted.Checked = True
        _batchID = ""
        LoadForm()
        If tdbcVoucherTypeID.Text = "" Then
            For i As Integer = 0 To dtVoucherTypeID.Rows.Count - 1
                If dtVoucherTypeID.Rows(i).Item("FormID").ToString = "D02F1008" Then
                    Dim sFormID As String = dtVoucherTypeID.Rows(i).Item("VoucherTypeID").ToString
                    tdbcVoucherTypeID.Text = sFormID
                    Exit Sub
                End If
            Next
        End If
    End Sub

    Private Sub LoadEdit()
        tdbcVoucherTypeID.Enabled = False
        txtVoucherNo.ReadOnly = True
        txtDescriptionMaster.Enabled = False
        LoadForm()
    End Sub

    Private Sub LoadForm()
        Dim sSQL As New StringBuilder()
        'sSQL.Append(" Select Convert(Varchar(20), A.RefDate,103) As RefDate, A.RefNo, A.SeriNo, A.ObjectTypeID, A.ObjectID,A.Description" & UnicodeJoin(gbUnicode) & " as Description,  " & vbCrLf)
        sSQL.Append(" Select RefDate, A.RefNo, A.SeriNo, A.ObjectTypeID, A.ObjectID,A.Description" & UnicodeJoin(gbUnicode) & " as Description,  " & vbCrLf)
        sSQL.Append("A.CurrencyID, A.ExchangeRate, A.OriginalAmount, A.ConvertedAmount, A.VATTypeID, A.VATNo," & vbCrLf)
        sSQL.Append("A.VATGroupID, A.ObjectName" & UnicodeJoin(gbUnicode) & " as ObjectName , A.CipID, A.TransactionTypeID, A.BatchID, A.ModuleID,A.Status," & vbCrLf)
        sSQL.Append("A.Ana01ID, A.Ana02ID, A.Ana03ID, A.Ana04ID, A.Ana05ID, A.Ana06ID, A.Ana07ID, A.Ana08ID, A.Ana09ID, A.Ana10ID, A.TransactionID, " & vbCrLf)
        sSQL.Append("A.VoucherTypeID, A.VoucherNo, A.VoucherDate, A.Posted, " & vbCrLf)
        sSQL.Append(" B.CipID, B.CipNo, B.CipName" & UnicodeJoin(gbUnicode) & " as CipName, B.Description" & UnicodeJoin(gbUnicode) & " As DescriptionMaster, B.AccountID, " & vbCrLf)
        sSQL.Append(" B.Disabled, B.Status " & vbCrLf)
        sSQL.Append(" From D02T0012 A WITH(NOLOCK) Inner Join D02T0100 B WITH(NOLOCK) On A.CipID = B.CipID  " & vbCrLf)
        sSQL.Append(" Inner Join Account On Account.AccountID = B.AccountID " & vbCrLf)
        sSQL.Append(" Where BatchID = " & SQLString(_batchID) & " And TransactionTypeID = 'SDXDCB'  And A.CipID =" & SQLString(_cipID) & vbCrLf)

        dtMain = ReturnDataTable(sSQL.ToString)
        If dtMain.Rows.Count > 0 Then
            With dtMain.Rows(0)
                'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b3:
                sEditVoucherTypeID = .Item("VoucherTypeID").ToString
                LoadTDBCombo()
                '-----------------------------------------------------------
                tdbcVoucherTypeID.Text = .Item("VoucherTypeID").ToString
                txtVoucherNo.Text = .Item("VoucherNo").ToString
                c1dateVoucherDate.Value = SQLDateShow(.Item("VoucherDate"))
                tdbcCipNo.Text = .Item("CipNo").ToString
                txtDescriptionMaster.Text = .Item("DescriptionMaster").ToString
                txtCipNoName.Text = .Item("CipName").ToString
                txtAccountID.Text = .Item("AccountID").ToString
                chkPosted.Checked = L3Bool(IIf(.Item("Posted").ToString = "0", False, True))
            End With
        End If
        LoadDataSource(tdbg, dtMain, gbUnicode)
    End Sub

    Private Sub LoadtdbdObjectID(ByVal sObjectTypeID As String)
        'LoadDataSource(tdbdObjectID, ReturnTableFilter(dtObject, "ObjectTypeID=" & SQLString(sObjectTypeID)), gbUnicode)
        If clsFilterDropdown.IsNewFilter Then
            tdbdObjectID.DisplayColumns("ObjectTypeID").Visible = (sObjectTypeID = "" Or sObjectTypeID = "-1")
            If sObjectTypeID = "" Then
                LoadDataSource(tdbdObjectID, ReturnTableFilter(dtObject, "", True), gbUnicode)
            Else
                LoadDataSource(tdbdObjectID, ReturnTableFilter(dtObject, "ObjectTypeID=" & SQLString(sObjectTypeID), True), gbUnicode)
            End If
        Else
            LoadDataSource(tdbdObjectID, ReturnTableFilter(dtObject, "ObjectTypeID=" & SQLString(sObjectTypeID)), gbUnicode)
        End If
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        ClearText(Me)
        btnSave.Enabled = True
        btnNext.Enabled = False
        tdbcCipNo.SelectedValue = ""
        tdbcVoucherTypeID.Text = ""
        txtVoucherNo.Text = ""
        LoadAddNew()

        tdbcVoucherTypeID.Focus()
    End Sub

    Private Sub tdbg_LockedColumns()
        'tdbg.Splits(SPLIT0).DisplayColumns(COL_ConvertedAmount).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
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



    'Private Sub tdbg_BeforeDelete(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.CancelEventArgs) Handles tdbg.BeforeDelete

    '    sArrTransactionID(row) = tdbg.Columns(COL_TransactionID).Text
    '    sArrTransactionTypeID(row) = tdbg.Columns(COL_TransactionTypeID).Text
    '    iLengthArr = row
    '    If row < sArrTransactionID.Length Then
    '        row = row + 1
    '    End If
    '    bDelete = True
    'End Sub

    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown

        If e.KeyCode = Keys.Enter Then
            If tdbg.Col = iLastCol Then
                HotKeyEnterGrid(tdbg, COL_RefDate, e)
            End If
        End If

        If e.Shift Then
            If e.KeyCode = Keys.Insert Then
                bInsertRow = True
                HotKeyShiftInsert(tdbg, 0, COL_RefDate, tdbg.Columns.Count)
            End If
        End If
        If e.KeyCode = Keys.F7 Then
            If tdbg.Splits(tdbg.SplitIndex).DisplayColumns(tdbg.Col).Locked = False Then
                HotKeyF7Other(tdbg)
            Else
                D99C0008.MsgL3(MsgLockedColumn, L3MessageBoxIcon.Exclamation)
                Return
            End If
        End If
        If e.KeyCode = Keys.F8 Then
            HotKeyF8(tdbg)
        End If
        If e.Alt And e.Control And e.KeyCode = Keys.C Then
            If tdbg.Splits(tdbg.SplitIndex).DisplayColumns(tdbg.Col).Locked = False Then
                CopyColumn(tdbg, tdbg.Col, tdbg.Columns(tdbg.Col).Text)
            Else
                D99C0008.MsgL3(MsgLockedColumn, L3MessageBoxIcon.Exclamation)
                Return
            End If
        End If

        If e.KeyCode = Keys.Delete OrElse (e.Control And e.KeyCode = Keys.Delete) Then
            sArrTransactionID(row) = tdbg.Columns(COL_TransactionID).Text
            sArrTransactionTypeID(row) = tdbg.Columns(COL_TransactionTypeID).Text
            iLengthArr = row
            If row < sArrTransactionID.Length Then
                row = row + 1
            End If
            bDelete = True
        End If
        HotKeyDownGrid(e, tdbg, COL_RefDate, 0, 1, True, True, True, COL_Description, txtDescriptionMaster.Text)

        If clsFilterDropdown.CheckKeydownFilterDropdown(tdbg, e) Then
            Select Case tdbg.Col
                Case COL_ObjectID
                    Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg, tdbg.Columns(tdbg.Col).DataField)
                    If tdbd Is Nothing Then Exit Select
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg, e, tdbd)
                    If dr Is Nothing Then Exit Sub
                    AfterColUpdate(tdbg.Col, dr)
                    Exit Sub
            End Select

        End If
    End Sub

    Public Sub HotKeyF7Other(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid)
        Try
            If c1Grid.RowCount < 1 Then Exit Sub

            If c1Grid(c1Grid.Row, c1Grid.Col).ToString() = "" Then
                c1Grid.Columns(c1Grid.Col).Text = c1Grid(c1Grid.Row - 1, c1Grid.Col).ToString
                If c1Grid.Col = COL_CurrencyID Then
                    c1Grid.Columns(COL_ExchangeRate).Text = c1Grid(c1Grid.Row - 1, COL_ExchangeRate).ToString
                ElseIf c1Grid.Col = COL_ObjectID Then
                    c1Grid.Columns(COL_VATNo).Text = c1Grid(c1Grid.Row - 1, COL_VATNo).ToString
                    c1Grid.Columns(COL_ObjectName).Text = c1Grid(c1Grid.Row - 1, COL_ObjectName).ToString
                ElseIf c1Grid.Col = COL_OriginalAmount Then
                    If c1Grid.Columns(COL_ExchangeRate).Text <> "" Then
                        'c1Grid.Columns(COL_ConvertedAmount).Text = c1Grid(c1Grid.Row - 1, COL_ConvertedAmount).ToString
                        CalcuteConvertedAmount()
                    Else
                        c1Grid.Columns(COL_ConvertedAmount).Text = "0"
                    End If
                End If
                tdbg.UpdateData()

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub tdbg_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tdbg.KeyPress
        Select Case tdbg.Col
            Case COL_SeriNo
                e.KeyChar = UCase(e.KeyChar)
                'Case COL_ExchangeRate
                '    e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
                'Case COL_OriginalAmount
                '    e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
                'Case COL_ConvertedAmount
                '    e.Handled = CheckKeyPress(e.KeyChar, EnumKey.NumberDot)
        End Select
    End Sub

    Private Sub tdbg_ComboSelect(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.ComboSelect
        Select Case e.ColIndex
            Case COL_CurrencyID
                GetExchangeRate()
                tdbg.Columns(COL_ExchangeRate).Text = dtExchangeRate.Rows(0).Item("ExchangeRate").ToString
                CalcuteConvertedAmount()
            Case COL_ObjectTypeID
                tdbg.Columns(COL_ObjectTypeID).Text = tdbdObjectTypeID.Columns("ObjectTypeID").Text
                tdbg.Columns(COL_ObjectID).Text = ""
                tdbg.Columns(COL_ObjectName).Text = ""
                tdbg.Columns(COL_VATNo).Text = ""
                'LoadtdbdObjectID(tdbg.Columns(COL_ObjectTypeID).Text)
            Case COL_ObjectID
                tdbg.Columns(COL_ObjectName).Text = tdbdObjectID.Columns("ObjectName").Text
                tdbg.Columns(COL_VATNo).Text = tdbdObjectID.Columns("VATNo").Text

        End Select
    End Sub

    Private Sub CalcuteConvertedAmount()
        Dim dExchangeRate As Double = 0
        Dim dOriginalAmount As Double = 0
        Dim dConvertedAmount As Double
        If tdbg.Columns(COL_ExchangeRate).Text <> "" And tdbg.Columns(COL_OriginalAmount).Text <> "" Then
            dExchangeRate = CDbl(tdbg.Columns(COL_ExchangeRate).Text)
            dOriginalAmount = CDbl(tdbg.Columns(COL_OriginalAmount).Text)
            If tdbdCurrencyID.Columns("Operator").Text <> "" Then
                If CInt(tdbdCurrencyID.Columns("Operator").Text) = 0 Then
                    dConvertedAmount = dExchangeRate * dOriginalAmount
                    tdbg.Columns(COL_ConvertedAmount).Text = SQLNumber(dConvertedAmount.ToString, DxxFormat.D90_ConvertedDecimals)
                Else
                    If dExchangeRate <> 0 Then
                        dConvertedAmount = dOriginalAmount / dExchangeRate
                        tdbg.Columns(COL_ConvertedAmount).Text = SQLNumber(dConvertedAmount.ToString, DxxFormat.D90_ConvertedDecimals)
                    Else
                        D99C0008.MsgL3(rL3("Nguyen_te_khong_hop_le"))
                        Exit Sub
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub tdbg_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg.BeforeColUpdate
        Select Case e.ColIndex
            'Case COL_RefDate
            '    tdbg.Columns(e.ColIndex).Text = L3DateValue(tdbg.Columns(e.ColIndex).Text)

            Case COL_ObjectTypeID
                If tdbg.Columns(COL_ObjectTypeID).Text <> tdbdObjectTypeID.Columns("ObjectTypeID").Text Then
                    tdbg.Columns(COL_ObjectTypeID).Text = ""
                    tdbg.Columns(COL_ObjectID).Text = ""
                    tdbg.Columns(COL_ObjectName).Text = ""
                    tdbg.Columns(COL_VATNo).Text = ""
                End If
            Case COL_ObjectID
                If clsFilterDropdown.IsNewFilter Then Exit Sub
                If tdbg.Columns(COL_ObjectID).Text <> tdbdObjectID.Columns("ObjectID").Text Then
                    tdbg.Columns(COL_ObjectID).Text = ""
                    tdbg.Columns(COL_ObjectName).Text = ""
                    tdbg.Columns(COL_VATNo).Text = ""
                End If

            Case COL_CurrencyID
                If tdbg.Columns(COL_CurrencyID).Text <> tdbdCurrencyID.Columns("CurrencyID").Text Then
                    tdbg.Columns(COL_CurrencyID).Text = ""
                    tdbg.Columns(COL_ExchangeRate).Text = ""
                End If
                'Case COL_ExchangeRate
                '    If Not IsNumeric(tdbg.Columns(COL_ExchangeRate).Text) Then e.Cancel = True
                'Case COL_OriginalAmount
                '    If Not IsNumeric(tdbg.Columns(COL_OriginalAmount).Text) Then e.Cancel = True
                'Case COL_ConvertedAmount
                '    If Not IsNumeric(tdbg.Columns(COL_ConvertedAmount).Text) Then e.Cancel = True
            Case COL_VATTypeID
                If tdbg.Columns(COL_VATTypeID).Text <> tdbdVATTypeID.Columns("VATTypeID").Text Then
                    tdbg.Columns(COL_VATTypeID).Text = ""
                End If
            Case COL_VATNo
            Case COL_VATGroupID
                If tdbg.Columns(COL_VATGroupID).Text <> tdbdVATGroupID.Columns("VATGroupID").Text Then
                    tdbg.Columns(COL_VATGroupID).Text = ""
                End If

                '--- Chuẩn Khoản mục b5: Kiểm tra Khoản mục lúc nhập liệu
                '---------------------------------------------
            Case COL_Ana01ID
                If tdbg.Columns(COL_Ana01ID).Text <> tdbdAna01ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(0) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana01ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana01ID).Text.Length > giArrAnaLength(0) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana01ID).Text = ""
                        End If
                    End If
                End If

            Case COL_Ana02ID
                If tdbg.Columns(COL_Ana02ID).Text <> tdbdAna02ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(1) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana02ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana02ID).Text.Length > giArrAnaLength(1) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana02ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana03ID
                If tdbg.Columns(COL_Ana03ID).Text <> tdbdAna03ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(2) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana03ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana03ID).Text.Length > giArrAnaLength(2) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana03ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana04ID
                If tdbg.Columns(COL_Ana04ID).Text <> tdbdAna04ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(3) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana04ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana04ID).Text.Length > giArrAnaLength(3) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana04ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana05ID
                If tdbg.Columns(COL_Ana05ID).Text <> tdbdAna05ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(4) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana05ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana05ID).Text.Length > giArrAnaLength(4) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana05ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana06ID
                If tdbg.Columns(COL_Ana06ID).Text <> tdbdAna06ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(5) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana06ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana06ID).Text.Length > giArrAnaLength(5) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana06ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana07ID
                If tdbg.Columns(COL_Ana07ID).Text <> tdbdAna07ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(6) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana07ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana07ID).Text.Length > giArrAnaLength(6) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana07ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana08ID
                If tdbg.Columns(COL_Ana08ID).Text <> tdbdAna08ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(7) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana08ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana08ID).Text.Length > giArrAnaLength(7) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana08ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana09ID
                If tdbg.Columns(COL_Ana09ID).Text <> tdbdAna09ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(8) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana09ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana09ID).Text.Length > giArrAnaLength(8) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana09ID).Text = ""
                        End If
                    End If
                End If
            Case COL_Ana10ID
                If tdbg.Columns(COL_Ana10ID).Text <> tdbdAna10ID.Columns("AnaID").Text Then
                    If gbArrAnaValidate(9) Then 'Kiểm tra nhập trong danh sách
                        tdbg.Columns(COL_Ana10ID).Text = ""
                    Else
                        If tdbg.Columns(COL_Ana10ID).Text.Length > giArrAnaLength(9) Then ' Kiểm tra chiều dài nhập vào
                            tdbg.Columns(COL_Ana10ID).Text = ""
                        End If
                    End If
                End If
                '---------------------------------------------
        End Select
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD91P0010
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 11/09/2007 10:36:53
    '# Modified User: 
    '# Modified Date: 
    '# Description: Lấy tỷ giá qui đổi
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD91P0010() As String
        Dim sSQL As String = ""
        sSQL &= "Exec D91P0010 "
        sSQL &= SQLString(tdbg.Columns(COL_CurrencyID).Text) & COMMA 'CurrencyID, varchar[20], NOT NULL
        'If tdbg.Columns(COL_RefDate).Text = "  /  /" Then
        '    sSQL &= SQLDateSave("") 'ExDate, datetime, NOT NULL
        'Else
        sSQL &= SQLDateSave(tdbg.Columns(COL_RefDate).Text) 'ExDate, datetime, NOT NULL
        'End If

        Return sSQL
    End Function

    Private Sub GetExchangeRate()
        Dim sSQL As String = ""
        sSQL = SQLStoreD91P0010()
        dtExchangeRate = ReturnDataTable(sSQL)
    End Sub

    Private Sub tdbg_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.AfterColUpdate
        If tdbg.Columns(COL_RefDate).Text = "" Then 'tdbg.Columns(COL_RefDate).Text = "  /  /"
            tdbg.Columns(COL_RefDate).Text = Date.Today.ToShortDateString
        End If
        Select Case e.ColIndex

            'Case COL_RefDate
            '    tdbg.Columns(e.ColIndex).Value = tdbg.Columns(e.ColIndex).Text
            Case COL_ObjectTypeID
                If tdbg.Columns(COL_ObjectTypeID).Text = "" Then
                    tdbg.Columns(COL_ObjectID).Text = ""
                    tdbg.Columns(COL_ObjectName).Text = ""
                    tdbg.Columns(COL_VATNo).Text = ""
                End If
            Case COL_ExchangeRate
                tdbg.Columns(COL_ExchangeRate).Text = SQLNumber(tdbg.Columns(COL_ExchangeRate).Text, DxxFormat.ExchangeRateDecimals)
                CalcuteConvertedAmount()

            Case COL_OriginalAmount
                tdbg.Columns(COL_OriginalAmount).Text = SQLNumber(tdbg.Columns(COL_OriginalAmount).Text, DxxFormat.DecimalPlaces)
                CalcuteConvertedAmount()

            Case COL_ConvertedAmount
                tdbg.Columns(COL_ConvertedAmount).Text = SQLNumber(tdbg.Columns(COL_ConvertedAmount).Text, DxxFormat.D90_ConvertedDecimals)

            Case COL_ObjectID
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg, e.Column.DataColumn.DataField)
                If tdbd Is Nothing Then Exit Select
                If clsFilterDropdown.IsNewFilter Then
                    Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg, e, tdbd)
                    AfterColUpdate(e.ColIndex, dr)
                    Exit Sub
                Else ' Nhập liệu dạng cũ (xổ dropdown)
                    '   Dim row As DataRow = ReturnDataRow(tdbd, tdbd.DisplayMember & "=" & SQLString(tdbg.Columns(e.ColIndex).Text))
                    Dim row As DataRow = Nothing
                    If tdbg.Columns(e.ColIndex).Text <> "" Then row = CType(tdbd.DataSource, DataTable).Rows(tdbd.Row) 'Sửa lỗi bị khi chọn Mă trùng 82152
                    AfterColUpdate(e.ColIndex, row)
                End If
        End Select
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If AskSave() = Windows.Forms.DialogResult.No Then Exit Sub
        tdbg.UpdateData()
        If Not AllowSave() Then Exit Sub

        'Kiểm tra Ngày phiếu có phù hợp với kỳ kế toán hiện tại không (gọi hàm CheckVoucherDateInPeriod)
        If Not CheckVoucherDateInPeriod(c1dateVoucherDate.Value.ToString) Then Exit Sub
        btnSave.Enabled = False
        btnClose.Enabled = False

        gbSavedOK = False
        Me.Cursor = Cursors.WaitCursor
        Dim sSQL As New StringBuilder
        Select Case _FormState
            Case EnumFormState.FormAdd
                _batchID = CreateIGE("D02T0012", "BatchID", "02", "BC", gsStringKey)
                '****************************************
                'Kiểm tra phiếu theo kiểu mới
                'Kiểm tra phiếu
                If tdbcVoucherTypeID.Columns("Auto").Text = "1" And bEditVoucherNo = False Then 'Sinh tự động và không nhấn F2
                    txtVoucherNo.Text = CreateIGEVoucherNoNew(tdbcVoucherTypeID, "D02T0012", _batchID)
                Else 'Không sinh tự động hay có nhấn F2
                    If bEditVoucherNo = False Then
                        'Kiểm tra trùng Số phiếu
                        If CheckDuplicateVoucherNoNew(D02, "D02T0012", _batchID, txtVoucherNo.Text) = True Then btnSave.Enabled = True : btnClose.Enabled = True : Me.Cursor = Cursors.Default : Exit Sub
                    Else 'Có nhấn F2 để sửa số phiếu
                        'Insert Số phiếu vào bảng D40T5558
                        InsertD02T5558(_batchID, sOldVoucherNo, txtVoucherNo.Text)
                    End If
                    'Insert VoucherNo vào bảng D91T9111
                    InsertVoucherNoD91T9111(txtVoucherNo.Text, "D02T0012", _batchID)
                End If
                bEditVoucherNo = False
                sOldVoucherNo = ""
                bFirstF2 = False
                '****************************************

                sSQL.Append(SQLInsertD02T0012s)
                If tdbcCipNo.Columns("Status").Value.ToString = "0" Then
                    sSQL.Append(SQLUpdateD02T0100())
                End If
                'If giAppMode = 1 Then ' Online
                '    If tdbcVoucherTypeID.Columns("Auto").Text <> "" Or Not IsDBNull(tdbcVoucherTypeID.Columns("Auto").Text) Then
                '        If CInt(tdbcVoucherTypeID.Columns("Auto").Text) <> 0 Then ' Tạo mã tự động
                '            SaveNewLastKey(tdbcVoucherTypeID)
                '        End If
                '    End If
                'Else
                '    If tdbcVoucherTypeID.Columns("Auto").Text <> "" Or Not IsDBNull(tdbcVoucherTypeID.Columns("Auto").Text) Then
                '        If CInt(tdbcVoucherTypeID.Columns("Auto").Text) <> 0 Then ' Tạo mã tự động
                '            CreateIGEVoucherNo(tdbcVoucherTypeID, True)
                '        End If
                '    End If
                'End If

                ''Kiểm tra trùng phiếu 
                'If CheckDuplicateVoucherNo(D02, "D02T0012", _batchID, txtVoucherNo.Text) Then
                '    Me.Cursor = Cursors.Default
                '    btnSave.Enabled = True
                '    btnClose.Enabled = True
                '    Exit Sub
                'End If
            Case EnumFormState.FormEdit
                If bDelete = True Then
                    sSQL.Append(SQLDeleteD02T0012s() & vbCrLf)
                    ExecuteSQL(sSQL.ToString)
                End If
                sSQL = New StringBuilder("")
                sSQL.Append(SQLUpdateD02T0012s)
                sSQL.Append(vbCrLf)
                If tdbcCipNo.Columns("Status").Value.ToString = "0" Then
                    sSQL.Append(SQLUpdateD02T0100())
                End If
                ' Thay tdbcCipNo.Columns("CipID").Value.ToString = CipID theo incident 52222 của Thị Hiệp bởi Văn Vinh
                sSQL.Append("If Not Exists (Select Top 1 1" & vbCrLf)
                sSQL.Append(" From D02T0012 WITH(NOLOCK) Where CipID=" & SQLString(CipID) & ")" & vbCrLf) 'tdbcCipNo.Columns("CipID").Value.ToString
                sSQL.Append(" Begin " & vbCrLf)
                sSQL.Append(" Update D02T0100 " & vbCrLf)
                sSQL.Append(" Set Status =0 " & vbCrLf)
                sSQL.Append(" Where CipID = " & SQLString(CipID) & vbCrLf)
                sSQL.Append("End ")
        End Select

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            gbSavedOK = True
            _cipID = tdbcCipNo.Columns("CipID").Text
            btnClose.Enabled = True
            Select Case _FormState
                Case EnumFormState.FormAdd
                    btnNext.Enabled = True
                    btnNext.Focus()
                Case EnumFormState.FormEdit
                    bDelete = False
                    btnSave.Enabled = True
                    btnClose.Focus()
            End Select
        Else
            If _FormState = EnumFormState.FormAdd Then
                DeleteVoucherNoD91T9111_Transaction(txtVoucherNo.Text, "D02T0012", "VoucherNo", tdbcVoucherTypeID, bEditVoucherNo)
            End If
            SaveNotOK()
            btnClose.Enabled = True
            btnSave.Enabled = True
        End If
    End Sub

    Private Function AllowSave() As Boolean
        If tdbcVoucherTypeID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Loai_phieu"))
            tdbcVoucherTypeID.Focus()
            Return False
        End If
        If txtVoucherNo.Text.Trim = "" Then
            D99C0008.MsgNotYetEnter(rL3("So_phieu"))
            txtVoucherNo.Focus()
            Return False
        End If
        If tdbcCipNo.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rL3("Ma_XDCB"))
            tdbcCipNo.Focus()
            Return False
        End If
        If tdbg.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tdbg.Focus()
            Return False
        End If
        For i As Integer = 0 To tdbg.RowCount - 1
            If tdbg(i, COL_ObjectTypeID).ToString <> "" Then
                If tdbg(i, COL_ObjectID).ToString = "" Then
                    D99C0008.MsgNotYetEnter(rL3("Ma_doi_tuong"))
                    tdbg.SplitIndex = SPLIT0
                    tdbg.Col = COL_ObjectID
                    tdbg.Bookmark = i
                    'tdbg.Focus()
                    Return False
                End If
            End If
            If tdbg(i, COL_ObjectID).ToString <> "" Then
                If tdbg(i, COL_ObjectTypeID).ToString = "" Then
                    D99C0008.MsgNotYetEnter(rL3("Ma_loai_doi_tuong"))
                    tdbg.SplitIndex = SPLIT0
                    tdbg.Col = COL_ObjectTypeID
                    tdbg.Bookmark = i
                    'tdbg.Focus()
                    Return False
                End If
            End If
            If tdbg(i, COL_CurrencyID).ToString = "" Then
                D99C0008.MsgNotYetEnter(rL3("Loai_tien"))
                tdbg.SplitIndex = SPLIT0
                tdbg.Col = COL_CurrencyID
                tdbg.Bookmark = i
                'tdbg.Focus()
                Return False
            End If
            If tdbg(i, COL_OriginalAmount).ToString = "" Then
                D99C0008.MsgNotYetEnter(rL3("Nguyen_te"))
                tdbg.SplitIndex = SPLIT0
                tdbg.Col = COL_OriginalAmount
                tdbg.Bookmark = i
                'tdbg.Focus()
                Return False
            End If
            If tdbg(i, COL_ExchangeRate).ToString <> "" Then
                If CDbl(tdbg(i, COL_ExchangeRate)) > MaxMoney Then
                    D99C0008.MsgNotYetEnter(rL3("Ty_gia_qua_lon"))
                    tdbg.SplitIndex = SPLIT0
                    tdbg.Col = COL_ExchangeRate
                    tdbg.Bookmark = i
                    'tdbg.Focus()
                    Return False
                End If
            End If
            If tdbg(i, COL_OriginalAmount).ToString <> "" Then
                If CDbl(tdbg(i, COL_OriginalAmount)) > MaxMoney Then
                    D99C0008.MsgNotYetEnter(rL3("Nguyen_te_qua_lon"))
                    tdbg.SplitIndex = SPLIT0
                    tdbg.Col = COL_OriginalAmount
                    tdbg.Bookmark = i
                    'tdbg.Focus()
                    Return False
                End If
            End If
        Next
        Return True
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0012s
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 27/09/2006 02:59:39
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0012s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        sTransactionID = ""
        Dim iCountIGE As Int32 = 0

        For i As Integer = 0 To tdbg.RowCount - 1
            If tdbg(i, COL_TransactionID).ToString = "" Then
                iCountIGE += 1
            End If
        Next
        For i As Integer = 0 To tdbg.RowCount - 1
            If tdbg(i, COL_TransactionID).ToString = "" Then
                sTransactionID = CreateIGEs("D02T0012", "TransactionID", "02", "TC", gsStringKey, sTransactionID, iCountIGE)
                tdbg(i, COL_TransactionID) = sTransactionID
            End If
            sSQL.Append("Insert Into D02T0012(")
            sSQL.Append("TransactionID, DivisionID, ModuleID, ")
            sSQL.Append("VoucherTypeID, VoucherNo, VoucherDate, TranMonth, TranYear, ")
            sSQL.Append(" CurrencyID, ExchangeRate, DebitAccountID, ")
            sSQL.Append(" OriginalAmount, ConvertedAmount, Status, TransactionTypeID, ")
            sSQL.Append("RefNo, RefDate, Disabled, CreateUserID, CreateDate, ")
            sSQL.Append("LastModifyUserID, LastModifyDate, SeriNo, ObjectTypeID, ObjectID, ")
            sSQL.Append("BatchID,  VATNo, VATGroupID, VATTypeID, ")
            sSQL.Append("Ana01ID, Ana02ID, Ana03ID, Ana04ID, Ana05ID, Ana06ID, Ana07ID, Ana08ID, ")
            sSQL.Append("Ana09ID, Ana10ID, CipID, Posted,DescriptionU,ObjectNameU ")
            sSQL.Append(") Values(")
            sSQL.Append(SQLString(tdbg(i, COL_TransactionID)) & COMMA) 'TransactionID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString("02") & COMMA) 'ModuleID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbcVoucherTypeID.Text) & COMMA) 'VoucherTypeID, varchar[20], NULL
            sSQL.Append(SQLString(txtVoucherNo.Text) & COMMA) 'VoucherNo, varchar[20], NULL
            sSQL.Append(SQLDateSave(c1dateVoucherDate.Value) & COMMA) 'VoucherDate, datetime, NULL
            sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NULL
            sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, smallint, NULL
            sSQL.Append(SQLString(tdbg(i, COL_CurrencyID)) & COMMA) 'CurrencyID, varchar[20], NOT NULL
            sSQL.Append(SQLMoney(tdbg(i, COL_ExchangeRate), DxxFormat.ExchangeRateDecimals) & COMMA) 'ExchangeRate, money, NOT NULL
            sSQL.Append(SQLString(txtAccountID.Text) & COMMA) 'DebitAccountID, varchar[20], NULL
            sSQL.Append(SQLMoney(tdbg(i, COL_OriginalAmount), DxxFormat.DecimalPlaces) & COMMA) 'OriginalAmount, money, NULL
            sSQL.Append(SQLMoney(tdbg(i, COL_ConvertedAmount), DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL
            sSQL.Append(SQLNumber(tdbg(i, COL_Status)) & COMMA) 'Status, tinyint, NOT NULL
            sSQL.Append(SQLString("SDXDCB") & COMMA) 'TransactionTypeID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_RefNo)) & COMMA) 'RefNo, varchar[20], NULL
            'If tdbg(i, COL_RefDate).ToString = " /  /" Then
            '    sSQL.Append(SQLDateSave(Date.Today) & COMMA) 'RefDate, datetime, NULL
            'Else
            sSQL.Append(SQLDateSave(tdbg(i, COL_RefDate)) & COMMA) 'RefDate, datetime, NULL
            'End If
            sSQL.Append("0" & COMMA) 'Disabled, bit, NOT NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
            sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
            sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
            sSQL.Append(SQLString(tdbg(i, COL_SeriNo)) & COMMA) 'SeriNo, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_ObjectTypeID)) & COMMA) 'ObjectTypeID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_ObjectID)) & COMMA) 'ObjectID, varchar[20], NULL
            sSQL.Append(SQLString(_batchID) & COMMA) 'BatchID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_VATNo)) & COMMA) 'VATNo, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_VATGroupID)) & COMMA) 'VATGroupID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_VATTypeID)) & COMMA) 'VATTypeID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_Ana01ID)) & COMMA) 'Ana01ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_Ana02ID)) & COMMA) 'Ana02ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_Ana03ID)) & COMMA) 'Ana03ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_Ana04ID)) & COMMA) 'Ana04ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_Ana05ID)) & COMMA) 'Ana05ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_Ana06ID)) & COMMA) 'Ana06ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_Ana07ID)) & COMMA) 'Ana07ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_Ana08ID)) & COMMA) 'Ana08ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_Ana09ID)) & COMMA) 'Ana09ID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_Ana10ID)) & COMMA) 'Ana10ID, varchar[20], NULL
            sSQL.Append(SQLString(ReturnValueC1Combo(tdbcCipNo, "CipID")) & COMMA) 'CipID, varchar[20], NULL
            sSQL.Append(SQLNumber(chkPosted.Checked) & COMMA) 'Posted, tinyint, NOT NULL
            sSQL.Append(SQLStringUnicode(tdbg(i, COL_Description), gbUnicode, True) & COMMA) 'Description, varchar[250], NULL
            sSQL.Append(SQLStringUnicode(tdbg(i, COL_ObjectName), gbUnicode, True)) 'ObjectName, varchar[250], NULL
            sSQL.Append(")")

            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0100
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 27/09/2006 03:37:00
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0100() As StringBuilder
        Dim sSQL As New StringBuilder

        sSQL.Append("Update D02T0100 Set ")
        sSQL.Append("Status = " & SQLNumber(1)) 'tinyint, NULL
        sSQL.Append(" Where ")
        sSQL.Append("CipID = " & SQLString(tdbcCipNo.Columns("CipID").Value.ToString))

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLStoreD02P1008
    '# Created User: KIM LONG
    '# Created Date: 02/08/2016 10:09:59
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLStoreD02P1008() As String
        Dim sSQL As String = ""
        sSQL &= ("-- 	-- Combo Ma XDCB" & vbCrLf)
        sSQL &= "Exec D02P1008 "
        sSQL &= SQLString(gsDivisionID) & COMMA 'DivisionID, varchar[50], NOT NULL
        sSQL &= SQLString(gsUserID) & COMMA 'UserID, varchar[50], NOT NULL
        sSQL &= SQLString(My.Computer.Name) & COMMA 'HostID, varchar[50], NOT NULL
        sSQL &= SQLString(Me.Name) & COMMA 'FormID, varchar[50], NOT NULL
        sSQL &= SQLNumber(giTranMonth) & COMMA 'TranMonth, int, NOT NULL
        sSQL &= SQLNumber(giTranYear) & COMMA 'TranYear, int, NOT NULL
        sSQL &= SQLNumber(gsLanguage) & COMMA 'Language, int, NOT NULL
        sSQL &= SQLNumber(gbUnicode) 'CodeTable, tinyint, NOT NULL
        Return sSQL
    End Function



    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLUpdateD02T0012s
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 27/09/2007 05:15:27
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLUpdateD02T0012s() As StringBuilder
        Dim sRet As New StringBuilder
        Dim sSQL As New StringBuilder
        sTransactionID = ""
        Dim iCountIGE As Int32 = 0
        For i As Integer = 0 To tdbg.RowCount - 1
            If tdbg(i, COL_TransactionID).ToString = "" Then
                iCountIGE += 1
            End If
        Next
        For i As Integer = 0 To tdbg.RowCount - 1
            If tdbg(i, COL_TransactionID).ToString = "" Then
                sTransactionID = CreateIGEs("D02T0012", "TransactionID", "02", "TC", gsStringKey, sTransactionID, iCountIGE)
                tdbg(i, COL_TransactionID) = sTransactionID
                sSQL.Append("Insert Into D02T0012(")
                sSQL.Append("TransactionID, DivisionID, ModuleID, ")
                sSQL.Append("VoucherTypeID, VoucherNo, VoucherDate, TranMonth, TranYear, ")
                sSQL.Append(" CurrencyID, ExchangeRate, DebitAccountID, ")
                sSQL.Append(" OriginalAmount, ConvertedAmount, Status, TransactionTypeID, ")
                sSQL.Append("RefNo, RefDate, Disabled, CreateUserID, CreateDate, ")
                sSQL.Append("LastModifyUserID, LastModifyDate, SeriNo, ObjectTypeID, ObjectID, ")
                sSQL.Append("BatchID,  VATNo, VATGroupID, VATTypeID, ")
                sSQL.Append("Ana01ID, Ana02ID, Ana03ID, Ana04ID, Ana05ID, Ana06ID, Ana07ID, Ana08ID, ")
                sSQL.Append("Ana09ID, Ana10ID, CipID, Posted,DescriptionU,ObjectNameU ")
                sSQL.Append(") Values(")
                sSQL.Append(SQLString(tdbg(i, COL_TransactionID)) & COMMA) 'TransactionID [KEY], varchar[20], NOT NULL
                sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID [KEY], varchar[20], NOT NULL
                sSQL.Append(SQLString("02") & COMMA) 'ModuleID [KEY], varchar[20], NOT NULL
                sSQL.Append(SQLString(tdbcVoucherTypeID.Text) & COMMA) 'VoucherTypeID, varchar[20], NULL
                sSQL.Append(SQLString(txtVoucherNo.Text) & COMMA) 'VoucherNo, varchar[20], NULL
                sSQL.Append(SQLDateSave(c1dateVoucherDate.Value) & COMMA) 'VoucherDate, datetime, NULL
                sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NULL
                sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, smallint, NULL
                sSQL.Append(SQLString(tdbg(i, COL_CurrencyID)) & COMMA) 'CurrencyID, varchar[20], NOT NULL
                sSQL.Append(SQLMoney(tdbg(i, COL_ExchangeRate), DxxFormat.ExchangeRateDecimals) & COMMA) 'ExchangeRate, money, NOT NULL
                sSQL.Append(SQLString(txtAccountID.Text) & COMMA) 'DebitAccountID, varchar[20], NULL
                sSQL.Append(SQLMoney(tdbg(i, COL_OriginalAmount), DxxFormat.DecimalPlaces) & COMMA) 'OriginalAmount, money, NULL
                sSQL.Append(SQLMoney(tdbg(i, COL_ConvertedAmount), DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL
                'sSQL.Append(SQLMoney(tdbg(i, COL_OriginalAmount)) & COMMA) 'OriginalAmount, money, NULL
                'sSQL.Append(SQLMoney(tdbg(i, COL_ConvertedAmount)) & COMMA) 'ConvertedAmount, money, NULL
                sSQL.Append(SQLNumber(tdbg(i, COL_Status)) & COMMA) 'Status, tinyint, NOT NULL
                sSQL.Append(SQLString("SDXDCB") & COMMA) 'TransactionTypeID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_RefNo)) & COMMA) 'RefNo, varchar[20], NULL
                'If tdbg(i, COL_RefDate).ToString = " /  /" Then
                '    sSQL.Append(SQLDateSave(Date.Today) & COMMA) 'RefDate, datetime, NULL
                'Else
                sSQL.Append(SQLDateSave(tdbg(i, COL_RefDate)) & COMMA) 'RefDate, datetime, NULL
                'End If

                sSQL.Append("0" & COMMA) 'Disabled, bit, NOT NULL
                sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
                sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
                sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
                sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
                sSQL.Append(SQLString(tdbg(i, COL_SeriNo)) & COMMA) 'SeriNo, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_ObjectTypeID)) & COMMA) 'ObjectTypeID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_ObjectID)) & COMMA) 'ObjectID, varchar[20], NULL
                sSQL.Append(SQLString(_batchID) & COMMA) 'BatchID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_VATNo)) & COMMA) 'VATNo, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_VATGroupID)) & COMMA) 'VATGroupID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_VATTypeID)) & COMMA) 'VATTypeID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_Ana01ID)) & COMMA) 'Ana01ID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_Ana02ID)) & COMMA) 'Ana02ID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_Ana03ID)) & COMMA) 'Ana03ID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_Ana04ID)) & COMMA) 'Ana04ID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_Ana05ID)) & COMMA) 'Ana05ID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_Ana06ID)) & COMMA) 'Ana06ID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_Ana07ID)) & COMMA) 'Ana07ID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_Ana08ID)) & COMMA) 'Ana08ID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_Ana09ID)) & COMMA) 'Ana09ID, varchar[20], NULL
                sSQL.Append(SQLString(tdbg(i, COL_Ana10ID)) & COMMA) 'Ana10ID, varchar[20], NULL
                sSQL.Append(SQLString(tdbcCipNo.Columns("CipID").Value.ToString) & COMMA) 'CipID, varchar[20], NULL
                sSQL.Append(SQLNumber(chkPosted.Checked) & COMMA) 'Posted, tinyint, NOT NULL
                sSQL.Append(SQLStringUnicode(tdbg(i, COL_Description), gbUnicode, True) & COMMA) 'Description, varchar[250], NULL
                sSQL.Append(SQLStringUnicode(tdbg(i, COL_ObjectName), gbUnicode, True)) 'ObjectName, varchar[250], NULL
                sSQL.Append(")")

                sRet.Append(sSQL.ToString & vbCrLf)
                sSQL.Remove(0, sSQL.Length)
            Else
                sSQL.Append("Update D02T0012 Set ")
                sSQL.Append("VoucherTypeID = " & SQLString(tdbcVoucherTypeID.Text) & COMMA) 'varchar[20], NULL
                'sSQL.Append("VoucherNo = " & SQLString(txtVoucherNo.Text) & COMMA) 'varchar[20], NULL
                sSQL.Append("VoucherDate = " & SQLDateSave(c1dateVoucherDate.Value) & COMMA) 'datetime, NULL
                sSQL.Append("TranMonth = " & SQLNumber(giTranMonth) & COMMA) 'tinyint, NULL
                sSQL.Append("TranYear = " & SQLNumber(giTranYear) & COMMA) 'smallint, NULL
                sSQL.Append("DescriptionU = " & SQLStringUnicode(tdbg(i, COL_Description), gbUnicode, True) & COMMA) 'varchar[250], NULL
                sSQL.Append("CurrencyID = " & SQLString(tdbg(i, COL_CurrencyID)) & COMMA) 'varchar[20], NOT NULL
                sSQL.Append("ExchangeRate = " & SQLMoney(tdbg(i, COL_ExchangeRate), DxxFormat.ExchangeRateDecimals) & COMMA) 'money, NOT NULL
                sSQL.Append("DebitAccountID = " & SQLString(txtAccountID.Text) & COMMA) 'varchar[20], NULL
                sSQL.Append("OriginalAmount = " & SQLMoney(tdbg(i, COL_OriginalAmount), DxxFormat.DecimalPlaces) & COMMA) 'money, NULL
                sSQL.Append("ConvertedAmount = " & SQLMoney(tdbg(i, COL_ConvertedAmount), DxxFormat.D90_ConvertedDecimals) & COMMA) 'money, NULL
                sSQL.Append("RefNo = " & SQLString(tdbg(i, COL_RefNo)) & COMMA) 'varchar[20], NULL
                sSQL.Append("RefDate = " & SQLDateSave(tdbg(i, COL_RefDate)) & COMMA) 'datetime, NULL
                sSQL.Append("LastModifyUserID = " & SQLString(gsUserID) & COMMA) 'varchar[20], NOT NULL
                sSQL.Append("LastModifyDate = GetDate()" & COMMA) 'datetime, NULL
                sSQL.Append("SeriNo = " & SQLString(tdbg(i, COL_SeriNo)) & COMMA) 'varchar[20], NULL
                sSQL.Append("ObjectTypeID = " & SQLString(tdbg(i, COL_ObjectTypeID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("ObjectID = " & SQLString(tdbg(i, COL_ObjectID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("ObjectNameU = " & SQLStringUnicode(tdbg(i, COL_ObjectName), gbUnicode, True) & COMMA) 'varchar[250], NULL
                sSQL.Append("VATNo = " & SQLString(tdbg(i, COL_VATNo)) & COMMA) 'varchar[20], NULL
                sSQL.Append("VATGroupID = " & SQLString(tdbg(i, COL_VATGroupID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("VATTypeID = " & SQLString(tdbg(i, COL_VATTypeID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("Ana01ID = " & SQLString(tdbg(i, COL_Ana01ID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("Ana02ID = " & SQLString(tdbg(i, COL_Ana02ID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("Ana03ID = " & SQLString(tdbg(i, COL_Ana03ID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("Ana04ID = " & SQLString(tdbg(i, COL_Ana04ID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("Ana05ID = " & SQLString(tdbg(i, COL_Ana05ID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("Ana06ID = " & SQLString(tdbg(i, COL_Ana06ID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("Ana07ID = " & SQLString(tdbg(i, COL_Ana07ID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("Ana08ID = " & SQLString(tdbg(i, COL_Ana08ID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("Ana09ID = " & SQLString(tdbg(i, COL_Ana09ID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("Ana10ID = " & SQLString(tdbg(i, COL_Ana10ID)) & COMMA) 'varchar[20], NULL
                sSQL.Append("CipID = " & SQLString(ReturnValueC1Combo(tdbcCipNo, "CipID")) & COMMA) 'varchar[20], NULL
                sSQL.Append("Posted = " & SQLNumber(chkPosted.Checked)) 'tinyint, NOT NULL
                sSQL.Append(" Where ")
                sSQL.Append("TransactionID = " & SQLString(tdbg(i, COL_TransactionID)) & " And ")
                sSQL.Append("VoucherNo = " & SQLString(txtVoucherNo.Text) & " And ")
                sSQL.Append("BatchID = " & SQLString(_batchID) & " And ")
                sSQL.Append("TransactionTypeID = " & SQLString("SDXDCB"))

                sRet.Append(sSQL.ToString & vbCrLf)
                sSQL.Remove(0, sSQL.Length)
            End If

        Next
        Return sRet
    End Function


    Private Sub tdbg_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbg.RowColChange
  If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
        Select Case tdbg.Col
            Case COL_ObjectID
                LoadtdbdObjectID(tdbg.Columns(COL_ObjectTypeID).Text)
        End Select

        If bInsertRow = True And tdbg.AddNewMode = C1.Win.C1TrueDBGrid.AddNewModeEnum.AddNewCurrent Then
            tdbg.Columns(COL_RefNo).Text = "" ' Gán 1 cột bất kỳ ="" cho lưới
            bInsertRow = False
        End If
    End Sub

    Private Sub btnHotKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHotKey.Click
        Dim f As New D02F7777
        With f
            .CallShowForm(Me.Name)
            .ShowDialog()
        End With
    End Sub

    Public Sub CopyColumn(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ColCopy As Integer, ByVal sValue As String)
        Dim sValue1 As String = ""
        Dim sValue2 As String = ""
        Dim sValue3 As String = ""
        Dim sValue4 As String = ""
        Dim Flag As DialogResult
        Flag = D99C0008.MsgCopyColumn()
        If ColCopy = COL_CurrencyID Then
            sValue1 = c1Grid.Columns(COL_CurrencyID).Text
            sValue2 = c1Grid.Columns(COL_ExchangeRate).Text
            sValue3 = c1Grid.Columns(COL_OriginalAmount).Text
            sValue4 = c1Grid.Columns(COL_ConvertedAmount).Text

        ElseIf ColCopy = COL_ObjectID Then
            sValue1 = c1Grid.Columns(COL_ObjectID).Text
            sValue2 = c1Grid.Columns(COL_ObjectName).Text
            sValue3 = c1Grid.Columns(COL_VATNo).Text
        ElseIf ColCopy = COL_OriginalAmount Then
            sValue1 = c1Grid.Columns(COL_OriginalAmount).Text
            sValue2 = c1Grid.Columns(COL_ConvertedAmount).Text
        End If
        If Flag = Windows.Forms.DialogResult.No Then ' Copy nhung dong con trong
            For i As Integer = 0 To c1Grid.RowCount - 1
                If c1Grid(i, ColCopy).ToString = "" Then
                    c1Grid(i, ColCopy) = sValue
                    If ColCopy = COL_CurrencyID Then
                        c1Grid(i, COL_CurrencyID) = sValue1
                        c1Grid(i, COL_ExchangeRate) = sValue2
                        c1Grid(i, COL_OriginalAmount) = sValue3
                        c1Grid(i, COL_ConvertedAmount) = sValue4
                    ElseIf ColCopy = COL_ObjectID Then
                        c1Grid(i, COL_ObjectID) = sValue1
                        c1Grid(i, COL_ObjectName) = sValue2
                        c1Grid(i, COL_VATNo) = sValue3
                    ElseIf ColCopy = COL_OriginalAmount Then
                        c1Grid(i, COL_OriginalAmount) = sValue1
                        c1Grid(i, COL_ConvertedAmount) = sValue2
                    End If
                End If
            Next
        ElseIf Flag = Windows.Forms.DialogResult.Yes Then ' Copy nhung dong con trong ' Copy het

            For i As Integer = 0 To c1Grid.RowCount - 1
                c1Grid(i, ColCopy) = sValue
                If ColCopy = COL_CurrencyID Then
                    c1Grid(i, COL_CurrencyID) = sValue1
                    c1Grid(i, COL_ExchangeRate) = sValue2
                    c1Grid(i, COL_OriginalAmount) = sValue3
                    c1Grid(i, COL_ConvertedAmount) = sValue4
                ElseIf ColCopy = COL_ObjectID Then
                    c1Grid(i, COL_ObjectID) = sValue1
                    c1Grid(i, COL_ObjectName) = sValue2
                    c1Grid(i, COL_VATNo) = sValue3
                ElseIf ColCopy = COL_OriginalAmount Then
                    c1Grid(i, COL_OriginalAmount) = sValue1
                    c1Grid(i, COL_ConvertedAmount) = sValue2
                End If
            Next
            c1Grid(0, ColCopy) = sValue
            If ColCopy = COL_CurrencyID Then
                c1Grid(0, COL_CurrencyID) = sValue1
                c1Grid(0, COL_ExchangeRate) = sValue2
                c1Grid(0, COL_OriginalAmount) = sValue3
                c1Grid(0, COL_ConvertedAmount) = sValue4
            ElseIf ColCopy = COL_ObjectID Then
                c1Grid(0, COL_ObjectID) = sValue1
                c1Grid(0, COL_ObjectName) = sValue2
                c1Grid(0, COL_VATNo) = sValue3
            ElseIf ColCopy = COL_OriginalAmount Then
                c1Grid(0, COL_OriginalAmount) = sValue1
                c1Grid(0, COL_ConvertedAmount) = sValue2
            End If
        Else
            Exit Sub
        End If
        tdbg.UpdateData()
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rL3("Cap_nhat_so_du_XDCB_-_D02F1008") & UnicodeCaption(gbUnicode) 'CËp nhËt sç d§ XDCB - D02F1008
        '================================================================ 
        lblVoucherTypeID.Text = rL3("Loai_phieu") 'Loại phiếu
        lblteVoucherDate.Text = rL3("Ngay_hach_toan") 'Ngày hạch toán
        lblCipNo.Text = rL3("Ma_XDCB") 'Mã XDCB
        lblDescriptionMaster.Text = rL3("Dien_giai") 'Diễn giải
        lblAccountID.Text = rL3("Tai_khoan_tap_hop") 'Tài khoản tập hợp
        lblVoucherNo.Text = rL3("So_phieu") 'Số phiếu
        '================================================================ 

        btnSave.Text = rL3("_Luu") '&Lưu
        btnNext.Text = rL3("Nhap__tiep") 'Nhập &tiếp
        btnClose.Text = rL3("Do_ng") 'Đó&ng
        btnHotKey.Text = rL3("_Phim_nong") '&Phím nóng
        '================================================================ 
        chkPosted.Text = rL3("Chuyen_but_toan_sang_module_tong_hop") 'Chuyển bút toán sang Module tổng hợp
        '================================================================ 
        grp1.Text = rL3("Chung_tu_hach_toan") 'Chứng từ hạch toán
        '================================================================ 
        tdbcCipNo.Columns("CipNo").Caption = rL3("Ma")  'Mã
        tdbcCipNo.Columns("CipName").Caption = rL3("Ten") 'Tên
        tdbcVoucherTypeID.Columns("VoucherTypeID").Caption = rL3("Loai_phieu") 'Loại phiếu
        tdbcVoucherTypeID.Columns("VoucherTypeName").Caption = rL3("Dien_giai") 'Diễn giải
        '================================================================ 
        tdbdObjectTypeID.Columns("ObjectTypeID").Caption = rL3("Ma") 'Mã
        tdbdObjectTypeID.Columns("ObjectTypeName").Caption = rL3("Ten") 'Tên
        tdbdObjectID.Columns("ObjectID").Caption = rL3("Ma") 'Mã
        tdbdObjectID.Columns("ObjectName").Caption = rL3("Ten") 'Tên
        tdbdObjectID.Columns("ObjectTypeID").Caption = rL3("Loai_DT") 'Tên
        tdbdCurrencyID.Columns("CurrencyID").Caption = rL3("Ma") 'Mã
        tdbdCurrencyID.Columns("CurrencyName").Caption = rL3("Ten") 'Tên
        tdbdVATTypeID.Columns("VATTypeID").Caption = rL3("Ma") 'Mã
        tdbdVATTypeID.Columns("Description").Caption = rL3("Ten") 'Tên
        tdbdVATGroupID.Columns("VATGroupID").Caption = rL3("Ma") 'Mã
        tdbdVATGroupID.Columns("VATGroupName").Caption = rL3("Ten") 'Tên

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
        tdbg.Columns("RefDate").Caption = rL3("Ngay_hoa_don") 'Ngày hóa đơn
        tdbg.Columns("RefNo").Caption = rL3("So_hoa_don") 'Số hóa đơn
        tdbg.Columns("SeriNo").Caption = rL3("So_Seri") 'Số Sêri
        tdbg.Columns("ObjectTypeID").Caption = rL3("Loai_doi_tuong") 'rl3("Ma_loai_doi_tuong") 'Mã loại đối tượng
        tdbg.Columns("ObjectID").Caption = rL3("Ma_doi_tuong") 'Mã đối tượng
        tdbg.Columns("Description").Caption = rL3("Dien_giai") 'Diễn giải
        tdbg.Columns("CurrencyID").Caption = rL3("Loai_tien") 'Loại tiền
        tdbg.Columns("ExchangeRate").Caption = rL3("Ty_gia") 'Tỷ giá
        tdbg.Columns("OriginalAmount").Caption = rL3("Nguyen_te") 'Nguyên tệ
        tdbg.Columns("ConvertedAmount").Caption = rL3("Quy_doi") 'Qui đổi
        tdbg.Columns("VATTypeID").Caption = rL3("Loai_hoa_don") 'Loại hóa đơn
        tdbg.Columns("VATNo").Caption = rL3("Ma_so_thue") 'Mã số thuế
        tdbg.Columns("VATGroupID").Caption = rL3("Nhom_thue") 'Nhóm thuế
        tdbg.Columns("ObjectName").Caption = rL3("Ten_doi_tuong_GTGT") 'Tên đối tượng GTGT
    End Sub

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T0012s
    '# Created User: Trần Thị Ái Trâm
    '# Created Date: 16/11/2009 11:14:08
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T0012s() As String
        Dim sRet As String = ""
        Dim sSQL As String = ""
        For i As Integer = 0 To iLengthArr
            sSQL &= "Delete From D02T0012"
            sSQL &= " Where "
            sSQL &= "TransactionID = " & SQLString(sArrTransactionID(i)) & " And "
            sSQL &= "VoucherNo = " & SQLString(txtVoucherNo.Text) & " And "
            sSQL &= "BatchID = " & SQLString(_batchID) & " And "
            sSQL &= "TransactionTypeID = " & SQLString(sArrTransactionTypeID(i))
            sRet &= sSQL & vbCrLf
        Next
        Return sRet
    End Function


    'Incident 	73638
#Region "Chuẩn hóa sinh số phiếu"
    Dim sOldVoucherNo As String = "" 'Lưu lại số phiếu cũ
    Dim bEditVoucherNo As Boolean = False '= True: có nhấn F2; = False: không
    Dim bFirstF2 As Boolean = False 'Nhấn F2 lần đầu tiên
    Dim iPer_F5558 As Integer = 0 'Phân quyền cho Sửa số phiếu


    Private Sub tdbcVoucherTypeID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.Close
        If tdbcVoucherTypeID.FindStringExact(tdbcVoucherTypeID.Text) = -1 Then
            tdbcVoucherTypeID.Text = ""
            txtVoucherNo.Text = ""
        End If
    End Sub

    Private Sub tdbcVoucherTypeID_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcVoucherTypeID.SelectedValueChanged
        bEditVoucherNo = False
        bFirstF2 = False
        If tdbcVoucherTypeID.SelectedValue Is Nothing OrElse tdbcVoucherTypeID.Text = "" Then
            txtVoucherNo.Text = ""
            ReadOnlyControl(txtVoucherNo)
            Exit Sub
        End If
        If _FormState = EnumFormState.FormAdd Then
            If tdbcVoucherTypeID.Columns("Auto").Text = "1" Then 'Sinh tự động
                txtVoucherNo.Text = CreateIGEVoucherNo(tdbcVoucherTypeID, False)
                ReadOnlyControl(txtVoucherNo)
            Else 'Không sinh tự động
                txtVoucherNo.Text = ""
                UnReadOnlyControl(txtVoucherNo, True)
            End If

        End If
    End Sub

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
                '    .TableName = "D02T0012" 'Tên bảng chứa số phiếu
                '    'Update 21/09/2010
                '    If _FormState = EnumFormState.FormAdd Then
                '        .VoucherID = "" 'Khóa sinh IGE là rỗng
                '    ElseIf _FormState = EnumFormState.FormEdit Then
                '        .VoucherID = _batchID 'Khóa sinh IGE
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
                '        gbSavedOK = True
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
        End If
    End Sub
#End Region

    Private Sub AfterColUpdate(ByVal iCol As Integer, ByVal dr() As DataRow)
        Dim iRow As Integer = tdbg.Row
        If dr Is Nothing OrElse dr.Length = 0 Then
            Dim row As DataRow = Nothing
            AfterColUpdate(iCol, row)
        ElseIf dr.Length = 1 Then
            If tdbg.Bookmark = tdbg.Row AndAlso tdbg.RowCount = tdbg.Row Then 'Đang đứng dòng mới
                Dim dr1 As DataRow = dtMain.NewRow
                ' dtGrid.Rows.InsertAt(dr1, tdbg.Row)'Bỏ 09/06/2017 vì dtGrid.DefaultView.RowFilter "" thì tdbg.Row luôn luôn = 0 nên gắn dữ liệu sai
                dtMain.Rows.Add(dr1) 'Luôn luôn add dưới table
                SetDefaultValues(tdbg, dr1) 'Bổ sung set giá trị mặc định 19/08/2015
                tdbg.Bookmark = tdbg.Row
            End If
            AfterColUpdate(iCol, dr(0))
        Else
            For Each row As DataRow In dr
                If tdbg.Bookmark = tdbg.Row AndAlso tdbg.RowCount = tdbg.Row Then 'Đang đứng dòng mới
                    Dim dr1 As DataRow = dtMain.NewRow
                    ' dtGrid.Rows.InsertAt(dr1, tdbg.Row)'Ánh Bỏ 09/06/2017 vì dtGrid.DefaultView.RowFilter "" thì tdbg.Row luôn luôn = 0 nên gắn dữ liệu sai
                    dtMain.Rows.Add(dr1) 'Luôn luôn add dưới table
                    SetDefaultValues(tdbg, dr1) 'Bổ sung set giá trị mặc định 19/08/2015
                    tdbg.Bookmark = tdbg.Row
                Else
                    tdbg.Row = iRow
                    tdbg.Bookmark = iRow
                End If
                AfterColUpdate(iCol, row)
                tdbg.UpdateData()
                iRow += 1
            Next
            tdbg.Focus()
        End If
    End Sub

    Private Sub AfterColUpdate(ByVal iCol As Integer, ByVal dr As DataRow)
        'Gán lại các giá trị phụ thuộc vào Dropdown
        Select Case iCol
            Case COL_ObjectID
                If dr Is Nothing OrElse dr.Item("ObjectID").ToString = "" Then
                    tdbg.Columns(COL_ObjectID).Text = ""
                    Exit Sub
                End If
                If tdbg.Columns(COL_ObjectTypeID).Text = "" Then
                    tdbg.Columns(COL_ObjectTypeID).Text = dr.Item("ObjectTypeID").ToString
                    LoadtdbdObjectID(tdbg.Columns(COL_ObjectTypeID).Text)
                End If
                tdbg.Columns(COL_ObjectID).Text = dr.Item("ObjectID").ToString
                tdbg.UpdateData()
        End Select
    End Sub

    Private Sub tdbg_ButtonClick(sender As Object, e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.ButtonClick
        If clsFilterDropdown.IsNewFilter = False Then Exit Sub
        If tdbg.AllowUpdate = False Then Exit Sub
        If tdbg.Splits(tdbg.SplitIndex).DisplayColumns(tdbg.Col).Locked Then Exit Sub
        Select Case tdbg.Col
            Case COL_ObjectID
                Dim tdbd As C1.Win.C1TrueDBGrid.C1TrueDBDropdown = clsFilterDropdown.GetDropdown(tdbg, tdbg.Columns(tdbg.Col).DataField)
                If tdbd Is Nothing Then Exit Select
                Dim dr() As DataRow = clsFilterDropdown.FilterDropdown(tdbg, e, tdbd)
                If dr Is Nothing Then Exit Sub
                AfterColUpdate(tdbg.Col, dr)
        End Select
    End Sub
End Class