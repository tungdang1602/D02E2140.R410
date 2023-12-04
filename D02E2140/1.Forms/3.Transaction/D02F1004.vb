'#-------------------------------------------------------------------------------------
'# Created Date: 01/10/2007 4:39:29 PM
'# Created User: Trần Thị ÁiTrâm
'# Modify Date: 01/10/2007 4:39:29 PM
'# Modify User: Trần Thị ÁiTrâm
'#-------------------------------------------------------------------------------------
Imports System.Text
Imports System

Public Class D02F1004

#Region "Const of tdbg"
    Private Const COL_RefDate As Integer = 0         ' Ngày hóa đơn
    Private Const COL_SeriNo As Integer = 1          ' Số Sêri
    Private Const COL_RefNo As Integer = 2           ' Số hóa đơn
    Private Const COL_Description As Integer = 3     ' Diễn giải
    Private Const COL_DebitAccountID As Integer = 4  ' Tài khoản nợ
    Private Const COL_CreditAccountID As Integer = 5 ' Tài khoản có
    Private Const COL_ExchangeRate As Integer = 6    ' Tỷ giá
    Private Const COL_OriginalAmount As Integer = 7  ' Nguyên tệ
    Private Const COL_ConvertedAmount As Integer = 8 ' Qui đổi
    Private Const COL_ObjectTypeID As Integer = 9    ' Loại đối tượng
    Private Const COL_ObjectID As Integer = 10       ' Mã đối tượng
    Private Const COL_ObjectName As Integer = 11     ' Tên đối tượng GTGT
    Private Const COL_VATNo As Integer = 12          ' Mã số thuế
    Private Const COL_VATTypeID As Integer = 13      ' Loại hóa đơn
    Private Const COL_VATGroupID As Integer = 14     ' Nhóm thuế
    Private Const COL_BatchID As Integer = 15        ' BatchID
    Private Const COL_Ana01ID As Integer = 16        ' Khoản mục 01
    Private Const COL_Ana02ID As Integer = 17        ' Khoản mục 02
    Private Const COL_Ana03ID As Integer = 18        ' Khoản mục 03
    Private Const COL_Ana04ID As Integer = 19        ' Khoản mục 04
    Private Const COL_Ana05ID As Integer = 20        ' Khoản mục 05
    Private Const COL_Ana06ID As Integer = 21        ' Khoản mục 06
    Private Const COL_Ana07ID As Integer = 22        ' Khoản mục 07
    Private Const COL_Ana08ID As Integer = 23        ' Khoản mục 08
    Private Const COL_Ana09ID As Integer = 24        ' Khoản mục 09
    Private Const COL_Ana10ID As Integer = 25        ' Khoản mục 10
    Private Const COL_TransactionID As Integer = 26  ' TransactionID
#End Region

    Private _auditCode As String
    Private _accountID As String
    Private _batchID As String
    Private _cipID As String
    Private _cipNo As String
    Private _cipName As String
    Private dtObject As DataTable
    Private dtExchangeRate As DataTable
    Dim sTransactionID As String
    Dim iLastCol As Integer
    Dim bInsertRow As Boolean = False
    Private createUserID As String = ""
    Private createDate As String = ""

    '---Kiểm tra khoản mục theo chuẩn gồm 6 bước
    '--- Chuẩn Khoản mục b1: Khai báo biến

#Region "Biến khai báo cho khoản mục"

    Private Const SplitAna As Int16 = 1 ' Ghi nhận Khoản mục chứa ở Split nào
    Dim bUseAna As Boolean 'Kiểm tra có sử dụng Khoản mục không, để set thuộc tính Enabled nút Khoản mục 
    'Dim iDisplayAnaCol As Integer = 0 ' Cột Khoản mục đầu tiên được hiển thị, khi nhấn nút Khoản mục thì Focus đến cột đó
    'Dim xCheckAna(9) As Boolean 'Khởi động tại Form_load: Ghi lại việc kiểm tra lần đầu Lưu, khi nhấn Lưu lần thứ 2 thì không cần kiểm tra nữa

#End Region

    'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b1:
    Dim sEditVoucherTypeID As String = ""

    Public WriteOnly Property sAuditCode() As String
        Set(ByVal value As String)
            _auditCode = value
        End Set
    End Property

    Private _byAudit As Byte

    Public WriteOnly Property ByAudit() As Byte
        Set(ByVal value As Byte)
            _byAudit = value
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

    Public Property CipNo() As String
        Get
            Return _cipNo
        End Get
        Set(ByVal value As String)
            If CipNo = value Then
                _cipNo = ""
                Return
            End If
            _cipNo = value
        End Set
    End Property

    Public Property CipName() As String
        Get
            Return _cipName
        End Get
        Set(ByVal value As String)
            If CipName = value Then
                _cipName = ""
                Return
            End If
            _cipName = value
        End Set
    End Property

    Public Property AccountID() As String
        Get
            Return _accountID
        End Get
        Set(ByVal value As String)
            If AccountID = value Then
                _accountID = ""
                Return
            End If
            _accountID = value
        End Set
    End Property

    Private _keyID As String = ""
    Public ReadOnly Property KeyID() As String
        Get
            Return _keyID
        End Get
    End Property

    Private _FormState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal value As EnumFormState)
            _FormState = value
            '--- Chuẩn Khoản mục b2: Lấy caption cho 10 khoản mục
            bUseAna = LoadTDBGridAnalysisCaption(D02, tdbg, COL_Ana01ID, SplitAna, True, gbUnicode)
            'SetNewXaCheckAna()
            'D91 có sử dụng Khoản mục
            'If bUseAna Then iDisplayAnaCol = 1
            If Not bUseAna Then tdbg.Splits(1).SplitSize = 0

            '------------------------------------
            'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b2:
           
            Select Case _FormState
                Case EnumFormState.FormAdd
                    btnSave.Enabled = True
                    btnNext.Enabled = False
                    LoadAddNew()
                    LoadTDBCombo()
                    LoadTDBDropDown()
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

    Private Sub LoadAddNew()
        c1dateVoucherDate.Value = Date.Today
        _batchID = ""
        LoadForm()
    End Sub

    Private Sub LoadForm()
        Dim sSQL As String = ""
        sSQL = "Select BatchID, CipID, VoucherTypeID, VoucherNo, VoucherDate, CurrencyID, " & vbCrLf
        sSQL &= "RefDate, SeriNo, RefNo, Description" & UnicodeJoin(gbUnicode) & " as Description, DebitAccountID, " & vbCrLf 'Convert(varchar(20), RefDate, 103) as
        sSQL &= "CreditAccountID, ExchangeRate, OriginalAmount, ConvertedAmount, ObjectTypeID, ObjectID, ObjectName" & UnicodeJoin(gbUnicode) & " as  ObjectName,  " & vbCrLf
        sSQL &= "VATNo, VATTypeID, VATGroupID, Ana01ID, Ana02ID, Ana03ID, Ana04ID, Ana05ID, Ana06ID, Ana07ID, Ana08ID, Ana09ID, Ana10ID, TransactionID, " & vbCrLf
        '*, Convert(varchar(20), RefDate, 103) as RefDate " & vbCrLf
        sSQL &= "CreateUserID, CreateDate" & vbCrLf
        sSQL &= "From D02T0012 WITH(NOLOCK) " & vbCrLf
        sSQL &= "Where BatchID = " & SQLString(_batchID) & " And CipID = " & SQLString(_cipID)
        Dim dt As DataTable = ReturnDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            With dt.Rows(0)
                sEditVoucherTypeID = .Item("VoucherTypeID").ToString
                LoadTDBCombo()
                'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b3:
                tdbcVoucherTypeID.Text = .Item("VoucherTypeID").ToString
                txtVoucherNo.Text = .Item("VoucherNo").ToString
                c1dateVoucherDate.Value = .Item("VoucherDate")
                tdbcCurrencyID.Text = .Item("CurrencyID").ToString

                createUserID = .Item("CreateUserID").ToString
                createDate = .Item("CreateDate").ToString
            End With
        End If
        LoadDataSource(tdbg, dt, gbUnicode)
    End Sub

    Private Sub LoadEdit()
        tdbcVoucherTypeID.Enabled = False
        txtVoucherNo.ReadOnly = True

        'c1dateVoucherDate.Enabled = False
        LoadForm()
    End Sub

    Private Sub D02F1004_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            UseEnterAsTab(Me)
        End If
        If e.KeyCode = Keys.F11 Then
            HotKeyF11(Me, tdbg)
        End If
        If e.Control And e.KeyCode = Keys.F1 Then
            btnHotKey_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub D02F1004_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        Loadlanguage()
        InputDateInTrueDBGrid(tdbg, COL_RefDate) '19/11/2018, id 115915-Lỗi xem phiếu tập hợp chi phí xây dựng cơ bản
        iPer_F5558 = ReturnPermission("D02F5558")
        SetBackColorObligatory()
        InputbyUnicode(Me, gbUnicode)
        ResetSplitDividerSize(tdbg)
        'LoadTDBGridAnalysisCaption(D02, tdbg, COL_Ana01ID, SPLIT1, True, gbUnicode)
        tdbg_LockedColumns()
        tdbg_NumberFormat()
        iLastCol = CountCol(tdbg, SPLIT1)
        SetResolutionForm(Me)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub LoadTDBCombo()
        Dim sSQL As String = ""
        'Trần Thị Ái Trâm - 10/12/2009 - Chuẩn load combo khi Sửa b5:
        'Load tdbcVoucherTypeID
        LoadVoucherTypeID(tdbcVoucherTypeID, D02, sEditVoucherTypeID, gbUnicode)
        'Load tdbcCurrencyID
        sSQL = "Select D91T0010.CurrencyID, D91T0010.CurrencyName" & UnicodeJoin(gbUnicode) & " as CurrencyName, D91T0010.ExchangeRate, D91T0010.Operator, " & vbCrLf
        sSQL &= "(Case  When D91T0010.CurrencyID=A.BaseCurrencyID" & vbCrLf
        sSQL &= "Then D90_ConvertedDecimals" & vbCrLf
        sSQL &= "Else D91T0010.DecimalPlaces End) As DecimalPlaces " & vbCrLf
        sSQL &= " From D91T0010 WITH(NOLOCK), (Select Top 1 BaseCurrencyID, D90_ConvertedDecimals" & vbCrLf
        sSQL &= " From D91T0025 WITH(NOLOCK)) As A" & vbCrLf
        sSQL &= "Order By CurrencyID  " & vbCrLf
        LoadDataSource(tdbcCurrencyID, sSQL, gbUnicode)
    End Sub

    Private Sub LoadTDBDropDown()
        Dim sSQL As String = ""
        Dim sUnicode As String = UnicodeJoin(gbUnicode)
        'Load tdbdObjectTypeID
        'sSQL = "Select ObjectTypeID," & IIf(geLanguage = EnumLanguage.Vietnamese, "ObjectTypeName", "ObjectTypeName01").ToString & " As ObjectTypeName " & vbCrLf
        'sSQL &= "  From D91T0005 Order By ObjectTypeID "
        'LoadDataSource(tdbdObjectTypeID, sSQL)
        LoadObjectTypeID(tdbdObjectTypeID, gbUnicode)
        'Load tdbdObjectID
        sSQL = "Select ObjectID, ObjectName" & sUnicode & " as ObjectName, ObjectTypeID, VATNo From Object WITH(NOLOCK) Where Disabled=0 Order By ObjectID "
        dtObject = ReturnDataTable(sSQL)

        'Load tdbdCreditAccountID
        sSQL = "Select AccountID," & IIf(geLanguage = EnumLanguage.Vietnamese, "AccountName", "AccountName01").ToString & sUnicode & " as AccountName, GroupID" & vbCrLf
        sSQL &= " From D90T0001 WITH(NOLOCK) Where Disabled=0 And AccountStatus=0 And OffAccount=0 Order by AccountID"
        LoadDataSource(tdbdCreditAccountID, sSQL, gbUnicode)

        'Load tdbdVATTypeID
        sSQL = "Select VATTypeID, Description" & sUnicode & " as Description" & vbCrLf
        sSQL &= "From D91T9001 WITH(NOLOCK) Where Language=" & SQLString(gsLanguage) & vbCrLf
        sSQL &= "Order by VATTypeID"
        LoadDataSource(tdbdVATTypeID, sSQL, gbUnicode)
        'Load tdbdVATGroupID
        sSQL = "Select VATGroupID,VATGroupName" & sUnicode & " as VATGroupName, RateTax" & vbCrLf
        sSQL &= "From D91T0040 WITH(NOLOCK)" & vbCrLf
        sSQL &= "Where Disabled = 0" & vbCrLf
        sSQL &= "Order By VATGroupID"
        LoadDataSource(tdbdVATGroupID, sSQL, gbUnicode)

        '--- Chuẩn Khoản mục b3: Load 10 khoản mục
        LoadTDBDropDownAna(tdbdAna01ID, tdbdAna02ID, tdbdAna03ID, tdbdAna04ID, tdbdAna05ID, tdbdAna06ID, tdbdAna07ID, tdbdAna08ID, tdbdAna09ID, tdbdAna10ID, tdbg, COL_Ana01ID, gbUnicode)
        '------------------------------------------
    End Sub

    Private Sub LoadtdbdObjectID(ByVal sObjectTypeID As String)
        LoadDataSource(tdbdObjectID, ReturnTableFilter(dtObject, "ObjectTypeID=" & SQLString(sObjectTypeID)), gbUnicode)
    End Sub


#Region "Events tdbcVoucherTypeID with txtVoucherNo"

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

    'Private Sub tdbcVoucherTypeID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcVoucherTypeID.KeyDown
    '    If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
    '        tdbcVoucherTypeID.Text = ""
    '        txtVoucherNo.Text = ""
    '    End If
    'End Sub

#End Region

#Region "Events tdbcCurrencyID"

    Private Sub tdbcCurrencyID_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tdbcCurrencyID.Close
        If tdbcCurrencyID.FindStringExact(tdbcCurrencyID.Text) = -1 Then tdbcCurrencyID.Text = ""
    End Sub

    Private Sub tdbcCurrencyID_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbcCurrencyID.KeyDown
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then tdbcCurrencyID.Text = ""
    End Sub

#End Region

    Private Sub tdbg_LockedColumns()
        tdbg.Splits(SPLIT0).DisplayColumns(COL_DebitAccountID).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT0).DisplayColumns(COL_ExchangeRate).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
        tdbg.Splits(SPLIT0).DisplayColumns(COL_ConvertedAmount).Style.BackColor = Color.FromArgb(COLOR_BACKCOLOR)
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



    Private Sub tdbg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tdbg.KeyDown
        'If e.KeyCode = Keys.Enter Then
        '    If tdbg.Col = iLastCol Then
        '        HotKeyEnterGrid(tdbg, COL_RefDate, e)
        '    End If
        'End If

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
        HotKeyDownGrid(e, tdbg, COL_RefDate, 0, 1, True, True, True, COL_Description, "")

    End Sub

    Public Sub HotKeyF7Other(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid)
        Try
            If c1Grid.RowCount < 1 Then Exit Sub

            If c1Grid(c1Grid.Row, c1Grid.Col).ToString() = "" Then
                c1Grid.Columns(c1Grid.Col).Text = c1Grid(c1Grid.Row - 1, c1Grid.Col).ToString
                If c1Grid.Col = COL_ObjectTypeID Then
                    c1Grid.Columns(COL_ObjectID).Text = ""
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
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Sub CopyColumn(ByVal c1Grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal ColCopy As Integer, ByVal sValue As String)
        Dim sValue1 As String = ""
        Dim sValue2 As String = ""
        Dim sValue3 As String = ""
        Dim sValue4 As String = ""
        Dim Flag As DialogResult
        Flag = D99C0008.MsgCopyColumn()
        If ColCopy = COL_ObjectTypeID Then
            sValue1 = c1Grid.Columns(COL_ObjectTypeID).Text
            'sValue2 = ""
            'sValue3 = c1Grid.Columns(COL_OriginalAmount).Text
            'sValue4 = c1Grid.Columns(COL_ConvertedAmount).Text

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
                    If ColCopy = COL_ObjectTypeID Then
                        c1Grid(i, COL_ObjectTypeID) = sValue1
                        c1Grid(i, COL_ObjectID) = ""
                        c1Grid(i, COL_ObjectName) = ""
                        c1Grid(i, COL_VATNo) = ""
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
                If ColCopy = COL_ObjectTypeID Then
                    c1Grid(i, COL_ObjectTypeID) = sValue1
                    c1Grid(i, COL_ObjectID) = ""
                    c1Grid(i, COL_ObjectName) = ""
                    c1Grid(i, COL_VATNo) = ""
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
            If ColCopy = COL_ObjectTypeID Then
                c1Grid(0, COL_ObjectTypeID) = sValue1
                c1Grid(0, COL_ObjectID) = ""
                c1Grid(0, COL_ObjectName) = ""
                c1Grid(0, COL_VATNo) = ""
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
            Case COL_CreditAccountID
                If tdbg.Columns(COL_DebitAccountID).Text = "" Then
                    tdbg.Columns(COL_DebitAccountID).DropDown = Nothing
                    tdbg.Columns(COL_DebitAccountID).Text = _accountID
                End If
                If tdbg.Columns(COL_ExchangeRate).Text = "" Then
                    tdbg.Columns(COL_ExchangeRate).Text = tdbcCurrencyID.Columns("ExchangeRate").Text
                End If
            Case COL_ObjectTypeID
                tdbg.Columns(COL_ObjectTypeID).Text = tdbdObjectTypeID.Columns("ObjectTypeID").Text
                tdbg.Columns(COL_ObjectID).Text = ""
                tdbg.Columns(COL_ObjectName).Text = ""
                tdbg.Columns(COL_VATNo).Text = ""
            Case COL_ObjectID
                tdbg.Columns(COL_ObjectName).Text = tdbdObjectID.Columns("ObjectName").Text()
                tdbg.Columns(COL_VATNo).Text = tdbdObjectID.Columns("VATNo").Text()
        End Select
    End Sub

    Private Sub tdbg_BeforeColEdit(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColEditEventArgs) Handles tdbg.BeforeColEdit
        Select Case e.ColIndex
            Case COL_ObjectID
                LoadtdbdObjectID(tdbg.Columns(COL_ObjectTypeID).Text)
        End Select
    End Sub

    Private Sub tdbg_BeforeColUpdate(ByVal sender As System.Object, ByVal e As C1.Win.C1TrueDBGrid.BeforeColUpdateEventArgs) Handles tdbg.BeforeColUpdate
        Select Case e.ColIndex
            
            Case COL_CreditAccountID
                If tdbg.Columns(COL_CreditAccountID).Text <> tdbdCreditAccountID.Columns("AccountID").Text Then
                    tdbg.Columns(COL_CreditAccountID).Text = ""
                End If
                'Case COL_ExchangeRate
                '    If Not IsNumeric(tdbg.Columns(COL_ExchangeRate).Text) Then e.Cancel = True
                'Case COL_OriginalAmount
                '    If Not IsNumeric(tdbg.Columns(COL_OriginalAmount).Text) Then e.Cancel = True
                'Case COL_ConvertedAmount
                '    If Not IsNumeric(tdbg.Columns(COL_ConvertedAmount).Text) Then e.Cancel = True
            Case COL_ObjectTypeID
                If tdbg.Columns(COL_ObjectTypeID).Text <> tdbdObjectTypeID.Columns("ObjectTypeID").Text Then
                    tdbg.Columns(COL_ObjectTypeID).Text = ""
                    tdbg.Columns(COL_ObjectID).Text = ""
                    tdbg.Columns(COL_ObjectName).Text = ""
                    tdbg.Columns(COL_VATNo).Text = ""
                End If
            Case COL_ObjectID
                If tdbg.Columns(COL_ObjectID).Text <> tdbdObjectID.Columns("ObjectID").Text Then
                    tdbg.Columns(COL_ObjectID).Text = ""
                    tdbg.Columns(COL_ObjectName).Text = ""
                    tdbg.Columns(COL_VATNo).Text = ""
                End If
            Case COL_VATTypeID
                If tdbg.Columns(COL_VATTypeID).Text <> tdbdVATTypeID.Columns("VATTypeID").Text Then
                    tdbg.Columns(COL_VATTypeID).Text = ""
                End If
            Case COL_VATGroupID
                If tdbg.Columns(COL_VATGroupID).Text <> tdbdVATGroupID.Columns("VATGroupID").Text Then
                    tdbg.Columns(COL_VATGroupID).Text = ""
                End If
            Case COL_BatchID
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
            Case COL_TransactionID
        End Select
    End Sub

    Private Sub tdbg_AfterColUpdate(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.ColEventArgs) Handles tdbg.AfterColUpdate
        If tdbg.Columns(COL_RefDate).Text = "" Then 'If tdbg.Columns(COL_RefDate).Text = "  /  /" Then
            tdbg.Columns(COL_RefDate).Text = Date.Today.ToShortDateString
        End If
        Select Case e.ColIndex
            Case COL_RefDate
                tdbg.Columns(e.ColIndex).Value = tdbg.Columns(e.ColIndex).Text
                If tdbg.Columns(COL_DebitAccountID).Text = "" Then
                    tdbg.Columns(COL_DebitAccountID).DropDown = Nothing
                    tdbg.Columns(COL_DebitAccountID).Text = _accountID
                End If
                If tdbg.Columns(COL_ExchangeRate).Text = "" Then
                    tdbg.Columns(COL_ExchangeRate).Text = tdbcCurrencyID.Columns("ExchangeRate").Text
                End If

            Case COL_RefNo
                If tdbg.Columns(COL_DebitAccountID).Text = "" Then
                    tdbg.Columns(COL_DebitAccountID).DropDown = Nothing
                    tdbg.Columns(COL_DebitAccountID).Text = _accountID
                End If
                If tdbg.Columns(COL_ExchangeRate).Text = "" Then
                    tdbg.Columns(COL_ExchangeRate).Text = tdbcCurrencyID.Columns("ExchangeRate").Text
                End If

            Case COL_Description
                If tdbg.Columns(COL_DebitAccountID).Text = "" Then
                    tdbg.Columns(COL_DebitAccountID).DropDown = Nothing
                    tdbg.Columns(COL_DebitAccountID).Text = _accountID
                End If
                If tdbg.Columns(COL_ExchangeRate).Text = "" Then
                    tdbg.Columns(COL_ExchangeRate).Text = tdbcCurrencyID.Columns("ExchangeRate").Text
                End If
            Case COL_CreditAccountID
               
            Case COL_ExchangeRate
                tdbg.Columns(COL_ExchangeRate).Text = SQLNumber(tdbg.Columns(COL_ExchangeRate).Text, DxxFormat.ExchangeRateDecimals)
                If tdbg.Columns(COL_DebitAccountID).Text = "" Then
                    tdbg.Columns(COL_DebitAccountID).DropDown = Nothing
                    tdbg.Columns(COL_DebitAccountID).Text = _accountID
                End If
                CalcuteConvertedAmount()
                'tdbg.Columns(COL_ConvertedAmount).Text = SQLNumber(tdbg.Columns(COL_ConvertedAmount).Text, D02Format.ConvertedAmount)
            Case COL_OriginalAmount
                tdbg.Columns(COL_OriginalAmount).Text = SQLNumber(tdbg.Columns(COL_OriginalAmount).Text, DxxFormat.DecimalPlaces)
                If tdbg.Columns(COL_DebitAccountID).Text = "" Then
                    tdbg.Columns(COL_DebitAccountID).DropDown = Nothing
                    tdbg.Columns(COL_DebitAccountID).Text = _accountID
                End If
                CalcuteConvertedAmount()
                'tdbg.Columns(COL_ConvertedAmount).Text = SQLNumber(tdbg.Columns(COL_ConvertedAmount).Text, D02Format.ConvertedAmount)
            Case COL_ConvertedAmount
                tdbg.Columns(COL_ConvertedAmount).Text = SQLNumber(tdbg.Columns(COL_ConvertedAmount).Text, DxxFormat.D90_ConvertedDecimals)
        End Select
    End Sub

    Private Sub tdbcCurrencyID_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tdbcCurrencyID.SelectedValueChanged
        GetExchangeRate()
        If tdbg.RowCount > 0 Then
            tdbg(tdbg.Row, COL_ExchangeRate) = tdbcCurrencyID.Columns("ExchangeRate").Text
            CalcuteConvertedAmount()
        End If
    End Sub

    Private Sub CalcuteConvertedAmount()
        Dim dExchangeRate As Double = 0
        Dim dOriginalAmount As Double = 0
        Dim dConvertedAmount As Double
        If tdbg.Columns(COL_ExchangeRate).Text <> "" And tdbg.Columns(COL_OriginalAmount).Text <> "" Then
            dExchangeRate = CDbl(tdbg.Columns(COL_ExchangeRate).Text)
            dOriginalAmount = CDbl(tdbg.Columns(COL_OriginalAmount).Text)
            If tdbcCurrencyID.Columns("Operator").Text <> "" Then
                If CInt(tdbcCurrencyID.Columns("Operator").Text) = 0 Then
                    dConvertedAmount = dExchangeRate * dOriginalAmount
                    tdbg.Columns(COL_ConvertedAmount).Text = dConvertedAmount.ToString
                Else
                    If dExchangeRate <> 0 Then
                        dConvertedAmount = dOriginalAmount / dExchangeRate
                        tdbg.Columns(COL_ConvertedAmount).Text = SQLNumber(dConvertedAmount.ToString, DxxFormat.D90_ConvertedDecimals)

                    Else
                        D99C0008.MsgL3(rl3("Nguyen_te_khong_hop_le"))
                        Exit Sub
                    End If
                End If

            End If
        End If
    End Sub

    Private Sub GetExchangeRate()
        Dim sSQL As String = ""
        sSQL = SQLStoreD91P0010()
        dtExchangeRate = ReturnDataTable(sSQL)
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
        sSQL &= SQLString(tdbcCurrencyID.Columns("CurrencyID").Text) & COMMA 'CurrencyID, varchar[20], NOT NULL
        'If tdbg.Columns(COL_RefDate).Text = "  /  /" Then
        '    sSQL &= SQLDateSave("") 'ExDate, datetime, NOT NULL
        'Else
        sSQL &= SQLDateSave(tdbg.Columns(COL_RefDate).Text) 'ExDate, datetime, NOT NULL
        'End If
        Return sSQL
    End Function

    Private Sub btnSetNewKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GetNewVoucherNo(tdbcVoucherTypeID, txtVoucherNo)
    End Sub

    Private Sub SetBackColorObligatory()
        txtVoucherNo.BackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcVoucherTypeID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
        tdbcCurrencyID.EditorBackColor = COLOR_BACKCOLOROBLIGATORY
    End Sub

    Private Function AllowSave() As Boolean
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
        If tdbcCurrencyID.Text.Trim = "" Then
            D99C0008.MsgNotYetChoose(rl3("Loai_tien"))
            tdbcCurrencyID.Focus()
            Return False
        End If
        If tdbg.RowCount <= 0 Then
            D99C0008.MsgNoDataInGrid()
            tdbg.Focus()
            Return False
        End If
        For i As Integer = 0 To tdbg.RowCount - 1
            If tdbg(i, COL_DebitAccountID).ToString = "" Then
                D99C0008.MsgNotYetEnter(rl3("Tai_khoan_no"))
                tdbg.SplitIndex = SPLIT0
                tdbg.Col = COL_DebitAccountID
                tdbg.Bookmark = i
                tdbg.Focus()
                Return False
            End If
            If tdbg(i, COL_CreditAccountID).ToString = "" Then
                D99C0008.MsgNotYetEnter(rl3("Tai_khoan_co"))
                tdbg.SplitIndex = SPLIT0
                tdbg.Col = COL_CreditAccountID
                tdbg.Bookmark = i
                tdbg.Focus()
                Return False
            End If
            If tdbg(i, COL_OriginalAmount).ToString = "" Then
                D99C0008.MsgNotYetEnter(rl3("Nguyen_te"))
                tdbg.SplitIndex = SPLIT0
                tdbg.Col = COL_OriginalAmount
                tdbg.Bookmark = i
                tdbg.Focus()
                Return False
            End If
            If tdbg(i, COL_ExchangeRate).ToString <> "" Then
                If CDbl(tdbg(i, COL_ExchangeRate).ToString) > MaxMoney Then
                    D99C0008.MsgL3(rl3("Ty_gia_qua_lon"))
                    tdbg.SplitIndex = SPLIT0
                    tdbg.Col = COL_ExchangeRate
                    tdbg.Bookmark = i
                    tdbg.Focus()
                    Return False
                End If

            End If
            If tdbg(i, COL_OriginalAmount).ToString <> "" Then
                If CDbl(tdbg(i, COL_OriginalAmount).ToString) > MaxMoney Then
                    D99C0008.MsgL3(rl3("Nguyen_te_qua_lon"))
                    tdbg.SplitIndex = SPLIT0
                    tdbg.Col = COL_OriginalAmount
                    tdbg.Bookmark = i
                    tdbg.Focus()
                    Return False
                End If

            End If
            If tdbg(i, COL_ObjectTypeID).ToString <> "" Then
                If tdbg(i, COL_ObjectID).ToString = "" Then
                    D99C0008.MsgNotYetEnter(rl3("Ma_doi_tuong"))
                    tdbg.SplitIndex = SPLIT0
                    tdbg.Col = COL_ObjectID
                    tdbg.Bookmark = i
                    tdbg.Focus()
                    Return False
                End If

            End If
            If tdbg(i, COL_ObjectID).ToString <> "" Then
                If tdbg(i, COL_ObjectTypeID).ToString = "" Then
                    D99C0008.MsgNotYetEnter(rl3("Doi_tuong"))
                    tdbg.SplitIndex = SPLIT0
                    tdbg.Col = COL_ObjectTypeID
                    tdbg.Bookmark = i
                    tdbg.Focus()
                    Return False
                End If

            End If
        Next
        Return True
    End Function

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
                _batchID = CreateIGE("D02T0012", "BatchID", "02", "BB", gsStringKey) 'BI
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

                sSQL.Append(SQLInsertD02T0018())
                sSQL.Append(vbCrLf)
                sSQL.Append(SQLInsertD02T0012s())

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
                sSQL.Append(SQLDeleteD02T0018)
                sSQL.Append(vbCrLf)
                sSQL.Append(SQLDeleteD02T0012)
                sSQL.Append(vbCrLf)
                sSQL.Append(SQLInsertD02T0018())
                sSQL.Append(vbCrLf)
                sSQL.Append(SQLInsertD02T0012s())
        End Select

        Dim bRunSQL As Boolean = ExecuteSQL(sSQL.ToString)
        Me.Cursor = Cursors.Default

        If bRunSQL Then
            SaveOK()
            gbSavedOK = True
            btnClose.Enabled = True
            _keyID = _batchID
            Select Case _FormState
                Case EnumFormState.FormAdd
                    'Kiểm tra và thiết lập Auditlog 
                    If _byAudit = 1 Then
                        'ExecuteAuditLog(_auditCode, "01", _cipNo, _cipName)
                        Lemon3.D91.RunAuditLog("02", _auditCode, "01", _cipNo, _cipName)
                    End If
                    btnNext.Enabled = True
                    btnNext.Focus()
                Case EnumFormState.FormEdit
                    'Kiểm tra và thiết lập Auditlog 
                    If _byAudit = 1 Then
                        'ExecuteAuditLog(_auditCode, "02", _cipNo, _cipName)
                        Lemon3.D91.RunAuditLog("02", _auditCode, "02", _cipNo, _cipName)
                    End If
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

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0018
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 09/10/2007 10:35:29
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLInsertD02T0018() As StringBuilder


        Dim sSQL As New StringBuilder
        sSQL.Append("Insert Into D02T0018(")
        sSQL.Append("BatchID, VoucherNo, VoucherTypeID, VoucherDate, Notes,NotesU, ")
        sSQL.Append("DivisionID, TranMonth, TranYear, CreateDate, CreateUserID, ")
        sSQL.Append("LastmodifyDate, LastmodifyUserID")
        sSQL.Append(") Values(")
        sSQL.Append(SQLString(_batchID) & COMMA) 'BatchID [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString(txtVoucherNo.Text) & COMMA) 'VoucherNo [KEY], varchar[20], NOT NULL
        sSQL.Append(SQLString(tdbcVoucherTypeID.Text) & COMMA) 'VoucherTypeID, varchar[20], NULL
        sSQL.Append(SQLDateSave(c1dateVoucherDate.Value) & COMMA) 'VoucherDate, datetime, NULL
        sSQL.Append("'',N''" & COMMA) 'Notes, varchar[250], NULL
        sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID, varchar[20], NULL
        sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NULL
        sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, smallint, NULL
        If _FormState = EnumFormState.FormAdd Then
            sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
            sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
        Else
            sSQL.Append(SQLDateTimeSave(createDate) & COMMA) 'CreateDate, datetime, NULL
            sSQL.Append(SQLString(createUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
        End If
        sSQL.Append("GetDate()" & COMMA) 'LastmodifyDate, datetime, NULL
        sSQL.Append(SQLString(gsUserID)) 'LastmodifyUserID, varchar[20], NULL
        sSQL.Append(")")

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLInsertD02T0012s
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 11/09/2007 11:24:08
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
                sTransactionID = CreateIGEs("D02T0012", "TransactionID", "02", "TB", gsStringKey, sTransactionID, iCountIGE) 'TI
                tdbg(i, COL_TransactionID) = sTransactionID
            End If

            sSQL.Append("Insert Into D02T0012(")
            sSQL.Append("TransactionID, DivisionID, ModuleID,")
            sSQL.Append("VoucherTypeID, VoucherNo, VoucherDate, TranMonth, TranYear, ")
            sSQL.Append("DescriptionU, CurrencyID, ExchangeRate, DebitAccountID, ")
            sSQL.Append("CreditAccountID, OriginalAmount,  ConvertedAmount, Status,")
            sSQL.Append("RefNo, RefDate, Disabled, CreateUserID, CreateDate, ")
            sSQL.Append("LastModifyUserID, LastModifyDate, SeriNo, ObjectTypeID, ObjectID, ")
            sSQL.Append("BatchID, ObjectNameU, VATNo, CipID, VATGroupID, VATTypeID,Ana01ID, Ana02ID, ")
            sSQL.Append("Ana03ID,Ana04ID, Ana05ID, Ana06ID, Ana07ID, Ana08ID, Ana09ID, Ana10ID, ")
            sSQL.Append("Notes,NotesU, Posted, SourceID, Internal")
            sSQL.Append(") Values(")
            sSQL.Append(SQLString(tdbg(i, COL_TransactionID)) & COMMA) 'TransactionID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(gsDivisionID) & COMMA) 'DivisionID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString("02") & COMMA) 'ModuleID [KEY], varchar[20], NOT NULL
            sSQL.Append(SQLString(tdbcVoucherTypeID.Text) & COMMA) 'VoucherTypeID, varchar[20], NULL
            sSQL.Append(SQLString(txtVoucherNo.Text) & COMMA) 'VoucherNo, varchar[20], NULL
            sSQL.Append(SQLDateSave(c1dateVoucherDate.Value) & COMMA) 'VoucherDate, datetime, NULL
            sSQL.Append(SQLNumber(giTranMonth) & COMMA) 'TranMonth, tinyint, NULL
            sSQL.Append(SQLNumber(giTranYear) & COMMA) 'TranYear, smallint, NULL
            sSQL.Append(SQLStringUnicode(tdbg(i, COL_Description), gbUnicode, True) & COMMA) 'Description, varchar[250], NULL
            sSQL.Append(SQLString(tdbcCurrencyID.Text) & COMMA) 'CurrencyID, varchar[20], NOT NULL
            sSQL.Append(SQLMoney(tdbg(i, COL_ExchangeRate), DxxFormat.ExchangeRateDecimals) & COMMA) 'ExchangeRate, money, NOT NULL
            sSQL.Append(SQLString(tdbg(i, COL_DebitAccountID)) & COMMA) 'DebitAccountID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_CreditAccountID)) & COMMA) 'CreditAccountID, varchar[20], NULL

            sSQL.Append(SQLMoney(tdbg(i, COL_OriginalAmount), DxxFormat.DecimalPlaces) & COMMA) 'OriginalAmount, money, NULL
            sSQL.Append(SQLMoney(tdbg(i, COL_ConvertedAmount), DxxFormat.D90_ConvertedDecimals) & COMMA) 'ConvertedAmount, money, NULL

            sSQL.Append(SQLNumber(0) & COMMA) 'Status, tinyint, NOT NULL
            sSQL.Append(SQLString(tdbg(i, COL_RefNo)) & COMMA) 'RefNo, varchar[20], NULL
            'If tdbg(i, COL_RefDate).ToString = " /  /" Then
            '    sSQL.Append(SQLDateSave(Date.Today) & COMMA) 'RefDate, datetime, NULL
            'Else
            sSQL.Append(SQLDateSave(tdbg(i, COL_RefDate)) & COMMA) 'RefDate, datetime, NULL
            'End If

            sSQL.Append(SQLNumber(0) & COMMA) 'Disabled, bit, NOT NULL
            If _FormState <> EnumFormState.FormAdd Then
                sSQL.Append(SQLString(createUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
                sSQL.Append(SQLDateTimeSave(createDate) & COMMA) 'CreateDate, datetime, NULL
            Else
                sSQL.Append(SQLString(gsUserID) & COMMA) 'CreateUserID, varchar[20], NOT NULL
                sSQL.Append("GetDate()" & COMMA) 'CreateDate, datetime, NULL
            End If

            sSQL.Append(SQLString(gsUserID) & COMMA) 'LastModifyUserID, varchar[20], NOT NULL
            sSQL.Append("GetDate()" & COMMA) 'LastModifyDate, datetime, NULL
            sSQL.Append(SQLString(tdbg(i, COL_SeriNo)) & COMMA) 'SeriNo, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_ObjectTypeID)) & COMMA) 'ObjectTypeID, varchar[20], NULL
            sSQL.Append(SQLString(tdbg(i, COL_ObjectID)) & COMMA) 'ObjectID, varchar[20], NULL
            sSQL.Append(SQLString(_batchID) & COMMA) 'BatchID, varchar[20], NULL
            sSQL.Append(SQLStringUnicode(tdbg(i, COL_ObjectName), gbUnicode, True) & COMMA) 'ObjectName, varchar[250], NULL
            sSQL.Append(SQLString(tdbg(i, COL_VATNo)) & COMMA) 'VATNo, varchar[20], NULL
            sSQL.Append(SQLString(_cipID) & COMMA) 'CipID, varchar[20], NULL
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
            sSQL.Append("'',N''" & COMMA) 'Notes, varchar[250], NULL
            sSQL.Append(SQLNumber(1) & COMMA) 'Posted, tinyint, NOT NULL
            sSQL.Append(SQLString("") & COMMA) 'SourceID, varchar[20], NULL
            sSQL.Append(SQLNumber(1)) 'Internal, tinyint, NOT NULL
            sSQL.Append(")")

            sRet.Append(sSQL.ToString & vbCrLf)
            sSQL.Remove(0, sSQL.Length)
        Next
        Return sRet
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T0018
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 09/10/2007 10:54:26
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T0018() As String
        Dim sSQL As String = ""
        sSQL &= "Delete From D02T0018"
        sSQL &= " Where "
        sSQL &= "BatchID = " & SQLString(_batchID)

        Return sSQL
    End Function

    '#---------------------------------------------------------------------------------------------------
    '# Title: SQLDeleteD02T0012
    '# Created User: Trần Thị ÁiTrâm
    '# Created Date: 09/10/2007 10:55:07
    '# Modified User: 
    '# Modified Date: 
    '# Description: 
    '#---------------------------------------------------------------------------------------------------
    Private Function SQLDeleteD02T0012() As String
        Dim sSQL As String = ""
        sSQL &= "Delete From D02T0012"
        sSQL &= " Where "
        sSQL &= "BatchID = " & SQLString(_batchID)
        Return sSQL
    End Function

    Private Sub tdbg_RowColChange(ByVal sender As Object, ByVal e As C1.Win.C1TrueDBGrid.RowColChangeEventArgs) Handles tdbg.RowColChange
  If e IsNot Nothing AndAlso e.LastRow = -1 Then Exit Sub
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

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        btnSave.Enabled = True
        btnNext.Enabled = False
        LoadAddNew()
        tdbcVoucherTypeID.Text = ""
        txtVoucherNo.Text = ""
        tdbcCurrencyID.Text = ""
        tdbcVoucherTypeID.Focus()
    End Sub

    Private Sub Loadlanguage()
        '================================================================ 
        Me.Text = rl3("Cap_nhat_chi_phi_XDCB_-_D02F1004") & UnicodeCaption(gbUnicode) 'CËp nhËt chi phÛ XDCB - D02F1004
        '================================================================ 
        lblVoucherTypeID.Text = rl3("Loai_phieu") 'Loại phiếu
        lblteVoucherDate.Text = rl3("Ngay_hach_toan") 'Ngày hạch toán
        lblCurrencyID.Text = rl3("Loai_tien") 'Loại tiền
        lblVoucherNo.Text = rl3("So_phieu") 'Số phiếu
        '================================================================ 
        btnSave.Text = rl3("_Luu") '&Lưu
        btnNext.Text = rl3("Nhap__tiep") 'Nhập &tiếp
        btnClose.Text = rl3("Do_ng") 'Đó&ng
        btnHotKey.Text = rl3("_Phim_nong") '&Phím nóng
        '================================================================ 
        grp1.Text = rl3("Chung_tu_hach_toan") 'Chứng từ hạch toán
        '================================================================ 
        tdbcCurrencyID.Columns("CurrencyID").Caption = rl3("Ma") 'Mã
        tdbcCurrencyID.Columns("CurrencyName").Caption = rl3("Ten") 'Tên
        tdbcVoucherTypeID.Columns("VoucherTypeID").Caption = rl3("Ma") 'Mã
        tdbcVoucherTypeID.Columns("VoucherTypeName").Caption = rl3("Dien_giai") 'Diễn giải
        '================================================================ 
        tdbdObjectTypeID.Columns("ObjectTypeID").Caption = rl3("Ma") 'Mã
        tdbdObjectTypeID.Columns("ObjectTypeName").Caption = rl3("Dien_giai") 'Diễn giải
        tdbdObjectID.Columns("ObjectID").Caption = rl3("Ma") 'Mã
        tdbdObjectID.Columns("ObjectName").Caption = rl3("Ten") 'Tên
        tdbdAna01ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna01ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna02ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna02ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna03ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna03ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna04ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna04ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna05ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna05ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna06ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna06ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna07ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna07ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna08ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna08ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna09ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna09ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdAna10ID.Columns("AnaID").Caption = rl3("Ma") 'Mã khoản mục
        tdbdAna10ID.Columns("AnaName").Caption = rl3("Ten") 'Tên khoản mục
        tdbdDebitAccountID.Columns("AccountID").Caption = rl3("Ma") 'Mã
        tdbdDebitAccountID.Columns("AccoutName").Caption = rl3("Ten") 'Tên
        tdbdCreditAccountID.Columns("AccountID").Caption = rl3("Ma") 'Mã
        tdbdCreditAccountID.Columns("AccountName").Caption = rl3("Ten") 'Tên
        tdbdVATTypeID.Columns("VATTypeID").Caption = rl3("Ma") 'Mã
        tdbdVATTypeID.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbdVATGroupID.Columns("VATGroupID").Caption = rl3("Ma") 'Mã
        tdbdVATGroupID.Columns("VATGroupName").Caption = rl3("Ten") 'Tên

        '================================================================ 
        tdbg.Columns("RefDate").Caption = rl3("Ngay_hoa_don") 'Ngày hóa đơn
        tdbg.Columns("SeriNo").Caption = rl3("So_Seri") 'Số Sêri
        tdbg.Columns("RefNo").Caption = rl3("So_hoa_don") 'Số hóa đơn
        tdbg.Columns("Description").Caption = rl3("Dien_giai") 'Diễn giải
        tdbg.Columns("DebitAccountID").Caption = rl3("TK_no") 'rl3("Tai_khoan_no") 'Tài khoản nợ
        tdbg.Columns("CreditAccountID").Caption = rl3("TK_co") 'rl3("Tai_khoan_co") 'Tài khoản có
        tdbg.Columns("ExchangeRate").Caption = rl3("Ty_gia") 'Tỷ giá
        tdbg.Columns("OriginalAmount").Caption = rl3("Nguyen_te") 'Nguyên tệ
        tdbg.Columns("ConvertedAmount").Caption = rl3("Quy_doi") 'Qui đổi
        tdbg.Columns("ObjectTypeID").Caption = rl3("Loai_doi_tuong") 'Loại đối tượng
        tdbg.Columns("ObjectID").Caption = rl3("Ma_doi_tuong") 'Mã đối tượng
        tdbg.Columns("ObjectName").Caption = rl3("Ten_doi_tuong_GTGT") 'Tên đối tượng GTGT
        tdbg.Columns("VATNo").Caption = rl3("Ma_so_thue") 'Mã số thuế
        tdbg.Columns("VATTypeID").Caption = rl3("Loai_hoa_don") 'Loại hóa đơn
        tdbg.Columns("VATGroupID").Caption = rl3("Nhom_thue") 'Nhóm thuế
    End Sub


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

End Class