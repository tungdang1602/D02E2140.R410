Imports System
Public Class D91F5558

    Private WithEvents backgroundWorker1 As System.ComponentModel.BackgroundWorker
    Private ChildName As String = "D91E0640"

    Private exe As D91E0640

    'Phân biệt các form được gọi trong class D91E0640
    Private _formName As String = ""
    Public WriteOnly Property FormName() As String
        Set(ByVal Value As String)
            _formName = Value
        End Set
    End Property

    Private _formPermission As String = ""
    Public WriteOnly Property FormPermission() As String
        Set(ByVal Value As String)
            _formPermission = Value
        End Set
    End Property

    Private _moduleID As String = ""
    Public WriteOnly Property ModuleID() As String
        Set(ByVal Value As String)
            _moduleID = Value
        End Set
    End Property

    Private _tableName As String = ""
    Public WriteOnly Property TableName() As String
        Set(ByVal Value As String)
            _tableName = Value
        End Set
    End Property

    Private _voucherID As String = ""
    Public WriteOnly Property VoucherID() As String
        Set(ByVal Value As String)
            _voucherID = Value
        End Set
    End Property

    Private _voucherNo As String = ""
    Public WriteOnly Property VoucherNo() As String
        Set(ByVal Value As String)
            _voucherNo = Value
        End Set
    End Property

    Private _mode As String = "0"
    Public WriteOnly Property Mode() As String 
        Set(ByVal Value As String )
            _mode = Value
        End Set
    End Property

    Private _keyID01 As String = ""
    Public WriteOnly Property KeyID01() As String
        Set(ByVal Value As String)
            _keyID01 = Value
        End Set
    End Property

    Private _keyID02 As String = ""
    Public WriteOnly Property KeyID02() As String
        Set(ByVal Value As String)
            _keyID02 = Value
        End Set
    End Property

    Private _keyID03 As String = ""
    Public WriteOnly Property KeyID03() As String
        Set(ByVal Value As String)
            _keyID03 = Value
        End Set
    End Property

    Private _keyID04 As String = ""
    Public WriteOnly Property KeyID04() As String
        Set(ByVal Value As String)
            _keyID04 = Value
        End Set
    End Property

    Private _keyID05 As String = ""
    Public WriteOnly Property KeyID05() As String
        Set(ByVal Value As String)
            _keyID05 = Value
        End Set
    End Property

    Private _iD01 As String = ""
    Public WriteOnly Property ID01() As String
        Set(ByVal Value As String)
            _iD01 = Value
        End Set
    End Property

    Private _iD02 As String = ""
    Public WriteOnly Property ID02() As String
        Set(ByVal Value As String)
            _iD02 = Value
        End Set
    End Property

    Private _iD03 As String = "0"
    Public WriteOnly Property ID03() As String
        Set(ByVal Value As String)
            _iD03 = Value
        End Set
    End Property

    Private _iD04 As String = ""
    Public WriteOnly Property ID04() As String
        Set(ByVal Value As String)
            _iD04 = Value
        End Set
    End Property

    Private _iD05 As String = ""
    Public WriteOnly Property ID05() As String
        Set(ByVal Value As String)
            _iD05 = Value
        End Set
    End Property

    'Trả về giá trị gia tri True, False; la trang thai da luu chua
    Private _Output01 As String
    Public ReadOnly Property Output01() As String
        Get
            Return _Output01
        End Get
    End Property

    'Trả về giá trị Số phiếu đã sửa từ form D91F5558
    Private _Output02 As String
    Public ReadOnly Property Output02() As String
        Get
            Return _Output02
        End Get
    End Property

    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backgroundWorker1.DoWork
        'Tạo một process gắn với exe con, process này sẽ quan sát exe con.
        Dim p As System.Diagnostics.Process

        Try
            p = Process.GetProcessesByName(ChildName)(0)

            If p Is Nothing Then
                Exit Sub
            End If

            'Chờ đợi exe con tắt tiến trình 
            p.EnableRaisingEvents = True
            p.WaitForExit()

        Catch ex As Exception
            MsgBox(ex.Message & " - " & ex.Source)
        End Try
    End Sub

    Public Sub FormLock_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Ẩn form trung gian
        Me.Size = New Size(0, 0)
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None

        '----Truyền tham số exe con------
        exe = New D91E0640(gsServer, gsCompanyID, gsConnectionUser, gsPassword, gsUserID, IIf(geLanguage = EnumLanguage.Vietnamese, "0", "10000").ToString(), gsDivisionID, giTranMonth, giTranYear)

        If _formName = "D91F5559" Then ' Sửa số tăng tự động
            exe.FormPermission = _formPermission
            exe.FormActive = D91E0640Form.D91F5559
            exe.ModuleID = _moduleID
        ElseIf _formName = "D91F5558" Then ' Sửa số phiếu
            With exe
                .FormPermission = _formPermission
                .FormActive = D91E0640Form.D91F5558
                .ModuleID = _moduleID
                .TableName = _tableName
                .VoucherID = _voucherID
                .VoucherNo = _voucherNo
                .Mode = _mode
                .KeyID01 = _keyID01
                .KeyID02 = _keyID02
                .KeyID03 = _keyID03
                .KeyID04 = _keyID04
                .KeyID05 = _keyID05
            End With
        ElseIf _formName = "D91F1655" Then
            With exe
                .FormPermission = _formPermission
                .FormActive = D91E0640Form.D91F1655
                .ID01 = _iD01
                .ID02 = _iD02
                .ID03 = _iD03
                .ID04 = _iD04
                .ID05 = _iD05
            End With
        End If

        exe.Run()


        If _formName = "D91F1301" Then ' Cơ chế không đợi
            Me.Close()
        Else ' Cơ chế đợi
            'Bắt đầu chạy cơ chế background
            backgroundWorker1 = New System.ComponentModel.BackgroundWorker
            backgroundWorker1.RunWorkerAsync()
        End If
    End Sub

    'sự kiện hoàn thành và dừng của Background
    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundWorker1.RunWorkerCompleted
        _Output01 = exe.Output01
        _Output02 = exe.Output02
        Me.Close()
    End Sub

End Class
