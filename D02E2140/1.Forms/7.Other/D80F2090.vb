Imports System.Text

Public Class D80F2090
    Private WithEvents backgroundWorker1 As System.ComponentModel.BackgroundWorker
    Private ChildName As String = "D80E0440"
    Dim exe As D80E0440
    Dim p As System.Diagnostics.Process

#Region "Property"

    Private _formActive As String
    Public WriteOnly Property FormActive() As String
        Set(ByVal Value As String)
            _formActive = Value
        End Set
    End Property

    Private _formState As EnumFormState
    Public WriteOnly Property FormState() As EnumFormState
        Set(ByVal Value As EnumFormState)
            _formState = Value
        End Set
    End Property

    Private _formPermission As String
    Public WriteOnly Property FormPermission() As String
        Set(ByVal Value As String)
            _formPermission = Value
        End Set
    End Property

    Private _moduleID As String
    Public WriteOnly Property ModuleID() As String
        Set(ByVal Value As String)
            _moduleID = Value
        End Set
    End Property

    Private _transTypeID As String
    Public WriteOnly Property TransTypeID() As String
        Set(ByVal Value As String)
            _transTypeID = Value
        End Set
    End Property

    Private _sFont As String = ""
    Public WriteOnly Property sFont() As String
        Set(ByVal Value As String)
            _sFont = Value
        End Set
    End Property

    Private _outPut01 As Boolean ' Kết quả trả về
    Public ReadOnly Property OutPut01() As Boolean
        Get
            Return _outPut01
        End Get
    End Property
#End Region

    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backgroundWorker1.DoWork
        'Tạo một process gắn với exe con, process này sẽ quan sát exe con.
        Dim p As System.Diagnostics.Process
        Try
            p = Process.GetProcessesByName(ChildName)(0)
            If p Is Nothing Then
                Exit Sub
            End If
            p.EnableRaisingEvents = True
            p.WaitForExit()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub FormLock_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Ẩn form trung gian
        Me.Size = New Size(0, 0)
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None

        '----Truyền tham số exe con------
        exe = New D80E0440(gsServer, gsCompanyID, gsConnectionUser, gsPassword, gsUserID, IIf(geLanguage = EnumLanguage.Vietnamese, "0", "10000").ToString, gsDivisionID, giTranMonth, giTranYear)
        exe.FormActive = D80E0440Form.D80F2090
        exe.FormPermission = _formPermission
        exe.ModuleID = _moduleID
        exe.TransTypeID = _transTypeID
        exe.sfont = _sFont 'VNI-UNICODE: - theo TL phân tích
        exe.Run()

        'Bắt đầu chạy cơ chế background
        backgroundWorker1 = New System.ComponentModel.BackgroundWorker
        backgroundWorker1.RunWorkerAsync()
    End Sub

    'sự kiện hoàn thành và dừng của Background
    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backgroundWorker1.RunWorkerCompleted
        _outPut01 = exe.Output01
        Me.Close()
    End Sub
End Class