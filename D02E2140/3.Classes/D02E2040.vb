
Public Class D02E2040
    Public Enum D02E2040Form
        ''' <summary>
        ''' D02F0502: 
        ''' D02F0503: 
        ''' </summary>
        D02E0502
    End Enum
    Private Const EXEMODULE As String = "D02"
    Private Const EXECHILD As String = "D02E2040"

    ''' <summary>
    ''' Khởi tạo exe con D02E0230
    ''' </summary>
    ''' <param name="Server">Server kết nối đến hệ thống</param>
    ''' <param name="Database">Database kết nối đến hệ thống</param>
    ''' <param name="UserDatabaseID">User Database kết nối đến hệ thống</param>
    ''' <param name="Password">Password kết nối đến hệ thống</param>
    ''' <param name="UserID">User Lemon3 kết nối đến hệ thống</param>
    ''' <param name="Language">Ngôn ngữ sử dụng</param>
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal UserDatabaseID As String, ByVal Password As String, ByVal UserID As String, ByVal Language As String)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ServerName", Server, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "DBName", Database, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ConnectionUserID", UserDatabaseID, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Password", Password, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "UserID", UserID, CodeOption.lmCode)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Language", Language)
    End Sub

    ''' <summary>
    ''' Khởi tạo exe con D02E0230
    ''' </summary>
    ''' <param name="Server">Server kết nối đến hệ thống</param>
    ''' <param name="Database">Database kết nối đến hệ thống</param>
    ''' <param name="UserDatabaseID">User Database kết nối đến hệ thống</param>
    ''' <param name="Password">Password kết nối đến hệ thống</param>
    ''' <param name="UserID">User Lemon3 kết nối đến hệ thống</param>
    ''' <param name="Language">Ngôn ngữ sử dụng</param>
    ''' <param name="DivisionID">Đơn vị hiện tại</param>
    ''' <param name="TranMonth">Tháng kế toán hiện tại</param>
    ''' <param name="TranYear">Năm kế toán hiện tại</param>
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal UserDatabaseID As String, ByVal Password As String, ByVal UserID As String, ByVal Language As String, ByVal DivisionID As String, ByVal TranMonth As Integer, ByVal TranYear As Integer)
        Me.New(Server, Database, UserDatabaseID, Password, UserID, Language)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "DivisionID", DivisionID)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "TranMonth", TranMonth.ToString)
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "TranYear", TranYear.ToString)
    End Sub

    ''' <summary>
    ''' Màn hình cần hiển thị cho exe con
    ''' </summary>
    Public WriteOnly Property FormActive() As D02E0230Form
        Set(ByVal Value As D02E0230Form)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl01", [Enum].GetName(GetType(D02E0230Form), Value))
        End Set
    End Property

    ''' <summary>
    ''' Màn hình phân quyền
    ''' </summary>
    Public WriteOnly Property FormPermision() As String
        Set(ByVal Value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl03", Value)
        End Set
    End Property

    Public WriteOnly Property ID01() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID01", value)
        End Set
    End Property

    Public WriteOnly Property ID02() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID02", value)
        End Set
    End Property

    Public WriteOnly Property ID03() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID03", value)
        End Set
    End Property


    ''' <summary>
    ''' Truyền vào mã moduleID (Dxx)
    ''' VD: D90
    ''' </summary>
    Public WriteOnly Property ModuleID() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID04", value)
        End Set
    End Property


    Public WriteOnly Property ID05() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID05", value)
        End Set
    End Property


    ''' <summary>
    ''' Kết quả trả về
    ''' </summary>
    Public ReadOnly Property Output01() As String
        Get
            Return D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Output01", "")
        End Get
    End Property

    ''' <summary>
    ''' Kết quả trả về
    ''' </summary>
    Public ReadOnly Property Output02() As String
        Get
            Return D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Output02", "")
        End Get
    End Property

    '''' <summary>
    '''' Kết quả sau khi tìm kiếm
    '''' </summary>
    'Public ReadOnly Property Output02() As String
    '    Get
    '        Return D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Output02", "")
    '    End Get
    'End Property


    ''' <summary>
    ''' Kiểm tra tồn tại exe con không ?
    ''' </summary>
    Private Function ExistFile(ByVal Path As String) As Boolean
        If System.IO.File.Exists(Path) Then Return True
        If geLanguage = EnumLanguage.Vietnamese Then
            D99C0008.MsgL3("Không tồn tại file " & EXECHILD & ".exe")
        Else
            D99C0008.MsgL3("Not exist file " & EXECHILD & ".exe")
        End If
        Return False
    End Function

    ''' <summary>
    ''' Thực thi exe con
    ''' </summary>
    Public Sub Run()
        If Not ExistFile(gsApplicationSetup & "\" & EXECHILD & ".exe") Then Exit Sub
        Dim pInfo As New System.Diagnostics.ProcessStartInfo(gsApplicationSetup & "\" & EXECHILD & ".exe")
        pInfo.Arguments = "/DigiNet Corporation"
        pInfo.WindowStyle = ProcessWindowStyle.Normal
        Process.Start(pInfo)

    End Sub
End Class
