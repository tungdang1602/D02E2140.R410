''' <summary>
''' Các màn hình của exe con D91E0240: Tìm kiếm mở rộng
''' </summary>
Public Enum D91E0240Form
    ''' <summary>
    ''' D91F6010: Tìm kiếm theo Đối tượng, Mã hàng
    ''' D91F6020: Tìm kiếm Tiêu thức khi in báo cáo
    ''' D91F0402: Tìm kiếm theo Diễn giải
    ''' </summary>
    D91F6010 = 0
    D91F6020 = 1
    D91F0402 = 2
End Enum

Public Class D91E0240

    Private Const EXEMODULE As String = "D91"
    Private Const EXECHILD As String = "D91E0240"

    ''' <summary>
    ''' Khởi tạo exe con D91E0240
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
        D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "CodeTable", gbUnicode.ToString)
    End Sub

    ''' <summary>
    ''' Khởi tạo exe con D91E0240
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
    Public WriteOnly Property FormActive() As D91E0240Form
        Set(ByVal Value As D91E0240Form)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "Ctrl01", [Enum].GetName(GetType(D91E0240Form), Value))
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

    ''' <summary>
    ''' Màn hình goi
    ''' </summary>
    Public WriteOnly Property FormCall() As String
        Set(ByVal Value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID06", Value)
        End Set
    End Property

    ''' <summary>
    ''' 1 : Tìm kiếm Loại chứng trừ
    ''' 2 : Tìm kiếm Đối tượng
    ''' 3 : Tìm kiếm Hàng hóa
    ''' </summary>
    Public WriteOnly Property InListID() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID01", value)
        End Set
    End Property

    ''' <summary>
    ''' Điều kiện tìm kiếm, truyen vao mot chuoi tim kiem
    ''' Loại ĐT: Đối với Cbo Đối tượng thì truyền 
    ''' IsDxx =1: Đối với Mã hàng (InventoryID) 
    ''' </summary>
    Public WriteOnly Property InWhere() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID02", value)
        End Set
    End Property


    ''' <summary>
    ''' Câu SQL để load lên lưới khi tìm kiếm theo Tiêu thức của báo cáo
    ''' </summary>
    Public WriteOnly Property SQLSelection() As String
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

    ''' <summary>
    ''' Điều kiện tìm kiếm, truyen vao gia tri can tim, con field can tim kiem thi trong du lieu da do ra
    ''' Loại ĐT: Đối với Cbo Đối tượng thi chi can truyen gia tri cua doi tuong, con truong ObjectID da co roi 
    ''' </summary>
    Public WriteOnly Property InWhereValue() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ID05", value)
        End Set
    End Property

    ''' <summary>
    ''' ModeSelect
    ''' </summary>
    Public WriteOnly Property ModeSelect() As String
        Set(ByVal value As String)
            D99C0007.SaveOthersSetting(EXEMODULE, EXECHILD, "ModeSelect", value)
        End Set
    End Property

    ''' <summary>
    ''' Kết quả sau khi tìm kiếm
    ''' </summary>
    Public ReadOnly Property Output01() As String
        Get
            Return D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Output01", "")
        End Get
    End Property

    ''' <summary>
    ''' Kết quả sau khi tìm kiếm
    ''' </summary>
    Public ReadOnly Property Output02() As String
        Get
            Return D99C0007.GetOthersSetting(EXEMODULE, EXECHILD, "Output02", "")
        End Get
    End Property


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
        If Not ExistFile(My.Application.Info.DirectoryPath & "\" & EXECHILD & ".exe") Then Exit Sub
        Dim pInfo As New System.Diagnostics.ProcessStartInfo(My.Application.Info.DirectoryPath & "\" & EXECHILD & ".exe")
        pInfo.Arguments = "/DigiNet Corporation"
        pInfo.WindowStyle = ProcessWindowStyle.Normal
        Process.Start(pInfo)

    End Sub
End Class
