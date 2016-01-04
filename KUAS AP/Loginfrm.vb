Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Net
Imports System.IO
Imports System.IO.Compression
Imports System.Text.RegularExpressions
Imports KUAS_AP.SilentWebModule
Imports System.Security.Cryptography
Imports System.Security
Imports System.Xml
Imports System.ComponentModel
Imports HtmlAgilityPack
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Threading

Public Class Loginfrm
    Public XmlPath As String = "Configs.xml"
    Public cookies As New CookieContainer()
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                    (ByVal hwnd As IntPtr, _
                     ByVal wMsg As Integer, _
                     ByVal wParam As IntPtr, _
                     ByVal lParam As Byte()) _
                     As Integer
    Public Const EM_SETCUEBANNER As Integer = &H1501
    Dim CheckUpdate As Thread
    Public Shared Function Encrypt(ByVal pToEncrypt As String, ByVal sKey As String) As String
        Dim des As New DESCryptoServiceProvider()
        Dim inputByteArray() As Byte
        inputByteArray = Encoding.Default.GetBytes(pToEncrypt)
        '建立加密對象的密鑰和偏移量
        '原文使用ASCIIEncoding.ASCII方法的GetBytes方法
        '使得輸入密碼必須輸入英文文本
        des.Key = ASCIIEncoding.ASCII.GetBytes(sKey)
        des.IV = ASCIIEncoding.ASCII.GetBytes(sKey)
        '寫二進制數組到加密流
        '(把內存流中的內容全部寫入)
        Dim ms As New System.IO.MemoryStream()
        Dim cs As New CryptoStream(ms, des.CreateEncryptor, CryptoStreamMode.Write)
        '寫二進制數組到加密流
        '(把內存流中的內容全部寫入)
        cs.Write(inputByteArray, 0, inputByteArray.Length)
        cs.FlushFinalBlock()

        '建立輸出字符串     
        Dim ret As New StringBuilder()
        Dim b As Byte
        For Each b In ms.ToArray()
            ret.AppendFormat("{0:X2}", b)
        Next

        Return ret.ToString()
    End Function

    '解密方法
    Public Shared Function Decrypt(ByVal pToDecrypt As String, ByVal sKey As String) As String
        Dim des As New DESCryptoServiceProvider()
        '把字符串放入byte數組
        Dim len As Integer
        len = pToDecrypt.Length / 2 - 1
        Dim inputByteArray(len) As Byte
        Dim x, i As Integer
        For x = 0 To len
            i = Convert.ToInt32(pToDecrypt.Substring(x * 2, 2), 16)
            inputByteArray(x) = CType(i, Byte)
        Next
        '建立加密對象的密鑰和偏移量，此值重要，不能修改
        des.Key = ASCIIEncoding.ASCII.GetBytes(sKey)
        des.IV = ASCIIEncoding.ASCII.GetBytes(sKey)
        Dim ms As New System.IO.MemoryStream()
        Dim cs As New CryptoStream(ms, des.CreateDecryptor, CryptoStreamMode.Write)
        cs.Write(inputByteArray, 0, inputByteArray.Length)
        cs.FlushFinalBlock()
        Return Encoding.Default.GetString(ms.ToArray)

    End Function
    Dim Course As Boolean
    Dim CourseUsername, CourseAccount, CoursePassword As String
    Private Sub StatusBtn_Click(sender As Object, e As EventArgs) Handles LoginBtn.Click
        Try
            User.Enabled = False
            Pwd.Enabled = False
            LoginBtn.Enabled = False
            RememberCheckBox.Enabled = False
            Me.Refresh()

            Dim userName As String = User.Text
            Dim password As String = Pwd.Text
            CourseAccount = userName
            CoursePassword = password

            Dim parameters As IDictionary(Of String, String) = New Dictionary(Of String, String)()
            parameters.Add("uid", userName)
            parameters.Add("pwd", password)
            Dim response As HttpWebResponse = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/perchk.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
            Dim reader As StreamReader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
            Dim respHTML As String = reader.ReadToEnd()

            parameters.Clear()
            response = HttpWebResponseUtility.CreateGetHttpResponse("http://140.127.113.231/kuas/f_head.jsp", Nothing, Nothing, cookies)
            reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
            respHTML = reader.ReadToEnd()

            Dim doc As New HtmlDocument()
            doc.LoadHtml(respHTML)
            Dim node As HtmlNode = doc.DocumentNode

            parameters.Clear()
            parameters.Add("UserName", userName)
            parameters.Add("Password", password)
            response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.109/Account/LogOn?ReturnUrl=%2f", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
            reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
            respHTML = reader.ReadToEnd()

            Try
                Me.Text = "KUAS AP (By Silent) @ " & WebUtility.HtmlDecode(node.SelectNodes("/html/body/div[1]/div/div[3]/span[3]")(0).InnerText)
            Catch ex As Exception
                MsgBox("登入失敗 , 請在嘗試一次 !", MsgBoxStyle.Critical)
                User.Enabled = True
                Pwd.Enabled = True
                LoginBtn.Enabled = True
                RememberCheckBox.Enabled = True
                Exit Sub
            End Try

            Dim configs As BindingList(Of Config) = Nothing
            configs = New BindingList(Of Config)()
            configs.Add(New Config() With {.Account = User.Text, .Pwd = Encrypt(Pwd.Text, "SilentKC"), .Remember = RememberCheckBox.Checked, .Manager = Guid.NewGuid})
            XmlSerialize.SerializeToXml("Configs.xml", configs)
            Dim Loginfrm As New AP_Frm(CourseAccount, CoursePassword, CourseUsername, Course, cookies)
            Loginfrm.Show()
            Me.Hide()
        Catch ex As Exception
            MsgBox("系統抓取異常 , 請稍後再試 :)", MsgBoxStyle.Critical, "Opps ! Something Error :(")
            User.Enabled = True
            Pwd.Enabled = True
            LoginBtn.Enabled = True
            RememberCheckBox.Enabled = True
        End Try
    End Sub
    Public Sub LoadSetting()
        If File.Exists(Application.StartupPath & "/" & XmlPath) Then
            Dim configs As BindingList(Of Config) = XmlSerialize.DeserializeFromXml(Of BindingList(Of Config))("Configs.xml")
            If Not configs.Item(0).Account = "" Then
                User.DataBindings.Add("Text", configs, "Account")
            End If
            If Not configs.Item(0).Pwd = "" Then
                Try
                    Pwd.DataBindings.Add("Text", configs, "Pwd")
                    Pwd.Text = Decrypt(Pwd.Text, "SilentKC")
                Catch ex As Exception

                End Try
            End If

            If configs.Item(0).Remember = True Then
                RememberCheckBox.Checked = True
            Else
                RememberCheckBox.Checked = False
                Pwd.Text = ""
            End If
        End If
    End Sub
    Private Sub Mainfrm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'SaveSetting()
    End Sub
    Sub CheckUpdateBackground()
        Try
            Dim response As HttpWebResponse = HttpWebResponseUtility.CreateGetHttpResponse("https://www.dropbox.com/s/zi6zb4sxoah1ov7/Update.txt?dl=1", Nothing, Nothing, Nothing)
            Dim reader As StreamReader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
            Dim respHTML As String = reader.ReadToEnd()
            Dim jsonDoc As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.JsonConvert.DeserializeObject(respHTML)
            If Not jsonDoc.Item("Version[0]").ToString() = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision Then
                If MsgBox("已有新版本 : " & jsonDoc.Item("Version[0]").ToString() & vbCrLf & "是否進行更新 ?", MsgBoxStyle.Information + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Dim myProcess As Process = System.Diagnostics.Process.Start(Application.StartupPath & "\" & "Auto Updater.exe")
                    End
                End If
            End If
            For i = 0 To Val(jsonDoc.Item("List")("Total").ToString) - 1
                If Not File.Exists(Application.StartupPath & "\" & jsonDoc.Item("List")("List[" & i & "]").ToString.Split("|")(0)) Then
                    MsgBox("遺失程式檔案 , 將進行自動下載 !", MsgBoxStyle.Exclamation)
                    Dim myProcess As Process = System.Diagnostics.Process.Start(Application.StartupPath & "\" & "Auto Updater.exe")
                    End
                End If
            Next

            If File.Exists(Application.StartupPath & "\" & "Auto Updater Ex.exe") Then
                Try
                    My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\" & "Auto Updater.exe")
                Catch ex As Exception

                End Try
                My.Computer.FileSystem.RenameFile(Application.StartupPath & "\" & "Auto Updater Ex.exe", "Auto Updater.exe")

                response = HttpWebResponseUtility.CreateGetHttpResponse("https://www.dropbox.com/s/3lrozgfig2n2yyx/Debug.txt?dl=1", Nothing, Nothing, Nothing)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                jsonDoc = Newtonsoft.Json.JsonConvert.DeserializeObject(respHTML)
                MsgBox("已更新至 v" & jsonDoc.Item("Version").ToString() & " !" & vbCrLf & vbCrLf & jsonDoc.Item("Update").ToString(), MsgBoxStyle.Information)
            End If
            LoginBtn.Enabled = True
        Catch ex As Exception
            MsgBox("目前無法檢查版本 !", MsgBoxStyle.Exclamation)
            LoginBtn.Enabled = True
        End Try
    End Sub
    Private Sub Mainfrm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Form.CheckForIllegalCrossThreadCalls = False
        LoadSetting()
        SendMessage(User.Handle, _
                     EM_SETCUEBANNER, _
                     IntPtr.Zero, _
                     System.Text.Encoding.Unicode.GetBytes("學號"))
        SendMessage(Pwd.Handle, _
                     EM_SETCUEBANNER, _
                     IntPtr.Zero, _
                     System.Text.Encoding.Unicode.GetBytes("密碼"))
        Try
            CheckUpdate = New Thread(AddressOf Me.CheckUpdateBackground)
            CheckUpdate.Start()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Loginfrm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Json.Close()
    End Sub
End Class
Namespace SilentWebModule
    ''' <summary>
    ''' 有關HTTP請求的模組
    ''' </summary>
    Public Class HttpWebResponseUtility
        Private Shared ReadOnly DefaultUserAgent As String = "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; WOW64; Trident/6.0; MAMIJS)"
        ''' <summary>
        ''' 創建GET方式的HTTP請求
        ''' </summary>
        ''' <param name="url">請求的URL</param>
        ''' <param name="timeout">請求的超時時間</param>
        ''' <param name="userAgent">請求的客戶端瀏覽器信息，可以為空</param>
        ''' <param name="cookies">隨同HTTP請求發送的Cookie信息，如果不需要身分驗證可以為空</param>
        ''' <returns></returns>
        Public Shared Function CreateGetHttpResponse(url As String, timeout As System.Nullable(Of Integer), userAgent As String, cookies As CookieContainer) As HttpWebResponse
            If String.IsNullOrEmpty(url) Then
                Throw New ArgumentNullException("url")
            End If
            Dim request As HttpWebRequest = TryCast(WebRequest.Create(url), HttpWebRequest)
            request.Method = "GET"
            request.KeepAlive = True
            request.UserAgent = DefaultUserAgent
            If Not String.IsNullOrEmpty(userAgent) Then
                request.UserAgent = userAgent
            End If
            If timeout.HasValue Then
                request.Timeout = timeout.Value
            End If
            If cookies IsNot Nothing Then
                request.CookieContainer = cookies
                'request.CookieContainer = New CookieContainer()
                'request.CookieContainer.Add(cookies)
            End If
            Return TryCast(request.GetResponse(), HttpWebResponse)
        End Function
        ''' <summary>
        ''' 創建POST方式的HTTP請求
        ''' </summary>
        ''' <param name="url">請求的URL</param>
        ''' <param name="parameters">隨同請求POST的參數名稱及參數值字典</param>
        ''' <param name="timeout">請求的超時時間</param>
        ''' <param name="userAgent">請求的客戶端瀏覽器信息，可以為空</param>
        ''' <param name="requestEncoding">發送HTTP請求時所用的編碼</param>
        ''' <param name="cookies">隨同HTTP請求發送的Cookie信息，如果不需要身分驗證可以為空</param>
        ''' <returns></returns>
        Public Shared Function CreatePostHttpResponse(url As String, parameters As IDictionary(Of String, String), timeout As System.Nullable(Of Integer), userAgent As String, requestEncoding As Encoding, cookies As CookieContainer) As HttpWebResponse
            If String.IsNullOrEmpty(url) Then
                Throw New ArgumentNullException("url")
            End If
            If requestEncoding Is Nothing Then
                Throw New ArgumentNullException("requestEncoding")
            End If
            Dim request As HttpWebRequest = Nothing
            '如果是發送HTTPS請求
            If url.StartsWith("https", StringComparison.OrdinalIgnoreCase) Then
                ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf CheckValidationResult)
                request = TryCast(WebRequest.Create(url), HttpWebRequest)
                request.ProtocolVersion = HttpVersion.Version10
            Else
                request = TryCast(WebRequest.Create(url), HttpWebRequest)
            End If
            request.Method = "POST"
            request.KeepAlive = True
            request.ContentType = "application/x-www-form-urlencoded"

            If url.Contains("http://140.127.113.155/Questionnaire/QuestionnaireInsert.aspx") Then
                request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
                request.ContentType = "application/x-www-form-urlencoded"
                request.Headers.Add("Accept-Encoding", "gzip, deflate")
                request.Headers.Add("Accept-Language", "zh-tw,en-us;q=0.7,en;q=0.3")
            End If

            If Not String.IsNullOrEmpty(userAgent) Then
                request.UserAgent = userAgent
            Else
                request.UserAgent = DefaultUserAgent
            End If

            If timeout.HasValue Then
                request.Timeout = timeout.Value
            End If
            If cookies IsNot Nothing Then
                request.CookieContainer = cookies
                'request.CookieContainer = New CookieContainer()
                'request.CookieContainer.Add(cookies)
            End If
            '如果需要POST數據
            If Not (parameters Is Nothing OrElse parameters.Count = 0) Then
                Dim buffer As New StringBuilder()
                Dim i As Integer = 0
                For Each key As String In parameters.Keys
                    If i > 0 Then
                        buffer.AppendFormat("&{0}={1}", key, parameters(key))
                    Else
                        buffer.AppendFormat("{0}={1}", key, parameters(key))
                    End If
                    i += 1
                Next
                Dim data As Byte() = requestEncoding.GetBytes(buffer.ToString())
                Using stream As Stream = request.GetRequestStream()
                    stream.Write(data, 0, data.Length)
                End Using
            End If

            Return TryCast(request.GetResponse(), HttpWebResponse)

        End Function

        Private Shared Function CheckValidationResult(sender As Object, certificate As X509Certificate, chain As X509Chain, errors As SslPolicyErrors) As Boolean
            Return True
            '總是接受
        End Function
    End Class
End Namespace

Namespace SystemAPI.Function.EncryptLibrary
    Public Class EncryptSHA
        ''' <summary>
        ''' 使用SHA加密訊息
        ''' </summary>
        ''' <param name="sourceMessage">原始資訊</param>
        ''' <param name="SHAType">SHA加密方式</param>
        ''' <returns>string</returns>
        Public Function Encrypt(sourceMessage As String, SHAType As EnumSHAType) As String
            If String.IsNullOrEmpty(sourceMessage) Then
                Return String.Empty
            End If

            '字串先轉成byte[]
            Dim Message As Byte() = Encoding.Unicode.GetBytes(sourceMessage)
            Dim HashImplement As HashAlgorithm = Nothing

            '選擇要使用的SHA加密方式
            Select Case SHAType
                Case EnumSHAType.SHA1
                    HashImplement = New SHA1Managed()
                    Exit Select
                Case EnumSHAType.SHA256
                    HashImplement = New SHA256Managed()
                    Exit Select
                Case EnumSHAType.SHA384
                    HashImplement = New SHA384Managed()
                    Exit Select
                Case EnumSHAType.SHA512
                    HashImplement = New SHA512Managed()
                    Exit Select
            End Select

            '取Hash值
            Dim HashValue As Byte() = HashImplement.ComputeHash(Message)

            '把byte[]轉成string後，再回傳
            Return BitConverter.ToString(HashValue).Replace("-", "").ToLower()

        End Function

        Public Enum EnumSHAType
            SHA1
            SHA256
            SHA384
            SHA512
        End Enum

    End Class
End Namespace