Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Net
Imports System.IO
Imports System.IO.Compression
Imports System.Text.RegularExpressions
Imports Auto_Updater.SilentWebModule
Imports System.Security.Cryptography
Imports System.Security
Imports System.Xml
Imports System.ComponentModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Threading
Public Class AutoUpdater
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer
    Private Declare Function GetSystemMenu Lib "User32 " (ByVal hwnd As Integer, ByVal bRevert As Long) As Integer
    Private Declare Function RemoveMenu Lib "User32 " (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
    Private Declare Function DrawMenuBar Lib "User32 " (ByVal hwnd As Integer) As Integer
    Private Declare Function GetMenuItemCount Lib "User32 " (ByVal hMenu As Integer) As Integer
    Private Const MF_BYPOSITION = &H400&
    Private Const MF_DISABLED = &H2&
    Private Const WM_KEYDOWN = &H100
    Private Const WM_KEYUP = &H101

    Dim UpdateURL As String = "https://www.dropbox.com/s/zi6zb4sxoah1ov7/Update.txt?dl=1"
    Dim ProgramName As String = "KUAS AP.exe"

    Dim LostList As New ListBox
    Dim UpdaterList As New ListBox
    Dim DownloadManager As New Downloader(Nothing, Nothing)
    Dim TotalUpdateCount As Integer
    Dim NowUpdateCount As Integer = 1
    Dim DownloadType As String = "Lost"
    Dim NowUpdating As String
    Dim UpdateThread As Thread
    Dim MainVersion As Version = New Version("0.0.0.0")
    Private Sub closeX(ByVal wnd As Form)
        Dim hMenu As Integer, nCount As Integer
        hMenu = GetSystemMenu(wnd.Handle.ToInt32, 0)
        nCount = GetMenuItemCount(hMenu)
        'MsgBox(nCount)
        Call RemoveMenu(hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION)
        DrawMenuBar(Me.Handle.ToInt32)
    End Sub
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Dim SC_CLOSE As Integer = 61536
        Dim WM_SYSCOMMAND As Integer = 274

        If m.Msg = WM_SYSCOMMAND AndAlso m.WParam.ToInt32 = SC_CLOSE Then
            '用戶按了“X” 
            Exit Sub
        End If
        MyBase.WndProc(m)
    End Sub
    Private Sub AutoUpdater_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            UpdateThread.Abort()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub AutoUpdater_Load(sender As Object, e As EventArgs) Handles Me.Load
        closeX(Me)
        Form.CheckForIllegalCrossThreadCalls = False
        Try
            MainVersion = New Version(FileVersionInfo.GetVersionInfo(Application.StartupPath & "\" & ProgramName).FileVersion.ToString())
        Catch ex As Exception

        End Try
        Me.Text = "Auto Updater - " & My.Application.Info.Version.ToString
        UpdateThread = New Thread(AddressOf Me.CheckUpdateBackground)
        UpdateThread.Start()
    End Sub
    Sub KillProcess(Program As String)
        Try
            My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\" & Program)
        Catch ex As Exception

        End Try
    End Sub
    Sub Download(Item As String, Type As String)
        NowUpdating = Item.Split("|")(0)
        DownloadType = Type
        Me.Refresh()
        Timer.Enabled = True
        If Item.Split("|")(0) = "Auto Updater.exe" Then
            KillProcess(Application.StartupPath & "\" & "Auto Updater Ex.exe")
            DownloadManager = New Downloader(Item.Split("|")(1), Application.StartupPath & "\", "Auto Updater Ex.exe")
        ElseIf Item.Split("|")(0) = "Newtonsoft.Json.dll" Then
            KillProcess(Application.StartupPath & "\" & "Newtonsoft.Json_Ex.dll")
            DownloadManager = New Downloader(Item.Split("|")(1), Application.StartupPath & "\", "Newtonsoft.Json_Ex.dll")
        Else
            KillProcess(Item.Split("|")(0))
            DownloadManager = New Downloader(Item.Split("|")(1), Application.StartupPath & "\", Item.Split("|")(0))
        End If

        DownloadManager.TimerGate(Timer, 100) '將Timer1實體傳入DownloadManager中，並設定Interval為100毫秒
        DownloadManager.StartDownload() '開始進行下載

        While DownloadManager.CheckDownloading '如果正在下載中，就永遠執行迴圈
            Application.DoEvents() '停頓(多工)
        End While

        Select Case DownloadManager.GetStatus '取得狀態傳回值
            Case 0
                MsgBox("更新或下載遺失檔案失敗 ! 請稍後再試 :)", MsgBoxStyle.Critical)
                End
            Case 1
                UpdateProgressBar.Value = 100
                If Not TotalUpdateCount = NowUpdateCount Then
                    NowUpdateCount += 1
                End If
        End Select
    End Sub
    Sub CheckUpdateBackground()
        'Try
        ProgressLabel.Text = "Status : Connect Update Server."
        Dim parameters As IDictionary(Of String, String) = New Dictionary(Of String, String)()
        Dim response As HttpWebResponse = HttpWebResponseUtility.CreateGetHttpResponse(UpdateURL, Nothing, Nothing, Nothing)
        Dim reader As StreamReader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
        Dim respHTML As String = reader.ReadToEnd()
        Dim jsonDoc As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.JsonConvert.DeserializeObject(respHTML)
        ProgressLabel.Text = "Status : Get Lost List."
        Thread.Sleep(500)
        For i = 0 To Val(jsonDoc.Item("List")("Total").ToString) - 1
            If Not File.Exists(Application.StartupPath & "\" & jsonDoc.Item("List")("List[" & i & "]").ToString.Split("|")(0)) Then
                If Not LostList.Items.Contains(jsonDoc.Item("List")("List[" & i & "]").ToString) Then
                    LostList.Items.Add(jsonDoc.Item("List")("List[" & i & "]").ToString)
                End If
            End If
        Next
        ProgressLabel.Text = "Status : Get Update List."
        Thread.Sleep(500)
        For i = 0 To Val(jsonDoc.Item("Total").ToString) - 1
            If Not New Version(jsonDoc.Item("Version")("Version[" & i & "]")("Version").ToString) > MainVersion Then
                Exit For
            End If
            If New Version(jsonDoc.Item("Version")("Version[" & i & "]")("Version").ToString) > MainVersion Then
                For j = 0 To Val(jsonDoc.Item("Version")("Version[" & i & "]")("Total").ToString) - 1
                    If Not UpdaterList.Items.Contains(jsonDoc.Item("Version")("Version[" & i & "]")("Update[" & j & "]").ToString) And Not LostList.Items.Contains(jsonDoc.Item("Version")("Version[" & i & "]")("Update[" & j & "]").ToString) Then
                        UpdaterList.Items.Add(jsonDoc.Item("Version")("Version[" & i & "]")("Update[" & j & "]").ToString)
                    End If
                Next
            End If
        Next
        TotalUpdateCount = LostList.Items.Count + UpdaterList.Items.Count
        ProgressLabel.Text = "Status : Start Download Lost."
        Thread.Sleep(500)
        For i = 0 To LostList.Items.Count - 1
            Download(LostList.Items.Item(i), "Lost")
        Next
        ProgressLabel.Text = "Status : Start Download Update."
        Thread.Sleep(500)
        For i = 0 To UpdaterList.Items.Count - 1
            Me.Refresh()
            Download(UpdaterList.Items.Item(i), "Update")
            Me.Refresh()
            Thread.Sleep(800)
            Me.Refresh()
        Next

        ProgressLabel.Text = "Status : Downloading " & DownloadType & " - " & NowUpdating & " [" & NowUpdateCount & "/" & TotalUpdateCount & "]"
        Timer.Enabled = False
        MsgBox("更新及下載遺失檔案完成 , 即將重啟程式 !", MsgBoxStyle.Information)
        ProgressLabel.Text = "Status : Finish."
        Thread.Sleep(500)
        Dim myProcess As Process = System.Diagnostics.Process.Start(Application.StartupPath & "\KUAS AP.exe")
        End
        'Catch ex As Exception
        '    MsgBox("更新或下載遺失檔案失敗 ! 請稍後再試 :)", MsgBoxStyle.Critical)
        '    End
        'End Try
    End Sub

    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles Timer.Tick
        Me.Refresh()
        ProgressLabel.Text = "Status : Downloading " & DownloadType & " - " & NowUpdating & " [" & NowUpdateCount & "/" & TotalUpdateCount & "]"
        UpdateProgressBar.Value = DownloadManager.GetProgressPercentage
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

