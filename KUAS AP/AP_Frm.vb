Imports KUAS_AP.SilentWebModule
Imports System.Threading
Imports System.Xml
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.ComponentModel
Imports System.Xml.Serialization
Imports HtmlAgilityPack
Imports System.Text.RegularExpressions

Public Class AP_Frm
    Dim CatchClassThread As Thread
    Dim CatchScoreThread As Thread
    Dim CatchAbsentThread As Thread
    Dim CatchCreditThread As Thread
    Dim ReadCreditThread As Thread
    Dim AutoQThread As Thread
    Dim AbsentCheck As Boolean = False
    Dim ScoreCheck As Boolean = False
    Dim ClassCheck As Boolean = False

    Dim ClassReady As Boolean = False
    Dim ScoreReady As Boolean = False
    Dim AbsentReady As Boolean = False

    Dim ClassFirst As Boolean = False
    Dim ScoreFirst As Boolean = False
    Dim AbsentFirst As Boolean = False

    Dim ymsAbsent As String = Nothing
    Dim ymsClass As String = Nothing
    Dim ymsScore As String = Nothing

    Dim StudentGrade As Integer

    ' Provides generic collection data binding. Notice that we are specifying 
    ' the type of object that is allowed to be added to the BindingList
    Private moConfigBindingList As BindingList(Of Config)
    Dim cookies As New CookieContainer
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                        (ByVal hwnd As IntPtr, _
                         ByVal wMsg As Integer, _
                         ByVal wParam As IntPtr, _
                         ByVal lParam As Byte()) _
                         As Integer
    Public Const EM_SETCUEBANNER As Integer = &H1501
    Public Sub New(Account As String, Password As String, Username As String, Course As Boolean, LoginCookies As CookieContainer)
        ' 此為設計工具所需的呼叫。
        InitializeComponent()
        ' 在 InitializeComponent() 呼叫之後加入任何初始設定。
        Me.Text = "KUAS AP (By Silent) @ " & Account
        Me.Icon = Loginfrm.Icon
        cookies = LoginCookies
    End Sub

    Private Sub AP_Frm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        SaveSetting()
        Try
            CatchClassThread.Abort()
        Catch ex As Exception

        End Try
        Try
            CatchScoreThread.Abort()
        Catch ex As Exception

        End Try
        Try
            CatchAbsentThread.Abort()
        Catch ex As Exception

        End Try
        Try
            ReadCreditThread.Abort()
        Catch ex As Exception

        End Try
        Try
            CatchCreditThread.Abort()
        Catch ex As Exception

        End Try
        End
    End Sub
    Public Sub SaveSetting()
        Dim WhiteList As String = Nothing
        Dim BlackList As String = Nothing
        Dim _Class As String = Nothing

        Dim configs As BindingList(Of Config) = Nothing
        configs = New BindingList(Of Config)()
        Dim config = XmlSerialize.DeserializeFromXml(Of BindingList(Of Config))("Configs.xml")
        If config.Item(0).Remember = True Then
            configs.Add(New Config() With {.Account = userName, .Pwd = Loginfrm.Encrypt(password, "SilentKC"), .Remember = True, .Manager = Guid.NewGuid})
        Else
            configs.Add(New Config() With {.Account = userName, .Pwd = "", .Remember = False, .Manager = Guid.NewGuid})
        End If
        XmlSerialize.SerializeToXml("Configs.xml", configs)
    End Sub
    Dim CourseTime As DateTime
    Public Sub LoadSetting()
        If File.Exists(Application.StartupPath & "/" & Loginfrm.XmlPath) Then
            Dim configs = XmlSerialize.DeserializeFromXml(Of BindingList(Of Config))("Configs.xml")
            userName = configs.Item(0).Account
            password = Loginfrm.Decrypt(configs.Item(0).Pwd, "SilentKC")
        End If
    End Sub
    Private Sub AP_Frm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Form.CheckForIllegalCrossThreadCalls = False
        LoadSetting()
        'SendMessage(AddTeacherTextbox.Handle, _
        '             EM_SETCUEBANNER, _
        '             IntPtr.Zero, _
        '             System.Text.Encoding.Unicode.GetBytes("老師名稱/課程名稱/體育項目"))
    End Sub
    Dim FrmX As Integer
    Private Sub AP_Frm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Loginfrm.Hide()
        LoginAP()
        FrmX = Me.Size.Width
        'RefreshCourse()
    End Sub
    Dim CreditTotalCount As Integer = 0
    Dim CreditNowCount As Integer = 0
    Private Sub TabControl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl.SelectedIndexChanged
        If Me.Text = "KUAS AP (Silent)" Then
            Exit Sub
        End If
        If TabControl.SelectedIndex = 1 Then
            Me.Size = New Point(FrmX, Me.Size.Height)
            Try
                CatchScoreThread = New Thread(AddressOf Me.RefreshScore)
                CatchScoreThread.Start()
            Catch ex As Exception

            End Try
        ElseIf TabControl.SelectedIndex = 0 Then
            Try
                CatchClassThread = New Thread(AddressOf Me.RefreshClass)
                CatchClassThread.Start()
            Catch ex As Exception

            End Try
        ElseIf TabControl.SelectedIndex = 2 Then
            Me.Size = New Point(FrmX, Me.Size.Height)
            Try
                CatchAbsentThread = New Thread(AddressOf Me.RefreshAbsent)
                CatchAbsentThread.Start()
            Catch ex As Exception

            End Try
        ElseIf TabControl.SelectedIndex = 3 Then
            Me.Size = New Point(FrmX, Me.Size.Height)
            If StudentGrade = 0 Then
                MsgBox("您目前就讀的系所無法試算學分 !", MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            Try
                If CreditDoubleCheck = True Then
                    Exit Sub
                End If
                CreditDoubleCheck = True
                CreditListView.Items.Clear()
                Dim yms As Integer
                If Now.Month < 9 Then
                    yms = Now.Year - 1912
                Else
                    yms = Now.Year - 1911
                End If
                If Val(ScoreSemester.Items.Item(0).ToString.Split(",")(0)) > yms Then : yms = Val(ScoreSemester.Items.Item(0).ToString.Split(",")(0)) : End If
                For i = 0 To ScoreSemester.Items.Count - 1
                    If Val(yms - (StudentGrade - 1)) <= Val(ScoreSemester.Items.Item(i).ToString.Split(",")(0)) And Val(ScoreSemester.Items.Item(i).ToString.Split(",")(0)) <= yms Then
                        CreditTotalCount += 1
                        CatchCreditThread = New Thread(AddressOf RefreshCredit)
                        CatchCreditThread.Start(ScoreSemester.Items.Item(i).ToString)
                    Else
                        Exit For
                    End If
                Next
                ReadCreditThread = New Thread(AddressOf ReadCredit)
                ReadCreditThread.Start()
            Catch ex As Exception

            End Try
        ElseIf TabControl.SelectedIndex = 4 Then
            'If MsgBox("是否進行教學評量自動填寫 ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            '    If (7 < Month(Now) And Month(Now) < 12) Or (2 < Month(Now) And Month(Now) < 5) Then
            '        AutoQ()
            '    Else
            '        AutoFinalQ()
            '    End If
            'End If
            'TabControl.SelectedIndex = 0
        End If
    End Sub
    Dim CreditDoubleCheck As Boolean = False
    Sub ReadCredit()
        While (1)
            If CreditTotalCount = CreditNowCount Then
                CreditDoubleCheck = False
                Dim Post As String = Nothing
                If CreditError = True Then
                    MsgBox("暫時無法試算學分 , 請稍後再試 !", MsgBoxStyle.Critical)
                    RequiredCredit = 0
                    OptionalCredit = 0
                    General1 = False
                    General2 = False
                    General3 = False
                    General4 = False
                    General5 = False
                    GeneralSociety = False
                    GeneralHistorical = False
                    GeneralTechnical = False
                    CreditError = False
                    CreditCheck = False
                    Exit While
                End If
                Post += "必修學分 : " & RequiredCredit & vbCrLf
                Post += "選修學分 : " & OptionalCredit & vbCrLf & vbCrLf & "目前已修過通識 : " & vbCrLf
                If General1 = True Then
                    Post += "核心通識(一)" & vbCrLf
                End If
                If General2 = True Then
                    Post += "核心通識(二)" & vbCrLf
                End If
                If General3 = True Then
                    Post += "核心通識(三)" & vbCrLf
                End If
                If General4 = True Then
                    Post += "核心通識(四)" & vbCrLf
                End If
                If General5 = True Then
                    Post += "核心通識(五)" & vbCrLf
                End If
                If GeneralHistorical = True Then
                    Post += "延伸通識(人文)" & vbCrLf
                End If
                If GeneralSociety = True Then
                    Post += "延伸通識(社會)" & vbCrLf
                End If
                If GeneralTechnical = True Then
                    Post += "延伸通識(科技)" & vbCrLf
                End If
                CreditTotalCount = 0
                CreditNowCount = 0
                Post += vbCrLf & "※ 以上資料僅供參考"
                MsgBox(Post, MsgBoxStyle.Information)
                RequiredCredit = 0
                OptionalCredit = 0
                General1 = False
                General2 = False
                General3 = False
                General4 = False
                General5 = False
                GeneralSociety = False
                GeneralHistorical = False
                GeneralTechnical = False
                CreditError = False
                CreditCheck = False
                Exit While
            End If
            Thread.Sleep(500)
        End While
    End Sub
    Dim userName As String
    Dim password As String
    Sub AutoFinalQ()
        Try
            Dim parameters As IDictionary(Of String, String) = New Dictionary(Of String, String)()
            Dim response As HttpWebResponse
            Dim reader As StreamReader
            Dim respHTML As String

            parameters.Clear()
            parameters.Add("ServicePoint", "load")
            parameters.Add("moduleId", "kuas_questionnaire")
            response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.109/Extjs/GetJavascript", parameters, Nothing, Nothing, Encoding.UTF8, cookies)

            parameters.Clear()
            parameters.Add("", "")
            response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.109/Questionnaire/Process/BrowseCheck", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
            reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
            respHTML = reader.ReadToEnd()
            Dim jsonDoc As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.JsonConvert.DeserializeObject(respHTML)
            If (jsonDoc.Item("isOpen").ToString = False) Then
                MsgBox("教學評量目前尚未開放填寫 !", MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            Dim _UserId As String = jsonDoc.Item("message")

            Dim _count As Integer = 0
            While (True)
                response = HttpWebResponseUtility.CreateGetHttpResponse("http://140.127.113.155/Questionnaire/QuestionnaireInsert.aspx?UserId=" & _UserId, Nothing, Nothing, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                Dim doc As New HtmlDocument()
                doc.LoadHtml(respHTML)
                Dim node As HtmlNode = doc.DocumentNode

                Dim __FUCKVIEWSTATE As String = ""
                Dim __VIEWSTATE As String = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__VIEWSTATE""]").Attributes.Item("value").Value)
                Dim __VIEWSTATEGENERATOR As String = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__VIEWSTATEGENERATOR""]").Attributes.Item("value").Value)
                Dim __EVENTVALIDATION As String = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__EVENTVALIDATION""]").Attributes.Item("value").Value)
                Dim __EVENTTARGET As String = Nothing

                parameters.Clear()
                response.Close()
                For i = 0 To respHTML.Split(vbNewLine).Count - 1
                    Dim _line As String = respHTML.Split(vbNewLine)(i)
                    If Not _line = Nothing Then
                        If (_line.Contains("$ContentPlaceHolder") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            If _line.Contains("value") Then
                                startPos = _line.IndexOf("value=""", StringComparison.CurrentCultureIgnoreCase) + 7
                                endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                                _value = _line.Substring(startPos, endPos - startPos)
                            End If
                            parameters.Add(_name, _value)
                        ElseIf (_line.Contains("$ContentPlaceHolder") And _line.Contains("javascript") And __EVENTTARGET = Nothing) Then
                            Dim startPos As Integer, endPos As Integer
                            startPos = _line.IndexOf("__doPostBack(&#39;", StringComparison.CurrentCultureIgnoreCase) + 18
                            endPos = _line.IndexOf("&#39", startPos, StringComparison.CurrentCultureIgnoreCase)
                            __EVENTTARGET = _line.Substring(startPos, endPos - startPos)
                        End If
                    End If
                Next
                If __EVENTTARGET = Nothing Then
                    Exit While
                End If
                While (__VIEWSTATE.Length > 32766)
                    __FUCKVIEWSTATE &= System.Uri.EscapeDataString(__VIEWSTATE.Substring(0, 32766))
                    __VIEWSTATE = __VIEWSTATE.Substring(32766)
                End While
                __FUCKVIEWSTATE &= System.Uri.EscapeDataString(__VIEWSTATE)
                parameters.Add("__EVENTTARGET", System.Uri.EscapeDataString(__EVENTTARGET))
                parameters.Add("__EVENTARGUMENT", "")
                parameters.Add("__VIEWSTATE", __FUCKVIEWSTATE)
                parameters.Add("__VIEWSTATEGENERATOR", System.Uri.EscapeDataString(__VIEWSTATEGENERATOR))
                parameters.Add("__EVENTVALIDATION", System.Uri.EscapeDataString(__EVENTVALIDATION))

                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.155/Questionnaire/QuestionnaireInsert.aspx?UserId=" & _UserId, parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                doc.LoadHtml(respHTML)
                node = doc.DocumentNode
                __VIEWSTATE = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__VIEWSTATE""]").Attributes.Item("value").Value)
                __VIEWSTATEGENERATOR = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__VIEWSTATEGENERATOR""]").Attributes.Item("value").Value)
                __EVENTVALIDATION = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__EVENTVALIDATION""]").Attributes.Item("value").Value)

                __FUCKVIEWSTATE = ""
                While (__VIEWSTATE.Length > 32766)
                    __FUCKVIEWSTATE &= System.Uri.EscapeDataString(__VIEWSTATE.Substring(0, 32766))
                    __VIEWSTATE = __VIEWSTATE.Substring(32766)
                End While
                __FUCKVIEWSTATE &= System.Uri.EscapeDataString(__VIEWSTATE)

                response.Close()
                parameters.Clear()
                For i = 0 To respHTML.Split(vbNewLine).Count - 1
                    Dim _line As String = respHTML.Split(vbNewLine)(i)
                    If Not _line = Nothing Then
                        If (_line.Contains("$ContentPlaceHolder") And _line.Contains("HiddenField") And Not _line.Contains("gvInProgress") And Not _line.Contains("GridViewContent") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            If _line.Contains("value") Then
                                startPos = _line.IndexOf("value=""", StringComparison.CurrentCultureIgnoreCase) + 7
                                endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                                _value = _line.Substring(startPos, endPos - startPos)
                            End If
                            parameters.Add(_name, System.Uri.EscapeDataString(_value))
                        ElseIf (_line.Contains("$ContentPlaceHolder") And _line.Contains("gvInProgress") And Not _line.Contains("GridView") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            If _line.Contains("value") Then
                                startPos = _line.IndexOf("value=""", StringComparison.CurrentCultureIgnoreCase) + 7
                                endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                                _value = _line.Substring(startPos, endPos - startPos)
                            End If
                            parameters.Add(_name, System.Uri.EscapeDataString(_value))
                        ElseIf (_line.Contains("$ContentPlaceHolder") And Not _line.Contains("gvInProgress") And _line.Contains("GridView") And Not _line.Contains("Radio") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            If _line.Contains("value") Then
                                startPos = _line.IndexOf("value=""", StringComparison.CurrentCultureIgnoreCase) + 7
                                endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                                _value = _line.Substring(startPos, endPos - startPos)
                            End If
                            parameters.Add(_name, System.Uri.EscapeDataString(_value))
                        ElseIf (_line.Contains("$ContentPlaceHolder") And Not _line.Contains("gvInProgress") And _line.Contains("GridView") And _line.Contains("Radio") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            Dim RndNum As New Random()
                            'ctl03 必修
                            If (_line.Contains("ctl02")) Then
                                _value = ctl02
                            ElseIf (_line.Contains("ctl04")) Then
                                _value = ctl04
                            ElseIf (_line.Contains("ctl05")) Then
                                _value = ctl05
                            ElseIf (_line.Contains("ctl06")) Then
                                _value = ctl06
                            Else
                                _value = 1
                            End If
                            If Not parameters.ContainsKey(_name) Then
                                parameters.Add(_name, System.Uri.EscapeDataString(_value))
                            End If
                        End If
                    End If
                Next
                parameters.Add("__EVENTTARGET", "")
                parameters.Add("__EVENTARGUMENT", "")
                parameters.Add("__VIEWSTATE", __FUCKVIEWSTATE)
                parameters.Add("__VIEWSTATEGENERATOR", System.Uri.EscapeDataString(__VIEWSTATEGENERATOR))
                parameters.Add("__EVENTVALIDATION", System.Uri.EscapeDataString(__EVENTVALIDATION))
                parameters.Add("ctl00$ContentPlaceHolder1$ButtonNextPart", System.Uri.EscapeDataString("第二部分(Second+Part)"))

                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.155/Questionnaire/QuestionnaireInsert.aspx?UserId=" & _UserId, parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                doc.LoadHtml(respHTML)
                node = doc.DocumentNode
                __VIEWSTATE = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__VIEWSTATE""]").Attributes.Item("value").Value)
                __VIEWSTATEGENERATOR = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__VIEWSTATEGENERATOR""]").Attributes.Item("value").Value)
                __EVENTVALIDATION = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__EVENTVALIDATION""]").Attributes.Item("value").Value)

                __FUCKVIEWSTATE = ""
                While (__VIEWSTATE.Length > 32766)
                    __FUCKVIEWSTATE &= System.Uri.EscapeDataString(__VIEWSTATE.Substring(0, 32766))
                    __VIEWSTATE = __VIEWSTATE.Substring(32766)
                End While
                __FUCKVIEWSTATE &= System.Uri.EscapeDataString(__VIEWSTATE)

                response.Close()
                parameters.Clear()
                For i = 0 To respHTML.Split(vbNewLine).Count - 1
                    Dim _line As String = respHTML.Split(vbNewLine)(i)
                    If Not _line = Nothing Then
                        If (_line.Contains("$ContentPlaceHolder") And _line.Contains("HiddenField") And Not _line.Contains("gvInProgress") And Not _line.Contains("GridViewContent") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            If _line.Contains("value") Then
                                startPos = _line.IndexOf("value=""", StringComparison.CurrentCultureIgnoreCase) + 7
                                endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                                _value = _line.Substring(startPos, endPos - startPos)
                            End If
                            parameters.Add(_name, System.Uri.EscapeDataString(_value))
                        ElseIf (_line.Contains("$ContentPlaceHolder") And _line.Contains("gvInProgress") And Not _line.Contains("GridView") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            If _line.Contains("value") Then
                                startPos = _line.IndexOf("value=""", StringComparison.CurrentCultureIgnoreCase) + 7
                                endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                                _value = _line.Substring(startPos, endPos - startPos)
                            End If
                            parameters.Add(_name, System.Uri.EscapeDataString(_value))
                        ElseIf (_line.Contains("$ContentPlaceHolder") And Not _line.Contains("gvInProgress") And _line.Contains("GridView") And Not _line.Contains("Radio") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            If _line.Contains("value") Then
                                startPos = _line.IndexOf("value=""", StringComparison.CurrentCultureIgnoreCase) + 7
                                endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                                _value = _line.Substring(startPos, endPos - startPos)
                            End If
                            parameters.Add(_name, System.Uri.EscapeDataString(_value))
                        ElseIf (_line.Contains("$ContentPlaceHolder") And Not _line.Contains("gvInProgress") And _line.Contains("GridView") And _line.Contains("Radio") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            _value = 1
                            If Not parameters.ContainsKey(_name) Then
                                parameters.Add(_name, System.Uri.EscapeDataString(_value))
                            End If
                        End If
                    End If
                Next
                parameters.Add("__EVENTTARGET", "")
                parameters.Add("__EVENTARGUMENT", "")
                parameters.Add("__VIEWSTATE", __FUCKVIEWSTATE)
                parameters.Add("__VIEWSTATEGENERATOR", System.Uri.EscapeDataString(__VIEWSTATEGENERATOR))
                parameters.Add("__EVENTVALIDATION", System.Uri.EscapeDataString(__EVENTVALIDATION))
                parameters.Add("ctl00$ContentPlaceHolder1$Button1", System.Uri.EscapeDataString("送出(Send)"))
                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.155/Questionnaire/QuestionnaireInsert.aspx?UserId=" & _UserId, parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                response.Close()
                _count += 1
            End While
            If _count > 0 Then
                If MsgBox("填寫完畢 , 共填寫完" & _count & "個教學評量 , 是否要至選課系統查看填寫結果 !?", MsgBoxStyle.Information + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    System.Diagnostics.Process.Start("http://140.127.113.107/Account/LogOn?ReturnUrl=%2f")
                End If
            Else
                If MsgBox("沒有教學評量需要填寫 , 是否要至選課系統查看 !?", MsgBoxStyle.Information + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    System.Diagnostics.Process.Start("http://140.127.113.107/Account/LogOn?ReturnUrl=%2f")
                End If
            End If
        Catch ex As Exception
            MsgBox("發生異常 , 暫無法自動填寫教學評量 !", MsgBoxStyle.Critical)
        End Try
        Button1.Enabled = True
        GroupBox1.Enabled = True
        GroupBox2.Enabled = True
        GroupBox3.Enabled = True
        GroupBox4.Enabled = True
    End Sub
    Sub AutoQ()
        Try
            Dim parameters As IDictionary(Of String, String) = New Dictionary(Of String, String)()
            Dim response As HttpWebResponse
            Dim reader As StreamReader
            Dim respHTML As String

            parameters.Clear()
            parameters.Add("ServicePoint", "load")
            parameters.Add("moduleId", "kuas_questionnaire")
            response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.109/Extjs/GetJavascript", parameters, Nothing, Nothing, Encoding.UTF8, cookies)

            parameters.Clear()
            parameters.Add("", "")
            response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.109/Questionnaire/Process/BrowseCheck", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
            reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
            respHTML = reader.ReadToEnd()
            Dim jsonDoc As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.JsonConvert.DeserializeObject(respHTML)
            If (jsonDoc.Item("isOpen").ToString = False) Then
                MsgBox("教學評量目前尚未開放填寫 !", MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            Dim _UserId As String = jsonDoc.Item("message")

            Dim _count As Integer = 0
            While (True)
                response = HttpWebResponseUtility.CreateGetHttpResponse("http://140.127.113.155/Questionnaire/QuestionnaireInsert.aspx?UserId=" & _UserId, Nothing, Nothing, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                Dim doc As New HtmlDocument()
                doc.LoadHtml(respHTML)
                Dim node As HtmlNode = doc.DocumentNode
                Dim __VIEWSTATE As String = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__VIEWSTATE""]").Attributes.Item("value").Value)
                Dim __VIEWSTATEGENERATOR As String = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__VIEWSTATEGENERATOR""]").Attributes.Item("value").Value)
                Dim __EVENTVALIDATION As String = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__EVENTVALIDATION""]").Attributes.Item("value").Value)
                Dim __EVENTTARGET As String = Nothing
                parameters.Clear()
                response.Close()
                For i = 0 To respHTML.Split(vbNewLine).Count - 1
                    Dim _line As String = respHTML.Split(vbNewLine)(i)
                    If Not _line = Nothing Then
                        If (_line.Contains("$ContentPlaceHolder") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            If _line.Contains("value") Then
                                startPos = _line.IndexOf("value=""", StringComparison.CurrentCultureIgnoreCase) + 7
                                endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                                _value = _line.Substring(startPos, endPos - startPos)
                            End If
                            parameters.Add(_name, _value)
                        ElseIf (_line.Contains("$ContentPlaceHolder") And _line.Contains("javascript") And __EVENTTARGET = Nothing) Then
                            Dim startPos As Integer, endPos As Integer
                            startPos = _line.IndexOf("__doPostBack(&#39;", StringComparison.CurrentCultureIgnoreCase) + 18
                            endPos = _line.IndexOf("&#39", startPos, StringComparison.CurrentCultureIgnoreCase)
                            __EVENTTARGET = _line.Substring(startPos, endPos - startPos)
                        End If
                    End If
                Next
                If __EVENTTARGET = Nothing Then
                    Exit While
                End If
                parameters.Add("__EVENTTARGET", System.Uri.EscapeDataString(__EVENTTARGET))
                parameters.Add("__EVENTARGUMENT", "")
                parameters.Add("__VIEWSTATE", System.Uri.EscapeDataString(__VIEWSTATE))
                parameters.Add("__VIEWSTATEGENERATOR", System.Uri.EscapeDataString(__VIEWSTATEGENERATOR))
                parameters.Add("__EVENTVALIDATION", System.Uri.EscapeDataString(__EVENTVALIDATION))

                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.155/Questionnaire/QuestionnaireInsert.aspx?UserId=" & _UserId, parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                doc.LoadHtml(respHTML)
                node = doc.DocumentNode
                __VIEWSTATE = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__VIEWSTATE""]").Attributes.Item("value").Value)
                __VIEWSTATEGENERATOR = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__VIEWSTATEGENERATOR""]").Attributes.Item("value").Value)
                __EVENTVALIDATION = WebUtility.HtmlDecode(node.SelectSingleNode("//*[@id=""__EVENTVALIDATION""]").Attributes.Item("value").Value)
                response.Close()
                parameters.Clear()
                For i = 0 To respHTML.Split(vbNewLine).Count - 1
                    Dim _line As String = respHTML.Split(vbNewLine)(i)
                    If Not _line = Nothing Then
                        If (_line.Contains("$ContentPlaceHolder") And _line.Contains("HiddenField") And Not _line.Contains("gvInProgress") And Not _line.Contains("GridViewContent") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            If _line.Contains("value") Then
                                startPos = _line.IndexOf("value=""", StringComparison.CurrentCultureIgnoreCase) + 7
                                endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                                _value = _line.Substring(startPos, endPos - startPos)
                            End If
                            parameters.Add(_name, System.Uri.EscapeDataString(_value))
                        ElseIf (_line.Contains("$ContentPlaceHolder") And _line.Contains("gvInProgress") And Not _line.Contains("GridViewContent") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            If _line.Contains("value") Then
                                startPos = _line.IndexOf("value=""", StringComparison.CurrentCultureIgnoreCase) + 7
                                endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                                _value = _line.Substring(startPos, endPos - startPos)
                            End If
                            parameters.Add(_name, System.Uri.EscapeDataString(_value))
                        ElseIf (_line.Contains("$ContentPlaceHolder") And Not _line.Contains("gvInProgress") And _line.Contains("GridViewContent") And Not _line.Contains("RadioButtonList1") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            If _line.Contains("value") Then
                                startPos = _line.IndexOf("value=""", StringComparison.CurrentCultureIgnoreCase) + 7
                                endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                                _value = _line.Substring(startPos, endPos - startPos)
                            End If
                            parameters.Add(_name, System.Uri.EscapeDataString(_value))
                        ElseIf (_line.Contains("$ContentPlaceHolder") And Not _line.Contains("gvInProgress") And _line.Contains("GridViewContent") And _line.Contains("RadioButtonList1") And Not _line.Contains("javascript")) Then
                            Dim startPos As Integer, endPos As Integer
                            Dim _name As String = Nothing
                            Dim _value As String = ""
                            startPos = _line.IndexOf("name=""", StringComparison.CurrentCultureIgnoreCase) + 6
                            endPos = _line.IndexOf("""", startPos, StringComparison.CurrentCultureIgnoreCase)
                            _name = _line.Substring(startPos, endPos - startPos)
                            Dim RndNum As New Random()
                            '_value = RndNum.Next(1, 4)
                            _value = 1
                            Try
                                parameters.Add(_name, System.Uri.EscapeDataString(_value))
                            Catch ex As Exception

                            End Try
                        End If
                    End If
                Next
                parameters.Add("__EVENTTARGET", "")
                parameters.Add("__EVENTARGUMENT", "")
                parameters.Add("__VIEWSTATE", System.Uri.EscapeDataString(__VIEWSTATE))
                parameters.Add("__VIEWSTATEGENERATOR", System.Uri.EscapeDataString(__VIEWSTATEGENERATOR))
                parameters.Add("__EVENTVALIDATION", System.Uri.EscapeDataString(__EVENTVALIDATION))
                parameters.Add("ctl00$ContentPlaceHolder1$ButtonSend", System.Uri.EscapeDataString("送出(Send)"))
                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.155/Questionnaire/QuestionnaireInsert.aspx?UserId=" & _UserId, parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                response.Close()
                _count += 1
            End While
            If _count > 0 Then
                If MsgBox("填寫完畢 , 共填寫完" & _count & "個教學評量 , 是否要至選課系統查看填寫結果 !?", MsgBoxStyle.Information + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    System.Diagnostics.Process.Start("http://140.127.113.107/Account/LogOn?ReturnUrl=%2f")
                End If
            Else
                If MsgBox("沒有教學評量需要填寫 , 是否要至選課系統查看 !?", MsgBoxStyle.Information + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    System.Diagnostics.Process.Start("http://140.127.113.107/Account/LogOn?ReturnUrl=%2f")
                End If
            End If
        Catch ex As Exception
            MsgBox("發生異常 , 暫無法自動填寫教學評量 !", MsgBoxStyle.Critical)
        End Try
        Button1.Enabled = True
        GroupBox1.Enabled = True
        GroupBox2.Enabled = True
        GroupBox3.Enabled = True
        GroupBox4.Enabled = True
    End Sub
    Sub LoginAP()
        Try
            'Dim parameters As IDictionary(Of String, String) = New Dictionary(Of String, String)()
            'parameters.Add("uid", userName)
            'parameters.Add("pwd", password)
            'Dim response As HttpWebResponse = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/perchk.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
            'Dim reader As StreamReader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
            'Dim respHTML As String = reader.ReadToEnd()
            ''If respHTML.Contains("script") Then
            ''    MsgBox("系統抓取異常 , 請稍後再試 :)", MsgBoxStyle.Critical, "Opps ! Something Error :(")
            ''    End
            ''End If

            Dim response As HttpWebResponse = HttpWebResponseUtility.CreateGetHttpResponse("http://140.127.113.231/kuas/f_head.jsp", Nothing, Nothing, cookies)
            Dim reader As StreamReader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
            Dim respHTML As String = reader.ReadToEnd()

            Dim doc As New HtmlDocument()
            doc.LoadHtml(respHTML)
            Dim node As HtmlNode = doc.DocumentNode

            Try
                TranGrade(WebUtility.HtmlDecode(node.SelectNodes("/html/body/div[1]/div/div[3]/span[2]")(0).InnerText))
                Me.Text = "KUAS AP (By Silent) @ " & WebUtility.HtmlDecode(node.SelectNodes("/html/body/div[1]/div/div[3]/span[3]")(0).InnerText)
            Catch ex As Exception
                MsgBox("系統抓取異常 , 請稍後再試 :)", MsgBoxStyle.Critical, "Opps ! Something Error :(")
                End
            End Try

            Me.Refresh()
            ReadClassSemester()
            Me.Refresh()
            ReadScoreSemester()
            Me.Refresh()
            ReadAbsentSemester()

            Try
                CatchClassThread = New Thread(AddressOf Me.RefreshClass)
                CatchClassThread.Start()
            Catch ex As Exception

            End Try
        Catch ex As Exception
            MsgBox("系統抓取異常 , 請稍後再試 :)", MsgBoxStyle.Critical, "Opps ! Something Error :(")
            End
        End Try
    End Sub
    Sub TranGrade(Grade As String)
        If Grade.Contains("一") Then
            StudentGrade = 1
        ElseIf Grade.Contains("二") Then
            StudentGrade = 2
        ElseIf Grade.Contains("三") Then
            StudentGrade = 3
        ElseIf Grade.Contains("四") Then
            StudentGrade = 4
        Else
            StudentGrade = 0
        End If
    End Sub
    Sub AddNewClassItem(Time As String, MON As String, TUE As String, WEN As String, THU As String, FRI As String, Optional SAT As String = Nothing, Optional SUN As String = Nothing)
        Dim item As New ListViewItem
        item.Text = Time
        item.SubItems.Add(MON)
        item.SubItems.Add(TUE)
        item.SubItems.Add(WEN)
        item.SubItems.Add(THU)
        item.SubItems.Add(FRI)
        If Not SAT = Nothing Then : item.SubItems.Add(SAT) : End If
        If Not SUN = Nothing Then : item.SubItems.Add(SUN) : End If
        If Not ClassListView.Items.Contains(item) Then : ClassListView.Items.Add(item) : End If
    End Sub
    Dim ClassSemester As New ListBox
    Sub ReadClassSemester()
        While (1)
            Try
                Dim parameters As IDictionary(Of String, String) = New Dictionary(Of String, String)()
                Dim response As HttpWebResponse
                Dim reader As StreamReader
                Dim respHTML As String

                parameters.Clear()
                parameters.Add("fncid", "AG222")
                If Now.Month < 9 Then
                    parameters.Add("sysyear", Now.Year - 1912)
                    parameters.Add("syssms", 2)
                Else
                    parameters.Add("sysyear", Now.Year - 1911)
                    parameters.Add("syssms", 1)
                End If
                parameters.Add("std_id", "")
                parameters.Add("local_ip", "")
                parameters.Add("online", "okey")
                parameters.Add("loginid", userName)
                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/fnc.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                response.Close()
                Dim doc As New HtmlDocument()
                doc.LoadHtml(respHTML)
                Dim node As HtmlNode = doc.DocumentNode
                Dim ls_randnum As String = WebUtility.HtmlDecode(node.SelectNodes("//input[9]")(0).Attributes.Item("value").Value)
                Dim _arg01 = WebUtility.HtmlDecode(node.SelectNodes("//input[1]")(0).Attributes.Item("value").Value)
                Dim _arg02 = WebUtility.HtmlDecode(node.SelectNodes("//input[2]")(0).Attributes.Item("value").Value)

                parameters.Clear()
                parameters.Add("fncid", "AG222")
                parameters.Add("arg01", _arg01)
                parameters.Add("arg02", _arg02)
                parameters.Add("arg03", userName)
                parameters.Add("arg04", "")
                parameters.Add("arg05", "")
                parameters.Add("arg06", "")
                parameters.Add("uid", userName)
                parameters.Add("ls_randnum", ls_randnum)
                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/system/sys001_00.jsp?spath=ag_pro/ag222.jsp?", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()

                HtmlAgilityPack.HtmlNode.ElementsFlags.Remove("option")
                doc.LoadHtml(respHTML)

                Dim nodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//option")
                ClassSemesterCombobox.Items.Clear()
                ClassSemester.Items.Clear()
                For Each nodex As HtmlNode In nodes
                    ClassSemesterCombobox.Items.Add(nodex.InnerText)
                    ClassSemester.Items.Add(nodex.Attributes.Item("value").Value)
                Next
                ClassSemesterCombobox.SelectedItem = doc.DocumentNode.SelectSingleNode("//option[@selected]").InnerText
                ymsClass = doc.DocumentNode.SelectSingleNode("//option[@selected]").Attributes.Item("value").Value
                ClassReady = True
                Exit While
            Catch ex As Exception
                'ClassReady = True
            End Try
        End While
    End Sub
    Sub RefreshClass()
        ClassSemesterCombobox.Enabled = False
        Dim ErrorCount As Integer = 0
        If ClassCheck = True Or ClassReady = False Then
            Exit Sub
        End If
        While (1)
            If ErrorCount >= 3 Then
                MsgBox("暫時無法查詢課表 , 請稍後再試 !", MsgBoxStyle.Critical)
                ClassCheck = False
                ClassFirst = True
                Exit While
            End If
            Try
                ClassCheck = True
                ClassListView.Items.Clear()
                Me.Size = New Point(FrmX, Me.Size.Height)

                Dim parameters As IDictionary(Of String, String) = New Dictionary(Of String, String)()
                Dim response As HttpWebResponse
                Dim reader As StreamReader
                Dim respHTML As String
                Dim doc As New HtmlDocument()

                Dim OpenB As Boolean = False
                Dim OpenH As Boolean = False
                'Debug.Print(WebUtility.HtmlDecode(node.SelectNodes("//select")(0).InnerText))

                'parameters.Clear()
                'parameters.Add("fncid", "AG222")
                'If Now.Month < 9 Then
                '    parameters.Add("sysyear", Now.Year - 1912)
                '    parameters.Add("syssms", 2)
                'Else
                '    parameters.Add("sysyear", Now.Year - 1911)
                '    parameters.Add("syssms", 1)
                'End If
                'parameters.Add("std_id", "")
                'parameters.Add("local_ip", "")
                'parameters.Add("online", "okey")
                'parameters.Add("loginid", userName)
                'response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/fnc.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                'reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                'respHTML = reader.ReadToEnd()
                'doc.LoadHtml(respHTML)
                'Dim node As HtmlNode = doc.DocumentNode
                'Dim ls_randnum As String = WebUtility.HtmlDecode(node.SelectNodes("//input[9]")(0).Attributes.Item("value").Value)

                parameters.Clear()
                If Not ymsClass = Nothing Then
                    parameters.Add("yms", System.Uri.EscapeDataString(ymsClass))
                    parameters.Add("arg01", ymsClass.Split(",")(0))
                    parameters.Add("arg02", ymsClass.Split(",")(1))
                Else
                    If Now.Month < 9 Then
                        parameters.Add("yms", System.Uri.EscapeDataString((Now.Year - 1912) & ",2"))
                        parameters.Add("arg01", Now.Year - 1912)
                        parameters.Add("arg02", 2)
                    Else
                        parameters.Add("yms", System.Uri.EscapeDataString((Now.Year - 1911) & ",1"))
                        parameters.Add("arg01", Now.Year - 1911)
                        parameters.Add("arg02", 1)
                    End If
                End If
                parameters.Add("spath", "ag_pro%2Fag222.jsp%3F")
                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/ag_pro/ag222.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                'Debug.Print(respHTML)
                If respHTML.Contains("無") Then
                    MsgBox("查無資料 !", MsgBoxStyle.Exclamation)
                    ClassCheck = False
                    ClassFirst = True
                    Exit While
                End If
                response.Close()

                doc.LoadHtml(respHTML)
                Dim node As HtmlNode = doc.DocumentNode

                For i = 12 To 16
                    For j = 2 To 8
                        If Not WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[" & j & "]")(0).InnerHtml.Split("<br>")(0)).Trim = "" Then
                            OpenB = True
                            GoTo FindB
                        End If
                    Next
                Next
FindB:
                For i = 2 To 16
                    For j = 7 To 8
                        If Not WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[" & j & "]")(0).InnerHtml.Split("<br>")(0)).Trim = "" Then
                            OpenH = True
                            GoTo FindH
                        End If
                    Next
                Next
FindH:
                If OpenB And OpenH Then
                    Me.Size = New Point(915, Me.Size.Height)
                    AddNewClassItem("M", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("一", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("二", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("三", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("四", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("A", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("五", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("六", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("七", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("八", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("B", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[12]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[12]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[12]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[12]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[12]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[12]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[12]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("十一", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[13]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[13]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[13]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[13]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[13]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[13]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[13]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("十二", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[14]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[14]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[14]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[14]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[14]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[14]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[14]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("十三", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[15]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[15]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[15]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[15]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[15]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[15]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[15]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("十四", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[16]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[16]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[16]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[16]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[16]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[16]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[16]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                ElseIf OpenB Then
                    Me.Size = New Point(FrmX, Me.Size.Height)
                    AddNewClassItem("M", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("一", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("二", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("三", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("四", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("A", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("五", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("六", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("七", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("八", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("B", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[12]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[12]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[12]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[12]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[12]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("十一", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[13]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[13]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[13]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[13]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[13]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("十二", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[14]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[14]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[14]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[14]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[14]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("十三", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[15]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[15]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[15]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[15]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[15]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("十四", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[16]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[16]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[16]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[16]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[16]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                ElseIf OpenH Then
                    Me.Size = New Point(915, Me.Size.Height)
                    AddNewClassItem("M", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("一", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("二", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("三", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("四", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("A", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("五", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("六", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("七", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                    AddNewClassItem("八", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[6]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[7]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[8]")(0).InnerHtml.Split("<br>")(0)))
                Else
                    GoTo DontOpen
                End If
                ClassCheck = False
                ClassFirst = True
                response.Close()
                Exit While
DontOpen:
                Me.Size = New Point(FrmX, Me.Size.Height)
                AddNewClassItem("M", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[2]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                AddNewClassItem("一", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[3]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                AddNewClassItem("二", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[4]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                AddNewClassItem("三", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[5]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                AddNewClassItem("四", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[6]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                AddNewClassItem("A", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[7]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                AddNewClassItem("五", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[8]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                AddNewClassItem("六", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[9]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                AddNewClassItem("七", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[10]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                AddNewClassItem("八", WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[2]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[3]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[4]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[5]")(0).InnerHtml.Split("<br>")(0)), WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[11]/td[6]")(0).InnerHtml.Split("<br>")(0)))
                ClassCheck = False
                ClassFirst = True
                response.Close()
                Exit While
            Catch ex As Exception
                'MsgBox("暫時無法查詢課表 , 請稍後再試 !", MsgBoxStyle.Critical)
                'ClassCheck = False
                ClassFirst = True
            End Try
        End While
        ClassSemesterCombobox.Enabled = True
    End Sub
    Sub AddNewScoreItem(Subject As String, Teacher As String, Score As String, Alert As String, AlertCause As String, YearScore As String)
        Subject = Subject.Trim()
        Teacher = Teacher.Trim()
        Score = Score.Trim()
        Alert = Alert.Trim()
        AlertCause = AlertCause.Trim()
        YearScore = YearScore.Trim()
        If Subject = Nothing Then : Subject = "*" : End If
        If Teacher = Nothing Then : Teacher = "*" : End If
        If Score = Nothing Then : Score = "*" : End If
        If Alert = Nothing Then : Alert = "*" : End If
        If AlertCause = Nothing Then : AlertCause = "*" : End If
        If YearScore = Nothing Then : YearScore = "*" : End If

        Dim item As New ListViewItem
        item.Text = Subject
        item.SubItems.Add(Teacher)
        item.SubItems.Add(Score)
        item.SubItems.Add(YearScore)
        item.SubItems.Add(Alert)
        item.SubItems.Add(AlertCause)
        MidtermListView.Items.Add(item)
    End Sub
    Dim ScoreSemester As New ListBox
    Sub ReadScoreSemester()
        While (1)
            Try
                Dim parameters As IDictionary(Of String, String) = New Dictionary(Of String, String)()
                Dim response As HttpWebResponse
                Dim reader As StreamReader
                Dim respHTML As String

                parameters.Clear()
                parameters.Add("fncid", "AG008")
                If Now.Month < 9 Then
                    parameters.Add("sysyear", Now.Year - 1912)
                    parameters.Add("syssms", 2)
                Else
                    parameters.Add("sysyear", Now.Year - 1911)
                    parameters.Add("syssms", 1)
                End If
                parameters.Add("std_id", "")
                parameters.Add("local_ip", "")
                parameters.Add("online", "okey")
                parameters.Add("loginid", userName)
                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/fnc.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                response.Close()
                Dim doc As New HtmlDocument()
                doc.LoadHtml(respHTML)
                Dim node As HtmlNode = doc.DocumentNode
                Dim ls_randnum As String = WebUtility.HtmlDecode(node.SelectNodes("//input[9]")(0).Attributes.Item("value").Value)
                Dim _arg01 = WebUtility.HtmlDecode(node.SelectNodes("//input[1]")(0).Attributes.Item("value").Value)
                Dim _arg02 = WebUtility.HtmlDecode(node.SelectNodes("//input[2]")(0).Attributes.Item("value").Value)

                parameters.Clear()
                parameters.Add("fncid", "AG008")
                parameters.Add("arg01", _arg01)
                parameters.Add("arg02", _arg02)
                parameters.Add("arg03", userName)
                parameters.Add("arg04", "")
                parameters.Add("arg05", "")
                parameters.Add("arg06", "")
                parameters.Add("uid", userName)
                parameters.Add("ls_randnum", ls_randnum)
                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/system/sys001_00.jsp?spath=ag_pro/ag008.jsp?", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                HtmlAgilityPack.HtmlNode.ElementsFlags.Remove("option")
                doc.LoadHtml(respHTML)

                Dim nodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//option")
                ScoreSemesterCombobox.Items.Clear()
                ScoreSemester.Items.Clear()
                For Each nodex As HtmlNode In nodes
                    ScoreSemesterCombobox.Items.Add(nodex.InnerText)
                    ScoreSemester.Items.Add(nodex.Attributes.Item("value").Value)
                Next
                ScoreSemesterCombobox.SelectedItem = doc.DocumentNode.SelectSingleNode("//option[@selected]").InnerText
                ymsScore = doc.DocumentNode.SelectSingleNode("//option[@selected]").Attributes.Item("value").Value
                ScoreReady = True
                Exit While
            Catch ex As Exception
                'ScoreReady = True
            End Try
        End While
    End Sub
    Sub RefreshScore()
        ScoreSemesterCombobox.Enabled = False
        Dim ErrorCount As Integer = 0
        If ScoreCheck = True Or ScoreReady = False Then
            Exit Sub
        End If
        While (1)
            If ErrorCount >= 3 Then
                MsgBox("暫時無法查詢成績 , 請稍後再試 !", MsgBoxStyle.Critical)
                ScoreCheck = False
                ScoreFirst = True
                Exit While
            End If
            Try
                ScoreLabel.Text = ""
                ScoreCheck = True
                MidtermListView.Items.Clear()
                Dim parameters As IDictionary(Of String, String) = New Dictionary(Of String, String)()
                Dim response As HttpWebResponse
                Dim reader As StreamReader
                Dim respHTML As String
                Dim GetList(50, 6) As String
                Dim FluckOutCount As Integer = 0
                Dim FluckOutCheck As Boolean = False
                Dim TotalCredit As Double = 0
                Dim TotalFluckOutCredit As Double = 0

                parameters.Clear()
                parameters.Add("fncid", "AG009")
                If Not ymsScore = Nothing Then
                    parameters.Add("sysyear", ymsScore.Split(",")(0))
                    parameters.Add("syssms", ymsScore.Split(",")(1))
                Else
                    If Now.Month < 9 Then
                        parameters.Add("sysyear", Now.Year - 1912)
                        parameters.Add("syssms", 2)
                    Else
                        parameters.Add("sysyear", Now.Year - 1911)
                        parameters.Add("syssms", 1)
                    End If
                End If
                parameters.Add("std_id", "")
                parameters.Add("local_ip", "")
                parameters.Add("online", "okey")
                parameters.Add("loginid", userName)
                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/fnc.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                Dim doc As New HtmlDocument()
                doc.LoadHtml(respHTML)
                Dim node As HtmlNode = doc.DocumentNode
                Dim ls_randnum As String = WebUtility.HtmlDecode(node.SelectNodes("//input[9]")(0).Attributes.Item("value").Value)
                response.Close()

                parameters.Clear()
                If Not ymsScore = Nothing Then
                    parameters.Add("yms", System.Uri.EscapeDataString(ymsScore))
                    parameters.Add("arg01", ymsScore.Split(",")(0))
                    parameters.Add("arg02", ymsScore.Split(",")(1))
                Else
                    If Now.Month < 9 Then
                        parameters.Add("yms", System.Uri.EscapeDataString((Now.Year - 1912) & ",2"))
                        parameters.Add("arg01", Now.Year - 1912)
                        parameters.Add("arg02", 2)
                    Else
                        parameters.Add("yms", System.Uri.EscapeDataString((Now.Year - 1911) & ",1"))
                        parameters.Add("arg01", Now.Year - 1911)
                        parameters.Add("arg02", 1)
                    End If
                End If
                parameters.Add("spath", "ag_pro/ag008.jsp?")
                Dim responsex As HttpWebResponse = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/ag_pro/ag008.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                Dim readerx As StreamReader = New StreamReader(responsex.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                Dim respHTMLx = readerx.ReadToEnd()
                'Debug.Print(respHTMLx)
                If respHTMLx.Contains("無") Then
                    MsgBox("查無資料 !", MsgBoxStyle.Exclamation)
                    ScoreCheck = False
                    ScoreFirst = True
                    Exit While
                End If
                response.Close()

                doc.LoadHtml(respHTMLx)
                node = doc.DocumentNode
                '<td>項次</td>
                '<td>科目名稱</td>
                '<td>學分數</td>
                '<td>授課時數</td>
                '<td>必選修</td>
                '<td>開課別</td>
                '<td>期中成績</td>
                '<td>學期成績</td>
                '<td>備註</td>
                For i = 2 To node.SelectNodes("//table[2]/tr").Count
                    GetList(i - 2, 0) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[2]")(0).InnerText).Trim
                    GetList(i - 2, 4) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[7]")(0).InnerText).Trim
                    GetList(i - 2, 5) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[8]")(0).InnerText).Trim

                    TotalCredit += Val(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[3]")(0).InnerText).Trim)
                    If WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[8]")(0).InnerText).Trim = "" Then
                        FluckOutCheck = True
                    ElseIf WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[8]")(0).InnerText).Trim = "不合格" Then
                        FluckOutCount += 1
                    ElseIf Val(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[8]")(0).InnerText).Trim) < 60 Then
                        If WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[8]")(0).InnerText).Trim.Contains("合格") Then
                            Continue For
                        End If
                        FluckOutCount += 1
                        TotalFluckOutCredit += Val(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[3]")(0).InnerText).Trim)
                    End If

                    'For j = 0 To 49
                    '    If GetList(j, 0) = Nothing Then : Exit For : End If
                    '    If GetList(j, 0) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[2]")(0).InnerText) Then
                    '        GetList(j, 4) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[7]")(0).InnerText)
                    '        GetList(j, 5) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[8]")(0).InnerText)
                    '        Exit For
                    '    End If
                    'Next
                Next
                Dim tmpScoreLabel As String = WebUtility.HtmlDecode(node.SelectNodes("//caption[1]/div[1]")(0).InnerHtml).Split("<br>")(0).Replace("　　　　", "     ")

                Dim ymsCheck As Boolean = False
                If ymsScore = Nothing Then
                    ymsCheck = False
                Else
                    If Now.Month < 9 Then
                        If Not (Now.Year - 1912) = ymsScore.Split(",")(0) Or Not 2 = ymsScore.Split(",")(1) Then
                            ymsCheck = True
                        End If
                    Else
                        If Not (Now.Year - 1911) = ymsScore.Split(",")(0) Or Not 1 = ymsScore.Split(",")(1) Then
                            ymsCheck = True
                        End If
                    End If
                End If
                ymsCheck = True
                If ymsCheck = False Then
                    parameters.Clear()
                    parameters.Add("fncid", "AG009")
                    If Not ymsScore = Nothing Then
                        parameters.Add("arg01", ymsScore.Split(",")(0))
                        parameters.Add("arg02", ymsScore.Split(",")(1))
                    Else
                        If Now.Month < 9 Then
                            parameters.Add("arg01", Now.Year - 1912)
                            parameters.Add("arg02", 2)
                        Else
                            parameters.Add("arg01", Now.Year - 1911)
                            parameters.Add("arg02", 1)
                        End If
                    End If
                    parameters.Add("arg03", userName)
                    parameters.Add("arg04", "")
                    parameters.Add("arg05", "")
                    parameters.Add("arg06", "")
                    parameters.Add("uid", userName)
                    parameters.Add("ls_randnum", ls_randnum)
                    response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/ag_pro/ag009.jsp?", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                    reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                    respHTML = reader.ReadToEnd()

                    doc.LoadHtml(respHTML)
                    node = doc.DocumentNode
                    '3 科目名稱
                    '5 老師名稱
                    '6 預警
                    '7 預警原因
                    If Not respHTML.Contains("目前無學生個人成績資料") Then
                        For i = 2 To node.SelectNodes("//table[2]/tr").Count
                            For j = 0 To 49
                                If GetList(j, 0) = Nothing Then : Exit For : End If
                                If GetList(j, 0) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[3]")(0).InnerText).Trim Then
                                    GetList(j, 1) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[5]")(0).InnerText).Trim
                                    GetList(j, 2) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[6]")(0).InnerText).Trim
                                    GetList(j, 3) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[7]")(0).InnerText).Trim
                                    Exit For
                                End If
                            Next
                        Next
                    End If

                    response.Close()
                    'For i = 2 To node.SelectNodes("//table[2]/tr").Count
                    '    GetList(i - 2, 0) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[3]")(0).InnerText)
                    '    GetList(i - 2, 1) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[5]")(0).InnerText)
                    '    GetList(i - 2, 2) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[6]")(0).InnerText)
                    '    GetList(i - 2, 3) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[7]")(0).InnerText)
                    'Next
                Else
                    parameters.Clear()
                    If Not ymsScore = Nothing Then
                        parameters.Add("yms", System.Uri.EscapeDataString(ymsScore))
                        parameters.Add("arg01", ymsScore.Split(",")(0))
                        parameters.Add("arg02", ymsScore.Split(",")(1))
                    Else
                        If Now.Month < 9 Then
                            parameters.Add("yms", System.Uri.EscapeDataString((Now.Year - 1912) & ",2"))
                            parameters.Add("arg01", Now.Year - 1912)
                            parameters.Add("arg02", 2)
                        Else
                            parameters.Add("yms", System.Uri.EscapeDataString((Now.Year - 1911) & ",1"))
                            parameters.Add("arg01", Now.Year - 1911)
                            parameters.Add("arg02", 1)
                        End If
                    End If
                    parameters.Add("spath", "ag_pro%2Fag222.jsp%3F")
                    response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/ag_pro/ag222.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                    reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                    respHTML = reader.ReadToEnd()
                    'Debug.Print(respHTML)
                    doc.LoadHtml(respHTML)
                    node = doc.DocumentNode
                    '3 科目名稱
                    '5 老師名稱
                    '6 預警
                    '7 預警原因
                    For i = 2 To node.SelectNodes("//table[1]/tr").Count
                        For j = 0 To 49
                            If GetList(j, 0) = Nothing Then : Exit For : End If
                            If GetList(j, 0) = WebUtility.HtmlDecode(node.SelectNodes("//table[1]/tr[" & i & "]/td[2]")(0).InnerText).Trim Then
                                GetList(j, 1) = WebUtility.HtmlDecode(node.SelectNodes("//table[1]/tr[" & i & "]/td[10]")(0).InnerText).Trim
                                Exit For
                            End If
                        Next
                    Next
                    response.Close()
                End If

                For i = 0 To 49
                    If GetList(i, 0) = Nothing Then : Exit For : End If
                    If GetList(i, 1) = Nothing Then : GetList(i, 1) = "*" : End If
                    If GetList(i, 2) = Nothing Then : GetList(i, 2) = "*" : End If
                    If GetList(i, 3) = Nothing Then : GetList(i, 3) = "*" : End If
                    AddNewScoreItem(GetList(i, 0), GetList(i, 1), GetList(i, 4), GetList(i, 2), GetList(i, 3), GetList(i, 5))
                Next
                ScoreCheck = False
                ScoreFirst = True

                ScoreLabel.Text = tmpScoreLabel

                If FluckOutCheck = True Then
                    If FluckOutCount = 0 Then
                        MsgBox("成績尚未完全公布 , 恭喜你目前歐趴 !", MsgBoxStyle.Information, ScoreSemesterCombobox.Items.Item(ScoreSemesterCombobox.SelectedIndex).ToString)
                    Else
                        If TotalFluckOutCredit > TotalCredit * 2 / 3 Then
                            MsgBox("成績尚未完全公布 , 目前被當" & FluckOutCount & "科 , 合計" & TotalFluckOutCredit & "學分 !" & vbCrLf & "已達三二標準 !", MsgBoxStyle.Critical, ScoreSemesterCombobox.Items.Item(ScoreSemesterCombobox.SelectedIndex).ToString.Trim)
                        ElseIf TotalFluckOutCredit > TotalCredit / 2 Then
                            MsgBox("成績尚未完全公布 , 目前被當" & FluckOutCount & "科 , 合計" & TotalFluckOutCredit & "學分 !" & vbCrLf & "已達二一標準 !", MsgBoxStyle.Information, ScoreSemesterCombobox.Items.Item(ScoreSemesterCombobox.SelectedIndex).ToString.Trim)
                        Else
                            MsgBox("成績尚未完全公布 , 目前被當" & FluckOutCount & "科 , 合計" & TotalFluckOutCredit & "學分 !" & vbCrLf & "恭喜你未達二一標準 !", MsgBoxStyle.Information, ScoreSemesterCombobox.Items.Item(ScoreSemesterCombobox.SelectedIndex).ToString.Trim)
                        End If
                    End If
                Else
                    If FluckOutCount = 0 Then
                        MsgBox("恭喜你歐趴 !", MsgBoxStyle.Information, ScoreSemesterCombobox.Items.Item(ScoreSemesterCombobox.SelectedIndex).ToString)
                    Else

                        If TotalFluckOutCredit > TotalCredit * 2 / 3 Then
                            MsgBox("被當" & FluckOutCount & "科 , 合計" & TotalFluckOutCredit & "學分 !" & vbCrLf & "已達三二標準 !", MsgBoxStyle.Critical, ScoreSemesterCombobox.Items.Item(ScoreSemesterCombobox.SelectedIndex).ToString.Trim)
                        ElseIf TotalFluckOutCredit > TotalCredit / 2 Then
                            MsgBox("被當" & FluckOutCount & "科 , 合計" & TotalFluckOutCredit & "學分 !" & vbCrLf & "已達二一標準 !", MsgBoxStyle.Exclamation, ScoreSemesterCombobox.Items.Item(ScoreSemesterCombobox.SelectedIndex).ToString.Trim)
                        Else
                            MsgBox("被當" & FluckOutCount & "科 , 合計" & TotalFluckOutCredit & "學分 !" & vbCrLf & "恭喜你未達二一標準 !", MsgBoxStyle.Information, ScoreSemesterCombobox.Items.Item(ScoreSemesterCombobox.SelectedIndex).ToString.Trim)
                        End If
                    End If
                End If
                Exit While
            Catch ex As Exception
                ErrorCount += 1
                'MsgBox("暫時無法查詢成績 , 請稍後再試 !", MsgBoxStyle.Critical)
                'ScoreCheck = False
                ScoreFirst = True
            End Try
        End While
        ScoreSemesterCombobox.Enabled = True
    End Sub
    Function TranAbsent(Item As String)
        Select Case Item
            Case "缺曠"
                Return "曠"
            Case "病假"
                Return "病"
            Case "公假"
                Return "公"
            Case "事假"
                Return "事"
            Case "產假"
                Return "產"
            Case "喪假"
                Return "喪"
            Case Else
                Return ""
        End Select
    End Function
    Sub AddNewAbsentItem(Time As String, MorningRally As String, MorningReading As String, a1 As String, a2 As String, a3 As String, a4 As String, B As String, a5 As String, a6 As String, a7 As String, a8 As String, C As String, a11 As String, a12 As String, a13 As String, a14 As String)
        Dim item As New ListViewItem
        item.Text = Time

        If Not MorningRally = Nothing Then
            item.SubItems.Add(MorningRally)
        Else
            item.SubItems.Add(MorningReading)
        End If
        item.SubItems.Add(a1)
        item.SubItems.Add(a2)
        item.SubItems.Add(a3)
        item.SubItems.Add(a4)
        item.SubItems.Add(B)
        item.SubItems.Add(a5)
        item.SubItems.Add(a6)
        item.SubItems.Add(a7)
        item.SubItems.Add(a8)
        AbsentListView.Items.Add(item)
    End Sub
    Dim AbsentSemester As New ListBox
    Sub ReadAbsentSemester()
        While (1)
            Try
                Dim parameters As IDictionary(Of String, String) = New Dictionary(Of String, String)()
                Dim response As HttpWebResponse
                Dim reader As StreamReader
                Dim respHTML As String

                parameters.Clear()
                parameters.Add("fncid", "AK002")
                If Now.Month < 9 Then
                    parameters.Add("sysyear", Now.Year - 1912)
                    parameters.Add("syssms", 2)
                Else
                    parameters.Add("sysyear", Now.Year - 1911)
                    parameters.Add("syssms", 1)
                End If
                parameters.Add("std_id", "")
                parameters.Add("local_ip", "")
                parameters.Add("online", "okey")
                parameters.Add("loginid", userName)
                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/fnc.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                Dim doc As New HtmlDocument()
                doc.LoadHtml(respHTML)
                Dim node As HtmlNode = doc.DocumentNode
                Dim ls_randnum As String = WebUtility.HtmlDecode(node.SelectNodes("//input[9]")(0).Attributes.Item("value").Value)
                Dim _arg01 = WebUtility.HtmlDecode(node.SelectNodes("//input[1]")(0).Attributes.Item("value").Value)
                Dim _arg02 = WebUtility.HtmlDecode(node.SelectNodes("//input[2]")(0).Attributes.Item("value").Value)

                parameters.Clear()
                parameters.Add("fncid", "AK002")
                parameters.Add("arg01", _arg01)
                parameters.Add("arg02", _arg02)
                parameters.Add("arg03", userName)
                parameters.Add("arg04", "")
                parameters.Add("arg05", "")
                parameters.Add("arg06", "")
                parameters.Add("uid", userName)
                parameters.Add("ls_randnum", ls_randnum)
                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/system/sys001_00.jsp?spath=ak_pro/ak002_01.jsp?", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                HtmlAgilityPack.HtmlNode.ElementsFlags.Remove("option")
                doc.LoadHtml(respHTML)

                Dim nodes As HtmlNodeCollection = doc.DocumentNode.SelectNodes("//option")
                AbsentSemesterCombobox.Items.Clear()
                AbsentSemester.Items.Clear()
                For Each nodex As HtmlNode In nodes
                    AbsentSemesterCombobox.Items.Add(nodex.InnerText)
                    AbsentSemester.Items.Add(nodex.Attributes.Item("value").Value)
                Next
                AbsentSemesterCombobox.SelectedItem = doc.DocumentNode.SelectSingleNode("//option[@selected]").InnerText
                ymsAbsent = doc.DocumentNode.SelectSingleNode("//option[@selected]").Attributes.Item("value").Value
                AbsentReady = True
                Exit While
            Catch ex As Exception
                'AbsentReady = True
            End Try
        End While
    End Sub
    Sub RefreshAbsent()
        AbsentSemesterCombobox.Enabled = False
        Dim ErrorCount As Integer = 0
        If AbsentCheck = True Or AbsentReady = False Then
            Exit Sub
        End If
        While (1)
            If ErrorCount >= 3 Then
                MsgBox("暫時無法查詢缺礦 , 請稍後再試 !", MsgBoxStyle.Critical)
                AbsentCheck = False
                AbsentFirst = True
                Exit While
            End If
            Try
                AbsentCheck = True
                AbsentListView.Items.Clear()
                Dim parameters As IDictionary(Of String, String) = New Dictionary(Of String, String)()
                Dim response As HttpWebResponse
                Dim reader As StreamReader
                Dim respHTML As String

                Dim doc As New HtmlDocument()

                parameters.Clear()
                If Not ymsAbsent = Nothing Then
                    parameters.Add("yms", System.Uri.EscapeDataString(ymsAbsent))
                    parameters.Add("arg01", ymsAbsent.Split(",")(0))
                    parameters.Add("arg02", ymsAbsent.Split(",")(1))
                Else
                    If Now.Month < 9 Then
                        parameters.Add("yms", System.Uri.EscapeDataString((Now.Year - 1912) & ",2"))
                        parameters.Add("arg01", Now.Year - 1912)
                        parameters.Add("arg02", 2)
                    Else
                        parameters.Add("yms", System.Uri.EscapeDataString((Now.Year - 1911) & ",1"))
                        parameters.Add("arg01", Now.Year - 1911)
                        parameters.Add("arg02", 1)
                    End If
                End If
                parameters.Add("spath", "ak_pro%2Fak002_01.jsp%3F")
                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/ak_pro/ak002_01.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()

                If respHTML.Contains("無") Then
                    MsgBox("本學期無缺曠課紀錄", MsgBoxStyle.Exclamation)
                    AbsentCheck = False
                    AbsentFirst = True
                    Exit While
                End If

                'Debug.Print(respHTML)
                doc.LoadHtml(respHTML)
                Dim node As HtmlNode = doc.DocumentNode

                For i = 2 To node.SelectNodes("//table[2]/tr").Count
                    AddNewAbsentItem(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[3]")(0).InnerText), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[4]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[5]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[6]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[7]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[8]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[9]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[10]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[11]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[12]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[13]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[14]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[15]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[16]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[17]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[18]")(0).InnerText)), TranAbsent(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[19]")(0).InnerText)))
                Next
                AbsentCheck = False
                AbsentFirst = True
                response.Close()
                Exit While
            Catch ex As Exception
                'MsgBox("暫時無法查詢缺礦 , 請稍後再試 !", MsgBoxStyle.Critical)
                'AbsentCheck = False
                AbsentFirst = True
                ErrorCount += 1
            End Try
        End While
        AbsentSemesterCombobox.Enabled = True
    End Sub

    Private Sub ClassSemesterCombobox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ClassSemesterCombobox.SelectedIndexChanged
        If ClassFirst = False Then
            Exit Sub
        End If
        ymsClass = ClassSemester.Items.Item(ClassSemesterCombobox.SelectedIndex)
        Try
            CatchClassThread = New Thread(AddressOf Me.RefreshClass)
            CatchClassThread.Start()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ScoreSemesterCombobox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ScoreSemesterCombobox.SelectedIndexChanged
        If ScoreFirst = False Then
            Exit Sub
        End If
        ymsScore = ScoreSemester.Items.Item(ScoreSemesterCombobox.SelectedIndex)
        Try
            CatchScoreThread = New Thread(AddressOf Me.RefreshScore)
            CatchScoreThread.Start()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub AbsentSemesterCombobox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles AbsentSemesterCombobox.SelectedIndexChanged
        If AbsentFirst = False Then
            Exit Sub
        End If
        ymsAbsent = AbsentSemester.Items.Item(AbsentSemesterCombobox.SelectedIndex)
        Try
            CatchAbsentThread = New Thread(AddressOf Me.RefreshAbsent)
            CatchAbsentThread.Start()
        Catch ex As Exception

        End Try
    End Sub
    Sub AddNewCreditItem(Subject As String, Teacher As String, Score As String, CreditInt As String, Credit As String, YearScore As String)
        Subject = Subject.Trim()
        Teacher = Teacher.Trim()
        Score = Score.Trim()
        CreditInt = CreditInt.Trim()
        Credit = Credit.Trim()
        YearScore = YearScore.Trim()
        If Subject = Nothing Then : Subject = "*" : End If
        If Teacher = Nothing Then : Teacher = "*" : End If
        If Score = Nothing Then : Score = "*" : End If
        If CreditInt = Nothing Then : CreditInt = "*" : End If
        If Credit = Nothing Then : Credit = "*" : End If
        If YearScore = Nothing Then : YearScore = "*" : End If

        Dim item As New ListViewItem
        item.Text = Subject
        item.SubItems.Add(Teacher)
        item.SubItems.Add(CreditInt)
        item.SubItems.Add(Credit)
        item.SubItems.Add(Score)
        item.SubItems.Add(YearScore)
        CreditListView.Items.Add(item)
    End Sub
    Dim RequiredCredit As Integer = 0
    Dim OptionalCredit As Integer = 0
    Dim General1 As Boolean = False
    Dim General2 As Boolean = False
    Dim General3 As Boolean = False
    Dim General4 As Boolean = False
    Dim General5 As Boolean = False
    Dim GeneralSociety As Boolean = False
    Dim GeneralHistorical As Boolean = False
    Dim GeneralTechnical As Boolean = False
    Dim CreditError As Boolean = False
    Dim CreditCheck As Boolean = False
    Sub RefreshCredit(ymsCredit As String)
        Dim ErrorCount As Integer = 0
        While (1)
            If CreditCheck = True Then
                Thread.Sleep(500)
                Continue While
            End If

            If ErrorCount >= 3 Then
                CreditError = True
                CreditNowCount += 1
                Exit While
            End If

            CreditCheck = True
            Try
                If ScoreReady = False Then
                    Exit Sub
                End If

                Dim parameters As IDictionary(Of String, String) = New Dictionary(Of String, String)()
                Dim response As HttpWebResponse
                Dim reader As StreamReader
                Dim respHTML As String
                Dim GetList(50, 6) As String

                parameters.Clear()
                parameters.Add("fncid", "AG009")
                If Not ymsCredit = Nothing Then
                    parameters.Add("sysyear", ymsCredit.Split(",")(0))
                    parameters.Add("syssms", ymsCredit.Split(",")(1))
                Else
                    CreditNowCount += 1
                    Exit While
                End If
                parameters.Add("std_id", "")
                parameters.Add("local_ip", "")
                parameters.Add("online", "okey")
                parameters.Add("loginid", userName)
                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/fnc.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()
                Dim doc As New HtmlDocument()
                doc.LoadHtml(respHTML)
                Dim node As HtmlNode = doc.DocumentNode
                Dim ls_randnum As String = WebUtility.HtmlDecode(node.SelectNodes("//input[9]")(0).Attributes.Item("value").Value)
                response.Close()

                parameters.Clear()
                If Not ymsCredit = Nothing Then
                    parameters.Add("yms", System.Uri.EscapeDataString(ymsCredit))
                    parameters.Add("arg01", ymsCredit.Split(",")(0))
                    parameters.Add("arg02", ymsCredit.Split(",")(1))
                Else
                    CreditNowCount += 1
                    Exit While
                End If
                parameters.Add("spath", "ag_pro/ag008.jsp?")
                Dim responsex As HttpWebResponse = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/ag_pro/ag008.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                Dim readerx As StreamReader = New StreamReader(responsex.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                Dim respHTMLx = readerx.ReadToEnd()

                If respHTMLx.Contains("無") Then
                    CreditCheck = False
                    CreditNowCount += 1
                    Exit Sub
                End If

                doc.LoadHtml(respHTMLx)
                node = doc.DocumentNode
                '<td>項次</td>
                '<td>科目名稱</td>
                '<td>學分數</td>
                '<td>授課時數</td>
                '<td>必選修</td>
                '<td>開課別</td>
                '<td>期中成績</td>
                '<td>學期成績</td>
                '<td>備註</td>
                Dim tmpRequiredCredit As Integer = 0
                Dim tmpOptionalCredit As Integer = 0
                For i = 2 To node.SelectNodes("//table[2]/tr").Count
                    GetList(i - 2, 0) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[2]")(0).InnerText).Trim
                    GetList(i - 2, 2) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[3]")(0).InnerText).Trim
                    GetList(i - 2, 3) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[5]")(0).InnerText).Trim
                    GetList(i - 2, 4) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[7]")(0).InnerText).Trim
                    GetList(i - 2, 5) = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[8]")(0).InnerText).Trim

                    If WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[5]")(0).InnerText).Trim.Contains("必修") Then
                        If Not WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[8]")(0).InnerText).Contains(".") Then
                            Continue For
                        End If
                        If Val(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[8]")(0).InnerText).Trim) >= 60 Or Val(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[3]")(0).InnerText).Trim) = 0 Then
                            TmpRequiredCredit += Val(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[3]")(0).InnerText).Trim)
                            Dim General As String = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[2]")(0).InnerText).Trim
                            If General.Contains("通識(一)") Then
                                General1 = True
                            ElseIf General.Contains("通識(二)") Then
                                General2 = True
                            ElseIf General.Contains("通識(三)") Then
                                General3 = True
                            ElseIf General.Contains("通識(四)") Then
                                General4 = True
                            ElseIf General.Contains("通識(五)") Then
                                General5 = True
                            ElseIf General.Contains("通識(社會)") Then
                                GeneralSociety = True
                            ElseIf General.Contains("通識(人文)") Then
                                GeneralHistorical = True
                            ElseIf General.Contains("通識(科技)") Then
                                GeneralTechnical = True
                            End If
                        End If
                    ElseIf WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[5]")(0).InnerText).Trim.Contains("選修") Then
                        If Not WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[8]")(0).InnerText).Contains(".") Or Val(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[3]")(0).InnerText).Trim) = 0 Then
                            Continue For
                        End If
                        If Val(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[8]")(0).InnerText).Trim) >= 60 Then
                            tmpOptionalCredit += Val(WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[3]")(0).InnerText).Trim)
                            Dim General As String = WebUtility.HtmlDecode(node.SelectNodes("//table[2]/tr[" & i & "]/td[2]")(0).InnerText).Trim
                            If General.Contains("通識(一)") Then
                                General1 = True
                            ElseIf General.Contains("通識(二)") Then
                                General2 = True
                            ElseIf General.Contains("通識(三)") Then
                                General3 = True
                            ElseIf General.Contains("通識(四)") Then
                                General4 = True
                            ElseIf General.Contains("通識(五)") Then
                                General5 = True
                            ElseIf General.Contains("通識(社會)") Then
                                GeneralSociety = True
                            ElseIf General.Contains("通識(人文)") Then
                                GeneralHistorical = True
                            ElseIf General.Contains("通識(科技)") Then
                                GeneralTechnical = True
                            End If
                        End If
                    End If
                Next
                response.Close()

                parameters.Clear()
                If Not ymsCredit = Nothing Then
                    parameters.Add("yms", System.Uri.EscapeDataString(ymsCredit))
                    parameters.Add("arg01", ymsCredit.Split(",")(0))
                    parameters.Add("arg02", ymsCredit.Split(",")(1))
                Else
                    If Now.Month < 9 Then
                        parameters.Add("yms", System.Uri.EscapeDataString((Now.Year - 1912) & ",2"))
                        parameters.Add("arg01", Now.Year - 1912)
                        parameters.Add("arg02", 2)
                    Else
                        parameters.Add("yms", System.Uri.EscapeDataString((Now.Year - 1911) & ",1"))
                        parameters.Add("arg01", Now.Year - 1911)
                        parameters.Add("arg02", 1)
                    End If
                End If
                parameters.Add("spath", "ag_pro%2Fag222.jsp%3F")
                response = HttpWebResponseUtility.CreatePostHttpResponse("http://140.127.113.231/kuas/ag_pro/ag222.jsp", parameters, Nothing, Nothing, Encoding.UTF8, cookies)
                reader = New StreamReader(response.GetResponseStream, System.Text.Encoding.GetEncoding("UTF-8"))
                respHTML = reader.ReadToEnd()

                doc.LoadHtml(respHTML)
                node = doc.DocumentNode
                For i = 2 To node.SelectNodes("//table[1]/tr").Count
                    For j = 0 To 49
                        If GetList(j, 0) = Nothing Then : Exit For : End If
                        If GetList(j, 0) = WebUtility.HtmlDecode(node.SelectNodes("//table[1]/tr[" & i & "]/td[2]")(0).InnerText).Trim Then
                            GetList(j, 1) = WebUtility.HtmlDecode(node.SelectNodes("//table[1]/tr[" & i & "]/td[10]")(0).InnerText).Trim
                            Exit For
                        End If
                    Next
                Next
                response.Close()

                For i = 0 To 49
                    If GetList(i, 0) = Nothing Then : Exit For : End If
                    If GetList(i, 1) = Nothing Then : GetList(i, 1) = "*" : End If
                    If GetList(i, 2) = Nothing Then : GetList(i, 2) = "*" : End If
                    If GetList(i, 3) = Nothing Then : GetList(i, 3) = "*" : End If
                    AddNewCreditItem(GetList(i, 0), GetList(i, 1), GetList(i, 4), GetList(i, 2), GetList(i, 3), GetList(i, 5))
                Next
                RequiredCredit += tmpRequiredCredit
                OptionalCredit += tmpOptionalCredit
                CreditNowCount += 1
                CreditCheck = False
                Exit While
            Catch ex As Exception
                ErrorCount += 1
                CreditCheck = False
            End Try
        End While
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MsgBox("是否進行教學評量自動填寫 ?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Button1.Enabled = False
            GroupBox1.Enabled = False
            GroupBox2.Enabled = False
            GroupBox3.Enabled = False
            GroupBox4.Enabled = False
            If (7 < Month(Now) And Month(Now) < 12) Or (2 < Month(Now) And Month(Now) < 5) Then
                AutoQThread = New Thread(AddressOf AutoQ)
                AutoQThread.Start()
            Else
                AutoQThread = New Thread(AddressOf AutoFinalQ)
                AutoQThread.Start()
            End If
        End If
    End Sub
    Dim ctl02 = 1
    Dim ctl04 = 1
    Dim ctl05 = 1
    Dim ctl06 = 1
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, RadioButton3.CheckedChanged, RadioButton4.CheckedChanged, RadioButton5.CheckedChanged, RadioButton6.CheckedChanged, RadioButton7.CheckedChanged, RadioButton8.CheckedChanged, RadioButton9.CheckedChanged, RadioButton10.CheckedChanged, RadioButton11.CheckedChanged, RadioButton12.CheckedChanged, RadioButton13.CheckedChanged, RadioButton14.CheckedChanged, RadioButton15.CheckedChanged, RadioButton16.CheckedChanged, RadioButton17.CheckedChanged, RadioButton18.CheckedChanged, RadioButton19.CheckedChanged, RadioButton20.CheckedChanged
        Dim RadioButtonx As RadioButton = CType(sender, RadioButton)
        Select Case Val(RadioButtonx.Name.Replace("RadioButton", ""))
            Case 1
                ctl02 = 1
            Case 2
                ctl02 = 2
            Case 3
                ctl04 = 1
            Case 4
                ctl04 = 2
            Case 5
                ctl04 = 3
            Case 6
                ctl04 = 4
            Case 7
                ctl05 = 1
            Case 8
                ctl05 = 2
            Case 9
                ctl05 = 3
            Case 10
                ctl05 = 4
            Case 11
                ctl05 = 5
            Case 12
                ctl05 = 6
            Case 13
                ctl05 = 7
            Case 14
                ctl05 = 8
            Case 15
                ctl05 = 9
            Case 16
                ctl06 = 1
            Case 17
                ctl06 = 2
            Case 18
                ctl06 = 3
            Case 19
                ctl06 = 4
            Case 20
                ctl06 = 5
        End Select
    End Sub
End Class
Public Class XmlSerialize
    Public Shared Sub SerializeToXml(FileName As String, [Object] As Object)
        Dim xml As XmlSerializer = Nothing
        Dim stream As Stream = Nothing
        Dim writer As StreamWriter = Nothing
        Try
            xml = New XmlSerializer([Object].[GetType]())
            stream = New FileStream(FileName, FileMode.Create, FileAccess.Write, FileShare.Read)
            writer = New StreamWriter(stream, Encoding.UTF8)
            xml.Serialize(writer, [Object])
        Catch ex As Exception
            Throw ex
        Finally
            writer.Close()
            stream.Close()
        End Try
    End Sub
    Public Shared Function DeserializeFromXml(Of T)(FileName As String) As T
        Dim xml As XmlSerializer = Nothing
        Dim stream As Stream = Nothing
        Dim reader As StreamReader = Nothing
        Try
            xml = New XmlSerializer(GetType(T))
            stream = New FileStream(FileName, FileMode.Open, FileAccess.Read, FileShare.None)
            reader = New StreamReader(stream, Encoding.UTF8)
            Dim obj As Object = xml.Deserialize(reader)
            If obj Is Nothing Then
                Return Nothing
            Else
                Return DirectCast(obj, T)
            End If
        Catch ex As Exception
            Throw ex
        Finally
            stream.Close()
            reader.Close()
        End Try
    End Function
End Class