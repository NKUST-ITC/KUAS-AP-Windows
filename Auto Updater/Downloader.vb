Public Class Downloader '下載器

    '全域物件
    Private WithEvents Client As New Net.WebClient() '宣告Client為WebClient物件，且帶有「事件」。
    Private TimerProgress As New Timer '宣告TimerProgress為Timer實體

    '字串常數
    Private Const DEFAULT_NONE_NAMED_NAME As String = "未命名-it-easy.tw" '預設檔名

    '字串變數
    Private _TargetURL As String '儲存下載的目標網址
    Private _SaveFolderPath As String '儲存檔案存放的資料夾路徑
    Private _SaveFileName As String '儲存檔案存放的檔名

    '字串不必變數
    Private DATE_DAY As String = " 天 " '天
    Private DATE_HOUR As String = " 小時 " '小時
    Private DATE_MINUTE As String = " 分鐘 " '分
    Private DATE_SECOND As String = " 秒" '秒
    Private DATE_LINK As String = "又 " '又

    '全域變數
    Private StartTime As Long '儲存開始下載的時間
    Private EndTime As Long '儲存結束下載的時間
    Private BufferTime As Long '暫存下載的時間
    Private BufferTimeBytes As Long '暫存下載的位元組數


    '全域旗標
    Private Status As Short = -2 '儲存下載狀態。如果傳回-2，代表尚未開始下載。傳回-1，代表下載正在進行中。傳回0，代表下載失敗。傳回1，代表下載成功。傳回2，代表中斷下載。

    '全域下載資訊
    Private BytesReceived As Long '儲存已下載的檔案大小
    Private TotalBytesToReceive As Long '儲存總共要下載的檔案大小
    Private ProgressPercentage As Integer '儲存目前的下載進度百分比(整數)
    Private BytesEverySecond As Long '儲存一秒內所傳的位元組數


    '-------------------------------------------------------------------建構子-------------------------------------------------------------------

    Public Sub New(ByVal InputTargetURL As String) 'Downloader建構子(多載一)。引數為InputTargetURL，載入下載的目標網址。
        Me.TargetURL = InputTargetURL '設定網址
        Call CreatSaveFileName() '簡單設定網址中提供的檔名
        Me.SaveFolderPath = Me.GetMyFolderPath() '預設檔案存放的資料夾路徑為程式本身的資料夾路徑
    End Sub

    Public Sub New(ByVal InputTargetURL As String, ByVal InputSaveFolderPath As String) 'Downloader建構子(多載二)。引數為InputTargetURL和InputTargetURL，載入下載的目標網址和檔案存放的資料夾路徑。
        Me.TargetURL = InputTargetURL '設定網址
        Me.SaveFolderPath = InputSaveFolderPath '設定路徑

        Call CreatSaveFileName() '簡單設定網址中提供的檔名
    End Sub

    Public Sub New(ByVal InputTargetURL As String, ByVal InputSaveFolderPath As String, ByVal InputSaveFileName As String) 'Downloader建構子(多載三)。引數為InputTargetURL、InputTargetURL和SaveFileName，載入下載的目標網址和檔案存放的資料夾路徑，還有檔案存放的名稱。
        Me.TargetURL = InputTargetURL '設定網址
        Me.SaveFolderPath = InputSaveFolderPath '設定路徑
        Me.SaveFileName = InputSaveFileName '設定檔名
    End Sub

    '--------------------------------------------------------------------------------------------------------------------------------------------

    '--------------------------------------------------------------------------------------------------------------------------------------------

    '-------------------------------------------------------------------物件屬性-----------------------------------------------------------------

    Public Property TargetURL() As String '取得或設定下載的目標網址
        Set(ByVal InputTargetURL As String)
            If Status <> -1 Then '若不為下載狀態
                _TargetURL = InputTargetURL '設定網址
            End If
        End Set

        Get
            Return _TargetURL '取得網址
        End Get
    End Property

    Public Property SaveFileName() As String '取得或設定檔案存放的檔名
        Set(ByVal InputSaveFileName As String)
            If Status <> -1 Then '若不為下載狀態
                '先檢查不合法字元，若遇到不合法字元自動刪除之
                Dim IllegalChar() As String = {"/", "", ":", "*", "?", """", "|", "<", ">"}
                For Each CheckChars As String In IllegalChar
                    InputSaveFileName = Strings.Replace(InputSaveFileName, CheckChars, "") '將不合法字元取代為空字串
                Next
                If InputSaveFileName = "" Then '若根本沒輸入檔名
                    InputSaveFileName = DEFAULT_NONE_NAMED_NAME '檔名為預設名稱
                End If


                _SaveFileName = InputSaveFileName '設定檔名
            End If
        End Set
        Get
            Return _SaveFileName '取得檔名
        End Get
    End Property

    Public Property SaveFolderPath() As String '取得或設定檔案存放的資料夾路徑
        Set(ByVal InputSaveFolderPath As String)
            If Status <> -1 Then '若不為下載狀態
                InputSaveFolderPath = CorrectFolderPath(InputSaveFolderPath) '更正路徑
                If IO.Directory.Exists(InputSaveFolderPath) Then '若資料夾路徑存在
                    _SaveFolderPath = InputSaveFolderPath '設定存放的資料夾路徑
                Else '若資料夾路徑不存在
                    _SaveFolderPath = Me.GetMyFolderPath
                End If
            End If
        End Set

        Get
            Return _SaveFolderPath '取得存放的資料夾路徑
        End Get
    End Property

    Public Property SaveFullPath() As String '取得或設定檔案存放的完整路徑
        Set(ByVal InputSaveFullPath As String)
            If Status <> -1 Then '若不為下載狀態
                Dim DividePath() As String = Strings.Split(InputSaveFullPath, "") '分割路徑
                If DividePath.Count > 0 Then '若有一個以上的反斜線
                    Dim FolderPathBuffer As String = Strings.Replace(InputSaveFullPath, DividePath(DividePath.Count - 1), "") '將檔名取代為空字串
                    Me.SaveFolderPath = FolderPathBuffer
                    Me.SaveFileName = DividePath(DividePath.Count - 1)
                Else '格式有誤
                    Me.SaveFolderPath = Me.GetMyFolderPath '資料夾路徑為程式自身路徑
                    Me.SaveFileName = DEFAULT_NONE_NAMED_NAME '檔名為預設名稱
                End If
            End If
        End Set

        Get
            Return Me.SaveFolderPath & Me.SaveFileName  '傳回檔案存放的完整路徑
        End Get
    End Property

    '--------------------------------------------------------------------------------------------------------------------------------------------

    '--------------------------------------------------------------------------------------------------------------------------------------------

    '---------------------------------------------------------------------函式-------------------------------------------------------------------

    Public Function CheckDownloading() As Boolean '檢查是否正在下載中
        If Status = -1 Then '若正在下載
            Return True  '傳回True布林值
        Else
            Return False '傳回False布林值
        End If
    End Function

    Public Function GetStatus() As Short '檢查下載器的狀態
        Return Status '傳回短整數
    End Function

    Public Function ChoseFolderPath() As String '開啟FolderBrowserDialog，讓使用者選取資料夾
        Dim FBD As New FolderBrowserDialog
        FBD.ShowDialog() '顯示資料夾選取視窗方塊

        Dim FolderPath As String = FBD.SelectedPath '取得選取路徑
        If FolderPath = "" Then '如果沒有選擇路徑
            FolderPath = Me.GetMyFolderPath() '等於自身路徑
        End If

        Return Me.CorrectFolderPath(FolderPath)
    End Function

    Public Function ChoseFullPath() As String '開啟SaveFileDialog，讓使用者選取檔案存放路徑
        Dim SFD As New SaveFileDialog
        SFD.FileName = Me.SaveFileName '暫時預設檔名在文字方塊內
        SFD.ShowDialog() '顯示檔案儲存路徑方塊

        Dim FullPath As String = SFD.FileName '取得選取路徑
        If FullPath = "" Then '如果沒有選擇路徑
            FullPath = Me.GetMyFolderPath() & DEFAULT_NONE_NAMED_NAME '等於自身路徑加上預設檔名
        ElseIf FullPath = Me.SaveFileName Then '若按下取消
            FullPath = Me.SaveFolderPath & Me.SaveFileName '不改變
        End If

        Return FullPath
    End Function

    Public Function GetMyFolderPath() As String '取得程式本身所在的資料夾路徑(不包括檔名)
        Return CorrectFolderPath(Application.StartupPath) '傳回程式本身的資料夾路徑
    End Function

    Public Function GetBytesReceived() As Long '取得已下載的檔案大小
        If Math.Abs(Status) = 1 Then '如果正在下載或是下載完成
            Return BytesReceived
        Else
            Return 0
        End If
    End Function

    Public Function GetTotalBytesToReceive() As Long '取得總共要下載的檔案大小
        If Math.Abs(Status) = 1 Then '如果正在下載或是下載完成
            Return TotalBytesToReceive
        Else
            Return 0
        End If
    End Function

    Public Function GetProgressPercentage() As Integer '取得目前的下載進度百分比(整數)
        If Math.Abs(Status) = 1 Then '如果正在下載或是下載完成
            Return ProgressPercentage
        Else
            Return 0
        End If
    End Function

    Public Function GetDoubleProgressPercentage(ByVal DecimalCount As Integer) As Double '取得目前的下載進度百分比(浮點數)，傳入小數位數
        If Math.Abs(Status) = 1 Then '如果正在下載或是下載完成
            Return Math.Round((Me.GetBytesReceived / Me.GetTotalBytesToReceive) * 100, DecimalCount)
        Else
            Return 0
        End If
    End Function

    Public Function GetEverySecondSpead() As Long '取得每秒所傳送的位元組數
        If Status = -1 Then '如果正在下載
            Return BytesEverySecond
        Else
            Return 0
        End If
    End Function

    Public Function GetLastBytes() As Long '取得剩餘的位元組數
        If Status = -1 Then '如果正在下載
            Return Me.GetTotalBytesToReceive - Me.GetBytesReceived
        Else
            Return 0
        End If
    End Function

    Public Function FormatBytes(ByVal InputBytes As Long, ByVal DecimalCount As Integer) As String '格式化顯示(多載一)。只傳入要格式化的位元組數字和小數位數。算是一種自動化設定。
        '從TB開始判斷到B
        Dim Unit As String '單位
        If InputBytes > Math.Pow(2, 40) Then '若大於1TB，則以TB顯示
            Unit = "TB"
        ElseIf InputBytes > Math.Pow(2, 30) Then '若大於1GB，則以GB顯示
            Unit = "GB"
        ElseIf InputBytes > Math.Pow(2, 20) Then '若大於1MB，則以MB顯示
            Unit = "MB"
        ElseIf InputBytes > Math.Pow(2, 10) Then '若大於1KB，則以KB顯示
            Unit = "KB"
        ElseIf InputBytes >= 0 Then '若大於0B，則以B顯示
            Unit = "Bytes"
        Else
            Unit = " -Error-"
        End If
        Return TransformBytes(InputBytes, DecimalCount, Unit) & " " & Unit
    End Function

    Public Function FormatBytes(ByVal InputBytes As Long, ByVal DecimalCount As Integer, ByVal FormatUnit As String) As String '格式化顯示(多載二)。傳入要格式化的位元組數字、小數位數和固定單位。
        Return TransformBytes(InputBytes, DecimalCount, FormatUnit) & " " & FormatUnit
    End Function

    Public Function TransformBytes(ByVal InputBytes As Long, ByVal DecimalCount As Integer, ByVal FormatUnit As String) As Double '轉換Bytes大小格式。
        FormatUnit = Strings.UCase(FormatUnit) '轉成大寫來判斷
        If Strings.Right(FormatUnit, 1) <> "B" Then '若結尾不是B
            FormatUnit = FormatUnit & "B" '加B
        End If

        Dim Divisor As Long '除數
        Select Case FormatUnit '選擇判斷
            Case "TB"
                Divisor = Math.Pow(2, 40)
            Case "GB"
                Divisor = Math.Pow(2, 30)
            Case "MB"
                Divisor = Math.Pow(2, 20)
            Case "KB"
                Divisor = Math.Pow(2, 10)
            Case "B", "BYTEB", "BYTESB", "位元組B", "個位元組B"
                Divisor = 1
            Case Else
                Return 0
        End Select

        Return Math.Round(InputBytes / Divisor, DecimalCount) '傳回運算值
    End Function

    Public Function GetDownloadingTime() As Long '取得下載的已用時間
        If Status = -1 Then '如果正在下載
            Return Environment.TickCount - StartTime '傳回目前時間減掉開始時間
        Else
            Return 0
        End If
    End Function

    Public Function GetDownloadedTime() As Long '取得整個下載進度的所用時間
        If Status = 1 Then '如果成功下載
            Return EndTime - StartTime '傳回結束時間減掉開始時間
        Else
            Return 0
        End If
    End Function

    Public Function GetLastTime() As Long '取得預估的剩餘時間
        If Status = -1 Then '如果正在下載
            Return (Me.GetLastBytes / Me.GetEverySecondSpead) * 1000 '轉成毫秒
        Else
            Return 0
        End If
    End Function

    Public Function FormatTime(ByVal TimeGap As Long) As String '轉換時間格式
        Dim iDay As Integer, iHour As Integer, iMinute As Integer, iSecond As Integer
        TimeGap = TimeGap / 1000 '毫秒轉換成秒
        '天
        While TimeGap >= 86400
            iDay += 1
            TimeGap -= 86400
        End While

        '時
        While TimeGap >= 3600
            iHour += 1
            TimeGap -= 3600
        End While

        '分
        While TimeGap >= 60
            iMinute += 1
            TimeGap -= 60
        End While

        '秒
        iSecond = TimeGap

        If iDay > 0 Then
            Return iDay & DATE_DAY & DATE_LINK & iHour & DATE_HOUR & iMinute & DATE_MINUTE & iSecond & DATE_SECOND
        ElseIf iHour > 0 Then
            Return iHour & DATE_HOUR & DATE_LINK & iMinute & DATE_MINUTE & iSecond & DATE_SECOND
        ElseIf iMinute > 0 Then
            Return iMinute & DATE_MINUTE & DATE_LINK & iSecond & DATE_SECOND
        ElseIf iSecond >= 0 Then
            Return iSecond & DATE_SECOND
        Else
            Return ""
        End If
    End Function


    Public Function CreateFolder(ByVal FolderPath As String) As Boolean '建立資料夾
        If IO.Directory.Exists(FolderPath) = False Then '若資料夾路徑不存在
            On Error Resume Next '如果遇到錯誤繼續執行
            IO.Directory.CreateDirectory(FolderPath) '建立資料夾
        End If

        Return IO.Directory.Exists(FolderPath) '建立完後若資料夾存在則傳回True；否則傳回False
    End Function

    Private Function CorrectFolderPath(ByVal InputFolderPath As String) As String '更正資料夾路徑格式
        If Strings.Right(InputFolderPath, 1) <> "" Then '判斷最後一個字元是否為""，若不是，則添加上去
            InputFolderPath = InputFolderPath & ""
        End If
        Return InputFolderPath
    End Function

    '--------------------------------------------------------------------------------------------------------------------------------------------

    '--------------------------------------------------------------------------------------------------------------------------------------------

    '--------------------------------------------------------------------副程式------------------------------------------------------------------

    Public Sub StartDownload() '開始進行下載程序
        On Error Resume Next '如果遇到錯誤繼續執行
        If Me.CheckDownloading = False Then '若目前沒有正在下載
            If IO.File.Exists(Me.SaveFullPath) Then IO.File.Delete(Me.SaveFullPath) '如果儲存路徑有殘留的檔案存在，刪除它
            Client.DownloadFileAsync(New Uri(TargetURL), Me.SaveFullPath) '開始下載，且不封鎖呼叫執行緒
            Status = -1 '正在下載
            StartTime = Environment.TickCount '儲存開始下載的時間點
            Debug.Print("正在從「{0}」取得資料...", Me.TargetURL)
        End If
    End Sub

    Public Sub StopDownloading() '停止下載程序
        If Status = -1 Then '如果目前是正在下載的狀態
            Status = 2
            Client.CancelAsync() '中斷下載
            Debug.Print("終止下載！")
        End If
    End Sub

    Private Sub TimerGate() '內部使用的Timer控制閘
        TimerProgress.Enabled = Me.CheckDownloading '開關Timer
    End Sub

    Public Sub TimerGate(ByRef TimerControl As Timer) 'Timer控制閘(多載一)，只傳入一個Timer物件實體
        TimerProgress = TimerControl '將TimerProgress參考到TimerControl
    End Sub

    Public Sub TimerGate(ByRef TimerControl As Timer, ByVal Interval As Long) 'Timer控制閘(多載二)，傳入一個Timer物件實體和時間
        TimerProgress = TimerControl '將TimerProgress參考到TimerControl
        TimerProgress.Interval = Interval '設定時間
    End Sub

    Private Sub CreatSaveFileName() '簡單設定網址中提供的檔名
        Dim DivideURL() As String = Strings.Split(Me.TargetURL, "/")
        If Strings.InStrRev(DivideURL(DivideURL.Count - 1), ".") > 0 Then '網址最後一段的字串中倒數搜尋是否有副檔名專用的「.」，若有則作為檔名。
            Me.SaveFileName = DivideURL(DivideURL.Count - 1)
        Else '若沒有則用預設名稱
            Me.SaveFileName = DEFAULT_NONE_NAMED_NAME
        End If
    End Sub

    Private Sub Client_DownloadProgressChanged(ByVal sender As Object, ByVal e As System.Net.DownloadProgressChangedEventArgs) Handles Client.DownloadProgressChanged '當Client正在下載時
        Call TimerGate() '內部使用的Timer控制閘

        '時間
        If Environment.TickCount - BufferTime > 1000 Then '若時間超過一秒
            BytesEverySecond = e.BytesReceived - BufferTimeBytes
            BufferTime = Environment.TickCount '等於目前時間
            BufferTimeBytes = e.BytesReceived '等於目前所傳的數
        End If

        BytesReceived = e.BytesReceived '儲存已下載的檔案大小
        TotalBytesToReceive = e.TotalBytesToReceive '儲存總共要下載的檔案大小
        ProgressPercentage = e.ProgressPercentage '儲存目前的下載進度百分比(整數)
        Debug.Print("已收到/總共：{0}/{1}　完成進度：{2}%　已使用：{3}　速度：{4}/s", Me.FormatBytes(Me.GetBytesReceived, 2), Me.FormatBytes(Me.GetTotalBytesToReceive, 2), Me.GetDoubleProgressPercentage(2), Me.FormatTime(Me.GetDownloadingTime), Me.FormatBytes(Me.GetEverySecondSpead, 2))
    End Sub

    Private Sub Client_DownloadFileCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs) Handles Client.DownloadFileCompleted  '當Clien結束下載
        On Error Resume Next '如果遇到錯誤繼續執行

        If Status = -1 Then
            If IO.File.Exists(Me.SaveFullPath) AndAlso My.Computer.FileSystem.GetFileInfo(SaveFullPath).Length > 0 AndAlso My.Computer.FileSystem.GetFileInfo(SaveFullPath).Length = Me.GetTotalBytesToReceive Then '若檔案存在，且大小都和網路上的相同。表示下載成功
                Status = 1 '下載成功
                EndTime = Environment.TickCount '儲存結束時間
                Debug.Print("下載成功！")
            Else
                Status = 0 '下載失敗
                Debug.Print("下載失敗！")
            End If
        End If

        If Status <> 1 Then '若下載不成功
            If IO.File.Exists(Me.SaveFullPath) Then IO.File.Delete(Me.SaveFullPath) '如果有殘留的檔案存在，刪除它
        End If
        Call TimerGate() '內部使用的Timer控制閘
    End Sub

    Public Sub SetTimeFormat(ByVal DayString As String, ByVal HourString As String, ByVal MinuteString As String, ByVal SecondString As String) '設定時間格式
        DATE_DAY = DayString
        DATE_HOUR = HourString
        DATE_MINUTE = MinuteString
        DATE_SECOND = SecondString
    End Sub

    '--------------------------------------------------------------------------------------------------------------------------------------------

    '--------------------------------------------------------------------------------------------------------------------------------------------

End Class
