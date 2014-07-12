Imports System.IO

Public Class Json
    Dim ProgramName As String = "KUAS AP.exe"
    Private Sub Json_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If File.Exists(Application.StartupPath & "\Newtonsoft.Json_Ex.dll") Then
            Try
                My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\Newtonsoft.Json.dll")
                FileSystem.Rename(Application.StartupPath & "\Newtonsoft.Json_Ex.dll", "Newtonsoft.Json.dll")
            Catch ex As Exception

            End Try
        End If
        Dim myProcess As Process = System.Diagnostics.Process.Start(Application.StartupPath & "\" & ProgramName)
        End
    End Sub
End Class
