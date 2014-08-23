Imports System.IO

Public Class Json
    Private Sub Json_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Try
            If File.Exists(Application.StartupPath & "\Json.exe") Then
                Try
                    My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\Json.exe")
                Catch ex As Exception

                End Try
            End If
            If File.Exists(Application.StartupPath & "\Newtonsoft.Json_Ex.dll") Then
                Try
                    My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\Newtonsoft.Json.dll")
                    FileSystem.Rename(Application.StartupPath & "\Newtonsoft.Json_Ex.dll", "Newtonsoft.Json.dll")
                Catch ex As Exception

                End Try
            End If
        Catch ex As Exception

        End Try
        Dim Loginfrm As New Loginfrm
        Loginfrm.Show()
        Me.Hide()
    End Sub
End Class
