<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Loginfrm
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意:  以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Loginfrm))
        Me.LoginBtn = New System.Windows.Forms.Button()
        Me.AccountLabel = New System.Windows.Forms.Label()
        Me.PasswordLabel = New System.Windows.Forms.Label()
        Me.User = New System.Windows.Forms.TextBox()
        Me.Pwd = New System.Windows.Forms.TextBox()
        Me.RememberCheckBox = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'LoginBtn
        '
        Me.LoginBtn.Enabled = False
        Me.LoginBtn.Font = New System.Drawing.Font("微軟正黑體", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.LoginBtn.Location = New System.Drawing.Point(305, 12)
        Me.LoginBtn.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.LoginBtn.Name = "LoginBtn"
        Me.LoginBtn.Size = New System.Drawing.Size(100, 25)
        Me.LoginBtn.TabIndex = 0
        Me.LoginBtn.Text = "Login"
        Me.LoginBtn.UseVisualStyleBackColor = True
        '
        'AccountLabel
        '
        Me.AccountLabel.AutoSize = True
        Me.AccountLabel.Location = New System.Drawing.Point(12, 15)
        Me.AccountLabel.Name = "AccountLabel"
        Me.AccountLabel.Size = New System.Drawing.Size(41, 16)
        Me.AccountLabel.TabIndex = 1
        Me.AccountLabel.Text = "學號 : "
        '
        'PasswordLabel
        '
        Me.PasswordLabel.AutoSize = True
        Me.PasswordLabel.Location = New System.Drawing.Point(12, 44)
        Me.PasswordLabel.Name = "PasswordLabel"
        Me.PasswordLabel.Size = New System.Drawing.Size(41, 16)
        Me.PasswordLabel.TabIndex = 2
        Me.PasswordLabel.Text = "密碼 : "
        '
        'User
        '
        Me.User.Location = New System.Drawing.Point(59, 12)
        Me.User.Name = "User"
        Me.User.Size = New System.Drawing.Size(237, 23)
        Me.User.TabIndex = 3
        '
        'Pwd
        '
        Me.Pwd.Location = New System.Drawing.Point(59, 41)
        Me.Pwd.Name = "Pwd"
        Me.Pwd.PasswordChar = Global.Microsoft.VisualBasic.ChrW(9679)
        Me.Pwd.Size = New System.Drawing.Size(237, 23)
        Me.Pwd.TabIndex = 4
        '
        'RememberCheckBox
        '
        Me.RememberCheckBox.AutoSize = True
        Me.RememberCheckBox.Location = New System.Drawing.Point(316, 44)
        Me.RememberCheckBox.Name = "RememberCheckBox"
        Me.RememberCheckBox.Size = New System.Drawing.Size(75, 20)
        Me.RememberCheckBox.TabIndex = 5
        Me.RememberCheckBox.Text = "記憶密碼"
        Me.RememberCheckBox.UseVisualStyleBackColor = True
        '
        'Loginfrm
        '
        Me.AcceptButton = Me.LoginBtn
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(417, 77)
        Me.Controls.Add(Me.RememberCheckBox)
        Me.Controls.Add(Me.Pwd)
        Me.Controls.Add(Me.User)
        Me.Controls.Add(Me.PasswordLabel)
        Me.Controls.Add(Me.AccountLabel)
        Me.Controls.Add(Me.LoginBtn)
        Me.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Loginfrm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "KUAS AP (By Silent)"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LoginBtn As System.Windows.Forms.Button
    Friend WithEvents AccountLabel As System.Windows.Forms.Label
    Friend WithEvents PasswordLabel As System.Windows.Forms.Label
    Friend WithEvents User As System.Windows.Forms.TextBox
    Friend WithEvents Pwd As System.Windows.Forms.TextBox
    Friend WithEvents RememberCheckBox As System.Windows.Forms.CheckBox

End Class
