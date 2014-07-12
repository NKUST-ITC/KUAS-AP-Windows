' Developer - Silent
' Date Created - 02/24/2014
'
' General Description - Simple Config class used for Business Object examples

Public Class Config

#Region " Modular Variables "

    Private _Account As String
    Private _Pwd As String
    Private _Remember As Boolean = True
    Private _Manager As Guid

#End Region

#Region " Constructors "

    Public Sub New()

    End Sub

#End Region

#Region " Public Properties "

    ''' <summary>
    ''' Property to hold the Account of the config
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>Config Account</returns>
    Public Property Account() As String
        Get
            Return _Account
        End Get
        Set(ByVal value As String)
            _Account = value
            'Console.WriteLine("Property {0} has been changed.", "Account")
        End Set
    End Property

    ''' <summary>
    ''' Property to hold the Pwd of the config
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>Config Pwd</returns>
    Public Property Pwd() As String
        Get
            Return _Pwd
        End Get
        Set(ByVal value As String)
            _Pwd = value
            'Console.WriteLine("Property {0} has been changed.", "Pwd")
        End Set
    End Property

    ''' <summary>
    ''' Property to hold the Remember of the config
    ''' </summary>
    ''' <value>Boolean</value>
    ''' <returns>Config Remember</returns>
    Public Property Remember() As Boolean
        Get
            Return _Remember
        End Get
        Set(ByVal value As Boolean)
            _Remember = value
            'Console.WriteLine("Property {0} has been changed.", "Remember")
        End Set
    End Property

    ''' <summary>
    ''' Holds config managers ID
    ''' </summary>
    ''' <value>Guid</value>
    ''' <returns>Config managers Guid</returns>
    Public Property Manager() As Guid
        Get
            Return _Manager
        End Get
        Set(ByVal value As Guid)
            _Manager = value
            'Console.WriteLine("Property {0} has been changed.", "Manager")
        End Set
    End Property

#End Region

End Class
