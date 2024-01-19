Public Class FTPconfig


    Private newftpenable As Boolean
    Public Property ftpenable() As Boolean
        Get
            Return newftpenable
        End Get
        Set(ByVal value As Boolean)
            newftpenable = value
        End Set
    End Property

    Private newftpIP As String
    Public Property ftpIP() As String
        Get
            Return newftpIP
        End Get
        Set(ByVal value As String)
            newftpIP = value
        End Set
    End Property



    Private Newftpusername As String
    Public Property ftpusername() As String
        Get
            Return Newftpusername
        End Get
        Set(ByVal value As String)
            Newftpusername = value
        End Set
    End Property


    Private newftppassword As String
    Public Property ftppassword() As String
        Get
            Return newftppassword
        End Get
        Set(ByVal value As String)
            newftppassword = value
        End Set
    End Property

    Private newftppath As String
    Public Property ftppath() As String
        Get
            Return newftppath
        End Get
        Set(ByVal value As String)
            newftppath = value
        End Set
    End Property




    Private newExtentionsAllowed As String
    Public Property ExtentionsAllowed() As String
        Get
            Return newExtentionsAllowed
        End Get
        Set(ByVal value As String)
            newExtentionsAllowed = value
        End Set
    End Property

    Public Sub New()

        Me.ftpenable = True

        Me.ftpIP = "ftp.p.nl"
        Me.ftpusername = "Email"
        Me.ftppassword = "pword"
        Me.ftppath = "/FTP-DealerUpload/Dealers/"
        Me.ExtentionsAllowed = System.Configuration.ConfigurationManager.AppSettings("ExtentionsAllowed")



    End Sub


End Class









