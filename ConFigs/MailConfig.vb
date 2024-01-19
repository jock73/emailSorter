Imports System.Configuration
Imports System.IO
Imports System.Reflection
Imports Microsoft.IdentityModel.Protocols
Imports PareX.PareX

Public Class MailConfig

    Private newLicenseKey As String
    Public Property LicenseKey() As String
        Get
            Return newLicenseKey
        End Get
        Set(ByVal value As String)
            newLicenseKey = value
        End Set
    End Property

    Private newOAuthEnabled As Boolean
    Public Property OAuthEnabled() As Boolean
        Get
            Return newOAuthEnabled
        End Get
        Set(ByVal value As Boolean)
            newOAuthEnabled = value
        End Set
    End Property

    Private newExchangeClientTimeZone As String
    Public Property ExchangeClientTimeZone() As String
        Get
            Return newExchangeClientTimeZone
        End Get
        Set(ByVal value As String)
            newExchangeClientTimeZone = value
        End Set
    End Property


    Private newClientId As String
    Public Property ClientId() As String
        Get
            Return newClientId
        End Get
        Set(ByVal value As String)
            newClientId = value
        End Set
    End Property


    Private newClientSecret As String
    Public Property ClientSecret() As String
        Get
            Return newClientSecret
        End Get
        Set(ByVal value As String)
            newClientSecret = value
        End Set
    End Property

    Private newScope As String
    Public Property Scope() As String
        Get
            Return newScope
        End Get
        Set(ByVal value As String)
            newScope = value
        End Set
    End Property

    Private newExchangeWebServiceURL As String
    Public Property ExchangeWebServiceURL() As String
        Get
            Return newExchangeWebServiceURL
        End Get
        Set(ByVal value As String)
            newExchangeWebServiceURL = value
        End Set
    End Property

    Private newTenantId As String
    Public Property TenantId() As String
        Get
            Return newTenantId
        End Get
        Set(ByVal value As String)
            newTenantId = value
        End Set
    End Property


    Private newServiceAccountEmail As String
    Public Property ServiceAccountEmail() As String
        Get
            Return newServiceAccountEmail
        End Get
        Set(ByVal value As String)
            newServiceAccountEmail = value
        End Set
    End Property

    Private newMaxEmails As Integer
    Public Property MaxEmails() As Integer
        Get
            Return newMaxEmails
        End Get
        Set(ByVal value As Integer)
            newMaxEmails = value
        End Set
    End Property


    Private newSaveFileLocalLocation As String
    Public Property SaveFileLocalLocation() As String
        Get
            Return newSaveFileLocalLocation
        End Get
        Set(ByVal value As String)
            newSaveFileLocalLocation = value
        End Set
    End Property


    Private newSubjectJsonReplaceFileLocation As String
    Public Property SubjectJsonReplaceFileLocation() As String
        Get
            Return newSubjectJsonReplaceFileLocation
        End Get
        Set(ByVal value As String)
            newSubjectJsonReplaceFileLocation = value
        End Set
    End Property


    Private newCleanSubject As String
    Public Property CleanSubject() As String
        Get
            Return newCleanSubject
        End Get
        Set(ByVal value As String)
            newCleanSubject = value
        End Set
    End Property

    Private NewPauseConsole As Boolean
    Public Property PauseConsole() As Boolean
        Get
            Return NewPauseConsole
        End Get
        Set(ByVal value As Boolean)
            NewPauseConsole = value
        End Set
    End Property

    Private NewDeleteDeleteinbox As Integer
    Public Property DeleteDeleteinbox() As Integer
        Get
            Return NewDeleteDeleteinbox
        End Get
        Set(ByVal value As Integer)
            NewDeleteDeleteinbox = value
        End Set
    End Property

    Public Sub New()
        '' colin todo, some of these parms should be in config
        Me.ServiceAccountEmail = "partsupload@parex.nl"
        Me.ExchangeWebServiceURL = "https://outlook.office365.com/EWS/Exchange.asmx"
        Me.Scope = "https://outlook.office365.com/.default"
        Me.TenantId = "TenantId"
        Me.ClientSecret = "ClientSecret"
        Me.ClientId = "ClientId"
        Me.ExchangeClientTimeZone = ""
        Me.OAuthEnabled = True
        Me.LicenseKey = "LicenseKey"


        Me.CleanSubject = System.Configuration.ConfigurationManager.AppSettings("CleanSubject")
        Me.MaxEmails = CInt(System.Configuration.ConfigurationManager.AppSettings("MaxEmails"))
        Me.SaveFileLocalLocation = System.Configuration.ConfigurationManager.AppSettings("SaveFileLocalLocation")
        Me.SubjectJsonReplaceFileLocation = System.Configuration.ConfigurationManager.AppSettings("SubjectJsonReplaceFileLocation") '' "D:\DEV\EmailSorter2FA\PareX\Config\subjectedit.json"

        Me.PauseConsole = System.Configuration.ConfigurationManager.AppSettings("PauseConsole")



        Me.DeleteDeleteinbox = System.Configuration.ConfigurationManager.AppSettings("deleteinboxpercentage")

        Write(SubjectJsonReplaceFileLocation)



        Write("Load Email Configuration of: " & ServiceAccountEmail)



    End Sub





End Class
