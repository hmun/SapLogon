' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports System.Collections.ObjectModel
Imports SAP.Middleware.Connector

Public Class SAPCon

    Const aParamWs As String = "Parameter"
    Const aConnectionWs As String = "SAP-Con"
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private aSapExcelDestinationConfiguration As SapExcelDestinationConfiguration
    Private _destination As RfcCustomDestination
    Private _connected As Boolean = False

    Private _dest As String = ""
    Private _client As String = ""
    Private _username As String = ""
    Private _password As String = ""
    Private _language As String = ""
    Private _sncmyname As String = ""

    Public Property Connected() As Boolean
        Get
            Return _connected
        End Get
        Set(ByVal value As Boolean)
            _connected = value
        End Set
    End Property

    Public Property Destination() As RfcCustomDestination
        Get
            Return _destination
        End Get
        Set(ByVal value As RfcCustomDestination)
            _destination = value
        End Set
    End Property

    Public Property Dest() As String
        Get
            Return _dest
        End Get
        Set(ByVal value As String)
            _dest = value
        End Set
    End Property

    Public Property Client() As String
        Get
            Return _client
        End Get
        Set(ByVal value As String)
            _client = value
        End Set
    End Property

    Public Property Username() As String
        Get
            Return _username
        End Get
        Set(ByVal value As String)
            _username = value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return _password
        End Get
        Set(ByVal value As String)
            _password = value
        End Set
    End Property

    Public Property Language() As String
        Get
            Return _language
        End Get
        Set(ByVal value As String)
            _language = value
        End Set
    End Property

    Public Property SncMyName() As String
        Get
            Return _sncmyname
        End Get
        Set(ByVal value As String)
            _sncmyname = value
        End Set
    End Property

    Public Sub New(pAssembly As String, ByRef pConParameter As ConParameter)
        Dim parameters As New RfcConfigParameters()

        log.Debug("New - " & "aDest=" & CStr(_dest))
        log.Debug("New - " & "setting up aSapExcelDestinationConfiguration")
        aSapExcelDestinationConfiguration = New SapExcelDestinationConfiguration
        Try
            aSapExcelDestinationConfiguration.ConfigAddOrChangeDestination(pAssembly)
        Catch Ex As System.Exception
            log.Error("New ConfigAddOrChangeDestination - Exception=" & Ex.ToString)
        End Try

        aSapExcelDestinationConfiguration.ExcelAddOrChangeDestination(pConParameter)
        aSapExcelDestinationConfiguration.SetUp()
        log.Debug("New - " & "finished setting up aSapExcelDestinationConfiguration")
        ' log.Debug("New - " & "calling setDest")
        log.Debug("New - " & "end")
    End Sub

    Public Function setDestination() As Integer
        Dim dest As RfcDestination = Nothing
        If _destination Is Nothing And Not _dest = "" Then
            Try
                log.Debug("setDestination - " & "getting dest from RfcDestinationManager")
                dest = RfcDestinationManager.GetDestination(_dest)
                log.Debug("setDestination - " & "creating destination")
                _destination = dest.CreateCustomDestination()
                log.Debug("setDestination - " & "using destination.Name=" & Destination.Name)
                setDestination = 0
            Catch Ex As System.Exception
                setDestination = 16
                log.Error("setDestination - Exception=" & Ex.ToString)
                Exit Function
            End Try
        Else
            setDestination = 4
        End If
    End Function
    Public Function checkCon(Optional pDoPing As Boolean = True) As Integer
        Dim formRet = 0
        If Not Connected And _destination.SncMode = "1" Then
            log.Debug("checkCon - " & "connecting using SNC destination")
            setCredentials_SNC(_client, _language, _sncmyname, _username)
        ElseIf Not Connected Then
            log.Debug("checkCon - " & "connecting using regular destination")
            setCredentials(_client, _username, _password, _language)
        End If
        If connected Or pDoPing Then
            Try
                log.Debug("checkCon - " & "calling destination.Ping")
                _destination.Ping()
                connected = True
                checkCon = 0
            Catch ex As RfcInvalidParameterException
                clearCredentials()
                Connected = False
                log.Error("checkCon - Exception=" & ex.ToString)
                Throw ex
            Catch ex As RfcBaseException
                clearCredentials()
                Connected = False
                log.Error("checkCon - Exception=" & ex.ToString)
                Throw ex
            End Try
        Else
            log.Debug("checkCon - " & "failed to connect")
            connected = False
            Destination = Nothing
            checkCon = 8
        End If
    End Function

    Public Sub setCredentials_SNC(aClient As String, aLanguage As String, Optional aSncMyName As String = "", Optional aUsername As String = "")
        log.Debug("setCredentials_SNC - " & "setting credentials")
        Try
            Destination.Client = aClient
            Destination.Language = aLanguage
            If Not String.IsNullOrEmpty(aSncMyName) Then
                Destination.SncMyName = aSncMyName
            End If
            If Not String.IsNullOrEmpty(aUsername) Then
                Destination.User = aUsername
            End If
            log.Debug("setCredentials_SNC - " & " Destination.Client=" & Destination.Client)
            log.Debug("setCredentials_SNC - " & " Destination.Language=" & Destination.Language)
            log.Debug("setCredentials_SNC - " & " Destination.User=" & Destination.User)
            log.Debug("setCredentials_SNC - " & " Destination.SncMyName=" & Destination.SncMyName)
        Catch ex As System.Exception
            '            MsgBox("setCredentials failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
            log.Error("setCredentials_SNC - Exception=" & ex.ToString)
        End Try
    End Sub

    Public Sub setCredentials(aClient As String, aUsername As String, aPassword As String, aLanguage As String)
        log.Debug("setCredentials - " & "setting credentials")
        Try
            Destination.Client = aClient
            Destination.User = aUsername
            Destination.Password = aPassword
            Destination.Language = aLanguage
        Catch ex As System.Exception
            '            MsgBox("setCredentials failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
            log.Error("setCredentials - Exception=" & ex.ToString)
        End Try
    End Sub

    Public Sub SAPlogoff()
        log.Debug("SAPlogoff - " & "closing connection")
        Destination = Nothing
        If _dest IsNot Nothing And _dest <> "" Then
            log.Debug("SAPlogoff - " & "calling aSapExcelDestinationConfiguration.TearDown, aDest=" & _dest)
            aSapExcelDestinationConfiguration.TearDown(_dest)
        Else
            log.Debug("SAPlogoff - " & "calling aSapExcelDestinationConfiguration.TearDown")
            aSapExcelDestinationConfiguration.TearDown()
        End If
        connected = False
    End Sub

    Public Sub clearCredentials()
        log.Debug("clearCredentials - " & "clearing credentials")
        Try
            Destination.User = ""
            Destination.Password = Nothing
        Catch ex As System.Exception
            '            MsgBox("clearCredentials failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
            log.Error("clearCredentials - Exception=" & ex.ToString)
        End Try
    End Sub

    Public Function getDestination() As RfcCustomDestination
        getDestination = _destination
        log.Debug("getDestination - " & "destination=" & Destination.Name)
    End Function

    Public Function GetDestinationList() As Collection(Of String)
        GetDestinationList = aSapExcelDestinationConfiguration.getDestinationList()
    End Function

End Class


