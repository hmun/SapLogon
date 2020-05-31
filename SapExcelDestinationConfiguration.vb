' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector
Imports System.Configuration
Imports System.Collections.Specialized
Imports System.Environment
Imports System.Uri
Imports System.IO
Imports SAPCommon
Imports System.Collections.ObjectModel

Public Class SapExcelDestinationConfiguration
        Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
        Private Shared inMemoryDestinationConfiguration As New SapInMemoryDestinationConfiguration()

        Public Shared Sub SetUp()
            '' register the in-memory destination configuration -- called before executing any of the examples
            log.Debug("SetUp - " & "RegisterDestinationConfiguration")
            Try
                RfcDestinationManager.RegisterDestinationConfiguration(inMemoryDestinationConfiguration)
            Catch Exc As System.Exception
                '            MsgBox(Exc.ToString, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapExcelDestinationConfiguration;SetUp")
                log.Error("SetUp - Exception=" & Exc.ToString)
                Exit Sub
            End Try
        End Sub

        Public Shared Sub TearDown(Optional destinationName As String = Nothing)
            '' unregister the in-memory destination configuration -- called after we are done working with the examples 
            RfcDestinationManager.UnregisterDestinationConfiguration(inMemoryDestinationConfiguration)
            If destinationName IsNot Nothing Then
                inMemoryDestinationConfiguration.RemoveDestination(destinationName)
            End If
        End Sub

        Public Shared Sub ConfigAddOrChangeDestination(pAssembly As String)
            Dim conParam() As String = {"Name", "AppServerHost", "SystemNumber", "SystemID", "Client", "Language", "SncMode", "SncPartnerName"}
            Dim conParameter As New ConParameter
            Dim parameters As New RfcConfigParameters()

        Dim appData As String = GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim configFile = Uri.UnescapeDataString(appData & "\SapExcel\" & pAssembly & "\sap_connections.config")
        If Not System.IO.File.Exists(configFile) Then
            appData = GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            configFile = Uri.UnescapeDataString(appData & "\SapExcel\" & pAssembly & "\sap_connections.config")
            If Not System.IO.File.Exists(configFile) Then
                appData = New Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).AbsolutePath
                appData = Path.GetDirectoryName(appData)
                configFile = Uri.UnescapeDataString(appData & "\sap_connections.config")
                If Not System.IO.File.Exists(configFile) Then
                        configFile = ""
                    End If
                End If
            End If
            Dim config As Configuration
            Dim configMap As New ExeConfigurationFileMap
        If Not configFile = "" Then
            Try
                configMap.ExeConfigFilename = configFile
                config = TryCast(ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None), Configuration)
                Dim sapConnectionsSection As SAPCommon.SapConnectionsSection = CType(config.Sections("SAPConnections"), SAPCommon.SapConnectionsSection)
                If sapConnectionsSection Is Nothing Then
                    log.Error("ConfigAddOrChangeDestination -" & "failed TypeOf read SAPConnections in " & configFile)
                Else
                    For i As Integer = 0 To sapConnectionsSection.SapConnections.Count - 1
                        log.Debug("ConfigAddOrChangeDestination - config file contains name=" & sapConnectionsSection.SapConnections(i).Name)
                        conParameter.addCon(i, sapConnectionsSection.SapConnections(i))
                    Next i
                End If
            Catch Exc As System.Exception
                log.Error("ConfigAddOrChangeDestination - parsing sap_connections.config, Exception=" & Exc.ToString)
            End Try
        End If
        Dim conRec As ConParamterRec
            For Each conRec In conParameter.aConCol.Values
                parameters = New RfcConfigParameters()
                parameters(RfcConfigParameters.Name) = conRec.aName.Value
                parameters(RfcConfigParameters.PeakConnectionsLimit) = "5"
                parameters(RfcConfigParameters.ConnectionIdleTimeout) = "600" '' 600 seconds, i.e. 10 minutes
                If conRec.aAppServerHost.Value IsNot Nothing Then
                    parameters(RfcConfigParameters.AppServerHost) = conRec.aAppServerHost.Value
                    parameters(RfcConfigParameters.SystemNumber) = CInt(conRec.aSystemNumber.Value)
                ElseIf conRec.aMessageServerHost.Value IsNot Nothing Then
                    parameters(RfcConfigParameters.MessageServerHost) = conRec.aMessageServerHost.Value
                    parameters(RfcConfigParameters.LogonGroup) = conRec.aLogonGroup.Value
                End If
                parameters(RfcConfigParameters.SystemID) = conRec.aSystemID.Value
                If conRec.aTrace.Value IsNot Nothing Then
                    parameters(RfcConfigParameters.Trace) = conRec.aTrace.Value
                End If
                If conRec.aClient.Value IsNot Nothing Then
                    parameters(RfcConfigParameters.Client) = conRec.aClient.Value
                End If
                If conRec.aLanguage.Value IsNot Nothing Then
                    parameters(RfcConfigParameters.Language) = conRec.aLanguage.Value
                End If
                If conRec.aSncMode.Value IsNot Nothing Then
                    parameters(RfcConfigParameters.SncMode) = conRec.aSncMode.Value
                    parameters(RfcConfigParameters.SncPartnerName) = conRec.aSncPartnerName.Value
                    If conRec.aSncMyName.Value IsNot Nothing Then
                        parameters(RfcConfigParameters.SncMyName) = conRec.aSncMyName.Value
                    End If
                End If
                log.Debug("ConfigAddOrChangeDestination - inMemoryDestinationConfiguration.AddOrEditDestination Name=" & conRec.aName.Value)
                Try
                    inMemoryDestinationConfiguration.AddOrEditDestination(parameters)
                Catch Exc As System.Exception
                    log.Error("ConfigAddOrChangeDestination - Exception=" & Exc.ToString)
                End Try
            Next
        End Sub

        Public Shared Sub ExcelAddOrChangeDestination(pConParameter As ConParameter)
        Dim parameters As New RfcConfigParameters()
        Dim conRec As ConParamterRec
        For Each conRec In pConParameter.aConCol.Values
            parameters = New RfcConfigParameters()
            parameters(RfcConfigParameters.Name) = conRec.aName.Value
            parameters(RfcConfigParameters.PeakConnectionsLimit) = "5"
            parameters(RfcConfigParameters.ConnectionIdleTimeout) = "600" '' 600 seconds, i.e. 10 minutes
            If conRec.aAppServerHost.Value IsNot Nothing Then
                parameters(RfcConfigParameters.AppServerHost) = conRec.aAppServerHost.Value
                parameters(RfcConfigParameters.SystemNumber) = CInt(conRec.aSystemNumber.Value)
            ElseIf conRec.aMessageServerHost.Value IsNot Nothing Then
                parameters(RfcConfigParameters.MessageServerHost) = conRec.aMessageServerHost.Value
                parameters(RfcConfigParameters.LogonGroup) = conRec.aLogonGroup.Value
            End If
            parameters(RfcConfigParameters.SystemID) = conRec.aSystemID.Value
            If conRec.aTrace.Value IsNot Nothing Then
                parameters(RfcConfigParameters.Trace) = conRec.aTrace.Value
            End If
            If conRec.aClient.Value IsNot Nothing Then
                parameters(RfcConfigParameters.Client) = conRec.aClient.Value
            End If
            If conRec.aLanguage.Value IsNot Nothing Then
                parameters(RfcConfigParameters.Language) = conRec.aLanguage.Value
            End If
            If conRec.aSncMode.Value IsNot Nothing Then
                parameters(RfcConfigParameters.SncMode) = conRec.aSncMode.Value
                parameters(RfcConfigParameters.SncPartnerName) = conRec.aSncPartnerName.Value
                If conRec.aSncMyName.Value IsNot Nothing Then
                    parameters(RfcConfigParameters.SncMyName) = conRec.aSncMyName.Value
                End If
            End If
            Try
                log.Debug("ExcelAddOrChangeDestination - inMemoryDestinationConfiguration.AddOrEditDestination Name=" & conRec.aName.Value)
                inMemoryDestinationConfiguration.AddOrEditDestination(parameters)
            Catch Exc As System.Exception
                log.Error("ExcelAddOrChangeDestination - Exception=" & Exc.ToString)
            End Try
        Next
    End Sub

        Public Function getDestinationList() As Collection(Of String)
            Dim list As New Collection(Of String)
            Dim availableDestinations As Dictionary(Of String, RfcConfigParameters)
            log.Debug("getDestinationList - getting availableDestinations")
            availableDestinations = inMemoryDestinationConfiguration.getAvailableDestinations()
            Dim key As String
            For Each key In availableDestinations.Keys
                list.Add(key)
            Next
            getDestinationList = list
            log.Debug("getDestinationList - getDestinationList.Count=" & CStr(getDestinationList.Count))
        End Function
    End Class
