' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAPCommon

Public Class ConParamterRec
    Public aID As TField
    Public aName As TField
    Public aAppServerHost As TField
    Public aSystemNumber As TField
    Public aSystemID As TField
    Public aMessageServerHost As TField
    Public aLogonGroup As TField
    Public aTrace As TField
    Public aClient As TField
    Public aLanguage As TField
    Public aSncMode As TField
    Public aSncMyName As TField
    Public aSncPartnerName As TField
    Public aSAPRouter As TField

    Public Sub New()
        aID = New TField
        aName = New TField
        aAppServerHost = New TField
        aSystemNumber = New TField
        aSystemID = New TField
        aMessageServerHost = New TField
        aLogonGroup = New TField
        aTrace = New TField
        aClient = New TField
        aLanguage = New TField
        aSncMode = New TField
        aSncMyName = New TField
        aSncPartnerName = New TField
        aSAPRouter = New TField
    End Sub

    Public Sub setValues(pId As Integer, pSapConnectionConfigElement As SAPCommon.SapConnectionConfigElement)
        aID = New TField("ID", CStr(pId))
        aName = New TField("Name", CStr(pSapConnectionConfigElement.Name))
        aSystemNumber = New TField("SystemNumber", CStr(pSapConnectionConfigElement.SystemNumber))
        aSystemID = New TField("SystemID", CStr(pSapConnectionConfigElement.SystemID))
        aClient = New TField("Client", CStr(pSapConnectionConfigElement.Client))
        aLanguage = New TField("Language", CStr(pSapConnectionConfigElement.Language))
        Try
            aAppServerHost = New TField("AppServerHost", CStr(pSapConnectionConfigElement.AppServerHost))
        Catch Exc As System.Exception
        End Try
        Try
            aMessageServerHost = New TField("MessageServerHost", CStr(pSapConnectionConfigElement.MessageServerHost))
        Catch Exc As System.Exception
        End Try
        Try
            aLogonGroup = New TField("LogonGroup", CStr(pSapConnectionConfigElement.LogonGroup))
        Catch Exc As System.Exception
        End Try
        Try
            aTrace = New TField("Trace", CStr(pSapConnectionConfigElement.Trace))
        Catch Exc As System.Exception
        End Try
        Try
            aSncMode = New TField("SncMode", CStr(pSapConnectionConfigElement.SncMode))
        Catch Exc As System.Exception
        End Try
        Try
            aSncMyName = New TField("SncMyName", CStr(pSapConnectionConfigElement.SncMyName))
        Catch Exc As System.Exception
        End Try
        Try
            aSncPartnerName = New TField("SncPartnerName", CStr(pSapConnectionConfigElement.SncPartnerName))
        Catch Exc As System.Exception
        End Try
        Try
            aSAPRouter = New TField("SAPRouter", CStr(pSapConnectionConfigElement.SAPRouter))
        Catch Exc As System.Exception
        End Try
    End Sub

    Public Sub setValues(pID As String, pName As String, pAppServerHost As String, pSystemNumber As String, pSystemID As String, pMessageServerHost As String, pLogonGroup As String, pTrace As String, pClient As String, pLanguage As String, pSncMode As String, pSncMyName As String, pSncPartnerName As String, Optional pSAPRouter As String = "")
        aID = New TField("ID", CStr(pID))
        aName = New TField("Name", CStr(pName))
        aAppServerHost = New TField("AppServerHost", CStr(pAppServerHost))
        aSystemNumber = New TField("SystemNumber", CStr(pSystemNumber))
        aSystemID = New TField("SystemID", CStr(pSystemID))
        aMessageServerHost = New TField("MessageServerHost", CStr(pMessageServerHost))
        aLogonGroup = New TField("LogonGroup", CStr(pLogonGroup))
        aTrace = New TField("Trace", CStr(pTrace))
        aClient = New TField("Client", CStr(pClient))
        aLanguage = New TField("Language", CStr(pLanguage))
        aSncMode = New TField("SncMode", CStr(pSncMode))
        aSncMyName = New TField("SncMyName", CStr(pSncMyName))
        aSncPartnerName = New TField("SncPartnerName", CStr(pSncPartnerName))
        aSAPRouter = New TField("SAPRouter", CStr(pSAPRouter))
    End Sub

    Public Sub setValue(pField As String, pValue As String)
        If pField = "ID" Then
            aID = New TField(pField, CStr(pValue))
        ElseIf pField = "Name" Then
            aName = New TField(pField, CStr(pValue))
        ElseIf pField = "AppServerHost" Then
            aAppServerHost = New TField(pField, CStr(pValue))
        ElseIf pField = "SystemNumber" Then
            aSystemNumber = New TField(pField, CStr(pValue))
        ElseIf pField = "SystemID" Then
            aSystemID = New TField(pField, CStr(pValue))
        ElseIf pField = "MessageServerHost" Then
            aMessageServerHost = New TField(pField, CStr(pValue))
        ElseIf pField = "LogonGroup" Then
            aLogonGroup = New TField(pField, CStr(pValue))
        ElseIf pField = "Trace" Then
            aTrace = New TField(pField, CStr(pValue))
        ElseIf pField = "Client" Then
            aClient = New TField(pField, CStr(pValue))
        ElseIf pField = "Language" Then
            aLanguage = New TField(pField, CStr(pValue))
        ElseIf pField = "SncMode" Then
            aSncMode = New TField(pField, CStr(pValue))
        ElseIf pField = "SncMyName" Then
            aSncMyName = New TField(pField, CStr(pValue))
        ElseIf pField = "SncPartnerName" Then
            aSncPartnerName = New TField(pField, CStr(pValue))
        ElseIf pField = "SAPRouter" Then
            aSAPRouter = New TField(pField, CStr(pValue))
        End If
    End Sub

    Public Function getKey() As String
            Dim aKey As String
            aKey = aID.Value
            getKey = aKey
        End Function

    Public Function getRKey() As String
        Dim aKey As String
        aKey = aID.Value
        getRKey = aKey
    End Function

End Class

Public Class ConParameter
    Public aConCol As Dictionary(Of String, ConParamterRec)
    Private sTField As TField

    Public Sub New()
        sTField = New TField
        aConCol = New Dictionary(Of String, ConParamterRec)
    End Sub

    Public Sub addCon(pID As Integer, pSapConnectionConfigElement As SAPCommon.SapConnectionConfigElement)
        Dim aConRec As New ConParamterRec
        Dim aKey As String
        aKey = CStr(pID)
        If aConCol.TryGetValue(aKey, aConRec) Then
            aConRec.setValues(pID, pSapConnectionConfigElement)
        Else
            aConRec = New ConParamterRec
            aConRec.setValues(pID, pSapConnectionConfigElement)
            aConCol.Add(aKey, aConRec)
        End If
    End Sub

    Public Sub addCon(pID As String, pName As String, pAppServerHost As String, pSystemNumber As String, pSystemID As String, pMessageServerHost As String, pLogonGroup As String, pTrace As String, pClient As String, pLanguage As String, pSncMode As String, pSncMyName As String, pSncPartnerName As String, Optional pSAPRouter As String = "")
        Dim aConRec As New ConParamterRec
        Dim aKey As String
        aKey = pID
        If aConCol.TryGetValue(aKey, aConRec) Then
            aConRec.setValues(pID, pName, pAppServerHost, pSystemNumber, pSystemID, pMessageServerHost, pLogonGroup, pTrace, pClient, pLanguage, pSncMode, pSncMyName, pSncPartnerName, pSAPRouter)
        Else
            aConRec = New ConParamterRec
            aConRec.setValues(pID, pName, pAppServerHost, pSystemNumber, pSystemID, pMessageServerHost, pLogonGroup, pTrace, pClient, pLanguage, pSncMode, pSncMyName, pSncPartnerName, pSAPRouter)
            aConCol.Add(aKey, aConRec)
        End If
    End Sub

    Public Sub addConValue(pID As String, pField As String, pValue As String)
        Dim aConRec As New ConParamterRec
        Dim aKey As String
        aKey = pID
        If aConCol.TryGetValue(aKey, aConRec) Then
            aConRec.setValue(pField, pValue)
        Else
            aConRec = New ConParamterRec
            aConRec.setValue("ID", pID)
            aConRec.setValue(pField, pValue)
            aConCol.Add(aKey, aConRec)
        End If
    End Sub

End Class

