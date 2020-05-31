' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports System.Reflection
Public Class SAPVersion
    Private _sapcon As SAPCon
    Public Sub New(ByRef pSapCon As SAPCon)
        _sapcon = pSapCon
    End Sub

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Public Function CheckVersionInSAP(pAddin As String, pVersion As String) As Integer
        Dim aSAPZ_BC_EXCEL_ADDIN_VERS_CHK As New SAPZ_BC_EXCEL_ADDIN_VERS_CHK(_sapcon)
        Dim aRet As Integer

        CheckVersionInSAP = True
        log.Debug("checkVersionInSAP - " & "aAddIn=" & CStr(pAddin))
        log.Debug("checkVersionInSAP - " & "calling aSAPZ_BC_EXCEL_ADDIN_VERS_CHK.checkVersion")
        aRet = aSAPZ_BC_EXCEL_ADDIN_VERS_CHK.checkVersion(pAddin, pVersion)
        If aRet <> 0 Then
            '            MsgBox("The Version " & cVersion & " of the Add-In " & aAddIn & " is not allowed in this SAP-System!",
            '            MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapGeneral")
            log.Debug("checkVersionInSAP - " & "The Version " & pVersion & " of the Add-In " & pAddin & " is not allowed in this SAP-System!")
            CheckVersionInSAP = False
        End If
    End Function

End Class

