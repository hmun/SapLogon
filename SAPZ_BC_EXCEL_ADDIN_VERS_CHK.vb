' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPZ_BC_EXCEL_ADDIN_VERS_CHK

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private _destination As RfcCustomDestination
    Private _sapcon As SAPCon

    Sub New(ByRef aSapCon As SAPCon)
        _sapcon = aSapCon
        _destination = aSapCon.getDestination()
        log.Debug("New - " & "creating Function Z_BC_EXCEL_ADDIN_VERS_CHK")
        Try
            oRfcFunction = _destination.Repository.CreateFunction("Z_BC_EXCEL_ADDIN_VERS_CHK")
            log.Debug("New - " & "oRfcFunction.Metadata.Name=" & oRfcFunction.Metadata.Name)
        Catch ex As Exception
            oRfcFunction = Nothing
            log.Warn("New - Exception=" & ex.ToString)
        End Try
    End Sub

    Public Function checkVersion(pAddIn As String, pVersion As String) As Integer
        _sapcon.checkCon()
        If oRfcFunction Is Nothing Then
            ' for systems that do not contain Z_BC_EXCEL_ADDIN_VERS_CHK we can not check the version
            checkVersion = 0
            log.Debug("checkVersion - " & "oRfcFunction is Nothing, skiping check. checkVersion=" & checkVersion)
        Else
            Try
                log.Debug("checkVersion - " & "Setting Function parameters")
                Dim oRETURN As IRfcTable = oRfcFunction.GetTable("T_RETURN")
                Dim oE_ALLOWED_VERSION As IRfcStructure = oRfcFunction.GetStructure("E_ALLOWED_VERSION")
                oRETURN.Clear()

                oRfcFunction.SetValue("I_ADDIN", pAddIn)
                oRfcFunction.SetValue("I_VERSION", pVersion)

                log.Debug("checkVersion - " & "invoking " & oRfcFunction.Metadata.Name)
                oRfcFunction.Invoke(_destination)
                log.Debug("checkVersion - " & "oRETURN.Count=" & CStr(oRETURN.Count))
                If oRETURN.Count > 0 Then
                    checkVersion = If(CType(oRETURN(0).GetValue("TYPE"), String) = "S", 0, 4)
                Else
                    checkVersion = 8
                End If
                log.Debug("checkVersion - " & "checkVersion=" & CStr(checkVersion))
            Catch abap_ex As RfcAbapBaseException
                Select Case abap_ex.Message
                    Case "WRONG_VERSION_FORMAT"
                        checkVersion = 1
                    Case "UNSUPPORTED_VERSION"
                        checkVersion = 2
                    Case "NO_VERSION_MAINTAINED"
                        checkVersion = 3
                    Case Else
                        checkVersion = 8
                End Select
                log.Debug("checkVersion - " & "abap_ex= " & abap_ex.Message & ", checkVersion=" & CStr(checkVersion))
                Exit Function
            Catch ex As Exception
                '                MsgBox("Exception in checkVersion! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPZ_BC_EXCEL_ADDIN_VERS_CHK")
                checkVersion = 8
                log.Error("checkVersion - " & "ex= " & ex.ToString & ", checkVersion=" & CStr(checkVersion))
            End Try
        End If
        Exit Function
    End Function
End Class

