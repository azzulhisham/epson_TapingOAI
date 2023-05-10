Imports System.Runtime.InteropServices
Imports System.Data
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.IO
Imports System.Drawing
Imports Microsoft.Win32
Imports System.Net
Imports Newtonsoft.Json

Module mdl_Database

    Public Function CreateDBConnection(ByVal dbFilePath As String) As OdbcConnection

        ' Connect string.
        Dim sConnStr As String = _
            "driver=Microsoft Access Driver (*.mdb); uid=admin; UserCommitSync=Yes; " & _
                    "Threads=3; SafeTransactions=0; PageTimeout=5; MaxScanRows=8; MaxBufferSize=2048; FIL=MS Access; DriverId=25; " & _
                    "DefaultDir=" & dbFilePath.Substring(0, dbFilePath.LastIndexOf("\")) & "; " & _
                    "DBQ=" & dbFilePath

        Try
            ' Open Connection.
            Dim oConn As New OdbcConnection(sConnStr)
            oConn.Open()
            ' Return Object.
            Return oConn
        Catch ex As Exception

        End Try

        Return Nothing

    End Function

    Public Function ReadMI_DataBase(ByVal DBSource As DB, ByRef Data As InspData, Optional ByRef SysID As String = "") As Long

        Dim oConn As OdbcConnection = CreateDBConnection(DBSource.Path & "\" & DBSource.FileName)

        Dim ch As Char = ChrW(39)
        Dim sGUID As String = ""
        Dim lRetVal As Long = 0


        If IsNothing(oConn) Then
            Return -1
        End If

        Dim sSQLcmd As String = String.Empty

        If SysID = "" Then
            Dim str_pLotNo As String = Data.P_LotNo
            sSQLcmd = "SELECT * FROM Record WHERE sLotNo ='"

            Do Until str_pLotNo = ""
                Application.DoEvents()

                If str_pLotNo.IndexOf(",") < 0 Then
                    sSQLcmd &= str_pLotNo.Trim & "'"
                    str_pLotNo = ""
                Else
                    sSQLcmd &= str_pLotNo.Substring(0, str_pLotNo.IndexOf(",")) & "'"
                    str_pLotNo = str_pLotNo.Substring(str_pLotNo.IndexOf(",") + 1).Trim

                    sSQLcmd &= " OR sLotNo='"
                End If
            Loop
        Else
            sSQLcmd = "SELECT * FROM Record WHERE sLotNo ='" & Data.LotNo & "'" & " ORDER BY InspType DESC"
        End If


        Dim OdbcCmd As New OdbcCommand(sSQLcmd, oConn)
        Dim oDBReader As OdbcDataReader = OdbcCmd.ExecuteReader()

        With oDBReader
            Dim iFieldCnt As Integer = .FieldCount
            Dim iRecNo As Integer = 0
            Dim pLot() As QtyCtrl


            If .HasRows Then
                Dim iLp As Integer = 0
                Dim sRetData(iFieldCnt - 1) As String

                If SysID = "" Then
                    ReDim Data.MarkData1(iRecNo)
                    ReDim Data.MarkData2(iRecNo)
                    ReDim Data.MFCSpec(iRecNo)
                    ReDim pLot(iRecNo)

                    Application.DoEvents()

                    Do While .Read()
                        Application.DoEvents()

                        ReDim Preserve Data.MarkData1(iRecNo)
                        ReDim Preserve Data.MarkData2(iRecNo)
                        ReDim Preserve Data.MFCSpec(iRecNo)
                        ReDim Preserve pLot(iRecNo)

                        With Data
                            .MFCSpec(iRecNo) = oDBReader.GetString(0)
                            .MarkData1(iRecNo) = oDBReader.GetString(7)
                            .MarkData2(iRecNo) = oDBReader.GetString(8).Replace(".", "")

                            If .MarkData2(iRecNo).Length > 6 Then
                                .MarkData2(iRecNo) = .MarkData2(iRecNo).Substring(1)
                            End If
                        End With

                        With Tp_OAI
                            Dim tmp As String = oDBReader.GetString(1)
                            Dim pData As System.Collections.Generic.IEnumerable(Of QtyCtrl) = .SysData(1).P_Lot.Where(Function(q) q.P_LotNo = tmp And q.WeekCode = "")

                            If pData.Count > 0 Then
                                pLot(iRecNo) = pData(0)
                                pLot(iRecNo).WeekCode = Data.MarkData2(iRecNo)
                            End If
                        End With

                        iRecNo += 1
                    Loop

                    lRetVal = iRecNo - 1
                    ReDim Tp_OAI.pLotQty(lRetVal)
                    Array.Copy(pLot, Tp_OAI.pLotQty, pLot.Length)
                Else
                    Tp_OAI.InspResult(1).GUID_No = oDBReader.GetString(0)
                    Tp_OAI.InspResult(1).LotNo = oDBReader.GetString(1)
                    Tp_OAI.InspResult(1).P_LotNo = oDBReader.GetString(2)
                    Tp_OAI.InspResult(1).InspType = oDBReader.GetString(3)
                    Tp_OAI.InspResult(1).InspOpt = oDBReader.GetString(4)
                    Tp_OAI.InspResult(1).InspMC = oDBReader.GetString(5)
                    Tp_OAI.InspResult(1).InspCnt = oDBReader.GetString(6)
                    Tp_OAI.InspResult(1).InspDate = oDBReader.GetString(7)

                    lRetVal = 1
                End If
            Else
                lRetVal = -1
            End If
        End With

        oConn.Close()
        oConn.Dispose()

        Return lRetVal

    End Function

    Public Function GetRecordsFromServer(ByVal Lot_No As String, ByRef RecData() As Rec) As Integer

        Dim RetVal As Integer = 0
        Dim sConnStr As String = _
            "SERVER=" & sqlServer & "; " & _
            "DataBase=" & sqlName & "; " & _
            "uid=" & sqluid & ";" & _
            "pwd=" & sqlpwd
        '"Integrated Security=SSPI"

        Dim dbConnection As New SqlConnection(sConnStr)
        Dim ch As Char = ChrW(39)
        Dim strSQL As String = _
            "SELECT * FROM Records WHERE Lot_No='" & Lot_No & "' " & _
            "ORDER BY Lot_No"


        Dim str_pLotNo() As String = Lot_No.Split(","c)
        strSQL = "SELECT * FROM Records WHERE Lot_No ='"


        For iLp As Integer = 0 To str_pLotNo.GetUpperBound(0)
            Application.DoEvents()

            strSQL &= str_pLotNo(iLp).Trim & "'"

            If iLp = str_pLotNo.GetUpperBound(0) Then
                strSQL &= " ORDER BY Lot_No"
            Else
                strSQL &= " OR Lot_No='"
            End If
        Next

        Try
            dbConnection.Open()

            Dim cmd As New SqlCommand(strSQL, dbConnection)
            cmd.ExecuteNonQuery()

            Dim sqlReader As SqlDataReader = cmd.ExecuteReader()

            With sqlReader
                Dim iFieldCnt As Integer = .FieldCount
                Dim iRecNo As Integer = 0

                ReDim RecData(iRecNo)

                If .HasRows Then
                    Dim sRetData(iFieldCnt - 1) As String

                    Do While .Read()
                        ReDim Preserve RecData(iRecNo)

                        With RecData(iRecNo)
                            .Lot_No = sqlReader.GetString(0)
                            .IMI_No = sqlReader.GetString(1)
                            .FreqVal = sqlReader.GetString(2)
                            .Opt = sqlReader.GetString(3)
                            .RecDate = sqlReader.GetDateTime(4).ToString
                            .Profile = sqlReader.GetString(5)
                            .CtrlNo = sqlReader.GetString(6)
                            .MacNo = sqlReader.GetString(7)
                            .MData1 = sqlReader.GetString(8)
                            .MData2 = sqlReader.GetString(9)
                            .MData3 = sqlReader.GetString(10)
                            .MData4 = sqlReader.GetString(11)
                            .MData5 = sqlReader.GetString(12)
                            .MData6 = sqlReader.GetString(13)
                        End With

                        iRecNo += 1
                    Loop

                    RetVal = iRecNo
                Else
                    RetVal = -1
                End If
            End With
        Catch sqlExc As SqlException
            RetVal = -1
        End Try

        dbConnection.Close()
        Return RetVal

    End Function

    Public Function GetExternalSourceData(ByVal vLotNo As String) As List(Of TapingLotFormData)

        Dim result As List(Of TapingLotFormData) = New List(Of TapingLotFormData)
        Dim EndPoint As String = String.Format("http://172.16.59.252/epson/Api/GetTapingLotFormDataPCS?vLotNo={0}", vLotNo)

        Dim request As HttpWebRequest = HttpWebRequest.Create(EndPoint)
        Dim response As HttpWebResponse

        request.Method = WebRequestMethods.Http.Get

        Try
            response = request.GetResponse()

            Using reader As New StreamReader(response.GetResponseStream())
                Dim json As String = reader.ReadToEnd()


                Dim tapingLotFormData As List(Of TapingLotFormData) =
                    JsonConvert.DeserializeObject(Of List(Of TapingLotFormData))(json)

                result = tapingLotFormData
            End Using
        Catch ex As Exception

            Dim er As String = ex.Message

        End Try

        Return result

    End Function

End Module
