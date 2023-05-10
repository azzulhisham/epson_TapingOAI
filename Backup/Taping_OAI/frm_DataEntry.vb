Public Class frm_DataEntry

    Dim MsgDisp(4) As String
    Dim TmpLotNo As String

    Dim LE_SeqNo As Integer = 0

    Dim fgLoad As Boolean = False
    Dim fgConf As Boolean = False


    Private Sub frm_DataEntry_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        fgLoad = False
        fgConf = False

    End Sub

    Private Sub frm_DataEntry_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If fgLoad = True Then Exit Sub
        fgLoad = True
        fgConf = False

        MsgDisp(0) = "Please Enter Lot No. ..."
        MsgDisp(1) = "Please input acceptance quantity..."
        MsgDisp(2) = "Scan P-Lot No., Press 'ENTER' to complete!"
        MsgDisp(3) = "Please Operator No. ..."
        MsgDisp(4) = "Processing... A Moment Please!"

        With Tp_OAI
            .RedoFlag = False
        End With

        With Me
            Tp_OAI.SysData(1).P_LotNo = ""

            .LE_SeqNo = 0
            .chk_TapeInsp.CheckState = IIf(Tp_OAI.InspTapeSeal.Mode <> 0, CheckState.Checked, CheckState.Unchecked)
            .txt_DataInput.Visible = True
            .lbl_AllLotNo.Text = ""
            .DispMsg()

            With .tmr_ReadIO
                .Interval = 50
                .Enabled = True
            End With
        End With

    End Sub


    Private Sub frm_DataEntry_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        With Tp_OAI
            If My.Computer.Network.Ping("172.16.59.2") Then
                If Not My.Computer.FileSystem.DirectoryExists(.RedoLotTmpRecLoc) Then
                    My.Computer.FileSystem.CreateDirectory(.RedoLotTmpRecLoc)
                End If
            End If
        End With

        With Me
            .txt_DataInput.Focus()
        End With

    End Sub

    Private Sub txt_DataInput_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt_DataInput.GotFocus

        'With Me.txt_DataInput
        '    .SelectionStart = 0
        '    .SelectionLength = .Text.Length
        'End With

    End Sub

    Private Sub DispMsg()

        Dim strLbl() As String = {"Lot No.", "Qty", "P-Lot No.", "Emp. No.", ""}


        With Me
            .lbl_EnterTitle.Text = MsgDisp(LE_SeqNo)
            .lbl_MsgBox.Text = strLbl(LE_SeqNo)

            If LE_SeqNo >= strLbl.GetUpperBound(0) Then
                .txt_DataInput.Visible = False
            Else
                With .txt_DataInput
                    'If LE_SeqNo = 2 Then
                    '    .Text = "3000"
                    '    .Text = ""
                    'End If

                    .Visible = True
                    .SelectAll()
                    .Focus()
                End With
            End If
        End With

    End Sub

    Private Sub tmr_ReadIO_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmr_ReadIO.Tick

        Static fg_Trg As Integer = 0

        If fg_Dbg = 0 Then
            With Tp_OAI
                If .IO.pb_EMG.BitState = cls_PCIBoard.BitsState.eBit_OFF Or .IO.pb_STOP.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                    .SysData(1).LotNo = ""
                    Me.LE_SeqNo = 0
                    Me.Close()
                End If
            End With
        End If


        'LED Blinking
        fg_Trg += 1

        If fg_Trg >= 4 Then
            fg_Trg = 0

            With Me
                .lbl_MsgBox.Visible = Not .lbl_MsgBox.Visible
            End With
        End If

    End Sub

    Private Function ReadRedoData(ByVal FileName As String) As Integer

        If My.Computer.Network.Ping("172.16.59.2") Then
            Dim RedoRec As String = My.Computer.FileSystem.ReadAllText(FileName)
            Dim RedoItem() As String = RedoRec.Split(vbLf)

            With Tp_OAI
                With .SysData(0)
                    .GUID = RedoItem(0).Trim
                    .LotNo = RedoItem(1).Trim & "_R"
                    .P_LotNo = RedoItem(2).Trim
                    .Prod = RedoItem(3).Trim
                    .Acceptance = RedoItem(4).Trim
                    .InspType = RedoItem(7).Trim
                    .ReDoLoc = RedoItem(8).Trim
                    .Insp = RedoItem(9).Trim


                    Dim MarkingData() As String = RedoItem(5).Split(","c)
                    Dim MarkingData_ As System.Collections.Generic.IEnumerable(Of String) = MarkingData.Where(Function(q) q.Length > 0 And q <> "")

                    ReDim .MarkData1(MarkingData_.Count - 1)

                    For iLp As Integer = 0 To MarkingData_.Count - 1
                        Application.DoEvents()
                        .MarkData1(iLp) = MarkingData_(iLp).Trim
                    Next

                    MarkingData = RedoItem(6).Split(","c)
                    MarkingData_ = MarkingData.Where(Function(q) q.Length > 0 And q <> "")

                    ReDim .MarkData2(MarkingData_.Count - 1)

                    For iLp As Integer = 0 To MarkingData_.Count - 1
                        Application.DoEvents()
                        .MarkData2(iLp) = MarkingData_(iLp).Trim
                    Next

                    .RedoInsp = True
                End With
            End With

            Return 0
        Else
            Return -1
        End If

    End Function


    Private Sub txt_DataInput_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_DataInput.KeyDown

        If e.KeyCode = Keys.Enter Then
            Dim strData As String = String.Empty

            strData = Me.txt_DataInput.Text
            Me.txt_DataInput.Text = ""


            With Tp_OAI
                Select Case LE_SeqNo
                    Case Is = 0
                        If strData = "" Then
                            MessageBox.Show("Invalid Lot No.! Please scan again...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Select
                        End If

                        Dim iChk As Integer = strData.IndexOf("/")

                        If iChk = 0 Then
                            If strData.IndexOf("/.") = 0 Then
                                .SysData(1).LotNo = strData.Substring(2).ToUpper
                                .SysData(0).RedoInsp = False

                                If Not .SysData(1).LotNo.IndexOf("V") = 0 Or Not .SysData(1).LotNo.Length = 10 Then
                                    .SysData(1).LotNo = ""
                                    MessageBox.Show("Invalid Lot No.! Lot No should begins with character 'V' and contains 10 characters...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Else
                                    With Me
                                        .txt_DataInput.Visible = False
                                        .lbl_EnterTitle.Text = "One Moment Please... Retrieving Data!"
                                    End With

                                    '--- Modify for re-do Lot ---
                                    Dim RedoLotData As String = .RedoLotTmpRecLoc & "\" & .SysData(1).LotNo & ".dat"

                                    If My.Computer.FileSystem.FileExists(RedoLotData) Then
                                        With Me
                                            If .fgConf = False Then
                                                .TmpLotNo = Tp_OAI.SysData(1).LotNo
                                                .fgConf = True
                                            Else
                                                If Tp_OAI.SysData(1).LotNo = .TmpLotNo Then
                                                    'Get Previous Record
                                                    ReadRedoData(RedoLotData)
                                                    Me.LE_SeqNo = 0
                                                    Me.Close()
                                                End If
                                            End If

                                            .txt_DataInput.Visible = True
                                            .lbl_MsgBox.Text = ""
                                        End With

                                        DispMsg()
                                        Exit Sub
                                    End If

                                    .InspResult(0).GUID_No = System.Guid.NewGuid.ToString.ToUpper

                                    If Not .SysData(1).LotNo.IndexOf("V_") < 0 Then
                                        LE_SeqNo += 2
                                    Else
                                        Dim iRecCnt As Integer = ReadMI_DataBase(.Database(1), .SysData(1), .InspResult(0).GUID_No)
                                        LE_SeqNo += 1
                                    End If

                                    With Me
                                        .txt_DataInput.Visible = True
                                        .lbl_MsgBox.Text = ""
                                    End With

                                    DispMsg()
                                End If
                            Else
                                MessageBox.Show("Invalid Lot No.! Please scan again...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If
                        ElseIf iChk > 0 Then
                            MessageBox.Show("Invalid Lot No.! Please scan again...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ElseIf iChk < 0 Then
                            .SysData(1).LotNo = strData.ToUpper
                            .SysData(0).RedoInsp = False

                            If Not .SysData(1).LotNo.IndexOf("V") = 0 Or Not .SysData(1).LotNo.Length = 10 Then
                                .SysData(1).LotNo = ""
                                MessageBox.Show("Invalid Lot No.! Lot No should begins with character 'V' and contains 10 characters...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Else
                                With Me
                                    .txt_DataInput.Visible = False
                                    .lbl_EnterTitle.Text = "One Moment Please... Retrieving Data!"
                                End With

                                '--- Modify for re-do Lot ---
                                Dim RedoLotData As String = .RedoLotTmpRecLoc & "\" & .SysData(1).LotNo & ".dat"

                                If My.Computer.FileSystem.FileExists(RedoLotData) Then
                                    With Me
                                        If .fgConf = False Then
                                            .TmpLotNo = Tp_OAI.SysData(1).LotNo
                                            .fgConf = True
                                        Else
                                            If Tp_OAI.SysData(1).LotNo = .TmpLotNo Then
                                                'Get Previous Record
                                                ReadRedoData(RedoLotData)
                                                Me.LE_SeqNo = 0
                                                Me.Close()
                                            End If
                                        End If

                                        .txt_DataInput.Visible = True
                                        .lbl_MsgBox.Text = ""
                                    End With

                                    DispMsg()
                                    Exit Sub
                                End If


                                .InspResult(0).GUID_No = System.Guid.NewGuid.ToString.ToUpper

                                If Not .SysData(1).LotNo.IndexOf("V_") < 0 Then
                                    LE_SeqNo += 2
                                Else
                                    Dim iRecCnt As Integer = ReadMI_DataBase(.Database(1), .SysData(1), .InspResult(0).GUID_No)
                                    LE_SeqNo += 1
                                End If

                                With Me
                                    .txt_DataInput.Visible = True
                                    .lbl_MsgBox.Text = ""
                                End With

                                DispMsg()
                            End If
                        End If
                    Case Is = 2
                        If strData = "" And .SysData(1).P_LotNo = "" Then
                            MessageBox.Show("Invalid P-Lot No.! Please scan again...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Select
                        End If

                        If strData = "" Then
                            Dim TotalQtyInput As Integer = 0

                            For iLp As Integer = 0 To .SysData(1).P_Lot.GetUpperBound(0)
                                Application.DoEvents()
                                TotalQtyInput += .SysData(1).P_Lot(iLp).QtyUsed
                            Next

                            'If Val(.SysData(1).Acceptance) <> TotalQtyInput Then
                            '    ReDim .SysData(1).P_Lot(0)
                            '    .SysData(1).P_LotNo = ""
                            '    Me.lbl_AllLotNo.Text = .SysData(1).P_LotNo
                            '    Exit Select
                            'Else
                            '    Me.lbl_AllLotNo.Text = ""
                            '    LE_SeqNo += 1
                            '    DispMsg()
                            '    Exit Select
                            'End If


                            Me.lbl_AllLotNo.Text = ""
                            LE_SeqNo += 1
                            DispMsg()
                            Exit Select
                        End If


                        Dim iChk As Integer = strData.IndexOf("/")
                        Dim p_LotNo As String = String.Empty

                        If iChk = 0 Then
                            If strData.IndexOf("/.") = 0 Then
                                '.SysData(1).P_LotNo = strData.Substring(2).ToUpper
                                p_LotNo = strData.Substring(2).ToUpper
                                Tp_OAI.P_Lot_No = p_LotNo

                                'frm_QtyEntry.ShowDialog(Me)
                                fg_Qty = 0

                                If Not p_LotNo.IndexOf("P") = 0 Or Not p_LotNo.IndexOf("-") = 3 Then
                                    MessageBox.Show("Invalid P-Lot No.! P-Lot No should begins with character 'P' and contains '-' characters...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Else
                                    'LE_SeqNo += 1

                                    If .SysData(1).P_LotNo = "" Then
                                        .SysData(1).P_LotNo = p_LotNo

                                        ReDim .SysData(1).P_Lot(0)

                                        With .SysData(1).P_Lot(0)
                                            .P_LotNo = p_LotNo
                                            .QtyUsed = fg_Qty
                                            .WeekCode = ""
                                        End With
                                    Else
                                        If .SysData(1).P_LotNo.IndexOf(p_LotNo) < 0 Then
                                            .SysData(1).P_LotNo = ", " & p_LotNo

                                            Dim P_LotCnt As Integer = .SysData(1).P_Lot.GetUpperBound(0) + 1
                                            ReDim Preserve .SysData(1).P_Lot(P_LotCnt)

                                            With .SysData(1).P_Lot(P_LotCnt)
                                                .P_LotNo = p_LotNo
                                                .QtyUsed = fg_Qty
                                                .WeekCode = ""
                                            End With
                                        End If
                                    End If

                                    Me.lbl_AllLotNo.Text = .SysData(1).P_LotNo
                                    DispMsg()
                                End If
                            Else
                                MessageBox.Show("Invalid P-Lot No.! Please scan again...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            End If
                        ElseIf iChk > 0 Then
                            MessageBox.Show("Invalid P-Lot No.! Please scan again...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ElseIf iChk < 0 Then
                            p_LotNo = strData.ToUpper
                            Tp_OAI.P_Lot_No = p_LotNo

                            'frm_QtyEntry.ShowDialog(Me)
                            fg_Qty = 0

                            If Not p_LotNo.IndexOf("P") = 0 Or Not p_LotNo.IndexOf("-") = 3 Then
                                MessageBox.Show("Invalid P-Lot No.! P-Lot No should begins with character 'P' and contains '-' characters...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Else
                                'LE_SeqNo += 1

                                If .SysData(1).P_LotNo = "" Then
                                    .SysData(1).P_LotNo = p_LotNo

                                    ReDim .SysData(1).P_Lot(0)

                                    With .SysData(1).P_Lot(0)
                                        .P_LotNo = p_LotNo
                                        .QtyUsed = fg_Qty
                                        .WeekCode = ""
                                    End With
                                Else
                                    If .SysData(1).P_LotNo.IndexOf(p_LotNo) < 0 Then
                                        .SysData(1).P_LotNo &= ", " & p_LotNo

                                        Dim P_LotCnt As Integer = .SysData(1).P_Lot.GetUpperBound(0) + 1
                                        ReDim Preserve .SysData(1).P_Lot(P_LotCnt)

                                        With .SysData(1).P_Lot(P_LotCnt)
                                            .P_LotNo = p_LotNo
                                            .QtyUsed = fg_Qty
                                            .WeekCode = ""
                                        End With
                                    End If
                                End If

                                Me.lbl_AllLotNo.Text = .SysData(1).P_LotNo
                                DispMsg()
                            End If
                        End If
                    Case Is = 1
                        If strData = "" Then
                            MessageBox.Show("Invalid Qty No.! Please scan again...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Select
                        End If


                        Dim iChk As Integer = strData.IndexOf("/")

                        If iChk = 0 Then
                            Dim ChkString As String = "//"

                            If strData.IndexOf(ChkString) <> 0 Then
                                MessageBox.Show("Invalid Qty No.! Please scan again...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Exit Select
                            Else
                                strData = strData.Substring(ChkString.Length)
                            End If
                        End If

                        If Not IsNumeric(strData) Or Val(strData) > 6000 Then
                            MessageBox.Show("Invalid Qty No.! Please scan again...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Select
                        End If

                        .SysData(1).Acceptance = strData.Trim
                        LE_SeqNo += 1

                        With Me
                            .txt_DataInput.Visible = True
                            .lbl_MsgBox.Text = ""
                        End With

                        DispMsg()
                    Case Is = 3
                        If strData = "" Then
                            MessageBox.Show("Invalid Emp. No.! Please scan again...", "Data Invalid...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Select
                        End If

                        .SysData(1).Insp = strData.ToUpper
                        LE_SeqNo += 1
                        DispMsg()

                        'Retrieve Inspection Data
                        If Not .SysData(1).LotNo.IndexOf("V_") < 0 Then
                            .SysData(1).GUID = .InspResult(0).GUID_No

                            .SysData(1).Prod = ProdID(frm_Main.cbo_Product.SelectedIndex)
                            .SysData(1).P_LotNo = frm_Main.txt_DmyLotNo.Text.ToUpper.Trim

                            Try
                                .SelectedSceneNo = Val(SceneNoDB(Array.IndexOf(ProdID, .SysData(1).Prod)))
                            Catch ex As Exception
                                .SelectedSceneNo = 0
                            End Try

                            frm_Main.txt_DataBlock1.Text = frm_Main.txt_DataBlock1.Text.Replace(", ", ",")
                            frm_Main.txt_DataBlock2.Text = frm_Main.txt_DataBlock2.Text.Replace(", ", ",")

                            Dim BlockData1() As String = frm_Main.txt_DataBlock1.Text.Split(","c)
                            Dim BlockData2() As String = frm_Main.txt_DataBlock2.Text.Split(","c)

                            Dim iDmy As Integer = 0
                            ReDim .SysData(1).MarkData1(BlockData1.GetUpperBound(0))
                            ReDim .SysData(1).MarkData2(BlockData2.GetUpperBound(0))

                            Array.Copy(BlockData1, .SysData(1).MarkData1, BlockData1.Length)
                            Array.Copy(BlockData2, .SysData(1).MarkData2, BlockData2.Length)

                            '--- Insp. Type --- 
                            '1: Include Tape Insp.
                            '2: Exclude Tape Insp.
                            .SysData(0) = .SysData(1)
                            .SysData(0).InspType = IIf(Me.chk_TapeInsp.CheckState = CheckState.Checked, "1", "0")

                            Me.Close()
                            Exit Sub
                        End If


                        'Get Database
                        .InspResult(0).GUID_No = System.Guid.NewGuid.ToString.ToUpper
                        Dim iRecCnt As Integer = ReadMI_DataBase(.Database(1), .SysData(1), .InspResult(0).GUID_No)
                        Dim iSQLrec As Integer = 0


                        If iRecCnt > 0 Then
                            .InspResult(0) = .InspResult(1)
                            .SysData(1).GUID = .InspResult(0).GUID_No

                            If .InspResult(0).InspType = 0 Then
                                MessageBox.Show("Inspection has been commpletely done for this Lot.", "Inspection Complete...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                .SysData(0).GUID = ""
                                Me.LE_SeqNo = 0
                                Me.Close()
                                Exit Sub
                            Else
                                Dim ChkSpec As Integer = 0
                                .SysData(1).InspType = (.InspResult(0).InspType - 1).ToString
                                '.SysData(1).P_LotNo = .InspResult(0).P_LotNo


                                With Tp_OAI
                                    iRecCnt = ReadMI_DataBase(.Database(0), .SysData(1))

                                    Dim SQLdbRecs() As Rec = Nothing
                                    iSQLrec = GetRecordsFromServer(.SysData(1).P_LotNo, SQLdbRecs)

                                    If iSQLrec > 0 Then
                                        If iRecCnt < 0 Then
                                            ReDim .SysData(1).MarkData1(SQLdbRecs.GetUpperBound(0))
                                            ReDim .SysData(1).MarkData2(SQLdbRecs.GetUpperBound(0))
                                            ReDim .SysData(1).MFCSpec(SQLdbRecs.GetUpperBound(0))
                                            ReDim .pLotQty(SQLdbRecs.GetUpperBound(0))

                                            For ilp As Integer = 0 To SQLdbRecs.GetUpperBound(0)
                                                .SysData(1).MarkData1(ilp) = SQLdbRecs(ilp).MData1
                                                .SysData(1).MarkData2(ilp) = SQLdbRecs(ilp).MData2
                                                .SysData(1).MFCSpec(ilp) = SQLdbRecs(ilp).IMI_No

                                                Dim tmp As String = SQLdbRecs(ilp).Lot_No
                                                Dim pData As System.Collections.Generic.IEnumerable(Of QtyCtrl) = .SysData(1).P_Lot.Where(Function(q) q.P_LotNo = tmp And q.WeekCode = "")

                                                If pData.Count > 0 Then
                                                    .pLotQty(ilp) = pData(0)
                                                    .pLotQty(ilp).WeekCode = SQLdbRecs(ilp).MData2
                                                End If
                                            Next
                                        Else
                                            Dim TotalRecs As Integer = .SysData(1).MFCSpec.GetUpperBound(0) + SQLdbRecs.Length

                                            ReDim Preserve .SysData(1).MarkData1(TotalRecs)
                                            ReDim Preserve .SysData(1).MarkData2(TotalRecs)
                                            ReDim Preserve .SysData(1).MFCSpec(TotalRecs)
                                            ReDim Preserve .pLotQty(TotalRecs)

                                            For ilp As Integer = .SysData(1).MFCSpec.Length To TotalRecs
                                                .SysData(1).MarkData1(ilp) = SQLdbRecs(ilp).MData1
                                                .SysData(1).MarkData2(ilp) = SQLdbRecs(ilp).MData2
                                                .SysData(1).MFCSpec(ilp) = SQLdbRecs(ilp).IMI_No

                                                Dim tmp As String = SQLdbRecs(ilp).Lot_No
                                                Dim pData As System.Collections.Generic.IEnumerable(Of QtyCtrl) = .SysData(1).P_Lot.Where(Function(q) q.P_LotNo = tmp And q.WeekCode = "")

                                                If pData.Count > 0 Then
                                                    .pLotQty(ilp) = pData(0)
                                                    .pLotQty(ilp).WeekCode = SQLdbRecs(ilp).MData2
                                                End If
                                            Next
                                        End If

                                        'Array.Copy(.pLotQty, .SysData(1).P_Lot, .pLotQty.Length)
                                    End If


                                    If iRecCnt < 0 And iSQLrec < 0 Then
                                        'No Rec
                                        If MessageBox.Show("No Marking Data Found!" & vbCrLf & "Anyway, do you wish to re-enter Lot Data?", "Invalid Data...", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                            .SysData(0).GUID = ""
                                            Me.LE_SeqNo = 0
                                            Me.Close()
                                        Else
                                            LE_SeqNo = 0
                                            DispMsg()
                                            Me.txt_DataInput.Focus()
                                        End If
                                    Else
                                        If iRecCnt > 0 Then
                                            For iLp As Integer = 1 To .SysData(1).MFCSpec.GetUpperBound(0)
                                                Application.DoEvents()

                                                If Not .SysData(1).MFCSpec(iLp) = .SysData(1).MFCSpec(iLp - 1) Then
                                                    ChkSpec = -1
                                                    Exit For
                                                End If
                                            Next
                                        Else
                                            If iRecCnt < 0 Then ChkSpec = -1
                                        End If
                                    End If
                                End With

                                If ChkSpec = 0 Then
                                    Dim MFCSpecFile As String = .SpecFileLocation & "\" & .SysData(1).MFCSpec(0) & ".dat"

                                    If My.Computer.FileSystem.FileExists(MFCSpecFile) Then
                                        Dim MFCSpecItem As String = My.Computer.FileSystem.ReadAllText(MFCSpecFile)
                                        Dim MFCSpecCode As String = MFCSpecItem.Substring(MFCSpecItem.IndexOf("L006"))

                                        MFCSpecCode = MFCSpecCode.Substring(MFCSpecCode.IndexOf(",") + 1).ToLower
                                        MFCSpecCode = MFCSpecCode.Substring(0, MFCSpecCode.LastIndexOf("format") - 1).ToUpper.Trim

                                        .SysData(1).Prod = MFCSpecCode
                                        .SysData(0) = .SysData(1)

                                        Try
                                            .SelectedSceneNo = Val(SceneNoDB(Array.IndexOf(ProdID, .SysData(1).Prod)))
                                        Catch ex As Exception
                                            .SelectedSceneNo = 0
                                        End Try
                                    Else
                                        MessageBox.Show("Manufacturing Instruction not found... Please Check Network Connection!", "IMI Not Found...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                        .SysData(0).GUID = ""
                                        Me.LE_SeqNo = 0
                                    End If
                                Else
                                    MessageBox.Show("Invalid Manufacturing Instruction... Check P-Lot IMI!", "Invalid IMI Found...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                    .SysData(0).GUID = ""
                                    Me.LE_SeqNo = 0
                                End If
                            End If

                            Me.Close()
                        Else
                            Dim ChkSpec As Integer = 0

                            With Tp_OAI
                                iRecCnt = ReadMI_DataBase(.Database(0), .SysData(1))
                                If Not iRecCnt < 0 Then iRecCnt += 1

                                Dim SQLdbRecs() As Rec = Nothing
                                iSQLrec = GetRecordsFromServer(.SysData(1).P_LotNo, SQLdbRecs)

                                If iSQLrec > 0 Then
                                    If iRecCnt < 0 Then
                                        ReDim .SysData(1).MarkData1(SQLdbRecs.GetUpperBound(0))
                                        ReDim .SysData(1).MarkData2(SQLdbRecs.GetUpperBound(0))
                                        ReDim .SysData(1).MFCSpec(SQLdbRecs.GetUpperBound(0))
                                        ReDim .pLotQty(SQLdbRecs.GetUpperBound(0))

                                        For ilp As Integer = 0 To SQLdbRecs.GetUpperBound(0)
                                            .SysData(1).MarkData1(ilp) = SQLdbRecs(ilp).MData1
                                            .SysData(1).MarkData2(ilp) = SQLdbRecs(ilp).MData2
                                            .SysData(1).MFCSpec(ilp) = SQLdbRecs(ilp).IMI_No

                                            Dim tmp As String = SQLdbRecs(ilp).Lot_No
                                            Dim pData As System.Collections.Generic.IEnumerable(Of QtyCtrl) = .SysData(1).P_Lot.Where(Function(q) q.P_LotNo = tmp And q.WeekCode = "")

                                            If pData.Count > 0 Then
                                                .pLotQty(ilp) = pData(0)
                                                .pLotQty(ilp).WeekCode = SQLdbRecs(ilp).MData2
                                            End If
                                        Next
                                    Else
                                        Dim OrgTecsCnt As Integer = .SysData(1).MarkData2.Length
                                        Dim TotalRecs As Integer = .SysData(1).MarkData2.GetUpperBound(0) + SQLdbRecs.Length

                                        ReDim Preserve .SysData(1).MarkData1(TotalRecs)
                                        ReDim Preserve .SysData(1).MarkData2(TotalRecs)
                                        ReDim Preserve .SysData(1).MFCSpec(TotalRecs)
                                        ReDim Preserve .pLotQty(TotalRecs)

                                        For ilp As Integer = OrgTecsCnt To TotalRecs
                                            .SysData(1).MarkData1(ilp) = SQLdbRecs(ilp - OrgTecsCnt).MData1
                                            .SysData(1).MarkData2(ilp) = SQLdbRecs(ilp - OrgTecsCnt).MData2
                                            .SysData(1).MFCSpec(ilp) = SQLdbRecs(ilp - OrgTecsCnt).IMI_No

                                            Dim tmp As String = SQLdbRecs(ilp - OrgTecsCnt).Lot_No
                                            Dim pData As System.Collections.Generic.IEnumerable(Of QtyCtrl) = .SysData(1).P_Lot.Where(Function(q) q.P_LotNo = tmp And q.WeekCode = "")

                                            If pData.Count > 0 Then
                                                .pLotQty(ilp) = pData(0)
                                                .pLotQty(ilp).WeekCode = SQLdbRecs(ilp - OrgTecsCnt).MData2
                                            End If
                                        Next
                                    End If

                                    'Array.Copy(.pLotQty, .SysData(1).P_Lot, .pLotQty.Length)
                                End If


                                If iRecCnt > 0 Or iSQLrec > 0 Then
                                    For iLp As Integer = 1 To .SysData(1).MFCSpec.GetUpperBound(0)
                                        Application.DoEvents()

                                        If Not .SysData(1).MFCSpec(iLp) = .SysData(1).MFCSpec(iLp - 1) Then
                                            ChkSpec = -1
                                            Exit For
                                        End If
                                    Next
                                Else
                                    If iRecCnt < 0 Then ChkSpec = -1
                                End If
                            End With


                            Dim SetRec As Integer = 0
                            ReDim .SysData(1).P_Lot(SetRec)

                            For iLp As Integer = 0 To .pLotQty.GetUpperBound(0)
                                ReDim Preserve .SysData(1).P_Lot(SetRec)

                                Dim tmp As String = .pLotQty(iLp).P_LotNo
                                Dim pData As System.Collections.Generic.IEnumerable(Of QtyCtrl) = .SysData(1).P_Lot.Where(Function(q) q.P_LotNo = tmp)

                                If pData.Count <= 0 Then
                                    .SysData(1).P_Lot(SetRec) = .pLotQty(iLp)
                                    SetRec += 1
                                End If
                            Next


                            If ChkSpec = 0 Then
                                Dim MFCSpecFile As String = .SpecFileLocation & "\" & .SysData(1).MFCSpec(0) & ".dat"

                                If My.Computer.FileSystem.FileExists(MFCSpecFile) Then
                                    Dim MFCSpecItem As String = My.Computer.FileSystem.ReadAllText(MFCSpecFile)
                                    Dim MFCSpecCode As String = MFCSpecItem.Substring(MFCSpecItem.IndexOf("L006"))

                                    MFCSpecCode = MFCSpecCode.Substring(MFCSpecCode.IndexOf(",") + 1).ToLower
                                    MFCSpecCode = MFCSpecCode.Substring(0, MFCSpecCode.LastIndexOf("format") - 1).ToUpper.Trim

                                    .SysData(1).GUID = .InspResult(0).GUID_No
                                    .SysData(1).Prod = MFCSpecCode
                                    .SysData(1).InspType = .InspResult(0).InspType

                                    Try
                                        .SelectedSceneNo = Val(SceneNoDB(Array.IndexOf(ProdID, .SysData(1).Prod)))
                                    Catch ex As Exception
                                        .SelectedSceneNo = 0
                                    End Try


                                    '--- Insp. Type --- 
                                    '1: Include Tape Insp.
                                    '2: Exclude Tape Insp.
                                    .SysData(0) = .SysData(1)
                                    .SysData(0).InspType = IIf(Me.chk_TapeInsp.CheckState = CheckState.Checked, "1", "0")
                                Else
                                    MessageBox.Show("Manufacturing Instruction not found... Please Check Network Connection!", "IMI Not Found...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                    .SysData(0).GUID = ""
                                    Me.LE_SeqNo = 0
                                End If
                            Else
                                MessageBox.Show("Invalid Manufacturing Instruction... Check P-Lot IMI!", "Invalid IMI Found...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                .SysData(0).GUID = ""
                                Me.LE_SeqNo = 0
                            End If

                            Me.Close()
                        End If
                End Select
            End With
        ElseIf e.KeyCode = Keys.Escape Then
            Tp_OAI.SysData(0).LotNo = ""
            Me.LE_SeqNo = 0
            Me.Close()
        End If

    End Sub

    Private Sub chk_TapeInsp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chk_TapeInsp.Click

        Me.txt_DataInput.Focus()

    End Sub

End Class