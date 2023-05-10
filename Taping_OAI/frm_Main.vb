
'---------------------------------------------------
'   PX - Taping OAI Development
'===================================================
'   Designed By : Zulhisham Tan
'   Module      : mdl_TapingOAI.vb
'   Date        : 28-Mar-2011
'   Version     : 2011.08.028.001
'---------------------------------------------------
'   Copyright (C) 2011-2014 az_Zulhisham
'---------------------------------------------------

Imports System.IO
Imports System.IO.Ports

Imports System.Math
Imports Microsoft.Win32
Imports System.Threading


Public Class frm_Main

    Public WithEvents TP_CAM1 As SerialPort = New SerialPort

    Delegate Sub UpdateControl(ByVal FormControl As Label, ByVal ThreadName As String, ByVal Text As String)
    Delegate Sub DispMsg(ByVal FormControl As Label, ByVal ThreadName As String, ByVal Text As String)
    Delegate Sub DispMsg_(ByVal FormControl As Button, ByVal ThreadName As String, ByVal Text As String)
    Delegate Sub DispCtrl(ByVal FormControl As PictureBox, ByVal ThreadName As String, ByVal Text As String)
    Delegate Sub DispCtrl_(ByVal FormControl As Label, ByVal ThreadName As String, ByVal Text As String)


    Dim KVM_Pos As Integer = 0
    Dim fg_UnloadMe As Integer = 0
    Dim fg_SerData As Integer = 0
    Dim fg_BootError As Integer = 0
    Dim fg_ModeSelectUnLock As Integer = 0

    Dim fg_A_min As Integer = 0
    Dim fg_A_max As Integer = 0
    Dim fg_B_min As Integer = 0
    Dim fg_B_max As Integer = 0



    Private Sub tmr_TimeTicker_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmr_TimeTicker.Tick

        Static AutoTick As Integer


        With Me
            If .fg_UnloadMe = 0 Then .DispCalender()
        End With


        With Tp_OAI
            If .IO.M2_CCW.BitState = cls_PCIBoard.BitsState.eBit_ON Or .IO.M2_CW.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                'AutoTick += 1

                'If AutoTick >= 30 Then
                '    AutoTick = 0
                '    .M2_CCW.Trigger_OFF()
                '    .M2_CW.Trigger_OFF()
                'End If
            Else
                AutoTick = 0
            End If
        End With

    End Sub

    Private Sub DispCalender()

        Dim MyDay As Date = Now


        With Me
            .lbl_YearVal.Text = String.Format("{0:D4}", MyDay.Year)
            .lbl_MonthVal.Text = MyMonth(MyDay.Month)
            .lbl_DayVal.Text = String.Format("{0:D2}", MyDay.Day)
            .lbl_WeekDayVal.Text = MyWeekDay(MyDay.DayOfWeek)
            .lbl_TimeVal.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", MyDay.Hour, MyDay.Minute, MyDay.Second)

            If .lbl_DayVal.Text.EndsWith("1") Then
                If MyDay.Day = 11 Then
                    .lbl_thLabel.Text = "th"
                Else
                    .lbl_thLabel.Text = "st"
                End If
            ElseIf .lbl_DayVal.Text.EndsWith("2") Then
                If MyDay.Day = 12 Then
                    .lbl_thLabel.Text = "th"
                Else
                    .lbl_thLabel.Text = "nd"
                End If
            ElseIf .lbl_DayVal.Text.EndsWith("3") Then
                If MyDay.Day = 13 Then
                    .lbl_thLabel.Text = "th"
                Else
                    .lbl_thLabel.Text = "rd"
                End If
            End If
        End With

        Application.DoEvents()

    End Sub

    Private Sub frm_Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If Tp_OAI.Mode = SysAppMode.app_AutoRun Then
            e.Cancel = True
            Exit Sub
        End If

        If Not Me.fg_BootError < 2 Then
            With Tp_OAI
                Dim Total As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

                If Total <> 0 Then
                    MessageBox.Show("You are not allow to quit the system while the inspection does not complete.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    e.Cancel = True
                    Exit Sub
                End If
            End With

            If MessageBox.Show("Are you sure you want to shutdown this application now?", "Taping-OAI...", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                e.Cancel = True
                Exit Sub
            End If
        End If

        With Me
            .fg_UnloadMe = 1
            .TP_CAM1.Close()

            .tmr_IOMonitor.Enabled = False
            .tmr_ReadIO.Enabled = False
            .tmr_TimeTicker.Enabled = False
        End With

    End Sub

    Private Sub frm_Main_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim DispSize As System.Drawing.Size
        Dim DispPt As System.Drawing.Point


        With DispSize
            .Width = 1366
            .Height = 768
        End With

        With DispPt
            .X = 0
            .Y = 0
        End With


        With Tp_OAI
            With .IO
                .START_LD.Trigger_OFF()
                .STOP_LD.Trigger_OFF()
                .Auto_LD.Trigger_OFF()
                .PL_G.Trigger_OFF()
                .PL_R.Trigger_OFF()
                .PL_Y.Trigger_OFF()
            End With

            SysDataReset()
            .Mode = SysAppMode.app_NotInit
        End With

        With Me
            .Text = "Taping OAI - " & String.Format("{0:D4}.{1:D4}.{2:D3}.{3:D3}", My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build, My.Application.Info.Version.Revision)
            .Opacity = 1

            .lbl_SWUpdate.Visible = False
            .pic_iMTEST.Visible = True
            .cmd_Instruction.Enabled = True

            .fg_UnloadMe = 0
            .fg_SerData = 0
            .fg_BootError = 0
            .fg_ModeSelectUnLock = 0

            .fg_A_max = 0
            .fg_A_min = 0
            .fg_B_max = 0
            .fg_B_min = 0

            .KVM_Pos = 0

            With .lbl_Msg
                .BackColor = Color.Transparent
                .TextAlign = ContentAlignment.TopLeft

                Dim ft As New Font("Tahoma", 9, FontStyle.Regular)
                .Font = ft
                .ForeColor = Color.OrangeRed
            End With

            Try
                .Size = DispSize
                .Location = DispPt
            Catch ex As Exception

            End Try

            ReadRegData()

            If InitSerialPort() = Func_Ret_Success Then
                FTP_Init()
                DelAllFZ_ImgFile()
                'GetImage()

                With .cmd_Instruction
                    .Text = "Home"
                    .Image = Me.pic_Home.BackgroundImage
                End With

                With .lbl_Msg
                    .Text = "Press the 'Home' button to initialize the mechanism."
                End With

                With .tmr_TimeTicker
                    .Interval = 200
                    .Enabled = True
                End With

                With .tmr_IOMonitor
                    .Interval = 60
                    .Enabled = True
                End With

                With .tmr_ReadIO
                    .Interval = 45
                    .Enabled = True
                End With

                .lbl_StepperOffSet.Text = Tp_OAI.MotionSys(TP_OAI_Axis_Z).Offset

                .txt_A_max.Text = Tp_OAI.InspTapeSeal.A_max
                .txt_A_min.Text = Tp_OAI.InspTapeSeal.A_min
                .txt_B_max.Text = Tp_OAI.InspTapeSeal.B_max
                .txt_B_min.Text = Tp_OAI.InspTapeSeal.B_min
                .chk_CheckSealSetting.CheckState = Tp_OAI.InspTapeSeal.Mode

                With .cbo_StickerColor
                    .Items.Clear()
                    .Sorted = False

                    .Items.Add("R")
                    .Items.Add("G")
                    .Items.Add("B")

                    .SelectedIndex = Tp_OAI.DefectStickerColor
                End With

                With .cbo_nPulse
                    .Items.Clear()
                    .Sorted = False

                    For Each element In n_Pulse
                        Application.DoEvents()
                        .Items.Add(element.ToString)
                    Next

                    .SelectedIndex = 0
                End With

                With .cbo_Product
                    .Items.Clear()
                    .Sorted = False

                    For Each Element In ProdID
                        Application.DoEvents()
                        .Items.Add(Element)
                    Next

                    .SelectedIndex = 0
                End With


                StartMainThread()
                .fg_BootError = 0
            Else
                Tp_OAI.Mode = SysAppMode.app_sysError
                .fg_BootError = 1

                With .cmd_Instruction
                    .Text = "Close"
                    .Image = Me.pic_Shutdown.BackgroundImage
                End With
            End If
        End With

    End Sub

    Private Sub tmr_IOMonitor_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr_IOMonitor.Tick

        Static fg_Trg As Integer = 0

        With Tp_OAI
            Me.lbl_SeqNo.Text = "Seq.: " & .AutoSeqNo.ToString

            Me.pic_SensorR.BackColor = GrnLED_OnOff(.IO.RR.BitState)
            Me.pic_SensorL.BackColor = GrnLED_OnOff(.IO.RL.BitState)
            Me.pic_Motor.BackColor = RedLED_OnOff(.IO.M2_CW.BitState)
            Me.pic_SensorR_.BackColor = GrnLED_OnOff(.IO.RR.BitState)
            Me.pic_SensorL_.BackColor = GrnLED_OnOff(.IO.RL.BitState)
            Me.pic_Motor_.BackColor = RedLED_OnOff(.IO.M2_CW.BitState)
            Me.pic_SCover.BackColor = OrgLED_OnOff(.IO.S_Cover.BitState Xor cls_PCIBoard.BitsState.eBit_ON)
            Me.pic_SCover_.BackColor = OrgLED_OnOff(.IO.S_Cover.BitState Xor cls_PCIBoard.BitsState.eBit_ON)


            Me.lbl_VisionCtrl.BackColor = IIf(.IO.FZ_Run.BitState, GrnLED_OnOff(1), RedLED_OnOff(0))

            Dim int_RetVal As Integer = cls_MotionCtrl.MtrGetStatus(.MotionSys(TP_OAI_Axis_Z).DevH, cls_MotionCtrl.MotionGetStatus.MTN_LIMIT_STATUS, .MotionSys(TP_OAI_Axis_Z).MotionStatus)
            Me.pic_StepOrg.BackColor = GrnLED_OnOff((.MotionSys(TP_OAI_Axis_Z).MotionStatus And &H20) / &H20)
            Me.pic_StepOrg_.BackColor = GrnLED_OnOff((.MotionSys(TP_OAI_Axis_Z).MotionStatus And &H20) / &H20)


            'LED Blinking
            fg_Trg += 1

            If fg_Trg = 4 Then
                fg_Trg = 0
                Dim ButtonLight As Color = Nothing

                Select Case .Mode
                    Case SysAppMode.app_NotInit
                        If .IO.START_LD.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                            .IO.START_LD.Trigger_OFF()
                            .IO.Auto_LD.Trigger_ON()
                        Else
                            .IO.START_LD.Trigger_ON()
                            .IO.Auto_LD.Trigger_OFF()
                        End If

                        ButtonLight = OrgLED_OnOff(Tp_OAI.IO.START_LD.BitState)
                    Case SysAppMode.app_Auto
                        If .IO.START_LD.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                            .IO.START_LD.Trigger_OFF()
                            .IO.PL_Y.Trigger_ON()
                        Else
                            .IO.START_LD.Trigger_ON()
                            .IO.PL_Y.Trigger_OFF()
                        End If

                        .IO.STOP_LD.Trigger_ON()
                        .IO.PL_G.Trigger_OFF()
                        .IO.PL_R.Trigger_OFF()

                        ButtonLight = GrnLED_OnOff(Tp_OAI.IO.START_LD.BitState)
                    Case SysAppMode.app_AutoRun
                        If .IO.STOP_LD.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                            .IO.STOP_LD.Trigger_OFF()
                        Else
                            .IO.STOP_LD.Trigger_ON()
                        End If

                        .IO.START_LD.Trigger_ON()
                        .IO.PL_Y.Trigger_ON()
                        .IO.PL_G.Trigger_OFF()
                        .IO.PL_R.Trigger_OFF()

                        ButtonLight = RedLED_OnOff(Tp_OAI.IO.STOP_LD.BitState)
                    Case SysAppMode.app_Manu
                        .IO.START_LD.Trigger_OFF()
                        .IO.STOP_LD.Trigger_ON()

                        .IO.PL_G.Trigger_ON()
                        .IO.PL_R.Trigger_OFF()
                        .IO.PL_Y.Trigger_OFF()
                    Case SysAppMode.app_Setting
                        .IO.PL_G.Trigger_OFF()
                        .IO.PL_R.Trigger_OFF()
                        .IO.PL_Y.Trigger_OFF()

                        If .IO.START_LD.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                            .IO.START_LD.Trigger_OFF()
                            .IO.STOP_LD.Trigger_ON()
                        Else
                            .IO.START_LD.Trigger_ON()
                            .IO.STOP_LD.Trigger_OFF()
                        End If

                        Me.btn_StepRight.BackColor = GrnLED_OnOff(.IO.START_LD.BitState)
                        Me.btn_StepLeft.BackColor = RedLED_OnOff(.IO.STOP_LD.BitState)
                        ButtonLight = Color.DarkGreen
                    Case SysAppMode.app_sysError
                        With .IO
                            .PL_R.Trigger_ON()
                            .PL_G.Trigger_OFF()
                            .PL_Y.Trigger_OFF()
                            .START_LD.Trigger_OFF()

                            If .STOP_LD.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                                .STOP_LD.Trigger_OFF()

                                With Me.lbl_Msg
                                    .BackColor = Color.Transparent
                                    .TextAlign = ContentAlignment.TopCenter

                                    Dim ft As New Font("Georgia", 12, FontStyle.Italic)
                                    .Font = ft
                                    .ForeColor = Color.Red
                                End With

                                With Me.lbl_Alarm
                                    .BackColor = Color.Transparent
                                    .BorderStyle = BorderStyle.None
                                End With

                                With Me.lbl_Alarm_
                                    .BackColor = Color.Transparent
                                    .BorderStyle = BorderStyle.None
                                End With
                            Else
                                .STOP_LD.Trigger_ON()

                                With Me.lbl_Msg
                                    .BackColor = Color.Red
                                    .TextAlign = ContentAlignment.TopCenter

                                    Dim ft As New Font("Georgia", 12, FontStyle.Italic)
                                    .Font = ft
                                    .ForeColor = Color.MistyRose
                                End With

                                With Me.lbl_Alarm
                                    .BackColor = Color.Salmon
                                    .BorderStyle = BorderStyle.FixedSingle
                                End With

                                With Me.lbl_Alarm_
                                    .BackColor = Color.Salmon
                                    .BorderStyle = BorderStyle.FixedSingle
                                End With
                            End If

                            ButtonLight = RedLED_OnOff(.STOP_LD.BitState)
                        End With
                End Select

                Me.cmd_Instruction.BackColor = ButtonLight
                Application.DoEvents()
            End If
        End With

    End Sub

    Private Sub pic_Motor__MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pic_Motor_.MouseDown

        With Tp_OAI
            If Not .Mode = SysAppMode.app_Manu Then Exit Sub

            If e.Button = Windows.Forms.MouseButtons.Left Then
                If Me.chk_Motor.CheckState = CheckState.Checked Then
                    .IO.M2_CCW.Trigger_ON()
                Else
                    .IO.M2_CW.Trigger_ON()
                End If
            End If
        End With

    End Sub

    Private Sub pic_Motor__MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pic_Motor_.MouseUp

        With Tp_OAI
            .IO.M2_CW.Trigger_OFF()
            .IO.M2_CCW.Trigger_OFF()
        End With

    End Sub

    Private Sub pic_Stepper__MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pic_Stepper_.MouseDown

        If e.Button = Windows.Forms.MouseButtons.Right Then
            Me.pic_Stepper_.ContextMenuStrip.Show()
        End If

        If e.Button = Windows.Forms.MouseButtons.Left Then
            With Tp_OAI
                If .SysData(0).InspCnt = 0 Then
                    mn_StepperMove()
                Else
                    MessageBox.Show("The function is temporary interlock because 'AutoRun' operation is still in progress.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End With
        End If

    End Sub

    Private Sub sub_StepperFR_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles sub_StepperFR.Click

        With Tp_OAI
            If .IO.M1_FR.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                .Mode = SysAppMode.app_NotInit
                .IO.M1_FR.Trigger_ON()

                With Me
                    With .cmd_Instruction
                        .Text = "Home"
                        .Image = Me.pic_Home.BackgroundImage
                    End With

                    With .lbl_Msg
                        .Text = "Press the 'Home' button to initialize the mechanism."
                    End With

                    .sub_StepperFR.Text = "Stepper Motor Brake"
                End With
            Else
                .IO.M1_FR.Trigger_OFF()
                Me.sub_StepperFR.Text = "Stepper Motor Free"
            End If
        End With

    End Sub

    Private Sub cmd_Instruction_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmd_Instruction.Click

        Me.cmd_Instruction.Enabled = False

        AppCommand()
        Me.cmd_Instruction.Enabled = True

    End Sub

    Private Sub AppCommand()

        Dim regSubKeyCamBGS_h As RegistryKey = regKey.CreateSubKey("Software\az_IOLogics\TapingOAI\CAM_Param\BGS_high")
        Dim regSubKeyCamBGS_L As RegistryKey = regKey.CreateSubKey("Software\az_IOLogics\TapingOAI\CAM_Param\BGS_Low")
        Dim regSubKeyCamDensity As RegistryKey = regKey.CreateSubKey("Software\az_IOLogics\TapingOAI\CAM_Param\DensityLevel")
        Dim regSubKeyFilterType As RegistryKey = regKey.CreateSubKey("Software\az_IOLogics\TapingOAI\CAM_Param\FilterType")

        Dim FuncRet As Integer = 0


        With Tp_OAI
            Select Case .Mode
                Case SysAppMode.app_NotInit
                    FuncRet = thd_MoveOrg(TP_OAI_Axis_Z, cls_MotionCtrl.MotionDirection.MTN_CCW)

                    If FuncRet = Func_Ret_Success Then
                        If Not .MotionSys(TP_OAI_Axis_Z).Offset = 0 Then
                            FuncRet = thd_MoveStep(TP_OAI_Axis_Z, .MotionSys(TP_OAI_Axis_Z).Offset)
                        End If

                        .Mode = SysAppMode.app_Auto
                        .IO.Auto_LD.Trigger_ON()
                        .IO.BZ.Trigger_ON()

                        Dim Tmr As Integer = My.Computer.Clock.TickCount

                        Do Until My.Computer.Clock.TickCount > Tmr + 500
                            Application.DoEvents()
                        Loop


                        With Me
                            With .cmd_Instruction
                                .Text = "Start"
                                .Image = Me.pic_Play.BackgroundImage
                            End With

                            With .lbl_Msg
                                .Text = "Ready to perform 'Auto' operation. The system will prompt for Data Entry if it was not done yet!"
                            End With

                            .TabControl1.SelectedIndex = 0
                        End With

                        .IO.BZ.Trigger_OFF()
                    End If
                Case SysAppMode.app_Auto
                    If .SysData(0).LotNo = "" Then
                        ClearDisp()
                        frm_DataEntry.ShowDialog()

                        If .SysData(0).LotNo <> "" Then
                            FuncRet = thd_MoveOrg(TP_OAI_Axis_Z, cls_MotionCtrl.MotionDirection.MTN_CCW)

                            If FuncRet = Func_Ret_Success Then
                                If Not .MotionSys(TP_OAI_Axis_Z).Offset = 0 Then
                                    FuncRet = thd_MoveStep(TP_OAI_Axis_Z, .MotionSys(TP_OAI_Axis_Z).Offset)
                                End If
                            End If

                            Me.lbl_MarkChar1.Text = .SysData(0).MarkData1(0)
                            Me.lbl_MarkChar2.Text = .SysData(0).MarkData2(0)

                            Dim AllWeekCode As String = String.Empty

                            For iLp As Integer = 0 To .SysData(0).MarkData2.GetUpperBound(0)
                                Application.DoEvents()
                                AllWeekCode &= .SysData(0).MarkData2(iLp) & " "

                                If .SysData(0).MarkData2(iLp).Contains("O") Then
                                    .SysData(0).MarkData2(iLp) = .SysData(0).MarkData2(iLp).Replace("O", "0")
                                End If
                            Next

                            Me.ToolTip1.SetToolTip(Me.lbl_MarkChar2, AllWeekCode.Trim)
                            Me.lbl_LotNo.Text = .SysData(0).LotNo & " (" & .SysData(0).Prod & ")"
                            Me.lbl_PLotNo.Text = .SysData(0).P_LotNo
                            Me.lbl_EmpNo.Text = .SysData(0).Insp
                            Me.lbl_Total.Text = "0/" & .SysData(0).Acceptance
                            Me.lbl_InspTapeType.Visible = Val(.SysData(0).InspType)

                            Dim TimeStart As Date = Now
                            .SysData(0).InspDate = TimeStart

                            If .SysData(0).RedoInsp = True Then
                                Me.lbl_LotName.Text = "Lot No. (R)"
                            Else
                                Me.lbl_LotName.Text = "Lot No."
                            End If

                            Me.lbl_Date.Text = String.Format("{0:D2}-{1:D2}-{2:D4} - {3:D2}:{4:D2}", TimeStart.Day, TimeStart.Month, TimeStart.Year, TimeStart.Hour, TimeStart.Minute)
                            Me.lbl_InspTapeType.Visible = Val(.SysData(0).InspType)

                            Dim MyImgFile As String = My.Application.Info.DirectoryPath & "\az2011.jpg"
                            Me.PictureBox4.BackgroundImage = Image.FromFile(MyImgFile)
                            Me.PictureBox9.BackgroundImage = Image.FromFile(MyImgFile)
                            Me.PictureBox11.BackgroundImage = Image.FromFile(MyImgFile)

                            DeleteSysTempFile()

                            With Me.lst_RawData
                                .Items.Clear()
                                .Refresh()
                            End With

                            Dim RetVal_ As String = String.Empty
                            Dim TotalChr As String = String.Empty

                            FZ_SerCmd("SCENE " & String.Format("{0:D2}", .SelectedSceneNo), RetVal_)
                            Thread.Sleep(300)

                            For iLp As Integer = 0 To 0
                                Application.DoEvents()
                                Me.SetInspChr(.SysData(0).MarkData1(iLp) & "A038", "22")
                            Next

                            .ColerationChk = 0

                            For iLp As Integer = 0 To .SysData(0).MarkData2.GetUpperBound(0)
                                Application.DoEvents()

                                If TotalChr = "" Then
                                    TotalChr &= .SysData(0).MarkData2(iLp)
                                Else
                                    If TotalChr.IndexOf(.SysData(0).MarkData2(iLp)) < 0 Then
                                        TotalChr &= .SysData(0).MarkData2(iLp)
                                    End If
                                End If

                                If .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "A" Then
                                    .ColerationChk = 0
                                    TotalChr &= "1468"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "B" Then
                                    .ColerationChk = 0
                                    TotalChr &= "358DG"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "C" Then
                                    .ColerationChk = 0
                                    TotalChr &= "016"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "D" Then
                                    .ColerationChk = 0
                                    TotalChr &= "018B"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "E" Then
                                    .ColerationChk = 0
                                    TotalChr &= "1378F"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "F" Then
                                    .ColerationChk = 0
                                    TotalChr &= "178"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "G" Then
                                    .ColerationChk = 0
                                    TotalChr &= "0368B"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "H" Then
                                    .ColerationChk = 0
                                    TotalChr &= "18"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "J" Then
                                    .ColerationChk = 0
                                    TotalChr &= "01A"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "K" Then
                                    .ColerationChk = 0
                                    TotalChr &= "8HX"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "L" Then
                                    .ColerationChk = 0
                                    TotalChr &= "1E"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "M" Then
                                    .ColerationChk = 0
                                    TotalChr &= "8H"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "N" Then
                                    .ColerationChk = 0
                                    TotalChr &= "R"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "P" Then
                                    .ColerationChk = 0
                                    TotalChr &= "0198"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "R" Then
                                    .ColerationChk = 0
                                    TotalChr &= "0138KN"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "S" Then
                                    .ColerationChk = 0
                                    TotalChr &= "0589"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "T" Then
                                    .ColerationChk = 0
                                    TotalChr &= "0187"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "U" Then
                                    .ColerationChk = 0
                                    TotalChr &= "018"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "V" Then
                                    .ColerationChk = 0
                                    TotalChr &= "0X"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "W" Then
                                    .ColerationChk = 0
                                    TotalChr &= "08"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "X" Then
                                    .ColerationChk = 0
                                    TotalChr &= "08AK"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "Y" Then
                                    .ColerationChk = 0
                                    TotalChr &= "01T"
                                ElseIf .SysData(0).MarkData2(iLp).Substring(.SysData(0).MarkData2(iLp).Length - 2, 1) = "Z" Then
                                    .ColerationChk = 0
                                    TotalChr &= "2"
                                Else
                                    .ColerationChk = 0
                                    TotalChr &= "ABQSU"
                                End If
                            Next


                            Me.SetInspChr(TotalChr, "23")


                            .BGSmaxSet = regSubKeyCamBGS_h.GetValue(String.Format("{0:D2}", .SelectedSceneNo), "110")
                            .BGSminSet = regSubKeyCamBGS_L.GetValue(String.Format("{0:D2}", .SelectedSceneNo), "105")
                            .DensityLvl = regSubKeyCamDensity.GetValue(String.Format("{0:D2}", .SelectedSceneNo), "190")
                            .FilterType = regSubKeyFilterType.GetValue(String.Format("{0:D2}", .SelectedSceneNo), "1")

                            'Set Filtering : Erosion
                            Me.FZ_SerCmd("UNITDATA 15 121 4")

                            'Set Filter Size : 0=3x3, 1=5x5
                            Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)

                            'Set min BGS Level
                            Me.FZ_SerCmd("UNITDATA 15 124 " & .BGSminSet)

                            'Set max BGS Level
                            Me.FZ_SerCmd("UNITDATA 15 125 " & .BGSmaxSet)
                        End If

                        Exit Sub
                    End If


                    Dim InterlockErrNo As Integer = 0

                    If Me.SysInterlock(InterlockErrNo) < 0 Then
                        SetAlarm(InterlockErrNo)
                        Exit Sub
                    End If

                    With Me
                        With .cmd_Instruction
                            .Text = "Stop"
                            .Image = Me.pic_Stop.BackgroundImage
                        End With

                        With .lbl_Msg
                            .Text = "Auto Running..."
                        End With
                    End With

                    If .TimeEllapse = 0 Then
                        .TimeEllapse = My.Computer.Clock.TickCount
                        CreateLotDataFolder()
                    End If

                    Dim RetVal As String = String.Empty
                    FZ_SerCmd("SCENE " & String.Format("{0:D2}", .SelectedSceneNo), RetVal)

                    'Set Filtering : Erosion
                    Me.FZ_SerCmd("UNITDATA 15 121 4")


                    .BGSmaxSet = regSubKeyCamBGS_h.GetValue(String.Format("{0:D2}", .SelectedSceneNo), "110")
                    .BGSminSet = regSubKeyCamBGS_L.GetValue(String.Format("{0:D2}", .SelectedSceneNo), "105")
                    .DensityLvl = regSubKeyCamDensity.GetValue(String.Format("{0:D2}", .SelectedSceneNo), "190")


                    'Set Filter Size : 0=3x3, 1=5x5
                    Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)

                    'Set min BGS Level
                    Me.FZ_SerCmd("UNITDATA 15 124 " & .BGSminSet)

                    'Set max BGS Level
                    Me.FZ_SerCmd("UNITDATA 15 125 " & .BGSmaxSet)

                    .Mode = SysAppMode.app_AutoRun
                Case SysAppMode.app_AutoRun
                    If .AutoSeqNo = 10 Then
                        Exit Select
                    End If

                    With Me
                        With .cmd_Instruction
                            .Text = "Start"
                            .Image = Me.pic_Play.BackgroundImage
                        End With

                        With .lbl_Msg
                            .Text = "Ready to perform 'Auto' operation. The system will prompt for Data Entry if it was not done yet!"
                        End With
                    End With

                    '.IO.M2_CCW.Trigger_OFF()
                    '.IO.M2_CW.Trigger_OFF()

                    .Mode = SysAppMode.app_Auto
                Case SysAppMode.app_sysError
                    SysErrorReset()
            End Select
        End With

    End Sub

    Private Function SetInspChr(ByVal InspData As String, ByVal ProgNo As String) As Integer

        Dim InspChar As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim _InspChar As String = "0123456789ABCDeFGHIJKLMNOPQRSTUVWXYZ"
        Dim RET As String = String.Empty


        Dim _chrSet As String = IIf(InspData.Contains("e"c), _InspChar, InspChar)

        For iLp As Integer = 140 To 175
            Application.DoEvents()

            If Not InspData.IndexOf(_chrSet.Substring(iLp - 140, 1)) < 0 Then
                Me.FZ_SerCmd("UNITDATA " & ProgNo & " " & iLp.ToString & " 1", RET)
            Else
                Me.FZ_SerCmd("UNITDATA " & ProgNo & " " & iLp.ToString & " 0", RET)
            End If
        Next

        If Not InspData.IndexOf("o") < 0 Then
            Me.FZ_SerCmd("UNITDATA " & ProgNo & " " & "164" & " 1", RET)
        End If

    End Function

    Private Sub CreateLotDataFolder()

        With Tp_OAI
            Dim InspDate As String = String.Format("{0:D2}-{1:D2}-{2:D4}", .SysData(0).InspDate.Day, .SysData(0).InspDate.Month, .SysData(0).InspDate.Year)
            Dim PathName As String = .LotDataPath & "\" & .SysData(0).Prod & "\" & InspDate & "\" & .SysData(0).LotNo


            If Not Val(.SysData(0).InspType) = 0 Then
                PathName &= "\SealInsp"
            Else
                PathName &= "\MarkInsp"
            End If


            If .SysData(0).LotNo.IndexOf("V_") < 0 Then
                If My.Computer.FileSystem.DirectoryExists(PathName) = False Then
                    My.Computer.FileSystem.CreateDirectory(PathName)
                Else
                    Try
                        My.Computer.FileSystem.DeleteDirectory(PathName, FileIO.DeleteDirectoryOption.DeleteAllContents)
                    Catch ex As Exception

                    End Try
                End If
            Else
                Try
                    My.Computer.FileSystem.DeleteDirectory(PathName, FileIO.DeleteDirectoryOption.DeleteAllContents)
                Catch ex As Exception

                End Try
            End If
        End With

    End Sub

    Private Function SysInterlock(Optional ByRef ErrorNo As Integer = 0) As Integer

        With Tp_OAI
            If .IO.RR.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                ErrorNo = 11
                Return -1
            End If

            If .IO.RL.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                ErrorNo = 12
                Return -1
            End If

            'If (.IO.S_Cover.BitState Xor cls_PCIBoard.BitsState.eBit_ON) = cls_PCIBoard.BitsState.eBit_OFF Then
            '    ErrorNo = 13
            '    Return -1
            'End If

            If Not .IO.FZ_Run.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                FZ_SerCmd("SCRSWITCH")

                Me.Opacity = 0.7
                frm_LngDelay.ShowDialog()
                Me.Opacity = 1

                If .IO.FZ_Run.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                    ErrorNo = 14
                    Return -1
                End If
            End If

            If .IO.M2_Alarm.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                ErrorNo = 15
                Return -1
            End If

            If .IO.CB_Alarm.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                ErrorNo = 16
                Return -1
            End If
        End With

        Return 0

    End Function

    Private Sub tmr_ReadIO_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr_ReadIO.Tick

        Static TmrTickCnt As Integer


        Me.tmr_ReadIO.Enabled = False

        With Tp_OAI
            Select Case .Mode
                Case Is = SysAppMode.app_Auto, SysAppMode.app_NotInit
                    If .Mode = SysAppMode.app_Auto Then
                        If .IO.pb_STOP.BitState = cls_PCIBoard.BitsState.eBit_ON And .IO.pb_START.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                            TmrTickCnt += 1

                            If TmrTickCnt >= 70 Then
                                TmrTickCnt = 0
                                SysDataReset()
                                ClearDisp()
                                .IO.BZ.Trigger_ON()

                                Do Until .IO.pb_STOP.BitState = cls_PCIBoard.BitsState.eBit_OFF And .IO.pb_START.BitState = cls_PCIBoard.BitsState.eBit_OFF
                                    Application.DoEvents()
                                Loop

                                .IO.BZ.Trigger_OFF()
                            End If
                        Else
                            TmrTickCnt = 0
                        End If
                    End If

                    If .IO.pb_START.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                        If .IO.pb_STOP.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                            Do Until .IO.pb_START.BitState = cls_PCIBoard.BitsState.eBit_OFF
                                Application.DoEvents()
                            Loop

                            TmrTickCnt = 0
                            Me.cmd_Instruction.PerformClick()
                        End If
                    End If
                Case Is = SysAppMode.app_AutoRun
                    TmrTickCnt = 0

                    With .IO
                        If .pb_STOP.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                            Do Until .pb_START.BitState = cls_PCIBoard.BitsState.eBit_OFF
                                Application.DoEvents()
                            Loop

                            Me.cmd_Instruction.PerformClick()
                        End If

                        If .RR.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                            SetAlarm(21)
                        End If

                        If .RL.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                            SetAlarm(22)
                        End If

                        'If Not .S_Cover.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                        '    SetAlarm(23)
                        'End If

                        If .FZ_Run.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                            SetAlarm(24)
                        End If

                        If Not .M2_Alarm.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                            SetAlarm(25)
                        End If

                        If Not .CB_Alarm.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                            SetAlarm(26)
                        End If
                    End With
                Case Is = SysAppMode.app_Manu
                    TmrTickCnt = 0

                    If .IO.pb_START.BitState = cls_PCIBoard.BitsState.eBit_ON And .IO.pb_STOP.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                        If .SysData(0).InspCnt = 0 Then
                            mn_StepperMove()
                        Else
                            MessageBox.Show("The function is temporary interlock because 'AutoRun' operation is still in progress.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If

                        Do Until .IO.pb_START.BitState = cls_PCIBoard.BitsState.eBit_OFF
                            Application.DoEvents()
                        Loop
                    End If
                Case Is = SysAppMode.app_Setting
                    TmrTickCnt = 0

                    If .IO.pb_STOP.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                        If .IO.pb_START.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                            mn_StepperMove_n()
                        End If
                    End If

                    If .IO.pb_START.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                        If .IO.pb_STOP.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                            mn_StepperMove_n(1)
                        End If
                    End If
                Case Is = SysAppMode.app_sysError
                    TmrTickCnt = 0

                    If .IO.pb_EMG.BitState = cls_PCIBoard.BitsState.eBit_ON And .IO.pb_STOP.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                        Me.AppCommand()
                    End If
            End Select

            If .IO.pb_EMG.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                If Not .Mode = SysAppMode.app_sysError Then
                    SysDataReset()
                    SetAlarm(1)
                End If
            End If


            Dim Total As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

            If Total <> 0 Then
                If .IO.RR.BitState = cls_PCIBoard.BitsState.eBit_OFF And .IO.RL.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                    If Not .Mode = SysAppMode.app_sysError Then
                        SetAlarm(91)
                        SysDataReset()
                    End If
                End If
            End If

            If .IO.pb_Auto.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                If Me.KVM_Pos <> 1 Then
                    KVM_Pos = 1
                    .IO.KVM_PC.Trigger_ON()
                    Thread.Sleep(100)
                    .IO.KVM_PC.Trigger_OFF()
                End If
            Else
                If Me.KVM_Pos <> 2 Then
                    KVM_Pos = 2
                    .IO.KVM_VZ.Trigger_ON()
                    Thread.Sleep(100)
                    .IO.KVM_VZ.Trigger_OFF()
                End If
            End If
        End With

        Me.tmr_ReadIO.Enabled = True
        Application.DoEvents()

    End Sub

    Private Sub SysErrorReset()

        With Me
            With .lbl_Msg
                .BackColor = Color.Transparent
                .TextAlign = ContentAlignment.TopLeft

                Dim ft As New Font("Tahoma", 9, FontStyle.Regular)
                .Font = ft
                .ForeColor = Color.OrangeRed
            End With

            With .ToolTip1
                .Hide(Me.lbl_Msg)
            End With

            .lbl_Alarm.Visible = False
            .lbl_Alarm_.Visible = False
            .lbl_ErrDesp.Visible = False
        End With

        With Tp_OAI
            .IO.BZ.Trigger_OFF()
            .Mode = .ErrorReset.AppMode

            If .Mode = SysAppMode.app_AutoRun Then
                .Mode = SysAppMode.app_Auto
            End If

            If .ErrorReset.ErrorCode = 1 Then
                .Mode = SysAppMode.app_NotInit
            End If

            Select Case .Mode
                Case Is = SysAppMode.app_NotInit
                    With Me
                        With .cmd_Instruction
                            .Image = Me.pic_Home.BackgroundImage
                            .Text = "Home"
                        End With

                        With .lbl_Msg
                            .Text = "Press the 'Home' button to initialize the mechanism."
                        End With
                    End With
                Case Is = SysAppMode.app_Auto
                    With Me
                        With .cmd_Instruction
                            .Image = Me.pic_Play.BackgroundImage
                            .Text = "Start"
                        End With

                        With .lbl_Msg
                            .Text = "Ready to perform 'Auto' operation. The system will prompt for Data Entry if it was not done yet!"
                        End With
                    End With
                Case Is = SysAppMode.app_Manu
                    With Me
                        With .cmd_Instruction
                            .Image = Me.pic_Stop.BackgroundImage
                            .Text = "Stop"
                        End With

                        With .lbl_Msg
                            .Text = "Manual Operation..."
                        End With
                    End With
            End Select
        End With

        With Me
            If .TabControl1.SelectedIndex <> 0 Then
                .fg_ModeSelectUnLock = 1
                .TabControl1.SelectedIndex = 0
            End If
        End With

    End Sub

    Private Sub SetAlarm(ByVal ErrorNo As Integer)

        With Me
            With Me.lbl_Msg
                .BackColor = Color.Red
                .TextAlign = ContentAlignment.TopCenter

                Dim ft As New Font("Georgia", 12, FontStyle.Italic)
                .Font = ft
                .ForeColor = Color.MistyRose


                'Update msg at different thread
                If .InvokeRequired Then
                    .Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_Msg, Thread.CurrentThread.Name, Tp_OAI.ErrorDisp(ErrorNo).str_ErrorDesc})
                Else
                    .Text = Tp_OAI.ErrorDisp(ErrorNo).str_ErrorDesc
                End If
            End With

            With .lbl_ErrDesp
                'Update msg at different thread
                If .InvokeRequired Then
                    .Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_ErrDesp, Thread.CurrentThread.Name, Tp_OAI.ErrorDisp(ErrorNo).str_ErrorToChk})
                Else
                    .Text = Tp_OAI.ErrorDisp(ErrorNo).str_ErrorToChk
                    .Visible = True
                End If
            End With

            With .lbl_Alarm
                'Update msg at different thread
                If .InvokeRequired Then
                    .Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_Alarm, Thread.CurrentThread.Name, Tp_OAI.ErrorDisp(ErrorNo).str_ErrorDesc})
                Else
                    .Text = Tp_OAI.ErrorDisp(ErrorNo).str_ErrorDesc
                End If

                Try
                    If Tp_OAI.Mode <> SysAppMode.app_AutoRun Then .Visible = True
                Catch ex As Exception

                End Try
            End With

            With .lbl_Alarm_
                'Update msg at different thread
                If .InvokeRequired Then
                    .Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_Alarm_, Thread.CurrentThread.Name, Tp_OAI.ErrorDisp(ErrorNo).str_ErrorDesc})
                Else
                    .Text = Tp_OAI.ErrorDisp(ErrorNo).str_ErrorDesc
                End If

                Try
                    If Tp_OAI.Mode <> SysAppMode.app_AutoRun Then .Visible = True
                Catch ex As Exception

                End Try
            End With

            With .cmd_Instruction
                .Image = Me.pic_Stop.BackgroundImage

                'Update msg at different thread
                If .InvokeRequired Then
                    Try
                        .Invoke(New frm_Main.DispMsg_(AddressOf Me.DispControl), New Object() {Me.cmd_Instruction, Thread.CurrentThread.Name, "Reset"})
                    Catch ex As Exception

                    End Try
                Else
                    .Text = "Reset"
                End If
            End With

            'With .ToolTip1
            '    If Tp_OAI.Mode <> SysAppMode.app_AutoRun Then
            '        Try
            '            .SetToolTip(Me.lbl_Msg, Tp_OAI.ErrorDisp(ErrorNo).str_ErrorToChk)
            '            .ShowAlways = True

            '            Dim pt As Point
            '            pt.X = Me.lbl_Msg.Location.X
            '            pt.Y = Me.lbl_Msg.Location.Y

            '            .Show(Tp_OAI.ErrorDisp(ErrorNo).str_ErrorToChk, Me, pt)
            '        Catch ex As Exception

            '        End Try
            '    End If
            'End With
        End With

        With Tp_OAI
            With .ErrorReset
                .ErrorCode = ErrorNo

                If Me.TabControl1.InvokeRequired Then
                    .AppMode = IIf(Tp_OAI.Mode <> SysAppMode.app_sysError, Tp_OAI.Mode, SysAppMode.app_Auto)
                Else
                    .AppMode = IIf(Tp_OAI.Mode <> SysAppMode.app_sysError, Tp_OAI.Mode, Me.TabControl1.SelectedIndex)
                End If

            End With

            .Mode = SysAppMode.app_sysError
            .IO.BZ.Trigger_ON()
        End With

    End Sub


    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged

        With Tp_OAI
            If Me.fg_ModeSelectUnLock = 0 Then .Mode = Me.TabControl1.SelectedIndex
            Me.fg_ModeSelectUnLock = 0

            With Me.lbl_Msg
                .BackColor = Color.Transparent
                .TextAlign = ContentAlignment.TopLeft

                Dim ft As New Font("Tahoma", 9, FontStyle.Regular)
                .Font = ft
                .ForeColor = Color.OrangeRed
            End With


            If .Mode = SysAppMode.app_Auto Then
                With Me
                    With .cmd_Instruction
                        .Image = Me.pic_Play.BackgroundImage
                        .Text = "Start"
                    End With

                    With .lbl_Msg
                        .Text = "Ready to perform 'Auto' operation. The system will prompt for Data Entry if it was not done yet!"
                    End With
                End With

                If .SysData(0).InspCnt = 0 And Not .SysData(0).LotNo.ToUpper.IndexOf("V_") < 0 Then
                    .SysData(1).Prod = ProdID(Me.cbo_Product.SelectedIndex)
                    .SysData(1).P_LotNo = Me.txt_DmyLotNo.Text.ToUpper.Trim

                    Try
                        .SelectedSceneNo = Val(SceneNoDB(Array.IndexOf(ProdID, .SysData(1).Prod)))
                    Catch ex As Exception
                        .SelectedSceneNo = 0
                    End Try

                    Me.txt_DataBlock1.Text = Me.txt_DataBlock1.Text.Replace(", ", ",")
                    Me.txt_DataBlock2.Text = Me.txt_DataBlock2.Text.Replace(", ", ",")

                    Dim BlockData1() As String = Me.txt_DataBlock1.Text.Split(","c)
                    Dim BlockData2() As String = Me.txt_DataBlock2.Text.Split(","c)

                    ReDim .SysData(0).MarkData1(BlockData1.GetUpperBound(0))
                    ReDim .SysData(0).MarkData2(BlockData2.GetUpperBound(0))

                    Array.Copy(BlockData1, .SysData(0).MarkData1, BlockData1.Length)
                    Array.Copy(BlockData2, .SysData(0).MarkData2, BlockData2.Length)

                    Me.lbl_MarkChar1.Text = .SysData(0).MarkData1(0)
                    Me.lbl_MarkChar2.Text = .SysData(0).MarkData2(0)
                End If
            ElseIf .Mode = SysAppMode.app_Manu Then
                With Me
                    With .cmd_Instruction
                        .Image = Me.pic_Stop.BackgroundImage
                        .Text = "Stop"
                    End With

                    With .lbl_Msg
                        .Text = "Manual Operation..."
                    End With
                End With
            ElseIf .Mode = SysAppMode.app_Setting Then
                With Me
                    With .cbo_Product
                        .Focus()
                    End With
                End With
            End If
        End With

    End Sub

    Private Sub pic_Stepper__Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pic_Stepper_.Click


    End Sub

    Private Function mn_StepperMove(Optional ByVal Pulse As Integer = -60) As Integer

        With Me
            .pic_Stepper.BackColor = RedLED_OnOff(1)
            .pic_Stepper_.BackColor = RedLED_OnOff(1)

            Pulse *= IIf(Me.chk_Stepper.CheckState = CheckState.Checked, -1, 1)

            If Tp_OAI.Mode = SysAppMode.app_AutoRun Then
                'If Pulse > 0 Then Pulse *= -1
            Else
            End If

            Dim FuncRet As Integer = thd_MoveStep(TP_OAI_Axis_Z, Pulse)

            .pic_Stepper.BackColor = RedLED_OnOff(0)
            .pic_Stepper_.BackColor = RedLED_OnOff(0)

            Return FuncRet
        End With

    End Function

    Private Function mn_StepperMove_n(Optional ByVal Direction As Integer = 0) As Integer

        With Me
            Dim Move_n As Integer = n_Pulse(.cbo_nPulse.SelectedIndex) * IIf(Direction = 0, -1, 1)

            .pic_Stepper.BackColor = RedLED_OnOff(1)
            .pic_Stepper_.BackColor = RedLED_OnOff(1)

            Dim FuncRet As Integer = thd_MoveStep(TP_OAI_Axis_Z, Move_n)

            If FuncRet = Func_Ret_Success Then
                Tp_OAI.MotionSys(TP_OAI_Axis_Z).Offset += Move_n
                .lbl_StepperOffSet.Text = Tp_OAI.MotionSys(TP_OAI_Axis_Z).Offset.ToString

                Dim regSubKey_ As RegistryKey = regKey.CreateSubKey("Software\az_IOLogics\TapingOAI\Motion")
                regSubKey_.SetValue("OffSet_0", Tp_OAI.MotionSys(TP_OAI_Axis_Z).Offset.ToString)
            End If

            .pic_Stepper.BackColor = RedLED_OnOff(0)
            .pic_Stepper_.BackColor = RedLED_OnOff(0)

            Return FuncRet
        End With

    End Function

    Private Sub TabControl1_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles TabControl1.Selecting

        With Tp_OAI
            If .Mode = SysAppMode.app_NotInit Or .Mode = SysAppMode.app_sysError Or .Mode = SysAppMode.app_AutoRun Then
                If Me.fg_ModeSelectUnLock = 0 Then
                    If .Mode = SysAppMode.app_Auto Or (.Mode = SysAppMode.app_NotInit And Me.TabControl1.SelectedIndex <> 0) Then
                        e.Cancel = True
                        If .Mode = SysAppMode.app_NotInit Then MessageBox.Show("The system need to be initialize first. Please press 'START' button to do so!", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        Me.fg_ModeSelectUnLock = 1
                        Exit Sub
                    End If
                End If
            End If

            If Me.TabControl1.SelectedIndex = 2 Then
                frm_PwdEntry.ShowDialog()

                If .AuthenticalCode = "" Then
                    e.Cancel = True
                    Exit Sub
                Else
                    If Not .AuthenticalCode = .Authentication Then
                        MessageBox.Show("Password unmatch...!", "Taping OAI...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
            End If
        End With

        With Me
            If .fg_A_max <> 0 Or .fg_A_min <> 0 Or .fg_B_max <> 0 Or .fg_B_min <> 0 Then
                e.Cancel = True
                MessageBox.Show("Range setting for upper and lower limit must be numeric.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

                If .fg_A_max <> 0 Then .txt_A_max.Focus() : Exit Sub
                If .fg_B_max <> 0 Then .txt_B_max.Focus() : Exit Sub
                If .fg_A_min <> 0 Then .txt_A_min.Focus() : Exit Sub
                If .fg_B_min <> 0 Then .txt_B_min.Focus() : Exit Sub
            End If
        End With

    End Sub

    Public Function FZ_SerCmd(ByVal strCmd As String, Optional ByRef strMeasDat As String = "") As Integer

        Dim str_Buffer As String = String.Empty
        Dim Buffer() As Char = ""


        With Me.TP_CAM1
            If .IsOpen = False Then
                .Open()
                Thread.Sleep(100)
            End If

            .DiscardInBuffer()
            str_Buffer = strCmd & vbCrLf
            .Write(str_Buffer)


            'Simply Delay Operation
            Thread.Sleep(80)

            Dim WaitReplyTimer As Integer = My.Computer.Clock.TickCount
            Dim fgData As Integer = 0

            Do Until Not .BytesToRead = 0
                Application.DoEvents()

                If My.Computer.Clock.TickCount > WaitReplyTimer + 900 Then
                    fgData = -1 : Exit Do
                End If
            Loop

            If fgData < 0 Then Return Func_Ret_Fail

            Dim ReadByteSize As Integer = .BytesToRead
            WaitReplyTimer = My.Computer.Clock.TickCount
            str_Buffer = ""


            If Not strCmd.IndexOf("MEASURE") < 0 Then
                If Not strCmd.IndexOf("/E") < 0 Or Not strCmd.IndexOf("/C") < 0 Then
                    Do Until ReadByteSize = 0
                        ReDim Buffer(ReadByteSize)
                        .Read(Buffer, 0, ReadByteSize)

                        For int_Dmy = 0 To Buffer.GetUpperBound(0)
                            'Application.DoEvents()

                            If Not Buffer(int_Dmy) = Nothing Then
                                If Buffer(int_Dmy) = vbCr Then
                                    str_Buffer &= " "
                                ElseIf Buffer(int_Dmy) = vbLf Then
                                    str_Buffer &= ""
                                Else
                                    str_Buffer &= Buffer(int_Dmy)
                                End If
                            End If
                        Next

                        ReadByteSize = .BytesToRead
                    Loop
                Else
                    Dim WaitVC_Respond As Integer = 700
                    Dim TimeOver As Integer = 0


                    Do Until Not str_Buffer.IndexOf("9999") < 0 Or Not str_Buffer.IndexOf("7777") < 0 Or Not str_Buffer.IndexOf("6666") < 0 Or Not str_Buffer.IndexOf("9990") < 0 Or Not str_Buffer.IndexOf("8888") < 0 Or Not str_Buffer.IndexOf("-1") < 0
                        If My.Computer.Clock.TickCount > WaitReplyTimer + WaitVC_Respond And .BytesToRead = 0 Then TimeOver = 1 : Exit Do
                        If Tp_OAI.Mode = SysAppMode.app_sysError Then TimeOver = 1 : Exit Do

                        ReDim Buffer(ReadByteSize)
                        .Read(Buffer, 0, ReadByteSize)

                        For int_Dmy = 0 To Buffer.GetUpperBound(0)
                            'Application.DoEvents()

                            If Not Buffer(int_Dmy) = Nothing Then
                                If Buffer(int_Dmy) = vbCr Then
                                    str_Buffer &= " "
                                ElseIf Buffer(int_Dmy) = vbLf Then
                                    str_Buffer &= ""
                                Else
                                    str_Buffer &= Buffer(int_Dmy)
                                End If
                            End If
                        Next

                        ReadByteSize = .BytesToRead
                    Loop

                    If Not TimeOver = 0 Then
                        Return -1
                    End If
                End If
            Else
                Do Until ReadByteSize = 0
                    ReDim Buffer(ReadByteSize)
                    .Read(Buffer, 0, ReadByteSize)

                    For int_Dmy = 0 To Buffer.GetUpperBound(0)
                        'Application.DoEvents()

                        If Not Buffer(int_Dmy) = Nothing Then
                            If Buffer(int_Dmy) = vbCr Then
                                str_Buffer &= " "
                            ElseIf Buffer(int_Dmy) = vbLf Then
                                str_Buffer &= ""
                            Else
                                str_Buffer &= Buffer(int_Dmy)
                            End If
                        End If
                    Next

                    ReadByteSize = .BytesToRead
                Loop
            End If

            str_Buffer = str_Buffer.Trim
            str_Buffer = str_Buffer.Replace(",", "o")
            strMeasDat = str_Buffer

            Return str_Buffer.Length
        End With

    End Function

    Private Sub FTP_Init()

        With Me.FtpConnection1
            .Timeout = 3000
            .UserName = "Anonymous"
            .Password = "1234"
            Me.SetServerAddress()
        End With

    End Sub

    Private Sub SetServerAddress(Optional ByVal ServerNo As Integer = 0)

        With Me.FtpConnection1
            Dim PathChar As String = "//"
            Dim ServerAddress As String = Tp_OAI.ftpServer.ServerName.Substring(Tp_OAI.ftpServer.ServerName.IndexOf(PathChar) + PathChar.Length)
            .ServerAddress = ServerAddress

            Dim _RemoteDir As String = IIf(Tp_OAI.ftpServer.DefaultDir.EndsWith("/"), Tp_OAI.ftpServer.DefaultDir, Tp_OAI.ftpServer.DefaultDir & "/")
            .ServerDirectory = _RemoteDir
        End With

    End Sub

    Private Function DelAllFZ_ImgFile() As Integer

        Dim GetTheFile As String = String.Empty
        Dim lRetVal As Integer = -1


        If Not GetConnectedAdapters() < 0 Then
            Try
                If Not My.Computer.Network.Ping(Me.FtpConnection1.ServerAddress) Then
                    Return lRetVal
                End If
            Catch ex As Exception
                Return lRetVal
            End Try

            With Tp_OAI
                .IO.FZ_DI0.Trigger_OFF()
                .IO.FZ_DI1.Trigger_OFF()

                Try
                    Me.FtpConnection1.Connect()
                Catch ex As Exception
                    Return lRetVal
                End Try

                Dim AllFiles As String() = Me.FtpConnection1.GetFiles()

                For Each GetTheFile In AllFiles
                    If Not GetTheFile.ToLower.IndexOf(".ifz") < 0 Then
                        'Delete the downloaded files
                        Me.FtpConnection1.DeleteFile(GetTheFile)
                    End If
                Next

                Me.FtpConnection1.Close()
            End With
        End If

    End Function

    Private Function GetImage(Optional ByRef pFilePath As String = "") As Integer

        Dim strMeasData As String = String.Empty
        Dim GetTheFile As String = String.Empty
        Dim imgFilePath As String = Tp_OAI.SysTempPath

        Dim CurrentTime As Date = Now

        Dim lRetVal As Integer = -1
        Dim WaitTmr As Integer = 0


        If Not GetConnectedAdapters() < 0 Then
            Try
                If Not My.Computer.Network.Ping(Me.FtpConnection1.ServerAddress) Then
                    Return lRetVal
                End If
            Catch ex As Exception
                pFilePath = "-"
                Return lRetVal
            End Try

            With Tp_OAI
                If Not .Mode = SysAppMode.app_AutoRun Then
                    Try
                        Me.lbl_CaptureMsg.Visible = True
                    Catch ex As Exception

                    End Try
                End If

                .IO.FZ_DI0.Trigger_ON()
                DelayProc(30)

                FZ_SerCmd("MEASURE", strMeasData)

                If Not .Mode = SysAppMode.app_AutoRun Then
                    Try
                        Me.lbl_VisionData.Text = strMeasData
                    Catch ex As Exception

                    End Try
                End If

                With My.Computer.FileSystem
                    If Not .DirectoryExists(imgFilePath) Then .CreateDirectory(imgFilePath)
                End With

                .IO.FZ_DI0.Trigger_OFF()

                Try
                    Me.FtpConnection1.Connect()
                Catch ex As Exception
                    pFilePath = "-"
                    Return lRetVal
                End Try

                'Thread.Sleep(1000)
                Dim AllFiles() As String = Nothing
                WaitTmr = My.Computer.Clock.TickCount

                Do
                    Application.DoEvents()
                    AllFiles = Me.FtpConnection1.GetFiles()

                    If Not .Mode = SysAppMode.app_AutoRun Then
                        If Not .Mode = SysAppMode.app_Manu And Not .Mode = SysAppMode.app_Setting Then
                            .IO.FZ_DI0.Trigger_OFF()
                            Me.FtpConnection1.Close()
                            Return -1
                        End If
                    End If

                    If My.Computer.Clock.TickCount > WaitTmr + 5000 Then
                        .IO.FZ_DI0.Trigger_OFF()
                        Me.FtpConnection1.Close()
                        MessageBox.Show("Time Over - Communication Error... No image files were being retrieved.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Return -1
                    End If
                Loop While AllFiles.Length = 0


                For Each GetTheFile In AllFiles
                    Application.DoEvents()

                    If Not GetTheFile.ToLower.IndexOf(".ifz") < 0 Then
                        If Not My.Computer.FileSystem.FileExists(imgFilePath & "\" & GetTheFile) Then
                            'Download the files
                            Me.FtpConnection1.DownloadFile(imgFilePath & "\" & GetTheFile, GetTheFile)
                            WaitTmr = My.Computer.Clock.TickCount

                            Do While Not My.Computer.FileSystem.FileExists(imgFilePath & "\" & GetTheFile)
                                Application.DoEvents()

                                If Not .Mode = SysAppMode.app_AutoRun Then
                                    If Not .Mode = SysAppMode.app_Manu And Not .Mode = SysAppMode.app_Setting Then
                                        .IO.FZ_DI0.Trigger_OFF()
                                        Me.FtpConnection1.Close()
                                        Return -1
                                    End If
                                End If

                                If My.Computer.Clock.TickCount > WaitTmr + 5000 Then
                                    .IO.FZ_DI0.Trigger_OFF()
                                    Me.FtpConnection1.Close()
                                    MessageBox.Show("Time Over - Communication Error... No image files were being retrieved.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                                    Return -1
                                End If
                            Loop

                            'Delete the downloaded files
                            Me.FtpConnection1.DeleteFile(GetTheFile)
                        End If
                    End If
                Next

                ShellCmd(GetTheFile)
                pFilePath = imgFilePath & "\" & GetTheFile
                WaitTmr = My.Computer.Clock.TickCount

                Do While Not My.Computer.FileSystem.FileExists(pFilePath.Replace(".ifz", ".jpg"))
                    Application.DoEvents()

                    If Not .Mode = SysAppMode.app_AutoRun Then
                        If Not .Mode = SysAppMode.app_Manu And Not .Mode = SysAppMode.app_Setting Then
                            .IO.FZ_DI0.Trigger_OFF()
                            Me.FtpConnection1.Close()
                            Return -1
                        End If
                    End If

                    If My.Computer.Clock.TickCount > WaitTmr + 5000 Then
                        .IO.FZ_DI0.Trigger_OFF()
                        Me.FtpConnection1.Close()
                        MessageBox.Show("Time Over - Communication Error... No image files were being retrieved.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        Return -1
                    End If
                Loop

                Try
                    Me.FtpConnection1.Close()
                Catch ex As Exception

                End Try

                Me.PictureBox9.BackgroundImage = Image.FromFile(pFilePath.Replace(".ifz", ".jpg"))
                Me.PictureBox11.BackgroundImage = Image.FromFile(pFilePath.Replace(".ifz", ".jpg"))

                If .Mode = SysAppMode.app_AutoRun Then
                    Me.PictureBox4.BackgroundImage = Image.FromFile(pFilePath.Replace(".ifz", ".jpg"))
                End If

                Try
                    Me.lbl_CaptureMsg.Visible = False

                Catch ex As Exception

                End Try

                lRetVal = 0
                Return lRetVal
            End With
        Else
            pFilePath = "-"
            Return lRetVal
        End If

    End Function

    Private Sub DelayProc(Optional ByVal tm As Integer = 1)

        Dim CurTmr As Integer = My.Computer.Clock.TickCount


        Do Until My.Computer.Clock.TickCount > CurTmr + tm
            Application.DoEvents()
        Loop

    End Sub

    Private Sub frm_Main_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        With Me
            If .fg_BootError = 1 Then
                With .lbl_Msg
                    .Text = "Unable to initialize Comm. Port. Please reload the application by closing the application now!"
                End With

                If MessageBox.Show("System initialization Error. Please try again by re-load the application. Do you wish to re-start now?", "Taping OAI...", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                    Application.Restart()
                    Exit Sub
                End If
            ElseIf .fg_BootError = 0 Then
                Dim SceneNo As Integer = 0
                Dim RetStr As String = String.Empty

                If FZ_SerCmd("SCENE " & String.Format("{0:D2}", SceneNo), RetStr) < 0 Then
                    If MessageBox.Show("Unable to establish communication with Vision System. Please try again by re-load the application. Do you wish to re-start now?", "Taping OAI...", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                        Application.Restart()
                        Exit Sub
                    End If
                Else
                    If RetStr.Trim.ToUpper <> "OK" Then
                        If MessageBox.Show("Communication with Vision System returns error. Please try again by re-load the application. Do you wish to re-start now?", "Taping OAI...", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                            Application.Restart()
                            Exit Sub
                        End If
                    Else
                        .fg_BootError = 2

                        If Not Tp_OAI.IO.FZ_Run.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                            FZ_SerCmd("SCRSWITCH")
                        End If
                    End If
                End If
            End If
        End With

    End Sub

    Private Sub pic_Camera__MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pic_Camera_.MouseDown, pic_VisionChk.MouseDown

        With Tp_OAI
            If Not .Mode = SysAppMode.app_Manu And Not .Mode = SysAppMode.app_Setting Then Exit Sub

            If e.Button = Windows.Forms.MouseButtons.Left Then
                If Not Tp_OAI.IO.FZ_Run.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                    If MessageBox.Show("the Vision System is currently not in 'Run' Mode. Do you wish to switch it to 'Run' Mode now?", "Taping-OAI...", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                        FZ_SerCmd("SCRSWITCH")
                    End If
                Else
                    pic_Camera.BackColor = RedLED_OnOff(1)
                    pic_Camera_.BackColor = RedLED_OnOff(1)
                    pic_VisionChk.BackColor = RedLED_OnOff(1)
                    Dim MeasData As String = String.Empty

                    If Me.chk_Vision.CheckState = CheckState.Checked Or sender.Equals(Me.pic_VisionChk) Then
                        If Me.lbl_CaptureMsg.Visible = True Then
                            pic_Camera.BackColor = RedLED_OnOff(0)
                            pic_Camera_.BackColor = RedLED_OnOff(0)
                            pic_VisionChk.BackColor = RedLED_OnOff(0)
                            Exit Sub
                        End If

                        If GetImage() < 0 Then
                            .IO.FZ_DI0.Trigger_OFF()

                            FZ_SerCmd("MEASURE", MeasData)
                            Me.lbl_VisionData.Text = MeasData

                            MessageBox.Show("Unabled to capture image. Please confirm network setting.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        End If
                    Else
                        .IO.FZ_DI0.Trigger_OFF()

                        If Me.chk_TapeSeal.Checked = True Then
                            .IO.FZ_DI1.Trigger_ON()
                        Else
                            .IO.FZ_DI1.Trigger_OFF()
                        End If

                        FZ_SerCmd("MEASURE", MeasData)

                        .IO.FZ_DI0.Trigger_OFF()
                        .IO.FZ_DI1.Trigger_OFF()
                        Me.lbl_VisionData.Text = MeasData
                    End If

                    pic_Camera.BackColor = RedLED_OnOff(0)
                    pic_Camera_.BackColor = RedLED_OnOff(0)
                    pic_VisionChk.BackColor = RedLED_OnOff(0)
                End If
            End If
        End With

    End Sub

    Private Sub btn_StepLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_StepLeft.Click

        Me.mn_StepperMove_n()

    End Sub

    Private Sub btn_StepRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_StepRight.Click

        Me.mn_StepperMove_n(1)

    End Sub

    Private Sub StartMainThread()

        Dim tsk_AutoRun As Thread = New Thread(AddressOf AutoRunningSeq)


        With tsk_AutoRun
            .Name = "tsk_AutoRun"
            .Start()
        End With

    End Sub

    Public Sub DispControlValue(ByVal FormControl As Label, ByVal ThreadName As String, ByVal Text As String)

        FormControl.Text = Text

        With Me
            Dim Data() As String = Text.Split(" "c)

            If Data.GetUpperBound(0) = 4 And Data(Data.GetUpperBound(0)) = "9999" Then
                Data(2) = Data(2) & " " & Data(3)
                Data(3) = "9999"
                ReDim Preserve Data(3)
            ElseIf Data.GetUpperBound(0) = 2 And Data(Data.GetUpperBound(0)) = "9999" Then
                If Tp_OAI.SysData(0).Prod.ToUpper = "RAKON" Then
                    ReDim Preserve Data(3)
                    Data(3) = "9999"
                    Data(2) = ""
                End If
            ElseIf Data.GetUpperBound(0) = 3 And Data(Data.GetUpperBound(0)) = "9999" Then
                If Data(1) = "" And Data(2) = "" Then
                    Exit Sub
                Else
                    If Data(2) = "" Then Data(2) = Data(1)
                    If Data(1) = "" Then Data(1) = Data(2)
                End If
            ElseIf Data.GetUpperBound(0) < 2 And Data(Data.GetUpperBound(0)) = "9999" Then
                Exit Sub
            End If


            .lst_RawData.Items.Add(Text)
            .lbl_VZData.Text = Data(1) & "  " & Data(2)
            .lbl_VisionData.Text = Text

            If .lbl_InspTapeType.Visible = False Then
                Dim Chk1 As Integer = Array.IndexOf(Tp_OAI.SysData(0).MarkData1, Data(1))
                Dim Chk2 As Integer = Array.IndexOf(Tp_OAI.SysData(0).MarkData2, Data(2))

                If Chk1 < 0 Then
                    .ErrorProvider5.SetError(.lbl_MarkChar1, "Data not match!")
                Else
                    .ErrorProvider5.Clear()
                End If

                If Chk2 < 0 Then
                    .ErrorProvider6.SetError(.lbl_MarkChar2, "Data not match!")
                Else
                    .ErrorProvider6.Clear()
                End If
            End If
        End With

    End Sub

    Public Sub DispLabel_(ByVal FormControl As Label, ByVal ThreadName As String, ByVal Text As String)

        If Text = "True" Then
            FormControl.Visible = True
        Else
            FormControl.Visible = False
        End If

    End Sub

    Public Sub DispControl_(ByVal FormControl As PictureBox, ByVal ThreadName As String, ByVal Text As String)

        If Text = "True" Then
            FormControl.Visible = True
        Else
            FormControl.Visible = False
        End If

    End Sub

    Public Sub DispLabel(ByVal FormControl As Label, ByVal ThreadName As String, ByVal Text As String)

        Try
            FormControl.Text = Text
            FormControl.Visible = True
        Catch ex As Exception

        End Try

    End Sub

    Public Sub DispToolTipLabel(ByVal FormControl As Label, ByVal ThreadName As String, ByVal Text As String)

        Try
            With Me.ToolTip1
                .SetToolTip(FormControl, Text)
            End With
        Catch ex As Exception

        End Try

    End Sub

    Public Sub DispControl(ByVal FormControl As Button, ByVal ThreadName As String, ByVal Text As String)

        FormControl.Text = Text

    End Sub

    Private Sub UpdateCountingData()

        With Tp_OAI
            Dim CalPercentage As String = String.Empty
            Dim Total As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

            'Display OK Quantity
            If Me.lbl_TotalOK.InvokeRequired Then
                Me.lbl_TotalOK.Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_TotalOK, Thread.CurrentThread.Name, .SysData(0).InspCnt.ToString})
            Else
                Me.lbl_TotalOK.Text = .SysData(0).InspCnt.ToString
            End If

            If Total > 0 Then
                CalPercentage = String.Format("{0:F2}%", (.SysData(0).InspCnt / Total) * 100)

                If Me.lbl_TotalOK_.InvokeRequired Then
                    Me.lbl_TotalOK_.Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_TotalOK_, Thread.CurrentThread.Name, CalPercentage})
                Else
                    Me.lbl_TotalOK_.Text = CalPercentage
                End If

                'Display NG Quantity
                If Me.lbl_TotalNG.InvokeRequired Then
                    Me.lbl_TotalNG.Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_TotalNG, Thread.CurrentThread.Name, .SysData(0).InspCnt_NG.ToString})
                Else
                    Me.lbl_TotalNG.Text = .SysData(0).InspCnt_NG.ToString
                End If

                CalPercentage = String.Format("{0:F2}%", (.SysData(0).InspCnt_NG / Total) * 100)

                If Me.lbl_TotalNG_.InvokeRequired Then
                    Me.lbl_TotalNG_.Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_TotalNG_, Thread.CurrentThread.Name, CalPercentage})
                Else
                    Me.lbl_TotalNG_.Text = CalPercentage
                End If

                'Display NG Quantity
                If Me.lbl_WeekCodeJump.InvokeRequired Then
                    Me.lbl_WeekCodeJump.Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_WeekCodeJump, Thread.CurrentThread.Name, .SysData(0).WeekCodeJump.ToString})
                Else
                    Me.lbl_WeekCodeJump.Text = .SysData(0).InspCnt_NG.ToString
                End If

                CalPercentage = String.Format("{0:F2}%", (.SysData(0).WeekCodeJump / Total) * 100)

                If Me.lbl_WeekCodeJump_.InvokeRequired Then
                    Me.lbl_WeekCodeJump_.Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_WeekCodeJump_, Thread.CurrentThread.Name, CalPercentage})
                Else
                    Me.lbl_WeekCodeJump_.Text = CalPercentage
                End If

                'Display Total Quantity
                If Me.lbl_Total.InvokeRequired Then
                    Me.lbl_Total.Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_Total, Thread.CurrentThread.Name, Total.ToString & "/" & .SysData(0).Acceptance})
                Else
                    Me.lbl_Total.Text = Total.ToString
                End If
            End If

            'Display Insp. Error Counter
            If Me.lbl_Judge.InvokeRequired Then
                Me.lbl_Judge.Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_Judge, Thread.CurrentThread.Name, .SysData(0).BuzzCnt & " Err."})
            Else
                Me.lbl_Judge.Text = Total.ToString
            End If
        End With

    End Sub

    Private Sub InspDone()

        With Tp_OAI
            'Saving Data
            Dim InspDate As String = String.Format("{0:D2}-{1:D2}-{2:D4}", .SysData(0).InspDate.Day, .SysData(0).InspDate.Month, .SysData(0).InspDate.Year)
            Dim PathName As String = .LotDataPath & "\" & .SysData(0).Prod & "\" & InspDate & "\" & .SysData(0).LotNo
            Dim RedoRecFile As String = .RedoLotTmpRecLoc & "\" & .SysData(0).LotNo & ".dat"
            Dim LotDataFile As String = String.Empty


            If Not Val(.SysData(0).InspType) = 0 Then
                LotDataFile = PathName & "\" & "\SealInsp" & ".dat"
            Else
                LotDataFile = PathName & "\" & "\MarkInsp" & ".dat"
            End If


            Dim LotEndSummarize As String = String.Empty
            Dim RedoSummarize As String = String.Empty

            With .SysData(0)
                LotEndSummarize = "Lot No." & vbTab & vbTab & vbTab & " : " & .LotNo & vbCrLf
                LotEndSummarize &= "P_Lot No." & vbTab & vbTab & " : " & .P_LotNo & vbCrLf
                LotEndSummarize &= "Product" & vbTab & vbTab & vbTab & " : " & .Prod & vbCrLf
                LotEndSummarize &= "Quantity" & vbTab & vbTab & " : " & .Acceptance & vbCrLf
                LotEndSummarize &= "Operator" & vbTab & vbTab & " : " & .Insp & vbCrLf
                LotEndSummarize &= "Date" & vbTab & vbTab & vbTab & " : " & Me.lbl_Date.Text & vbCrLf
                LotEndSummarize &= "Duration" & vbTab & vbTab & " : " & Me.lbl_TimeEllapse.Text & vbCrLf
                LotEndSummarize &= vbCrLf
                LotEndSummarize &= "Total OK" & vbTab & vbTab & " : " & Me.lbl_TotalOK.Text & "   (" & Me.lbl_TotalOK_.Text & ")" & vbCrLf
                LotEndSummarize &= "Total NG" & vbTab & vbTab & " : " & Me.lbl_TotalNG.Text & "   (" & Me.lbl_TotalNG_.Text & ")" & vbCrLf
                LotEndSummarize &= "Total" & vbTab & vbTab & vbTab & " : " & Me.lbl_Total.Text & vbCrLf
                LotEndSummarize &= "Error Cnt." & vbTab & vbTab & " : " & Me.lbl_Judge.Text & vbCrLf


                If Val(Me.lbl_TotalNG.Text) <> 0 Then
                    RedoSummarize = .GUID & vbCrLf
                    RedoSummarize &= .LotNo & vbCrLf
                    RedoSummarize &= .P_LotNo & vbCrLf
                    RedoSummarize &= .Prod & vbCrLf
                    RedoSummarize &= .Acceptance & vbCrLf

                    Dim str As String = String.Empty

                    For iLp As Integer = 0 To .MarkData1.GetUpperBound(0)
                        Application.DoEvents()

                        If str <> "" Then
                            If str.IndexOf(.MarkData1(iLp)) < 0 Then
                                str &= .MarkData1(iLp) & ","
                            End If
                        Else
                            str &= .MarkData1(iLp) & ","
                        End If
                    Next

                    RedoSummarize &= str & vbCrLf
                    str = ""

                    For iLp As Integer = 0 To .MarkData2.GetUpperBound(0)
                        Application.DoEvents()

                        If str <> "" Then
                            If str.IndexOf(.MarkData2(iLp)) < 0 Then
                                str &= .MarkData2(iLp) & ","
                            End If
                        Else
                            str &= .MarkData2(iLp) & ","
                        End If
                    Next

                    RedoSummarize &= str & vbCrLf
                    RedoSummarize &= .InspType & vbCrLf
                    RedoSummarize &= Tp_OAI.RedoLoc_ & vbCrLf
                    RedoSummarize &= .Insp & vbCrLf
                    RedoSummarize &= Me.lbl_Judge.Text & vbCrLf
                    RedoSummarize &= Me.lbl_Date.Text & vbCrLf
                    RedoSummarize &= Me.lbl_TotalOK.Text & vbCrLf
                    RedoSummarize &= Me.lbl_TotalNG.Text & vbCrLf

                    'My.Computer.FileSystem.WriteAllText(RedoRecFile, RedoSummarize & vbCrLf, True, System.Text.Encoding.ASCII)
                Else
                    If My.Computer.FileSystem.FileExists(RedoRecFile) Then
                        My.Computer.FileSystem.DeleteFile(RedoRecFile)
                    End If
                End If
            End With

            My.Computer.FileSystem.WriteAllText(LotDataFile, LotEndSummarize & vbCrLf, False, System.Text.Encoding.ASCII)
        End With

        SysDataReset()
        DeleteSysTempFile()

    End Sub

    Private Sub ClearDisp()

        With Me
            .lbl_TimeEllapse.Text = "0.0 min Elapse"
            .lbl_CycTime.Text = "0.0 sec/pcs"
            .lbl_LotNo.Text = ""
            .lbl_PLotNo.Text = ""
            .lbl_EmpNo.Text = ""
            .lbl_Date.Text = ""
            .lbl_VZData.Text = ""
            .lbl_TotalOK.Text = ""
            .lbl_TotalOK_.Text = ""
            .lbl_TotalNG.Text = ""
            .lbl_TotalNG_.Text = ""
            .lbl_WeekCodeJump.Text = ""
            .lbl_WeekCodeJump_.Text = ""

            .lbl_Total.Text = ""
            .lbl_Judge.Text = "---"

            With .lst_RawData
                .Items.Clear()
                .Refresh()
            End With
        End With

    End Sub

    Private Sub AutoRunningSeq()

        Dim MotorTurn As Integer = 0
        Dim FmtRetry As Integer = 0
        Dim ReTry As Integer = 0

        Dim WkJmpReTry As Integer = 0
        Dim EmptyPocket As Integer = 0
        Dim PullerTmr As Integer = 0
        Dim CycTime As Integer = 0
        Dim SqCnt As Integer = 0
        Dim Mtr_Tmr As Integer = 0

        Dim AutoUpdate As Integer = 1
        Dim ChkUpdateTempOffTmr As Integer = 0

        Dim fg_ChkWkJmp As Integer = 0              ' Disable checking week code jumping



        With Me
            Do Until .fg_UnloadMe = 1
                Application.DoEvents()

                With Tp_OAI
                    If .Mode = SysAppMode.app_AutoRun Then
                        Select Case .AutoSeqNo
                            Case Is = 0
                                .WC_Error = False
                                CycTime = My.Computer.Clock.TickCount

                                pic_Camera.BackColor = RedLED_OnOff(1)
                                pic_Camera_.BackColor = RedLED_OnOff(1)
                                pic_VisionChk.BackColor = RedLED_OnOff(1)
                                Dim MeasData As String = String.Empty

                                .IO.FZ_DI0.Trigger_OFF()

                                'Check Tape Seal ???
                                If Me.lbl_InspTapeType.Visible = True Then
                                    .IO.FZ_DI1.Trigger_ON()
                                End If

                                FZ_SerCmd("MEASURE", MeasData)
                                Dim MeasData_ As String = MeasData
                                .IO.FZ_DI1.Trigger_OFF()

                                pic_Camera.BackColor = RedLED_OnOff(0)
                                pic_Camera_.BackColor = RedLED_OnOff(0)
                                pic_VisionChk.BackColor = RedLED_OnOff(0)


                                If Me.lbl_VZData.InvokeRequired Then
                                    Me.lbl_VZData.Invoke(New frm_Main.DispMsg(AddressOf Me.DispToolTipLabel), New Object() {Me.lbl_VZData, Thread.CurrentThread.Name, MeasData.Trim})
                                End If


                                '--- Tape Seal Inspection Judgment ---
                                'Respond from Vision -> 'OK 72.28,46.79 6666
                                If Not Val(.SysData(0).InspType) = 0 Then
                                    'Saving Data
                                    Dim InspDate As String = String.Format("{0:D2}-{1:D2}-{2:D4}", .SysData(0).InspDate.Day, .SysData(0).InspDate.Month, .SysData(0).InspDate.Year)
                                    Dim PathName As String = .LotDataPath & "\" & .SysData(0).Prod & "\" & InspDate & "\" & .SysData(0).LotNo

                                    If Not Val(.SysData(0).InspType) = 0 Then
                                        PathName &= "\SealInsp\RawData"
                                    Else
                                        PathName &= "\MarkInsp\RawData"
                                    End If

                                    If My.Computer.FileSystem.DirectoryExists(PathName) = False Then
                                        My.Computer.FileSystem.CreateDirectory(PathName)
                                    End If

                                    Dim RawDataFile As String = PathName & "\" & .SysData(0).LotNo & ".dat"
                                    My.Computer.FileSystem.WriteAllText(RawDataFile, MeasData_ & vbCrLf, True, System.Text.Encoding.ASCII)


                                    If Not MeasData.IndexOf("6666") < 0 Then
                                        MeasData = MeasData.Replace(","c, " "c)
                                        Dim Seal_() As String = MeasData.Split(" "c)
                                        Dim SealDim_ As System.Collections.Generic.IEnumerable(Of String) = Seal_.Where(Function(q) q.Length > 0 And q <> "6666" And q <> "OK")

                                        'Check Range Here
                                        If Val(SealDim_(0)) < .InspTapeSeal.A_min Or Val(SealDim_(0)) > .InspTapeSeal.A_max Or Val(SealDim_(1)) < .InspTapeSeal.B_min Or Val(SealDim_(1)) > .InspTapeSeal.B_max Then
                                            If Not ReTry < 2 Then
                                                ReTry = 0
                                                .RT_Error = 90
                                                .AutoSeqNo = 10
                                                Exit Select
                                            Else
                                                ReTry += 1
                                                mn_StepperMove(-8)
                                                Thread.Sleep(180)
                                                mn_StepperMove(8)

                                                Thread.Sleep(80)
                                                .AutoSeqNo = 0
                                                Exit Select
                                            End If
                                        End If

                                        'Searching First Xtal
                                        Dim Total As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

                                        If Total = 0 Then
                                            FZ_SerCmd("MEASURE", MeasData)

                                            If Not MeasData.IndexOf("8888") < 0 Then
                                                'Empty pocket
                                                .AutoSeqNo += 1
                                                Exit Select
                                            Else
                                                Dim Data_() As String = MeasData.Split(" "c)
                                                Dim Is_rgbData_ As System.Collections.Generic.IEnumerable(Of String) = Data_.Where(Function(q) q.Length > 0 And q <> "9999" And q <> "OK")

                                                If Is_rgbData_.Count < 2 Then
                                                    .AutoSeqNo += 1
                                                    Exit Select
                                                Else
                                                    'First Xtal found here
                                                    .SysData(0).InspCnt = 1
                                                    SqCnt = 3
                                                    UpdateCountingData()
                                                    .AutoSeqNo += 1

                                                    Exit Select
                                                End If
                                            End If
                                        Else
                                            SqCnt += 1
                                            .SysData(0).InspCnt = SqCnt \ 3
                                            UpdateCountingData()

                                            Total = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

                                            If Total >= Val(.SysData(0).Acceptance) + 25 Then
                                                'Finished Inspection (included 25 empty pocket at the end)
                                                SetAlarm(99)
                                                .AutoSeqNo = 0

                                                InspDone()
                                                Exit Select
                                            Else
                                                'Tape Seal Inspection Continue
                                                .AutoSeqNo += 1
                                                Exit Select
                                            End If
                                        End If
                                    End If
                                End If              '--- Tape Seal Inspection Judgment Procedure End ---


                                '--- Marking Inspection ---
                                Dim Data() As String = MeasData.Split(" "c)
                                Dim ChkDataFmt As Integer = -1

                                If Data.GetUpperBound(0) = 4 And Data(Data.GetUpperBound(0)) = "9999" Then
                                    If Not .SysData(0).Prod.ToUpper.IndexOf("3225") < 0 Then
                                        Data(2) = Data(2) & " " & Data(3)
                                        Data(3) = "9999"
                                        ReDim Preserve Data(3)
                                        ChkDataFmt = 0
                                    End If
                                ElseIf Data.GetUpperBound(0) = 2 And Data(Data.GetUpperBound(0)) = "9999" Then
                                    If .SysData(0).Prod.ToUpper = "RAKON" Then
                                        ReDim Preserve Data(3)
                                        Data(3) = "9999"
                                        Data(2) = Data(1)
                                        ChkDataFmt = 0
                                    End If
                                ElseIf Data.GetUpperBound(0) = 3 And Data(Data.GetUpperBound(0)) = "9999" Then
                                    If Data(1).Trim.Length < .SysData(0).MarkData1(0).Length Or Data(2).Trim = "" Then
                                        ChkDataFmt = -1
                                    Else
                                        If Data(2) = "" Then Data(2) = Data(1)
                                        If Data(1) = "" Then Data(1) = Data(2)
                                        ChkDataFmt = 0
                                    End If
                                ElseIf Not Array.IndexOf(Data, "8888") < 0 Then
                                    ChkDataFmt = 0
                                End If


                                'In correct Vision Data Format
                                If ChkDataFmt < 0 Then
                                    If Me.lbl_VZData.InvokeRequired Then
                                        Me.lbl_VZData.Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_VZData, Thread.CurrentThread.Name, MeasData.Trim})
                                    Else
                                        Me.lbl_VZData.Text = MeasData.Trim
                                    End If

                                    Dim Skip As Boolean = False
                                    FmtRetry += 1

                                    If FmtRetry > 7 Then
                                        Dim RetVal As String = String.Empty

                                        Me.FZ_SerCmd("UNITDATA 15 125", RetVal)
                                        Dim FilterSize() As String = RetVal.Split(" "c)

                                        If Not Val(FilterSize(0)) = Val(.BGSmaxSet) Then
                                            Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)

                                            'Set min BGS Level
                                            Me.FZ_SerCmd("UNITDATA 15 124 " & .BGSminSet)

                                            'Set max BGS Level
                                            Me.FZ_SerCmd("UNITDATA 15 125 " & .BGSmaxSet)

                                            .AutoSeqNo = 0
                                            Exit Select
                                        Else
                                            Dim ChkTotalCnt As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump
                                            FmtRetry = 0
                                            ReTry = 0

                                            If ChkTotalCnt < Val(.SysData(0).Acceptance) Then
                                                .RT_Error = 90
                                                .AutoSeqNo = 10
                                                Exit Select
                                            Else
                                                MeasData = "OK 8888"
                                                MeasData_ = "OK 8888"
                                                Skip = True
                                            End If
                                        End If
                                    Else
                                        If Data.GetUpperBound(0) < 2 OrElse (Data(1) <> "" Or Data(2) <> "") Then
                                            Dim RetVal As String = String.Empty
                                            Dim BGSmin() As String = {}
                                            Dim BGSmax() As String = {}


                                            Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)
                                            Dim FilterSize() As String = RetVal.Split(" "c)

                                            If Val(FilterSize(0)) = 0 Then
                                                Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)
                                            Else
                                                Me.FZ_SerCmd("UNITDATA 18 10", RetVal)

                                                If Not RetVal.IndexOf("OK") < 0 Then
                                                    Dim DensityData() As String = RetVal.Split(" "c)

                                                    If Val(DensityData(0)) > IIf(Val(FilterSize(0)) = 0, Val(.DensityLvl) + 15, Val(.DensityLvl)) AndAlso Val(DensityData(0)) < Val(.DensityLvl) + 50 Then
                                                        Me.FZ_SerCmd("UNITDATA 15 124", RetVal)
                                                        BGSmin = RetVal.Split(" "c)
                                                        Me.FZ_SerCmd("UNITDATA 15 124 " & Val(BGSmin(0) + 5).ToString, RetVal)

                                                        Me.FZ_SerCmd("UNITDATA 15 125", RetVal)
                                                        BGSmax = RetVal.Split(" "c)
                                                        Me.FZ_SerCmd("UNITDATA 15 125 " & Val(BGSmax(0) + 5).ToString, RetVal)
                                                    End If


                                                    If Val(DensityData(0)) < Val(.DensityLvl) - 10 AndAlso Val(DensityData(0)) > Val(.DensityLvl) - 60 Then
                                                        Me.FZ_SerCmd("UNITDATA 15 124", RetVal)
                                                        BGSmin = RetVal.Split(" "c)
                                                        Me.FZ_SerCmd("UNITDATA 15 124 " & Val(BGSmin(0) - 5).ToString, RetVal)


                                                        Me.FZ_SerCmd("UNITDATA 15 125", RetVal)
                                                        BGSmax = RetVal.Split(" "c)
                                                        Me.FZ_SerCmd("UNITDATA 15 125 " & Val(BGSmax(0) - 5).ToString, RetVal)
                                                    End If
                                                End If
                                            End If
                                        End If


                                        mn_StepperMove(-8)
                                        Thread.Sleep(180)
                                        mn_StepperMove(8)

                                        Thread.Sleep(80)
                                        .AutoSeqNo = 0
                                        Exit Select
                                    End If

                                    If Skip = False Then Exit Select
                                End If

                                FmtRetry = 0
                                Dim Is_rgbData As System.Collections.Generic.IEnumerable(Of String) = Data.Where(Function(q) q.Length >= 1 And q <> "9999" And q <> "OK")

                                If MeasData.IndexOf("8888") < 0 Then
                                    If Not EmptyPocket = 0 Then
                                        ReTry = 0
                                        .RT_Error = 71
                                        .AutoSeqNo = 10
                                        Exit Select
                                    End If

                                    EmptyPocket = 0

                                    If Me.lbl_VZData.InvokeRequired Then
                                        If Is_rgbData.Count = 2 Then
                                            Me.lbl_VZData.Invoke(New frm_Main.UpdateControl(AddressOf Me.DispControlValue), New Object() {Me.lbl_VZData, Thread.CurrentThread.Name, MeasData})
                                        End If
                                    Else
                                        Me.lbl_VZData.Text = MeasData
                                        Me.lbl_VisionData.Text = MeasData
                                    End If
                                Else
                                    Dim Total As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

                                    If .SysData(0).InspCnt <> 0 And Total < Val(.SysData(0).Acceptance) Then
                                        ReTry = 0
                                        .RT_Error = 71
                                        .AutoSeqNo = 10
                                        Exit Select
                                    Else
                                        If Total = 0 Then
                                            .AutoSeqNo = 1
                                            Exit Select
                                        End If

                                        If Total >= Val(.SysData(0).Acceptance) Then
                                            If EmptyPocket >= 25 Then
                                                EmptyPocket = 0
                                                PullerTmr = 0
                                                MotorTurn = 0

                                                SetAlarm(99)
                                                .AutoSeqNo = 0

                                                Dim RunDuration As Integer = My.Computer.Clock.TickCount - .TimeEllapse
                                                Dim TotalTimeEllapse As String = String.Format("{0:F0} min {1:F0} sec Elapse", RunDuration \ 60000, ((RunDuration Mod 60000) / 1000))

                                                If Me.lbl_TimeEllapse.InvokeRequired Then
                                                    Me.lbl_TimeEllapse.Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_TimeEllapse, Thread.CurrentThread.Name, TotalTimeEllapse})
                                                Else
                                                    Me.lbl_TimeEllapse.Text = TotalTimeEllapse
                                                End If

                                                InspDone()
                                                Exit Select
                                            Else
                                                EmptyPocket += 1
                                                .AutoSeqNo = 1
                                                Exit Select
                                            End If
                                        End If
                                    End If
                                End If


                                '--- Inspect Vision Data ---
                                Dim RedoLocSeek As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump + 1
                                .RedoFlag = False


                                '--- Implement Fast Inspection during ReDo ---
                                If .SysData(0).ReDoLoc <> "" Then
                                    Dim NG_Pos() As String = .SysData(0).ReDoLoc.Split(",")

                                    If Not Array.IndexOf(NG_Pos, RedoLocSeek.ToString.Trim) < 0 Then
                                        .RedoFlag = True
                                    End If
                                Else
                                    .RedoFlag = True
                                End If

                                .RedoFlag = True    'Bypass Fast Inspection

                                If Is_rgbData.Count = 2 And .RedoFlag = True Then
                                    If Me.lbl_InspTapeType.Visible = False Then
                                        Dim Chk1 As Integer = -1
                                        Dim Chk2 As Integer = -1

                                        'Inspection Vision data
                                        If Not .SysData(0).MarkData1(0) = "" Then
                                            Chk1 = Array.IndexOf(.SysData(0).MarkData1, Data(1))
                                        Else
                                            Chk1 = 1
                                        End If

                                        'Inspection Vision data
                                        'Chk2 = Array.IndexOf(.SysData(0).MarkData2, Data(2))

                                        'No Check Week Code
                                        If .SysData(0).MarkData2(0).Substring(0, 1) = Data(2).Substring(0, 1) AndAlso .SysData(0).MarkData2(0).Substring(.SysData(0).MarkData2(0).Length - 1) = Data(2).Substring(Data(2).Length - 1) Then
                                            Chk2 = 1
                                        End If


                                        '
                                        'Recover Mis-judge Data
                                        If Chk1 < 0 And Data(1).Length = 5 Then
                                            Dim RcvData As String = String.Empty

                                            Select Case .SysData(0).Prod
                                                Case Is = "FA-20HDOT"
                                                    RcvData = Data(1).Substring(0, 3)
                                                    RcvData = RcvData.Replace("A", "4")
                                                    RcvData = RcvData.Replace("B", "8")
                                                    RcvData = RcvData.Replace("C", "0")
                                                    RcvData = RcvData.Replace("D", "0")
                                                    RcvData = RcvData.Replace("E", "1")
                                                    RcvData = RcvData.Replace("F", "1")
                                                    RcvData = RcvData.Replace("H", "8")
                                                    RcvData = RcvData.Replace("L", "1")
                                                    RcvData = RcvData.Replace("O", "0")
                                                    RcvData = RcvData.Replace("R", "3")
                                                    RcvData = RcvData.Replace("S", "5")
                                                    RcvData = RcvData.Replace("T", "1")
                                                    RcvData = RcvData.Replace("U", "0")

                                                    Data(1) = RcvData & Data(1).Substring(3, 2)
                                                Case Else
                                                    RcvData = Data(1).Substring(0, 4)
                                                    RcvData = RcvData.Replace("A", "4")
                                                    RcvData = RcvData.Replace("B", "8")
                                                    RcvData = RcvData.Replace("C", "0")
                                                    RcvData = RcvData.Replace("D", "0")
                                                    RcvData = RcvData.Replace("E", "1")
                                                    RcvData = RcvData.Replace("F", "1")
                                                    RcvData = RcvData.Replace("H", "8")
                                                    RcvData = RcvData.Replace("L", "1")
                                                    RcvData = RcvData.Replace("O", "0")
                                                    RcvData = RcvData.Replace("R", "0")
                                                    RcvData = RcvData.Replace("S", "5")
                                                    RcvData = RcvData.Replace("T", "1")
                                                    RcvData = RcvData.Replace("U", "0")

                                                    Data(1) = RcvData & Data(1).Substring(4, 1)
                                            End Select

                                            If Data(1).EndsWith("1") Or Data(1).EndsWith("F") Then
                                                Data(1) = Data(1).Substring(0, 4) & "P"
                                            End If

                                            If Data(1).EndsWith("0") Or Data(1).EndsWith("E") Then
                                                Data(1) = Data(1).Substring(0, 4) & "C"
                                            End If

                                            If Data(1).StartsWith("7") And Data(1).Length = 5 And IsNumeric(Data(1).Substring(0, 3)) Then
                                                Data(1) = "2" & Data(1).Substring(1)
                                            End If

                                            Chk1 = Array.IndexOf(.SysData(0).MarkData1, Data(1))
                                        End If


                                        'Recover Vision Control Data
                                        If Chk1 < 0 Then
                                            If Data(1).Length > .SysData(0).MarkData1(0).Length Then
                                                For iLp As Integer = 0 To .SysData(0).MarkData1.GetUpperBound(0)
                                                    Application.DoEvents()

                                                    If Not Data(1).IndexOf(.SysData(0).MarkData1(iLp)) < 0 Then
                                                        Chk1 = iLp
                                                        Exit For
                                                    End If
                                                Next
                                            End If
                                        End If

                                        If Chk2 < 0 Then
                                            If .SysData(0).MarkData2(0).Length = 6 And .SysData(0).MarkData2(0).StartsWith("o") Then
                                                If Data(2).Length > 1 AndAlso Data(2).Substring(1, 1) = "E" And (Data(2).StartsWith("6") Or Data(2).StartsWith("9") Or Data(2).StartsWith("P")) Then
                                                    Data(2) = "o" & Data(2).Substring(1)
                                                    Chk2 = Array.IndexOf(.SysData(0).MarkData2, Data(2))
                                                End If
                                            End If

                                            If Data(2).Length > .SysData(0).MarkData2(0).Length Then
                                                For iLp As Integer = 0 To .SysData(0).MarkData2.GetUpperBound(0)
                                                    Application.DoEvents()

                                                    If Not Data(2).IndexOf(.SysData(0).MarkData2(iLp)) < 0 Then
                                                        Chk2 = iLp
                                                        Exit For
                                                    End If
                                                Next
                                            End If

                                            If Chk2 < 0 Then
                                                For iLp = 0 To .SysData(0).MarkData2.GetUpperBound(0)
                                                    Application.DoEvents()
                                                    If Not Chk2 < 0 Then Exit For

                                                    If Data(2).Length = .SysData(0).MarkData2(iLp).Length Then
                                                        For iLp_ As Integer = 0 To .SysData(0).MarkData2(iLp).Length - 1
                                                            Application.DoEvents()

                                                            Try
                                                                If Not Data(2).Substring(iLp_, 1) = .SysData(0).MarkData2(iLp).Substring(iLp_, 1) Then
                                                                    Select Case .SysData(0).Prod
                                                                        Case Is = "FA-20HDOT"
                                                                            Select Case iLp_
                                                                                Case Is = 1
                                                                                    If Data(2).Substring(iLp_, 1) = "1"c And .SysData(0).MarkData2(iLp).Substring(iLp_, 1) = "E"c Then
                                                                                        Data(2) = Data(2).Substring(iLp_ - 1, 1) & "E" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                                Case Is = 2 'Recover Year Data Value (1234567890)
                                                                                    If Not IsNumeric(Data(2).Substring(iLp_, 1)) Then
                                                                                        If Data(2).Substring(iLp_, 1) = "T" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        End If
                                                                                    End If
                                                                                Case Is = 3 'Recover Month Data Value (123456789XYZ)
                                                                                    If Data(2).Substring(iLp_, 1) = "O" Or Data(2).Substring(iLp_, 1) = "0" Then
                                                                                        'Data(2) = Data(2).Substring(0, iLp_) & "C" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "T" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                                Case Is = 4 'Recover Day Data Value (123456789XYZ)
                                                                                    If Data(2).Substring(iLp_, 1) = "o" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "P" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                            End Select

                                                                        Case Is = "FA-20H", "FA-128", "FA-23", "FA-23A", "FA-238", "FA-118T", "FA-206"
                                                                            Select Case iLp_
                                                                                Case Is = 0
                                                                                    If Data(2).Substring(iLp_, 1) = "1"c And .SysData(0).MarkData2(iLp).Substring(iLp_, 1) = "E"c Then
                                                                                        Data(2) = "E" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                                Case Is = 1 'Recover Year Data Value (1234567890)
                                                                                    If Not IsNumeric(Data(2).Substring(iLp_, 1)) Then
                                                                                        If Data(2).Substring(iLp_, 1) = "T" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        End If
                                                                                    End If
                                                                                Case Is = 2 'Recover Month Data Value (123456789XYZ)
                                                                                    If Data(2).Substring(iLp_, 1) = "O" Or Data(2).Substring(iLp_, 1) = "0" Then
                                                                                        'Data(2) = Data(2).Substring(0, iLp_) & "C" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "T" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "W" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                            End Select

                                                                        Case Is = "TSX-3225", "TD3225N"
                                                                            Select Case iLp_
                                                                                Case Is = 0
                                                                                    If Data(2).Substring(iLp_, 1) = "1"c And .SysData(0).MarkData2(iLp).Substring(iLp_, 1) = "E"c Then
                                                                                        Data(2) = "E" & Data(2).Substring(iLp_ + 1)
                                                                                    End If

                                                                                    If (Data(2).Substring(iLp_, 1) = "1"c) And .SysData(0).MarkData2(iLp).Substring(iLp_, 1) = "T"c Then
                                                                                        Data(2) = "T" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                                Case Is = 2 'Recover Year Data Value (1234567890)
                                                                                    If Not IsNumeric(Data(2).Substring(iLp_, 1)) Then
                                                                                        If Data(2).Substring(iLp_, 1) = "T" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "W" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        End If
                                                                                    End If
                                                                                Case Is = 3 'Recover Month Data Value (123456789XYZ)
                                                                                    If Data(2).Substring(iLp_, 1) = "O" Or Data(2).Substring(iLp_, 1) = "0" Then
                                                                                        'Data(2) = Data(2).Substring(0, iLp_) & "C" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "T" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "W" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                                Case Is = 4 'Recover Version Value (dd)
                                                                                    If Data(2).Substring(iLp_, 1) = "C" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "T" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "W" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                                Case Is = 5 'Recover Version Value (dd)
                                                                                    If Data(2).Substring(iLp_, 1) = "C" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "T" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "4"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "3"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "5"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "W" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1"
                                                                                    End If
                                                                            End Select
                                                                    End Select

                                                                    Chk2 = Array.IndexOf(.SysData(0).MarkData2, Data(2))
                                                                End If
                                                            Catch ex As Exception
                                                                Chk2 = -1
                                                            End Try

                                                            If Not Chk2 < 0 Then Exit For
                                                        Next
                                                    End If

                                                    If Not Chk2 < 0 Then Exit For
                                                Next
                                            End If
                                        End If


                                        If Chk1 < 0 Or Chk2 < 0 Then
                                            If Not ReTry < 7 Then
                                                Dim RetVal As String = String.Empty

                                                Me.FZ_SerCmd("UNITDATA 15 125", RetVal)
                                                Dim FilterSize() As String = RetVal.Split(" "c)

                                                If Not Val(FilterSize(0)) = Val(.BGSmaxSet) Then
                                                    Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)

                                                    'Set min BGS Level
                                                    Me.FZ_SerCmd("UNITDATA 15 124 " & .BGSminSet)

                                                    'Set max BGS Level
                                                    Me.FZ_SerCmd("UNITDATA 15 125 " & .BGSmaxSet)

                                                    .AutoSeqNo = 0
                                                    Exit Select
                                                Else
                                                    ReTry = 0
                                                    Thread.Sleep(800)
                                                    .AutoSeqNo = 11
                                                    Exit Select
                                                End If
                                            Else
                                                Dim RetVal As String = String.Empty
                                                Dim BGSmin() As String = {}
                                                Dim BGSmax() As String = {}


                                                Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)
                                                Dim FilterSize() As String = RetVal.Split(" "c)

                                                If Val(FilterSize(0)) = 0 Then
                                                    Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)
                                                Else
                                                    Me.FZ_SerCmd("UNITDATA 18 10", RetVal)

                                                    If Not RetVal.IndexOf("OK") < 0 Then
                                                        Dim DensityData() As String = RetVal.Split(" "c)

                                                        If Val(DensityData(0)) > IIf(Val(FilterSize(0)) = 0, Val(.DensityLvl) + 15, Val(.DensityLvl)) AndAlso Val(DensityData(0)) < Val(.DensityLvl) + 50 Then
                                                            Me.FZ_SerCmd("UNITDATA 15 124", RetVal)
                                                            BGSmin = RetVal.Split(" "c)
                                                            Me.FZ_SerCmd("UNITDATA 15 124 " & Val(BGSmin(0) + 5).ToString, RetVal)

                                                            Me.FZ_SerCmd("UNITDATA 15 125", RetVal)
                                                            BGSmax = RetVal.Split(" "c)
                                                            Me.FZ_SerCmd("UNITDATA 15 125 " & Val(BGSmax(0) + 5).ToString, RetVal)
                                                        End If


                                                        If Val(DensityData(0)) < Val(.DensityLvl) - 10 AndAlso Val(DensityData(0)) > Val(.DensityLvl) - 60 Then
                                                            Me.FZ_SerCmd("UNITDATA 15 124", RetVal)
                                                            BGSmin = RetVal.Split(" "c)
                                                            Me.FZ_SerCmd("UNITDATA 15 124 " & Val(BGSmin(0) - 5).ToString, RetVal)


                                                            Me.FZ_SerCmd("UNITDATA 15 125", RetVal)
                                                            BGSmax = RetVal.Split(" "c)
                                                            Me.FZ_SerCmd("UNITDATA 15 125 " & Val(BGSmax(0) - 5).ToString, RetVal)
                                                        End If
                                                    End If
                                                End If


                                                ReTry += 1
                                                mn_StepperMove(-8)
                                                Thread.Sleep(180)
                                                mn_StepperMove(8)

                                                Thread.Sleep(80)
                                                .AutoSeqNo = 0
                                                Exit Select
                                            End If
                                        Else
                                            If .ColerationChk <> 0 Then
                                                Dim RetVal(1) As String

                                                If Data(2).Length = .SysData(0).MarkData2(0).Length Then
                                                    Me.FZ_SerCmd("UNITDATA 23 1163", RetVal(0))
                                                    'Me.FZ_SerCmd("UNITDATA 23 1164", RetVal(1))
                                                Else
                                                    For iLp As Integer = 0 To .SysData(0).MarkData2.GetUpperBound(0)
                                                        Application.DoEvents()

                                                        If Not Data(2).IndexOf(.SysData(0).MarkData2(iLp)) < 0 Then
                                                            Chk2 = iLp
                                                            Exit For
                                                        End If
                                                    Next

                                                    '1E1XAA
                                                    Dim datPos As Integer = Data(2).IndexOf(.SysData(0).MarkData2(Chk2))
                                                    Me.FZ_SerCmd("UNITDATA 23 " & (1163 + datPos).ToString.Trim, RetVal(0))
                                                    'Me.FZ_SerCmd("UNITDATA 24 " & (1163 + datPos).ToString.Trim, RetVal(1))
                                                End If


                                                Dim rslt1() As String = RetVal(0).Split(" "c)
                                                'Dim rslt2() As String = RetVal(1).Split(" "c)


                                                If Not CType(rslt1(0), Single) > .ColerationChk Then
                                                    If ReTry > 3 Then
                                                        .RT_Error = 90
                                                        .AutoSeqNo = 10
                                                    Else
                                                        mn_StepperMove(-8)
                                                        Thread.Sleep(180)
                                                        mn_StepperMove(8)

                                                        Thread.Sleep(80)
                                                        ReTry += 1
                                                        .AutoSeqNo = 0
                                                    End If

                                                    Exit Select
                                                End If
                                            End If

                                            'Inspection Data OK
                                            ReTry = 0

                                            '> 1 Mother Lot
                                            If Not fg_ChkWkJmp = 0 Then
                                                If .SysData(0).P_Lot.Length > 1 Then
                                                    If .SysData(0).pLotRun = "" Then
                                                        'First Change or First detection
                                                        .SysData(0).pLotRun = .SysData(0).MarkData2(Chk2)
                                                    ElseIf Not .SysData(0).pLotRun = .SysData(0).MarkData2(Chk2) Then
                                                        'When Week Code Change
                                                        Dim Total As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

                                                        If Not .SysData(0).P_Lot_r.Length = .SysData(0).P_Lot.Length Then
                                                            ReDim .SysData(0).P_Lot_r(.SysData(0).P_Lot.GetUpperBound(0))
                                                        End If

                                                        'Calculate Actual Quantity For The New Weekcode
                                                        Dim tmp As String = .SysData(0).pLotRun
                                                        Dim pData_r As System.Collections.Generic.IEnumerable(Of QtyCtrl) = .SysData(0).P_Lot_r.Where(Function(q) q.WeekCode <> tmp)

                                                        If pData_r.Count > 0 Then
                                                            For iLp As Integer = 0 To pData_r.Count - 1
                                                                Total -= pData_r(iLp).QtyUsed
                                                            Next
                                                        End If

                                                        'Get The Total Quantity For The Weekcode From The Input
                                                        For iLp As Integer = 0 To .SysData(0).P_Lot.Length - 1
                                                            If Not .SysData(0).P_Lot(iLp).WeekCode = Nothing AndAlso Not .SysData(0).P_Lot(iLp).WeekCode.IndexOf("O") < 0 Then
                                                                .SysData(0).P_Lot(iLp).WeekCode = .SysData(0).P_Lot(iLp).WeekCode.Replace("O", "0")
                                                            End If
                                                        Next


                                                        Dim pTotal As System.Collections.Generic.IEnumerable(Of QtyCtrl) = .SysData(0).P_Lot.Where(Function(q) q.WeekCode = tmp)
                                                        Dim Total_p As Integer = 0
                                                        Dim comp As Integer = -1
                                                        Dim wkchg1 As Integer = -1
                                                        Dim wkchg2 As Integer = -1
                                                        Dim wk_Qty() As Integer

                                                        If pTotal.Count > 0 Then
                                                            ReDim wk_Qty(pTotal.Count - 1)

                                                            For iLp As Integer = 0 To pTotal.Count - 1
                                                                Total_p += pTotal(iLp).QtyUsed
                                                                wk_Qty(iLp) = pTotal(iLp).QtyUsed
                                                            Next

                                                            comp = Array.IndexOf(wk_Qty, Total)

                                                            For iLp As Integer = 0 To wk_Qty.GetUpperBound(0)
                                                                If Total >= wk_Qty(iLp) - WkCdChgRng Then
                                                                    wkchg1 = 0
                                                                    Exit For
                                                                End If

                                                                If Total >= wk_Qty(iLp) + WkCdChgRng Then
                                                                    wkchg1 = 0
                                                                    Exit For
                                                                End If
                                                            Next
                                                        End If

                                                        Dim Total_x As Integer = Total - .SysData(0).WeekCodeJump


                                                        If pTotal.Count > 0 And (Total >= Total_p Or Not comp < 0 Or Not wkchg1 < 0) Then
                                                            'Weekcode completely end
                                                            For iLp = 0 To .SysData(0).P_Lot_r.GetUpperBound(0)
                                                                If .SysData(0).P_Lot_r(iLp).WeekCode = "" Then
                                                                    .SysData(0).P_Lot_r(iLp) = pTotal(0)
                                                                    .SysData(0).P_Lot_r(iLp).QtyUsed = Total
                                                                    Exit For
                                                                End If
                                                            Next

                                                            .ContWkJmp = 0
                                                            .SysData(0).pLotRun = .SysData(0).MarkData2(Chk2)
                                                        Else
                                                            Dim pData As System.Collections.Generic.IEnumerable(Of QtyCtrl)

                                                            'Weekcode not completely finish
                                                            If WkJmpReTry >= 3 Then
                                                                WkJmpReTry = 0

                                                                'Wrong Qty Counting
                                                                If .ContWkJmp > ContWkJmpMax Then
                                                                    .ContWkJmp = 0
                                                                    pData = .SysData(1).P_Lot.Where(Function(q) q.WeekCode = tmp)

                                                                    If pData.Count = 1 Then
                                                                        For iLp = 0 To .SysData(0).P_Lot_r.GetUpperBound(0)
                                                                            If .SysData(0).P_Lot_r(iLp).WeekCode = "" Then
                                                                                .SysData(0).P_Lot_r(iLp) = pData(0)
                                                                                Exit For
                                                                            End If
                                                                        Next
                                                                    Else
                                                                        Dim n As Integer = 0
                                                                        Dim d As Integer = 0
                                                                        Dim sq As Integer = 0

                                                                        'Finding nearest qty setting
                                                                        For iLp As Integer = 0 To pData.Count - 1
                                                                            d = pData(iLp).QtyUsed - Total

                                                                            If n <> 0 Then
                                                                                If d > n Then
                                                                                    n = d
                                                                                    sq = iLp
                                                                                End If
                                                                            Else
                                                                                n = d
                                                                                sq = iLp
                                                                            End If
                                                                        Next

                                                                        For iLp = 0 To .SysData(0).P_Lot_r.GetUpperBound(0)
                                                                            If .SysData(0).P_Lot_r(iLp).WeekCode = "" Then
                                                                                .SysData(0).P_Lot_r(iLp) = pData(sq)
                                                                                Exit For
                                                                            End If
                                                                        Next
                                                                    End If


                                                                    WkJmpReTry = 0
                                                                    .SysData(0).pLotRun = .SysData(0).MarkData2(Chk2)
                                                                    .SysData(0).InspCnt += 1
                                                                    .SysData(0).WeekCodeJump -= ContWkJmpMax
                                                                    .SysData(0).InspCnt += ContWkJmpMax


                                                                    Dim InspDate_ As String = String.Format("{0:D2}-{1:D2}-{2:D4}", .SysData(0).InspDate.Day, .SysData(0).InspDate.Month, .SysData(0).InspDate.Year)
                                                                    Dim PathName_ As String = .LotDataPath & "\" & .SysData(0).Prod & "\" & InspDate_ & "\" & .SysData(0).LotNo

                                                                    If Not Val(.SysData(0).InspType) = 0 Then
                                                                        PathName_ &= "\SealInsp\RawData"
                                                                    Else
                                                                        PathName_ &= "\MarkInsp\RawData"
                                                                    End If

                                                                    If My.Computer.FileSystem.DirectoryExists(PathName_) = False Then
                                                                        My.Computer.FileSystem.CreateDirectory(PathName_)
                                                                    End If

                                                                    Dim RawDataFile_ As String = PathName_ & "\" & .SysData(0).LotNo & ".dat"
                                                                    My.Computer.FileSystem.WriteAllText(RawDataFile_, MeasData & vbCrLf, True, System.Text.Encoding.ASCII)

                                                                    UpdateCountingData()
                                                                    .AutoSeqNo += 1
                                                                Else
                                                                    .WC_Error = True
                                                                    .RT_Error = 70
                                                                    .AutoSeqNo = 10
                                                                End If
                                                            Else
                                                                WkJmpReTry += 1

                                                                mn_StepperMove(-8)
                                                                Thread.Sleep(180)
                                                                mn_StepperMove(8)

                                                                Thread.Sleep(80)
                                                                .AutoSeqNo = 0
                                                            End If

                                                            Exit Select
                                                        End If
                                                    End If
                                                End If
                                            End If

                                            WkJmpReTry = 0
                                            .ContWkJmp = 0
                                            .SysData(0).InspCnt += 1

                                            Dim InspDate As String = String.Format("{0:D2}-{1:D2}-{2:D4}", .SysData(0).InspDate.Day, .SysData(0).InspDate.Month, .SysData(0).InspDate.Year)
                                            Dim PathName As String = .LotDataPath & "\" & .SysData(0).Prod & "\" & InspDate & "\" & .SysData(0).LotNo

                                            If Not Val(.SysData(0).InspType) = 0 Then
                                                PathName &= "\SealInsp\RawData"
                                            Else
                                                PathName &= "\MarkInsp\RawData"
                                            End If

                                            If My.Computer.FileSystem.DirectoryExists(PathName) = False Then
                                                My.Computer.FileSystem.CreateDirectory(PathName)
                                            End If

                                            Dim RawDataFile As String = PathName & "\" & .SysData(0).LotNo & ".dat"
                                            My.Computer.FileSystem.WriteAllText(RawDataFile, MeasData & vbCrLf, True, System.Text.Encoding.ASCII)
                                        End If
                                    Else
                                        .AutoSeqNo += 1
                                    End If
                                Else
                                    If Is_rgbData.Count = 2 And .RedoFlag = False Then
                                        ReTry = 0
                                        .SysData(0).InspCnt += 1
                                    Else
                                        ReTry = 0
                                        .RT_Error = 90
                                        .AutoSeqNo = 10

                                        Exit Select
                                    End If
                                End If

                                UpdateCountingData()
                                .AutoSeqNo += 1
                            Case Is = 1
                                'Tape Seal Check move 20 deg/step.
                                'Marking Check move 60 deg/step.

                                Dim Total As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

                                If PullerTmr = 0 Then
                                    MotorTurn = 0
                                    .IO.M2_CW.Trigger_ON()
                                    PullerTmr = My.Computer.Clock.TickCount
                                End If

                                If Me.lbl_InspTapeType.Visible = True Then
                                    mn_StepperMove(-20)
                                    MotorTurn += 1

                                    If MotorTurn >= 30 * 3 Then
                                        MotorTurn = 0
                                        .IO.M2_CW.Trigger_ON()
                                        PullerTmr = My.Computer.Clock.TickCount
                                    End If
                                Else
                                    mn_StepperMove()
                                    MotorTurn += 1

                                    If MotorTurn >= 30 Then
                                        MotorTurn = 0
                                        .IO.M2_CW.Trigger_ON()
                                        PullerTmr = My.Computer.Clock.TickCount
                                    End If
                                End If

                                Thread.Sleep(60)
                                .AutoSeqNo = 0

                                Dim TimePerCyc As String = String.Format("{0:F2}sec/pcs", (My.Computer.Clock.TickCount - CycTime) / 1000)

                                If Me.lbl_CycTime.InvokeRequired Then
                                    Me.lbl_CycTime.Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_CycTime, Thread.CurrentThread.Name, TimePerCyc})
                                Else
                                    Me.lbl_CycTime.Text = TimePerCyc
                                End If

                                Dim RunDuration As Integer = My.Computer.Clock.TickCount - .TimeEllapse
                                Dim TotalTimeEllapse As String = String.Format("{0:F0} min {1:F0} sec Elapse", RunDuration \ 60000, ((RunDuration Mod 60000) / 1000))

                                If Me.lbl_TimeEllapse.InvokeRequired Then
                                    Me.lbl_TimeEllapse.Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_TimeEllapse, Thread.CurrentThread.Name, TotalTimeEllapse})
                                Else
                                    Me.lbl_TimeEllapse.Text = TotalTimeEllapse
                                End If
                            Case Is = 2
                                Dim Total As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

                                If Total < .SysData(0).Acceptance Then
                                    .AutoSeqNo = 0
                                Else
                                    .AutoSeqNo = 0
                                End If
                            Case Is = 10        'System Error Routine
                                Dim ImageFile As String = String.Empty
                                Dim InspDate As String = String.Format("{0:D2}-{1:D2}-{2:D4}", .SysData(0).InspDate.Day, .SysData(0).InspDate.Month, .SysData(0).InspDate.Year)
                                Dim PathName As String = .LotDataPath & "\" & .SysData(0).Prod & "\" & InspDate & "\" & .SysData(0).LotNo

                                If Not Val(.SysData(0).InspType) = 0 Then
                                    PathName &= "\SealInsp\ImgData"
                                Else
                                    PathName &= "\MarkInsp\ImgData"
                                End If

                                If My.Computer.FileSystem.DirectoryExists(PathName) = False Then
                                    My.Computer.FileSystem.CreateDirectory(PathName)
                                End If

                                Me.GetImage(ImageFile)

                                If ImageFile <> "-" Then
                                    Dim imgFile() As String = ImageFile.Split("\"c)

                                    Try
                                        My.Computer.FileSystem.CopyFile(ImageFile.Replace(".ifz", ".jpg"), PathName & "\" & imgFile(imgFile.GetUpperBound(0)).Replace(".ifz", ".jpg"))
                                    Catch ex As Exception

                                    End Try
                                End If

                                Dim RetVal As String = Nothing
                                FZ_SerCmd("SCENE 31", RetVal)

                                SetAlarm(.RT_Error)
                                .SysData(0).BuzzCnt += 1
                                UpdateCountingData()

                                Do While .Mode = SysAppMode.app_AutoRun
                                    Application.DoEvents()
                                Loop

                                'Check NG Label Sticker
                                .AutoSeqNo = 11
                            Case Is = 11
                                '.WC_Error = False
                                pic_Camera.BackColor = RedLED_OnOff(1)
                                pic_Camera_.BackColor = RedLED_OnOff(1)
                                pic_VisionChk.BackColor = RedLED_OnOff(1)
                                Dim MeasData As String = String.Empty

                                .IO.FZ_DI0.Trigger_OFF()

                                'Check Tape Seal ???
                                If Me.lbl_InspTapeType.Visible = True Then
                                    .IO.FZ_DI1.Trigger_ON()
                                End If

                                FZ_SerCmd("MEASURE", MeasData)
                                Dim MeasData_ As String = MeasData

                                pic_Camera.BackColor = RedLED_OnOff(0)
                                pic_Camera_.BackColor = RedLED_OnOff(0)
                                pic_VisionChk.BackColor = RedLED_OnOff(0)


                                If Me.lbl_VZData.InvokeRequired Then
                                    Me.lbl_VZData.Invoke(New frm_Main.DispMsg(AddressOf Me.DispToolTipLabel), New Object() {Me.lbl_VZData, Thread.CurrentThread.Name, MeasData.Trim})
                                End If


                                '--- Tape Seal Inspection Judgment ---
                                'Respond from Vision -> 'OK 72.28,46.79 6666
                                If Not Val(.SysData(0).InspType) = 0 Then
                                    'Saving Data
                                    Dim InspDate As String = String.Format("{0:D2}-{1:D2}-{2:D4}", .SysData(0).InspDate.Day, .SysData(0).InspDate.Month, .SysData(0).InspDate.Year)
                                    Dim PathName As String = .LotDataPath & "\" & .SysData(0).Prod & "\" & InspDate & "\" & .SysData(0).LotNo

                                    If Not Val(.SysData(0).InspType) = 0 Then
                                        PathName &= "\SealInsp\RawData"
                                    Else
                                        PathName &= "\MarkInsp\RawData"
                                    End If

                                    If My.Computer.FileSystem.DirectoryExists(PathName) = False Then
                                        My.Computer.FileSystem.CreateDirectory(PathName)
                                    End If

                                    Dim RawDataFile As String = PathName & "\" & .SysData(0).LotNo & ".dat"
                                    My.Computer.FileSystem.WriteAllText(RawDataFile, MeasData_ & vbCrLf, True, System.Text.Encoding.ASCII)


                                    If Not MeasData.IndexOf("6666") < 0 Then
                                        MeasData = MeasData.Replace(","c, " "c)
                                        Dim Seal_() As String = MeasData.Split(" "c)
                                        Dim SealDim_ As System.Collections.Generic.IEnumerable(Of String) = Seal_.Where(Function(q) q.Length > 0 And q <> "6666" And q <> "OK")

                                        'Check Range Here
                                        If Val(SealDim_(0)) < .InspTapeSeal.A_min Or Val(SealDim_(0)) > .InspTapeSeal.A_max Or Val(SealDim_(1)) < .InspTapeSeal.B_min Or Val(SealDim_(1)) > .InspTapeSeal.B_max Then
                                            If Not ReTry < 2 Then
                                                ReTry = 0
                                                .RT_Error = 90
                                                .AutoSeqNo = 10
                                                Exit Select
                                            Else
                                                ReTry += 1
                                                mn_StepperMove(-8)
                                                Thread.Sleep(180)
                                                mn_StepperMove(8)

                                                Thread.Sleep(80)
                                                .AutoSeqNo = 11
                                                Exit Select
                                            End If
                                        End If

                                        'Searching First Xtal
                                        Dim Total As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

                                        If Total = 0 Then
                                            FZ_SerCmd("MEASURE", MeasData)

                                            If Not MeasData.IndexOf("8888") < 0 Then
                                                'Empty pocket
                                                .AutoSeqNo += 1
                                                Exit Select
                                            Else
                                                Dim Data_() As String = MeasData.Split(" "c)
                                                Dim Is_rgbData_ As System.Collections.Generic.IEnumerable(Of String) = Data_.Where(Function(q) q.Length > 0 And q <> "9999" And q <> "OK")

                                                If Is_rgbData_.Count < 2 Then
                                                    .AutoSeqNo += 1
                                                    Exit Select
                                                Else
                                                    'First Xtal found here
                                                    .SysData(0).InspCnt = 1
                                                    SqCnt = 3
                                                    UpdateCountingData()
                                                    .AutoSeqNo += 1

                                                    Exit Select
                                                End If
                                            End If
                                        Else
                                            SqCnt += 1
                                            .SysData(0).InspCnt = SqCnt \ 3
                                            UpdateCountingData()

                                            Total = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

                                            If Total >= Val(.SysData(0).Acceptance) + 25 Then
                                                'Finished Inspection (included 25 empty pocket at the end)
                                                SetAlarm(99)
                                                .AutoSeqNo = 0

                                                InspDone()
                                                Exit Select
                                            Else
                                                'Tape Seal Inspection Continue
                                                .AutoSeqNo += 1
                                                Exit Select
                                            End If
                                        End If
                                    End If
                                End If              '--- Tape Seal Inspection Judgment Procedure End ---


                                '--- Marking Inspection ---
                                Dim Data() As String = MeasData.Split(" "c)
                                Dim ChkDataFmt As Integer = -1

                                If Data.GetUpperBound(0) = 4 And Data(Data.GetUpperBound(0)) = "9999" Then
                                    If Not .SysData(0).Prod.ToUpper.IndexOf("3225") < 0 Then
                                        Data(2) = Data(2) & " " & Data(3)
                                        Data(3) = "9999"
                                        ReDim Preserve Data(3)
                                        ChkDataFmt = 0
                                    End If
                                ElseIf Data.GetUpperBound(0) = 2 And Data(Data.GetUpperBound(0)) = "9999" Then
                                    If .SysData(0).Prod.ToUpper = "RAKON" Then
                                        ReDim Preserve Data(3)
                                        Data(3) = "9999"
                                        Data(2) = Data(1)
                                        ChkDataFmt = 0
                                    End If
                                ElseIf Data.GetUpperBound(0) = 3 And Data(Data.GetUpperBound(0)) = "9999" Then
                                    If Data(1).Trim.Length < .SysData(0).MarkData1(0).Length Or Data(2).Trim = "" Then
                                        ChkDataFmt = -1
                                    Else
                                        If Data(2) = "" Then Data(2) = Data(1)
                                        If Data(1) = "" Then Data(1) = Data(2)
                                        ChkDataFmt = 0
                                    End If
                                ElseIf Not Array.IndexOf(Data, "8888") < 0 Then
                                    ChkDataFmt = 0
                                End If


                                'In correct Vision Data Format
                                If ChkDataFmt < 0 Then
                                    If Me.lbl_VZData.InvokeRequired Then
                                        Me.lbl_VZData.Invoke(New frm_Main.DispMsg(AddressOf Me.DispLabel), New Object() {Me.lbl_VZData, Thread.CurrentThread.Name, MeasData.Trim})
                                    Else
                                        Me.lbl_VZData.Text = MeasData.Trim
                                    End If

                                    FmtRetry += 1

                                    If FmtRetry > 7 Then
                                        Dim RetVal As String = String.Empty

                                        Me.FZ_SerCmd("UNITDATA 15 125", RetVal)
                                        Dim FilterSize() As String = RetVal.Split(" "c)

                                        If Not Val(FilterSize(0)) = Val(.BGSmaxSet) Then
                                            Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)

                                            'Set min BGS Level
                                            Me.FZ_SerCmd("UNITDATA 15 124 " & .BGSminSet)

                                            'Set max BGS Level
                                            Me.FZ_SerCmd("UNITDATA 15 125 " & .BGSmaxSet)

                                            .AutoSeqNo = 11
                                            Exit Select
                                        Else
                                            FmtRetry = 0
                                            ReTry = 0
                                            .AutoSeqNo = 20
                                        End If
                                    Else
                                        If Data.GetUpperBound(0) < 2 OrElse (Data(1) <> "" Or Data(2) <> "") Then
                                            Dim RetVal As String = String.Empty
                                            Dim BGSmin() As String = {}
                                            Dim BGSmax() As String = {}


                                            Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)
                                            Dim FilterSize() As String = RetVal.Split(" "c)

                                            If Val(FilterSize(0)) = 0 Then
                                                Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)
                                            Else
                                                Me.FZ_SerCmd("UNITDATA 18 10", RetVal)

                                                If Not RetVal.IndexOf("OK") < 0 Then
                                                    Dim DensityData() As String = RetVal.Split(" "c)

                                                    If Val(DensityData(0)) > IIf(Val(FilterSize(0)) = 0, Val(.DensityLvl) + 15, Val(.DensityLvl)) AndAlso Val(DensityData(0)) < Val(.DensityLvl) + 50 Then
                                                        Me.FZ_SerCmd("UNITDATA 15 124", RetVal)
                                                        BGSmin = RetVal.Split(" "c)
                                                        Me.FZ_SerCmd("UNITDATA 15 124 " & Val(BGSmin(0) + 5).ToString, RetVal)

                                                        Me.FZ_SerCmd("UNITDATA 15 125", RetVal)
                                                        BGSmax = RetVal.Split(" "c)
                                                        Me.FZ_SerCmd("UNITDATA 15 125 " & Val(BGSmax(0) + 5).ToString, RetVal)
                                                    End If


                                                    If Val(DensityData(0)) < Val(.DensityLvl) - 10 AndAlso Val(DensityData(0)) > Val(.DensityLvl) - 60 Then
                                                        Me.FZ_SerCmd("UNITDATA 15 124", RetVal)
                                                        BGSmin = RetVal.Split(" "c)
                                                        Me.FZ_SerCmd("UNITDATA 15 124 " & Val(BGSmin(0) - 5).ToString, RetVal)


                                                        Me.FZ_SerCmd("UNITDATA 15 125", RetVal)
                                                        BGSmax = RetVal.Split(" "c)
                                                        Me.FZ_SerCmd("UNITDATA 15 125 " & Val(BGSmax(0) - 5).ToString, RetVal)
                                                    End If
                                                End If
                                            End If
                                        End If


                                        mn_StepperMove(-8)
                                        Thread.Sleep(180)
                                        mn_StepperMove(8)

                                        Thread.Sleep(80)
                                        .AutoSeqNo = 11
                                    End If

                                    Exit Select
                                End If

                                FmtRetry = 0
                                Dim Is_rgbData As System.Collections.Generic.IEnumerable(Of String) = Data.Where(Function(q) q.Length >= 1 And q <> "9999" And q <> "OK")

                                If MeasData.IndexOf("8888") < 0 Then
                                    If Not EmptyPocket = 0 Then
                                        ReTry = 0
                                        .RT_Error = 71
                                        .AutoSeqNo = 10
                                        Exit Select
                                    End If

                                    If Me.lbl_VZData.InvokeRequired Then
                                        If Is_rgbData.Count = 2 Then
                                            Me.lbl_VZData.Invoke(New frm_Main.UpdateControl(AddressOf Me.DispControlValue), New Object() {Me.lbl_VZData, Thread.CurrentThread.Name, MeasData})
                                        End If
                                    Else
                                        Me.lbl_VZData.Text = MeasData
                                        Me.lbl_VisionData.Text = MeasData
                                    End If
                                Else
                                    If .SysData(0).InspCnt <> 0 And .SysData(0).InspCnt + .SysData(0).InspCnt_NG < Val(.SysData(0).Acceptance) Then
                                        ReTry = 0
                                        .RT_Error = 71
                                        .AutoSeqNo = 10
                                        Exit Select
                                    Else
                                        If .SysData(0).InspCnt = 0 Then
                                            ReTry = 0
                                            .AutoSeqNo = 0
                                            Exit Select
                                        End If
                                    End If
                                End If


                                '--- Inspect Vision Data ---
                                If Is_rgbData.Count = 2 Then
                                    If Me.lbl_InspTapeType.Visible = False Then
                                        Dim Chk1 As Integer = -1
                                        Dim Chk2 As Integer = -1


                                        If Not .SysData(0).MarkData1(0) = "" Then
                                            Chk1 = Array.IndexOf(.SysData(0).MarkData1, Data(1))
                                        Else
                                            Chk1 = 1
                                        End If


                                        'Check Week Code
                                        'Chk2 = Array.IndexOf(.SysData(0).MarkData2, Data(2))

                                        'No Check Week Code
                                        If .SysData(0).MarkData2(0).Substring(0, 1) = Data(2).Substring(0, 1) AndAlso .SysData(0).MarkData2(0).Substring(.SysData(0).MarkData2(0).Length - 1) = Data(2).Substring(Data(2).Length - 1) Then
                                            Chk2 = 1
                                        End If


                                        '
                                        'Recover Mis-judge Data
                                        If Chk1 < 0 And Data(1).Length = 5 Then
                                            Dim RcvData As String = String.Empty

                                            Select Case .SysData(0).Prod
                                                Case Is = "FA-20HDOT"
                                                    RcvData = Data(1).Substring(0, 3)
                                                    RcvData = RcvData.Replace("A", "4")
                                                    RcvData = RcvData.Replace("B", "3")
                                                    RcvData = RcvData.Replace("C", "0")
                                                    RcvData = RcvData.Replace("D", "0")
                                                    RcvData = RcvData.Replace("E", "1")
                                                    RcvData = RcvData.Replace("F", "1")
                                                    RcvData = RcvData.Replace("O", "0")
                                                    RcvData = RcvData.Replace("R", "3")
                                                    RcvData = RcvData.Replace("S", "5")
                                                    RcvData = RcvData.Replace("T", "1")
                                                    RcvData = RcvData.Replace("U", "0")
                                                    RcvData = RcvData.Replace("H", "8")
                                                    RcvData = RcvData.Replace("L", "1")

                                                    Data(1) = RcvData & Data(1).Substring(3, 2)
                                                Case Else
                                                    RcvData = Data(1).Substring(0, 4)
                                                    RcvData = RcvData.Replace("A", "4")
                                                    RcvData = RcvData.Replace("B", "3")
                                                    RcvData = RcvData.Replace("C", "0")
                                                    RcvData = RcvData.Replace("D", "0")
                                                    RcvData = RcvData.Replace("E", "1")
                                                    RcvData = RcvData.Replace("F", "1")
                                                    RcvData = RcvData.Replace("O", "0")
                                                    RcvData = RcvData.Replace("R", "0")
                                                    RcvData = RcvData.Replace("S", "5")
                                                    RcvData = RcvData.Replace("T", "1")
                                                    RcvData = RcvData.Replace("U", "0")
                                                    RcvData = RcvData.Replace("H", "8")
                                                    RcvData = RcvData.Replace("L", "1")

                                                    Data(1) = RcvData & Data(1).Substring(4, 1)
                                            End Select

                                            If Data(1).EndsWith("1") Or Data(1).EndsWith("F") Then
                                                Data(1) = Data(1).Substring(0, 4) & "P"
                                            End If

                                            If Data(1).EndsWith("0") Or Data(1).EndsWith("E") Then
                                                Data(1) = Data(1).Substring(0, 4) & "C"
                                            End If

                                            Chk1 = Array.IndexOf(.SysData(0).MarkData1, Data(1))
                                        End If


                                        'Recover Error
                                        If Chk1 < 0 Then
                                            If Data(1).Length > .SysData(0).MarkData1(0).Length Then
                                                For iLp As Integer = 0 To .SysData(0).MarkData1.GetUpperBound(0)
                                                    Application.DoEvents()

                                                    If Not Data(1).IndexOf(.SysData(0).MarkData1(iLp)) < 0 Then
                                                        Chk1 = iLp
                                                        Exit For
                                                    End If
                                                Next
                                            End If
                                        End If

                                        If Chk2 < 0 Then
                                            If .SysData(0).MarkData2(0).Length = 6 And .SysData(0).MarkData2(0).StartsWith("o") Then
                                                If Data(2).Length > 1 AndAlso Data(2).Substring(1, 1) = "E" And (Data(2).StartsWith("6") Or Data(2).StartsWith("9") Or Data(2).StartsWith("P")) Then
                                                    Data(2) = "o" & Data(2).Substring(1)
                                                    Chk2 = Array.IndexOf(.SysData(0).MarkData2, Data(2))
                                                End If
                                            End If

                                            If Data(2).Length > .SysData(0).MarkData2(0).Length Then
                                                For iLp As Integer = 0 To .SysData(0).MarkData2.GetUpperBound(0)
                                                    Application.DoEvents()

                                                    If Not Data(2).IndexOf(.SysData(0).MarkData2(iLp)) < 0 Then
                                                        Chk2 = iLp
                                                        Exit For
                                                    End If
                                                Next
                                            End If

                                            If Chk2 < 0 Then
                                                For iLp = 0 To .SysData(0).MarkData2.GetUpperBound(0)
                                                    Application.DoEvents()
                                                    If Not Chk2 < 0 Then Exit For

                                                    If Data(2).Length = .SysData(0).MarkData2(iLp).Length Then
                                                        For iLp_ As Integer = 0 To .SysData(0).MarkData2(iLp).Length - 1
                                                            Application.DoEvents()

                                                            Try
                                                                If Not Data(2).Substring(iLp_, 1) = .SysData(0).MarkData2(iLp).Substring(iLp_, 1) Then
                                                                    Select Case .SysData(0).Prod
                                                                        Case Is = "FA-20HDOT"
                                                                            Select Case iLp_
                                                                                Case Is = 1
                                                                                    If Data(2).Substring(iLp_, 1) = "1"c And .SysData(0).MarkData2(iLp).Substring(iLp_, 1) = "E"c Then
                                                                                        Data(2) = Data(2).Substring(iLp_ - 1, 1) & "E" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                                Case Is = 2 'Recover Year Data Value (1234567890)
                                                                                    If Not IsNumeric(Data(2).Substring(iLp_, 1)) Then
                                                                                        If Data(2).Substring(iLp_, 1) = "T" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        End If
                                                                                    End If
                                                                                Case Is = 3 'Recover Month Data Value (123456789XYZ)
                                                                                    If Data(2).Substring(iLp_, 1) = "O" Or Data(2).Substring(iLp_, 1) = "0" Then
                                                                                        'Data(2) = Data(2).Substring(0, iLp_) & "C" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "T" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                                Case Is = 4 'Recover Day Data Value (123456789XYZ)
                                                                                    If Data(2).Substring(iLp_, 1) = "o" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "9" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                            End Select

                                                                        Case Is = "FA-20H", "FA-128", "FA-23", "FA-23A", "FA-238", "FA-118T", "FA-206"
                                                                            Select Case iLp_
                                                                                Case Is = 0
                                                                                    If Data(2).Substring(iLp_, 1) = "1"c And .SysData(0).MarkData2(iLp).Substring(iLp_, 1) = "E"c Then
                                                                                        Data(2) = "E" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                                Case Is = 1 'Recover Year Data Value (1234567890)
                                                                                    If Not IsNumeric(Data(2).Substring(iLp_, 1)) Then
                                                                                        If Data(2).Substring(iLp_, 1) = "T" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        End If
                                                                                    End If
                                                                                Case Is = 2 'Recover Month Data Value (123456789XYZ)
                                                                                    If Data(2).Substring(iLp_, 1) = "O" Or Data(2).Substring(iLp_, 1) = "0" Then
                                                                                        'Data(2) = Data(2).Substring(0, iLp_) & "C" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "T" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "W" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                            End Select

                                                                        Case Is = "TSX-3225", "TD3225N"
                                                                            Select Case iLp_
                                                                                Case Is = 0
                                                                                    If Data(2).Substring(iLp_, 1) = "1"c And .SysData(0).MarkData2(iLp).Substring(iLp_, 1) = "E"c Then
                                                                                        Data(2) = "E" & Data(2).Substring(iLp_ + 1)
                                                                                    End If

                                                                                    If (Data(2).Substring(iLp_, 1) = "1"c) And .SysData(0).MarkData2(iLp).Substring(iLp_, 1) = "T"c Then
                                                                                        Data(2) = "T" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                                Case Is = 2 'Recover Year Data Value (1234567890)
                                                                                    If Not IsNumeric(Data(2).Substring(iLp_, 1)) Then
                                                                                        If Data(2).Substring(iLp_, 1) = "T" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "W" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                        ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                            Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                        End If
                                                                                    End If
                                                                                Case Is = 3 'Recover Month Data Value (123456789XYZ)
                                                                                    If Data(2).Substring(iLp_, 1) = "O" Or Data(2).Substring(iLp_, 1) = "0" Then
                                                                                        'Data(2) = Data(2).Substring(0, iLp_) & "C" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "T" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "W" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                                Case Is = 4 'Recover Version Value (dd)
                                                                                    If Data(2).Substring(iLp_, 1) = "C" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "T" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "4" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "C" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "3" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "5" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "W" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0" & Data(2).Substring(iLp_ + 1)
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1" & Data(2).Substring(iLp_ + 1)
                                                                                    End If
                                                                                Case Is = 5 'Recover Version Value (dd)
                                                                                    If Data(2).Substring(iLp_, 1) = "C" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "T" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "B" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "A" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "4"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "D" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "E" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "F" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "H" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "O" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "R" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "3"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "S" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "5"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "W" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "8"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "U" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "0"
                                                                                    ElseIf Data(2).Substring(iLp_, 1) = "L" Then
                                                                                        Data(2) = Data(2).Substring(0, iLp_) & "1"
                                                                                    End If
                                                                            End Select
                                                                    End Select

                                                                    Chk2 = Array.IndexOf(.SysData(0).MarkData2, Data(2))
                                                                End If
                                                            Catch ex As Exception
                                                                Chk2 = -1
                                                            End Try

                                                            If Not Chk2 < 0 Then Exit For
                                                        Next
                                                    End If

                                                    If Not Chk2 < 0 Then Exit For
                                                Next
                                            End If
                                        End If


                                        If Chk1 < 0 Or Chk2 < 0 Then
                                            If Not ReTry < 3 Then
                                                Dim RetVal As String = String.Empty

                                                Me.FZ_SerCmd("UNITDATA 15 125", RetVal)
                                                Dim FilterSize() As String = RetVal.Split(" "c)

                                                If Not Val(FilterSize(0)) = Val(.BGSmaxSet) Then
                                                    Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)

                                                    'Set min BGS Level
                                                    Me.FZ_SerCmd("UNITDATA 15 124 " & .BGSminSet)

                                                    'Set max BGS Level
                                                    Me.FZ_SerCmd("UNITDATA 15 125 " & .BGSmaxSet)

                                                    .AutoSeqNo = 11
                                                    Exit Select
                                                Else
                                                    ReTry = 0
                                                    .AutoSeqNo = 20
                                                    Exit Select
                                                End If
                                            Else
                                                Dim RetVal As String = String.Empty
                                                Dim BGSmin() As String = {}
                                                Dim BGSmax() As String = {}


                                                Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)
                                                Dim FilterSize() As String = RetVal.Split(" "c)

                                                If Val(FilterSize(0)) = 0 Then
                                                    Me.FZ_SerCmd("UNITDATA 15 123 " & .FilterType)
                                                Else
                                                    Me.FZ_SerCmd("UNITDATA 18 10", RetVal)

                                                    If Not RetVal.IndexOf("OK") < 0 Then
                                                        Dim DensityData() As String = RetVal.Split(" "c)

                                                        If Val(DensityData(0)) > IIf(Val(FilterSize(0)) = 0, Val(.DensityLvl) + 15, Val(.DensityLvl)) AndAlso Val(DensityData(0)) < Val(.DensityLvl) + 50 Then
                                                            Me.FZ_SerCmd("UNITDATA 15 124", RetVal)
                                                            BGSmin = RetVal.Split(" "c)
                                                            Me.FZ_SerCmd("UNITDATA 15 124 " & Val(BGSmin(0) + 5).ToString, RetVal)

                                                            Me.FZ_SerCmd("UNITDATA 15 125", RetVal)
                                                            BGSmax = RetVal.Split(" "c)
                                                            Me.FZ_SerCmd("UNITDATA 15 125 " & Val(BGSmax(0) + 5).ToString, RetVal)
                                                        End If


                                                        If Val(DensityData(0)) < Val(.DensityLvl) - 10 AndAlso Val(DensityData(0)) > Val(.DensityLvl) - 60 Then
                                                            Me.FZ_SerCmd("UNITDATA 15 124", RetVal)
                                                            BGSmin = RetVal.Split(" "c)
                                                            Me.FZ_SerCmd("UNITDATA 15 124 " & Val(BGSmin(0) - 5).ToString, RetVal)


                                                            Me.FZ_SerCmd("UNITDATA 15 125", RetVal)
                                                            BGSmax = RetVal.Split(" "c)
                                                            Me.FZ_SerCmd("UNITDATA 15 125 " & Val(BGSmax(0) - 5).ToString, RetVal)
                                                        End If
                                                    End If
                                                End If


                                                ReTry += 1
                                                mn_StepperMove(-8)
                                                Thread.Sleep(180)
                                                mn_StepperMove(8)

                                                Thread.Sleep(80)
                                                .AutoSeqNo = 11
                                                Exit Select
                                            End If
                                        Else
                                            If .ColerationChk <> 0 Then
                                                Dim RetVal(1) As String

                                                If Data(2).Length = .SysData(0).MarkData2(0).Length Then
                                                    Me.FZ_SerCmd("UNITDATA 23 1163", RetVal(0))
                                                    'Me.FZ_SerCmd("UNITDATA 23 1164", RetVal(1))
                                                Else
                                                    For iLp As Integer = 0 To .SysData(0).MarkData2.GetUpperBound(0)
                                                        Application.DoEvents()

                                                        If Not Data(2).IndexOf(.SysData(0).MarkData2(iLp)) < 0 Then
                                                            Chk2 = iLp
                                                            Exit For
                                                        End If
                                                    Next

                                                    '1E1XAA
                                                    Dim datPos As Integer = Data(2).IndexOf(.SysData(0).MarkData2(Chk2))
                                                    Me.FZ_SerCmd("UNITDATA 23 " & (1163 + datPos).ToString.Trim, RetVal(0))
                                                    'Me.FZ_SerCmd("UNITDATA 24 " & (1163 + datPos).ToString.Trim, RetVal(1))
                                                End If


                                                Dim rslt1() As String = RetVal(0).Split(" "c)
                                                'Dim rslt2() As String = RetVal(1).Split(" "c)


                                                If Not CType(rslt1(0), Single) > .ColerationChk Then
                                                    If ReTry > 7 Then
                                                        .RT_Error = 90
                                                        .AutoSeqNo = 10
                                                    Else
                                                        mn_StepperMove(-8)
                                                        Thread.Sleep(180)
                                                        mn_StepperMove(8)

                                                        Thread.Sleep(80)
                                                        ReTry += 1
                                                        .AutoSeqNo = 11
                                                    End If

                                                    Exit Select
                                                End If
                                            End If


                                            If Not fg_ChkWkJmp = 0 Then
                                                If .SysData(0).P_Lot.Length > 1 Then
                                                    If .SysData(0).pLotRun = "" Then
                                                        .SysData(0).pLotRun = .SysData(0).MarkData2(Chk2)
                                                    ElseIf Not .SysData(0).pLotRun = .SysData(0).MarkData2(Chk2) Then
                                                        'When Week Code Change
                                                        Dim Total As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

                                                        If Not .SysData(0).P_Lot_r.Length = .SysData(0).P_Lot.Length Then
                                                            ReDim .SysData(0).P_Lot_r(.SysData(0).P_Lot.GetUpperBound(0))
                                                        End If

                                                        'Calculate Actual Quantity For The New Weekcode
                                                        Dim tmp As String = .SysData(0).pLotRun
                                                        Dim pData_r As System.Collections.Generic.IEnumerable(Of QtyCtrl) = .SysData(0).P_Lot_r.Where(Function(q) q.WeekCode <> tmp)

                                                        If pData_r.Count > 0 Then
                                                            For iLp As Integer = 0 To pData_r.Count - 1
                                                                Total -= pData_r(iLp).QtyUsed
                                                            Next
                                                        End If

                                                        'Get The Total Quantity For The Weekcode From The Input
                                                        For iLp As Integer = 0 To .SysData(0).P_Lot.Length - 1
                                                            If Not .SysData(0).P_Lot(iLp).WeekCode.IndexOf("O") < 0 Then
                                                                .SysData(0).P_Lot(iLp).WeekCode = .SysData(0).P_Lot(iLp).WeekCode.Replace("O", "0")
                                                            End If
                                                        Next


                                                        Dim pTotal As System.Collections.Generic.IEnumerable(Of QtyCtrl) = .SysData(0).P_Lot.Where(Function(q) q.WeekCode = tmp)
                                                        Dim Total_p As Integer = 0
                                                        Dim comp As Integer = -1
                                                        Dim wkchg1 As Integer = -1
                                                        Dim wkchg2 As Integer = -1
                                                        Dim wk_Qty() As Integer


                                                        If pTotal.Count > 0 Then
                                                            ReDim wk_Qty(pTotal.Count - 1)

                                                            For iLp As Integer = 0 To pTotal.Count - 1
                                                                Total_p += pTotal(iLp).QtyUsed
                                                                wk_Qty(iLp) = pTotal(iLp).QtyUsed
                                                            Next

                                                            comp = Array.IndexOf(wk_Qty, Total)

                                                            For iLp As Integer = 0 To wk_Qty.GetUpperBound(0)
                                                                If Total >= wk_Qty(iLp) - WkCdChgRng Then
                                                                    wkchg1 = 0
                                                                    Exit For
                                                                End If

                                                                If Total >= wk_Qty(iLp) + WkCdChgRng Then
                                                                    wkchg1 = 0
                                                                    Exit For
                                                                End If
                                                            Next
                                                        End If


                                                        Dim Total_x As Integer = Total - .SysData(0).WeekCodeJump

                                                        If pTotal.Count > 0 And (Total >= Total_p Or Not comp < 0 Or Not wkchg1 < 0) Then
                                                            'Weekcode completely end
                                                            For iLp = 0 To .SysData(0).P_Lot_r.GetUpperBound(0)
                                                                If .SysData(0).P_Lot_r(iLp).WeekCode = "" Then
                                                                    .SysData(0).P_Lot_r(iLp) = pTotal(0)
                                                                    .SysData(0).P_Lot_r(iLp).QtyUsed = Total
                                                                    Exit For
                                                                End If
                                                            Next

                                                            .ContWkJmp = 0
                                                            .SysData(0).pLotRun = .SysData(0).MarkData2(Chk2)
                                                        Else
                                                            Dim pData As System.Collections.Generic.IEnumerable(Of QtyCtrl)

                                                            'Weekcode not completely finish
                                                            If WkJmpReTry >= 3 Then
                                                                WkJmpReTry = 0

                                                                'Wrong Qty Counting
                                                                If .ContWkJmp > ContWkJmpMax Then
                                                                    .ContWkJmp = 0
                                                                    pData = .SysData(1).P_Lot.Where(Function(q) q.WeekCode = tmp)

                                                                    If pData.Count = 1 Then
                                                                        For iLp = 0 To .SysData(0).P_Lot_r.GetUpperBound(0)
                                                                            If .SysData(0).P_Lot_r(iLp).WeekCode = "" Then
                                                                                .SysData(0).P_Lot_r(iLp) = pData(0)
                                                                                Exit For
                                                                            End If
                                                                        Next
                                                                    Else
                                                                        Dim n As Integer = 0
                                                                        Dim d As Integer = 0
                                                                        Dim sq As Integer = 0

                                                                        'Finding nearest qty setting
                                                                        For iLp As Integer = 0 To pData.Count - 1
                                                                            d = pData(iLp).QtyUsed - Total

                                                                            If n <> 0 Then
                                                                                If d > n Then
                                                                                    n = d
                                                                                    sq = iLp
                                                                                End If
                                                                            Else
                                                                                n = d
                                                                                sq = iLp
                                                                            End If
                                                                        Next

                                                                        For iLp = 0 To .SysData(0).P_Lot_r.GetUpperBound(0)
                                                                            If .SysData(0).P_Lot_r(iLp).WeekCode = "" Then
                                                                                .SysData(0).P_Lot_r(iLp) = pData(sq)
                                                                                Exit For
                                                                            End If
                                                                        Next
                                                                    End If


                                                                    WkJmpReTry = 0
                                                                    .SysData(0).pLotRun = .SysData(0).MarkData2(Chk2)
                                                                    .SysData(0).InspCnt += 1
                                                                    .SysData(0).WeekCodeJump -= ContWkJmpMax
                                                                    .SysData(0).InspCnt += ContWkJmpMax


                                                                    Dim InspDate_ As String = String.Format("{0:D2}-{1:D2}-{2:D4}", .SysData(0).InspDate.Day, .SysData(0).InspDate.Month, .SysData(0).InspDate.Year)
                                                                    Dim PathName_ As String = .LotDataPath & "\" & .SysData(0).Prod & "\" & InspDate_ & "\" & .SysData(0).LotNo

                                                                    If Not Val(.SysData(0).InspType) = 0 Then
                                                                        PathName_ &= "\SealInsp\RawData"
                                                                    Else
                                                                        PathName_ &= "\MarkInsp\RawData"
                                                                    End If

                                                                    If My.Computer.FileSystem.DirectoryExists(PathName_) = False Then
                                                                        My.Computer.FileSystem.CreateDirectory(PathName_)
                                                                    End If

                                                                    Dim RawDataFile_ As String = PathName_ & "\" & .SysData(0).LotNo & ".dat"
                                                                    My.Computer.FileSystem.WriteAllText(RawDataFile_, MeasData & vbCrLf, True, System.Text.Encoding.ASCII)

                                                                    UpdateCountingData()
                                                                    .AutoSeqNo = 1
                                                                Else
                                                                    .WC_Error = True
                                                                    .RT_Error = 70
                                                                    .AutoSeqNo = 10
                                                                End If
                                                            Else
                                                                WkJmpReTry += 1

                                                                mn_StepperMove(-8)
                                                                Thread.Sleep(180)
                                                                mn_StepperMove(8)

                                                                Thread.Sleep(80)
                                                                .AutoSeqNo = 0
                                                            End If

                                                            Exit Select
                                                        End If
                                                    End If
                                                End If
                                            End If


                                            WkJmpReTry = 0
                                            .ContWkJmp = 0
                                            .SysData(0).InspCnt += 1

                                            Dim InspDate As String = String.Format("{0:D2}-{1:D2}-{2:D4}", .SysData(0).InspDate.Day, .SysData(0).InspDate.Month, .SysData(0).InspDate.Year)
                                            Dim PathName As String = .LotDataPath & "\" & .SysData(0).Prod & "\" & InspDate & "\" & .SysData(0).LotNo

                                            If Not Val(.SysData(0).InspType) = 0 Then
                                                PathName &= "\SealInsp\RawData"
                                            Else
                                                PathName &= "\MarkInsp\RawData"
                                            End If

                                            If My.Computer.FileSystem.DirectoryExists(PathName) = False Then
                                                My.Computer.FileSystem.CreateDirectory(PathName)
                                            End If

                                            Dim RawDataFile As String = PathName & "\" & .SysData(0).LotNo & ".dat"
                                            My.Computer.FileSystem.WriteAllText(RawDataFile, MeasData & vbCrLf, True, System.Text.Encoding.ASCII)
                                        End If
                                    Else
                                        .AutoSeqNo = 1
                                        Exit Select
                                    End If
                                Else
                                    ReTry = 0
                                    .RT_Error = 90
                                    .AutoSeqNo = 10
                                    Exit Select
                                End If

                                UpdateCountingData()
                                .AutoSeqNo = 1
                            Case Is = 20        'Chk & Confirm Sticker is put
                                Dim RetVal As String = String.Empty
                                FZ_SerCmd("SCENE 31", RetVal)

                                pic_Camera.BackColor = RedLED_OnOff(1)
                                pic_Camera_.BackColor = RedLED_OnOff(1)
                                pic_VisionChk.BackColor = RedLED_OnOff(1)


                                Dim MeasData As String = String.Empty
                                .IO.FZ_DI0.Trigger_OFF()
                                FZ_SerCmd("MEASURE", MeasData)

                                pic_Camera.BackColor = RedLED_OnOff(0)
                                pic_Camera_.BackColor = RedLED_OnOff(0)
                                pic_VisionChk.BackColor = RedLED_OnOff(0)


                                FZ_SerCmd("SCENE " & String.Format("{0:D2}", .SelectedSceneNo), RetVal)

                                If MeasData.IndexOf("9999") < 0 Then
                                    .RT_Error = 92
                                    .AutoSeqNo = 10      'Invalid data returns alarm
                                    Exit Select
                                Else
                                    Dim Data() As String = MeasData.Split(" "c)
                                    Dim rgbData As System.Collections.Generic.IEnumerable(Of String) = Data.Where(Function(q) q.Length > 0 And q <> "9999" And q <> "OK")

                                    If rgbData.Count > 0 Then
                                        Dim idx() As Integer = {0, 1, 2}
                                        Dim NoneSelectedIdx As System.Collections.Generic.IEnumerable(Of Integer) = idx.Where(Function(q) q <> .DefectStickerColor)

                                        If Val(rgbData(.DefectStickerColor)) < Val(rgbData(NoneSelectedIdx(0))) * 1.3 Or Val(rgbData(.DefectStickerColor)) < Val(rgbData(NoneSelectedIdx(1))) * 1.3 Then
                                            .RT_Error = 92
                                            .AutoSeqNo = 10
                                            Exit Select
                                        Else
                                            Dim InspDate As String = String.Format("{0:D2}-{1:D2}-{2:D4}", .SysData(0).InspDate.Day, .SysData(0).InspDate.Month, .SysData(0).InspDate.Year)
                                            Dim PathName As String = .LotDataPath & "\" & .SysData(0).Prod & "\" & InspDate & "\" & .SysData(0).LotNo

                                            If Not Val(.SysData(0).InspType) = 0 Then
                                                PathName &= "\SealInsp\RawData"
                                            Else
                                                PathName &= "\MarkInsp\RawData"
                                            End If

                                            If My.Computer.FileSystem.DirectoryExists(PathName) = False Then
                                                My.Computer.FileSystem.CreateDirectory(PathName)
                                            End If

                                            Dim RawDataFile As String = PathName & "\" & .SysData(0).LotNo & ".dat"
                                            My.Computer.FileSystem.WriteAllText(RawDataFile, MeasData & vbCrLf, True, System.Text.Encoding.ASCII)

                                            If .WC_Error = True Then
                                                .SysData(0).WeekCodeJump += 1
                                                .ContWkJmp += 1
                                            Else
                                                .SysData(0).InspCnt_NG += 1
                                                .ContWkJmp = 0
                                            End If

                                            UpdateCountingData()
                                            .WC_Error = False
                                            .AutoSeqNo = 1

                                            Dim NG_Pos As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump
                                            .RedoLoc_ &= NG_Pos.ToString.Trim & ","

                                            Exit Select
                                        End If
                                    Else
                                        .RT_Error = 92
                                        .AutoSeqNo = 10      'Invalid data returns alarm
                                        Exit Select
                                    End If
                                End If
                        End Select


                        Dim Total_ As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

                        With .IO
                            If Tp_OAI.Mode <> SysAppMode.app_AutoRun Then
                                Mtr_Tmr = 0
                                .M2_CCW.Trigger_OFF()
                                .M2_CW.Trigger_OFF()
                            End If

                            If .M2_CCW.BitState = cls_PCIBoard.BitsState.eBit_ON Or .M2_CW.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                                If Mtr_Tmr = 0 Then
                                    Select Case Total_
                                        Case 0 To 249
                                            Mtr_Tmr = 7000
                                        Case 250 To 499
                                            Mtr_Tmr = 6000
                                        Case 500 To 999
                                            Mtr_Tmr = 5000
                                        Case 1000 To 1499
                                            Mtr_Tmr = 5000
                                        Case 1500 To 1999
                                            Mtr_Tmr = 3800
                                        Case 2000 To 2499
                                            Mtr_Tmr = 3800
                                        Case Else
                                            Mtr_Tmr = 5000
                                    End Select
                                End If

                                If My.Computer.Clock.TickCount > PullerTmr + Mtr_Tmr Then
                                    PullerTmr = 0
                                    Mtr_Tmr = 0
                                    .M2_CCW.Trigger_OFF()
                                    .M2_CW.Trigger_OFF()
                                End If
                            End If


                            '.S_Cover.BitState = cls_PCIBoard.BitsState.eBit_ON - This bit has been remove...
                            If .RR.BitState = cls_PCIBoard.BitsState.eBit_OFF Or _
                                .RL.BitState = cls_PCIBoard.BitsState.eBit_OFF Or _
                                .FZ_Run.BitState = cls_PCIBoard.BitsState.eBit_OFF Then

                                PullerTmr = 0
                                Mtr_Tmr = 0
                                .M2_CCW.Trigger_OFF()
                                .M2_CW.Trigger_OFF()
                                SetAlarm(27)

                                Do While Tp_OAI.Mode = SysAppMode.app_AutoRun
                                    Application.DoEvents()
                                Loop
                            End If
                        End With
                    Else        'Not AutoRun Mode
                        If PullerTmr <> 0 Then
                            PullerTmr = 0
                            Mtr_Tmr = 0
                            .IO.M2_CCW.Trigger_OFF()
                            .IO.M2_CW.Trigger_OFF()
                        End If

                        'Check For Software Update
                        Dim Total As Integer = .SysData(0).InspCnt + .SysData(0).InspCnt_NG + .SysData(0).WeekCodeJump

                        If Total = 0 And AutoUpdate = 1 Then
                            If Not .Mode = SysAppMode.app_AutoRun And Not .Mode = SysAppMode.app_sysError Then
                                If My.Computer.Network.IsAvailable Then
                                    Try
                                        If My.Computer.Network.Ping("172.16.59.254") Then
                                            If System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed And System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CheckForUpdate Then
                                                With Me
                                                    .tmr_IOMonitor.Enabled = False
                                                    .tmr_ReadIO.Enabled = False
                                                    .cmd_Instruction.Enabled = False

                                                    If .lbl_SWUpdate.InvokeRequired Then
                                                        .lbl_SWUpdate.Invoke(New frm_Main.DispCtrl_(AddressOf Me.DispLabel_), New Object() {Me.lbl_SWUpdate, Thread.CurrentThread.Name, "True"})
                                                    Else
                                                        .lbl_SWUpdate.Visible = True
                                                    End If

                                                    If .pic_iMTEST.InvokeRequired Then
                                                        .pic_iMTEST.Invoke(New frm_Main.DispCtrl(AddressOf Me.DispControl_), New Object() {Me.pic_iMTEST, Thread.CurrentThread.Name, "True"})
                                                    Else
                                                        .pic_iMTEST.Visible = True
                                                    End If

                                                    Try
                                                        .Opacity = 0.7
                                                    Catch ex As Exception

                                                    End Try
                                                End With


                                                '--- Switch KVM ---
                                                .IO.KVM_PC.Trigger_OFF()
                                                .IO.KVM_VZ.Trigger_OFF()
                                                Thread.Sleep(50)

                                                KVM_Pos = 1
                                                .IO.KVM_PC.Trigger_ON()
                                                Thread.Sleep(100)
                                                .IO.KVM_PC.Trigger_OFF()

                                                'KVM_Pos = 2
                                                '.IO.KVM_VZ.Trigger_ON()
                                                'Thread.Sleep(100)
                                                '.IO.KVM_VZ.Trigger_OFF()

                                                'Display The Message
                                                MessageBox.Show("The program detected newer version is available on the publisher's network ! Please click 'OK' to proceed the update now!", "Network Deployment...", MessageBoxButtons.OK, MessageBoxIcon.Information)

                                                AddHandler System.Deployment.Application.ApplicationDeployment.CurrentDeployment.UpdateCompleted, AddressOf AppReStart
                                                System.Deployment.Application.ApplicationDeployment.CurrentDeployment.UpdateAsync()

                                                Exit Sub
                                            End If
                                        End If
                                    Catch ex As Exception
                                        If Me.pic_iMTEST.InvokeRequired Then
                                            Me.pic_iMTEST.Invoke(New frm_Main.DispCtrl(AddressOf Me.DispControl_), New Object() {Me.pic_iMTEST, Thread.CurrentThread.Name, "False"})
                                        Else
                                            Me.pic_iMTEST.Visible = False
                                        End If

                                        AutoUpdate = 0

                                        If ChkUpdateTempOffTmr = 0 Then
                                            ChkUpdateTempOffTmr = My.Computer.Clock.TickCount
                                        End If
                                    End Try
                                Else
                                    If Me.pic_iMTEST.InvokeRequired Then
                                        Me.pic_iMTEST.Invoke(New frm_Main.DispCtrl(AddressOf Me.DispControl_), New Object() {Me.pic_iMTEST, Thread.CurrentThread.Name, "False"})
                                    Else
                                        Me.pic_iMTEST.Visible = False
                                    End If

                                    AutoUpdate = 0

                                    If ChkUpdateTempOffTmr = 0 Then
                                        ChkUpdateTempOffTmr = My.Computer.Clock.TickCount
                                    End If
                                End If
                            End If
                        Else
                            If AutoUpdate = 0 And Not ChkUpdateTempOffTmr = 0 Then
                                If My.Computer.Clock.TickCount > ChkUpdateTempOffTmr + (1000 * 180) Then
                                    ChkUpdateTempOffTmr = 0
                                    AutoUpdate = 1

                                    If Me.pic_iMTEST.InvokeRequired Then
                                        Me.pic_iMTEST.Invoke(New frm_Main.DispCtrl(AddressOf Me.DispControl_), New Object() {Me.pic_iMTEST, Thread.CurrentThread.Name, "True"})
                                    Else
                                        Me.pic_iMTEST.Visible = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End With
            Loop
        End With

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        mn_StepperMove(3)
        mn_StepperMove(-3)

    End Sub

    Private Sub txt_A_max_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt_A_max.TextChanged, txt_A_min.TextChanged, txt_B_max.TextChanged, txt_B_min.TextChanged

        With Me
            If sender.Equals(.txt_A_max) Then
                If Not IsNumeric(.txt_A_max.Text) Then
                    .ErrorProvider1.SetError(.txt_A_max, "Must be a numeric value!")
                    .fg_A_max = 1
                Else
                    .ErrorProvider1.Clear()
                    .fg_A_max = 0
                    Tp_OAI.InspTapeSeal.A_max = CType(.txt_A_max.Text, Decimal)
                    regSubKey.SetValue("Seal_A_max", .txt_A_max.Text, RegistryValueKind.String)
                End If
            ElseIf sender.Equals(.txt_A_min) Then
                If Not IsNumeric(.txt_A_min.Text) Then
                    .ErrorProvider2.SetError(.txt_A_min, "Must be a numeric value!")
                    .fg_A_min = 1
                Else
                    .ErrorProvider2.Clear()
                    .fg_A_min = 0
                    Tp_OAI.InspTapeSeal.A_min = CType(.txt_A_min.Text, Decimal)
                    regSubKey.SetValue("Seal_A_min", .txt_A_min.Text, RegistryValueKind.String)
                End If
            ElseIf sender.Equals(.txt_B_max) Then
                If Not IsNumeric(.txt_B_max.Text) Then
                    .ErrorProvider3.SetError(.txt_B_max, "Must be a numeric value!")
                    .fg_B_max = 1
                Else
                    .ErrorProvider3.Clear()
                    .fg_B_max = 0
                    Tp_OAI.InspTapeSeal.B_max = CType(.txt_B_max.Text, Decimal)
                    regSubKey.SetValue("Seal_B_max", .txt_B_max.Text, RegistryValueKind.String)
                End If
            ElseIf sender.Equals(.txt_B_min) Then
                If Not IsNumeric(.txt_B_min.Text) Then
                    .ErrorProvider4.SetError(.txt_B_min, "Must be a numeric value!")
                    .fg_B_min = 1
                Else
                    .ErrorProvider4.Clear()
                    .fg_B_min = 0
                    Tp_OAI.InspTapeSeal.B_min = CType(.txt_B_min.Text, Decimal)
                    regSubKey.SetValue("Seal_B_min", .txt_B_min.Text, RegistryValueKind.String)
                End If
            End If
        End With

    End Sub

    Private Sub chk_CheckSealSetting_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chk_CheckSealSetting.CheckStateChanged

        With Me
            Tp_OAI.InspTapeSeal.Mode = IIf(.chk_CheckSealSetting.CheckState = CheckState.Checked, 1, 0)
            regSubKey.SetValue("Check_TapeSeal_Mode", IIf(.chk_CheckSealSetting.CheckState = CheckState.Checked, "1", "0"), RegistryValueKind.String)
        End With

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim RetVal As String = String.Empty
        FZ_SerCmd("SCENE 00", RetVal)

    End Sub

    Private Sub cbo_StickerColor_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbo_StickerColor.SelectedIndexChanged

        With Me
            Tp_OAI.DefectStickerColor = cbo_StickerColor.SelectedIndex
            regSubKey.SetValue("DefectStickerColor", cbo_StickerColor.SelectedIndex.ToString, RegistryValueKind.String)
        End With

    End Sub

    Private Sub btn_AdvanceSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_AdvanceSet.Click, AdvanceSettingToolStripMenuItem.Click

        frm_AdvSetting.ShowDialog()

        With Me
            With .cbo_Product
                .Items.Clear()
                .Sorted = False

                For Each Element In ProdID
                    Application.DoEvents()
                    .Items.Add(Element)
                Next

                .SelectedIndex = 0
            End With
        End With

    End Sub

    Private Sub cbo_Product_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbo_Product.SelectedIndexChanged

        With Me
            With .cbo_Product
                If .SelectedIndex = 3 Then
                    With Me.txt_DataBlock1
                        .Text = ""
                        .Enabled = False
                    End With

                    With Me.txt_DataBlock2
                        .SelectAll()
                        .Focus()
                    End With
                Else
                    With Me.txt_DataBlock1
                        .Enabled = True
                        .Text = "2400P"
                        .SelectAll()
                        .Focus()
                    End With
                End If
            End With
        End With

    End Sub

    Private Sub AppReStart()

        'RemoveHandler System.Deployment.Application.ApplicationDeployment.CurrentDeployment.UpdateCompleted, AddressOf Me.AppReStart
        'MessageBox.Show("The Application has completely update.", "Network Deployment...", MessageBoxButtons.OK, MessageBoxIcon.Information)

        regSubKey.SetValue("NetworkUpdate", "1")
        'Me.lbl_SWUpdate.Visible = False
        Me.fg_BootError = 0
        Application.Restart()

    End Sub

    Private Sub btn_Home_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Home.Click

        Dim FuncRet As Integer = 0


        With Tp_OAI
            FuncRet = thd_MoveOrg(TP_OAI_Axis_Z, cls_MotionCtrl.MotionDirection.MTN_CCW)

            If FuncRet = Func_Ret_Success Then
                If Not .MotionSys(TP_OAI_Axis_Z).Offset = 0 Then
                    FuncRet = thd_MoveStep(TP_OAI_Axis_Z, .MotionSys(TP_OAI_Axis_Z).Offset)
                End If
            End If
        End With

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click


        Dim RET As String = String.Empty


        Me.FZ_SerCmd("UNITDATA 23 1163", RET)


    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class