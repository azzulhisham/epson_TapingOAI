'---------------------------------------------------
'   PX - Taping OAI Development
'===================================================
'   Designed By : Zulhisham Tan
'   Module      : frm_Splash.vb
'   Date        : 08-Mar-2011
'   Version     : 2011.03.001.001
'---------------------------------------------------
'   Copyright (C) 2011-2014 az_IOLogics
'---------------------------------------------------

Imports Microsoft.Win32
Imports System.Threading


Public Class frm_Splash

    Dim KVM_Pos As Integer = 0


    Private Sub FadeAway()

        With Me.tmr_Enter
            .Interval = 30
            .Enabled = True
        End With

    End Sub

    Private Sub btn_Start_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_Start.Click

        With Me
            If btn_Start.Enabled = False Then Exit Sub
            If Not .Tag = "-" Then Exit Sub

            .Tag = "1"
            .lbl_BootMsg.Text = "System is starting up..."
            .btn_Start.Enabled = False
            .tmr_WaitKey.Enabled = False
            .tmr_ReadIO.Enabled = False

            With Tp_OAI.IO
                .START_LD.Trigger_ON()
                .STOP_LD.Trigger_OFF()
                .Auto_LD.Trigger_OFF()
            End With

            .btn_Start.BackColor = GrnLED_OnOff(Tp_OAI.IO.START_LD.BitState)
            .btn_Close.BackColor = OrgLED_OnOff(Tp_OAI.IO.STOP_LD.BitState)
        End With

        FadeAway()

    End Sub

    Private Sub btn_Close_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_Close.Click

        With Me
            If Not .Tag = "-" Then Exit Sub
            .lbl_BootMsg.Text = "System shutting down..."
            .Tag = "0"
            .tmr_WaitKey.Enabled = False
            .tmr_ReadIO.Enabled = False

            If .btn_Start.Visible = True Then
                With Tp_OAI.IO
                    .START_LD.Trigger_OFF()
                    .STOP_LD.Trigger_ON()
                    .Auto_LD.Trigger_OFF()
                End With

                .btn_Start.BackColor = GrnLED_OnOff(Tp_OAI.IO.START_LD.BitState)
                .btn_Close.BackColor = OrgLED_OnOff(Tp_OAI.IO.STOP_LD.BitState)
            End If
        End With

        FadeAway()

    End Sub

    Private Sub frm_Splash_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If frm_Main.IsHandleCreated = True Then End

        With Me
            fg_Dbg = 0
            Application.DoEvents()

            .lbl_AppVersionNo.Text = "Version No. : " & String.Format("{0:D4}.{1:D4}.{2:D3}.{3:D3}", My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build, My.Application.Info.Version.Revision)
            .lbl_BootMsg.Text = "System Initializing..."
            .Tag = "-"

            GrnLED_OnOff(0) = .pic_GrColorSigOff.BackColor
            GrnLED_OnOff(1) = .pic_GrColorSigOn.BackColor
            RedLED_OnOff(0) = .pic_RdColorSigOff.BackColor
            RedLED_OnOff(1) = .pic_RdColorSigOn.BackColor
            OrgLED_OnOff(0) = .pic_OgColorSigOff.BackColor
            OrgLED_OnOff(1) = .pic_OgColorSigOn.BackColor
        End With

        With Tp_OAI
            ReDim .ErrorDisp(100)
            ReDim .MotionSys(1)
            ReDim .SysData(1)
            ReDim .Database(1)
            ReDim .InspResult(1)

            .AutoSeqNo = 0
            SysDataReset()
            DeleteSysTempFile()


            '--- Debug Data Entry Code Page  ---
            'fg_Dbg = 1
            'ReadRegData()
            'frm_DataEntry.ShowDialog()
            'frm_AdvSetting.ShowDialog()
            '-----------------------------------


            With .ErrorDisp(1)
                .str_ErrorCode = "- 1 -"
                .str_ErrorDesc = "Emergency Alert or System Fault..."
                .str_ErrorToChk = "Check Emergency button and Circuit Breaker."
            End With

            With .ErrorDisp(11)
                .str_ErrorCode = "- 11 -"
                .str_ErrorDesc = "System Interlock..."
                .str_ErrorToChk = "Check Sensor (R) !"
            End With

            With .ErrorDisp(12)
                .str_ErrorCode = "- 12 -"
                .str_ErrorDesc = "System Interlock..."
                .str_ErrorToChk = "Check Sensor (L) !"
            End With

            With .ErrorDisp(13)
                .str_ErrorCode = "- 13 -"
                .str_ErrorDesc = "System Interlock..."
                .str_ErrorToChk = "Check Safety Cover! The Cover should closed."
            End With

            With .ErrorDisp(14)
                .str_ErrorCode = "- 14 -"
                .str_ErrorDesc = "System Interlock..."
                .str_ErrorToChk = "Vision system not in Run Mode."
            End With

            With .ErrorDisp(15)
                .str_ErrorCode = "- 15 -"
                .str_ErrorDesc = "System Interlock..."
                .str_ErrorToChk = "Unloader Motor Error!"
            End With


            With .ErrorDisp(16)
                .str_ErrorCode = "- 16 -"
                .str_ErrorDesc = "AutoRun Interlock..."
                .str_ErrorToChk = "CB Motor Error!"
            End With

            With .ErrorDisp(17)
                .str_ErrorCode = "- 17 -"
                .str_ErrorDesc = "System Interlock..."
                .str_ErrorToChk = "Motor Error!"
            End With

            With .ErrorDisp(21)
                .str_ErrorCode = "- 21 -"
                .str_ErrorDesc = "AutoRun Interlock..."
                .str_ErrorToChk = "Check Sensor (R) !"
            End With

            With .ErrorDisp(22)
                .str_ErrorCode = "- 22 -"
                .str_ErrorDesc = "AutoRun Interlock..."
                .str_ErrorToChk = "Check Sensor (L) !"
            End With

            With .ErrorDisp(23)
                .str_ErrorCode = "- 23 -"
                .str_ErrorDesc = "AutoRun Interlock..."
                .str_ErrorToChk = "Check Safety Cover! The Cover should closed."
            End With

            With .ErrorDisp(24)
                .str_ErrorCode = "- 24 -"
                .str_ErrorDesc = "AutoRun Interlock..."
                .str_ErrorToChk = "Vision system not in Run Mode."
            End With

            With .ErrorDisp(25)
                .str_ErrorCode = "- 25 -"
                .str_ErrorDesc = "AutoRun Interlock..."
                .str_ErrorToChk = "Motor Error!"
            End With

            With .ErrorDisp(26)
                .str_ErrorCode = "- 26 -"
                .str_ErrorDesc = "AutoRun Interlock..."
                .str_ErrorToChk = "CB Motor Error!"
            End With

            With .ErrorDisp(27)
                .str_ErrorCode = "- 27 -"
                .str_ErrorDesc = "AutoRun Interrupt..."
                .str_ErrorToChk = "Interlock Device Trigger!"
            End With

            With .ErrorDisp(70)
                .str_ErrorCode = "- 70 -"
                .str_ErrorDesc = "Weekcode Jumping..."
                .str_ErrorToChk = "Weekcode Jumping !!!"
            End With

            With .ErrorDisp(71)
                .str_ErrorCode = "- 71 -"
                .str_ErrorDesc = "Quantity NG..."
                .str_ErrorToChk = "Quantity NG!!!"
            End With

            With .ErrorDisp(90)
                .str_ErrorCode = "- 90 -"
                .str_ErrorDesc = "Inspection NG..."
                .str_ErrorToChk = "One of the Data Block is not match!"
            End With

            With .ErrorDisp(91)
                .str_ErrorCode = "- 91 -"
                .str_ErrorDesc = "Inspection Cancel..."
                .str_ErrorToChk = "Re-Inspect!!!"
            End With

            With .ErrorDisp(92)
                .str_ErrorCode = "- 92 -"
                .str_ErrorDesc = "RGB Data (Sticker) NG..."
                .str_ErrorToChk = "Re-Setting Vision System!!!"
            End With

            With .ErrorDisp(99)
                .str_ErrorCode = "- 99 -"
                .str_ErrorDesc = "Inspection Finish..."
                .str_ErrorToChk = "Change Lot..."
            End With
        End With

    End Sub

    Private Sub tmr_Enter_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr_Enter.Tick

        With Me
            Application.DoEvents()

            If .Opacity = 0 Then
                .tmr_Enter.Enabled = False

                If .Tag = "1" Then
                    .Hide()
                    frm_Main.ShowDialog(Me)
                End If

                .Close()
                ApplicationClose()
                Application.Exit()

                End
                Exit Sub
            End If

            .Opacity -= 0.02
        End With

    End Sub

    Private Sub frm_Splash_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        'Read Motion Moving Condition Setting From Reg.
        ReadRegMotionSetting()

        Dim iRetVal As Integer = InitHardWare()

        If Not iRetVal = Func_Ret_Success Then
            With Me.btn_Close
                .BackColor = OrgLED_OnOff(1)
                .Visible = True
            End With

            Me.lbl_BootMsg.Text = "Hardware Error..."
            Exit Sub
        Else
            With Me
                .lbl_BootMsg.Text = "System is in standby..."

                With .tmr_ReadIO
                    .Interval = 100
                    .Enabled = True
                End With

                .btn_Close.Visible = True
                .btn_Start.Visible = True
            End With
        End If

    End Sub

    Private Sub tmr_ReadIO_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr_ReadIO.Tick

        Static fg_Trg As Integer = 0


        With Tp_OAI.IO
            If Val(regSubKey.GetValue("NetworkUpdate", "0")) <> 0 Then
                regSubKey.SetValue("NetworkUpdate", "0")

                Me.tmr_ReadIO.Enabled = False
                Me.btn_Start.PerformClick()
                Exit Sub
            End If

            If .pb_START.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                Me.tmr_ReadIO.Enabled = False

                Do Until .pb_START.BitState = cls_PCIBoard.BitsState.eBit_OFF
                    Application.DoEvents()
                Loop

                Me.btn_Start.PerformClick()
                Exit Sub
            End If

            If .pb_EMG.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                With Me
                    .tmr_ReadIO.Enabled = False

                    MessageBox.Show("The system is in Emergency mode or the system has no power supplied.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    .btn_Close.PerformClick()
                End With

                Exit Sub
            End If


            If .pb_Auto.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                If Me.KVM_Pos <> 1 Then
                    KVM_Pos = 1
                    .KVM_PC.Trigger_ON()
                    Thread.Sleep(100)
                    .KVM_PC.Trigger_OFF()
                End If
            Else
                If Me.KVM_Pos <> 2 Then
                    KVM_Pos = 2
                    .KVM_VZ.Trigger_ON()
                    Thread.Sleep(100)
                    .KVM_VZ.Trigger_OFF()
                End If
            End If

            'LED Blinking
            fg_Trg += 1

            If fg_Trg = 2 Then
                fg_Trg = 0

                If .START_LD.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                    .START_LD.Trigger_OFF()
                    .Auto_LD.Trigger_OFF()

                    .PL_G.Trigger_ON()
                    .PL_R.Trigger_ON()
                    .PL_Y.Trigger_ON()
                Else
                    .START_LD.Trigger_ON()
                    .Auto_LD.Trigger_ON()

                    .PL_G.Trigger_OFF()
                    .PL_R.Trigger_OFF()
                    .PL_Y.Trigger_OFF()
                End If

                Me.btn_Start.BackColor = GrnLED_OnOff(Tp_OAI.IO.START_LD.BitState)
            End If
        End With

    End Sub

    Private Sub btn_Close_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btn_Close.MouseDown

        If e.Button = Windows.Forms.MouseButtons.Right Then
            Me.btn_Close.ContextMenuStrip.Show()
        End If

    End Sub

    Private Sub sub_ShutdownPC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles sub_ShutdownPC.Click

        Me.btn_Close.PerformClick()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'With Me.FtpConnection1
        '    .AutoLogin = True
        'End With

        'ShellCmd("")
        'Me.PictureBox3.BackgroundImageLayout = ImageLayout.Zoom
        'Me.PictureBox3.BackgroundImage = Image.FromFile("c:\ifz\a2011.jpg")

    End Sub

End Class
