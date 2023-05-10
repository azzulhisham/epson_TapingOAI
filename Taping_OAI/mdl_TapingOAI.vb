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

Imports System
Imports System.Math
Imports System.Threading
Imports System.IO
Imports System.IO.Ports
Imports System.Management
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Imports System.Diagnostics
Imports System.ComponentModel


Module mdl_TapingOAI

    Public Const sqlServer As String = "172.16.59.254\SQLEXPRESS"
    Public Const sqlName As String = "Marking"
    Public Const sqluid As String = "VB-SQL"
    Public Const sqlpwd As String = "Anyn0m0us"

    Public Const Func_Ret_Success = 0
    Public Const Func_Ret_Fail = -1

    Public Const fg_OFF = 0
    Public Const fg_ON = 1

    Public Const ContWkJmpMax As Integer = 5
    Public Const WkCdChgRng As Integer = 5

    Public Enum SuccessError
        retError = -1
        Terminated = 0
        retSuccess = 1
    End Enum

    Public Const TP_OAI_Axis_X = 0
    Public Const TP_OAI_Axis_Y = 1
    Public Const TP_OAI_Axis_P = 2
    Public Const TP_OAI_Axis_Z = 0

    Public Enum AxisNumber
        Axis_X = 0
        Axis_Y = 1
        Axis_P = 2
        Axis_Z = 3
    End Enum


    Public Enum MotionThreadEvt
        evt_NotInit = -1
        evt_MoveOrg = 0
        evt_MovePTP
        evt_AutoRun
        evt_Return
        evt_Busy
        evt_Rdy
        evt_Succeed
        evt_Jog
        evt_Quit
    End Enum

    Public Enum MotionType
        Absolute = 0
        Increament
    End Enum

    Public Enum SysAppMode
        app_Auto = 0
        app_Manu
        app_Setting
        app_NotInit
        app_AutoRun
        app_sysError
    End enum

    Public Structure DB
        Public FileName As String
        Public Path As String
    End Structure

    Public Structure CommPortData
        Public PortName As String
        Public DataBits As Integer
        Public BaudRate As Integer
        Public StopBits As System.IO.Ports.StopBits
        Public Parity As System.IO.Ports.Parity
    End Structure

    Public Structure MotionDev
        Public DevH As Integer
        Public HWD_Resolution As Integer
        Public MotionStatus As Integer
        Public BaseClock As Integer
        Public Offset As Integer

        Public CurPosition As Decimal
        Public MoveData As Decimal
        Public ScrewSize As Decimal

        Public MotionEvent As MotionThreadEvt
        Public MotionTask As MotionThreadEvt
        Public NormalMotionSetting As cls_MotionCtrl.MTRMOTION
        Public HomingMotionSetting As cls_MotionCtrl.MTRMOTION
        Public HomingMotionSetting_ex As cls_MotionCtrl.MTRMOTION
    End Structure

    Public Structure SysError
        Public str_ErrorCode As String
        Public str_ErrorDesc As String
        Public str_ErrorToChk As String
    End Structure

    Public Structure ErrorResetState
        Public ErrorCode As Integer
        Public AppMode As SysAppMode
    End Structure


    Public Structure hw_Confg
        '---  DIO Declaration ---
        Public DIO_0 As cls_DIO2727

        Public OutPort1 As cls_DIO_Port
        Public OutPort2 As cls_DIO_Port
        Public OutPort3 As cls_DIO_Port
        Public OutPort4 As cls_DIO_Port

        Public pb_EMG As cls_ioBits
        Public pb_START As cls_ioBits
        Public pb_STOP As cls_ioBits
        Public pb_Auto As cls_ioBits
        Public pb_Manu As cls_ioBits
        Public CB_Alarm As cls_ioBits
        Public M2_Alarm As cls_ioBits
        Public RL As cls_ioBits

        Public RR As cls_ioBits
        Public S_Cover As cls_ioBits
        Public FZ_Run As cls_ioBits
        Public nc_1 As cls_ioBits
        Public nc_2 As cls_ioBits
        Public nc_3 As cls_ioBits
        Public nc_4 As cls_ioBits
        Public nc_5 As cls_ioBits

        Public Auto_LD As cls_ioBits
        Public START_LD As cls_ioBits
        Public STOP_LD As cls_ioBits
        Public BZ As cls_ioBits
        Public PL_R As cls_ioBits
        Public PL_Y As cls_ioBits
        Public PL_G As cls_ioBits
        Public KVM_PC As cls_ioBits

        Public KVM_VZ As cls_ioBits
        Public M2_CW As cls_ioBits
        Public M2_CCW As cls_ioBits
        Public M2_BK As cls_ioBits
        Public M1_FR As cls_ioBits
        Public M2_FHS As cls_ioBits
        Public FZ_DI0 As cls_ioBits
        Public FZ_DI1 As cls_ioBits
    End Structure

    Public Structure Server
        Public ServerName As String
        Public DefaultDir As String
        Public State As Integer
    End Structure

    Public Structure QtyCtrl
        Public P_LotNo As String
        Public WeekCode As String
        Public QtyUsed As Integer
    End Structure

    Public Structure InspData
        Public MarkData1() As String
        Public MarkData2() As String
        Public MFCSpec() As String
        Public Version As String
        Public IMI As String
        Public Prod As String
        Public LotNo As String
        Public P_LotNo As String
        Public Acceptance As String
        Public Freq As String
        Public Insp As String
        Public GUID As String
        Public InspType As String
        Public ReDoLoc As String
        Public RedoInsp As Boolean

        Public InspCnt As Integer
        Public InspCnt_NG As Integer
        Public WeekCodeJump As Integer
        Public BuzzCnt As Integer
        Public Empty As Integer

        Public InspDate As Date
        Public InspStart As Date

        Public P_Lot() As QtyCtrl
        Public P_Lot_r() As QtyCtrl
        Public pLotRun As String
    End Structure

    Public Structure InspRec
        Public GUID_No As String
        Public LotNo As String
        Public P_LotNo As String
        Public IMI As String
        Public InspOpt As String
        Public InspMC As String

        Public InspType As Integer
        Public InspCnt As Integer
        Public InspCnt_NG As Integer
        Public WeekCodeJump As Integer
        Public Empty As Integer

        Public InspDate As Date
        Public RecDate As Date
    End Structure

    Public Structure Rec
        Public Lot_No As String
        Public IMI_No As String
        Public FreqVal As String
        Public Opt As String
        Public RecDate As String
        Public Profile As String
        Public CtrlNo As String
        Public MacNo As String
        Public MData1 As String
        Public MData2 As String
        Public MData3 As String
        Public MData4 As String
        Public MData5 As String
        Public MData6 As String
    End Structure

    Public Structure TapeSealMeas
        Public A_min As Decimal
        Public A_max As Decimal
        Public B_min As Decimal
        Public B_max As Decimal
        Public Mode As Integer
    End Structure

    Public Structure SystemConfg
        Public MC_No As String
        Public DataPath As String
        Public LotDataPath As String
        Public SysTempPath As String
        Public P_Lot_No As String
        Public Ims_SpecNo As String

        Public Authentication As String
        Public AuthenticalCode As String
        Public SpecFileLocation As String
        Public RedoLotTmpRecLoc As String
        Public RedoLoc_ As String

        Public BGSmaxSet As String
        Public BGSminSet As String
        Public DensityLvl As String
        Public FilterType As String

        Public RedoFlag As Boolean

        Public AutoSeqNo As Integer
        Public TimeEllapse As Integer
        Public DefectStickerColor As Integer
        Public SelectedSceneNo As Integer
        Public RT_Error As Integer
        Public ContWkJmp As Integer
        Public ColerationChk As Single

        Public WC_Error As Boolean

        Public IO As hw_Confg
        Public ErrorDisp() As SysError
        Public ErrorReset As ErrorResetState
        Public MotionSys() As MotionDev
        Public Mode As SysAppMode
        Public Database() As DB
        Public SysData() As InspData
        Public InspResult() As InspRec
        Public pLotQty() As QtyCtrl
        Public InspTapeSeal As TapeSealMeas

        Public OAI_CAM1 As CommPortData
        Public ftpServer As Server
    End Structure


    Public GrnLED_OnOff(1) As Color
    Public RedLED_OnOff(1) As Color
    Public OrgLED_OnOff(1) As Color
    Public BlueLED_OnOff(1) As Color
    Public CrayLED_OnOff(1) As Color

    Public Tp_OAI As SystemConfg

    Public fg_Dbg As Integer
    Public fg_Qty As Integer

    Public MyWeekDay() As String = {"", "Monday", "Tuesday", "Wednesday", "Thurday", "Friday", "Saturday", "Sunday"}
    Public MyMonth() As String = {"", "Jan", "Feb", "Mac", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
    Public WeekDay() As String = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}

    Public Motion_Axis() As String = {"FBIMC1", "FBIMC2"}
    Public ProdID() As String
    Public SceneNoDB() As String
    Public n_Pulse() As Integer = {1, 2, 3, 4, 5, 10, 20, 30, 40, 50}

    Public regKey As RegistryKey = Registry.CurrentUser
    Public regSubKey As RegistryKey = regKey.CreateSubKey("Software\az_IOLogics\TapingOAI")


    Public Function ApplicationClose() As Integer

        With Tp_OAI
            With .IO
                .pb_EMG = Nothing
                .pb_START = Nothing
                .pb_STOP = Nothing
                .pb_Auto = Nothing
                .pb_Manu = Nothing
                .FZ_Run = Nothing
                .RL = Nothing
                .RR = Nothing

                .M2_Alarm = Nothing
                .CB_Alarm = Nothing
                .S_Cover = Nothing
                .nc_1 = Nothing
                .nc_2 = Nothing
                .nc_3 = Nothing
                .nc_4 = Nothing
                .nc_5 = Nothing

                .START_LD = Nothing
                .STOP_LD = Nothing
                .Auto_LD = Nothing
                .BZ = Nothing
                .FZ_DI0 = Nothing
                .FZ_DI1 = Nothing
                .KVM_PC = Nothing
                .KVM_VZ = Nothing

                .M1_FR = Nothing
                .M2_BK = Nothing
                .M2_CCW = Nothing
                .M2_CW = Nothing
                .M2_FHS = Nothing
                .PL_G = Nothing
                .PL_R = Nothing
                .PL_Y = Nothing

                .OutPort1 = Nothing
                .OutPort2 = Nothing
                .OutPort3 = Nothing
                .OutPort4 = Nothing

                .DIO_0.DIO_Close()
                .DIO_0 = Nothing
            End With
        End With

        System.GC.Collect()

    End Function

    Public Function ReadRegMotionSetting() As Integer

        Dim regSubKey_ As RegistryKey = regKey.CreateSubKey("Software\az_IOLogics\TapingOAI\Motion")
        Dim regSubKey_Hm As RegistryKey = regKey.CreateSubKey("Software\az_IOLogics\TapingOAI\Motion\Homing")
        Dim regSubKey_Rn As RegistryKey = regKey.CreateSubKey("Software\az_IOLogics\TapingOAI\Motion\Running")
        Dim str_Dmy As String = String.Empty


        With Tp_OAI
            For iLp As Integer = 0 To Motion_Axis.GetUpperBound(0)
                Application.DoEvents()

                With .MotionSys(iLp)
                    .Offset = CType(regSubKey_.GetValue("OffSet_" & iLp.ToString.Trim, "0"), Integer)
                End With

                With .MotionSys(iLp).HomingMotionSetting
                    .dwLowSpeed = CType(regSubKey_Hm.GetValue("LowSpeed_" & iLp.ToString.Trim, "100"), Integer)
                    .dwSpeed = CType(regSubKey_Hm.GetValue("Speed_" & iLp.ToString.Trim, "2000"), Integer)
                    .dwAcc = CType(regSubKey_Hm.GetValue("Acc_" & iLp.ToString.Trim, "100"), Integer)
                    .dwDec = CType(regSubKey_Hm.GetValue("Dec_" & iLp.ToString.Trim, "100"), Integer)
                    .dwSSpeed = CType(regSubKey_Hm.GetValue("SSpeed_" & iLp.ToString.Trim, "0"), Integer)
                    .nStep = CType(regSubKey_Hm.GetValue("nStep_" & iLp.ToString.Trim, cls_MotionCtrl.MTR_CW), Integer)
                End With

                With .MotionSys(iLp).NormalMotionSetting
                    .dwLowSpeed = CType(regSubKey_Rn.GetValue("LowSpeed_" & iLp.ToString.Trim, "1000"), Integer)
                    .dwSpeed = CType(regSubKey_Rn.GetValue("Speed_" & iLp.ToString.Trim, "20000"), Integer)
                    .dwAcc = CType(regSubKey_Rn.GetValue("Acc_" & iLp.ToString.Trim, "300"), Integer)
                    .dwDec = CType(regSubKey_Rn.GetValue("Dec_" & iLp.ToString.Trim, "300"), Integer)
                    .dwSSpeed = CType(regSubKey_Rn.GetValue("SSpeed_" & iLp.ToString.Trim, "0"), Integer)
                    .nStep = CType(regSubKey_Rn.GetValue("nStep_" & iLp.ToString.Trim, cls_MotionCtrl.MTR_CW), Integer)
                End With
            Next
        End With

    End Function

    Public Function GetMotionHnd() As Integer

        Dim int_RetHnd As Integer = -1
        Dim int_RetVal As Integer = Func_Ret_Success


        With Tp_OAI
            For iLp As Integer = 0 To Motion_Axis.GetUpperBound(0)
                Application.DoEvents()
                .MotionSys(iLp).DevH = -1
            Next

            For iLp As Integer = 0 To Motion_Axis.GetUpperBound(0)
                Application.DoEvents()
                int_RetHnd = OpenMotionDev(iLp)

                If int_RetHnd = cls_MotionCtrl.MTR_INVALID_HANDLE_VALUE Then
                    Return Func_Ret_Fail * iLp - 1
                Else
                    .MotionSys(iLp).DevH = int_RetHnd
                End If
            Next

        End With


        Return int_RetVal

    End Function

    Public Function OpenMotionDev(ByVal AxisNumber As Integer) As Integer

        Try
            Return cls_MotionCtrl.MtrOpen(Motion_Axis(AxisNumber), cls_MotionCtrl.MotionOpenFlag.MTN_FLAG_NORMAL)
        Catch ex As Exception
            Return Func_Ret_Fail
        End Try

    End Function

    Public Function CloseMotionDev(ByVal AxisNumber As Integer) As Integer

        With Tp_OAI
            Return cls_MotionCtrl.MtrClose(.MotionSys(AxisNumber).DevH)
        End With

    End Function

    Public Function SetMotionBaseClock(ByVal AxisNumber As Integer) As Integer

        Dim DevInfo As cls_MotionCtrl.MTRDEVICE


        With Tp_OAI
            If cls_MotionCtrl.MtrGetDeviceInfo(.MotionSys(AxisNumber).DevH, DevInfo) = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                If Not DevInfo.dwType = 7204 And Not DevInfo.dwType = 7208 Then
                    Return cls_MotionCtrl.MtrSetBaseClock(.MotionSys(AxisNumber).DevH, 99)
                End If
            Else
                Return Func_Ret_Fail
            End If
        End With

    End Function

    Public Function SetMotionInterlockOff(ByVal AxisNumber As Integer) As Integer

        With Tp_OAI
            Return cls_MotionCtrl.MtrOffInterLock(.MotionSys(AxisNumber).DevH)
        End With

    End Function

    Public Function SetMotionPulseOut(ByVal AxisNumber As Integer, ByVal int_Mode As cls_MotionCtrl.MotionPulseOut) As Integer

        With Tp_OAI
            Return cls_MotionCtrl.MtrSetPulseOut(.MotionSys(AxisNumber).DevH, int_Mode, &H8)
        End With

    End Function



    Public Function InitHardWare() As Integer

        Dim iRetVal As Integer = GetMotionHnd()
        Dim ioBoardCnt As Integer = 0


        If iRetVal = Func_Ret_Success Then
            For iLp As Integer = 0 To Motion_Axis.GetUpperBound(0)
                Application.DoEvents()

                If Not SetMotionBaseClock(iLp) = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                    MessageBox.Show("Unabled to set Base Clock of Axis (" & iLp.ToString & "). Please confirm the Motion Controller had installed properly.", "TapingOAI - System Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Func_Ret_Fail
                End If

                If Not SetMotionInterlockOff(iLp) = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                    MessageBox.Show("Unabled to turn off software interlock of Axis (" & iLp.ToString & "). Please confirm the Motion Controller had installed properly.", "TapingOAI - System Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Func_Ret_Fail
                End If

                If Not SetMotionPulseOut(iLp, cls_MotionCtrl.MotionPulseOut.MTN_METHOD) = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                    MessageBox.Show("Unabled to set the method of output pulse on Axis (" & iLp.ToString & "). Please confirm the Motion Controller had installed properly.", "TapingOAI - System Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return Func_Ret_Fail
                End If
            Next
        Else
            MessageBox.Show("Fail on getting Motion Device Handle. Please refer to the system engineer.", "TapingOAI - System Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Func_Ret_Fail
        End If


        With Tp_OAI.IO
            .DIO_0 = New cls_DIO2727("FBIDIO1", 0, "Interface")
            ioBoardCnt = cls_DIO2727.Count

            If .DIO_0.BoardHnd > 0 Then
                .OutPort1 = New cls_DIO_Port("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, cls_DIO_Port.PortDirection.Out_Port)
                .OutPort2 = New cls_DIO_Port("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, cls_DIO_Port.PortDirection.Out_Port)

                .OutPort1.NewPortValue(0)
                .OutPort2.NewPortValue(0)

                .pb_EMG = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 1, cls_ioBits.BitDirection.In_Bit)
                .pb_START = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 2, cls_ioBits.BitDirection.In_Bit)
                .pb_STOP = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 3, cls_ioBits.BitDirection.In_Bit)
                .pb_Auto = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 4, cls_ioBits.BitDirection.In_Bit)
                .pb_Manu = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 5, cls_ioBits.BitDirection.In_Bit)
                .CB_Alarm = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 6, cls_ioBits.BitDirection.In_Bit)
                .M2_Alarm = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 7, cls_ioBits.BitDirection.In_Bit)
                .RL = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 8, cls_ioBits.BitDirection.In_Bit)

                .RR = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 1, cls_ioBits.BitDirection.In_Bit)
                .S_Cover = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 2, cls_ioBits.BitDirection.In_Bit)
                .FZ_Run = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 3, cls_ioBits.BitDirection.In_Bit)
                .nc_1 = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 4, cls_ioBits.BitDirection.In_Bit)
                .nc_2 = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 5, cls_ioBits.BitDirection.In_Bit)
                .nc_3 = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 6, cls_ioBits.BitDirection.In_Bit)
                .nc_4 = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 7, cls_ioBits.BitDirection.In_Bit)
                .nc_5 = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 8, cls_ioBits.BitDirection.In_Bit)

                .Auto_LD = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 1, cls_ioBits.BitDirection.Out_Bit)
                .START_LD = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 2, cls_ioBits.BitDirection.Out_Bit)
                .STOP_LD = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 3, cls_ioBits.BitDirection.Out_Bit)
                .BZ = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 4, cls_ioBits.BitDirection.Out_Bit)
                .PL_R = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 5, cls_ioBits.BitDirection.Out_Bit)
                .PL_Y = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 6, cls_ioBits.BitDirection.Out_Bit)
                .PL_G = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 7, cls_ioBits.BitDirection.Out_Bit)
                .KVM_PC = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT1_8, 8, cls_ioBits.BitDirection.Out_Bit)

                .KVM_VZ = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 1, cls_ioBits.BitDirection.Out_Bit)
                .M2_CW = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 2, cls_ioBits.BitDirection.Out_Bit)
                .M2_CCW = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 3, cls_ioBits.BitDirection.Out_Bit)
                .M2_BK = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 4, cls_ioBits.BitDirection.Out_Bit)
                .M1_FR = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 5, cls_ioBits.BitDirection.Out_Bit)
                .M2_FHS = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 6, cls_ioBits.BitDirection.Out_Bit)
                .FZ_DI0 = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 7, cls_ioBits.BitDirection.Out_Bit)
                .FZ_DI1 = New cls_ioBits("FBIDIO1", 0, "Interface", .DIO_0.BoardHnd, FBIDIO_OUT9_16, 8, cls_ioBits.BitDirection.Out_Bit)

                iRetVal = Func_Ret_Success
            Else
                iRetVal = Func_Ret_Fail
            End If

            Return iRetVal
        End With

    End Function

    Public Function thd_MoveOrg(ByVal AxisNumber As Integer, ByVal Direction As cls_MotionCtrl.MotionDirection) As Integer

        Dim int_RetVal As Integer = Func_Ret_Success
        Dim int_Wait As Integer = 0


        'Motion Structure (MTRMOTION structure)
        With Tp_OAI.MotionSys(AxisNumber).HomingMotionSetting
            .dwMode = cls_MotionCtrl.MTR_ACC_NORMAL
            .nStep = Direction
            .nReserved = 0
        End With


        With Tp_OAI
            int_RetVal = cls_MotionCtrl.MtrGetStatus(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionGetStatus.MTN_LIMIT_STATUS, .MotionSys(AxisNumber).MotionStatus)
            .MotionSys(AxisNumber).MotionEvent = MotionThreadEvt.evt_MoveOrg

            If int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                If AxisNumber = TP_OAI_Axis_Z Then
                    .IO.M1_FR.Trigger_OFF()
                    .MotionSys(AxisNumber).MotionStatus = .MotionSys(AxisNumber).MotionStatus And &H20

                    If .MotionSys(AxisNumber).MotionStatus = &H20 Then
                        Do
                            int_RetVal = thd_MoveStep(AxisNumber, 1)

                            If .IO.pb_EMG.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                                StopAxisMotion(AxisNumber)
                                Return Func_Ret_Fail
                            End If

                            int_RetVal = cls_MotionCtrl.MtrGetStatus(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionGetStatus.MTN_LIMIT_STATUS, .MotionSys(AxisNumber).MotionStatus)
                            .MotionSys(AxisNumber).MotionStatus = .MotionSys(AxisNumber).MotionStatus And &H20
                        Loop While .MotionSys(AxisNumber).MotionStatus = &H20
                    End If
                Else
                    .MotionSys(AxisNumber).MotionStatus = .MotionSys(AxisNumber).MotionStatus And IIf((Direction = cls_MotionCtrl.MotionDirection.MTN_CW), &H1, &H2)

                    If .MotionSys(AxisNumber).MotionStatus = IIf((Direction = cls_MotionCtrl.MotionDirection.MTN_CW), &H1, &H2) Or .MotionSys(AxisNumber).CurPosition <= 0 Then
                        int_RetVal = thd_MoveStep(AxisNumber, IIf((.MotionSys(AxisNumber).CurPosition <= 0), 28, 20))
                        If Not int_RetVal = Func_Ret_Success Then Return Func_Ret_Fail

                        int_Wait = My.Computer.Clock.TickCount

                        Do While My.Computer.Clock.TickCount < int_Wait + 180
                            Application.DoEvents()

                            If .IO.pb_EMG.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                                StopAxisMotion(AxisNumber)
                                Return Func_Ret_Fail
                            End If
                        Loop
                    End If
                End If
            Else
                Return Func_Ret_Fail
            End If


            int_RetVal = cls_MotionCtrl.MtrSetLimitConfig(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.LimitConfig.MTN_SIGNAL_FILTER, cls_MotionCtrl.MTR_ON)

            If Not int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                Return Func_Ret_Fail
            End If

            int_RetVal = cls_MotionCtrl.MtrSetLimitConfig(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.LimitConfig.MTN_ORG_FUNC, cls_MotionCtrl.MTR_ORG_EZ_STOP)

            If Not int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                Return Func_Ret_Fail
            End If

            int_RetVal = cls_MotionCtrl.MtrSetLimitConfig(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.LimitConfig.MTN_SD_FUNC, cls_MotionCtrl.MTR_CHANGE_SD_SPEED)

            If Not int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                Return Func_Ret_Fail
            End If

            int_RetVal = cls_MotionCtrl.MtrSetMotion(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionStartMotion.MTN_ORG, .MotionSys(AxisNumber).HomingMotionSetting)

            If Not int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                Return Func_Ret_Fail
            End If

            .MotionSys(AxisNumber).MotionEvent = MotionThreadEvt.evt_Busy
            int_RetVal = cls_MotionCtrl.MtrStartMotion(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionStartMotion.MTN_ORG)

            If Not int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                Return Func_Ret_Fail
            End If

            If Not AxisNumber = TP_OAI_Axis_Z Then
                int_Wait = My.Computer.Clock.TickCount

                Do
                    Application.DoEvents()

                    If .IO.pb_EMG.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                        StopAxisMotion(AxisNumber)
                        Return Func_Ret_Fail
                    End If

                    int_RetVal = cls_MotionCtrl.MtrGetStatus(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionGetStatus.MTN_LIMIT_STATUS, .MotionSys(AxisNumber).MotionStatus)

                    If int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                        .MotionSys(AxisNumber).MotionStatus = .MotionSys(AxisNumber).MotionStatus And IIf((Direction = cls_MotionCtrl.MotionDirection.MTN_CW), &H1, &H2)
                    End If

                    If My.Computer.Clock.TickCount > int_Wait + 15000 Then
                        cls_MotionCtrl.MtrStopMotion(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionStopMotion.MTN_IMMEDIATE_STOP)
                        Return Func_Ret_Fail
                    End If
                Loop While Not .MotionSys(AxisNumber).MotionStatus = IIf((Direction = cls_MotionCtrl.MotionDirection.MTN_CW), &H1, &H2)
            End If

            int_RetVal = cls_MotionCtrl.MtrSetLimitConfig(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.LimitConfig.MTN_ORG_FUNC, cls_MotionCtrl.MTR_ORG_STOP)

            If Not int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                cls_MotionCtrl.MtrStopMotion(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionStopMotion.MTN_IMMEDIATE_STOP)
                Return Func_Ret_Fail
            End If

            int_Wait = My.Computer.Clock.TickCount

            Do
                Application.DoEvents()

                If .IO.pb_EMG.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                    StopAxisMotion(AxisNumber)
                    Return Func_Ret_Fail
                End If

                int_RetVal = cls_MotionCtrl.MtrGetStatus(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionGetStatus.MTN_LIMIT_STATUS, .MotionSys(AxisNumber).MotionStatus)

                If int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                    .MotionSys(AxisNumber).MotionStatus = .MotionSys(AxisNumber).MotionStatus And &H20
                End If

                If My.Computer.Clock.TickCount > int_Wait + 10000 Then
                    cls_MotionCtrl.MtrStopMotion(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionStopMotion.MTN_IMMEDIATE_STOP)
                    Return Func_Ret_Fail
                End If
            Loop While Not .MotionSys(AxisNumber).MotionStatus = &H20


            Do
                Application.DoEvents()

                If .IO.pb_EMG.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                    StopAxisMotion(AxisNumber)
                    Return Func_Ret_Fail
                End If

                int_RetVal = cls_MotionCtrl.MtrGetStatus(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionGetStatus.MTN_BUSY, .MotionSys(AxisNumber).MotionStatus)

                If int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                    .MotionSys(AxisNumber).MotionStatus = .MotionSys(AxisNumber).MotionStatus And &H1
                End If

                If My.Computer.Clock.TickCount > int_Wait + 10000 Then
                    cls_MotionCtrl.MtrStopMotion(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionStopMotion.MTN_IMMEDIATE_STOP)
                    Return Func_Ret_Fail
                End If
            Loop While Not .MotionSys(AxisNumber).MotionStatus = &H0


            .MotionSys(AxisNumber).CurPosition = 0
            int_RetVal = cls_MotionCtrl.MtrSetLimitConfig(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.LimitConfig.MTN_SD_FUNC, cls_MotionCtrl.MTR_SD_OFF)

            If Not int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                .MotionSys(AxisNumber).MotionEvent = MotionThreadEvt.evt_NotInit
                Return Func_Ret_Fail
            Else
                .MotionSys(AxisNumber).MotionEvent = MotionThreadEvt.evt_Rdy
            End If
        End With

    End Function

    Public Function StopAxisMotion(ByVal AxisNumber As Integer) As Integer

        Dim int_RetVal As Integer = Func_Ret_Success


        With Tp_OAI.MotionSys(AxisNumber)
            int_RetVal = cls_MotionCtrl.MtrStopMotion(.DevH, cls_MotionCtrl.MotionStopMotion.MTN_IMMEDIATE_STOP)

            If Not int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                Return Func_Ret_Fail
            End If
        End With

    End Function

    Public Function thd_MoveStep(ByVal AxisNumber As Integer, ByVal DistanceInMM As Decimal, Optional ByVal MoveType As MotionType = MotionType.Absolute) As Integer

        Dim int_RetVal As Integer = Func_Ret_Success
        Dim int_Wait As Integer = 0

        Dim sng_MoveStep As Decimal = 0
        Dim str_Dmy As String = String.Empty


        With Tp_OAI
            If Not .MotionSys(AxisNumber).MotionEvent = MotionThreadEvt.evt_MoveOrg Then
                .MotionSys(AxisNumber).MotionEvent = MotionThreadEvt.evt_Busy
            End If

            If MoveType = MotionType.Absolute Then
                sng_MoveStep = DistanceInMM
                'sng_MoveStep = DistanceInMM - .MotionSys(AxisNumber).CurPosition
                '.CurPosition = .CurPosition + sng_MoveStep

                'str_Dmy = String.Format("{0:f3}", sng_MoveStep)
                'sng_MoveStep = Val(str_Dmy)
            Else
                sng_MoveStep = DistanceInMM
                '.CurPosition = .CurPosition + DistanceInMM
            End If

            .MotionSys(AxisNumber).NormalMotionSetting.dwMode = cls_MotionCtrl.MTR_ACC_NORMAL

            If AxisNumber = TP_OAI_Axis_Z Then
                .MotionSys(AxisNumber).NormalMotionSetting.nStep = DistanceInMM
            Else
                .MotionSys(AxisNumber).NormalMotionSetting.nStep = CType((sng_MoveStep * Tp_OAI.MotionSys(AxisNumber).HWD_Resolution) / .MotionSys(AxisNumber).ScrewSize, Integer)
            End If


            .MotionSys(AxisNumber).NormalMotionSetting.nReserved = 0

            If .MotionSys(AxisNumber).NormalMotionSetting.nStep = 0 Then
                .MotionSys(AxisNumber).MotionEvent = MotionThreadEvt.evt_Rdy
                Return Func_Ret_Success
            End If

            If AxisNumber = TP_OAI_Axis_Z Then
                int_RetVal = cls_MotionCtrl.MtrSetLimitConfig(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.LimitConfig.MTN_ORG_FUNC, cls_MotionCtrl.MTR_ORG_EZ_STOP)

                If Not int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                    Return Func_Ret_Fail
                End If

                int_RetVal = cls_MotionCtrl.MtrSetMotion(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionStartMotion.MTN_PTP, .MotionSys(AxisNumber).NormalMotionSetting)
            Else
                'int_RetVal = cls_MotionCtrl.MtrSetMotion(.DevH, cls_MotionCtrl.MotionStartMotion.MTN_PTP, .NormalMotionSetting)

                If Tp_OAI.Mode = SysAppMode.app_NotInit Then
                    .MotionSys(AxisNumber).HomingMotionSetting_ex = .MotionSys(AxisNumber).HomingMotionSetting
                    .MotionSys(AxisNumber).HomingMotionSetting_ex.nStep = .MotionSys(AxisNumber).NormalMotionSetting.nStep

                    int_RetVal = cls_MotionCtrl.MtrSetMotion(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionStartMotion.MTN_PTP, .MotionSys(AxisNumber).HomingMotionSetting_ex)
                Else
                    int_RetVal = cls_MotionCtrl.MtrSetMotion(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionStartMotion.MTN_PTP, .MotionSys(AxisNumber).NormalMotionSetting)
                End If
            End If


            If Not int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                Return Func_Ret_Fail
            End If

            If AxisNumber = TP_OAI_Axis_Z Then
                int_RetVal = cls_MotionCtrl.MtrStartMotion(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionStartMotion.MTN_PTP)
            Else
                int_RetVal = cls_MotionCtrl.MtrStartMotion(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionStartMotion.MTN_PTP)
            End If

            If Not int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                Return Func_Ret_Fail
            End If


            int_Wait = My.Computer.Clock.TickCount

            'If AxisNumber = TP_OAI_Axis_Z Then
            '    If Not .MotionSys(AxisNumber).MotionEvent = MotionThreadEvt.evt_MoveOrg Then
            '        .MotionSys(AxisNumber).MotionEvent = MotionThreadEvt.evt_Rdy
            '        Return int_RetVal
            '    End If

            '    Do While My.Computer.Clock.TickCount < int_Wait + 280
            '        Application.DoEvents()
            '    Loop

            '    Return int_RetVal
            'End If


            Do
                Application.DoEvents()

                If .IO.pb_EMG.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                    StopAxisMotion(AxisNumber)
                    Return Func_Ret_Fail
                End If

                int_RetVal = cls_MotionCtrl.MtrGetStatus(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionGetStatus.MTN_BUSY, .MotionSys(AxisNumber).MotionStatus)

                If int_RetVal = cls_MotionCtrl.MTR_ERROR_SUCCESS Then
                    .MotionSys(AxisNumber).MotionStatus = .MotionSys(AxisNumber).MotionStatus And &H1
                End If

                If My.Computer.Clock.TickCount > int_Wait + 10000 Then
                    cls_MotionCtrl.MtrStopMotion(.MotionSys(AxisNumber).DevH, cls_MotionCtrl.MotionStopMotion.MTN_IMMEDIATE_STOP)
                    Return Func_Ret_Fail
                End If
            Loop While Not .MotionSys(AxisNumber).MotionStatus = &H0

            If MoveType = MotionType.Absolute Then
                '.MotionSys(AxisNumber).CurPosition = .MotionSys(AxisNumber).CurPosition + sng_MoveStep
            Else
                '.MotionSys(AxisNumber).CurPosition = .MotionSys(AxisNumber).CurPosition + DistanceInMM
            End If

            'str_Dmy = String.Format("{0:f3}", CType(.CurPosition, Single))
            '.CurPosition = Val(str_Dmy)

            If Not .MotionSys(AxisNumber).MotionEvent = MotionThreadEvt.evt_MoveOrg Then
                .MotionSys(AxisNumber).MotionEvent = MotionThreadEvt.evt_Rdy
            End If
        End With

    End Function

    Public Function ShellCmd(ByVal ImageFile As String) As String

        Dim ShlCmd As New System.Diagnostics.Process


        With ShlCmd
            .StartInfo.WorkingDirectory = "c:\ifz"
            .StartInfo.FileName = "c:\\ifz\byrtobmp.exe"
            .StartInfo.Arguments = ImageFile & " /o:" & Chr(34) & "c:\ifz" & Chr(34) & " /f:jpg /h"
            .StartInfo.UseShellExecute = False
            .StartInfo.RedirectStandardInput = True
            .StartInfo.RedirectStandardOutput = True

            '.StartInfo.UserName = "Administrator"
            .StartInfo.CreateNoWindow = True
            .StartInfo.WindowStyle = ProcessWindowStyle.Hidden

            If System.Environment.OSVersion.Version.Major >= 6 Then ' Windows Vista or higher
                .StartInfo.Verb = "runas"
            Else
                ' No need to prompt to run as admin
            End If

            Try
                .Start()
            Catch ex As Exception
                Return "Shell Cmd. Error"
            End Try

            Dim myStreamWriter As StreamWriter = .StandardInput
            myStreamWriter.WriteLine("" & vbCrLf)
            myStreamWriter.WriteLine("Exit" & vbCrLf)

            Dim mystreamreader As StreamReader = .StandardOutput
            Dim retstr As String = mystreamreader.ReadToEnd.ToLower

            .WaitForExit(8000)
            .Close()

            mystreamreader.Close()

            With My.Computer.FileSystem
                If .FileExists(ShlCmd.StartInfo.WorkingDirectory & "\" & ImageFile) Then
                    Try
                        .DeleteFile(ShlCmd.StartInfo.WorkingDirectory & "\" & ImageFile)
                    Catch ex As Exception

                    End Try
                End If
            End With

            Return retstr
        End With

    End Function

    Public Sub ReadRegData()

        Dim regSubKeyComm_1 As RegistryKey = regKey.CreateSubKey("Software\az_IOLogics\TapingOAI\CAM_1")


        With Tp_OAI
            With .Database(0)
                .FileName = regSubKey.GetValue("SourceDataBaseName", "Record.mdb")
                .Path = regSubKey.GetValue("SourceDataBasePath", "\\172.16.59.2\epmmn\CP\PX Line\Laser Marking\Data\Record")
                '.Path = "C:\Data"           'Remark this line when data source is from the network
            End With

            With .Database(1)
                .FileName = regSubKey.GetValue("SysDataBaseName", "TapingOAI.mdb")
                .Path = regSubKey.GetValue("SysDataBasePath", "C:\Data")
            End With

            ProdID = regSubKey.GetValue("ProdID", "FA-20H,TSX-3225,FA-118T,RAKON").ToString.Split(",")
            SceneNoDB = regSubKey.GetValue("SceneNoDB", "0,1,2,3").ToString.Split(",")


            .MC_No = regSubKey.GetValue("MachineNo", "1")
            .DataPath = regSubKey.GetValue("DataPath", "C:\Data")
            .LotDataPath = regSubKey.GetValue("LotDataPath", "C:\Data\LotEndData")
            .SysTempPath = regSubKey.GetValue("SysTempPath", "c:\iFz")

            .Authentication = regSubKey.GetValue("Authentication", "123")
            .SpecFileLocation = regSubKey.GetValue("SpecFileLocation", "\\172.16.59.2\epmmn\Control\CP\PX Line\MI\LaserMarking")
            '.SpecFileLocation = regSubKey.GetValue("SpecFileLocation", "C:\Data\IMI")       'Remark this line when data source is from the network

            .RedoLotTmpRecLoc = regSubKey.GetValue("RedoLotTmpRecLoc", "\\172.16.59.2\epmmn\CP\PX Line\TapingOAI")

            .ftpServer.ServerName = regSubKey.GetValue("ftpServerName", "ftp://10.5.5.100")
            .ftpServer.DefaultDir = regSubKey.GetValue("ftpServerDir", "/RAMDisk/")

            .DefectStickerColor = CType(regSubKey.GetValue("DefectStickerColor", "0"), Integer)
            .Ims_SpecNo = regSubKey.GetValue("Ims_SpecNo", "D1010003:FA-118|D1010003:FA-118")

            With .InspTapeSeal
                .A_min = CType(regSubKey.GetValue("Seal_A_min", "0"), Decimal)
                .A_max = CType(regSubKey.GetValue("Seal_A_max", "0"), Decimal)
                .B_min = CType(regSubKey.GetValue("Seal_B_min", "0"), Decimal)
                .B_max = CType(regSubKey.GetValue("Seal_B_max", "0"), Decimal)
                .Mode = CType(regSubKey.GetValue("Check_TapeSeal_Mode", "1"), Decimal)
            End With

            With .OAI_CAM1
                .PortName = regSubKeyComm_1.GetValue("CommPortName", "COM1")

                .BaudRate = CType(regSubKeyComm_1.GetValue("CommBaudRate"), Integer)
                .BaudRate = IIf((.BaudRate = 0), 115200, .BaudRate)

                .DataBits = CType(regSubKeyComm_1.GetValue("CommDataBits"), Integer)
                .DataBits = IIf((.DataBits = 0), 8, .DataBits)

                .StopBits = CType(regSubKeyComm_1.GetValue("CommStopBits"), System.IO.Ports.StopBits)
                .StopBits = IIf((.StopBits = 0), Ports.StopBits.One, .StopBits)

                .Parity = CType(regSubKeyComm_1.GetValue("CommParity"), System.IO.Ports.Parity)
            End With
        End With

    End Sub

    Public Function InitSerialPort() As Integer

        Dim lng_RetVal As Long = Func_Ret_Fail


        'Initialize Comm. Port For Pick & Place Unit (FZ3)
        Try
            With frm_Main.TP_CAM1
                If .IsOpen = True Then
                    .Close()
                End If

                .PortName = Tp_OAI.OAI_CAM1.PortName
                .Parity = Tp_OAI.OAI_CAM1.Parity
                .BaudRate = Tp_OAI.OAI_CAM1.BaudRate
                .StopBits = Tp_OAI.OAI_CAM1.StopBits
                .DataBits = Tp_OAI.OAI_CAM1.DataBits
                .ReceivedBytesThreshold = 1
                .ReadBufferSize = 1024

                '.Open()
            End With
        Catch
            Return lng_RetVal
        End Try


        Return Func_Ret_Success

    End Function

    Public Function GetConnectedAdapters() As Integer

        Dim SearchConnectedAdapters As Integer = -1


        For Each nic As Net.NetworkInformation.NetworkInterface In Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces
            If nic.OperationalStatus = Net.NetworkInformation.OperationalStatus.Up And Not nic.NetworkInterfaceType = Net.NetworkInformation.NetworkInterfaceType.Loopback Then
                'Debug.Print(String.Format("{0} {1}", nic.Description, nic.GetPhysicalAddress))
                SearchConnectedAdapters += 1
            End If
        Next

        Return SearchConnectedAdapters

    End Function

    Public Function GetFilesList(ByVal PathName As String, ByVal ExtName As String, ByRef allFiles() As String) As String

        Dim filesCount As Integer = -1


        Try
            Dim files = From file In My.Computer.FileSystem.GetFiles(PathName) _
            Order By file _
            Select file

            Dim filesinfo = From file In files _
                    Select My.Computer.FileSystem.GetFileInfo(file)

            Dim SelectInfo = From file In filesinfo _
                    Where file.Extension = ExtName _
                    Select file.FullName

            allFiles = SelectInfo.ToArray
            filesCount = allFiles.GetUpperBound(0)
        Catch ex As Exception
            filesCount = -1
        End Try

        Return filesCount

    End Function

    Public Sub DeleteSysTempFile()

        Dim datFiles() As String = {}


        With Tp_OAI
            If Not GetFilesList(.SysTempPath, ".jpg", datFiles) < 0 Then
                If datFiles.Length > 0 Then
                    For iLp As Integer = 0 To datFiles.GetUpperBound(0)
                        Application.DoEvents()

                        Try
                            My.Computer.FileSystem.DeleteFile(datFiles(iLp))
                        Catch ex As Exception
                            'nop
                        End Try
                    Next
                End If
            End If
        End With

    End Sub

    Public Sub SysDataReset()

        With Tp_OAI
            For iLp As Integer = 0 To .SysData.GetUpperBound(0)
                Application.DoEvents()

                With .SysData(iLp)
                    .LotNo = ""
                    .P_LotNo = ""
                    .Acceptance = ""
                    .Freq = ""
                    .IMI = ""
                    .InspType = ""
                    .ReDoLoc = ""
                    .pLotRun = ""

                    .RedoInsp = False

                    .Empty = 0
                    .WeekCodeJump = 0
                    .InspCnt = 0
                    .InspCnt_NG = 0
                    .BuzzCnt = 0

                    ReDim .P_Lot(0)
                    ReDim .P_Lot_r(0)
                End With
            Next

            .RedoLoc_ = ""
            .ContWkJmp = 0
            .AutoSeqNo = 0
            .TimeEllapse = 0
        End With

    End Sub

End Module
