'---------------------------------------------------
'   PX - Taping OAI Development
'===================================================
'   Designed By : Zulhisham Tan
'   Module      : cls_MotionCtrl.vb
'   Date        : 03-Jun-2009
'   Version     : 2009.06.03.001
'---------------------------------------------------
'   Copyright (C) 2007-2009 az_Zulhisham
'---------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.Runtime.InteropServices


'Namespace InterfaceCorpMotionDll


Public Class cls_MotionCtrl

    ' Open
    Public Const MTR_FLAG_NORMAL As Integer = 0
    Public Const MTR_FLAG_OVERLAPPED As Integer = 1

    Public Enum MotionOpenFlag
        MTN_FLAG_NORMAL = 0
        MTN_FLAG_OVERLAPPED
    End Enum

    ' Set/GetBaseClock
    Public Const MTR_CLOCK_1M As Integer = 0
    Public Const MTR_CLOCK_1_4M As Integer = 1
    Public Const MTR_CLOCK_1_16M As Integer = 2

    Public Enum MotionBaseClock
        MTN_CLOCK_1M = 0
        MTN_CLOCK_1_4M
        MTN_CLOCK_1_16M
    End Enum

    ' Set/GetPulseOut
    Public Const MTR_METHOD As Integer = 0
    Public Const MTR_IDLING As Integer = 1
    Public Const MTR_FINISH_FLAG As Integer = 3
    Public Const MTR_0P5PULSE As Integer = 1
    Public Const MTR_1P5PULSE As Integer = 3
    Public Const MTR_WIDTH As Integer = 2
    Public Const MTR_PULSE_OUT As Integer = 0
    Public Const MTR_INP As Integer = 1
    Public Const MTR_PULSE_OFF As Integer = 2

    Public Enum MotionPulseOut
        MTN_METHOD = 0
        MTN_IDLING
        MTN_FINISH_FLAG
    End Enum

    ' Set/GetLimitConfig
    Public Const MTR_MASK As Integer = 0
    Public Const MTR_LOGIC As Integer = 1
    Public Const MTR_SD_FUNC As Integer = 2
    Public Const MTR_SD_ACTIVE As Integer = 3
    Public Const MTR_ORG_FUNC As Integer = 4
    Public Const MTR_ORG_ACTIVE As Integer = 5
    Public Const MTR_ORG_EZ_COUNT As Integer = 6
    Public Const MTR_ALM_FUNC As Integer = 7
    Public Const MTR_SIGNAL_FILTER As Integer = 8
    Public Const MTR_1MICRO As Integer = &H10
    Public Const MTR_10MICRO As Integer = &H11
    Public Const MTR_100MICRO As Integer = &H12
    Public Const MTR_DI As Integer = &H0
    Public Const MTR_LIMIT As Integer = &H100
    Public Const MTR_SYNC_EXT As Integer = &H200

    Public Enum LimitConfig
        MTN_MASK = 0
        MTN_LOGIC
        MTN_SD_FUNC
        MTN_SD_ACTIVE
        MTN_ORG_FUNC
        MTN_ORG_ACTIVE
        MTN_ORG_EZ_COUNT
        MTN_ALM_FUNC
        MTN_SIGNAL_FILTER
    End Enum


    Public Const MTR_CHANGE_SD_SPEED As Integer = 0
    ' outupt pulses at the startup velocity.
    Public Const MTR_DEC_STOP_SIGNAL As Integer = 1
    Public Const MTR_SD_OFF As Integer = 2

    Public Const MTR_SIGNAL_LEVEL As Integer = 0
    Public Const MTR_SIGNAL_EDGE As Integer = 1
    Public Const MTR_SIGNAL_LATCH As Integer = 2

    Public Const MTR_ORG_STOP As Integer = 0
    Public Const MTR_ORG_DEC_EZ_STOP As Integer = 1

    ' then stop it when the specified number of the EZ pulses are counted.
    Public Const MTR_ORG_EZ_STOP As Integer = 2
    '  after the ORG signal is asserted.

    Public Const MTR_ALM_STOP As Integer = 0
    Public Const MTR_ALM_DEC_STOP As Integer = 1
    Public Const MTR_OFF As Integer = 0
    Public Const MTR_ON As Integer = 1

    ' Set/GetCounterConfig
    Public Enum MotionCounterConfig
        MTN_ENCODER_MODE = 0
        MTN_ENCODER_CLEAR
        MTN_COUNTER_CLEAR
    End Enum

    Public Const MTR_ENCODER_MODE As Integer = 0
    Public Const MTR_ENCODER_CLEAR As Integer = 1
    Public Const MTR_COUNTER_CLEAR As Integer = 2

    Public Const MTR_NOT_COUNT As Integer = 0
    Public Const MTR_SINGLE As Integer = 1
    Public Const MTR_DOUBLE As Integer = 2
    Public Const MTR_QUAD As Integer = 3
    Public Const MTR_UP_DOWN As Integer = 4

    Public Const MTR_NOT_CLEAR As Integer = &H0
    Public Const MTR_STOPPED_BY_ORG As Integer = &H1
    Public Const MTR_STOPPED_BY_ORG_Z As Integer = &H2
    Public Const MTR_NORMAL_STOPPED As Integer = &H4
    Public Const MTR_STOPPED_BY_LIMIT As Integer = &H8

    ' Set/GetComparator
    Public Const MTR_COMP_CONFIG1 As Integer = 0
    Public Const MTR_COMP_CONFIG2 As Integer = 1
    Public Const MTR_COMP_CONFIG3 As Integer = 2
    Public Const MTR_NO_COMP As Integer = 0
    Public Const MTR_GT_COMP As Integer = 1
    Public Const MTR_EQ_COMP As Integer = 2
    Public Const MTR_PRESET1 As Integer = 3
    Public Const MTR_PRESET2 As Integer = 4
    Public Const MTR_PRESET3 As Integer = 5
    Public Const MTR_COMP_OBJECT As Integer = 6
    Public Const MTR_COMP_OBJECT1 As Integer = 6
    Public Const MTR_COMP_OBJECT2 As Integer = 7
    Public Const MTR_COMP_OUT As Integer = 9

    ' Set/GetCLR
    Public Const MTR_ONESHOT As Integer = 0
    Public Const MTR_LEVEL As Integer = 1
    Public Const MTR_CLR_OUT As Integer = 2

    ' Set/Get/StartMotion
    Public Enum MotionStartMotion
        MTN_JOG = 0
        MTN_ORG
        MTN_PTP
        MTN_ORG_PLS
        MTN_TIMER
    End Enum

    Public Const MTR_JOG As Integer = 0
    Public Const MTR_ORG As Integer = 1
    Public Const MTR_PTP As Integer = 2
    Public Const MTR_ORG_PLS As Integer = 3
    Public Const MTR_TIMER As Integer = 4
    Public Const MTR_LINE As Integer = 6
    Public Const MTR_ARC As Integer = 7
    Public Const MTR_DRAW_IN As Integer = 8
    Public Const MTR_RESTART As Integer = 9

    Public Const MTR_NORMAL As Integer = &H0
    Public Const MTR_CONST As Integer = &H10000

    Public Const MTR_ACC_NORMAL As Integer = 0
    Public Const MTR_ACC_SIN As Integer = 1
    Public Const MTR_ACC_EXP As Integer = 2
    Public Const MTR_ACC_ORIGINAL As Integer = 3
    Public Const MTR_FH As Integer = &H100

    Public Const MTR_STOP_ORG As Integer = &H0
    Public Const MTR_STOP_ORG_EZ As Integer = &H100

    Public Const MTR_SQUEAR_ROOT2 As Integer = &H200
    Public Const MTR_MARK As Integer = &H400
    Public Const MTR_MARK_ACC As Integer = &H800

    Public Const MTR_CW As Integer = 1
    Public Const MTR_CCW As Integer = -1

    Public Enum MotionDirection
        MTN_CCW = -1
        MTN_CW = 1
    End Enum

    ' StopMotion
    Public Const MTR_DEC_STOP As Integer = 0
    Public Const MTR_IMMEDIATE_STOP As Integer = 1
    Public Const MTR_STOP_ALL As Integer = 2

    Public Enum MotionStopMotion
        MTN_DEC_STOP = 0
        MTN_IMMEDIATE_STOP
        MTN_STOP_ALL
    End Enum

    ' Set/GetSync
    Public Const MTR_RESET_SYNC_START As Integer = 0
    Public Const MTR_SET_SYNC_START As Integer = 1

    ' ChangeSpeed
    Public Const MTR_IMMEDIATE_CHANGE As Integer = 0
    Public Const MTR_ACCDEC_CHANGE As Integer = 1
    Public Const MTR_LOW_SPEED As Integer = 3
    Public Const MTR_DEC_LOW_SPEED As Integer = 4
    Public Const MTR_RELEASE_DEC As Integer = 5
    Public Const MTR_HOLD_NOW_SPEED As Integer = 6
    Public Const MTR_RELEASE_HOLD As Integer = 7

    ' Read/Write/ClearCounter, SetSampleCounter
    Public Const MTR_ENCODER As Integer = 0
    Public Const MTR_COUNTER As Integer = 1
    Public Const MTR_REMAINS As Integer = 2

    ' GetStatus
    Public Const MTR_BUSY As Integer = 0
    Public Const MTR_FINISH_STATUS As Integer = 1
    Public Const MTR_LIMIT_STATUS As Integer = 2
    Public Const MTR_INTERLOCK_STATUS As Integer = 4

    Public Const MTR_ILOCK_OFF As Integer = 0
    Public Const MTR_ILOCK_ON As Integer = 1

    Public Const MTR_ERROR_CODE As Integer = 5
    Public Const MTR_CONFIRM_END As Integer = 6
    Public Const MTR_COMPLETE As Integer = 0
    Public Const MTR_NOT_COMPLETE As Integer = 1
    Public Const MTR_COMPARATOR As Integer = 7

    Public Enum MotionGetStatus
        MTN_BUSY = 0
        MTN_FINISH_STATUS
        MTN_LIMIT_STATUS
        MTN_INTERLOCK_STATUS
        MTN_ERROR_CODE
        MTN_COMPARATOR
    End Enum

    ' SetSampleCounter
    Public Const MTR_NORMAL_SAMP As Integer = &H0
    Public Const MTR_BUSY_SAMP As Integer = &H100
    Public Const MTR_ONETIME As Integer = &H0
    Public Const MTR_REPEAT As Integer = &H1000

    Public Const MTR_SAMPLE_FINISHED As Integer = 0
    Public Const MTR_NOW_SAMPLING As Integer = 1
    Public Const MTR_WAITING_BUSY As Integer = 2
    Public Const MTR_FULL_BUFFER As Integer = 3

    ' -----------------------------------------------------------------------
    '
    '              Code
    '
    ' -----------------------------------------------------------------------
    Public Const MTR_ERROR_SUCCESS As Integer = &H0
    Public Const MTR_ERROR_NOT_DEVICE As Integer = &HC0000001
    Public Const MTR_ERROR_NOT_OPEN As Integer = &HC0000002
    Public Const MTR_ERROR_INVALID_HANDLE As Integer = &HC0000003
    Public Const MTR_ERROR_ALREADY_OPEN As Integer = &HC0000004
    Public Const MTR_ERROR_NOT_SUPPORTED As Integer = &HC0000009
    Public Const MTR_ERROR_NOW_MOVING As Integer = &HC0001000
    Public Const MTR_ERROR_NOW_STOPPED As Integer = &HC0001001
    Public Const MTR_ERROR_WRITE_FAILED As Integer = &HC0001010
    Public Const MTR_ERROR_READ_FAILED As Integer = &HC0001011
    Public Const MTR_ERROR_INVALID_DEVICE As Integer = &HC0001012
    Public Const MTR_ERROR_INVALID_AXIS As Integer = &HC0001013
    Public Const MTR_ERROR_INVALID_SPEED As Integer = &HC0001014
    Public Const MTR_ERROR_INVALID_ACCDEC As Integer = &HC0001015
    Public Const MTR_ERROR_INVALID_PULSE As Integer = &HC0001016
    Public Const MTR_ERROR_INVALID_PARAMETER As Integer = &HC0001017
    Public Const MTR_ERROR_NOW_INTERLOCKED As Integer = &HC0001020
    Public Const MTR_ERROR_IMPOSSIBLE As Integer = &HC0001021
    Public Const MTR_ERROR_WRITE_FAILED_EEPROM As Integer = &HC0001022
    Public Const MTR_ERROR_READ_FAILED_EEPROM As Integer = &HC0001023
    Public Const MTR_ERROR_NOT_ALLOCATE_MEMORY As Integer = &HC0001024
    Public Const MTR_ERROR_NOW_WAIT_STA As Integer = &HC0001025

    Public Const MTR_INVALID_HANDLE_VALUE As Integer = -1

    ' -----------------------------------------------------------------------
    '   CallBack
    ' -----------------------------------------------------------------------
    Delegate Sub MTRCALLBACK(ByVal EventInfo As MTREVENTTABLE, ByVal dwUser As Integer)

    ' -----------------------------------------------------------------------
    '      Board information
    ' -----------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MTRDEVICE
        Public dwType As Integer
        Public dwRSW As Integer
        Public dwAxis As Integer
    End Structure

    ' -----------------------------------------------------------------------
    '      Motion parameters
    ' -----------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MTRMOTION
        Public dwMode As Integer
        Public dwLowSpeed As Integer
        Public dwSpeed As Integer
        Public dwAcc As Integer
        Public dwDec As Integer
        Public dwSSpeed As Integer
        Public nStep As Integer
        Public nReserved As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MTRMOTIONEX
        Public dwMode As Integer
        Public fLowSpeed As Single
        Public fSpeed As Single
        Public dwAcc As Integer
        Public dwDec As Integer
        Public fSSpeed As Single
        Public nStep As Integer
        Public nReserved As Integer
    End Structure

    ' -----------------------------------------------------------------------
    '      Motion parameters for linear interpolation
    ' -----------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MTRMOTIONLINE
        Public dwMode As Integer
        Public dwLowSpeed As Integer
        Public dwSpeed As Integer
        Public dwAcc As Integer
        Public dwDec As Integer
        Public dwSSpeed As Integer
        Public nXStep As Integer
        Public nYStep As Integer
        Public nReserved As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MTRMOTIONLINEEX
        Public dwMode As Integer
        Public fLowSpeed As Single
        Public fSpeed As Single
        Public dwAcc As Integer
        Public dwDec As Integer
        Public fSSpeed As Single
        Public nXStep As Integer
        Public nYStep As Integer
        Public nReserved As Integer
    End Structure

    ' -----------------------------------------------------------------------
    '      User-supplied motion profile
    ' -----------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MTRORIGINALACC
        Public dwSpeed As Integer
        Public dwAcc As Integer
    End Structure

    ' -----------------------------------------------------------------------
    '      Interrupt event table (status/mask)
    ' -----------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MTREVENTTABLE
        Public dwPulseOut As Integer
        Public dwCounter As Integer
        Public dwComparator As Integer
        Public dwSignal As Integer
        Public dwReserved As Integer
    End Structure

    ' -----------------------------------------------------------------------
    '      Event messages and user-supplied callback routine
    ' -----------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MTREVENTREQ
        Public hWnd As Integer
        Public uPulseOut As Integer
        Public uCounter As Integer
        Public uComparator As Integer
        Public uSignal As Integer
        Public uReserved As Integer
        Public lpCallBackProc As MTRCALLBACK
        Public dwUser As Integer
    End Structure

    ' -----------------------------------------------------------------------
    '      Sampling condition
    ' -----------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure MTRSAMPLEREQ
        Public hWnd As Integer
        Public uStopMsg As Integer
        Public uSampleMsg As Integer
        Public dwAxis As Integer
        Public dwMode As Integer
        Public dwSampleRate As Integer
        Public dwBufferSize As Integer
        Public dwSampleCounter As Integer
    End Structure


    '-------------------------------------------------------------------------------------------------
    '
    '       Declare Function
    '
    '-------------------------------------------------------------------------------------------------
    Declare Function MtrOpen Lib "FbiMtr.DLL" (ByVal lpszName As String, ByVal dwFlags As Integer) As Integer
    Declare Function MtrClose Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer) As Integer
    Declare Function MtrGetDeviceInfo Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pDevInfo As MTRDEVICE) As Integer
    Declare Function MtrReset Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer) As Integer
    Declare Function MtrOffInterLock Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer) As Integer

    Declare Function MtrSetBaseClock Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwClock As Integer) As Integer
    Declare Function MtrSetPulseOut Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByVal dwConfig As Integer) As Integer
    Declare Function MtrSetLimitConfig Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByVal dwConfig As Integer) As Integer
    Declare Function MtrSetCounterConfig Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByVal dwConfig As Integer) As Integer
    Declare Function MtrSetComparator Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByVal dwConfig As Integer) As Integer
    Declare Function MtrSetCLR Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByVal dwConfig As Integer) As Integer
    Declare Function MtrSetAccCurve Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwStageNum As Integer, ByRef pOriginalAcc As MTRORIGINALACC) As Integer
    Declare Function MtrSetAccCurve Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwStageNum As Integer, ByVal pOriginalAcc() As MTRORIGINALACC) As Integer
    Declare Function MtrSetMotion Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMove As Integer, ByRef pMotion As MTRMOTION) As Integer
    Declare Function MtrSetMotionEx Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMove As Integer, ByRef pMotion As MTRMOTIONEX) As Integer
    Declare Function MtrSetMotionLine Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef MotionLine As MTRMOTIONLINE) As Integer
    Declare Function MtrSetMotionLineEx Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef MotionLine As MTRMOTIONLINEEX) As Integer

    Declare Function MtrStartMotion Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMove As Integer) As Integer
    Declare Function MtrStopMotion Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer) As Integer
    Declare Function MtrRestart Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer) As Integer
    Declare Function MtrSetSync Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwSync As Integer) As Integer
    Declare Function MtrStartSync Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer) As Integer
    Declare Function MtrSingleStep Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal nDir As Integer) As Integer
    Declare Function MtrChangeSpeed Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByVal dwSpeed As Integer) As Integer
    Declare Function MtrChangeSpeedEx Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByVal fSpeed As Single) As Integer

    Declare Function MtrGetStatus Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByRef pdwStatus As Integer) As Integer
    Declare Function MtrReadSpeed Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwSpeed As Integer) As Integer
    Declare Function MtrReadSpeedEx Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pfSpeed As Single) As Integer
    Declare Function MtrReadCounter Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByRef pnPos As Integer) As Integer
    Declare Function MtrWriteCounter Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByVal nPos As Integer) As Integer
    Declare Function MtrClearCounter Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer) As Integer
    Declare Function MtrOutputDO Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwDO As Integer) As Integer
    Declare Function MtrInputDI Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwDI As Integer) As Integer
    Declare Function MtrOutputDO1 Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwDO1 As Integer) As Integer
    Declare Function MtrOutputDO2 Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwDO2 As Integer) As Integer
    Declare Function MtrOutputDO3 Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwDO3 As Integer) As Integer
    Declare Function MtrOutputDO4 Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwDO4 As Integer) As Integer
    Declare Function MtrOutputCLR Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwCLR As Integer) As Integer
    Declare Function MtrOutputSON Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwSON As Integer) As Integer

    Declare Function MtrSetEvent Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pEventReq As MTREVENTREQ) As Integer
    Declare Function MtrSetEventMask Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pEventMask As MTREVENTTABLE) As Integer

    Declare Function MtrSetSampleCounter Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pSample As MTRSAMPLEREQ) As Integer
    Declare Function MtrStartSampleCounter Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer) As Integer
    Declare Function MtrStopSampleCounter Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer) As Integer
    Declare Function MtrGetSampleStatus Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwStatus As Integer, ByRef pdwSampNum As Integer) As Integer
    Declare Function MtrGetSampleData Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwSampNum As Integer, ByRef pnPos As Integer) As Integer
    Declare Function MtrGetSampleData Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwSampNum As Integer, ByVal pnPos() As Integer) As Integer

    Declare Function MtrGetBaseClock Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwClock As Integer) As Integer
    Declare Function MtrGetPulseOut Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByRef pdwConfig As Integer) As Integer
    Declare Function MtrGetLimitConfig Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByRef pdwConfig As Integer) As Integer
    Declare Function MtrGetCounterConfig Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByRef pdwConfig As Integer) As Integer
    Declare Function MtrGetComparator Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByRef pdwConfig As Integer) As Integer
    Declare Function MtrGetCLR Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMode As Integer, ByRef pdwConfig As Integer) As Integer
    Declare Function MtrGetAccCurve Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwStageNum As Integer, ByRef pOriginalAcc As MTRORIGINALACC) As Integer
    Declare Function MtrGetAccCurve Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwStageNum As Integer, ByVal pOriginalAcc() As MTRORIGINALACC) As Integer
    Declare Function MtrGetMotion Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMove As Integer, ByRef pMotion As MTRMOTION) As Integer
    Declare Function MtrGetMotionEx Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByVal dwMove As Integer, ByRef pMotion As MTRMOTIONEX) As Integer
    Declare Function MtrGetMotionLine Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pMotionLine As MTRMOTIONLINE) As Integer
    Declare Function MtrGetMotionLineEx Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pMotionLine As MTRMOTIONLINEEX) As Integer
    Declare Function MtrGetSync Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwSync As Integer) As Integer
    Declare Function MtrGetEvent Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pEventReq As MTREVENTREQ) As Integer
    Declare Function MtrGetEventMask Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pEventMask As MTREVENTTABLE) As Integer
    Declare Function MtrGetSampleCounter Lib "FbiMtr.DLL" (ByVal hDeviceHandle As Integer, ByRef pSample As MTRSAMPLEREQ) As Integer

End Class

'End Namespace

