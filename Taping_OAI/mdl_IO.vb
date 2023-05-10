'---------------------------------------------------
'   PX - Taping OAI Development
'===================================================
'   Designed By : Zulhisham Tan
'   Module      : mdl_IO.vb
'   Date        : 03-Jun-2009
'   Version     : 2009.06.03.001
'---------------------------------------------------
'   Copyright (C) 2007-2009 az_Zulhisham
'---------------------------------------------------


Option Strict Off
Option Explicit On


Module mdl_IO

    Public Structure OVERLAPPED
        Public Internal As Int32
        Public InternalHigh As Int32
        Public Offset As Int32
        Public OffsetHigh As Int32
        Public hEvent As IntPtr
    End Structure

    Public Const FBIDIO_RSTIN_MASK As Short = 1 ' The symbols used when carrying out the mask of the RSTIN
    Public Const FBIDIO_FLAG_SHARE As Short = &H2S ' The flag is applicable to the DioOpen function. This flag shows that the device is opened as shareable.

    Public Const FBIDIO_IN1_8 As Short = 0 ' Read the data from IN1 through IN8.
    Public Const FBIDIO_IN9_16 As Short = 1 ' Read the data from IN9 through IN16.
    Public Const FBIDIO_IN17_24 As Short = 2 ' Read the data from IN17 through IN24.
    Public Const FBIDIO_IN25_32 As Short = 3 ' Read the data from IN25 through IN32.
    Public Const FBIDIO_IN33_40 As Short = 4 ' Read the data from IN33 through IN40.
    Public Const FBIDIO_IN41_48 As Short = 5 ' Read the data from IN41 through IN48.
    Public Const FBIDIO_IN49_56 As Short = 6 ' Read the data from IN49 through IN56.
    Public Const FBIDIO_IN57_64 As Short = 7 ' Read the data from IN57 through IN64.


    Public Const FBIDIO_IN1_16 As Short = 0 ' Read the data from IN1 through IN16.
    Public Const FBIDIO_IN17_32 As Short = 2 ' Read the data from IN17 through IN32.
    Public Const FBIDIO_IN33_48 As Short = 4 ' Read the data from IN33 through IN48.
    Public Const FBIDIO_IN49_64 As Short = 6 ' Read the data from IN49 through IN64.

    Public Const FBIDIO_IN1_32 As Short = 0 ' Read the data from IN1 through IN32.
    Public Const FBIDIO_IN33_64 As Short = 4 ' Read the data from IN33 through IN64.

    Public Const FBIDIO_OUT1_8 As Short = 0 ' Write the data to OUT1 through OUT8.
    Public Const FBIDIO_OUT9_16 As Short = 1 ' Write the data to OUT9 through OUT16.
    Public Const FBIDIO_OUT17_24 As Short = 2 ' Write the data to OUT17 through OUT24.
    Public Const FBIDIO_OUT25_32 As Short = 3 ' Write the data to OUT25 through OUT32.
    Public Const FBIDIO_OUT33_40 As Short = 4 ' Write the data to OUT33 through OUT40.
    Public Const FBIDIO_OUT41_48 As Short = 5 ' Write the data to OUT41 through OUT48.
    Public Const FBIDIO_OUT49_56 As Short = 6 ' Write the data to OUT49 through OUT56.
    Public Const FBIDIO_OUT57_64 As Short = 7 ' Write the data to OUT57 through OUT64.

    Public Const FBIDIO_OUT1_16 As Short = 0 ' Write the data to OUT1 through OUT16.
    Public Const FBIDIO_OUT17_32 As Short = 2 ' Write the data to OUT17 through OUT32.
    Public Const FBIDIO_OUT33_48 As Short = 4 ' Write the data to OUT33 through OUT48.
    Public Const FBIDIO_OUT49_64 As Short = 6 ' Write the data to OUT49 through OUT64.

    Public Const FBIDIO_OUT1_32 As Short = 0 ' Write the data to OUT1 through OUT32.
    Public Const FBIDIO_OUT33_64 As Short = 4 ' Write the data to OUT33 through OUT64.

    Public Const FBIDIO_STB1_ENABLE As Short = &H1S ' The STB1 event is eabled.
    Public Const FBIDIO_STB1_HIGH_EDGE As Short = &H10S ' The rising edge for STB1 is enabled.

    Public Const FBIDIO_ACK2_ENABLE As Short = &H4S ' The ACK2 event is enabled.
    Public Const FBIDIO_ACK2_HIGH_EDGE As Short = &H40S ' The rising edge for ACK2 is enabled.


    ' -----------------------------------------------------------------------
    '       Return Value
    ' -----------------------------------------------------------------------
    Public Const FBIDIO_ERROR_SUCCESS As Short = 0 ' The process is successfully completed.
    Public Const FBIDIO_ERROR_NOT_DEVICE As Integer = &HC0000001 ' The specified driver cannot be called.
    Public Const FBIDIO_ERROR_NOT_OPEN As Integer = &HC0000002 ' The specified driver cannot be opened.
    Public Const FBIDIO_ERROR_INVALID_HANDLE As Integer = &HC0000003 ' The device handle is invalid.
    Public Const FBIDIO_ERROR_ALREADY_OPEN As Integer = &HC0000004 ' The device has already been opened.
    Public Const FBIDIO_ERROR_HANDLE_EOF As Integer = &HC0000005 ' End of file is reached.
    Public Const FBIDIO_ERROR_MORE_DATA As Integer = &HC0000006 ' More available data exists.
    Public Const FBIDIO_ERROR_INSUFFICIENT_BUFFER As Integer = &HC0000007 ' The data area passed to the system call is too small.
    Public Const FBIDIO_ERROR_IO_PENDING As Integer = &HC0000008 ' Overlapped I/O operations are in progress.
    Public Const FBIDIO_ERROR_NOT_SUPPORTED As Integer = &HC0000009 ' The specified function is not supported.
    Public Const FBIDIO_ERROR_MEMORY_NOTALLOCATED As Integer = &HC0001000 ' Not enough memory.
    Public Const FBIDIO_ERROR_PARAMETER As Integer = &HC0001001 ' he specified parameter is invalid.
    Public Const FBIDIO_ERROR_INVALID_CALL As Integer = &HC0001002 ' Invalid function is called.
    Public Const FBIDIO_ERROR_DRVCAL As Integer = &HC0001003 ' Failed to call the driver.
    Public Const FBIDIO_ERROR_NULL_POINTER As Integer = &HC0001004 ' A NULL pointer is passed between the driver and the DLL.

    ' -----------------------------------------------------------------------
    '       DLL
    ' -----------------------------------------------------------------------
    Public Declare Function DioOpen Lib "FbiDio.DLL" (ByVal lpszName As String, ByVal fdwAttrs As Integer) As Integer
    Public Declare Function DioClose Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer) As Integer
    Public Declare Function DioInputPoint Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pBuffer As Integer, ByVal dwStartNum As Integer, ByVal dwNum As Integer) As Integer
    Public Declare Function DioOutputPoint Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pBuffer As Integer, ByVal dwStartNum As Integer, ByVal dwNum As Integer) As Integer
    Public Declare Function DioGetBackGroundUseTimer Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pnUse As Integer) As Integer
    Public Declare Function DioSetBackGroundUseTimer Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nUse As Integer) As Integer
    Public Declare Function DioSetBackGround Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal dwStartPoint As Integer, ByVal dwPointNum As Integer, ByVal dwValueNum As Integer, ByVal dwCycle As Integer, ByVal dwCount As Integer, ByVal dwOption As Integer) As Integer
    Public Declare Function DioFreeBackGround Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal hBackGroundHandle As Integer) As Integer
    Public Declare Function DioStopBackGround Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal hBackGroundHandle As Integer) As Integer
    Public Declare Function DioGetBackGroundStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal hBackGroundHandle As Integer, ByRef pnStartPoint As Integer, ByRef pnPointNum As Integer, ByRef pnValueNum As Integer, ByRef pnCycle As Integer, ByRef pnCount As Integer, ByRef pnOption As Integer, ByRef pnExecute As Integer, ByRef pnExecCount As Integer, ByRef pnBufferOffset As Integer, ByRef pnOver As Integer) As Integer
    Public Declare Function DioInputPointBack Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal hBackGroundHandle As Integer, ByRef pBuffer As Integer, ByVal nNumOfBytesToRead As Integer, ByRef pOverlapped As OVERLAPPED) As Integer
    Public Declare Function DioOutputPointBack Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal hBackGroundHandle As Integer, ByRef pBuffer As Integer, ByVal nNumOfBytesToWrite As Integer, ByRef lpOverlapped As OVERLAPPED) As Integer
    Public Declare Function DioWatchPointBack Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal hBackGroundHandle As Integer, ByRef pBuffer As Integer, ByVal nNumOfBytesToRead As Integer, ByRef lpOverlapped As OVERLAPPED) As Integer
    Public Declare Function DioGetInputHandShakeConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pnInputHandShakeContig As Integer, ByRef pdwBitMask1 As Integer, ByRef pdwBitMask2 As Integer) As Integer
    Public Declare Function DioSetInputHandShakeConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nInputHandShakeContig As Integer, ByVal dwBitMask1 As Integer, ByVal dwBitMask2 As Integer) As Integer
    Public Declare Function DioGetOutputHandShakeConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pnOutputHandShakeContig As Integer, ByRef pdwBitMask1 As Integer, ByRef pdwBitMask2 As Integer) As Integer
    Public Declare Function DioSetOutputHandShakeConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nOutputHandShakeConfig As Integer, ByVal dwBitMask1 As Integer, ByVal dwBitMask2 As Integer) As Integer
    Public Declare Function DioInputHandShake Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef lpBuffer As Integer, ByVal nNumOfBytesToRead As Integer, ByRef lpNumOfBytesRead As Integer, ByRef lpOverlapped As OVERLAPPED) As Integer
    Public Declare Function DioInputHandShakeEx Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef lpBuffer As Integer, ByVal nNumOfBytesToRead As Integer, ByRef lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Integer) As Integer
    Public Declare Function DioOutputHandShake Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef lpBuffer As Integer, ByVal nNumOfBytesToWrite As Integer, ByRef lpNumOfBytesWritten As Integer, ByRef lpOverlapped As OVERLAPPED) As Integer
    Public Declare Function DioOutputHandShakeEx Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef lpBuffer As Integer, ByVal nNumOfBytesToWrite As Integer, ByRef lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Integer) As Integer
    Public Declare Function DioStopInputHandShake Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer) As Integer
    Public Declare Function DioStopOutputHandShake Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer) As Integer
    Public Declare Function DioGetHandShakeStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwpDeviceStatus As Integer, ByRef pdwpInputedBuffNum As Integer, ByRef pdwpOutputedBuffNum As Integer) As Integer
    Public Declare Function DioInputByte Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nNo As Integer, ByRef pbValue As Byte) As Integer
    Public Declare Function DioInputWord Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nNo As Integer, ByRef pwValue As Short) As Integer
    Public Declare Function DioInputDword Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nNo As Integer, ByRef pdwValue As Integer) As Integer
    Public Declare Function DioOutputByte Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nNo As Integer, ByVal bValue As Byte) As Integer
    Public Declare Function DioOutputWord Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nNo As Integer, ByVal wValue As Short) As Integer
    Public Declare Function DioOutputDword Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nNo As Integer, ByVal dwValue As Integer) As Integer
    Public Declare Function DioGetAckStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pbAckStatus As Byte) As Integer
    Public Declare Function DioSetAckPulseCommand Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal bCommand As Byte) As Integer
    Public Declare Function DioGetStbStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pbStbStatus As Byte) As Integer
    Public Declare Function DioSetStbPulseCommand Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal bCommand As Byte) As Integer
    Public Declare Function DioInputUniversalPoint Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwUniversalPoint As Integer) As Integer
    Public Declare Function DioOutputUniversalPoint Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal dwUniversalPoint As Integer) As Integer
    Public Declare Function DioSetTimeOut Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal dwInputTotalTimeout As Integer, ByVal dwInputIntervalTimeout As Integer, ByVal dwOutputTotalTimeout As Integer, ByVal dwOutputIntervalTimeout As Integer) As Integer
    Public Declare Function DioGetTimeOut Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwpInputTotalTimeout As Integer, ByRef pdwpInputIntervalTimeout As Integer, ByRef pdwpOutputTotalTimeout As Integer, ByRef pdwpOutputIntervalTimeout As Integer) As Integer
    Public Declare Function DioSetIrqMask Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal bIrqMask As Byte) As Integer
    Public Declare Function DioGetIrqMask Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pbIrqMask As Byte) As Integer
    Public Declare Function DioSetIrqConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal bIrqConfig As Byte) As Integer
    Public Declare Function DioGetIrqConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pbIrqConfig As Byte) As Integer
    Public Declare Function DioGetDeviceConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwDeviceConfig As Integer) As Integer
    Public Declare Function DioSetTimerConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal bTimerConfigValue As Byte) As Integer
    Public Declare Function DioGetTimerConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pbTimerConfigValue As Byte) As Integer
    Public Declare Function DioGetTimerCount Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pbTimerCount As Byte) As Integer
    Public Declare Function DioSetLatchStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal bLatchStatus As Byte) As Integer
    Public Declare Function DioGetLatchStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pbLatchStatus As Byte) As Integer
    Public Declare Function DioGetResetInStatus Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pbResetInStatus As Byte) As Integer
    Public Declare Function DioEventRequestPending Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal dwEventEnableMask As Integer, ByRef pEventBuf As Integer, ByRef lpOverlapped As OVERLAPPED) As Integer
    Public Declare Function DioCommonGetPciDeviceInfo Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwDeviceID As Integer, ByRef pdwVenderID As Integer, ByRef pdwClassCode As Integer, ByRef pdwRevisionID As Integer, ByRef pdwBaseAddress0 As Integer, ByRef pdwBaseAddress1 As Integer, ByRef pdwBaseAddress2 As Integer, ByRef pdwBaseAddress3 As Integer, ByRef pdwBaseAddress4 As Integer, ByRef pdwBaseAddress5 As Integer, ByRef pdwSubsystemID As Integer, ByRef pdwSubsystemVenderID As Integer, ByRef pdwInterruptLine As Integer, ByRef pdwBoardID As Integer) As Integer
    Public Declare Function DioEintSetIrqMask Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal dwSetIrqMask As Integer) As Integer
    Public Declare Function DioEintGetIrqMask Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwGetIrqMask As Integer) As Integer
    Public Declare Function DioEintSetEdgeConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal dwSetFallEdgeConfig As Integer, ByVal dwSetRiseEdgeConfig As Integer) As Integer
    Public Declare Function DioEintGetEdgeConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwGetFallEdgeConfig As Integer, ByRef pdwGetRiseEdgeConfig As Integer) As Integer
    Public Declare Function DioEintInputPoint Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pBuffer As Integer, ByVal dwStartNum As Integer, ByVal dwNum As Integer) As Integer
    Public Declare Function DioEintInputByte Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nNo As Integer, ByRef pbFallValue As Byte, ByRef pbRiseValue As Byte) As Integer
    Public Declare Function DioEintInputWord Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nNo As Integer, ByRef pwFallValue As Short, ByRef pwRiseValue As Short) As Integer
    Public Declare Function DioEintInputDword Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nNo As Integer, ByRef pdwFallValue As Integer, ByRef pdwRiseValue As Integer) As Integer
    Public Declare Function DioEintSetFilterConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nNo As Integer, ByVal nSetFilterConfig As Integer) As Integer
    Public Declare Function DioEintGetFilterConfig Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByVal nNo As Integer, ByRef pnGetFilterConfig As Integer) As Integer
    Public Declare Function DioEventRequestPendingEx Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwEventEnableMask As Integer, ByRef pEventBuf As Integer, ByRef lpOverlapped As OVERLAPPED) As Integer
    Public Declare Function DioGetDeviceConfigEx Lib "FbiDio.DLL" (ByVal hDeviceHandle As Integer, ByRef pdwDeviceConfig As Integer, ByRef pdwDeviceConfigEx As Integer) As Integer

End Module
