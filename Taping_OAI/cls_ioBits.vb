'---------------------------------------------------
'   PX - Taping OAI Development
'===================================================
'   Designed By : Zulhisham Tan
'   Module      : cls_ioBits.vb
'   Date        : 03-Jun-2009
'   Version     : 2009.06.03.001
'---------------------------------------------------
'   Copyright (C) 2007-2009 az_Zulhisham
'---------------------------------------------------


Public Class cls_ioBits

    Inherits cls_DIO_Port


    Public Enum BitDirection
        In_Bit = 0
        Out_Bit
    End Enum


    Private i_BitNo As Integer
    Private i_BitType As BitDirection

    Private bi_BitData As Byte
    Private eBitState As BitsState


    Public Sub New()

        i_BitNo = 0
        eBitState = BitsState.eBit_OFF

    End Sub

    Public Sub New(ByVal sbrdName As String, ByVal lbrdNo As Long, ByVal sbrdMaker As String, ByVal lbrdHnd As Long, ByVal shPortNo As Short, ByVal shBitNo As Short, ByVal iBitType As BitDirection)

        MyBase.New(sbrdName, lbrdNo, sbrdMaker, lbrdHnd, shPortNo, iBitType)

        i_BitNo = Convert.ToInt32(shBitNo)
        i_BitType = iBitType

        If i_BitType = BitDirection.In_Bit Then
            eBitState = MyBase.PortData / (2 ^ (i_BitNo - 1))
        Else
            eBitState = BitsState.eBit_OFF
        End If

    End Sub

    Public Overrides Function Trigger_ON() As Long

        Dim iRetVal As Integer


        iRetVal = DioOutputPoint(MyBase.BoardHnd, Bit_ON, (MyBase.PortNo * 8) + i_BitNo, 1)
        eBitState = BitsState.eBit_ON

    End Function

    Public Overrides Function Trigger_OFF() As Long

        Dim iRetVal As Integer


        iRetVal = DioOutputPoint(MyBase.BoardHnd, Bit_OFF, (MyBase.PortNo * 8) + i_BitNo, 1)
        eBitState = BitsState.eBit_OFF

    End Function

    Public Overrides ReadOnly Property BitNo() As Long

        Get
            Return i_BitNo
        End Get

    End Property

    Public Overrides ReadOnly Property BitState() As BitsState

        Get
            If i_BitType = BitDirection.In_Bit Then
                eBitState = (MyBase.DataValue And (2 ^ (i_BitNo - 1))) / (2 ^ (i_BitNo - 1))
            End If

            Return eBitState
        End Get

    End Property

End Class
