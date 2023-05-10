'---------------------------------------------------
'   PX - Taping OAI Development
'===================================================
'   Designed By : Zulhisham Tan
'   Module      : cls_DIO_Port.vb
'   Date        : 03-Jun-2009
'   Version     : 2009.06.03.001
'---------------------------------------------------
'   Copyright (C) 2007-2009 az_Zulhisham
'---------------------------------------------------


Public Class cls_DIO_Port

    Inherits cls_DIO2727

    Private i_PortNo As Integer
    Private i_PortType As PortDirection

    Private bi_PortData As Byte

    Public Sub New()

        bi_PortData = 0

    End Sub

    Public Sub New(ByVal sbrdName As String, ByVal lbrdNo As Long, ByVal sbrdMaker As String, ByVal lbrdHnd As Long, ByVal shPortNo As Short, ByVal iPortType As PortDirection)

        MyBase.New(sbrdName, lbrdNo, sbrdMaker, lbrdHnd)
        i_PortNo = Convert.ToInt32(shPortNo)
        i_PortType = iPortType

    End Sub

    Public Overrides Function DataValue() As Byte

        Dim iRetVal As Integer = -1

        If i_PortType = PortDirection.In_Port Then
            iRetVal = DioInputByte(MyBase.BoardHnd, i_PortNo, bi_PortData)
        End If

        Return bi_PortData

    End Function

    Public Overrides Function NewPortValue(ByVal bi_Data As Byte) As Integer

        Dim iRetVal As Integer = -1

        If i_PortType = PortDirection.Out_Port Then
            iRetVal = DioOutputByte(MyBase.BoardHnd, i_PortNo, bi_Data)

            If iRetVal = FBIDIO_ERROR_SUCCESS Then bi_PortData = bi_Data
        End If

        Return 0

    End Function

    Public Overrides Function ToString() As String

        Return MyBase.BoardName & " : " & i_PortNo & "-" & bi_PortData

    End Function

    Public Overrides ReadOnly Property PortNo() As Long

        Get
            Return i_PortNo
        End Get

    End Property

    Public Property PortData() As Byte

        Get
            Return bi_PortData
        End Get

        Set(ByVal value As Byte)
            bi_PortData = value
        End Set

    End Property

End Class
