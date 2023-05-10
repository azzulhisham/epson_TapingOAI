'---------------------------------------------------
'   PX - Taping OAI Development
'===================================================
'   Designed By : Zulhisham Tan
'   Module      : cls_PCIBoard.vb
'   Date        : 03-Jun-2009
'   Version     : 2009.06.03.001
'---------------------------------------------------
'   Copyright (C) 2007-2009 az_Zulhisham
'---------------------------------------------------


Public MustInherit Class cls_PCIBoard

    Public Enum PortDirection
        In_Port = 0
        Out_Port
    End Enum

    Public Enum BitsState
        eBit_OFF = Bit_OFF
        eBit_ON = Bit_ON
    End Enum

    Public Const Bit_OFF = 0
    Public Const Bit_ON = 1


    Public MustOverride Property BoardName() As String
    Public MustOverride Property BoardType() As String
    Public MustOverride Property BoardMaker() As String
    Public MustOverride Property BoardNo() As Long


    'Public Overridable Sub NewPortValue(ByVal bi_Data As Byte)

    Public Overridable Function DataValue() As Byte

        Return 0

    End Function

    Public Overridable Function NewPortValue(ByVal bi_Data As Byte) As Integer

        Return 0

    End Function

    Public Overridable Function Trigger_ON() As Long

        Return 0

    End Function

    Public Overridable Function Trigger_OFF() As Long

        Return 0

    End Function

    Public Overridable ReadOnly Property PortNo() As Long

        Get
            Return 0
        End Get

    End Property

    Public Overridable ReadOnly Property BitNo() As Long

        Get
            Return 0
        End Get

    End Property

    Public Overridable ReadOnly Property BitState() As BitsState

        Get
            Return 0
        End Get

    End Property

End Class
