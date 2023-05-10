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


Imports System.Windows.Forms


Public Class cls_DIO2727

    Inherits cls_PCIBoard

    Private sBoardName As String
    Private sBoardType As String
    Private sBoardMaker As String
    Private lBoardNo As Long

    Private lHnd As Long = -1
    Private Shared DIO_Count As Integer


    Public Sub New()

        sBoardType = "Digital I/O"

    End Sub

    Public Sub New(ByVal lbrdHnd As Long)

        lHnd = lbrdHnd

    End Sub

    Public Sub New(ByVal sbrdName As String, ByVal lbrdNo As Long, ByVal sbrdMaker As String)

        sBoardType = "Digital I/O"
        sBoardName = sbrdName
        sBoardMaker = sbrdMaker

        lBoardNo = lbrdNo
        DIO_Open()

        DIO_Count += 1

    End Sub

    Public Sub New(ByVal sbrdName As String, ByVal lbrdNo As Long, ByVal sbrdMaker As String, ByVal lbrdHnd As Long)

        sBoardType = "Digital I/O"
        sBoardName = sbrdName
        sBoardMaker = sbrdMaker

        lBoardNo = lbrdNo
        lHnd = lbrdHnd

    End Sub

    Public Function DIO_Open() As Long

        lHnd = -1

        Try
            lHnd = DioOpen(sBoardName & Chr(0), FBIDIO_FLAG_SHARE)
        Catch
            MessageBox.Show("Open DIO error. Please check hardware and system configuration.", "Hardware Initialize Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return lHnd

    End Function

    Public Function DIO_Close() As Long

        If lHnd > 0 Then
            Return (DioClose(lHnd))
        End If

    End Function

    Public Overrides Function ToString() As String

        Return sBoardMaker & " : " & sBoardName & "-" & lHnd

    End Function

    Public Overrides Property BoardName() As String

        'FBIDIO1
        Get
            Return sBoardName
        End Get

        Set(ByVal value As String)
            sBoardName = value
        End Set

    End Property

    Public Overrides Property BoardType() As String

        Get
            Return sBoardType
        End Get

        Set(ByVal value As String)
            sBoardType = value
        End Set

    End Property

    Public Overrides Property BoardMaker() As String

        Get
            Return sBoardMaker
        End Get

        Set(ByVal value As String)
            sBoardMaker = value
        End Set

    End Property

    Public Overrides Property BoardNo() As Long

        Get
            Return lBoardNo
        End Get

        Set(ByVal value As Long)
            lBoardNo = value
        End Set

    End Property

    Public ReadOnly Property BoardHnd() As Long

        Get
            Return lHnd
        End Get

    End Property

    Public Shared ReadOnly Property Count() As Integer

        Get
            Return DIO_Count
        End Get

    End Property

End Class
