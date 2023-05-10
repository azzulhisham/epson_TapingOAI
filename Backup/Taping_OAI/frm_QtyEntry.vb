Public Class frm_QtyEntry

    Dim fgLoad As Boolean = False


    Private Sub txt_DataInput_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_DataInput.KeyDown

        With Me
            If e.KeyCode = Keys.Enter Then
                If IsNumeric(.txt_DataInput.Text) AndAlso Val(.txt_DataInput.Text) > 0 Then
                    fg_Qty = CType(.txt_DataInput.Text, Integer)
                    Me.Close()
                Else
                    MessageBox.Show("The Qty. value is invalid...!", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    With .txt_DataInput
                        .Text = "0"
                        .SelectAll()
                        .Focus()
                    End With
                End If
            End If
        End With

    End Sub

    Private Sub frm_QtyEntry_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        With Me
            .fgLoad = False
            Me.tmr_ReadIO.Enabled = False
        End With

    End Sub

    Private Sub frm_QtyEntry_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If fgLoad = True Then Exit Sub
        fgLoad = True

        With Me
            With .lbl_PLotNo
                .Text = Tp_OAI.P_Lot_No
            End With

            With .txt_DataInput
                fg_Qty = 0

                If Tp_OAI.SysData(1).P_Lot.GetUpperBound(0) = 0 And Tp_OAI.SysData(1).P_Lot(0).QtyUsed = 0 Then
                    .Text = Tp_OAI.SysData(1).Acceptance
                Else
                    Dim QtyCnt As Integer = (Val(Tp_OAI.SysData(1).Acceptance))

                    For iLp As Integer = 0 To Tp_OAI.SysData(1).P_Lot.GetUpperBound(0)
                        QtyCnt -= Tp_OAI.SysData(1).P_Lot(iLp).QtyUsed
                    Next

                    .Text = QtyCnt.ToString
                End If

                .SelectAll()
            End With

            With .tmr_ReadIO
                .Interval = 250
                .Enabled = True
            End With
        End With

    End Sub

    Private Sub frm_QtyEntry_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        With Me.txt_DataInput
            .Focus()
        End With

    End Sub

    Private Sub tmr_ReadIO_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr_ReadIO.Tick

        Static fg_Trg As Integer = 0

        If fg_Dbg = 0 Then
            With Tp_OAI
                If .IO.pb_EMG.BitState = cls_PCIBoard.BitsState.eBit_OFF Or .IO.pb_STOP.BitState = cls_PCIBoard.BitsState.eBit_ON Then
                    fg_Qty = 0
                    Me.Close()
                End If
            End With
        End If


        With Me
            .lbl_MsgBox.Visible = Not .lbl_MsgBox.Visible
        End With

    End Sub

End Class