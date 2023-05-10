Public Class frm_PwdEntry

    Private Sub frm_PwdEntry_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        With Me
            .tmr_MonitorEnv.Enabled = False
        End With

    End Sub

    Private Sub frm_PwdEntry_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With Tp_OAI
            .AuthenticalCode = ""
        End With

        With Me
            With .txt_PwdEnter
                .Text = ""
            End With

            With .tmr_MonitorEnv
                .Interval = 200
                .Enabled = True
            End With
        End With

    End Sub

    Private Sub frm_PwdEntry_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        With Me
            .txt_PwdEnter.Focus()
        End With

    End Sub

    Private Sub txt_PwdEnter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_PwdEnter.KeyDown

        If e.KeyCode = Keys.Enter Then
            With Tp_OAI
                .AuthenticalCode = Me.txt_PwdEnter.Text
                Me.Close()
            End With
        ElseIf e.KeyCode = Keys.Escape Then
            Me.Close()
        End If

    End Sub

    Private Sub tmr_MonitorEnv_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmr_MonitorEnv.Tick

        With Tp_OAI.IO
            If .pb_STOP.BitState = cls_PCIBoard.BitsState.eBit_ON Or .pb_EMG.BitState = cls_PCIBoard.BitsState.eBit_OFF Then
                Me.Close()
            End If
        End With

    End Sub

End Class