Public Class frm_LngDelay

    Private Sub frm_LngDelay_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        With Me
            .tmr_TimeOver.Enabled = False
            .tmr_Blink.Enabled = False
        End With

    End Sub

    Private Sub frm_LngDelay_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With Me
            With .tmr_Blink
                .Interval = 250
                .Enabled = True
            End With

            With .tmr_TimeOver
                .Interval = 1000 * (13)
                .Enabled = True
            End With
        End With

    End Sub

    Private Sub tmr_Blink_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr_Blink.Tick

        With Me
            .Label1.Visible = Not .Label1.Visible
        End With

    End Sub

    Private Sub tmr_TimeOver_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmr_TimeOver.Tick

        With Me
            .tmr_TimeOver.Enabled = False
            .tmr_Blink.Enabled = False
            .Close()
        End With

    End Sub

End Class