Imports Microsoft.Win32



Public Class frm_AdvSetting

    Dim LstViewSelectedIndex As Integer = -1


    Private Sub frm_AdvSetting_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        LstViewSelectedIndex = -1

        With Tp_OAI
            Me.txt_DataPath.Text = .DataPath
            Me.txt_LotDataPath.Text = .LotDataPath
            Me.txt_MachineNo.Text = .MC_No

            Me.txt_LowSpeed.Text = .MotionSys(0).NormalMotionSetting.dwLowSpeed
            Me.txt_Speed.Text = .MotionSys(0).NormalMotionSetting.dwSpeed
            Me.txt_Acc.Text = .MotionSys(0).NormalMotionSetting.dwAcc
            Me.txt_Dec.Text = .MotionSys(0).NormalMotionSetting.dwDec
        End With

        With Me.ListView1
            .Clear()
            .Columns.Add("Product", 70, HorizontalAlignment.Center)
            .Columns.Add("Scene No.", 70, HorizontalAlignment.Center)


            For iLp As Integer = 0 To ProdID.GetUpperBound(0)
                Dim itm As ListViewItem = Me.ListView1.Items.Add(ProdID(iLp))
                itm.SubItems.Add(SceneNoDB(iLp))
            Next
        End With

    End Sub

    Private Sub frm_AdvSetting_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        With Me
            With .txt_DataPath
                .SelectAll()
                .Focus()
            End With
        End With

    End Sub

    Private Function ValidateField() As Integer

        With Me
            With .txt_DataPath
                Try
                    My.Computer.FileSystem.CreateDirectory(.Text)
                    My.Computer.FileSystem.DeleteDirectory(.Text, FileIO.DeleteDirectoryOption.DeleteAllContents)
                Catch ex As Exception
                    MessageBox.Show("Inavlid Path Name.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Me.ErrorProvider1.SetError(Me.txt_DataPath, "Invalid Path Name.")
                    .SelectAll()
                    .Focus()
                    Return -1
                End Try
            End With

            With .txt_LotDataPath
                Try
                    My.Computer.FileSystem.CreateDirectory(.Text)
                    My.Computer.FileSystem.DeleteDirectory(.Text, FileIO.DeleteDirectoryOption.DeleteAllContents)
                Catch ex As Exception
                    MessageBox.Show("Inavlid Path Name.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Me.ErrorProvider1.SetError(Me.txt_LotDataPath, "Invalid Path Name.")
                    .SelectAll()
                    .Focus()
                    Return -1
                End Try
            End With

            With .txt_LowSpeed
                If Not IsNumeric(.Text) Or Val(.Text) < 0 Or Val(.Text) > 3000 Then
                    MessageBox.Show("Inavlid value - It must be a numeric and more than zero.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Me.ErrorProvider1.SetError(Me.txt_LowSpeed, "Invalid Value.")
                    .SelectAll()
                    .Focus()
                    Return -1
                End If
            End With

            With .txt_Speed
                If Not IsNumeric(.Text) Or Val(.Text) < 0 Or Val(.Text) > 27000 Then
                    MessageBox.Show("Inavlid value - It must be a numeric and more than zero.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Me.ErrorProvider1.SetError(Me.txt_Speed, "Invalid Value.")
                    .SelectAll()
                    .Focus()
                    Return -1
                End If
            End With

            With .txt_Acc
                If Not IsNumeric(.Text) Or Val(.Text) < 0 Or Val(.Text) > 1000 Then
                    MessageBox.Show("Inavlid value - It must be a numeric and more than zero.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Me.ErrorProvider1.SetError(Me.txt_Acc, "Invalid Value.")
                    .SelectAll()
                    .Focus()
                    Return -1
                End If
            End With

            With .txt_Dec
                If Not IsNumeric(.Text) Or Val(.Text) < 0 Or Val(.Text) > 1000 Then
                    MessageBox.Show("Inavlid value - It must be a numeric and more than zero.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Me.ErrorProvider1.SetError(Me.txt_Dec, "Invalid Value.")
                    .SelectAll()
                    .Focus()
                    Return -1
                End If
            End With
        End With

        Return 0

    End Function

    Private Sub btn_Save_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_Save.Click

        Dim regSubKey_Rn As RegistryKey = regKey.CreateSubKey("Software\az_IOLogics\TapingOAI\Motion\Running")


        If ValidateField() < 0 Then
            Exit Sub
        End If

        With Tp_OAI
            .DataPath = Me.txt_DataPath.Text.Trim
            .LotDataPath = Me.txt_LotDataPath.Text.Trim
            .MC_No = Me.txt_MachineNo.Text.Trim

            .MotionSys(0).NormalMotionSetting.dwLowSpeed = Val(Me.txt_LowSpeed.Text)
            .MotionSys(0).NormalMotionSetting.dwSpeed = Val(Me.txt_Speed.Text)
            .MotionSys(0).NormalMotionSetting.dwAcc = Val(Me.txt_Acc.Text)
            .MotionSys(0).NormalMotionSetting.dwDec = Val(Me.txt_Dec.Text)

            regSubKey.SetValue("DataPath", .DataPath.Trim, RegistryValueKind.String)
            regSubKey.SetValue("LotDataPath", .LotDataPath.Trim, RegistryValueKind.String)
            regSubKey.SetValue("MachineNo", .MC_No.Trim, RegistryValueKind.String)

            regSubKey_Rn.SetValue("LowSpeed_", .MotionSys(0).NormalMotionSetting.dwLowSpeed.ToString, RegistryValueKind.String)
            regSubKey_Rn.SetValue("Speed_", .MotionSys(0).NormalMotionSetting.dwSpeed.ToString, RegistryValueKind.String)
            regSubKey_Rn.SetValue("Acc_", .MotionSys(0).NormalMotionSetting.dwAcc.ToString, RegistryValueKind.String)
            regSubKey_Rn.SetValue("Dec_", .MotionSys(0).NormalMotionSetting.dwDec.ToString, RegistryValueKind.String)
        End With

        Me.Close()

    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged

        If ListView1.SelectedItems.Count > 0 Then
            Me.TextBox1.Text = ListView1.SelectedItems(0).Text
            Me.TextBox2.Text = ListView1.SelectedItems(0).SubItems(1).Text

            LstViewSelectedIndex = Array.IndexOf(ProdID, ListView1.SelectedItems(0).Text)
        End If

    End Sub

    Private Sub btn_Edit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Edit.Click

        With Me
            If LstViewSelectedIndex < 0 Then
                MessageBox.Show("Please select a record from the list to edit.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            If .TextBox1.Text = "" Or .TextBox2.Text = "" Then
                MessageBox.Show("Record items can not be left empty.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Information)

                If .TextBox1.Text = "" Then
                    .TextBox1.Focus()
                End If

                If .TextBox2.Text = "" Then
                    .TextBox2.Focus()
                End If

                Exit Sub
            End If

            If ListView1.SelectedItems.Count > 0 Then
                ListView1.SelectedItems(0).Text = .TextBox1.Text.Trim
                ProdID(.LstViewSelectedIndex) = .TextBox1.Text.Trim

                ListView1.SelectedItems(0).SubItems(1).Text = .TextBox2.Text.Trim
                SceneNoDB(.LstViewSelectedIndex) = .TextBox2.Text.Trim

                Dim NewProdLst As String = String.Empty

                For Each element In ProdID
                    NewProdLst &= element & ","
                Next

                NewProdLst = NewProdLst.Substring(0, NewProdLst.Length - 1)
                regSubKey.SetValue("ProdID", NewProdLst)

                Dim NewSceneLst As String = String.Empty

                For Each element In SceneNoDB
                    NewSceneLst &= element & ","
                Next

                NewSceneLst = NewSceneLst.Substring(0, NewSceneLst.Length - 1)
                regSubKey.SetValue("SceneNoDB", NewSceneLst)
            End If
        End With

    End Sub

    Private Sub btn_Add_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_Add.Click

        With Me
            If Not Array.IndexOf(ProdID, .TextBox1.Text.ToUpper) < 0 Then
                MessageBox.Show("The Product ID already exist is the system. Please use 'Edit' function in this case.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            If .TextBox1.Text = "" Or .TextBox2.Text = "" Then
                MessageBox.Show("Record items can not be left empty.", "Taping-OAI...", MessageBoxButtons.OK, MessageBoxIcon.Information)

                If .TextBox1.Text = "" Then
                    .TextBox1.Focus()
                End If

                If .TextBox2.Text = "" Then
                    .TextBox2.Focus()
                End If

                Exit Sub
            End If

            Dim itm As ListViewItem = Me.ListView1.Items.Add(.TextBox1.Text.ToUpper)
            itm.SubItems.Add(.TextBox2.Text.ToUpper)

            ReDim Preserve ProdID(ProdID.Length)
            ReDim Preserve SceneNoDB(SceneNoDB.Length)

            ProdID(ProdID.GetUpperBound(0)) = .TextBox1.Text.Trim
            SceneNoDB(SceneNoDB.GetUpperBound(0)) = .TextBox2.Text.Trim

            Dim NewProdLst As String = String.Empty

            For Each element In ProdID
                NewProdLst &= element & ","
            Next

            NewProdLst = NewProdLst.Substring(0, NewProdLst.Length - 1)
            regSubKey.SetValue("ProdID", NewProdLst)

            Dim NewSceneLst As String = String.Empty

            For Each element In SceneNoDB
                NewSceneLst &= element & ","
            Next

            NewSceneLst = NewSceneLst.Substring(0, NewSceneLst.Length - 1)
            regSubKey.SetValue("SceneNoDB", NewSceneLst)
        End With

    End Sub

End Class