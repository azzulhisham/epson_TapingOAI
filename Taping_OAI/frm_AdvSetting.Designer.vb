<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_AdvSetting
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_AdvSetting))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.txt_Dec = New System.Windows.Forms.TextBox
        Me.txt_Acc = New System.Windows.Forms.TextBox
        Me.txt_Speed = New System.Windows.Forms.TextBox
        Me.txt_LowSpeed = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txt_MachineNo = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.txt_LotDataPath = New System.Windows.Forms.TextBox
        Me.txt_DataPath = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.btn_Save = New System.Windows.Forms.Button
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.btn_Edit = New System.Windows.Forms.Button
        Me.btn_Add = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.txt_Dec)
        Me.GroupBox1.Controls.Add(Me.txt_Acc)
        Me.GroupBox1.Controls.Add(Me.txt_Speed)
        Me.GroupBox1.Controls.Add(Me.txt_LowSpeed)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.GroupBox1.Location = New System.Drawing.Point(12, 132)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(316, 143)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Motion Parameter"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label12.Location = New System.Drawing.Point(253, 108)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(44, 14)
        Me.Label12.TabIndex = 15
        Me.Label12.Text = "pls/sec"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label11.Location = New System.Drawing.Point(253, 83)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(44, 14)
        Me.Label11.TabIndex = 14
        Me.Label11.Text = "pls/sec"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label10.Location = New System.Drawing.Point(253, 59)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(44, 14)
        Me.Label10.TabIndex = 13
        Me.Label10.Text = "pls/sec"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label9.Location = New System.Drawing.Point(253, 35)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(44, 14)
        Me.Label9.TabIndex = 12
        Me.Label9.Text = "pls/sec"
        '
        'txt_Dec
        '
        Me.txt_Dec.Location = New System.Drawing.Point(147, 105)
        Me.txt_Dec.Name = "txt_Dec"
        Me.txt_Dec.Size = New System.Drawing.Size(100, 22)
        Me.txt_Dec.TabIndex = 5
        '
        'txt_Acc
        '
        Me.txt_Acc.Location = New System.Drawing.Point(147, 80)
        Me.txt_Acc.Name = "txt_Acc"
        Me.txt_Acc.Size = New System.Drawing.Size(100, 22)
        Me.txt_Acc.TabIndex = 4
        '
        'txt_Speed
        '
        Me.txt_Speed.Location = New System.Drawing.Point(147, 56)
        Me.txt_Speed.Name = "txt_Speed"
        Me.txt_Speed.Size = New System.Drawing.Size(100, 22)
        Me.txt_Speed.TabIndex = 3
        '
        'txt_LowSpeed
        '
        Me.txt_LowSpeed.Location = New System.Drawing.Point(147, 32)
        Me.txt_LowSpeed.Name = "txt_LowSpeed"
        Me.txt_LowSpeed.Size = New System.Drawing.Size(100, 22)
        Me.txt_LowSpeed.TabIndex = 2
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label7.Location = New System.Drawing.Point(130, 107)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(11, 14)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = ":"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label8.Location = New System.Drawing.Point(130, 83)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(11, 14)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = ":"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label6.Location = New System.Drawing.Point(130, 59)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(11, 14)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = ":"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label5.Location = New System.Drawing.Point(130, 35)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(11, 14)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = ":"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(28, 108)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(75, 14)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Deceleration"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(28, 83)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 14)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Acceleration"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(28, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(42, 14)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Speed"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(28, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 14)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Low Speed"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txt_MachineNo)
        Me.GroupBox2.Controls.Add(Me.Label17)
        Me.GroupBox2.Controls.Add(Me.Label18)
        Me.GroupBox2.Controls.Add(Me.txt_LotDataPath)
        Me.GroupBox2.Controls.Add(Me.txt_DataPath)
        Me.GroupBox2.Controls.Add(Me.Label13)
        Me.GroupBox2.Controls.Add(Me.Label14)
        Me.GroupBox2.Controls.Add(Me.Label15)
        Me.GroupBox2.Controls.Add(Me.Label16)
        Me.GroupBox2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.GroupBox2.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(316, 114)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "General"
        '
        'txt_MachineNo
        '
        Me.txt_MachineNo.Location = New System.Drawing.Point(147, 78)
        Me.txt_MachineNo.Name = "txt_MachineNo"
        Me.txt_MachineNo.Size = New System.Drawing.Size(150, 22)
        Me.txt_MachineNo.TabIndex = 10
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label17.Location = New System.Drawing.Point(130, 81)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(11, 14)
        Me.Label17.TabIndex = 12
        Me.Label17.Text = ":"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label18.Location = New System.Drawing.Point(28, 81)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(74, 14)
        Me.Label18.TabIndex = 11
        Me.Label18.Text = "Machine No."
        '
        'txt_LotDataPath
        '
        Me.txt_LotDataPath.Location = New System.Drawing.Point(147, 54)
        Me.txt_LotDataPath.Name = "txt_LotDataPath"
        Me.txt_LotDataPath.Size = New System.Drawing.Size(150, 22)
        Me.txt_LotDataPath.TabIndex = 1
        '
        'txt_DataPath
        '
        Me.txt_DataPath.Location = New System.Drawing.Point(147, 30)
        Me.txt_DataPath.Name = "txt_DataPath"
        Me.txt_DataPath.Size = New System.Drawing.Size(150, 22)
        Me.txt_DataPath.TabIndex = 0
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label13.Location = New System.Drawing.Point(130, 57)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(11, 14)
        Me.Label13.TabIndex = 9
        Me.Label13.Text = ":"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label14.Location = New System.Drawing.Point(130, 33)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(11, 14)
        Me.Label14.TabIndex = 8
        Me.Label14.Text = ":"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label15.Location = New System.Drawing.Point(28, 57)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(83, 14)
        Me.Label15.TabIndex = 7
        Me.Label15.Text = "Lot Data Path"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label16.Location = New System.Drawing.Point(28, 33)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(61, 14)
        Me.Label16.TabIndex = 6
        Me.Label16.Text = "Data Path"
        '
        'btn_Save
        '
        Me.btn_Save.Location = New System.Drawing.Point(510, 247)
        Me.btn_Save.Name = "btn_Save"
        Me.btn_Save.Size = New System.Drawing.Size(80, 28)
        Me.btn_Save.TabIndex = 6
        Me.btn_Save.Text = "Save"
        Me.btn_Save.UseVisualStyleBackColor = True
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'ListView1
        '
        Me.ListView1.Location = New System.Drawing.Point(17, 27)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(163, 136)
        Me.ListView1.TabIndex = 7
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.TextBox2)
        Me.GroupBox3.Controls.Add(Me.TextBox1)
        Me.GroupBox3.Controls.Add(Me.Label22)
        Me.GroupBox3.Controls.Add(Me.Label21)
        Me.GroupBox3.Controls.Add(Me.Label20)
        Me.GroupBox3.Controls.Add(Me.Label19)
        Me.GroupBox3.Controls.Add(Me.btn_Edit)
        Me.GroupBox3.Controls.Add(Me.btn_Add)
        Me.GroupBox3.Controls.Add(Me.ListView1)
        Me.GroupBox3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.GroupBox3.Location = New System.Drawing.Point(337, 18)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(252, 216)
        Me.GroupBox3.TabIndex = 8
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Scene Setting"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(109, 191)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(133, 22)
        Me.TextBox2.TabIndex = 15
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(109, 167)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(133, 22)
        Me.TextBox1.TabIndex = 14
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.ForeColor = System.Drawing.Color.Black
        Me.Label22.Location = New System.Drawing.Point(92, 194)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(11, 14)
        Me.Label22.TabIndex = 13
        Me.Label22.Text = ":"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.ForeColor = System.Drawing.Color.Black
        Me.Label21.Location = New System.Drawing.Point(92, 170)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(11, 14)
        Me.Label21.TabIndex = 12
        Me.Label21.Text = ":"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.ForeColor = System.Drawing.Color.Black
        Me.Label20.Location = New System.Drawing.Point(15, 194)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(64, 14)
        Me.Label20.TabIndex = 11
        Me.Label20.Text = "Scene No."
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.ForeColor = System.Drawing.Color.Black
        Me.Label19.Location = New System.Drawing.Point(14, 170)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(50, 14)
        Me.Label19.TabIndex = 10
        Me.Label19.Text = "Product"
        '
        'btn_Edit
        '
        Me.btn_Edit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_Edit.Location = New System.Drawing.Point(194, 56)
        Me.btn_Edit.Name = "btn_Edit"
        Me.btn_Edit.Size = New System.Drawing.Size(48, 24)
        Me.btn_Edit.TabIndex = 9
        Me.btn_Edit.Text = "Edit"
        Me.btn_Edit.UseVisualStyleBackColor = True
        '
        'btn_Add
        '
        Me.btn_Add.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_Add.Location = New System.Drawing.Point(194, 27)
        Me.btn_Add.Name = "btn_Add"
        Me.btn_Add.Size = New System.Drawing.Size(48, 23)
        Me.btn_Add.TabIndex = 8
        Me.btn_Add.Text = "Add"
        Me.btn_Add.UseVisualStyleBackColor = True
        '
        'frm_AdvSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(601, 283)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.btn_Save)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_AdvSetting"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Taping OAI - Advance Setting"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txt_Dec As System.Windows.Forms.TextBox
    Friend WithEvents txt_Acc As System.Windows.Forms.TextBox
    Friend WithEvents txt_Speed As System.Windows.Forms.TextBox
    Friend WithEvents txt_LowSpeed As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txt_LotDataPath As System.Windows.Forms.TextBox
    Friend WithEvents txt_DataPath As System.Windows.Forms.TextBox
    Friend WithEvents btn_Save As System.Windows.Forms.Button
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents txt_MachineNo As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btn_Add As System.Windows.Forms.Button
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents btn_Edit As System.Windows.Forms.Button
End Class
