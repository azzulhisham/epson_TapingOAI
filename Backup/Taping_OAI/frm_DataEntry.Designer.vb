<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_DataEntry
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_DataEntry))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lbl_EnterTitle = New System.Windows.Forms.Label
        Me.txt_DataInput = New System.Windows.Forms.TextBox
        Me.lbl_MsgBox = New System.Windows.Forms.Label
        Me.lbl_AllLotNo = New System.Windows.Forms.Label
        Me.tmr_ReadIO = New System.Windows.Forms.Timer(Me.components)
        Me.chk_TapeInsp = New System.Windows.Forms.CheckBox
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PictureBox1.Location = New System.Drawing.Point(-21, -17)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(132, 190)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'lbl_EnterTitle
        '
        Me.lbl_EnterTitle.AutoSize = True
        Me.lbl_EnterTitle.Location = New System.Drawing.Point(96, 38)
        Me.lbl_EnterTitle.Name = "lbl_EnterTitle"
        Me.lbl_EnterTitle.Size = New System.Drawing.Size(247, 14)
        Me.lbl_EnterTitle.TabIndex = 1
        Me.lbl_EnterTitle.Text = "Scan P-Lot No., Press 'ENTER' to complete!"
        '
        'txt_DataInput
        '
        Me.txt_DataInput.Location = New System.Drawing.Point(97, 58)
        Me.txt_DataInput.Name = "txt_DataInput"
        Me.txt_DataInput.Size = New System.Drawing.Size(271, 22)
        Me.txt_DataInput.TabIndex = 2
        '
        'lbl_MsgBox
        '
        Me.lbl_MsgBox.Font = New System.Drawing.Font("Georgia", 15.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_MsgBox.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lbl_MsgBox.Location = New System.Drawing.Point(230, 4)
        Me.lbl_MsgBox.Name = "lbl_MsgBox"
        Me.lbl_MsgBox.Size = New System.Drawing.Size(150, 25)
        Me.lbl_MsgBox.TabIndex = 3
        Me.lbl_MsgBox.Text = "Lot No."
        Me.lbl_MsgBox.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbl_AllLotNo
        '
        Me.lbl_AllLotNo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_AllLotNo.ForeColor = System.Drawing.Color.Maroon
        Me.lbl_AllLotNo.Location = New System.Drawing.Point(97, 83)
        Me.lbl_AllLotNo.Name = "lbl_AllLotNo"
        Me.lbl_AllLotNo.Size = New System.Drawing.Size(269, 17)
        Me.lbl_AllLotNo.TabIndex = 4
        Me.lbl_AllLotNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tmr_ReadIO
        '
        '
        'chk_TapeInsp
        '
        Me.chk_TapeInsp.AutoSize = True
        Me.chk_TapeInsp.Checked = True
        Me.chk_TapeInsp.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chk_TapeInsp.Font = New System.Drawing.Font("Georgia", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_TapeInsp.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.chk_TapeInsp.Location = New System.Drawing.Point(205, 103)
        Me.chk_TapeInsp.Name = "chk_TapeInsp"
        Me.chk_TapeInsp.Size = New System.Drawing.Size(165, 19)
        Me.chk_TapeInsp.TabIndex = 5
        Me.chk_TapeInsp.Text = "Top Tape Seal Inspection"
        Me.chk_TapeInsp.UseVisualStyleBackColor = True
        '
        'frm_DataEntry
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(380, 123)
        Me.ControlBox = False
        Me.Controls.Add(Me.chk_TapeInsp)
        Me.Controls.Add(Me.txt_DataInput)
        Me.Controls.Add(Me.lbl_AllLotNo)
        Me.Controls.Add(Me.lbl_MsgBox)
        Me.Controls.Add(Me.lbl_EnterTitle)
        Me.Controls.Add(Me.PictureBox1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_DataEntry"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lbl_EnterTitle As System.Windows.Forms.Label
    Friend WithEvents txt_DataInput As System.Windows.Forms.TextBox
    Friend WithEvents lbl_MsgBox As System.Windows.Forms.Label
    Friend WithEvents lbl_AllLotNo As System.Windows.Forms.Label
    Friend WithEvents tmr_ReadIO As System.Windows.Forms.Timer
    Friend WithEvents chk_TapeInsp As System.Windows.Forms.CheckBox
End Class
