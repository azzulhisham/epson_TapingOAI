<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_QtyEntry
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_QtyEntry))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lbl_MsgBox = New System.Windows.Forms.Label
        Me.txt_DataInput = New System.Windows.Forms.TextBox
        Me.tmr_ReadIO = New System.Windows.Forms.Timer(Me.components)
        Me.lbl_PLotNo = New System.Windows.Forms.Label
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PictureBox1.Location = New System.Drawing.Point(-37, -43)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(93, 135)
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'lbl_MsgBox
        '
        Me.lbl_MsgBox.Font = New System.Drawing.Font("Georgia", 12.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_MsgBox.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lbl_MsgBox.Location = New System.Drawing.Point(84, 1)
        Me.lbl_MsgBox.Name = "lbl_MsgBox"
        Me.lbl_MsgBox.Size = New System.Drawing.Size(95, 23)
        Me.lbl_MsgBox.TabIndex = 4
        Me.lbl_MsgBox.Text = "Qty. Input"
        Me.lbl_MsgBox.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txt_DataInput
        '
        Me.txt_DataInput.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_DataInput.Location = New System.Drawing.Point(62, 27)
        Me.txt_DataInput.Name = "txt_DataInput"
        Me.txt_DataInput.Size = New System.Drawing.Size(117, 26)
        Me.txt_DataInput.TabIndex = 5
        Me.txt_DataInput.Text = "0"
        Me.txt_DataInput.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'tmr_ReadIO
        '
        '
        'lbl_PLotNo
        '
        Me.lbl_PLotNo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_PLotNo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lbl_PLotNo.Location = New System.Drawing.Point(62, 56)
        Me.lbl_PLotNo.Name = "lbl_PLotNo"
        Me.lbl_PLotNo.Size = New System.Drawing.Size(117, 16)
        Me.lbl_PLotNo.TabIndex = 6
        Me.lbl_PLotNo.Text = "Pxx-00000"
        Me.lbl_PLotNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frm_QtyEntry
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(185, 77)
        Me.ControlBox = False
        Me.Controls.Add(Me.lbl_PLotNo)
        Me.Controls.Add(Me.txt_DataInput)
        Me.Controls.Add(Me.lbl_MsgBox)
        Me.Controls.Add(Me.PictureBox1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frm_QtyEntry"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lbl_MsgBox As System.Windows.Forms.Label
    Friend WithEvents txt_DataInput As System.Windows.Forms.TextBox
    Friend WithEvents tmr_ReadIO As System.Windows.Forms.Timer
    Friend WithEvents lbl_PLotNo As System.Windows.Forms.Label
End Class
