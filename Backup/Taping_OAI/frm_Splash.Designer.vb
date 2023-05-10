<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Splash
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Splash))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lbl_BootMsg = New System.Windows.Forms.Label
        Me.btn_Close = New System.Windows.Forms.Button
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.sub_ShutdownPC = New System.Windows.Forms.ToolStripMenuItem
        Me.btn_Start = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.pic_OgColorSigOff = New System.Windows.Forms.PictureBox
        Me.pic_OgColorSigOn = New System.Windows.Forms.PictureBox
        Me.pic_RdColorSigOff = New System.Windows.Forms.PictureBox
        Me.pic_RdColorSigOn = New System.Windows.Forms.PictureBox
        Me.pic_GrColorSigOff = New System.Windows.Forms.PictureBox
        Me.pic_GrColorSigOn = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.tmr_ReadIO = New System.Windows.Forms.Timer(Me.components)
        Me.tmr_WaitKey = New System.Windows.Forms.Timer(Me.components)
        Me.tmr_Enter = New System.Windows.Forms.Timer(Me.components)
        Me.Button1 = New System.Windows.Forms.Button
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.FtpConnection1 = New EnterpriseDT.Net.Ftp.FTPConnection(Me.components)
        Me.lbl_AppVersionNo = New System.Windows.Forms.Label
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        CType(Me.pic_OgColorSigOff, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pic_OgColorSigOn, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pic_RdColorSigOff, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pic_RdColorSigOn, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pic_GrColorSigOff, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pic_GrColorSigOn, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(-117, -15)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(352, 376)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'lbl_BootMsg
        '
        Me.lbl_BootMsg.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_BootMsg.ForeColor = System.Drawing.Color.Teal
        Me.lbl_BootMsg.Location = New System.Drawing.Point(265, 101)
        Me.lbl_BootMsg.Name = "lbl_BootMsg"
        Me.lbl_BootMsg.Size = New System.Drawing.Size(171, 22)
        Me.lbl_BootMsg.TabIndex = 17
        Me.lbl_BootMsg.Text = "System is starting up..."
        Me.lbl_BootMsg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btn_Close
        '
        Me.btn_Close.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_Close.ContextMenuStrip = Me.ContextMenuStrip1
        Me.btn_Close.ForeColor = System.Drawing.Color.Maroon
        Me.btn_Close.Image = CType(resources.GetObject("btn_Close.Image"), System.Drawing.Image)
        Me.btn_Close.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_Close.Location = New System.Drawing.Point(447, 233)
        Me.btn_Close.Name = "btn_Close"
        Me.btn_Close.Padding = New System.Windows.Forms.Padding(8, 3, 3, 3)
        Me.btn_Close.Size = New System.Drawing.Size(106, 41)
        Me.btn_Close.TabIndex = 16
        Me.btn_Close.Text = "Shutdown"
        Me.btn_Close.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btn_Close.UseVisualStyleBackColor = False
        Me.btn_Close.Visible = False
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.sub_ShutdownPC})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(147, 26)
        '
        'sub_ShutdownPC
        '
        Me.sub_ShutdownPC.Name = "sub_ShutdownPC"
        Me.sub_ShutdownPC.Size = New System.Drawing.Size(146, 22)
        Me.sub_ShutdownPC.Text = "Shutdown PC"
        '
        'btn_Start
        '
        Me.btn_Start.BackColor = System.Drawing.Color.Lime
        Me.btn_Start.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Start.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btn_Start.Image = CType(resources.GetObject("btn_Start.Image"), System.Drawing.Image)
        Me.btn_Start.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_Start.Location = New System.Drawing.Point(447, 180)
        Me.btn_Start.Name = "btn_Start"
        Me.btn_Start.Padding = New System.Windows.Forms.Padding(7, 3, 7, 3)
        Me.btn_Start.Size = New System.Drawing.Size(106, 47)
        Me.btn_Start.TabIndex = 15
        Me.btn_Start.Text = "Start"
        Me.btn_Start.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btn_Start.UseVisualStyleBackColor = False
        Me.btn_Start.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Gray
        Me.Label2.Location = New System.Drawing.Point(265, 57)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(226, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Windows Embedded Standard and Windows 7"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Georgia", 9.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label1.Location = New System.Drawing.Point(489, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 15)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "az_Logics"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pic_OgColorSigOff
        '
        Me.pic_OgColorSigOff.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.pic_OgColorSigOff.Location = New System.Drawing.Point(316, 234)
        Me.pic_OgColorSigOff.Name = "pic_OgColorSigOff"
        Me.pic_OgColorSigOff.Size = New System.Drawing.Size(20, 21)
        Me.pic_OgColorSigOff.TabIndex = 25
        Me.pic_OgColorSigOff.TabStop = False
        Me.pic_OgColorSigOff.Visible = False
        '
        'pic_OgColorSigOn
        '
        Me.pic_OgColorSigOn.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.pic_OgColorSigOn.Location = New System.Drawing.Point(342, 234)
        Me.pic_OgColorSigOn.Name = "pic_OgColorSigOn"
        Me.pic_OgColorSigOn.Size = New System.Drawing.Size(20, 21)
        Me.pic_OgColorSigOn.TabIndex = 24
        Me.pic_OgColorSigOn.TabStop = False
        Me.pic_OgColorSigOn.Visible = False
        '
        'pic_RdColorSigOff
        '
        Me.pic_RdColorSigOff.BackColor = System.Drawing.Color.Maroon
        Me.pic_RdColorSigOff.Location = New System.Drawing.Point(264, 262)
        Me.pic_RdColorSigOff.Name = "pic_RdColorSigOff"
        Me.pic_RdColorSigOff.Size = New System.Drawing.Size(20, 21)
        Me.pic_RdColorSigOff.TabIndex = 23
        Me.pic_RdColorSigOff.TabStop = False
        Me.pic_RdColorSigOff.Visible = False
        '
        'pic_RdColorSigOn
        '
        Me.pic_RdColorSigOn.BackColor = System.Drawing.Color.Red
        Me.pic_RdColorSigOn.Location = New System.Drawing.Point(238, 262)
        Me.pic_RdColorSigOn.Name = "pic_RdColorSigOn"
        Me.pic_RdColorSigOn.Size = New System.Drawing.Size(20, 21)
        Me.pic_RdColorSigOn.TabIndex = 22
        Me.pic_RdColorSigOn.TabStop = False
        Me.pic_RdColorSigOn.Visible = False
        '
        'pic_GrColorSigOff
        '
        Me.pic_GrColorSigOff.BackColor = System.Drawing.Color.Green
        Me.pic_GrColorSigOff.Location = New System.Drawing.Point(264, 234)
        Me.pic_GrColorSigOff.Name = "pic_GrColorSigOff"
        Me.pic_GrColorSigOff.Size = New System.Drawing.Size(20, 21)
        Me.pic_GrColorSigOff.TabIndex = 21
        Me.pic_GrColorSigOff.TabStop = False
        Me.pic_GrColorSigOff.Visible = False
        '
        'pic_GrColorSigOn
        '
        Me.pic_GrColorSigOn.BackColor = System.Drawing.Color.Lime
        Me.pic_GrColorSigOn.Location = New System.Drawing.Point(238, 234)
        Me.pic_GrColorSigOn.Name = "pic_GrColorSigOn"
        Me.pic_GrColorSigOn.Size = New System.Drawing.Size(20, 21)
        Me.pic_GrColorSigOn.TabIndex = 20
        Me.pic_GrColorSigOn.TabStop = False
        Me.pic_GrColorSigOn.Visible = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackgroundImage = CType(resources.GetObject("PictureBox2.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PictureBox2.Location = New System.Drawing.Point(385, 73)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(296, 255)
        Me.PictureBox2.TabIndex = 26
        Me.PictureBox2.TabStop = False
        '
        'tmr_ReadIO
        '
        Me.tmr_ReadIO.Interval = 30
        '
        'tmr_WaitKey
        '
        Me.tmr_WaitKey.Interval = 200
        '
        'tmr_Enter
        '
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(268, 155)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(86, 36)
        Me.Button1.TabIndex = 27
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'PictureBox3
        '
        Me.PictureBox3.Location = New System.Drawing.Point(268, 197)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(154, 30)
        Me.PictureBox3.TabIndex = 28
        Me.PictureBox3.TabStop = False
        Me.PictureBox3.Visible = False
        '
        'FtpConnection1
        '
        Me.FtpConnection1.ParentControl = Me
        Me.FtpConnection1.TransferNotifyInterval = CType(4096, Long)
        '
        'lbl_AppVersionNo
        '
        Me.lbl_AppVersionNo.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_AppVersionNo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.lbl_AppVersionNo.Location = New System.Drawing.Point(265, 73)
        Me.lbl_AppVersionNo.Name = "lbl_AppVersionNo"
        Me.lbl_AppVersionNo.Size = New System.Drawing.Size(184, 19)
        Me.lbl_AppVersionNo.TabIndex = 29
        Me.lbl_AppVersionNo.Text = "Version No. : 2011-0328-001-001"
        Me.lbl_AppVersionNo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frm_Splash
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(558, 280)
        Me.ControlBox = False
        Me.Controls.Add(Me.lbl_AppVersionNo)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.pic_OgColorSigOff)
        Me.Controls.Add(Me.pic_OgColorSigOn)
        Me.Controls.Add(Me.pic_RdColorSigOff)
        Me.Controls.Add(Me.pic_RdColorSigOn)
        Me.Controls.Add(Me.pic_GrColorSigOn)
        Me.Controls.Add(Me.pic_GrColorSigOff)
        Me.Controls.Add(Me.lbl_BootMsg)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btn_Start)
        Me.Controls.Add(Me.btn_Close)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox2)
        Me.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frm_Splash"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        CType(Me.pic_OgColorSigOff, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pic_OgColorSigOn, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pic_RdColorSigOff, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pic_RdColorSigOn, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pic_GrColorSigOff, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pic_GrColorSigOn, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lbl_BootMsg As System.Windows.Forms.Label
    Friend WithEvents btn_Close As System.Windows.Forms.Button
    Friend WithEvents btn_Start As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pic_OgColorSigOff As System.Windows.Forms.PictureBox
    Friend WithEvents pic_OgColorSigOn As System.Windows.Forms.PictureBox
    Friend WithEvents pic_RdColorSigOff As System.Windows.Forms.PictureBox
    Friend WithEvents pic_RdColorSigOn As System.Windows.Forms.PictureBox
    Friend WithEvents pic_GrColorSigOff As System.Windows.Forms.PictureBox
    Friend WithEvents pic_GrColorSigOn As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents tmr_ReadIO As System.Windows.Forms.Timer
    Friend WithEvents tmr_WaitKey As System.Windows.Forms.Timer
    Friend WithEvents tmr_Enter As System.Windows.Forms.Timer
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents sub_ShutdownPC As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents FtpConnection1 As EnterpriseDT.Net.Ftp.FTPConnection
    Friend WithEvents lbl_AppVersionNo As System.Windows.Forms.Label

End Class
