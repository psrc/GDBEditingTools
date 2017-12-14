<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWalkAccess
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
        Me.cboBusStops = New System.Windows.Forms.ComboBox
        Me.cboBlocks = New System.Windows.Forms.ComboBox
        Me.cboTAZ = New System.Windows.Forms.ComboBox
        Me.cboTempStops = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.txtBusBuffer = New System.Windows.Forms.TextBox
        Me.txtTAZBuffer = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'cboBusStops
        '
        Me.cboBusStops.FormattingEnabled = True
        Me.cboBusStops.Location = New System.Drawing.Point(106, 24)
        Me.cboBusStops.Name = "cboBusStops"
        Me.cboBusStops.Size = New System.Drawing.Size(121, 21)
        Me.cboBusStops.TabIndex = 0
        '
        'cboBlocks
        '
        Me.cboBlocks.FormattingEnabled = True
        Me.cboBlocks.Location = New System.Drawing.Point(106, 66)
        Me.cboBlocks.Name = "cboBlocks"
        Me.cboBlocks.Size = New System.Drawing.Size(121, 21)
        Me.cboBlocks.TabIndex = 1
        '
        'cboTAZ
        '
        Me.cboTAZ.FormattingEnabled = True
        Me.cboTAZ.Location = New System.Drawing.Point(106, 112)
        Me.cboTAZ.Name = "cboTAZ"
        Me.cboTAZ.Size = New System.Drawing.Size(121, 21)
        Me.cboTAZ.TabIndex = 2
        '
        'cboTempStops
        '
        Me.cboTempStops.FormattingEnabled = True
        Me.cboTempStops.Location = New System.Drawing.Point(106, 156)
        Me.cboTempStops.Name = "cboTempStops"
        Me.cboTempStops.Size = New System.Drawing.Size(121, 21)
        Me.cboTempStops.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Bus Stops"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 69)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Census Blocks"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 115)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Buffered TAZ "
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 156)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Temp  Bus Stops"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(242, 32)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(86, 13)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Bus Stops Buffer"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(242, 81)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(62, 13)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "TAZ  Buffer"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(289, 132)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(133, 23)
        Me.Button1.TabIndex = 12
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtBusBuffer
        '
        Me.txtBusBuffer.Location = New System.Drawing.Point(357, 32)
        Me.txtBusBuffer.Name = "txtBusBuffer"
        Me.txtBusBuffer.Size = New System.Drawing.Size(100, 20)
        Me.txtBusBuffer.TabIndex = 13
        Me.txtBusBuffer.Text = "528"
        '
        'txtTAZBuffer
        '
        Me.txtTAZBuffer.Location = New System.Drawing.Point(357, 81)
        Me.txtTAZBuffer.Name = "txtTAZBuffer"
        Me.txtTAZBuffer.Size = New System.Drawing.Size(100, 20)
        Me.txtTAZBuffer.TabIndex = 14
        Me.txtTAZBuffer.Text = "500"
        '
        'frmWalkAccess
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(469, 262)
        Me.Controls.Add(Me.txtTAZBuffer)
        Me.Controls.Add(Me.txtBusBuffer)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboTempStops)
        Me.Controls.Add(Me.cboTAZ)
        Me.Controls.Add(Me.cboBlocks)
        Me.Controls.Add(Me.cboBusStops)
        Me.Name = "frmWalkAccess"
        Me.Text = "frmWalkAccess"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboBusStops As System.Windows.Forms.ComboBox
    Friend WithEvents cboBlocks As System.Windows.Forms.ComboBox
    Friend WithEvents cboTAZ As System.Windows.Forms.ComboBox
    Friend WithEvents cboTempStops As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtBusBuffer As System.Windows.Forms.TextBox
    Friend WithEvents txtTAZBuffer As System.Windows.Forms.TextBox
End Class
