<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNodeScanner
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cboLayer = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.chkFixIJNode = New System.Windows.Forms.CheckBox()
        Me.chkSwitchAttributes = New System.Windows.Forms.CheckBox()
        Me.chkSwitchNodes = New System.Windows.Forms.CheckBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.btnFlipEdges = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cboLayer
        '
        Me.cboLayer.FormattingEnabled = True
        Me.cboLayer.Location = New System.Drawing.Point(122, 27)
        Me.cboLayer.Name = "cboLayer"
        Me.cboLayer.Size = New System.Drawing.Size(121, 21)
        Me.cboLayer.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Layer to Scan:"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(135, 157)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(145, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Check Edges & Projects"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'chkFixIJNode
        '
        Me.chkFixIJNode.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.chkFixIJNode.AutoSize = True
        Me.chkFixIJNode.Location = New System.Drawing.Point(26, 101)
        Me.chkFixIJNode.Name = "chkFixIJNode"
        Me.chkFixIJNode.Size = New System.Drawing.Size(123, 17)
        Me.chkFixIJNode.TabIndex = 4
        Me.chkFixIJNode.Text = "Fix invalid I & J Nodes"
        Me.chkFixIJNode.UseVisualStyleBackColor = True
        '
        'chkSwitchAttributes
        '
        Me.chkSwitchAttributes.AutoSize = True
        Me.chkSwitchAttributes.Location = New System.Drawing.Point(26, 124)
        Me.chkSwitchAttributes.Name = "chkSwitchAttributes"
        Me.chkSwitchAttributes.Size = New System.Drawing.Size(136, 17)
        Me.chkSwitchAttributes.TabIndex = 5
        Me.chkSwitchAttributes.Text = "Switch Edge Attributes "
        Me.chkSwitchAttributes.UseVisualStyleBackColor = True
        '
        'chkSwitchNodes
        '
        Me.chkSwitchNodes.AutoSize = True
        Me.chkSwitchNodes.Location = New System.Drawing.Point(26, 78)
        Me.chkSwitchNodes.Name = "chkSwitchNodes"
        Me.chkSwitchNodes.Size = New System.Drawing.Size(103, 17)
        Me.chkSwitchNodes.TabIndex = 6
        Me.chkSwitchNodes.Text = "Switch IJ Nodes"
        Me.chkSwitchNodes.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(135, 186)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(145, 23)
        Me.Button2.TabIndex = 7
        Me.Button2.Text = "Check Tranist  Points"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'btnFlipEdges
        '
        Me.btnFlipEdges.Location = New System.Drawing.Point(135, 215)
        Me.btnFlipEdges.Name = "btnFlipEdges"
        Me.btnFlipEdges.Size = New System.Drawing.Size(145, 23)
        Me.btnFlipEdges.TabIndex = 8
        Me.btnFlipEdges.Text = "FlipEdges"
        Me.btnFlipEdges.UseVisualStyleBackColor = True
        '
        'frmNodeScanner
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.btnFlipEdges)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.chkSwitchNodes)
        Me.Controls.Add(Me.chkSwitchAttributes)
        Me.Controls.Add(Me.chkFixIJNode)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboLayer)
        Me.Name = "frmNodeScanner"
        Me.Text = "frmNodeScanner"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboLayer As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents chkFixIJNode As System.Windows.Forms.CheckBox
    Friend WithEvents chkSwitchAttributes As System.Windows.Forms.CheckBox
    Friend WithEvents chkSwitchNodes As System.Windows.Forms.CheckBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents btnFlipEdges As System.Windows.Forms.Button
End Class
