<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Step7API_EEC
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
        Me.FindParentFolder = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.txtStatus = New System.Windows.Forms.Label()
        Me.Generate = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'FindParentFolder
        '
        Me.FindParentFolder.Location = New System.Drawing.Point(12, 12)
        Me.FindParentFolder.Name = "FindParentFolder"
        Me.FindParentFolder.Size = New System.Drawing.Size(115, 46)
        Me.FindParentFolder.TabIndex = 0
        Me.FindParentFolder.Text = "Zoek hoofdmap"
        Me.FindParentFolder.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(12, 64)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(242, 20)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = "\\srvia03\WorkspaceDevelopment"
        '
        'txtStatus
        '
        Me.txtStatus.AutoSize = True
        Me.txtStatus.Location = New System.Drawing.Point(9, 87)
        Me.txtStatus.MaximumSize = New System.Drawing.Size(200, 20)
        Me.txtStatus.MinimumSize = New System.Drawing.Size(200, 20)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.Size = New System.Drawing.Size(200, 20)
        Me.txtStatus.TabIndex = 2
        '
        'Generate
        '
        Me.Generate.Location = New System.Drawing.Point(12, 110)
        Me.Generate.Name = "Generate"
        Me.Generate.Size = New System.Drawing.Size(115, 47)
        Me.Generate.TabIndex = 3
        Me.Generate.Text = "Genereren"
        Me.Generate.UseVisualStyleBackColor = True
        '
        'Step7API_EEC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(276, 167)
        Me.Controls.Add(Me.FindParentFolder)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Generate)
        Me.Name = "Step7API_EEC"
        Me.Text = "Step7Api_EEC"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FindParentFolder As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents txtStatus As System.Windows.Forms.Label
    Friend WithEvents Generate As System.Windows.Forms.Button

End Class
