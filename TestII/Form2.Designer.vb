<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
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
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.txtFilename = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(30, 92)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(97, 69)
        Me.btnBrowse.TabIndex = 0
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(161, 92)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(97, 69)
        Me.btnImport.TabIndex = 1
        Me.btnImport.Text = "Import"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'txtFilename
        '
        Me.txtFilename.Location = New System.Drawing.Point(30, 12)
        Me.txtFilename.Multiline = True
        Me.txtFilename.Name = "txtFilename"
        Me.txtFilename.Size = New System.Drawing.Size(209, 59)
        Me.txtFilename.TabIndex = 2
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.txtFilename)
        Me.Controls.Add(Me.btnImport)
        Me.Controls.Add(Me.btnBrowse)
        Me.Name = "Form2"
        Me.Text = "Form2"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents btnImport As System.Windows.Forms.Button
    Friend WithEvents txtFilename As System.Windows.Forms.TextBox
End Class
