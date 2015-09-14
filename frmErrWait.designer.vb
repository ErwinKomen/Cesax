<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmErrWait
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmErrWait))
    Me.cmdOk = New System.Windows.Forms.Button
    Me.SuspendLayout()
    '
    'cmdOk
    '
    Me.cmdOk.Location = New System.Drawing.Point(129, 12)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 0
    Me.cmdOk.Text = "OK"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'frmErrWait
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(337, 49)
    Me.Controls.Add(Me.cmdOk)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmErrWait"
    Me.Text = "Press OK when ready reviewing the error"
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents cmdOk As System.Windows.Forms.Button
End Class
