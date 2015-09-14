<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFeatAdd
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
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFeatAdd))
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.tbName = New System.Windows.Forms.TextBox()
    Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
    Me.Label1 = New System.Windows.Forms.Label()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.tbValue = New System.Windows.Forms.TextBox()
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.SuspendLayout()
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Location = New System.Drawing.Point(143, 115)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 3
    Me.cmdOk.Text = "Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(62, 115)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 4
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'tbName
    '
    Me.tbName.Location = New System.Drawing.Point(15, 25)
    Me.tbName.Name = "tbName"
    Me.tbName.Size = New System.Drawing.Size(204, 20)
    Me.tbName.TabIndex = 1
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 9)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(162, 13)
    Me.Label1.TabIndex = 6
    Me.Label1.Text = "Name of the feature (no spaces):"
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 63)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(142, 13)
    Me.Label2.TabIndex = 8
    Me.Label2.Text = "Default value for all features:"
    '
    'tbValue
    '
    Me.tbValue.Location = New System.Drawing.Point(15, 79)
    Me.tbValue.Name = "tbValue"
    Me.tbValue.Size = New System.Drawing.Size(204, 20)
    Me.tbValue.TabIndex = 2
    '
    'Timer1
    '
    '
    'frmFeatAdd
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(230, 150)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.tbValue)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.tbName)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmFeatAdd"
    Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Adding a feature"
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents tbName As System.Windows.Forms.TextBox
  Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents tbValue As System.Windows.Forms.TextBox
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
