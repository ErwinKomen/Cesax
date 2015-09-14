<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmError
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmError))
    Me.Label1 = New System.Windows.Forms.Label
    Me.tbErrText = New System.Windows.Forms.RichTextBox
    Me.cmdContinue = New System.Windows.Forms.Button
    Me.cmdInterrupt = New System.Windows.Forms.Button
    Me.cmdSave = New System.Windows.Forms.Button
    Me.cmdExit = New System.Windows.Forms.Button
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.GroupBox1.SuspendLayout()
    Me.SuspendLayout()
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 9)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(216, 13)
    Me.Label1.TabIndex = 0
    Me.Label1.Text = "An error has occurred with the following text:"
    '
    'tbErrText
    '
    Me.tbErrText.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbErrText.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbErrText.Location = New System.Drawing.Point(15, 25)
    Me.tbErrText.Name = "tbErrText"
    Me.tbErrText.Size = New System.Drawing.Size(797, 210)
    Me.tbErrText.TabIndex = 1
    Me.tbErrText.Text = ""
    '
    'cmdContinue
    '
    Me.cmdContinue.Location = New System.Drawing.Point(6, 19)
    Me.cmdContinue.Name = "cmdContinue"
    Me.cmdContinue.Size = New System.Drawing.Size(155, 23)
    Me.cmdContinue.TabIndex = 3
    Me.cmdContinue.Text = "&Continue anyway"
    Me.cmdContinue.UseVisualStyleBackColor = True
    '
    'cmdInterrupt
    '
    Me.cmdInterrupt.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdInterrupt.Location = New System.Drawing.Point(167, 19)
    Me.cmdInterrupt.Name = "cmdInterrupt"
    Me.cmdInterrupt.Size = New System.Drawing.Size(155, 23)
    Me.cmdInterrupt.TabIndex = 4
    Me.cmdInterrupt.Text = "&Interrupt current action"
    Me.cmdInterrupt.UseVisualStyleBackColor = True
    '
    'cmdSave
    '
    Me.cmdSave.Location = New System.Drawing.Point(6, 48)
    Me.cmdSave.Name = "cmdSave"
    Me.cmdSave.Size = New System.Drawing.Size(155, 23)
    Me.cmdSave.TabIndex = 5
    Me.cmdSave.Text = "&Save work and exit"
    Me.cmdSave.UseVisualStyleBackColor = True
    '
    'cmdExit
    '
    Me.cmdExit.Location = New System.Drawing.Point(167, 48)
    Me.cmdExit.Name = "cmdExit"
    Me.cmdExit.Size = New System.Drawing.Size(155, 23)
    Me.cmdExit.TabIndex = 6
    Me.cmdExit.Text = "&Exit without saving"
    Me.cmdExit.UseVisualStyleBackColor = True
    '
    'GroupBox1
    '
    Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox1.Controls.Add(Me.cmdContinue)
    Me.GroupBox1.Controls.Add(Me.cmdExit)
    Me.GroupBox1.Controls.Add(Me.cmdInterrupt)
    Me.GroupBox1.Controls.Add(Me.cmdSave)
    Me.GroupBox1.Location = New System.Drawing.Point(475, 241)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(337, 83)
    Me.GroupBox1.TabIndex = 7
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "What would you like to do?"
    '
    'frmError
    '
    Me.AcceptButton = Me.cmdContinue
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdInterrupt
    Me.ClientSize = New System.Drawing.Size(824, 336)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.tbErrText)
    Me.Controls.Add(Me.Label1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmError"
    Me.Text = "Take a deep breath..."
    Me.GroupBox1.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents tbErrText As System.Windows.Forms.RichTextBox
  Friend WithEvents cmdContinue As System.Windows.Forms.Button
  Friend WithEvents cmdInterrupt As System.Windows.Forms.Button
  Friend WithEvents cmdSave As System.Windows.Forms.Button
  Friend WithEvents cmdExit As System.Windows.Forms.Button
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
End Class
