<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFind
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
    Me.lbSearch = New System.Windows.Forms.Label
    Me.tbFind = New System.Windows.Forms.TextBox
    Me.cmdPrev = New System.Windows.Forms.Button
    Me.cmdNext = New System.Windows.Forms.Button
    Me.chbWhole = New System.Windows.Forms.CheckBox
    Me.chbCase = New System.Windows.Forms.CheckBox
    Me.SuspendLayout()
    '
    'lbSearch
    '
    Me.lbSearch.AutoSize = True
    Me.lbSearch.Location = New System.Drawing.Point(0, 9)
    Me.lbSearch.Name = "lbSearch"
    Me.lbSearch.Size = New System.Drawing.Size(59, 13)
    Me.lbSearch.TabIndex = 0
    Me.lbSearch.Text = "Search for:"
    '
    'tbFind
    '
    Me.tbFind.Location = New System.Drawing.Point(2, 25)
    Me.tbFind.Name = "tbFind"
    Me.tbFind.Size = New System.Drawing.Size(395, 20)
    Me.tbFind.TabIndex = 1
    '
    'cmdPrev
    '
    Me.cmdPrev.Location = New System.Drawing.Point(234, 60)
    Me.cmdPrev.Name = "cmdPrev"
    Me.cmdPrev.Size = New System.Drawing.Size(79, 23)
    Me.cmdPrev.TabIndex = 2
    Me.cmdPrev.Text = "&Previous"
    Me.cmdPrev.UseVisualStyleBackColor = True
    '
    'cmdNext
    '
    Me.cmdNext.Location = New System.Drawing.Point(319, 60)
    Me.cmdNext.Name = "cmdNext"
    Me.cmdNext.Size = New System.Drawing.Size(79, 23)
    Me.cmdNext.TabIndex = 3
    Me.cmdNext.Text = "&Next"
    Me.cmdNext.UseVisualStyleBackColor = True
    '
    'chbWhole
    '
    Me.chbWhole.AutoSize = True
    Me.chbWhole.Location = New System.Drawing.Point(4, 60)
    Me.chbWhole.Name = "chbWhole"
    Me.chbWhole.Size = New System.Drawing.Size(88, 17)
    Me.chbWhole.TabIndex = 4
    Me.chbWhole.Text = "&Whole words"
    Me.chbWhole.UseVisualStyleBackColor = True
    '
    'chbCase
    '
    Me.chbCase.AutoSize = True
    Me.chbCase.Location = New System.Drawing.Point(104, 59)
    Me.chbCase.Name = "chbCase"
    Me.chbCase.Size = New System.Drawing.Size(94, 17)
    Me.chbCase.TabIndex = 5
    Me.chbCase.Text = "&Case sensitive"
    Me.chbCase.UseVisualStyleBackColor = True
    '
    'frmFind
    '
    Me.AcceptButton = Me.cmdNext
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(401, 101)
    Me.ControlBox = False
    Me.Controls.Add(Me.chbCase)
    Me.Controls.Add(Me.chbWhole)
    Me.Controls.Add(Me.cmdNext)
    Me.Controls.Add(Me.cmdPrev)
    Me.Controls.Add(Me.tbFind)
    Me.Controls.Add(Me.lbSearch)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.KeyPreview = True
    Me.MaximizeBox = False
    Me.MaximumSize = New System.Drawing.Size(407, 126)
    Me.MinimizeBox = False
    Me.Name = "frmFind"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Search"
    Me.TopMost = True
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents lbSearch As System.Windows.Forms.Label
  Friend WithEvents tbFind As System.Windows.Forms.TextBox
  Friend WithEvents cmdPrev As System.Windows.Forms.Button
  Friend WithEvents cmdNext As System.Windows.Forms.Button
  Friend WithEvents chbWhole As System.Windows.Forms.CheckBox
  Friend WithEvents chbCase As System.Windows.Forms.CheckBox
End Class
