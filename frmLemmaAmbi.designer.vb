<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLemmaAmbi
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLemmaAmbi))
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.tbVern = New System.Windows.Forms.TextBox
    Me.Label1 = New System.Windows.Forms.Label
    Me.Label2 = New System.Windows.Forms.Label
    Me.tbPos = New System.Windows.Forms.TextBox
    Me.lboLemma = New System.Windows.Forms.ListBox
    Me.Label3 = New System.Windows.Forms.Label
    Me.cmdOk = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.Label4 = New System.Windows.Forms.Label
    Me.cmdSkip = New System.Windows.Forms.Button
    Me.Label5 = New System.Windows.Forms.Label
    Me.tbFeat = New System.Windows.Forms.TextBox
    Me.wbDict = New System.Windows.Forms.WebBrowser
    Me.tbDict = New System.Windows.Forms.TextBox
    Me.StatusStrip1.SuspendLayout()
    Me.SuspendLayout()
    '
    'Timer1
    '
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 295)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(613, 22)
    Me.StatusStrip1.TabIndex = 0
    Me.StatusStrip1.Text = "StatusStrip1"
    '
    'ToolStripStatusLabel1
    '
    Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
    Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(19, 17)
    Me.ToolStripStatusLabel1.Text = "..."
    '
    'ToolStripProgressBar1
    '
    Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
    Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 16)
    '
    'tbVern
    '
    Me.tbVern.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbVern.Location = New System.Drawing.Point(105, 12)
    Me.tbVern.Name = "tbVern"
    Me.tbVern.Size = New System.Drawing.Size(508, 20)
    Me.tbVern.TabIndex = 1
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 15)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(87, 13)
    Me.Label1.TabIndex = 2
    Me.Label1.Text = "Vernacular word:"
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 41)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(81, 13)
    Me.Label2.TabIndex = 4
    Me.Label2.Text = "Part-of-Speech:"
    '
    'tbPos
    '
    Me.tbPos.Location = New System.Drawing.Point(105, 38)
    Me.tbPos.Name = "tbPos"
    Me.tbPos.Size = New System.Drawing.Size(70, 20)
    Me.tbPos.TabIndex = 3
    '
    'lboLemma
    '
    Me.lboLemma.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.lboLemma.FormattingEnabled = True
    Me.lboLemma.Location = New System.Drawing.Point(15, 97)
    Me.lboLemma.Name = "lboLemma"
    Me.lboLemma.Size = New System.Drawing.Size(216, 160)
    Me.lboLemma.TabIndex = 5
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(12, 81)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(89, 13)
    Me.Label3.TabIndex = 6
    Me.Label3.Text = "Possible lemma's:"
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Location = New System.Drawing.Point(538, 269)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 7
    Me.cmdOk.Text = "&Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(376, 269)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 8
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(247, 77)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(142, 13)
    Me.Label4.TabIndex = 10
    Me.Label4.Text = "Dictionary entry (if available):"
    '
    'cmdSkip
    '
    Me.cmdSkip.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdSkip.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdSkip.Location = New System.Drawing.Point(457, 269)
    Me.cmdSkip.Name = "cmdSkip"
    Me.cmdSkip.Size = New System.Drawing.Size(75, 23)
    Me.cmdSkip.TabIndex = 11
    Me.cmdSkip.Text = "S&kip"
    Me.cmdSkip.UseVisualStyleBackColor = True
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(193, 41)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(57, 13)
    Me.Label5.TabIndex = 13
    Me.Label5.Text = "Feature(s):"
    '
    'tbFeat
    '
    Me.tbFeat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeat.Location = New System.Drawing.Point(256, 38)
    Me.tbFeat.Name = "tbFeat"
    Me.tbFeat.Size = New System.Drawing.Size(357, 20)
    Me.tbFeat.TabIndex = 12
    '
    'wbDict
    '
    Me.wbDict.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.wbDict.Location = New System.Drawing.Point(250, 97)
    Me.wbDict.MinimumSize = New System.Drawing.Size(20, 20)
    Me.wbDict.Name = "wbDict"
    Me.wbDict.Size = New System.Drawing.Size(363, 166)
    Me.wbDict.TabIndex = 14
    '
    'tbDict
    '
    Me.tbDict.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDict.Location = New System.Drawing.Point(395, 74)
    Me.tbDict.Name = "tbDict"
    Me.tbDict.Size = New System.Drawing.Size(218, 20)
    Me.tbDict.TabIndex = 15
    '
    'frmLemmaAmbi
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(613, 317)
    Me.Controls.Add(Me.tbDict)
    Me.Controls.Add(Me.wbDict)
    Me.Controls.Add(Me.Label5)
    Me.Controls.Add(Me.tbFeat)
    Me.Controls.Add(Me.cmdSkip)
    Me.Controls.Add(Me.Label4)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.lboLemma)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.tbPos)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.tbVern)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmLemmaAmbi"
    Me.Text = "Lemma disambiguation"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents tbVern As System.Windows.Forms.TextBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents tbPos As System.Windows.Forms.TextBox
  Friend WithEvents lboLemma As System.Windows.Forms.ListBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents cmdSkip As System.Windows.Forms.Button
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents tbFeat As System.Windows.Forms.TextBox
  Friend WithEvents wbDict As System.Windows.Forms.WebBrowser
  Friend WithEvents tbDict As System.Windows.Forms.TextBox
End Class
