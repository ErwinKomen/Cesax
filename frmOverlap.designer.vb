<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOverlap
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOverlap))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.cmdAccept = New System.Windows.Forms.Button
    Me.cmdExit = New System.Windows.Forms.Button
    Me.Label1 = New System.Windows.Forms.Label
    Me.tbChainMain = New System.Windows.Forms.RichTextBox
    Me.tbChainOverlap = New System.Windows.Forms.RichTextBox
    Me.Label2 = New System.Windows.Forms.Label
    Me.cmdReject = New System.Windows.Forms.Button
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
    Me.Label3 = New System.Windows.Forms.Label
    Me.tbSrcNode = New System.Windows.Forms.RichTextBox
    Me.tbDstNode = New System.Windows.Forms.RichTextBox
    Me.Label4 = New System.Windows.Forms.Label
    Me.StatusStrip1.SuspendLayout()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 359)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(571, 22)
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
    'Timer1
    '
    '
    'cmdAccept
    '
    Me.cmdAccept.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdAccept.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdAccept.Location = New System.Drawing.Point(242, 304)
    Me.cmdAccept.Name = "cmdAccept"
    Me.cmdAccept.Size = New System.Drawing.Size(75, 23)
    Me.cmdAccept.TabIndex = 1
    Me.cmdAccept.Text = "&Accept"
    Me.cmdAccept.UseVisualStyleBackColor = True
    '
    'cmdExit
    '
    Me.cmdExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdExit.Location = New System.Drawing.Point(242, 333)
    Me.cmdExit.Name = "cmdExit"
    Me.cmdExit.Size = New System.Drawing.Size(75, 23)
    Me.cmdExit.TabIndex = 2
    Me.cmdExit.Text = "E&xit"
    Me.cmdExit.UseVisualStyleBackColor = True
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(3, 85)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(62, 13)
    Me.Label1.TabIndex = 3
    Me.Label1.Text = "Main chain:"
    '
    'tbChainMain
    '
    Me.tbChainMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbChainMain.Location = New System.Drawing.Point(3, 101)
    Me.tbChainMain.Name = "tbChainMain"
    Me.tbChainMain.Size = New System.Drawing.Size(241, 255)
    Me.tbChainMain.TabIndex = 4
    Me.tbChainMain.Text = ""
    '
    'tbChainOverlap
    '
    Me.tbChainOverlap.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbChainOverlap.Location = New System.Drawing.Point(3, 101)
    Me.tbChainOverlap.Name = "tbChainOverlap"
    Me.tbChainOverlap.Size = New System.Drawing.Size(233, 255)
    Me.tbChainOverlap.TabIndex = 5
    Me.tbChainOverlap.Text = ""
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(3, 85)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(143, 13)
    Me.Label2.TabIndex = 6
    Me.Label2.Text = "Combine into the main chain:"
    '
    'cmdReject
    '
    Me.cmdReject.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdReject.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdReject.Location = New System.Drawing.Point(242, 275)
    Me.cmdReject.Name = "cmdReject"
    Me.cmdReject.Size = New System.Drawing.Size(75, 23)
    Me.cmdReject.TabIndex = 7
    Me.cmdReject.Text = "&Reject"
    Me.cmdReject.UseVisualStyleBackColor = True
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer1.Name = "SplitContainer1"
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.tbSrcNode)
    Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
    Me.SplitContainer1.Panel1.Controls.Add(Me.tbChainMain)
    Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbDstNode)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbChainOverlap)
    Me.SplitContainer1.Panel2.Controls.Add(Me.cmdReject)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
    Me.SplitContainer1.Panel2.Controls.Add(Me.cmdExit)
    Me.SplitContainer1.Panel2.Controls.Add(Me.cmdAccept)
    Me.SplitContainer1.Size = New System.Drawing.Size(571, 359)
    Me.SplitContainer1.SplitterDistance = 247
    Me.SplitContainer1.TabIndex = 8
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(3, 9)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(161, 13)
    Me.Label3.TabIndex = 5
    Me.Label3.Text = "Change the antecedent of node:"
    '
    'tbSrcNode
    '
    Me.tbSrcNode.Location = New System.Drawing.Point(3, 25)
    Me.tbSrcNode.Name = "tbSrcNode"
    Me.tbSrcNode.Size = New System.Drawing.Size(241, 57)
    Me.tbSrcNode.TabIndex = 6
    Me.tbSrcNode.Text = ""
    '
    'tbDstNode
    '
    Me.tbDstNode.Location = New System.Drawing.Point(3, 25)
    Me.tbDstNode.Name = "tbDstNode"
    Me.tbDstNode.Size = New System.Drawing.Size(233, 57)
    Me.tbDstNode.TabIndex = 8
    Me.tbDstNode.Text = ""
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(3, 9)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(155, 13)
    Me.Label4.TabIndex = 9
    Me.Label4.Text = "The new antecedent becomes:"
    '
    'frmOverlap
    '
    Me.AcceptButton = Me.cmdAccept
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdExit
    Me.ClientSize = New System.Drawing.Size(571, 381)
    Me.Controls.Add(Me.SplitContainer1)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmOverlap"
    Me.Text = "Check for overlapping chains..."
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel1.PerformLayout()
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    Me.SplitContainer1.Panel2.PerformLayout()
    Me.SplitContainer1.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents cmdAccept As System.Windows.Forms.Button
  Friend WithEvents cmdExit As System.Windows.Forms.Button
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents tbChainMain As System.Windows.Forms.RichTextBox
  Friend WithEvents tbChainOverlap As System.Windows.Forms.RichTextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents cmdReject As System.Windows.Forms.Button
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents tbSrcNode As System.Windows.Forms.RichTextBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents tbDstNode As System.Windows.Forms.RichTextBox
End Class
