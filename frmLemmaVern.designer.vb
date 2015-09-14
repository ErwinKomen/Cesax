<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLemmaVern
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLemmaVern))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
    Me.dgvLemmaVern = New System.Windows.Forms.DataGridView()
    Me.Label6 = New System.Windows.Forms.Label()
    Me.cboAction = New System.Windows.Forms.ComboBox()
    Me.Label5 = New System.Windows.Forms.Label()
    Me.tbPos = New System.Windows.Forms.TextBox()
    Me.Label4 = New System.Windows.Forms.Label()
    Me.tbPosHead = New System.Windows.Forms.TextBox()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.tbVern = New System.Windows.Forms.TextBox()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.tbLemma = New System.Windows.Forms.TextBox()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.chbAllDelete = New System.Windows.Forms.CheckBox()
    Me.StatusStrip1.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    CType(Me.dgvLemmaVern, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 719)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(775, 22)
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
    'GroupBox1
    '
    Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox1.Controls.Add(Me.SplitContainer1)
    Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(682, 704)
    Me.GroupBox1.TabIndex = 1
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Lemma's that might be added"
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer1.Location = New System.Drawing.Point(3, 16)
    Me.SplitContainer1.Name = "SplitContainer1"
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.dgvLemmaVern)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.chbAllDelete)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
    Me.SplitContainer1.Panel2.Controls.Add(Me.cboAction)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbPos)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbPosHead)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbVern)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbLemma)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
    Me.SplitContainer1.Size = New System.Drawing.Size(676, 685)
    Me.SplitContainer1.SplitterDistance = 386
    Me.SplitContainer1.TabIndex = 0
    '
    'dgvLemmaVern
    '
    Me.dgvLemmaVern.AllowUserToAddRows = False
    Me.dgvLemmaVern.AllowUserToDeleteRows = False
    Me.dgvLemmaVern.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvLemmaVern.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvLemmaVern.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgvLemmaVern.Location = New System.Drawing.Point(0, 0)
    Me.dgvLemmaVern.Name = "dgvLemmaVern"
    Me.dgvLemmaVern.ReadOnly = True
    Me.dgvLemmaVern.Size = New System.Drawing.Size(386, 685)
    Me.dgvLemmaVern.TabIndex = 0
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.ForeColor = System.Drawing.Color.Maroon
    Me.Label6.Location = New System.Drawing.Point(4, 161)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(193, 13)
    Me.Label6.TabIndex = 11
    Me.Label6.Text = "Accept or adapt the action to be taken:"
    '
    'cboAction
    '
    Me.cboAction.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboAction.FormattingEnabled = True
    Me.cboAction.Location = New System.Drawing.Point(70, 177)
    Me.cboAction.Name = "cboAction"
    Me.cboAction.Size = New System.Drawing.Size(213, 21)
    Me.cboAction.TabIndex = 10
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(4, 180)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(40, 13)
    Me.Label5.TabIndex = 9
    Me.Label5.Text = "Action:"
    '
    'tbPos
    '
    Me.tbPos.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbPos.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbPos.Location = New System.Drawing.Point(70, 93)
    Me.tbPos.Name = "tbPos"
    Me.tbPos.ReadOnly = True
    Me.tbPos.Size = New System.Drawing.Size(78, 20)
    Me.tbPos.TabIndex = 7
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(3, 96)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(57, 13)
    Me.Label4.TabIndex = 6
    Me.Label4.Text = "Vern POS:"
    '
    'tbPosHead
    '
    Me.tbPosHead.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbPosHead.Location = New System.Drawing.Point(70, 33)
    Me.tbPosHead.Name = "tbPosHead"
    Me.tbPosHead.Size = New System.Drawing.Size(78, 20)
    Me.tbPosHead.TabIndex = 5
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(3, 36)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(69, 13)
    Me.Label3.TabIndex = 4
    Me.Label3.Text = "Lemma POS:"
    '
    'tbVern
    '
    Me.tbVern.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbVern.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbVern.Location = New System.Drawing.Point(70, 67)
    Me.tbVern.Name = "tbVern"
    Me.tbVern.ReadOnly = True
    Me.tbVern.Size = New System.Drawing.Size(213, 20)
    Me.tbVern.TabIndex = 3
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(3, 70)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(58, 13)
    Me.Label2.TabIndex = 2
    Me.Label2.Text = "Vernacular"
    '
    'tbLemma
    '
    Me.tbLemma.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbLemma.Location = New System.Drawing.Point(70, 7)
    Me.tbLemma.Name = "tbLemma"
    Me.tbLemma.Size = New System.Drawing.Size(213, 20)
    Me.tbLemma.TabIndex = 1
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(3, 10)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(41, 13)
    Me.Label1.TabIndex = 0
    Me.Label1.Text = "Lemma"
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Location = New System.Drawing.Point(700, 693)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 2
    Me.cmdOk.Text = "&Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(700, 664)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 3
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'Timer1
    '
    '
    'chbAllDelete
    '
    Me.chbAllDelete.AutoSize = True
    Me.chbAllDelete.Location = New System.Drawing.Point(7, 223)
    Me.chbAllDelete.Name = "chbAllDelete"
    Me.chbAllDelete.Size = New System.Drawing.Size(151, 17)
    Me.chbAllDelete.TabIndex = 12
    Me.chbAllDelete.Text = "Mark all entries as [Delete]"
    Me.chbAllDelete.UseVisualStyleBackColor = True
    '
    'frmLemmaVern
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(775, 741)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmLemmaVern"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Lemma addition specifier"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    Me.SplitContainer1.Panel2.PerformLayout()
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer1.ResumeLayout(False)
    CType(Me.dgvLemmaVern, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents dgvLemmaVern As System.Windows.Forms.DataGridView
  Friend WithEvents tbPos As System.Windows.Forms.TextBox
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents tbPosHead As System.Windows.Forms.TextBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents tbVern As System.Windows.Forms.TextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents tbLemma As System.Windows.Forms.TextBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents cboAction As System.Windows.Forms.ComboBox
  Friend WithEvents chbAllDelete As System.Windows.Forms.CheckBox
End Class
