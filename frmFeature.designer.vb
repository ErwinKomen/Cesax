<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFeature
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFeature))
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.Label1 = New System.Windows.Forms.Label()
    Me.dgvFeature = New System.Windows.Forms.DataGridView()
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
    Me.tbFeatContent = New System.Windows.Forms.TextBox()
    Me.tbFeatValue = New System.Windows.Forms.RichTextBox()
    Me.Label6 = New System.Windows.Forms.Label()
    Me.Label5 = New System.Windows.Forms.Label()
    Me.tbFeatName = New System.Windows.Forms.TextBox()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.cmdDel = New System.Windows.Forms.Button()
    Me.cmdAdd = New System.Windows.Forms.Button()
    Me.tbFeatType = New System.Windows.Forms.TextBox()
    Me.cmdApply = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
    CType(Me.dgvFeature, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    Me.StatusStrip1.SuspendLayout()
    Me.SuspendLayout()
    '
    'Timer1
    '
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(3, 6)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(63, 13)
    Me.Label1.TabIndex = 1
    Me.Label1.Text = "Features of:"
    '
    'dgvFeature
    '
    Me.dgvFeature.AllowUserToAddRows = False
    Me.dgvFeature.AllowUserToDeleteRows = False
    Me.dgvFeature.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgvFeature.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvFeature.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvFeature.Location = New System.Drawing.Point(6, 25)
    Me.dgvFeature.Name = "dgvFeature"
    Me.dgvFeature.ReadOnly = True
    Me.dgvFeature.Size = New System.Drawing.Size(170, 242)
    Me.dgvFeature.TabIndex = 3
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer1.Name = "SplitContainer1"
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.tbFeatContent)
    Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
    Me.SplitContainer1.Panel1.Controls.Add(Me.dgvFeature)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbFeatValue)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbFeatName)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
    Me.SplitContainer1.Panel2.Controls.Add(Me.cmdDel)
    Me.SplitContainer1.Panel2.Controls.Add(Me.cmdAdd)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbFeatType)
    Me.SplitContainer1.Panel2.Controls.Add(Me.cmdApply)
    Me.SplitContainer1.Panel2.Controls.Add(Me.cmdCancel)
    Me.SplitContainer1.Size = New System.Drawing.Size(351, 270)
    Me.SplitContainer1.SplitterDistance = 179
    Me.SplitContainer1.TabIndex = 4
    '
    'tbFeatContent
    '
    Me.tbFeatContent.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeatContent.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbFeatContent.Location = New System.Drawing.Point(72, 3)
    Me.tbFeatContent.Name = "tbFeatContent"
    Me.tbFeatContent.ReadOnly = True
    Me.tbFeatContent.Size = New System.Drawing.Size(104, 20)
    Me.tbFeatContent.TabIndex = 4
    '
    'tbFeatValue
    '
    Me.tbFeatValue.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeatValue.Location = New System.Drawing.Point(8, 103)
    Me.tbFeatValue.Name = "tbFeatValue"
    Me.tbFeatValue.Size = New System.Drawing.Size(157, 106)
    Me.tbFeatValue.TabIndex = 3
    Me.tbFeatValue.Text = ""
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.Location = New System.Drawing.Point(5, 87)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(37, 13)
    Me.Label6.TabIndex = 22
    Me.Label6.Text = "Value:"
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(5, 48)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(104, 13)
    Me.Label5.TabIndex = 20
    Me.Label5.Text = "Name of the feature:"
    '
    'tbFeatName
    '
    Me.tbFeatName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeatName.Location = New System.Drawing.Point(8, 64)
    Me.tbFeatName.Name = "tbFeatName"
    Me.tbFeatName.Size = New System.Drawing.Size(157, 20)
    Me.tbFeatName.TabIndex = 2
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(5, 9)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(122, 13)
    Me.Label2.TabIndex = 18
    Me.Label2.Text = "Category for this feature:"
    '
    'cmdDel
    '
    Me.cmdDel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdDel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdDel.Location = New System.Drawing.Point(90, 215)
    Me.cmdDel.Name = "cmdDel"
    Me.cmdDel.Size = New System.Drawing.Size(75, 23)
    Me.cmdDel.TabIndex = 17
    Me.cmdDel.Text = "&Delete"
    Me.cmdDel.UseVisualStyleBackColor = True
    '
    'cmdAdd
    '
    Me.cmdAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdAdd.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdAdd.Location = New System.Drawing.Point(9, 215)
    Me.cmdAdd.Name = "cmdAdd"
    Me.cmdAdd.Size = New System.Drawing.Size(75, 23)
    Me.cmdAdd.TabIndex = 16
    Me.cmdAdd.Text = "&Add"
    Me.cmdAdd.UseVisualStyleBackColor = True
    '
    'tbFeatType
    '
    Me.tbFeatType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeatType.Location = New System.Drawing.Point(8, 25)
    Me.tbFeatType.Name = "tbFeatType"
    Me.tbFeatType.Size = New System.Drawing.Size(157, 20)
    Me.tbFeatType.TabIndex = 1
    '
    'cmdApply
    '
    Me.cmdApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdApply.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdApply.Location = New System.Drawing.Point(9, 244)
    Me.cmdApply.Name = "cmdApply"
    Me.cmdApply.Size = New System.Drawing.Size(75, 23)
    Me.cmdApply.TabIndex = 14
    Me.cmdApply.Text = "&Save"
    Me.cmdApply.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(90, 244)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 10
    Me.cmdCancel.Text = "E&xit"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 274)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(351, 22)
    Me.StatusStrip1.TabIndex = 5
    Me.StatusStrip1.Text = "StatusStrip1"
    '
    'ToolStripStatusLabel1
    '
    Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
    Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(19, 17)
    Me.ToolStripStatusLabel1.Text = "..."
    '
    'frmFeature
    '
    Me.AcceptButton = Me.cmdApply
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(351, 296)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.SplitContainer1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.KeyPreview = True
    Me.Name = "frmFeature"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
    Me.Text = "Feature editor"
    CType(Me.dgvFeature, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel1.PerformLayout()
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    Me.SplitContainer1.Panel2.PerformLayout()
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer1.ResumeLayout(False)
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents dgvFeature As System.Windows.Forms.DataGridView
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents cmdApply As System.Windows.Forms.Button
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents tbFeatName As System.Windows.Forms.TextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents cmdDel As System.Windows.Forms.Button
  Friend WithEvents cmdAdd As System.Windows.Forms.Button
  Friend WithEvents tbFeatType As System.Windows.Forms.TextBox
  Friend WithEvents tbFeatValue As System.Windows.Forms.RichTextBox
  Friend WithEvents tbFeatContent As System.Windows.Forms.TextBox
End Class
