<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCatType
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCatType))
    Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
    Me.OK_Button = New System.Windows.Forms.Button
    Me.Cancel_Button = New System.Windows.Forms.Button
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.Label1 = New System.Windows.Forms.Label
    Me.tbCatType = New System.Windows.Forms.TextBox
    Me.Label2 = New System.Windows.Forms.Label
    Me.lboCatType = New System.Windows.Forms.ListBox
    Me.Label4 = New System.Windows.Forms.Label
    Me.wbClause = New System.Windows.Forms.WebBrowser
    Me.Label5 = New System.Windows.Forms.Label
    Me.tbCategory = New System.Windows.Forms.TextBox
    Me.tbValues = New System.Windows.Forms.RichTextBox
    Me.Label6 = New System.Windows.Forms.Label
    Me.Label7 = New System.Windows.Forms.Label
    Me.TabControl1 = New System.Windows.Forms.TabControl
    Me.tpClause = New System.Windows.Forms.TabPage
    Me.tpPsd = New System.Windows.Forms.TabPage
    Me.tbPsd = New System.Windows.Forms.RichTextBox
    Me.TableLayoutPanel1.SuspendLayout()
    Me.TabControl1.SuspendLayout()
    Me.tpClause.SuspendLayout()
    Me.tpPsd.SuspendLayout()
    Me.SuspendLayout()
    '
    'TableLayoutPanel1
    '
    Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TableLayoutPanel1.ColumnCount = 2
    Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
    Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
    Me.TableLayoutPanel1.Location = New System.Drawing.Point(426, 458)
    Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
    Me.TableLayoutPanel1.RowCount = 1
    Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
    Me.TableLayoutPanel1.TabIndex = 0
    '
    'OK_Button
    '
    Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
    Me.OK_Button.Location = New System.Drawing.Point(3, 3)
    Me.OK_Button.Name = "OK_Button"
    Me.OK_Button.Size = New System.Drawing.Size(67, 23)
    Me.OK_Button.TabIndex = 0
    Me.OK_Button.Text = "OK"
    '
    'Cancel_Button
    '
    Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
    Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
    Me.Cancel_Button.Name = "Cancel_Button"
    Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
    Me.Cancel_Button.TabIndex = 1
    Me.Cancel_Button.Text = "Cancel"
    '
    'Timer1
    '
    '
    'Label1
    '
    Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(127, 311)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(138, 13)
    Me.Label1.TabIndex = 1
    Me.Label1.Text = "Correct (one word) category"
    '
    'tbCatType
    '
    Me.tbCatType.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.tbCatType.Location = New System.Drawing.Point(130, 345)
    Me.tbCatType.Name = "tbCatType"
    Me.tbCatType.Size = New System.Drawing.Size(90, 20)
    Me.tbCatType.TabIndex = 6
    '
    'Label2
    '
    Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 311)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(83, 13)
    Me.Label2.TabIndex = 7
    Me.Label2.Text = "Possible values:"
    '
    'lboCatType
    '
    Me.lboCatType.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.lboCatType.FormattingEnabled = True
    Me.lboCatType.Location = New System.Drawing.Point(15, 327)
    Me.lboCatType.Name = "lboCatType"
    Me.lboCatType.Size = New System.Drawing.Size(95, 160)
    Me.lboCatType.TabIndex = 8
    '
    'Label4
    '
    Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label4.AutoSize = True
    Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
    Me.Label4.Location = New System.Drawing.Point(127, 368)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(134, 13)
    Me.Label4.TabIndex = 9
    Me.Label4.Text = "E.g: Time, Dir, Contrast etc"
    '
    'wbClause
    '
    Me.wbClause.Dock = System.Windows.Forms.DockStyle.Fill
    Me.wbClause.Location = New System.Drawing.Point(3, 3)
    Me.wbClause.MinimumSize = New System.Drawing.Size(20, 20)
    Me.wbClause.Name = "wbClause"
    Me.wbClause.Size = New System.Drawing.Size(540, 261)
    Me.wbClause.TabIndex = 10
    '
    'Label5
    '
    Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(127, 397)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(55, 13)
    Me.Label5.TabIndex = 11
    Me.Label5.Text = "Category: "
    '
    'tbCategory
    '
    Me.tbCategory.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.tbCategory.Location = New System.Drawing.Point(130, 413)
    Me.tbCategory.Name = "tbCategory"
    Me.tbCategory.ReadOnly = True
    Me.tbCategory.Size = New System.Drawing.Size(100, 20)
    Me.tbCategory.TabIndex = 12
    '
    'tbValues
    '
    Me.tbValues.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbValues.Location = New System.Drawing.Point(287, 327)
    Me.tbValues.Name = "tbValues"
    Me.tbValues.ReadOnly = True
    Me.tbValues.Size = New System.Drawing.Size(282, 125)
    Me.tbValues.TabIndex = 13
    Me.tbValues.Text = ""
    '
    'Label6
    '
    Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label6.AutoSize = True
    Me.Label6.Location = New System.Drawing.Point(127, 325)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(150, 13)
    Me.Label6.TabIndex = 14
    Me.Label6.Text = "descriptive for this constituent:"
    '
    'Label7
    '
    Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label7.AutoSize = True
    Me.Label7.Location = New System.Drawing.Point(284, 311)
    Me.Label7.Name = "Label7"
    Me.Label7.Size = New System.Drawing.Size(81, 13)
    Me.Label7.TabIndex = 15
    Me.Label7.Text = "Related values:"
    '
    'TabControl1
    '
    Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TabControl1.Controls.Add(Me.tpClause)
    Me.TabControl1.Controls.Add(Me.tpPsd)
    Me.TabControl1.Location = New System.Drawing.Point(15, 12)
    Me.TabControl1.Name = "TabControl1"
    Me.TabControl1.SelectedIndex = 0
    Me.TabControl1.Size = New System.Drawing.Size(554, 293)
    Me.TabControl1.TabIndex = 16
    '
    'tpClause
    '
    Me.tpClause.Controls.Add(Me.wbClause)
    Me.tpClause.Location = New System.Drawing.Point(4, 22)
    Me.tpClause.Name = "tpClause"
    Me.tpClause.Padding = New System.Windows.Forms.Padding(3)
    Me.tpClause.Size = New System.Drawing.Size(546, 267)
    Me.tpClause.TabIndex = 0
    Me.tpClause.Text = "Clause"
    Me.tpClause.UseVisualStyleBackColor = True
    '
    'tpPsd
    '
    Me.tpPsd.Controls.Add(Me.tbPsd)
    Me.tpPsd.Location = New System.Drawing.Point(4, 22)
    Me.tpPsd.Name = "tpPsd"
    Me.tpPsd.Padding = New System.Windows.Forms.Padding(3)
    Me.tpPsd.Size = New System.Drawing.Size(546, 267)
    Me.tpPsd.TabIndex = 1
    Me.tpPsd.Text = "Syntax"
    Me.tpPsd.UseVisualStyleBackColor = True
    '
    'tbPsd
    '
    Me.tbPsd.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tbPsd.Location = New System.Drawing.Point(3, 3)
    Me.tbPsd.Name = "tbPsd"
    Me.tbPsd.Size = New System.Drawing.Size(540, 261)
    Me.tbPsd.TabIndex = 0
    Me.tbPsd.Text = ""
    '
    'frmCatType
    '
    Me.AcceptButton = Me.OK_Button
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.Cancel_Button
    Me.ClientSize = New System.Drawing.Size(584, 499)
    Me.Controls.Add(Me.TabControl1)
    Me.Controls.Add(Me.Label7)
    Me.Controls.Add(Me.Label6)
    Me.Controls.Add(Me.tbValues)
    Me.Controls.Add(Me.tbCategory)
    Me.Controls.Add(Me.Label5)
    Me.Controls.Add(Me.Label4)
    Me.Controls.Add(Me.lboCatType)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.tbCatType)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.TableLayoutPanel1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmCatType"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Please supply the correct category description for the indicated constituent"
    Me.TableLayoutPanel1.ResumeLayout(False)
    Me.TabControl1.ResumeLayout(False)
    Me.tpClause.ResumeLayout(False)
    Me.tpPsd.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
  Friend WithEvents OK_Button As System.Windows.Forms.Button
  Friend WithEvents Cancel_Button As System.Windows.Forms.Button
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents tbCatType As System.Windows.Forms.TextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents lboCatType As System.Windows.Forms.ListBox
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents wbClause As System.Windows.Forms.WebBrowser
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents tbCategory As System.Windows.Forms.TextBox
  Friend WithEvents tbValues As System.Windows.Forms.RichTextBox
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents Label7 As System.Windows.Forms.Label
  Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
  Friend WithEvents tpClause As System.Windows.Forms.TabPage
  Friend WithEvents tpPsd As System.Windows.Forms.TabPage
  Friend WithEvents tbPsd As System.Windows.Forms.RichTextBox

End Class
