<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGetResult
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmGetResult))
    Me.cmdOk = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.tbResTextId = New System.Windows.Forms.TextBox
    Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
    Me.Label1 = New System.Windows.Forms.Label
    Me.Label2 = New System.Windows.Forms.Label
    Me.tbResLoc = New System.Windows.Forms.TextBox
    Me.Label3 = New System.Windows.Forms.Label
    Me.tbResFile = New System.Windows.Forms.TextBox
    Me.tbResEtreeId = New System.Windows.Forms.TextBox
    Me.Label17 = New System.Windows.Forms.Label
    Me.tbResForestId = New System.Windows.Forms.TextBox
    Me.Label15 = New System.Windows.Forms.Label
    Me.tbResPeriod = New System.Windows.Forms.TextBox
    Me.Label16 = New System.Windows.Forms.Label
    Me.Label4 = New System.Windows.Forms.Label
    Me.tbResText = New System.Windows.Forms.RichTextBox
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
    Me.tbResFoll = New System.Windows.Forms.RichTextBox
    Me.tbResPrec = New System.Windows.Forms.RichTextBox
    Me.tbResPsd = New System.Windows.Forms.RichTextBox
    Me.Label5 = New System.Windows.Forms.Label
    Me.Label6 = New System.Windows.Forms.Label
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    Me.SuspendLayout()
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Location = New System.Drawing.Point(569, 365)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 11
    Me.cmdOk.Text = "Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(479, 365)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 12
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'tbResTextId
    '
    Me.tbResTextId.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbResTextId.Location = New System.Drawing.Point(62, 6)
    Me.tbResTextId.Name = "tbResTextId"
    Me.tbResTextId.Size = New System.Drawing.Size(210, 20)
    Me.tbResTextId.TabIndex = 1
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 9)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(40, 13)
    Me.Label1.TabIndex = 6
    Me.Label1.Text = "TextId:"
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(278, 9)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(51, 13)
    Me.Label2.TabIndex = 8
    Me.Label2.Text = "Location:"
    '
    'tbResLoc
    '
    Me.tbResLoc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbResLoc.Location = New System.Drawing.Point(335, 6)
    Me.tbResLoc.Name = "tbResLoc"
    Me.tbResLoc.Size = New System.Drawing.Size(309, 20)
    Me.tbResLoc.TabIndex = 2
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(12, 35)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(26, 13)
    Me.Label3.TabIndex = 10
    Me.Label3.Text = "File:"
    '
    'tbResFile
    '
    Me.tbResFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbResFile.Location = New System.Drawing.Point(62, 32)
    Me.tbResFile.Name = "tbResFile"
    Me.tbResFile.Size = New System.Drawing.Size(582, 20)
    Me.tbResFile.TabIndex = 3
    '
    'tbResEtreeId
    '
    Me.tbResEtreeId.BackColor = System.Drawing.Color.White
    Me.tbResEtreeId.Location = New System.Drawing.Point(281, 58)
    Me.tbResEtreeId.Name = "tbResEtreeId"
    Me.tbResEtreeId.Size = New System.Drawing.Size(65, 20)
    Me.tbResEtreeId.TabIndex = 6
    '
    'Label17
    '
    Me.Label17.AutoSize = True
    Me.Label17.Location = New System.Drawing.Point(232, 61)
    Me.Label17.Name = "Label17"
    Me.Label17.Size = New System.Drawing.Size(47, 13)
    Me.Label17.TabIndex = 33
    Me.Label17.Text = "eTreeId:"
    '
    'tbResForestId
    '
    Me.tbResForestId.BackColor = System.Drawing.Color.White
    Me.tbResForestId.Location = New System.Drawing.Point(161, 58)
    Me.tbResForestId.Name = "tbResForestId"
    Me.tbResForestId.Size = New System.Drawing.Size(65, 20)
    Me.tbResForestId.TabIndex = 5
    '
    'Label15
    '
    Me.Label15.AutoSize = True
    Me.Label15.Location = New System.Drawing.Point(107, 61)
    Me.Label15.Name = "Label15"
    Me.Label15.Size = New System.Drawing.Size(48, 13)
    Me.Label15.TabIndex = 32
    Me.Label15.Text = "ForestId:"
    '
    'tbResPeriod
    '
    Me.tbResPeriod.BackColor = System.Drawing.Color.White
    Me.tbResPeriod.Location = New System.Drawing.Point(62, 58)
    Me.tbResPeriod.Name = "tbResPeriod"
    Me.tbResPeriod.Size = New System.Drawing.Size(39, 20)
    Me.tbResPeriod.TabIndex = 4
    '
    'Label16
    '
    Me.Label16.AutoSize = True
    Me.Label16.Location = New System.Drawing.Point(12, 61)
    Me.Label16.Name = "Label16"
    Me.Label16.Size = New System.Drawing.Size(40, 13)
    Me.Label16.TabIndex = 31
    Me.Label16.Text = "Period:"
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(7, 2)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(31, 13)
    Me.Label4.TabIndex = 37
    Me.Label4.Text = "Text:"
    '
    'tbResText
    '
    Me.tbResText.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbResText.BackColor = System.Drawing.Color.White
    Me.tbResText.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbResText.Location = New System.Drawing.Point(10, 95)
    Me.tbResText.Name = "tbResText"
    Me.tbResText.Size = New System.Drawing.Size(298, 88)
    Me.tbResText.TabIndex = 8
    Me.tbResText.Text = ""
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.SplitContainer1.Location = New System.Drawing.Point(5, 84)
    Me.SplitContainer1.Name = "SplitContainer1"
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.tbResFoll)
    Me.SplitContainer1.Panel1.Controls.Add(Me.tbResPrec)
    Me.SplitContainer1.Panel1.Controls.Add(Me.tbResText)
    Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbResPsd)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
    Me.SplitContainer1.Size = New System.Drawing.Size(652, 275)
    Me.SplitContainer1.SplitterDistance = 311
    Me.SplitContainer1.TabIndex = 39
    '
    'tbResFoll
    '
    Me.tbResFoll.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbResFoll.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbResFoll.Location = New System.Drawing.Point(10, 189)
    Me.tbResFoll.Name = "tbResFoll"
    Me.tbResFoll.Size = New System.Drawing.Size(298, 83)
    Me.tbResFoll.TabIndex = 9
    Me.tbResFoll.Text = ""
    '
    'tbResPrec
    '
    Me.tbResPrec.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbResPrec.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbResPrec.Location = New System.Drawing.Point(10, 18)
    Me.tbResPrec.Name = "tbResPrec"
    Me.tbResPrec.Size = New System.Drawing.Size(299, 71)
    Me.tbResPrec.TabIndex = 7
    Me.tbResPrec.Text = ""
    '
    'tbResPsd
    '
    Me.tbResPsd.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbResPsd.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbResPsd.Location = New System.Drawing.Point(3, 18)
    Me.tbResPsd.Name = "tbResPsd"
    Me.tbResPsd.Size = New System.Drawing.Size(321, 254)
    Me.tbResPsd.TabIndex = 10
    Me.tbResPsd.Text = ""
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(0, 2)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(65, 13)
    Me.Label5.TabIndex = 39
    Me.Label5.Text = "Psd/Syntax:"
    '
    'Label6
    '
    Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label6.AutoSize = True
    Me.Label6.ForeColor = System.Drawing.Color.Maroon
    Me.Label6.Location = New System.Drawing.Point(12, 370)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(281, 13)
    Me.Label6.TabIndex = 40
    Me.Label6.Text = "Note: other information is filled in at the database tab page"
    '
    'frmGetResult
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(656, 400)
    Me.Controls.Add(Me.Label6)
    Me.Controls.Add(Me.SplitContainer1)
    Me.Controls.Add(Me.tbResEtreeId)
    Me.Controls.Add(Me.Label17)
    Me.Controls.Add(Me.tbResForestId)
    Me.Controls.Add(Me.Label15)
    Me.Controls.Add(Me.tbResPeriod)
    Me.Controls.Add(Me.Label16)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.tbResFile)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.tbResLoc)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.tbResTextId)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmGetResult"
    Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Provide basic information about the new corpus result"
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel1.PerformLayout()
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    Me.SplitContainer1.Panel2.PerformLayout()
    Me.SplitContainer1.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents tbResTextId As System.Windows.Forms.TextBox
  Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents tbResLoc As System.Windows.Forms.TextBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents tbResFile As System.Windows.Forms.TextBox
  Friend WithEvents tbResEtreeId As System.Windows.Forms.TextBox
  Friend WithEvents Label17 As System.Windows.Forms.Label
  Friend WithEvents tbResForestId As System.Windows.Forms.TextBox
  Friend WithEvents Label15 As System.Windows.Forms.Label
  Friend WithEvents tbResPeriod As System.Windows.Forms.TextBox
  Friend WithEvents Label16 As System.Windows.Forms.Label
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents tbResText As System.Windows.Forms.RichTextBox
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents tbResPsd As System.Windows.Forms.RichTextBox
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents tbResFoll As System.Windows.Forms.RichTextBox
  Friend WithEvents tbResPrec As System.Windows.Forms.RichTextBox
End Class
