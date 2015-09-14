Imports System.Windows.Forms
Imports System.Xml
Public Class frmCatType
  ' ================================= LOCAL VARIABLES =========================================================
  Private loc_strCatType As String = ""     ' Local copy of CatType
  Private loc_strCategory As String = ""    ' Local copy of Category (Adv, Vb etc)
  Private loc_intLevelSyntax As Integer           ' Level of syntactical embedding
  Private loc_intId As Integer          ' the Id of the node in which the current selection is
  Private fntNormal As System.Drawing.Font = New Font("Courier New", 10, FontStyle.Regular)
  Private fntBold As System.Drawing.Font = New Font("Courier New", 10, FontStyle.Bold)
  Private fntSub As System.Drawing.Font = New Font("Courier New", 7, FontStyle.Regular)
  ' ================================= LOCAL CONSTANTS =========================================================
  ' ===========================================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   CatValue
  ' Goal:   Get or set the CatValue property
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Property CatValue() As String
    Get
      ' Return the local copy
      CatValue = loc_strCatType
    End Get
    Set(ByVal value As String)
      Dim arCatValue() As String  ' Array of ambiguous values
      Dim intI As Integer         ' COunter

      ' Split the value into an array
      arCatValue = Split(value, ";")
      ' Clear the listbox
      With Me.lboCatType
        .Items.Clear()
        ' Load the values
        For intI = 0 To UBound(arCatValue)
          .Items.Add(arCatValue(intI))
        Next intI
        ' Sort the items alphabetically
        .Sorted = True
        ' Set the selected value
        .SelectedIndex = 0
      End With
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Clause
  ' Goal:   Set the HTML text of the [wbClause]
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property Clause() As String
    Set(ByVal value As String)
      Me.wbClause.DocumentText = value
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Syntax
  ' Goal:   Set the syntax (PSD) text
  ' History:
  ' 17-02-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub Syntax(ByRef ndxThis As XmlNode)
    Dim ndxForest As XmlNode  ' <forest> ancestor of me

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Exit Sub
      If (ndxThis.Name <> "eTree") Then Exit Sub
      ' Extract the ID number
      loc_intId = ndxThis.Attributes("Id").Value
      ' Get the parent, the forest Id
      ndxForest = ndxThis.SelectSingleNode("./ancestor::forest")
      ' Clear what was there
      Me.tbPsd.Text = ""
      If (Not ndxForest Is Nothing) Then
        ' Perform the syntax operation
        If (Not GetSyntax(ndxForest, Me.tbPsd)) Then
          ' Something went wrong
          Status("Error deriving syntax")
        End If
      End If
    Catch ex As Exception
      ' Show error
      MsgBox("frmCatType/Syntax error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   Alike
  ' Goal:   Fill the table with the values that are like me
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property Alike() As String
    Set(ByVal value As String)
      Me.tbValues.Text = value
      Me.tbValues.SelectionStart = 1
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Category
  ' Goal:   Get or set the category (Adv, Vb etc)
  ' History:
  ' 10-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Property Category() As String
    Get
      ' Return the local topy
      Category = loc_strCategory
    End Get
    Set(ByVal value As String)
      ' Get the category
      loc_strCategory = value
      ' Show the category appropriately
      Me.tbCategory.Text = loc_strCategory
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   OK_Button_Click
  ' Goal:   Return a positive dialogue result
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Cancel_Button_Click
  ' Goal:   Return negative dialogue result
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmCatType_Load
  ' Goal:   Trigger initialisation
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmCatType_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set owner
    Me.Owner = frmMain
    Me.StartPosition = FormStartPosition.CenterScreen
    ' Initially set the category...
    Me.tbCategory.Text = "(Unknown)"
    ' Make sure the initialisation gets triggered
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Perform initialisation
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off the timer
      Me.Timer1.Enabled = False
      ' Show the first element of the listbox
      Me.lboCatType.SelectedIndex = 0
      ' Show we are ready
      Status("ready")
      lboCatType_SelectedIndexChanged(sender, e)
    Catch ex As Exception
      ' Show error
      MsgBox("frmCatType/Timer1 (initialisation) error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   lboCatType_DoubleClick
  ' Goal:   Make sure we end with a positive note
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboCatType_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lboCatType.DoubleClick
    ' Show the correct value
    Me.tbCatType.Text = Me.lboCatType.SelectedItem
    ' Set the dialog result
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    ' Leave this form
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   lboCatType_SelectedIndexChanged
  ' Goal:   Adapt what is displayed
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboCatType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lboCatType.SelectedIndexChanged
    Try
      ' Validate
      If (Not Me.Visible) Then Exit Sub
      ' Show the correct value
      Me.tbCatType.Text = Me.lboCatType.SelectedItem
    Catch ex As Exception
      ' Show error
      MsgBox("frmCatType/CatType_SelectedIndexChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   tbCatType_TextChanged
  ' Goal:   Adapt the local copy
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub tbCatType_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCatType.TextChanged
    ' Set the local copy
    loc_strCatType = Me.tbCatType.Text
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmCatType_VisibleChanged
  ' Goal:   When it becomes visible, I want the correct tabpage to show up
  ' History:
  ' 17-02-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmCatType_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
    ' Check if it becomes visible
    If (Me.Visible) Then
      ' Show the correct tabpage
      Me.TabControl1.SelectedTab = Me.tpClause
    End If
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   GetSyntax
  ' Goal:   Get the syntax of one node
  ' History:
  ' 28-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetSyntax(ByRef ndxThis As XmlNode, ByRef rtfSynt As RichTextBox) As Boolean
    Dim ndxChild As XmlNode   ' Start node
    Dim ndxParent As XmlNode  ' Parent node
    Dim strLabel As String    ' The label of an <eTree>
    Dim strParentL As String  ' Label of parent

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (rtfSynt Is Nothing) Then Return False
      ' Determine what kind of node this is
      Select Case ndxThis.Name
        Case "eLeaf"
          ' Is this the selected Id?
          If (ndxThis.SelectSingleNode("./ancestor::eTree[@Id=" & loc_intId & "]") Is Nothing) Then
            ' Just add the text of this node
            AppendToRtb(rtfSynt, ndxThis.Attributes("Text").Value, Color.Black, fntNormal)
          Else
            ' The text of the selected Id is set in red.
            AppendToRtb(rtfSynt, ndxThis.Attributes("Text").Value, Color.Red, fntBold)
          End If
        Case "forest"
          ' Walk all the children, which are 1 level deeper
          loc_intLevelSyntax += 1
          For Each ndxChild In ndxThis.ChildNodes
            ' Apply action on this child
            If (Not GetSyntax(ndxChild, rtfSynt)) Then
              ' Error, so retreat
              Return False
            End If
          Next ndxChild
          ' Go one level down again
          loc_intLevelSyntax -= 1
        Case "eTree"
          ' Action depends on what kind of label we have
          strLabel = ndxThis.Attributes("Label").Value
          ' Make sure CODE nodes are not shown, unless desired by the user...
          If (strLabel <> "CODE") OrElse (bShowCode) Then
            ' Check if we are IP or CP, or if our parents are
            If (strLabel Like "[IC]P-*") Then
              ' Need an LF
              rtfSynt.AppendText(vbCrLf & Strings.StrDup(loc_intLevelSyntax, " "))
            Else
              ' Try get parent
              ndxParent = ndxThis.ParentNode
              If (Not ndxParent Is Nothing) Then
                If (Not ndxParent.Attributes("Label") Is Nothing) Then
                  ' Get parent label name
                  strParentL = ndxParent.Attributes("Label").Value
                  ' Check parent label
                  If (strParentL Like "[IC]P-*") Then
                    ' Need an LF
                    rtfSynt.AppendText(vbCrLf & Strings.StrDup(loc_intLevelSyntax, " "))
                  End If
                End If
              End If
            End If
            ' First give the bracket
            AppendToRtb(rtfSynt, "(", Color.Black, fntNormal)
            ' Give the label of this node
            AppendToRtb(rtfSynt, strLabel & " ", Color.Blue, fntSub)
            ' Continue with the children, which are 1 level more down
            loc_intLevelSyntax += 1
            For Each ndxChild In ndxThis.ChildNodes
              ' Apply action on this child
              If (Not GetSyntax(ndxChild, rtfSynt)) Then
                ' Error, so retreat
                Return False
              End If
            Next ndxChild
            ' Go one level down again
            loc_intLevelSyntax -= 1
            ' Append a closing bracket
            AppendToRtb(rtfSynt, ")", Color.Black, fntNormal)
            ' AFter a CODE node, we need a linefeed
            If (strLabel = "CODE") Then
              AppendToRtb(rtfSynt, vbCrLf, Color.Black, fntNormal)
            End If
          End If
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      MsgBox("frmCatType/GetSyntax error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
End Class
