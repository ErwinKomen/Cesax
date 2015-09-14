Imports System.Xml
Public Class frmDbFilter
  ' ================================= LOCAL VARIABLES =========================================================
  Private loc_strFilter As String = _
    "./descendant-or-self::Result[" & vbCrLf & _
    "  child::Feature[@Name='(FeatureName)' and tb:matches(@Value, '(FeatureValue)')]" & vbCrLf & _
    "]" & vbCrLf
  Private colFeature As New StringColl  ' Collection of feature names and values
  Private colResultA As New StringColl  ' Collection of result attribute/node names and values
  Private arResultAttr() As String = {"Search", "TextId", "Cat", "Period", "Notes", "Status"}
  Private arResultNode() As String = {"Text", "Psd"}
  ' ============================================================================================================
  '---------------------------------------------------------------------------------------------------------
  ' Name:       XpathFilter()
  ' Goal:       Return the value of [loc_strFilter]
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public ReadOnly Property XpathFilter() As String
    Get
      Return loc_strFilter
    End Get
  End Property
  '---------------------------------------------------------------------------------------------------------
  ' Name:       frmDbFilter_Load()
  ' Goal:       Start-up material
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub frmDbFilter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set me somewhere
    Me.Owner = frmMain
    ' Start the timer for initialisation
    Me.Timer1.Enabled = True
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       Timer1_Tick()
  ' Goal:       initialise
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Dim ndxList As XmlNodeList  ' List of nodex
    ' Dim ndxThis As XmlNode      ' Working node
    Dim crsThis As Windows.Forms.Cursor = Cursor.Current
    Dim intI As Integer         ' Counter

    Try
      ' only continue if visible
      If (Not Me.Visible) Then Exit Sub
      ' Switch off the timer
      Me.Timer1.Enabled = False
      ' Show we are busy
      Status("Initialising...") : Cursor.Current = Cursors.WaitCursor
      ' Perform initialisations...
      ' (1) Determine the names of the features from looking at the first record containing them
      ndxList = pdxCrpDbase.SelectNodes("./descendant::Result[1]/child::Feature")
      If (ndxList.Count > 0) Then
        ' Reset the combobox
        With Me.cboFeatureName
          ' Clear it
          .Items.Clear()
          ' Get a list of features and load the combobox
          For intI = 0 To ndxList.Count - 1
            .Items.Add(ndxList(intI).Attributes("Name").Value)
          Next intI
          ' Set the first one
          .SelectedIndex = 0
        End With
      End If
      ' (2) Add the names for Result characteristics
      With Me.cboResultName
        ' Clear
        .Items.Clear()
        ' Load the attributes
        For intI = 0 To arResultAttr.Length - 1
          .Items.Add(arResultAttr(intI))
        Next intI
        ' Load the nodes
        For intI = 0 To arResultNode.Length - 1
          .Items.Add(arResultNode(intI))
        Next intI
        ' Set the first one
        .SelectedIndex = 0
      End With
      ' (3) Load a default filter
      AdaptXpath()
      ' Me.tbFilterXpath.Text = loc_strFilter
      ' Show we are ready
      Status("Ready")
      Cursor.Current = crsThis
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbFilter/Timer1 error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       cmdCancel_Click()
  ' Goal:       Close and provide the value "Cancel"
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       cmdOk_Click()
  ' Goal:       Close and provide the value "Ok"
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    Me.DialogResult = Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       tbFilterXpath_TextChanged()
  ' Goal:       Read the latest version of the filter
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub tbFilterXpath_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFilterXpath.TextChanged
    ' Adapt the filter
    loc_strFilter = Me.tbFilterXpath.Text
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       cmdAdd_Click()
  ' Goal:       Add one feature name/value combination
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddF.Click
    Try
      ' Validate
      If (Me.cboFeatureName.SelectedIndex < 0) Then Exit Sub
      ' If the feature exists: remove it
      If (colFeature.Exists(Me.cboFeatureName.SelectedItem)) Then colFeature.DelItem(colFeature.Find(Me.cboFeatureName.SelectedItem))
      ' Add a feature with the indicated name and value
      colFeature.Add(Me.cboFeatureName.SelectedItem, Me.tbFeatureValue.Text)
      ' Adapt the Xpath filter expression
      AdaptXpath()
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbFilter/cmdAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub cmdAddR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddR.Click
    Try
      ' Validate
      If (Me.cboResultName.SelectedIndex < 0) Then Exit Sub
      ' If the feature exists: remove it
      If (colResultA.Exists(Me.cboResultName.SelectedItem)) Then colResultA.DelItem(colResultA.Find(Me.cboResultName.SelectedItem))
      ' Add a feature with the indicated name and value
      colResultA.Add(Me.cboResultName.SelectedItem, Me.tbResultValue.Text)
      ' Adapt the Xpath filter expression
      AdaptXpath()
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbFilter/cmdAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       cmdDel_Click()
  ' Goal:       Delete one feature name/value combination
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub cmdDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelF.Click
    Try
      ' Validate
      If (Me.cboFeatureName.SelectedIndex < 0) Then Exit Sub
      ' If the feature exists: remove it
      If (colFeature.Exists(Me.cboFeatureName.SelectedItem)) Then colFeature.DelItem(colFeature.Find(Me.cboFeatureName.SelectedItem))
      ' Adapt the Xpath filter expression
      AdaptXpath()
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbFilter/AdaptXpath error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub cmdDelR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelR.Click
    Try
      ' Validate
      If (Me.cboResultName.SelectedIndex < 0) Then Exit Sub
      ' If the feature exists: remove it
      If (colResultA.Exists(Me.cboResultName.SelectedItem)) Then colResultA.DelItem(colResultA.Find(Me.cboResultName.SelectedItem))
      ' Adapt the Xpath filter expression
      AdaptXpath()
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbFilter/AdaptXpath error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       cmdReset_Click()
  ' Goal:       Remove all feature/value combinations
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdResetF.Click
    colFeature.Clear() : colResultA.Clear()
    AdaptXpath()
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       AdaptXpath()
  ' Goal:       Build an Xpath expression and show it
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub AdaptXpath()
    Dim bFirst As Boolean = True  ' Flag for the first feature/node
    Dim strXpath As String  ' Query
    Dim intI As Integer     ' Counter

    Try
      ' Are there any results?
      If (colFeature.Count = 0) AndAlso (colResultA.Count = 0) Then
        ' Make simple query
        Me.tbFilterXpath.Text = "./descendant-or-self::Result" & vbCrLf
      Else
        ' Start
        strXpath = "./descendant-or-self::Result[" & vbCrLf
        ' Walk all possible attributes/nodes
        For intI = 0 To colResultA.Count - 1
          ' Add preamble
          If (bFirst) Then
            strXpath &= "  "
            bFirst = False
          Else
            strXpath &= "  and "
          End If
          ' Is this a result attribute or node?
          If (Array.Exists(arResultAttr, Function(strIn As String) strIn = colResultA.Item(intI))) Then
            ' This is an attribute
            strXpath &= "tb:matches(string(@" & colResultA.Item(intI) & "), '" & colResultA.Exmp(intI) & "')" & vbCrLf
          Else
            ' This is a child node
            strXpath &= "tb:matches(string(child::" & colResultA.Item(intI) & "), '" & colResultA.Exmp(intI) & "')" & vbCrLf
          End If
        Next intI
        ' Walk all features
        For intI = 0 To colFeature.Count - 1
          ' Add preamble
          If (bFirst) Then
            strXpath &= "  "
            bFirst = False
          Else
            strXpath &= "  and "
          End If
          ' Add this one
          strXpath &= "child::Feature[@Name = '" & colFeature.Item(intI) & "' and tb:matches(@Value, '" & colFeature.Exmp(intI) & "')]" & vbCrLf
        Next intI
        ' Finish
        strXpath &= "]" & vbCrLf
        Me.tbFilterXpath.Text = strXpath
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbFilter/AdaptXpath error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class