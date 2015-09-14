Imports System.Xml
Imports Netron
Module modEditor
  ' ================================= LOCAL STRUCTURE =========================================================
  Private Structure NPinfo
    Dim NodeId As Integer     ' The NodeID of this NP
    Dim NPtext As String      ' The text of this NP
    Dim NPtype As String      ' THe kind of NP
    Dim NPcat As String       ' The syntactical category (label) of the NP
    Dim Size As Integer       ' Size of the NP
    Dim NPcase As String      ' Grammatical case of the NP
    Dim CorType As String     ' Coreference type of the NP
    Dim CorDist As Integer    ' Coreference distance of the NP
    Dim Hierarchy As Integer  ' The NP hierarchy number (see GetNPhierarchy function)
    Dim IsPP As Boolean       ' Whether this NP is part of a PP
  End Structure
  ' ================================= LOCAL VARIABLES =========================================================
  Private bEdtSpace As Boolean = False  ' Whether we are working under the pressure of Space-Pressing
  Private bBusy As Boolean = False      ' Flag to help work while not being doubly interrupted
  Private bTreeBusy As Boolean = False  ' Treestack is internally being worked on
  Private loc_intPos As Integer         ' Start of selected position
  Private loc_intLen As Integer         ' Length of current selection
  Private loc_intLineNum As Integer     ' The line number where we are
  Private loc_intLinePrev As Integer = -1 ' Line number previous time of selection
  Private loc_intCharNum As Integer     ' The character position where we are
  Private loc_intListNum As Integer     ' The currently selected number from the list
  Private loc_intLeft As Integer        ' Position where the current line starts left
  Private loc_intId As Integer          ' the Id of the node in which the current selection is
  Private loc_intCorNum As Integer = 0  ' Number of elements in loc_arCoref
  Private loc_xndList As XmlNodeList    ' A list of currently possible selectable constituents
  Private loc_arCoref As New NodeColl   ' Collection of <eTree> nodes: first source, then goal(s) for coref
  Private loc_ndxLast As XmlNode = Nothing        ' Node that was selected lastly
  Private loc_arPoints(0 To 1) As Drawing.PointF  ' Array of points for the drawing
  Private loc_grThis As Graphics                  ' Local graphics stuff
  Private loc_arLine As New NodeColl              ' A chain down with coreferential nodes
  Private loc_arSel As New NodeColl               ' Currently selected and valid nodes
  Private loc_intLevelSyntax As Integer           ' Level of syntactical embedding
  Private loc_intAutoPos As Integer = 1           ' Postion for autostuff
  Private loc_rtfEditor As RichTextBox = Nothing  ' The Editor's richtextbox
  Private strNoText As String = "CODE|META|METADATA|E_S"
  Private ndxCurrentSel As XmlNode = Nothing      ' Currently selected node
  Private fntNormal As System.Drawing.Font = New Font("Courier New", 10, FontStyle.Regular)
  Private fntBold As System.Drawing.Font = New Font("Courier New", 10, FontStyle.Bold)
  Private fntSub As System.Drawing.Font = New Font("Courier New", 7, FontStyle.Regular)
  Private fntPde As System.Drawing.Font = New Font("Courier New", 9, FontStyle.Regular)
  Private fntPdeEmph As System.Drawing.Font = New Font("Courier New", 9, FontStyle.Bold)
  Private loc_colTreeEdit As New StringColl       ' Collection of edits for "undo" operation
  ' ================================= GLOBAL VARIABLES ========================================================
  Public colCorefSource As Color = Color.Yellow   ' Coreference source (background color)
  Public colCorefDest As Color = Color.LightGreen ' Coreference destination (background color)
  Public colSelect As Color = Color.LightGray     ' Simple selection color
  Public strRefNew As String = "New"              ' Completely new NP
  Public strRefNewVar As String = "NewVar"        ' Introductino of a variable
  Public strRefAssumed As String = "Assumed"      ' Assumed/world knowledge
  Public strRefIdentity As String = "Identity"    ' Identity coreference relation
  Public strRefInferred As String = "Inferred"    ' Inferred (part/whole etc) relation
  Public strRefInert As String = "Inert"          ' Does not refer to something and cannot be referred to
  Public strRefCross As String = "CrossSpeech"    ' Cross a speech boundary
  Public strRefOneArg As String = strRefNew & "|" & strRefNewVar & "|" & strRefAssumed & "|" & strRefInert
  Public strRefTwoArg As String = strRefIdentity & "|" & strRefInferred & "|" & strRefCross
  Public bAnyConstituent As Boolean = False      ' Select any constituent or not?
  ' ===========================================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   InitAutoCoref
  ' Goal:   Set a local copy for the richtextbox of the Coreference Editor
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub InitAutoCoref(ByRef rtfThis As RichTextBox)
    ' Set the editor
    loc_rtfEditor = rtfThis
    ' Make sure we are not interrupted
    bEdtSpace = True
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   FinishAutoCoref
  ' Goal:   Make sure the [SelectionChanged] can do its proper work again
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub FinishAutoCoref()
    ' Make sure we are interruptable again
    bEdtSpace = False
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   SelectNode
  ' Goal:   Select the indicated node as source or destination
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SelectNode(ByRef ndxThis As XmlNode, ByVal strType As String)
    Dim colThis As Color
    Dim intLeft As Integer  ' Position on the left of the line we are using

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Exit Sub
      ' Determine which color to take
      Select Case strType
        Case "source", "Source"
          colThis = colCorefSource
          ' Set this as selected node
          ndxCurrentSel = ndxThis
        Case "target", "Target"
          colThis = colCorefDest
        Case "blank", "Blank"
          colThis = Color.White
      End Select
      ' Get left position
      intLeft = GetLeft(ndxThis)
      ' Validate
      If (intLeft < 0) Then Exit Sub
      ' Store this position for the translation...
      loc_intAutoPos = intLeft + ndxThis.Attributes("from").Value
      ' Select this item
      SelectEtree(loc_rtfEditor, ndxThis, intLeft, colThis)
      ' Check whether we are still visible
      If (loc_rtfEditor.GetPositionFromCharIndex(intLeft).Y + 50 > loc_rtfEditor.Height) OrElse _
         (loc_rtfEditor.GetPositionFromCharIndex(intLeft).Y < 0) Then
        ' Set the start of the selection
        bEdtSpace = True
        loc_rtfEditor.SelectionStart = intLeft
        bEdtSpace = False
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/SelectNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetMySection
  ' Goal:   Get the number of the section I belong to
  ' History:
  ' 01-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function GetMySection(ByRef ndxThis As XmlNode) As Integer
    Dim ndxForest As XmlNode  ' Forest of current node
    Dim intI As Integer       ' Counter
    Dim intForestId As Integer ' My forest id
    Dim intSect As Integer    ' Section number

    Try
      ' Get this node's forest
      ndxForest = Nothing
      If (Not GetForestNode(ndxThis, ndxForest)) Then Exit Function
      ' Get my forest ide
      intForestId = ndxForest.Attributes("forestId").Value
      ' Initially set my section to zero
      intSect = 0
      ' Check if there are more than 1 sections
      If (pdxSection.Count > 1) Then
        ' Check if this is section 0
        If (intForestId >= pdxSection(1).Attributes("forestId").Value) Then
          ' Walk all sections
          For intI = pdxSection.Count - 1 To 1 Step -1
            ' Check if this is the correct section
            If (intForestId >= pdxSection(intI).Attributes("forestId").Value)  Then
              ' this is the section
              intSect = intI
              Exit For
            End If
          Next intI
        End If
      End If
      ' Return the section number
      Return intSect
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetMySection error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return at least something appropriate
      Return 0
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GotoNode
  ' Goal:   Set the cursor in the indicated node
  ' History:
  ' 15-06-2010  ERK Created
  ' 08-12-2012  ERK Need to make sure that the actual node, and not only the left hand is selected
  ' ------------------------------------------------------------------------------------
  Public Sub GotoNode(ByRef ndxThis As XmlNode)
    Dim intLeft As Integer    ' Position on the left of the line we are using
    Dim intSect As Integer    ' My section number
    Dim ndxWork As XmlNode    ' Working node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Exit Sub
      ' Get the head-noun
      ndxWork = GetHeadNounNode(ndxThis)
      ' Get my section number
      intSect = GetMySection(ndxWork)
      ' Check if this is the current section
      If (intSect <> intCurrentSection) Then
        ' Go to this section
        frmMain.GoShowSection(intSect)
      End If
      ' Get left position
      intLeft = GetLeft(ndxWork)
      ' Validate
      If (intLeft < 0) Then Exit Sub
      ' Store this position for the translation...
      loc_intAutoPos = intLeft + ndxWork.Attributes("from").Value
      ' Put the cursor there
      frmMain.tbEdtMain.SelectionStart = loc_intAutoPos - 1
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GotoNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   InitGraphics
  ' Goal:   Initialise the possibility to work with graphics (lines) within the RTB
  ' History:
  ' 18-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub InitGraphics(ByRef rtfThis As RichTextBox)
    Try
      ' Validate
      If (rtfThis Is Nothing) Then Exit Sub
      ' Set the graphics object
      loc_grThis = rtfThis.CreateGraphics
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/InitGraphics error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   GoToCurrentLine
  ' Goal:   Set the cursor to the current line
  ' History:
  ' 10-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub GoToCurrentLine(ByRef rtfThis As RichTextBox)
    Dim intCount As Integer   ' Number of possibilities
    Dim intI As Integer       ' Counter
    Dim intPos As Integer     ' Position within string

    Try
      ' Determine how many LF's we have to count
      intCount = loc_intLineNum : intPos = 0
      ' Find this in [strCurrentText]
      For intI = 1 To intCount
        ' Find the next position 
        intPos = InStr(intPos + 1, strCurrentText, vbLf)
      Next intI
      ' Debug.Print(frmMain.TabControl1.SelectedTab.Name)
      ' Set this position in the rich textbox
      rtfThis.SelectionStart = intPos
      rtfThis.SelectionLength = 0
      ' Check which tab page is selected
      Select Case frmMain.TabControl1.SelectedTab.Name
        Case frmMain.tpSyntax.Name
          ' Make sure the correct syntax is being shown
          ShowSyntax(frmMain.tbEdtSyntax)
        Case frmMain.tpTree.Name
          ' Make sure the correct tree is being shown
          ShowTree(frmMain.litTree)
        Case frmMain.tpDep.Name
          ' Make sure the correct dependency is being shown
          ShowCurrentDep()
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GoToCurrentLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   GetCurrentText
  ' Goal:   Get the text of the currently selected line
  ' History:
  ' 19-03-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetCurrentText(ByRef intForestId As Integer, Optional ByVal bShowLineNum As Boolean = False) As String
    Dim ndxFor As XmlNode = Nothing ' Current forest
    Dim ndxCns As XmlNode = Nothing ' Current etree
    Dim strBack As String = ""      ' What we return

    Try
      ' Get current forest
      If (Not GetCurrentForest(ndxFor, ndxCns)) Then Return "(no forest)"
      intForestId = ndxFor.Attributes("forestId").Value
      ' Get the text of this line
      If (bShowLineNum) Then
        strBack = "[" & intForestId & "] " & GetSeg(ndxFor, "org")
      Else
        strBack = GetSeg(ndxFor, "org")
      End If
      ' Return the result
      Return strBack
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetCurrentText error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return "(error)"
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetCurrentLine
  ' Goal:   Get the currently selected line number
  ' History:
  ' 19-03-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetCurrentLine() As Integer
    Try
      ' Return the currently selected line
      Return loc_intLineNum
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetCurrentLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   GetCurrentForest
  ' Goal:   Get the currently selected forest
  ' History:
  ' 19-03-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetCurrentForest(ByRef ndxThis As XmlNode, ByRef ndxEtree As XmlNode) As Boolean
    Try
      ' Validate
      If (loc_intLineNum < 0) OrElse (loc_intLineNum >= intSectSize) Then
        ' Set the starting point at the firs forest
        If (Not GetFirstForest(pdxCurrentFile, ndxThis)) Then Return False
      Else
        ' Get the currently selected forest
        ndxThis = GetForest(loc_intLineNum)
      End If
      ' Get the currently selected etree element
      If (loc_xndList Is Nothing) Then
        ndxEtree = Nothing
      ElseIf (loc_xndList.Count = 0) OrElse (loc_intListNum < 0) Then
        ndxEtree = Nothing
      Else
        ndxEtree = loc_xndList(loc_intListNum)
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetCurrentForest error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   EditorSelectionChanged
  ' Goal:   Do what is needed for the editor selection changes...
  ' History:
  ' 28-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub EditorSelectionChanged(ByRef rtfThis As RichTextBox, ByRef rtfSynt As RichTextBox, _
                                    ByRef rtfPde As RichTextBox, ByRef rtfFeat As RichTextBox)
    Dim intI As Integer               ' Counter
    Dim ndxForest As XmlNode          ' Forest entry
    Dim ndxNext As XmlNode = Nothing  ' Dummy <eTree>

    Try
      ' Make sure we are not processing a space
      If (bEdtSpace) Then Exit Sub
      ' Show that we are calculating
      Status("Calculating...")
      ' Undo the coloring of the previous selection
      SelectFromList(rtfThis, loc_intListNum, Color.White)
      ' Clear the previous line
      ClearCorefLine(rtfThis)
      ' Determine where we are
      loc_intPos = rtfThis.SelectionStart
      loc_intLen = rtfThis.SelectionLength
      ' Get the text of the RTF box
      strCurrentText = rtfThis.Text
      ' Determine the position
      ' NOTE: this can NOT be done using [GetLineFromCharIndex] or [GetFirstCharIndexOfCurrentLine]
      '   because these functions work only on truncated lines
      '   What we want is non-truncated full lines, as stored in [arLeft()] for instance
      If (GetRtfPosition(loc_intPos, loc_intLineNum, loc_intCharNum, loc_intLeft)) Then
        ' Check if the linenumber is valid
        If (loc_intLineNum >= intSectSize) Then
          ' Show we are out of bounds
          Status("(End of file)")
        Else
          ' Get the correct forest node
          ndxForest = GetForest(loc_intLineNum)
          If (ndxForest Is Nothing) Then Exit Sub
          ' Sort the relevant constituents from this line 
          loc_xndList = ndxForest.SelectNodes(".//eTree[@from<=" & loc_intCharNum + 1 & _
                            " and @to>" & loc_intCharNum + loc_intLen & "]")
          ' Build a stack of selectable nodes from this collection
          loc_arSel.Clear()
          For intI = 0 To loc_xndList.Count - 1
            ' what project are we?

            ' Are we allowed to select ANY constituent?
            ' AD-HOC Sonar: (pdxCurrentFile.SelectSingleNode("./descendant::forestGrp").Attributes("File").Value Like "WR-*") 
            If (bAnyConstituent) OrElse (Not bRefEditMode) Then
              ' Add this to the selectables stack
              loc_arSel.Add(loc_xndList(intI))
            Else
              ' Check if this element should be added to the stack of selectables
              Select Case LabelType(loc_xndList(intI))
                Case "Must", "MayAny", "MayDst"
                  ' Double check: don't allow selection of a unique child (e.g. PRO$) that has 
                  '  a referential parent (e.g. NP-OB1)
                  If (Not UniqueChildOfCorefXP(loc_xndList(intI), ndxNext)) Then
                    ' Add this to the selectables stack
                    loc_arSel.Add(loc_xndList(intI))
                  End If
              End Select
            End If
          Next intI
          ' Also store the DEEPEST (lowest level) node
          ndxCurrentNode = loc_xndList(loc_xndList.Count - 1)
          ' Set the index to the top of the array, which will be decremented immediately
          '   when [ShowConstituent()] down below is being called
          loc_intListNum = loc_arSel.Count
          ' Show the translation of this line
          ShowPde(rtfThis, rtfPde, rtfThis.SelectionStart)
        End If
      End If
      '' Only continue if we are not in AutoBusy)
      'If (bAutoBusy) Then Exit Sub
      ' Make sure [Calculating] does not stay on infinitely
      Status("")
      ' At least show the first constituent, going 1 down from [loc_arSel.Count] 
      ShowConstituent(rtfThis, False)
      ' Show the features of the currently selected node
      ShowFeat(rtfFeat)
      ' Switch off the ANy constituent flag
      bAnyConstituent = False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/EditorSelectionChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ClearCorefLine
  ' Goal:   Clear the currently drawn lines (if any)
  ' History:
  ' 12-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub ClearCorefLine(ByRef rtbThis As RichTextBox)
    Try
      ' Validate
      If (loc_grThis Is Nothing) Then InitGraphics(rtbThis)
      '' Redraw the rich text area
      'rtbThis.Refresh()
      ' Clear the previous line(s)
      ShowCorefDown(rtbThis, loc_ndxLast, True)
      'loc_grThis.DrawLines(Pens.Transparent, loc_arPoints)
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ClearCorefLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ChainDown
  ' Goal:   Starting from [xndThis] build a chain down
  ' History:
  ' 18-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function ChainDown(ByRef xndThis As XmlNode) As Boolean
    Dim xndNext As XmlNode  ' Next node in the chain down
    Dim xndDst As XmlNode   ' Destination node
    Dim bInferred As Boolean = False  ' We met an "Inferred" one
    Dim intMaxDepth As Integer = 40   ' Maximum number of nodes we show...

    Try
      ' Start from [xndThis]
      xndNext = xndThis
      ' Clear the stack
      loc_arLine.Clear()
      ' Loop through
      While (Not xndNext Is Nothing) AndAlso (Not bInferred) AndAlso (intMaxDepth > 0)
        ' Don't go further down if some relation is only "Inferred"
        bInferred = (CorefAttr(xndNext, "RefType") = strRefInferred)
        ' Is it already on the stack?
        If (loc_arLine.Exists(xndNext)) Then
          ' There is a circular coreference chain!
          MsgBox("Circular coreference chain...")
          ' Return failure
          Return False
        Else
          ' Push it on the stack
          loc_arLine.Add(xndNext)
        End If
        ' Get the destination node
        xndDst = CorefDst(xndNext)
        ' Go to the next one
        xndNext = xndDst
        ' Decrement the counter
        intMaxDepth -= 1
      End While
      ' Was the last relation "Inferred"?
      If (bInferred) Then
        ' Then the last one should still be pushed on the stack
        loc_arLine.Add(xndNext)
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ChainDown error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowLine
  ' Goal:   Show the line from the source to the destination node in the specified color
  ' History:
  ' 18-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub ShowLine(ByRef rtbThis As RichTextBox, ByRef xndSrc As XmlNode, ByRef xndDst As XmlNode, _
       ByVal bClear As Boolean, ByVal bIsContinuation As Boolean, ByVal bDotted As Boolean)
    Dim penThis As Pen  ' How the current pen looks like
    Dim capArrow As New Drawing2D.LineCap

    Try
      ' Validate
      If (rtbThis Is Nothing) OrElse (xndSrc Is Nothing) OrElse (xndDst Is Nothing) Then Exit Sub
      ' Possibly initialize drawing
      If (loc_grThis Is Nothing) Then InitGraphics(rtbThis)
      ' Try get the position of the source
      loc_arPoints(0) = GetPoint(rtbThis, xndSrc)
      ' Validate
      If (loc_arPoints(0).X = 0) AndAlso (loc_arPoints(0).Y = 0) Then Exit Sub
      ' Try get the position of the destination
      loc_arPoints(1) = GetPoint(rtbThis, xndDst)
      ' Validate
      If (loc_arPoints(1).X = 0) AndAlso (loc_arPoints(1).Y = 0) Then Exit Sub
      ' Adapt the destination slightly
      loc_arPoints(1).Y += 10
      ' If this line is a continuation, then the start point should also be offset
      If (bIsContinuation) Then loc_arPoints(0).Y += 10
      ' Make up the pen
      If (bClear) Then
        penThis = New Pen(Color.Transparent)
      Else
        penThis = New Pen(Color.DarkRed)
        ' Could be: penThis = New Pen(GetRefColor(CorefAttr(xndSrc, "RefType")))
      End If
      ' Set the start and end linecap
      With penThis
        .StartCap = Drawing2D.LineCap.RoundAnchor
        .EndCap = Drawing2D.LineCap.ArrowAnchor
        ' Should the line be dotted?
        If (bDotted) Then .DashStyle = Drawing2D.DashStyle.Dash
      End With
      ' Draw the line
      loc_grThis.DrawLines(penThis, loc_arPoints)
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ShowLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   GetPoint
  ' Goal:   Get the point on the screen where the node [xndThis] is located
  ' History:
  ' 18-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetPoint(ByRef rtbThis As RichTextBox, ByRef xndThis As XmlNode) As Point
    Dim intPos As Integer     ' Character position

    Try
      ' Validate
      If (rtbThis Is Nothing) OrElse (xndThis Is Nothing) Then Exit Function
      ' Possibly initialize drawing
      If (loc_grThis Is Nothing) Then InitGraphics(rtbThis)
      ' Calculate the points
      intPos = GetCharIndex(xndThis)
      If (intPos >= 0) Then
        ' First point is there...
        Return rtbThis.GetPositionFromCharIndex(intPos)
      Else
        ' No luck - return failure
        Return Nothing
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetPoint error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowCorefDown
  ' Goal:   Draw lines from source to destination, the whole chain DOWN
  ' History:
  ' 18-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub ShowCorefDown(ByRef rtbThis As RichTextBox, ByRef xndThis As XmlNode, _
                            ByVal bClear As Boolean)
    Dim intI As Integer       ' Counter
    Dim intNum As Integer     ' Number of lines to go down to
    Dim xndSrc As XmlNode     ' Source node
    Dim xndDst As XmlNode     ' Destination node
    Dim bDotted As Boolean    ' Whether the line should be solid or dotted

    Try
      ' Build the chain down
      If (Not ChainDown(xndThis)) Then Exit Sub
      ' Note the number of lines
      intNum = loc_arLine.Count
      ' There should at least be TWO points!!!
      If (loc_arLine.Count < 2) Then Exit Sub
      ' Refresh the scene since we are actually going to draw...
      rtbThis.Refresh()
      ' Start with the first node
      xndSrc = loc_arLine.Item(0)
      ' Walk the chain down
      For intI = 1 To intNum - 1
        ' Get the destination
        xndDst = loc_arLine.Item(intI)
        ' Determine whether the line should be dotted or not
        bDotted = (CorefAttr(xndSrc, "RefType") = strRefInferred)
        ' Draw a line
        ShowLine(rtbThis, xndSrc, xndDst, bClear, intI > 1, bDotted)
        ' Go to the next one
        xndSrc = xndDst
      Next intI
      ' Note this line for future reference
      loc_ndxLast = xndThis
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ShowCorefDown error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowCorefLine
  ' Goal:   Draw lines from each source to destination node
  ' History:
  ' 12-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub ShowCorefLine(ByRef rtbThis As RichTextBox, ByRef xndThis As XmlNode)
    Dim pntThis As New Point          ' one point
    Dim intPos As Integer             ' Character position
    Dim xndDst As XmlNode             ' Destination node

    Try
      ' Validate
      If (loc_grThis Is Nothing) Then InitGraphics(rtbThis)
      '' Redraw the rich text area
      'rtbThis.Refresh()
      ' Clear the previous line
      loc_grThis.DrawLines(Pens.Transparent, loc_arPoints)
      ' Calculate the points
      intPos = GetCharIndex(xndThis)
      If (intPos >= 0) Then
        ' First point is there...
        pntThis = rtbThis.GetPositionFromCharIndex(intPos)
        loc_arPoints(0) = pntThis
        ' The second point should be calculated from the destination
        xndDst = CorefDst(xndThis)
        If (Not xndDst Is Nothing) Then
          ' Calculate the position
          intPos = GetCharIndex(xndDst)
          If (intPos >= 0) Then
            ' Add this point
            pntThis = rtbThis.GetPositionFromCharIndex(intPos)
            ' Subtract something from the Y-coordinate
            pntThis.Y += 10
            loc_arPoints(1) = pntThis
            ' Draw the line
            loc_grThis.DrawLines(Pens.DarkRed, loc_arPoints)
          End If
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ShowCorefLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   GetCharIndex
  ' Goal:   Get the character position for this node
  ' History:
  ' 12-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetCharIndex(ByRef xndThis As XmlNode) As Integer
    Dim intLeft As Integer  ' The left position of the selected node

    Try
      ' Validate
      If (xndThis Is Nothing) Then Return -1
      ' Get the left position
      intLeft = GetLeft(xndThis)
      ' Is this valid?
      If (intLeft < 0) Then Return -1
      ' Get the character position
      Return intLeft + xndThis.Attributes("from").Value
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetCharIndex error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetAutoPos
  ' Goal:   Show the text in the PDE window
  ' History:
  ' 01-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetAutoPos() As Integer
    Return loc_intAutoPos
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowPde
  ' Goal:   Show the text in the PDE window
  ' History:
  ' 10-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ShowPde(ByRef rtfThis As RichTextBox, ByRef rtfPdeView As RichTextBox, _
                     ByVal intPos As Integer)
    Dim intLine As Integer    ' The current line number

    Try
      ' Validate
      If (rtfPdeView Is Nothing) Then Exit Sub
      ' Determine the current line number (if possible)
      If (Not GetRtfLineNumber(rtfThis, intPos, intLine)) Then Exit Sub
      ' Switch over
      ShowPdeLine(rtfThis, rtfPdeView, intLine)
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ShowPde error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowPdeEtree
  ' Goal:   Show the PDE translation with the current <eTree> as center
  ' History:
  ' 10-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ShowPdeEtree(ByVal intEtreeId As Integer)
    Dim intLine As Integer    ' The current line number
    Dim intMin As Integer = 0 ' Minimum
    Dim intMax As Integer     'Maximum

    Try
      ' Validate
      If (intEtreeId < 0) Then Exit Sub
      ' Get the line number
      intMax = pdxList.Count - 1
      intLine = GetLineIdx(intEtreeId, intMin, intMax)
      ' Show the PDE
      ShowPdeLine(frmMain.tbEdtMain, frmMain.tbMainPde, intLine)
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ShowPdeEtree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowPdeLine
  ' Goal:   Show the text in the PDE window
  ' History:
  ' 10-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ShowPdeLine(ByRef rtfThis As RichTextBox, ByRef rtfPdeView As RichTextBox, _
                     ByVal intLine As Integer)
    Dim ndxThis As XmlNode    ' Start node
    Dim intI As Integer       ' Counter
    Dim intMin As Integer     ' Where we start from
    Dim intMax As Integer     ' The maximum line number
    Dim strId As String       ' The ID of this line
    Dim intBefore As Integer = 5  ' Number of lines that need to precede me
    Dim intOffset As Integer = 20 ' Number of lines of the window we show
    Dim intPos As Integer         ' Position in string

    Try
      ' Validate
      If (rtfPdeView Is Nothing) Then Exit Sub
      ' Clear the syntax RTF
      rtfPdeView.Text = "" : intPos = 1
      ' Make sure we are not scrolling (how??)
      rtfPdeView.Visible = False
      ' Determine the minimum and maximum line number
      intMin = intLine - intOffset
      ' Debug.Print(loc_intLineNum)
      If (intMin < 0) Then intMin = intSectFirst Else intMin = Math.Max(intSectFirst, intMin + intSectFirst)
      intMax = intLine + intOffset
      If (intMax > intSectLast) Then intMax = intSectLast Else intMax = Math.Min(intSectLast, intMax + intSectFirst)
      ' Visit all the lines in this section
      For intI = intMin To intMax
        ' Set our poitn at this line
        ndxThis = pdxList(intI)
        If (ndxThis IsNot Nothing) Then
          ' Get the ID of this line
          If (ndxThis.Attributes("Location") Is Nothing) Then
            strId = "[s" & Format(ndxThis.Attributes("forestId").Value, "0000") & "] "
          Else
            strId = "[" & ndxThis.Attributes("Location").Value & "] "
          End If
          ' Is this the currently selected line?
          If (intI = intLine + intSectFirst) Then
            ' Show line emphatic
            AppendToRtb(rtfPdeView, strId, Color.DarkRed, fntPdeEmph)
            AppendToRtb(rtfPdeView, ForestText(ndxThis, strCurrentTransLang) & vbCrLf, Color.DarkBlue, fntPde)
            'ElseIf (intI > intLine + intSectFirst + 2) AndAlso (intSectLast - intSectFirst > 400) Then
            '  ' Get out before we need to do the whole translation!
            '  Exit For
          Else
            ' Are we at the right line?
            If (intI = intLine + intSectFirst - intBefore) Then
              ' Keep track of where we need to be
              intPos = rtfPdeView.SelectionStart
            End If
            ' Show line normal
            AppendToRtb(rtfPdeView, strId & ForestText(ndxThis, strCurrentTransLang) & vbCrLf, Color.Black, fntPde)
          End If
        Else
          ' There really is a problem...
          Stop
        End If
      Next intI
      ' Make sure we can scroll again
      rtfPdeView.Visible = True
      '' Set the selection to the correct point
      'rtfPdeView.SelectionStart = intPos
      ' Should we scroll?
      If (rtfPdeView.GetPositionFromCharIndex(intPos).Y + 200 > rtfPdeView.Height) Then
        ' Yes, force scrolling 
        rtfPdeView.Focus()
        rtfPdeView.SelectionStart = intPos
        ' Set focus back to the main window
        rtfThis.Focus()
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ShowPdeLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowFeat
  ' Goal:   Show the features associated with the currently selected node
  ' History:
  ' 27-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ShowFeat(ByRef rtfFeat As RichTextBox)
    Dim strFsType As String   ' The type of <fs>
    Dim ndxThis As XmlNode    ' Start node
    Dim ndxRoot As XmlNode    ' Root node
    Dim ndxLf As XmlNode      ' Leaf
    Dim ndxFs As XmlNodeList  ' All <fs> elements that are my children
    Dim ndxF As XmlNodeList   ' All <f> elements children of <fs>
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    Dim strL As String = ""   ' Lemma
    Dim strD As String = ""   ' Definition
    Dim strW As String = ""   ' Word
    Dim strP As String = ""   ' POS

    Try
      ' Validate
      If (loc_arSel Is Nothing) Then Exit Sub
      If (rtfFeat Is Nothing) Then Exit Sub
      If (loc_intListNum < 0) OrElse (loc_intListNum >= loc_arSel.Count) Then Exit Sub
      ' Clear the features RTF
      rtfFeat.Text = ""
      ' Get the currently selected node
      ndxThis = loc_arSel.Item(loc_intListNum)
      ' Validate again
      If (ndxThis Is Nothing) Then Exit Sub
      ' Do we have an IP number?
      If (Not ndxThis.Attributes("IPnum") Is Nothing) Then
        ' Show the IPnumber
        rtfFeat.AppendText("IPnum=" & ndxThis.Attributes("IPnum").Value & vbCrLf)
      End If
      ' Get all the <fs> children
      ndxFs = ndxThis.SelectNodes("./fs")
      ' Walk them
      For intI = 0 To ndxFs.Count - 1
        ' Retrieve the type attribute of this one
        strFsType = ndxFs(intI).Attributes("type").Value
        ' Add to the RTF
        rtfFeat.AppendText(strFsType & ": " & vbCrLf)
        ' Get the <f> children
        ndxF = ndxFs(intI).SelectNodes("./f")
        ' Walk all the <f> children
        For intJ = 0 To ndxF.Count - 1
          ' Retrieve attribute name and value
          If (ndxF(intJ) Is Nothing) Then
            ' No features specified
            rtfFeat.AppendText(vbTab & "(no features specified)" & vbCrLf)
          Else
            ' Add this feature and value
            With ndxF(intJ)
              rtfFeat.AppendText(vbTab & .Attributes("name").Value & "=" & .Attributes("value").Value & vbCrLf)
            End With
          End If
        Next intJ
      Next intI
      ' Show myself
      rtfFeat.AppendText("This node:" & vbCrLf & vbTab & NodeInfo(ndxThis) & vbCrLf)
      ' Try to add lemma information
      ' Get the lemma and the definition
      If (LemmaOneWord(ndxCurrentNode, strL, strD)) Then
        strP = ndxCurrentNode.Attributes("Label").Value
        ndxLf = ndxCurrentNode.SelectSingleNode("./descendant::eLeaf[@Type='Vern'][1]")
        If (ndxLf Is Nothing) Then
          strW = "-"
        Else
          strW = ndxLf.Attributes("Text").Value
        End If
        rtfFeat.AppendText(vbCrLf & "This word: [" & strW & "] (" & strP & ")" & vbCrLf & "Lemma:" & vbTab & strL & vbCrLf & "Definition(s):" & vbCrLf & strD & vbCrLf)
      End If
      ' Also try to determine and show where we point to
      ndxRoot = GetRootNode(ndxThis)
      ' Is this the same?
      If (Not EqualNodes(ndxThis, ndxRoot)) AndAlso (Not ndxRoot Is Nothing) Then
        ' Show the chain root
        rtfFeat.AppendText("Chain root:" & vbCrLf & vbTab & NodeInfo(ndxRoot) & vbCrLf)
      End If
      ' Possibly adapt what is being shown
      frmMain.FeatShowAdapt(ndxCurrentNode)
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ShowSyntax error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowSyntax
  ' Goal:   Show the syntax of the constituent we are in
  ' History:
  ' 28-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ShowSyntax(ByRef rtfSynt As RichTextBox)
    Dim ndxForest As XmlNode = Nothing  ' This line's forest

    Try
      ' Check if we are visible
      If (Not rtfSynt.Visible) Then Exit Sub
      ' Check the line number
      If (loc_intLineNum < 0) OrElse (loc_intListNum < 0) Then Exit Sub
      ' The following is perhaps not necessary?
      '' Check whether the line number has changed or not
      'If (loc_intLineNum <> loc_intLinePrev) Then
      ' Clear the syntax RTF
      rtfSynt.Text = ""
      '' Get the currently selected node
      'ndxThis = loc_arSel.Item(loc_intListNum)
      ' Do we have a selected node?
      If (ndxCurrentSel Is Nothing) Then
        ' Get the correct forest node
        ndxForest = GetForest(loc_intLineNum)
      Else
        ' Derive forest from current selected node
        ndxForest = ndxCurrentSel.SelectSingleNode("./ancestor::forest[1]")
      End If
      If (ndxForest Is Nothing) Then Exit Sub
      ' Reset the syntax level to 0
      loc_intLevelSyntax = 0
      ' Retrieve the text for this node
      If (GetSyntax(ndxForest, rtfSynt)) Then
        ' Everything okay!
      End If
      ' Adapt the line number
      loc_intLinePrev = loc_intLineNum
      ' Show the current tree in the edit line
      frmMain.tbSyntLine.Text = "[" & ndxForest.Attributes("forestId").Value & "] " & _
        GetSeg(ndxForest, "org")
      'End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ShowSyntax error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowTree
  ' Goal:   Show the syntax of the constituent we are in
  ' History:
  ' 03-11-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ShowTree(ByRef litThis As Lithium.LithiumControl)
    Dim ndxForest As XmlNode = Nothing  ' This line's forest
    Dim ndxThis As XmlNode = Nothing    ' Selected line

    Try
      ' Check the line number
      If (loc_intLineNum < 0) OrElse (loc_intListNum < 0) Then Exit Sub
      ' Get the currently selected node
      ndxThis = ndxCurrentSel
      ' If nothing is selected, we cannot show a tree
      If (ndxCurrentSel Is Nothing) Then
        ' Show that nothing has been selected
        Status("No particular constituent is selected")
        ' Get the forest from the selected linenumber
        ndxForest = GetForest(loc_intLineNum)
      Else
        ' Get the correct forest node
        ' ndxForest = GetForest(loc_intLineNum)
        ' Derive forest from current selected node
        ndxForest = ndxThis.SelectSingleNode("./ancestor::forest[1]")
      End If
      ' Validate on forest
      If (ndxForest Is Nothing) Then Exit Sub
      ' ========= DEBUG =========
      With litThis
        .LayoutEnabled = True
        .LayoutDirection = TreeDirection.Vertical
        .ConnectionType = Lithium.ConnectionType.Traditional
      End With
      ' =========================
      ' Retrieve the text for this node
      If (Not MakeLitTree(litThis, ndxForest, ndxThis)) Then
        ' There is a problem
        Status("Could not draw a tree")
      Else
        ' Show we are ready
        Status("Ready")
      End If
      ' Show the current text in the edit line
      frmMain.tbTreeLine.Text = "[" & ndxForest.Attributes("forestId").Value & "] " & _
        GetSeg(ndxForest, "org")
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ShowTree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
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
            If (DoLike(strLabel, "[IC]P[-_]*|S|SBAR")) Then
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
                  If (DoLike(strParentL, "[IC]P[_-]*|S|SBAR")) Then
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
      HandleErr("modEditor/GetSyntax error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AppendSyntax
  ' Goal:   Append text to the richtextbox in the indicated manner
  ' History:
  ' 28-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub AppendToRtb(ByRef rtfSynt As RichTextBox, ByVal strText As String, _
                ByVal colThis As System.Drawing.Color, ByRef fntThis As System.Drawing.Font)
    Try
      With rtfSynt
        .SelectionFont = fntThis
        .SelectionColor = colThis
        .AppendText(strText)
        .SelectionFont = fntNormal
        .SelectionColor = Color.Black
      End With
    Catch ex As Exception
      ' Retreat silently
      Exit Sub
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   GetSelectedConst
  ' Goal:   Find out which constituent is selected
  ' History:
  ' 23-11-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetSelectedConst(Optional ByVal bBottom As Boolean = False) As XmlNode
    Try
      ' Determine which one is now selected
      If (loc_arSel Is Nothing) OrElse (loc_arSel.Count = 0) OrElse (loc_intListNum < 0) Then
        ' No valid constituent selected
        Return Nothing
      End If
      ' A valid constituent must have been selected - get it
      If (bBottom) Then
        Return loc_arSel.Item(0)
      Else
        Return loc_arSel.Item(loc_intListNum)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetSelectedConst error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SelectCoref
  ' Goal:   Select the source or the goal of a coreference relation
  '         The SOURCE is selected if this is the first selection
  '         The DESTINATION can be more than one, and those are the next selections
  ' History:
  ' 11-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SelectCoref(ByRef rtfThis As RichTextBox)
    Dim xndThis As XmlNode  ' Currently selected node
    Dim xndNext As XmlNode  ' Parent <eTree>
    Dim colThis As Color    ' The background colour to be selected
    Dim intIdx As Integer   ' Index of an element
    Dim intI As Integer     ' The number of the selected constituent
    Dim intLeft As Integer  ' Result of [GetLeft]

    Try
      ' Determine which one is now selected
      If (loc_arSel Is Nothing) OrElse (loc_arSel.Count = 0) OrElse (loc_intListNum < 0) Then
        ' No valid constituent selected
        Status("No valid constituent is selected!")
        Exit Sub
      End If
      ' A valid constituent must have been selected - get it
      xndThis = loc_arSel.Item(loc_intListNum)
      ' Does it already exist?
      If (loc_arCoref.Exists(xndThis)) Then
        ' It already exists, so we need to DESELECT
        ' Delete this from the collection
        intI = loc_arCoref.Find(xndThis)
        ' Was this the FIRST item?
        If (intI = 0) Then
          ' Then all corefs must be deselected
          ClearCoref(rtfThis)
        Else
          ' Clear its selection in the richtextbox
          colThis = Color.White
          intLeft = GetLeft(xndThis)
          ' Do validate
          If (intLeft < 0) Then Exit Sub
          ' If all okay: select
          SelectEtree(rtfThis, xndThis, intLeft, colThis)
          ' Delete this item from the collection
          loc_arCoref.DelItem(intI)
          ' Redraw ALL constituents that are (1) ancestor of [xndThis] and (2) selected
          xndNext = xndThis.ParentNode
          ' Go into a loop
          While (Not xndNext Is Nothing)
            ' See if this needs redrawing
            If (xndNext.Name = "eTree") Then
              ' Check if it exists in the set of selections
              intIdx = loc_arCoref.Find(xndNext)
              ' Is this index positive?
              Select Case intIdx
                Case 0
                  ' This is a source node
                  SelectEtree(rtfThis, xndNext, loc_intLeft, colCorefSource)
                Case Is > 0
                  ' This is a destination node
                  SelectEtree(rtfThis, xndNext, loc_intLeft, colCorefDest)
              End Select
            End If
            ' Go to the next parent
            xndNext = xndNext.ParentNode
          End While
        End If
      Else
        ' It doesn't exist yet, so we need to SELECT
        ' Determine the color
        If (loc_arCoref.Count = 0) Then
          ' Set the source colour:
          colThis = colCorefSource
        Else
          ' Set the destination colour: green
          colThis = colCorefDest
        End If
        ' Set its selection in the richtextbox
        SelectEtree(rtfThis, xndThis, loc_intLeft, colThis)
        ' Add this to the collection
        loc_arCoref.AddUnique(xndThis)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/SelectCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ClearCoref
  ' Goal:   Clear all selected constituents
  ' History:
  ' 11-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ClearCoref(ByRef rtfThis As RichTextBox)
    Dim xndThis As XmlNode  ' Currently selected node
    Dim intI As Integer     ' The number of the selected constituent
    Dim intNum As Integer   ' Number of constituents selected
    Dim intLeft As Integer  ' Result of getleft
    Dim colThis As Color    ' The background colour to be selected

    Try
      ' How many are there?
      intNum = loc_arCoref.Count
      ' Set the deselect colour
      colThis = Color.White
      ' Walk through all constituents selected
      For intI = intNum - 1 To 0 Step -1
        ' Get this node
        xndThis = loc_arCoref.Item(intI)
        ' Clear its selection in the richtextbox
        intLeft = GetLeft(xndThis) : If (intLeft < 0) Then Exit Sub
        SelectEtree(rtfThis, xndThis, intLeft, colThis)
        ' Make sure its correct colors are shown
        ShowOneCoref(rtfThis, xndThis)
        ' Delete it
        loc_arCoref.DelItem(intI)
      Next intI
      ' Make sure the selection is to the left
      rtfThis.SelectionLength = 0
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ClearCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   GetLeft
  ' Goal:   Get the position of the start of the line containing this XML node
  ' History:
  ' 11-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetLeft(ByRef xndThis As XmlNode) As Integer
    Dim intId As Integer      ' ForestId value
    Dim xndForest As XmlNode  ' The forest ancestor

    Try
      ' Is the input valid?
      If (xndThis Is Nothing) Then Return -1
      ' Do we have a valide [arLeft] array?
      If (arLeft Is Nothing) Then Return -1
      ' Get the forest ancestor
      xndForest = xndThis.SelectSingleNode("./ancestor::forest[1]")
      ' Any valid return
      If (xndForest Is Nothing) Then
        ' Return failure
        Return -1
      Else
        ' Get the associated left value
        intId = xndForest.Attributes("forestId").Value - intSectFirst
        ' If this value is <0 then it is beyond our reach
        If (intId < 0) Then
          Return -1
        Else
          ' Validate
          If (intId > UBound(arLeft)) Then Return -1
          Return arLeft(intId - 1)
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetLeft error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowConstituent
  ' Goal:   Show the next higher or lower constituent
  ' History:
  ' 28-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ShowConstituent(ByRef rtfThis As RichTextBox, ByVal bIncrement As Boolean)
    Dim intSize As Integer    ' The size of the selected constituent's array

    Try
      ' Do we have a list?
      If (loc_arSel.Count = 0) Then
        ' Set the Id number
        loc_intId = -1
        ndxCurrentSel = Nothing
        ' Exit
        Exit Sub
      End If
      ' Get the size
      intSize = loc_arSel.Count
      ' Increment the listnumber, if this is possible
      If (loc_intListNum < intSize - 1) Then
        ' Make the previous selection WHITE
        SelectFromList(rtfThis, loc_intListNum, Color.White)
        ' Increment it
        loc_intListNum += 1
        ' Is this one still okay?
        If (LabelType(loc_arSel.Item(loc_intListNum)) = "") AndAlso (Not bAnyConstituent) Then
          ' Show that nothing is selected
          Status("(no constituent selected)")
          ' Reset the Id number
          loc_intId = -1
          ndxCurrentSel = Nothing
        Else
          ' Try to set this one to gray
          SelectFromList(rtfThis, loc_intListNum, colSelect)
          ' Show where we are now
          ShowPosition(rtfThis)
          ' Set the Id number
          loc_intId = loc_arSel.Item(loc_intListNum).Attributes("Id").Value
          ' Set the selected constituent
          ndxCurrentSel = loc_arSel.Item(loc_intListNum)
        End If
      Else
        ' Try to decrement if this is possible
        If (loc_intListNum > 0) Then
          ' Decrement it
          loc_intListNum -= 1
          ' Is this one still okay?
          If (LabelType(loc_arSel.Item(loc_intListNum)) = "") AndAlso (Not bAnyConstituent) Then
            ' Show that nothing is selected
            Status("(no constituent selected)")
            ' Reset the Id number
            loc_intId = -1
            ndxCurrentSel = Nothing
          Else
            ' Try to set this one to gray
            SelectFromList(rtfThis, loc_intListNum, colSelect)
            ' Show where we are now
            ShowPosition(rtfThis)
            ' Set the Id number
            loc_intId = loc_arSel.Item(loc_intListNum).Attributes("Id").Value
            ' Set the selected constituent
            ndxCurrentSel = loc_arSel.Item(loc_intListNum)
          End If
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ShowConstituent error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   SelectFromList
  ' Goal:   Try to set element with number [intNum] from the list in the specified color
  ' History:
  ' 28-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SelectFromList(ByRef rtfThis As RichTextBox, ByVal intNum As Integer, _
                            ByVal colThis As System.Drawing.Color)
    Dim xndThis As XmlNode  ' The currently selected element

    Try
      ' Was the datarowlist already built up?
      If (loc_arSel Is Nothing) Then Exit Sub
      ' Is the index within bounds?
      If (loc_intListNum < 0) OrElse (loc_intListNum >= loc_arSel.Count) Then Exit Sub
      ' Get the currently selected element
      xndThis = loc_arSel.Item(loc_intListNum)
      ' Is this element already selected?
      If (Not IsProtected(xndThis)) Then
        ' Process this element
        SelectEtree(rtfThis, xndThis, loc_intLeft, colThis)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/SelectFromList error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   IsProtected
  ' Goal:   Is this node protected from overwriting?
  '         This is true, if the node or one of its parents exists in [loc_arCoref]
  ' History:
  ' 11-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsProtected(ByRef xndThis As XmlNode) As Boolean
    Dim xndNext As XmlNode  ' Next node

    Try
      ' Take the first step
      xndNext = xndThis
      ' Go into a loop
      Do
        ' Is this one in the array?
        If (loc_arCoref.Exists(xndNext)) Then
          ' Return true
          Return True
        End If
        ' Take the next step up
        xndNext = xndNext.ParentNode
      Loop Until (xndNext Is Nothing) OrElse (xndNext.Name <> "eTree")
      ' Check all the children
      For Each xndNext In xndThis.ChildNodes
        ' Check if the child (or its children) are protected
        If (ChildProtected(xndNext)) Then
          ' Return success
          Return True
        End If
      Next xndNext
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/IsProtected error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsProtected
  ' Goal:   Is this node or its children protected from overwriting?
  '         This is true, if the node or one of its parents exists in [loc_arCoref]
  ' History:
  ' 11-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function ChildProtected(ByRef xndThis As XmlNode) As Boolean
    Dim xndNext As XmlNode  ' Next node

    Try
      ' Is this node protected?
      If (xndThis.Name = "eTree") AndAlso (loc_arCoref.Exists(xndThis)) Then Return True
      ' Otherwise check its children
      For Each xndNext In xndThis.ChildNodes
        If (ChildProtected(xndNext)) Then Return True
      Next xndNext
      ' Otherwise return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ChildProtected error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CurrentItem
  ' Goal:   Get the currently selected item (if any)
  ' History:
  ' 11-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CurrentItem() As XmlNode
    Try
      ' Was the datarowlist already built up?
      If (loc_arSel Is Nothing) Then Return Nothing
      ' Is the index within bounds?
      If (loc_intListNum < 0) OrElse (loc_intListNum >= loc_arSel.Count) Then Return Nothing
      ' Return this element
      Return loc_arSel.Item(loc_intListNum)
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/CurrentItem error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SelectEtree
  ' Goal:   Select the XML node in the specified color
  ' History:
  ' 28-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SelectEtree(ByRef rtfThis As RichTextBox, ByRef ndxThis As XmlNode, _
                         ByVal intLeft As Integer, ByVal colThis As System.Drawing.Color)
    Dim intPos As Integer     ' Start of selected position
    Dim intLen As Integer     ' Length of selection

    Try
      ' Check validity
      If (rtfThis Is Nothing) OrElse (ndxThis Is Nothing) Then Exit Sub
      ' Access this element
      With ndxThis
        ' Determine what should be selected
        intPos = .Attributes("from").Value
        intLen = .Attributes("to").Value - intPos + 1
        ' Select this in the indicated colour
        SelectBackground(rtfThis, colThis, intLeft + intPos - 1, intLen)
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/SelectEtree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   SetSelection
  ' Goal:   Set the selection to the XML node's starting point
  ' History:
  ' 29-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SetSelection(ByRef rtfThis As RichTextBox, ByRef ndxThis As XmlNode, _
                          ByVal intLeft As Integer)
    Dim intPos As Integer     ' Start of selected position
    Dim intLen As Integer     ' Length of selection

    Try
      ' Check validity
      If (rtfThis Is Nothing) OrElse (ndxThis Is Nothing) Then Exit Sub
      ' Access this element
      With ndxThis
        ' Determine what should be selected
        intPos = .Attributes("from").Value
        intLen = 0
        ' Put selection there
        With rtfThis
          .SelectionStart = intLeft + intPos - 1
          .SelectionLength = intLen
        End With
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/SetSelection error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   SelectBackground
  ' Goal:   Select the background colour of the indicated [start] + [len].
  '         Return the selection to the point where it was before.
  ' History:
  ' 28-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SelectBackground(ByRef rtfThis As RichTextBox, ByVal colThis As System.Drawing.Color, _
               ByVal intStart As Integer, ByVal intLen As Integer)
    Dim intThisPos As Integer ' The current selection
    Dim intThisLen As Integer ' Length of the current selection
    Dim fntNoUl As FontStyle  ' Currentfontstyle, but no underlining
    Dim fntDoUl As FontStyle  ' Current fontstyle, but with underlining

    Try
      ' Validate
      If (rtfThis Is Nothing) OrElse (intStart <= 0) Then Exit Sub
      ' make sure we are not interruptable
      bEdtSpace = True
      ' Note the current selection + length
      With rtfThis
        intThisPos = .SelectionStart
        intThisLen = .SelectionLength
        ' Select what needs to be selected according to the parameters [intStart], [intLen]
        .SelectionStart = intStart
        .SelectionLength = intLen
        ' Set the color of this selection
        .SelectionBackColor = colThis
        ' Do we need a rectangle around us?
        Select Case colThis
          Case colCorefSource
            ' The source should be underlined
            ' fntDoUl = .SelectionFont.Style Or FontStyle.Underline
            fntDoUl = .Font.Style Or FontStyle.Underline
            .SelectionFont = New Font(.Font, fntDoUl)
          Case colCorefDest
            ' Destination does NOT get special treatment...
          Case Else
            ' Anything that is NOT the source should not be underlined
            ' fntNoUl = .SelectionFont.Style And Not FontStyle.Underline
            fntNoUl = .Font.Style And Not FontStyle.Underline
            .SelectionFont = New Font(.Font, fntNoUl)
        End Select
        ' Return the selection to where it was
        .SelectionStart = intThisPos
        .SelectionLength = intThisLen
      End With
      ' Again interruptable
      bEdtSpace = False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/SelectBackground error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   HasCoref
  ' Goal:   Check whether the indicated XML node of type <eTree> and has a child
  '           of type <fs> with an attribute @type=coref
  ' History:
  ' 29-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function HasCoref(ByRef ndxThis As XmlNode) As Boolean
    Dim ndxChild As XmlNode  ' Result of query

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Is it of the correct type?
      If (ndxThis.Name = "eTree") Then
        ' Get the first child of type <fs> with attribute @type=coref
        ndxChild = ndxThis.SelectSingleNode("./fs[@type='coref']")
        ' Return the result
        Return (Not ndxChild Is Nothing)
      Else
        ' Return failure
        Return False
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/HasCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowPosition
  ' Goal:   Show where we are to the user
  ' History:
  ' 10-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ShowPosition(ByRef rtfThis As RichTextBox)
    Dim ndxThis As XmlNode      ' One node
    Dim ndxDst As XmlNode       ' Possible target/destination node
    Dim ndxForest As XmlNode    ' This line's forest
    Dim strType As String = ""  ' Possible reference type
    Dim strLabel As String = "" ' The label of the currently selected node
    Dim strLoc As String = ""   ' The location of the currently selected node
    Dim strState As String = "" ' The status to be shown to the user

    Try
      ' Validate input
      If (rtfThis Is Nothing) Then Exit Sub
      ' Is everything valid?
      If (loc_arSel Is Nothing) OrElse (loc_intListNum < 0) Then
        ' Show we are not sure where we are
        Status("(Cannot determine position)")
      ElseIf (loc_arSel.Count = 0) Then
        ' Only give the line number, not the label
        ' Get the correct forest node
        ndxForest = GetForest(loc_intLineNum)
        If (ndxForest Is Nothing) Then Exit Sub
        ' Get the location where we are
        ndxThis = ndxForest
        ' Got something?
        If (Not ndxThis Is Nothing) Then
          ' Get the location string
          strLoc = ndxThis.Attributes("Location").Value
          ' Show where we are
          Status("[" & strLoc & "] (no constituent selected)")
        Else
          ' Show we are not sure where we are
          Status("(Cannot determine position)")
        End If
      Else
        ' Get the current position
        If (loc_intListNum = loc_arSel.Count) Then
          ' There is no need to process this one!
          ndxThis = loc_arSel.Item(loc_intListNum - 1)
          'Exit Sub
        Else
          ndxThis = loc_arSel.Item(loc_intListNum)
        End If
        ' Access this element
        With ndxThis
          ' Show the label of this constituent
          strLabel = .Attributes("Label").Value
        End With
        ' Get the location where we are
        strLoc = NodeLocation(ndxThis)
        ' Show where we are
        strState = "[" & strLoc & "] " & strLabel
        ' Try get the reference type (if there is some)
        strType = CorefAttr(ndxThis, "RefType")
        ' Need to add coreference type?
        If (strType <> "") Then strState &= " (Coref Type=" & strType & ")"
        ' Try to get a target
        ndxDst = CorefDst(ndxThis)
        ' Did we get a target?
        If (Not ndxDst Is Nothing) Then
          ' Add the target's location etc
          strState &= " --> [" & NodeLocation(ndxDst) & "] " & NodeLabel(ndxDst) & " (" & _
            VernToEnglish(NodeText(ndxDst)) & ")"
          ' Draw a line
          ShowCorefDown(rtfThis, ndxThis, False)
        End If
        ' Show status
        Status(strState)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ShowPosition error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   NodeLocation
  ' Goal:   Get the <forest> location where we are!!
  ' History:
  ' 12-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function NodeLocation(ByRef ndxThis As XmlNode) As String
    Dim ndxForest As XmlNode    ' The forest parent

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      ' Get the nearest forest parent
      ndxForest = ndxThis.SelectSingleNode("./ancestor::forest[1]")
      ' Any valid result?
      If (ndxForest Is Nothing) Then
        ' Return failure
        Return Nothing
      End If
      ' Return the forest's <Location> attribute
      Return ndxForest.Attributes("Location").Value
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/NodeLocation error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   NodeLabel
  ' Goal:   Get the label (if any) of this node
  ' History:
  ' 12-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function NodeLabel(ByRef ndxThis As XmlNode) As String
    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      If (ndxThis.Name <> "eTree") Then Return ""
      ' Return the node's label
      Return ndxThis.Attributes("Label").Value
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/NodeLabel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   NodeText
  ' Goal:   Get the <eLeaf> text of this node
  ' History:
  ' 12-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function NodeText(ByRef ndxThis As XmlNode, Optional ByVal bDoStar As Boolean = False) As String
    Dim xndList As XmlNodeList  ' Result of SELECT
    Dim strBack As String = ""  ' Return string
    Dim intI As Integer         ' Counter
    Dim bPunct As Boolean       ' Flag for punctuation
    Dim bIsStar As Boolean      ' Flag for starred elements

    Try
      ' Validate input
      If (ndxThis Is Nothing) Then Return ""
      ' Get a list of <eLeaf> nodes
      xndList = ndxThis.SelectNodes(".//eLeaf[@Type='Vern' or @Type='Punct' or @Type='Star']")
      ' Any result?
      If (xndList Is Nothing) Then
        ' Return failure
        Return ""
      ElseIf (xndList.Count = 0) Then
        ' Return failure
        Return ""
      Else
        ' Set the starred flag
        bIsStar = (xndList(0).Attributes("Type").Value = "Star")
        ' Test for starred element
        If (bDoStar OrElse (Not bIsStar)) Then
          ' Get the first word
          strBack = xndList(0).Attributes("Text").Value
        End If
        ' Set the punctuation flag
        bPunct = (xndList(0).Attributes("Type").Value = "Punct")
        ' Build the return string
        For intI = 1 To xndList.Count - 1
          ' Set the starred flag
          bIsStar = (xndList(intI).Attributes("Type").Value = "Star")
          ' Check if this is a "Star" element
          If (bDoStar OrElse (Not bIsStar)) Then
            ' Add a space or not...
            If (Not bPunct) Then strBack &= " "
            ' Add a word
            strBack &= xndList(intI).Attributes("Text").Value
          End If
          ' Set the punctuation flag
          bPunct = (xndList(intI).Attributes("Type").Value = "Punct")
        Next intI
        ' Return the result
        Return Trim(strBack)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/NodeText error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   NodeSyntax
  ' Goal:   Get the text of this node broken up in syntactic units
  ' History:
  ' 10-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function NodeSyntax(ByRef ndxThis As XmlNode, Optional ByVal bHtml As Boolean = False) As String
    Dim strBack As String = ""  ' String to be returned

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      ' Do a recursive operation
      If (TravNode(ndxThis, "Syntax", strBack, bHtml)) Then
        ' Return the result
        Return strBack
      Else
        ' Return empty
        Return ""
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/NodeSyntax error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return ""
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  GetMood
  ' Goal :  Try to get the mood, if this is a finite verb
  ' History:
  ' 05-05-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Function GetMood(ByVal strLabel As String) As String
    Try
      Select Case strLabel
        Case "BEI", "HVI", "AXI", "MDI", "VBI"
          GetMood = "Imperative"
        Case "BEPH", "BEPI", "BEPS", "BEP", "BEDI", "BEDS", "BED", _
          "HVPI", "HVPS", "HVP", "HVDI", "HVDS", "HVD", _
          "AXPI", "AXPS", "AXP", "AXDI", "AXDS", "AXD", _
          "MDPI", "MDPS", "MDP", "MDDI", "MDDS", "MDD", _
          "VBPH", "VBPI", "VBPS", "VBP", "VBDI", "VBDS", "VBD", _
          "NEG+BEDI", "NEG+HVD", "NEG+MDD", _
          "RP+VBDI", "RP+VBPI", "RP+VBD"
          GetMood = "Declarative"
        Case Else
          GetMood = "Unknown"
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetMood error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return ""
    End Try
  End Function
  ' -------------------------------------------------------------------------------------------------------
  ' Name: GetCase
  ' Goal: Determine what kind of grammatical case this node has (if it is an NP)
  ' History:
  ' 08-04-2009    ERK Created
  ' -------------------------------------------------------------------------------------------------------
  Private Function GetCase(ByVal strLabel As String, Optional ByVal strDefault As String = "") As String
    Dim intPos As Integer         ' Position in string

    Try
      ' Is this an NP label
      If (Left(strLabel, 3) <> "NP-") Then
        ' Perhaps this is just a simple NP?
        If (strLabel = "NP") Then
          ' Return nominative case
          Return "NOM"
        Else
          ' Return failure
          Return ""
        End If
      End If
      ' Strip the beginning from the label
      strLabel = Mid(strLabel, 4)
      ' Is there any hyphen following?
      intPos = InStr(strLabel, "-")
      If (intPos > 0) Then
        ' Strip everything from hyphen onwards
        strLabel = Left(strLabel, intPos - 1)
      End If
      ' What if the result is empty?
      If (strLabel = "") Then
        strLabel = strDefault
      End If
      ' Return the result
      Return strLabel
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetCase error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return ""
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  GetNPhierarchy
  ' Goal :  Get the hierarchy-number depending on the NP type and the coreference type
  ' Note :  The higher the number, the more definite/specific/potentially referential
  ' History:
  ' 06-05-2009  ERK Created
  ' 11-10-2012  ERK Added the [bIsPP] test
  ' ----------------------------------------------------------------------------------------------------------
  Private Function GetNPhierarchy(ByVal strNPtype As String, ByVal strCorType As String, ByVal bIsPP As Boolean) As Integer
    Dim intAdd As Integer = 0   ' Amount to be added if this is NOT a PP object

    Try
      ' We can add the maximum (5) if this is not a PP object)
      If (Not bIsPP) Then intAdd = 5
      ' Check the NP type
      Select Case strNPtype
        Case "empty", "ZeroSbj"      ' For the moment DON'T count "trace" as empty
          GetNPhierarchy = 5 + intAdd
        Case "DemPro", "Dem"
          GetNPhierarchy = 4 + intAdd
        Case "PrsPro", "Pro", "PossPro"
          GetNPhierarchy = 3 + intAdd
        Case "DefNP"
          GetNPhierarchy = 2 + intAdd
        Case Else
          ' Check if it is a referring NP (which can be the case, for instance, for Proper nouns)
          If (DoLike(strCorType, "Identity|CrossSpeech|Assumed|Inferred")) Then
            GetNPhierarchy = 1 + intAdd
          Else
            GetNPhierarchy = 0 + intAdd
          End If
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetNPhierarchy error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return 0
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  GetSubjNP
  ' Goal :  Find out whether there is a unique subject NP
  ' Note :  For the moment we only look at the grammatical case. 
  '         Is there one unique NP with nominative case?
  ' History:
  ' 06-05-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Function GetSubjNP(ByRef arNPinfo() As NPinfo, ByVal intNPnum As Integer) As Integer
    Dim bNominative As Boolean = False  ' Whether we found a nominative case NP or not
    Dim intI As Integer                 ' Counter
    Dim intIdx As Integer               ' Index of nominative NP

    Try
      ' Loop through all NP's
      For intI = 1 To intNPnum
        ' Access this element
        With arNPinfo(intI)
          ' Is this one nominative?
          If (DoLike(.NPcase, "NOM|SBJ")) Then
            ' Did we already have a nominative one?
            If (bNominative) Then
              ' There are 2 nominative case NP's, so return failure
              Return 0
            Else
              ' Set the nominative flag
              bNominative = True
              ' Set the index
              intIdx = intI
            End If
          End If
        End With
      Next intI
      ' Check the net result
      If (bNominative) Then
        ' Unique nominative NP, return success
        GetSubjNP = intIdx
      Else
        ' No nominative NP found - return failure
        GetSubjNP = 0
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetSubjNP error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return 0
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  GetUniqueNP
  ' Goal :  Find out whether the array of NP's contains one that is at the top of the hierarchy
  '           empty > DemPro > Pro > Dem+NP > IdentityNP > other NP
  ' Return: - the index of the found unique NP, or 0 if not found
  '         - the Coreference Type of the unique NP found
  ' History:
  ' 06-05-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Function GetUniqueNP(ByRef arNPinfo() As NPinfo, ByVal intNPnum As Integer, _
    ByRef strCorType As String) As Integer
    Dim intWinner As Integer = -1       ' For the moment there is no winner
    Dim intI As Integer                 ' Counter
    Dim intIdx As Integer               ' Index of nominative NP
    Dim bUnique As Boolean = False      ' Whether this is a unique winner or not

    Try
      ' Loop through all NP's
      For intI = 1 To intNPnum
        ' Access this element
        With arNPinfo(intI)
          ' Is the hierarchy higher?
          If (.Hierarchy > intWinner) Then
            ' Adapt the winning number
            intWinner = .Hierarchy
            ' Also get the index of this one
            intIdx = intI
            ' And also get the coreference type of this one
            strCorType = .CorType
            ' Set uniqueness
            bUnique = True
          ElseIf (.Hierarchy = intWinner) Then
            ' Reset uniqueness
            bUnique = False
          End If
        End With
      Next intI
      ' Evaluate the answer
      If (bUnique) Then
        ' There is a unique winner, return it
        GetUniqueNP = intIdx
      Else
        ' There is no unique winner
        GetUniqueNP = 0
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetUniqueNP error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return 0
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  GetNPpos
  ' Goal :  Determine the position of the NP (intNodeId) within the clause (intclauseNode)
  ' History:
  ' 09-04-2009  ERK Created
  ' 11-10-2012  ERK Adapted for Cesax
  ' ----------------------------------------------------------------------------------------------------------
  Private Function GetNPpos(ByVal intNodeId As Integer, ByRef ndxParent As XmlNode) As String
    Dim ndxChild As XmlNode     ' One node
    Dim strPos As String = ""   ' The layout of this CP/IP/PP as a string
    Dim strAdd As String = ""   ' The symbol to be added

    Try
      ' Is there a valid clause node ID?
      If (ndxParent Is Nothing) Then
        ' Return failure
        Return "n/a"
      End If
      ' Don't look at PP stuff
      If (DoLike(ndxParent.Attributes("Label").Value, "PP*")) Then
        ' Return information
        Return "pp"
      End If
      ' Walk all the [eTree] children
      ndxChild = ndxParent.FirstChild
      While (Not ndxChild Is Nothing)
        ' This must be an [eTree]
        If (ndxChild.Name = "eTree") Then
          ' Check what we have here
          If (ndxChild.Attributes("Id").Value = intNodeId) Then
            ' This is ME
            strAdd = "N"
          ElseIf (DoLike(ndxChild.Attributes("Label").Value, strFiniteVerb)) Then
            ' This is the finite verb
            strAdd = "V"
          Else
            strAdd = "x"
          End If
          ' Avoid duplicate "x"
          If (Not ((strAdd = "x") AndAlso (Right(strPos, 1) = "x"))) Then
            ' Add to the total
            strPos &= strAdd
          End If

        End If
        ' Go to next child
        ndxChild = ndxChild.NextSibling
      End While
      ' Return the total string
      Return strPos
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetNPpos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   NodeIS
  ' Goal:   Produce HTML output that can be used for IS processing (topic guessing)
  ' History:
  ' 11-10-2012  ERK Adapted from Cesac2008
  ' ------------------------------------------------------------------------------------
  Private Function NodeIS(ByRef ndxThis As XmlNode) As String
    Dim intThisId As Integer          ' The ID of the node under consideration (may differ from previous one)
    Dim intNPsize As Integer          ' Size of the NP in number of elements
    Dim intCorDist As Integer         ' The distance to the coreferenced element
    Dim intNPnum As Integer           ' The number of the NP being looked at within an IP/CP
    Dim intTopIdx As Integer          ' index of the topic
    Dim intI As Integer               ' Counter
    Dim strIPloc As String = ""       ' Location of this IP
    Dim strText As String = ""        ' The text of this IP/CP
    Dim strBack As String = ""        ' What we return
    Dim strMood As String = ""        ' The mood of the IP/CP
    Dim strBracket As String = ""     ' The text of this IP, but now bracketed
    Dim strSrc As String = ""         ' Contents of source node
    Dim strDst As String = ""         ' Contents of destination node
    Dim strRoot As String = ""        ' Root of the chain
    Dim strCorType As String = ""     ' The type of the reference
    Dim strLblSrc As String = ""      ' Label of hte source
    Dim strLblDst As String = ""      ' Label of the destination
    Dim strClauseType As String = ""  ' Type of clause we are in
    Dim strCase As String = ""        ' Grammatical case (if determined)
    Dim strNPtype As String = ""      ' The kind of NP we are having
    Dim strNPpos As String = ""       ' The position of this NP within the PP/IP/CP
    Dim strIStype As String = ""      ' The Infor.Structure type: topic-comment, arg-focus, sent-focus etc.
    Dim strNPtext As String = ""      ' The text of this NP
    Dim strNPcat As String = ""       ' The syntactical category of the NP
    Dim strTopic As String = ""       ' The topic as a string
    Dim strReason As String = ""      ' The reason for choosing this particular IS type
    Dim strCorDist As String = ""     ' Coreference distance as string
    Dim bAddIt As Boolean = False     ' Whether information from this node needs to be added
    Dim bIsPP As Boolean = False      ' Whether the node actually is a PP
    Dim bHasDem As Boolean            ' Whether one of the NP's is an independant demonstrative
    Dim ndxChild As XmlNode           ' One child
    Dim ndxWork As XmlNode            ' Working node
    Dim ndxSel As XmlNodeList         ' List of selected nodes
    Dim arNPinfo(10) As NPinfo        ' Array of NP info elements (allow maximum of 10)

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      If (ndxThis.Name <> "eTree") Then Return ""
      ' (1) - Get the location code of his IP/CP
      strIPloc = ndxThis.SelectSingleNode("ancestor::forest").Attributes("Location").Value
      ' (2) - Get the text of this IP/CP
      strText = NodeText(ndxThis)
      ' (3) - Visit all the children of this IP/CP
      ndxSel = ndxThis.SelectNodes("child::eTree")
      strMood = "" : intNPnum = 0 : intTopIdx = 0 : strTopic = "-" : strReason = "unknown"
      bIsPP = False : bHasDem = False
      For intI = 0 To ndxSel.Count - 1
        ' Initialisations
        ndxWork = Nothing : bIsPP = False
        ' Access this child
        ndxChild = ndxSel(intI)
        ' Get the label and the id
        strLblSrc = ndxChild.Attributes("Label").Value : intThisId = ndxChild.Attributes("Id").Value
        ' Add this to the bracketed model of the line
        If (strBracket <> "") Then strBracket &= " "
        strBracket &= "[" & strLblSrc & " " & Trim(NodeText(ndxChild)) & "]"
        ' Do some preprocessing to get the NP of a PP if any
        If (DoLike(strLblSrc, "PP*")) Then
          ' Get the NP that is part of the PP
          ndxWork = ndxChild.SelectSingleNode("child::eTree[tb:Like(@Label, 'NP*')]", conTb)
          ' Got anything?
          If (Not ndxWork Is Nothing) Then
            ' There is an NP child: get its label
            strLblSrc = ndxWork.Attributes("Label").Value
            ' Indicate this is a PP
            bIsPP = True
          End If
        ElseIf (DoLike(strLblSrc, "NP*")) Then
          ' Get the NP
          ndxWork = ndxChild
        End If
        ' Is this the finite verb?
        If (DoLike(strLblSrc, strFiniteVerb)) Then
          ' Try determine the mood
          strMood = GetMood(strLblSrc)
        ElseIf (bIsPP) OrElse (DoLike(strLblSrc, "NP*")) Then
          ' THis is a referential element that needs processing
          ' Get the text of this NP
          strNPtext = NodeText(ndxWork)
          ' Determine the kind of NP this is from the NPtype feature
          strNPtype = GetFeature(ndxWork, "NP", "NPtype")
          ' Check for interrogative mood
          If (DoLike(strNPtype, "whPro|Wh*|wh*")) Then
            strMood = "Interrogative"
          End If
          ' Determine the size of the node
          intNPsize = ndxWork.SelectNodes("descendant::eLeaf[@Type='Vern']").Count
          ' Try to get the kind of case
          strCase = GetCase(strLblSrc, "n/a")
          ' Adapt case to "Oblique", if this is an NP part of a PP
          If (DoLike(strCase, "NOM*|SBJ*|OB*")) AndAlso (bIsPP) Then strCase = "OBL"
          ' Determine the reference type and other general information
          strCorType = GetFeature(ndxWork, "coref", "RefType")
          ' Get the coreference distance
          strCorDist = GetFeature(ndxWork, "coref", "NdDist")
          If (strCorDist = "") Then intCorDist = 0 Else intCorDist = CInt(strCorDist)
          ' Get the syntactical category of the NP
          If (bIsPP) Then
            strNPcat = "PP-OBJ"
          Else
            strNPcat = ndxWork.Attributes("Label").Value
          End If
          ' Add this NP to the stack
          intNPnum += 1
          ' =========== DEBUG ===========
          ' If (InStr(strNPtext, "Wenlo") > 0) Then Stop
          ' =============================
          ' Store the information in the array
          With arNPinfo(intNPnum)
            .NodeId = intThisId
            .NPtext = strNPtext
            .NPtype = strNPtype
            .NPcat = strNPcat
            .Size = intNPsize
            .NPcase = strCase
            .CorType = strCorType
            .CorDist = intCorDist
            .Hierarchy = GetNPhierarchy(strNPtype, strCorType, bIsPP)
            .IsPP = bIsPP
          End With
        End If
      Next intI
      ' Determine what kind of Information Structure type this clause is
      '  NOTE: for the moment only topic-comment or something else (=other)
      ' Step #1: action depends on the mood
      If (strMood = "Declarative") Then
        ' Step #2: Determine the number of NPs
        Select Case intNPnum
          Case 0    ' This is not a topic-comment clause
            strIStype = "other"
            strReason = "no NP"
          Case 1    ' This could be a topic-comment clause
            ' Is this a referring NP?
            If (DoLike(strCorType, "Identity|CrossSpeech|Assumed|Inferred")) Then
              ' This is a topic-comment clause
              strIStype = "TopCom"
              intTopIdx = 1
              strTopic = arNPinfo(intTopIdx).NPtext
              strNPcat = arNPinfo(intTopIdx).NPcat
              strReason = "1NP=ref"
              ' Further differentiate IS-type
              If (arNPinfo(intTopIdx).IsPP) Then strIStype &= "-PP"
              ' Also take into account demonstratives...
              If (DoLike(arNPinfo(intTopIdx).NPtype, "DemPro|Dem")) Then strIStype &= "-Dem"
            Else
              ' This must be argument-focus or sentence-focus
              strIStype = "other"
              strReason = "1NP=noref"
            End If
          Case Else ' Ask the next question
            ' Step #3: is there a unique NP above the hierarchy?
            intI = GetUniqueNP(arNPinfo, intNPnum, strCorType)
            If (intI > 0) Then
              ' There is a unique NP on top. But is it of the right referential type?
              If (DoLike(strCorType, "Identity|CrossSpeech|Assumed|Inferred")) Then
                ' Yes, this is a unique NP, which refers back to something
                strIStype = "TopCom"
                intTopIdx = intI
                strTopic = arNPinfo(intTopIdx).NPtext
                strNPcat = arNPinfo(intTopIdx).NPcat
                strReason = "unqNP=ref"
                ' Further differentiate IS-type
                If (arNPinfo(intTopIdx).IsPP) Then strIStype &= "-PP"
                ' Also take into account demonstratives...
                If (arNPinfo(intTopIdx).NPtype = "DemPro") Then strIStype &= "-Dem"
              Else
                ' The NP is unique, but it does not refer back...
                strIStype = "other"
                strReason = "unqNP=noref"
              End If
            Else
              ' There is no unique NP type at the top. so look further!!
              ' Step #4: look at subjecthood
              intI = GetSubjNP(arNPinfo, intNPnum)
              If (intI > 0) Then
                ' There is a unique Subject NP
                strIStype = "TopCom"
                intTopIdx = intI
                strTopic = arNPinfo(intTopIdx).NPtext
                strNPcat = arNPinfo(intTopIdx).NPcat
                strReason = "NP=Su"
                ' Further differentiate IS-type
                If (arNPinfo(intTopIdx).IsPP) Then strIStype &= "-PP"
                ' Also take into account demonstratives...
                If (arNPinfo(intTopIdx).NPtype = "DemPro") Then strIStype &= "-Dem"
              Else
                ' There is not a unique Subject NP
                ' Step #5: look at animacy -- TODO: implement animacy
                strIStype = "other"
                strReason = "NP=noSu"
              End If
            End If
        End Select
      Else
        ' This is not topic-comment
        strIStype = "other"
        strReason = "mood"
      End If
      ' Has a topic NP been found?
      If (intTopIdx > 0) Then
        ' Determine the "position" of the topic within this IP/CP
        strNPpos = GetNPpos(arNPinfo(intTopIdx).NodeId, ndxThis)
      Else
        ' No position can be determined, because no Topic is there
        strNPpos = "-"
      End If
      ' Adapt topic if it is empty
      If (strTopic = "") Then strTopic = "(elided)"
      ' Provide HTML output: Loc - Type - Reason - Topic - Syntax
      strBack = "<tr><td><Font Size=" & """" & "1" & """" & ">" & strIPloc & "</Font></td>" & _
        "<td>" & strIStype & "</td>" & _
        "<td>" & strReason & "</td>" & _
        "<td>" & strNPpos & "</td>" & _
        "<td>" & strNPcat & "</td>" & _
        "<td>" & strTopic & "</td>" & _
        "<td>" & strBracket & "</td>" & _
        "</tr>" & vbCrLf
      ' Return what we have found
      Return strBack
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/NodeIS error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   TravNode
  ' Goal:   Travel a node recursively, fulfilling the indicated operation [strOp]
  '         Return the result in [strBack], which you have to make empty before starting!!
  ' History:
  ' 10-05-2010  ERK Created
  ' 11-10-2012  ERK Added topic guessing (from Cesac2008)
  ' 24-04-2013  ERK Added "DepChe" format
  ' ------------------------------------------------------------------------------------
  Public Function TravNode(ByRef ndxThis As XmlNode, ByVal strOp As String, ByRef strBack As String, _
                           Optional ByVal bHtml As Boolean = False) As Boolean
    Dim strLabel As String = ""         ' The label of this node
    Dim strParentL As String = ""       ' Label of parent
    Dim strOrg As String = ""           ' Original text
    Dim strPde As String = ""           ' PDE translation
    Dim strState As String = ""         ' A particular cognitive state
    Dim strCat As String = ""           ' category
    Dim strAnt As String = ""           ' Antecedent ID
    Dim strDepRel As String = ""        ' Dependency relation type
    Dim strFeats As String = ""         ' Features belonging to this node
    'Dim intEtreeId As Integer           ' ID of current  node
    'Dim intDepId As Integer             ' Dependency ID
    'Dim intDepHd As Integer             ' Dependency head ID
    'Dim intI As Integer                 ' COunter
    Dim ndxChild As XmlNode = Nothing   ' Children of me
    Dim ndxParent As XmlNode = Nothing  ' My parent
    Dim ndxHead As XmlNode = Nothing    ' Head node
    'Dim ndxList As XmlNodeList          ' List of children
    Dim bIsFsEleaf As Boolean = False   ' Whether this <eTree> contains <fs> elements and also an <eLeaf> child

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Perform initial action
      Select Case strOp
        Case "Rating"  ' Interrater agreement comparison
          Select Case ndxThis.Name
            Case "forest"
              ' Walk all the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Apply action on this child
                If (Not TravNode(ndxChild, strOp, strBack)) Then
                  ' Error, so retreat
                  Return False
                End If
              Next ndxChild
            Case "eTree"    ' Only note interrater agreement of <eTree> elements
              ' Should this node be included?
              If (DoLike(ndxThis.Attributes("Label").Value, "NP|NP-*|PRO$|NPR$")) Then
                ' Get the RefType and the antecedent ID
                strState = CorefAttr(ndxThis, "RefType") : strState = IIf(strState = "", "(empty)", strState)
                strAnt = CorefDstId(ndxThis) : strAnt = IIf(strAnt = "-1", "0", strAnt)
                ' Store the result
                strBack &= ndxThis.Attributes("Id").Value & ";" & strAnt & ";" & strState & vbCrLf
              End If
              ' Process the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Try to process the child
                If (Not (TravNode(ndxChild, strOp, strBack))) Then Return False
              Next ndxChild
          End Select
        Case "DepChe", "DepEng", "DepNl" ' Dependancy formats for Chechen, English etc
          ' Bottom-Up processing, where the heads and dependencies are resolved
          Select Case ndxThis.Name
            Case "forest"
              ' The forest's children are all <root>s
              For Each ndxChild In ndxThis.SelectNodes("./child::eTree")
                ' Forest-children are roots
                AddFeature(pdxCurrentFile, ndxChild, "dep", "rel", "root")
                ' Try to process the child
                If (Not (TravNode(ndxChild, strOp, strBack))) Then Return False
              Next ndxChild
            Case "eTree"
              ' ============= DEBUG ==============
              ' Debug.Print("TravNode Label=" & ndxThis.Attributes("Label").Value)
              'If (ndxThis.SelectSingleNode("./ancestor::forest[@forestId='4']") IsNot Nothing) Then
              '  Stop
              'End If
              If (ndxThis.Attributes("Id").Value = 80) OrElse (ndxThis.Attributes("Id").Value = 143) Then Stop
              ' ==================================
              ' Check if this node must have a head:
              '    Are there any non-punctuation <eTree> children?
              ' Debug.Print(ndxThis.SelectSingleNode("./child::eTree[count(child::eLeaf[@Type='Punct'])>0]").ToString)
              'If (ndxThis.SelectSingleNode("./child::eTree[count(child::eLeaf)=0" & _
              '                             " or count(child::eLeaf[@Type!='Punct' and @Type!='Zero'])>0]") IsNot Nothing) Then
              ' OLD: If (ndxThis.SelectSingleNode("./child::eTree[count(child::eLeaf)=0 or count(child::eLeaf[@Type!='Punct'])>0]") IsNot Nothing) Then
              ' (1) Determine the head for this phrase
              If (ndxThis.SelectSingleNode("./child::eTree[tb:hdcandi(self::eTree)]", conTb) IsNot Nothing) Then
                ndxHead = GetHeadNodeDep(ndxThis, strOp)
                If (ndxHead Is Nothing) Then
                  ' We're in trouble -- no head!!!
                  bInterrupt = True
                  Stop
                  Return False
                End If
                ' (2) Mark it as head
                AddFeature(pdxCurrentFile, ndxHead, "dep", "rel", "hd")
                ' (3) Visit all children that are NOT the head
                For Each ndxChild In ndxThis.SelectNodes("./child::eTree[not(@Id=" & ndxHead.Attributes("Id").Value & ")]")
                  ' (4) Add the dependency relation of this child
                  AddFeature(pdxCurrentFile, ndxChild, "dep", "rel", GetDepRel(ndxChild, strOp))
                Next ndxChild
                ' Process the children, unless we are at the bottom!
                For Each ndxChild In ndxThis.SelectNodes("./child::eTree[count(child::eLeaf)=0]")
                  ' Try to process the child
                  If (Not (TravNode(ndxChild, strOp, strBack))) Then Return False
                Next ndxChild
              ElseIf (ndxThis.SelectSingleNode("./child::eLeaf[(@Type='Star')" & _
                      " and (@Text='*exp*' or " & _
                      "tb:matches(@Text, '*-[123456789]|*-[123456789][0123456789]'))]", conTb) IsNot Nothing) Then
                ' This is a node that points to a dislocated node, so this one IS the head
                AddFeature(pdxCurrentFile, ndxThis, "dep", "rel", "hd")
              Else
                ' Double check
                If (GetFeature(ndxThis, "dep", "rel").ToString = "hd") Then
                  Stop
                  ' We should not modify nodes that are already assigned to be heads
                Else
                  ' Check if all is going well...
                  Stop
                  ' This is an end-node, and an end-node can never be a head, but only a modifier or something
                  AddFeature(pdxCurrentFile, ndxThis, "dep", "rel", GetDepRel(ndxThis, strOp))
                End If
              End If

              ' ============== OLD CODE THAT WORKED OKAY =====================
              '' Check if this node must have a head
              'If (ndxThis.SelectSingleNode("./child::eTree[count(child::eLeaf[@Type!='Punct'])>0]") IsNot Nothing) Then
              '  ' (1) Determine the head for this phrase
              '  ndxHead = GetHeadNodeDep(ndxThis, strOp)
              '  If (ndxHead Is Nothing) Then
              '    ' We're in trouble -- no head!!!
              '    bInterrupt = True
              '    Return False
              '    Stop
              '  End If
              '  ' (2) Mark it as head
              '  AddFeature(pdxCurrentFile, ndxHead, "dep", "rel", "hd")
              '  ' (3) Visit all children that are NOT the head
              '  For Each ndxChild In ndxThis.SelectNodes("./child::eTree[not(@Id=" & ndxHead.Attributes("Id").Value & ")]")
              '    ' (4) Add the dependency relation of this child
              '    AddFeature(pdxCurrentFile, ndxChild, "dep", "rel", GetDepRel(ndxChild, strOp))
              '  Next ndxChild
              '  ' Process the children, unless we are at the bottom!
              '  For Each ndxChild In ndxThis.SelectNodes("./child::eTree[count(child::eLeaf)=0]")
              '    ' Try to process the child
              '    If (Not (TravNode(ndxChild, strOp, strBack))) Then Return False
              '  Next ndxChild
              'Else
              '  ' This is an end-node, and an end-node can never be a head, but only a modifier or something
              '  AddFeature(pdxCurrentFile, ndxThis, "dep", "rel", GetDepRel(ndxThis, strOp))
              'End If
              ' ==============================================================
          End Select
        Case "PosPTB"      ' Penn Tree Bank part-of-speech words per line
          Select Case ndxThis.Name
            Case "forest"
              ' Walk all the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Apply action on this child
                If (Not TravNode(ndxChild, strOp, strBack)) Then
                  ' Error, so retreat
                  Return False
                End If
              Next ndxChild
              ' Add a new line after this forest has finished
              strBack &= vbLf
            Case "eTree"
              ' Process the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Try to process the child
                If (Not (TravNode(ndxChild, strOp, strBack))) Then Return False
              Next ndxChild
            Case "eLeaf"
              ' Check if this is not an empty category
              If (ndxThis.Attributes("Type").Value <> "Star") Then
                ' Get label of parent
                strLabel = ndxThis.ParentNode.Attributes("Label").Value
                ' Replace label's semicolon with underscore
                strLabel = strLabel.Replace(";", "_")
                ' Add vernacular + POS tag
                strBack &= ndxThis.Attributes("Text").Value & "/" & strLabel & " "
              End If
          End Select
        Case "Token"      ' Tokenization, as close to [PosSimple] as possible
          Select Case ndxThis.Name
            Case "forest"
              ' Walk all the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Apply action on this child
                If (Not TravNode(ndxChild, strOp, strBack)) Then
                  ' Error, so retreat
                  Return False
                End If
              Next ndxChild
              ' Add a <utt> line after this forest has finished
              strBack &= "<utt>" & vbLf
            Case "eTree"
              ' Process the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Try to process the child
                If (Not (TravNode(ndxChild, strOp, strBack))) Then Return False
              Next ndxChild
            Case "eLeaf"
              ' Check if this is not an empty category
              If (ndxThis.Attributes("Type").Value <> "Star") Then
                ' Get label of parent
                strLabel = ndxThis.ParentNode.Attributes("Label").Value
                ' Replace label's semicolon with underscore
                strLabel = strLabel.Replace(";", "_")
                ' Add vernacular + POS tag
                strBack &= ndxThis.Attributes("Text").Value & vbLf
              End If
          End Select
        Case "PosSimple"  ' Simple POS tags outputting
          Select Case ndxThis.Name
            Case "forest"
              ' Walk all the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Apply action on this child
                If (Not TravNode(ndxChild, strOp, strBack)) Then
                  ' Error, so retreat
                  Return False
                End If
              Next ndxChild
              ' Add a <utt> line after this forest has finished
              strBack &= "<utt>" & vbLf
            Case "eTree"
              ' Process the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Try to process the child
                If (Not (TravNode(ndxChild, strOp, strBack))) Then Return False
              Next ndxChild
            Case "eLeaf"
              ' Check if this is not an empty category
              If (ndxThis.Attributes("Type").Value <> "Star") Then
                ' Get label of parent
                strLabel = ndxThis.ParentNode.Attributes("Label").Value
                ' Replace label's semicolon with underscore
                strLabel = strLabel.Replace(";", "_")
                ' Add vernacular + POS tag
                strBack &= ndxThis.Attributes("Text").Value & vbTab & strLabel & vbLf
              End If
          End Select
        Case "PosOrg"  ' Simple POS tags outputting
          Select Case ndxThis.Name
            Case "forest"
              ' Walk all the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Apply action on this child
                If (Not TravNode(ndxChild, strOp, strBack)) Then
                  ' Error, so retreat
                  Return False
                End If
              Next ndxChild
              '' Add a newline
              'strBack &= vbLf
            Case "eTree"
              ' Process the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Try to process the child
                If (Not (TravNode(ndxChild, strOp, strBack))) Then Return False
              Next ndxChild
            Case "eLeaf"
              ' Check if this is not an empty category
              If (ndxThis.Attributes("Type").Value <> "Star") Then
                ' Get label of parent
                ndxParent = ndxThis.ParentNode
                If (ndxParent Is Nothing) Then Return False
                strLabel = ndxParent.Attributes("Label").Value.Replace(";", "_")
                ' Do we have Cat and Cyr values?
                If (ndxParent.SelectSingleNode("./child::fs/child::f[@name='Cat']") Is Nothing) Then
                  ' Do it simple
                  strCat = strLabel : strOrg = ndxThis.Attributes("Text").Value
                Else
                  ' Get category and cyr values
                  strCat = ndxParent.SelectSingleNode("./child::fs/child::f[@name='Cat']").Attributes("value").Value
                  strOrg = ndxParent.SelectSingleNode("./child::fs/child::f[@name='Cyr']").Attributes("value").Value
                End If
                ' Add vernacular + POS tag
                strBack &= strOrg & vbTab & strCat & vbTab & ndxThis.Attributes("Text").Value & vbTab & strLabel & vbLf
              End If
          End Select
        Case "Syntax" ' Determine the syntax
          ' Is this an <eTree> or an <eLeaf> node?
          Select Case ndxThis.Name
            Case "eTree"
              ' Filter out unwanted nodes
              strLabel = ndxThis.Attributes("Label").Value
              ' Don't continue to process CODE nodes
              If (strLabel Like "CODE*") Then Return True
              ' Do start the output of other nodes
              If (bHtml) Then
                strBack &= "[<font size='1' color='blue'>" & strLabel & " </font>"
              Else
                strBack &= "[" & strLabel & " "
              End If
              ' Process the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Try to process the child
                If (Not (TravNode(ndxChild, strOp, strBack, bHtml))) Then Return False
              Next ndxChild
            Case "eLeaf"
              ' Return the leaf value, if this is applicable
              strBack &= ndxThis.Attributes("Text").Value
          End Select
        Case "CgnState"  ' Output for cognitive status part
          ' Is this an <eTree> or an <eLeaf> node?
          Select Case ndxThis.Name
            Case "forest"
              ' Walk all the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Apply action on this child
                If (Not TravNode(ndxChild, strOp, strBack)) Then
                  ' Error, so retreat
                  Return False
                End If
              Next ndxChild
            Case "eTree"
              ' Filter out unwanted nodes
              strLabel = ndxThis.Attributes("Label").Value
              ' Don't continue to process CODE nodes
              If (strLabel Like "CODE*") Then Return True
              ' Skip some other nodes too
              If (DoLike(GetFeature(ndxThis, "NP", "NPtype"), "Trace|Zero")) Then Return True
              ' Check if this is a node with a particular cognitive state
              strState = GetFeature(ndxThis, "NP", "CgState")
              If (strState <> "") Then
                ' Do start the output of other nodes
                strBack &= "[<font size='1' color='blue'>" & strLabel & " </font>"
                ' Add the state in a particular color
                Select Case strState
                  Case "inf"
                    strBack &= "<font size='1'>InFoc</font> <font color='red'>"
                  Case "act"
                    strBack &= "<font size='1'>Act</font> <font color='blue'> "
                  Case "fam"
                    strBack &= "<font size='1'>Fam</font> <font color='pink'> "
                  Case "unq"
                    strBack &= "<font size='1'>UnqId</font> <font color='green'> "
                  Case "ref"
                    strBack &= "<font size='1'>Ref</font> <font color='orange'> "
                  Case "typ"
                    strBack &= "<font size='1'>TpId</font> <font color='blueviolet'> "
                  Case Else
                    strBack &= "<font size='1'>non</font> <font color='brown'> "
                End Select
              End If
              ' Process the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Try to process the child
                If (Not (TravNode(ndxChild, strOp, strBack, bHtml))) Then Return False
              Next ndxChild
              ' Finish this one
              If (strState <> "") Then strBack &= "</font>]"
            Case "eLeaf"
              ' What is the type of eleaf?
              Select Case ndxThis.Attributes("Type").Value
                Case "Zero"
                  ' skip this
                Case "Vern"
                  ' Return the leaf value, if this is applicable
                  strBack &= VernToEnglish(ndxThis.Attributes("Text").Value) & " "
                Case "Punct"
                  ' Return the leaf value, if this is applicable
                  strBack &= VernToEnglish(ndxThis.Attributes("Text").Value) & " "
                Case Else
                  ' Return the leaf value, if this is applicable
                  strBack &= VernToEnglish(ndxThis.Attributes("Text").Value) & " "
              End Select
          End Select
        Case "ISoutput"
          ' Check the kind of node we have: IS output needs an [eTree
          Select Case ndxThis.Name
            Case "eTree"
              ' Only look at main clauses and sub clauses
              If (DoLike(ndxThis.Attributes("Label").Value, "IP-MAT*|IP-SUB*")) Then
                ' Add the IS output of this clause
                strBack &= NodeIS(ndxThis)
              End If
            Case "forest"
              ' Walk all the children
              For Each ndxChild In ndxThis.ChildNodes
                ' Apply action on this child
                If (Not TravNode(ndxChild, strOp, strBack)) Then
                  ' Error, so retreat
                  Return False
                End If
              Next ndxChild
          End Select
        Case "PdeOutput"
          ' We only need to consider the child(ren) of a <forest> node
          If (ndxThis.Name = "forest") Then
            ' Get the org and english translation
            strOrg = XmlString(ndxThis.SelectSingleNode("./div[@lang='org']").InnerText)
            strPde = XmlString(ndxThis.SelectSingleNode("./div[@lang='eng']").InnerText)
            ' Determine the output
            strBack &= "<forest forestId='" & ndxThis.Attributes("forestId").Value & "'" & _
              " location='" & ndxThis.Attributes("Location").Value & _
              "'>" & vbCrLf & _
              "  <div lang='org'><seg>" & strOrg & "</seg></div>" & vbCrLf & _
              "  <div lang='eng'>" & vbCrLf & "    <seg>" & strPde & "</seg>" & _
              vbCrLf & "  </div>" & vbCrLf & "</forest>"
          End If
        Case "Tiger"
          Select Case ndxThis.Name
            Case "forest"
              ' Walk all <eTree> children
              loc_intLevelSyntax = 1
              For Each ndxChild In ndxThis.ChildNodes
                ' Apply action on this child
                If (Not TravNode(ndxChild, strOp, strBack)) Then
                  ' Error, so retreat
                  Return False
                End If
              Next ndxChild
              ' Go one level down again
              loc_intLevelSyntax = 1
            Case "eTree"
              strLabel = gettigeridnumber(GetFeature(ndxThis, "tig", "id"))
              strBack &= StrDup((loc_intLevelSyntax - 1) * 2, " ") & strLabel & vbCrLf
              ' Continue with the children, which are 1 level more down
              loc_intLevelSyntax += 1
              For Each ndxChild In ndxThis.ChildNodes
                ' Apply action on this child
                If (Not TravNode(ndxChild, strOp, strBack)) Then
                  ' Error, so retreat
                  Return False
                End If
              Next ndxChild
              ' Go one level down again
              loc_intLevelSyntax -= 1
          End Select
        Case "PsdOutput"
          ' Action depends on the kind of node
          Select Case ndxThis.Name
            Case "eLeaf"
              ' Append the value of the leaf
              strBack &= ndxThis.Attributes("Text").Value
            Case "forest"
              ' We need to start with a left opening bracket
              strBack &= "("
              ' Walk all the children, which are 1 level deeper
              loc_intLevelSyntax = 1
              For Each ndxChild In ndxThis.ChildNodes
                ' Apply action on this child
                If (Not TravNode(ndxChild, strOp, strBack)) Then
                  ' Error, so retreat
                  Return False
                End If
              Next ndxChild
              ' Go one level down again
              loc_intLevelSyntax = 1
              ' Does this forest HAVE an ID node?
              If (Not ndxThis.Attributes("Location") Is Nothing) Then
                ' Do add an ID node
                strBack &= " (ID " & ndxThis.Attributes("TextId").Value & "," & _
                  ndxThis.Attributes("Location").Value & ")" & vbCrLf
              End If
              ' Finish this forest appropriately
              strBack &= ")" & vbCrLf
            Case "eTree"
              ' Action depends on what kind of label we have
              strLabel = ndxThis.Attributes("Label").Value
              ' Check if we are IP or CP, or if our parents are
              If (strLabel Like "[IC]P-*") Then
                ' Add an LF and some spaces
                strBack &= vbCrLf & Strings.StrDup(loc_intLevelSyntax, " ")
              Else
                ' Try get parent
                ndxParent = ndxThis.ParentNode
                If (Not ndxParent Is Nothing) Then
                  ' Only check <eTree> parents
                  If (ndxParent.Name = "eTree") Then
                    ' Get parent label name
                    strParentL = ndxParent.Attributes("Label").Value
                    ' Check parent label
                    If (strParentL Like "[IC]P-*") Then
                      ' Add an LF and some spaces
                      strBack &= vbCrLf & Strings.StrDup(loc_intLevelSyntax, " ")
                    End If
                  End If
                End If
              End If
              ' First give the bracket and the node's label with a space
              strBack &= "(" & strLabel & " "
              ' Possibly process <fs> content
              If (Not ndxThis.SelectSingleNode("./fs") Is Nothing) Then
                ' Process all <fs/f> children
                For Each ndxChild In ndxThis.SelectNodes("./fs/f")
                  ' Make sure only the correct attributes are saved
                  Select Case ndxChild.Attributes("name").Value
                    Case "history"
                      ' Don't include this one
                    Case Else
                      ' Add this <fs/f> child
                      strBack &= "(FS-" & ndxChild.Attributes("name").Value & " " & _
                        ndxChild.Attributes("value").Value & ")"
                  End Select
                Next ndxChild
                ' This node has <fs> children - check if it is an <eLeaf> parent or an <eTree> one
                If (Not ndxThis.SelectSingleNode("./eLeaf") Is Nothing) Then
                  ' This is an <eLeaf> parent, so add extra structure
                  strBack &= "(LEX "
                  bIsFsEleaf = True
                End If
              End If
              ' Continue with the children, which are 1 level more down
              loc_intLevelSyntax += 1
              For Each ndxChild In ndxThis.ChildNodes
                ' Apply action on this child
                If (Not TravNode(ndxChild, strOp, strBack)) Then
                  ' Error, so retreat
                  Return False
                End If
              Next ndxChild
              ' Go one level down again
              loc_intLevelSyntax -= 1
              ' Append a closing bracket
              strBack &= ")"
              ' Possibly append one more closing bracket
              If (bIsFsEleaf) Then
                strBack &= ")"
                bIsFsEleaf = False
              End If
          End Select
        Case "PsdSimple", "PsdPTB"
          ' Action depends on the kind of node
          Select Case ndxThis.Name
            Case "eLeaf"
              ' Append the value of the leaf
              strBack &= ndxThis.Attributes("Text").Value
            Case "forest"
              ' We need to start with a left opening bracket
              strBack &= "("
              ' PTB output requires root label to have "ROOT"
              If (strOp = "PsdPTB") Then strBack &= "ROOT "
              ' Walk all the children, which are 1 level deeper
              loc_intLevelSyntax = 1
              For Each ndxChild In ndxThis.ChildNodes
                ' Apply action on this child
                If (Not TravNode(ndxChild, strOp, strBack)) Then
                  ' Error, so retreat
                  Return False
                End If
              Next ndxChild
              ' Go one level down again
              loc_intLevelSyntax = 1
              ' Does this forest HAVE an ID node?
              If (strOp <> "PsdPTB") AndAlso (Not ndxThis.Attributes("Location") Is Nothing) Then
                ' Do add an ID node
                strBack &= " (ID " & ndxThis.Attributes("TextId").Value & "," & _
                  ndxThis.Attributes("Location").Value & ")" & vbCrLf
              End If
              ' Finish this forest appropriately
              strBack &= ")" & vbCrLf
            Case "eTree"
              ' Action depends on what kind of label we have
              strLabel = ndxThis.Attributes("Label").Value
              ' Check if we are IP or CP, or if our parents are
              If (strLabel Like "[IC]P-*") Then
                ' Add an LF and some spaces
                strBack &= vbCrLf & Strings.StrDup(loc_intLevelSyntax, " ")
              Else
                ' Try get parent
                ndxParent = ndxThis.ParentNode
                If (Not ndxParent Is Nothing) Then
                  ' Only check <eTree> parents
                  If (ndxParent.Name = "eTree") Then
                    ' Get parent label name
                    strParentL = ndxParent.Attributes("Label").Value
                    ' Check parent label
                    If (strParentL Like "[IC]P-*") Then
                      ' Add an LF and some spaces
                      strBack &= vbCrLf & Strings.StrDup(loc_intLevelSyntax, " ")
                    End If
                  End If
                End If
              End If
              ' First give the bracket and the node's label with a space
              strBack &= "(" & strLabel & " "
              ' Continue with the children, which are 1 level more down
              loc_intLevelSyntax += 1
              For Each ndxChild In ndxThis.ChildNodes
                ' Apply action on this child
                If (Not TravNode(ndxChild, strOp, strBack)) Then
                  ' Error, so retreat
                  Return False
                End If
              Next ndxChild
              ' Go one level down again
              loc_intLevelSyntax -= 1
              ' Append a closing bracket
              strBack &= ")"
          End Select
      End Select
      ' Perform finishing actions
      Select Case strOp
        Case "Syntax"
          ' If this is an <eTree>, then close this node
          If (ndxThis.Name = "eTree") Then strBack &= "]"
      End Select
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/TravNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   FirstMust
  ' Goal:   Find the first constituent that MUST receive coreference information
  ' History:
  ' 29-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function FirstMust(ByRef rtfThis As RichTextBox) As Boolean
    Dim intLine As Integer      ' The line in the array of XML nodes
    Dim intPtc As Integer       ' Percentage of progress

    Try
      ' Check whether busy
      If (bBusy) Then Return False
      ' Set busy
      bBusy = True
      ' Walk through all the <forest> nodes
      For intLine = 0 To intSectSize - 1
        ' Show where we are
        intPtc = (100 * intLine) \ intSectSize
        Status("Checking must " & intPtc & "%", intPtc)
        ' Is there something here?
        If (GetMust(rtfThis, intLine, True)) Then
          ' Reset busy
          bBusy = False
          ' Return success
          Return True
        End If
      Next intLine
      ' Everything is done
      Status("No later constituent MUST receive coreference information.")
      ' Reset busy
      bBusy = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/FirstMust error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------
  ' Name:   FindSource
  ' Goal:   Find and goto the constituent that points to me
  ' History:
  ' 23-11-2010  ERK Created
  ' 08-12-2012  ERK Need to make sure that the actual source, and not only the left hand is selected
  ' ------------------------------------------------------------------------------------------------
  Public Sub FindSource(ByRef ndxThis As XmlNode, Optional ByVal strLabel As String = "", _
                        Optional ByVal bIdentity As Boolean = True)
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intId As Integer          ' The Id we are looking for
    Dim ndxForest As XmlNode      ' The current forest being looked at
    Dim ndxNext As XmlNode        ' Working node
    Dim ndxResult As XmlNodeList  ' Result of selection
    Dim strRefType As String      ' The reference type we are dealing with
    Dim strIdtRef As String = "Identity|CrossSpeech"

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Exit Sub
      ' Get my Id
      intId = ndxThis.Attributes("Id").Value
      ' Show we are busy
      Status("Looking for a source node...")
      ' Start looking from this point in [ar]
      For intI = loc_intLineNum To intSectSize
        ' Get this forest
        ndxForest = pdxList(intSectFirst + intI)
        ' Any results at all?
        If (Not ndxForest Is Nothing) Then
          ' Check if the reference is in this forestn
          ndxResult = ndxForest.SelectNodes(".//ref[@target=" & intId & "]")
          ' Any result?
          If (ndxResult.Count > 0) Then
            ' Do we need to take care of the reftype?
            If (strLabel = "") Then
              ' Get to the first result's parent
              ndxNext = ndxResult(0).ParentNode
              ' This must be an [eTree]
              If (ndxNext.Name = "eTree") Then
                ' We have found it, so now let's try and select it
                GotoNode(ndxNext)
                ' Show we found it
                Status("Found source node")
                Exit Sub
              End If
            Else
              ' We need to watch for the correct reftype
              For intJ = 0 To ndxResult.Count - 1
                ' Get this node
                ndxNext = ndxResult(intJ).ParentNode
                ' Check if it is okay
                If (ndxNext.Name = "eTree") Then
                  ' Check if the label is okay
                  If (Not DoLike(ndxNext.Attributes("Label").Value, strLabel)) Then
                    ' Get the reference type
                    strRefType = GetFeature(ndxNext, "coref", "RefType")
                    ' Check this condition
                    If (bIdentity AndAlso (DoLike(strRefType, strIdtRef))) OrElse _
                       (Not bIdentity AndAlso (Not DoLike(strRefType, strIdtRef))) Then
                      ' We have found it, so now let's try and select it
                      GotoNode(ndxNext)
                      ' Show we found it
                      Status("Found source node")
                      Exit Sub
                    End If
                  End If
                End If
              Next intJ
            End If
          End If
        End If
      Next intI
      ' Failure - no selection
      Status("Could not find a source node.")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/FindSource error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------
  ' Name:   FindAntecedent
  ' Goal:   Find and goto the antecedent I point to
  ' History:
  ' 08-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------
  Public Sub FindAntecedent(ByRef ndxThis As XmlNode)
    Dim intId As Integer          ' The Id we are looking for
    Dim ndxNext As XmlNode        ' Working node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Exit Sub
      ' See if this has an antecedent
      ndxNext = ndxThis.SelectSingleNode("./child::ref")
      If (Not ndxNext Is Nothing) Then
        ' Get the target id
        intId = ndxNext.Attributes("target").Value
        ' Validate
        If (intId >= 0) Then
          ' Find the node belonging to this
          ndxNext = IdToNode(intId)
          ' Check result
          If (Not ndxNext Is Nothing) Then
            ' Go to this node
            GotoNode(ndxNext)
            ' Show we found it
            Status("Found antecedent")
            Exit Sub
          End If
        End If
      End If
      ' We have not found an antecedent
      Status("Could not find an antecedent.")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/FindAntecedent error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   GetEdtLine
  ' Goal:   Get the line number within the editor's textbox
  ' History:
  ' 14-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetEdtLine() As Integer
    Try
      ' Validate
      If (loc_intLineNum < 0) Then
        Return -1
      End If
      ' Just return the local line number + intsectfirst
      Return loc_intLineNum + intSectFirst
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/GetEdtLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetForest
  ' Goal:   Find the appropriate <forest> XML node
  ' History:
  ' 14-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetForest(ByVal intLine As Integer) As XmlNode
    Try
      ' First validation
      If (pdxList Is Nothing) Then Return Nothing
      ' Validate
      If (intLine < 0) OrElse (intSectFirst + intLine >= pdxList.Count) Then
        ' Return failure
        Return Nothing
      End If
      ' Return the appropriate XML node
      Return pdxList(intSectFirst + intLine)
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetForest error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  Public Function GetForestWithId(ByVal intForestId As Integer) As XmlNode
    Dim intI As Integer ' Counter

    Try
      ' First validation
      If (pdxList Is Nothing) Then Logging("GetForestWithId: empty pdxlist") : Return Nothing
      If (intForestId < 0) Then Logging("GetForestWithId: negative forestid") : Return Nothing
      ' Visit all lines
      For intI = 0 To pdxList.Count - 1
        ' Check this id
        If (pdxList(intI).Attributes("forestId").Value = intForestId) Then Return pdxList(intI)
      Next intI
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetForestWithId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   NextMust
  ' Goal:   Find the next constituent that MUST receive coreference information
  ' History:
  ' 29-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function NextMust(ByRef rtfThis As RichTextBox) As Boolean
    Dim intLine As Integer      ' The line in the array of XML nodes
    Dim intPtc As Integer       ' Percentage of progress

    Try
      ' Check whether busy
      If (bBusy) Then Return False
      ' Set busy
      bBusy = True
      ' Walk through all the <forest> nodes
      For intLine = loc_intLineNum To intSectSize - 1
        ' Show where we are
        intPtc = (100 * (intLine - loc_intLineNum)) \ (intSectSize - loc_intLineNum)
        Status("Checking must " & intPtc & "%", intPtc)
        ' Try get a must
        If (GetMust(rtfThis, intLine, True)) Then
          ' Reset busy
          bBusy = False
          ' Return success
          Return True
        End If
      Next intLine
      ' Everything is done
      Status("No later constituent MUST receive coreference information.")
      ' Reset busy
      bBusy = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/NextMust error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Reset busy
      bBusy = False
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   PrevMust
  ' Goal:   Find a previous constituent that MUST receive coreference information
  ' History:
  ' 14-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function PrevMust(ByRef rtfThis As RichTextBox) As Boolean
    Dim intLine As Integer      ' The line in the array of XML nodes
    Dim intPtc As Integer       ' Percentage of progress

    Try
      ' Check whether busy
      If (bBusy) Then Return False
      ' Validate
      If (loc_intLineNum = 0) Then Return False
      ' Set busy
      bBusy = True
      ' Walk through all the <forest> nodes
      For intLine = loc_intLineNum To 0 Step -1
        ' Show where we are
        intPtc = (100 * (intLine - loc_intLineNum)) \ (loc_intLineNum)
        Status("Checking must " & intPtc & "%", intPtc)
        ' Try to get one here
        If (GetMust(rtfThis, intLine, False)) Then
          ' Reset busy
          bBusy = False
          ' Return success
          Return True
        End If
      Next intLine
      ' Nothing before is found
      Status("No earlier constituent MUST receive coreference information.")
      ' Reset busy
      bBusy = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/PrevMust error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Reset busy
      bBusy = False
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetMust
  ' Goal:   Find a constituent that MUST receive coreference information
  '           on the indicated [intLine]
  ' History:
  ' 14-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetMust(ByRef rtfThis As RichTextBox, ByVal intLine As Integer, _
                          ByVal bForward As Boolean) As Boolean
    Dim intI As Integer         ' Counter of <eTree> elements
    Dim intLeft As Integer      ' The position in the textbox where this line starts
    Dim intPos As Integer       ' The correct side of the current selection
    Dim ndxList As XmlNodeList  ' List of <eTree> elements
    Dim ndxThis As XmlNode      ' One node
    Dim ndxForest As XmlNode    ' Result of GetForest

    Try
      ' Validate
      If (rtfThis Is Nothing) OrElse (intLine > UBound(arLeft)) Then Return False
      ' Determine what the position is
      intLeft = arLeft(intLine)
      ' Get the correct forest node
      ndxForest = GetForest(intLine)
      If (ndxForest Is Nothing) Then Return False
      ' Should we seek in our own local line, where we are??
      If (intLine = loc_intLineNum) AndAlso (Not loc_arSel Is Nothing) AndAlso _
        (loc_arSel.Count > 0) Then
        ' Should we go forwards or backwards?
        If (bForward) Then
          ' Only select those <eTree> elements that are further away from me!!
          intPos = loc_arSel.Item(loc_intListNum).Attributes("to").Value
          ndxList = ndxForest.SelectNodes(".//eTree[from>" & intPos & "]")
        Else
          ' Only select those <eTree> elements that are earlier than me!!
          intPos = loc_arSel.Item(loc_intListNum).Attributes("from").Value
          ndxList = ndxForest.SelectNodes(".//eTree[to<" & intPos & "]")
        End If
      Else
        ' Examine all the <eTree> elements in this line
        ndxList = ndxForest.SelectNodes(".//eTree")
      End If
      For intI = 1 To ndxList.Count
        ' Get the node
        ndxThis = ndxList(intI - 1)
        ' Is this one on the list of those that must still be done?
        If (LabelType(ndxThis) = "Must") AndAlso (Not HasCoref(ndxThis)) Then
          ' Found it! First position the cursor there
          SetSelection(rtfThis, ndxThis, intLeft)
          ' Now select it...
          SelectEtree(rtfThis, ndxThis, intLeft, colSelect)
          ' Return success
          Return True
        End If
      Next intI
      ' Return failure - nothing is found
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetMust error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetMatches
  ' Goal:   Calculate and show all possible matches
  ' History:
  ' 28-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetMatches(ByRef dtrList As DataRow()) As String
    'Dim intI As Integer         ' Counter
    'Dim intId As Integer        ' The ID of the <eTree> element (if it is there)
    'Dim intFrom As Integer      ' Value of @from attribute
    'Dim intTo As Integer        ' Value of @to attribute
    'Dim intLineNum As Integer   ' The line number of the selected element
    Dim strText As String = ""  ' The text of one line
    Dim strSeg As String = ""   ' The text of <seg>
    Dim strLabel As String = "" ' The label of the current constituent
    Dim strBack As String = ""  ' To be returned

    Try
      ' Return success
      Return strBack
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetMatches error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   EditPDE
  ' Goal:   Process the last changes in the PDE translation
  ' Note:   Where we are is in the line indicated by loc_intListNum
  ' History:
  ' 10-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function EditPDE(ByRef strPDE As String) As Boolean
    Dim ndxThis As XmlNode    ' One node
    Dim ndxForest As XmlNode  ' This line's forest

    Try
      ' Is something selected?
      If (loc_intLineNum < 0) Then
        ' Show nothing is selected
        Status("First select a line in the main editor window!!")
        ' return failure
        Return False
      End If
      ' Get the correct forest node
      ndxForest = GetForest(loc_intLineNum)
      If (ndxForest Is Nothing) Then Return False
      ' Get the node of this line
      ndxThis = ndxForest.SelectSingleNode(".//div[@lang='eng']/seg")
      ' Did we get something?
      If (ndxThis Is Nothing) Then
        ' SOmething is wrong
        MsgBox("EditPDE: could not find a location for the translation of the selected line")
        ' return failure
        Return False
      Else
        ' Add this translation
        ndxThis.InnerText = strPDE
        ' Return success
        Return True
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/EditPDE error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   NextLine
  ' Goal:   Go to the next line if possible
  ' History:
  ' 10-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function NextLine() As Boolean
    Try
      ' Let's see if there is a current number
      If (loc_intLineNum < 0) Then
        ' Do we have a [pdxList], and does it have something at all?
        If (pdxList Is Nothing) OrElse (intSectSize = 0) Then
          ' Don't do anything
        Else
          ' Adapt the line number
          loc_intLinePrev = loc_intLineNum
          ' This means we can just go to line #1
          loc_intLineNum = 0
        End If
      ElseIf (loc_intLineNum >= intSectSize - 1) Then
        ' impossible to increment...
        Return False
      Else
        ' Adapt the line number
        loc_intLinePrev = loc_intLineNum
        ' Okay we can increment
        loc_intLineNum += 1
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/NextLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   PrevLine
  ' Goal:   Go to the previous line if possible
  ' History:
  ' 10-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function PrevLine() As Boolean
    Try
      ' Let's see if there is a current number
      If (loc_intLineNum <= 0) Then
        ' Do we have a [pdxList], and does it have something at all?
        If (pdxList Is Nothing) OrElse (intSectSize = 0) Then
          ' Don't do anything
        Else
          ' Adapt the line number
          loc_intLinePrev = loc_intLineNum
          ' This means we can just go to line #1
          loc_intLineNum = 0
        End If
      Else
        ' Adapt the line number
        loc_intLinePrev = loc_intLineNum
        ' Okay we can decrement
        loc_intLineNum -= 1
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/PrevLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DelCoref
  ' Goal:   Delete the coreference relation of the currently selected element
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DelCoref(ByRef rtfThis As RichTextBox, ByVal bDoAll As Boolean) As Boolean
    Dim ndxSrc As XmlNode       ' Source node
    Dim ndxDst As XmlNode       ' Destination node
    Dim intPos As Integer       ' Note where our position was

    Try
      ' Validate
      If (loc_arSel Is Nothing) Then Return False
      If (loc_intListNum < 0) OrElse (loc_intListNum >= loc_arSel.Count) Then Return False
      ' Note our first position
      intPos = rtfThis.SelectionStart
      ' Show we are busy
      bEdtSpace = True
      ' Get the source node
      ndxSrc = loc_arSel.Item(loc_intListNum)
      ' Do we need to do all of them?
      If (bDoAll) Then
        ' Better ask first!!!
        Select Case MsgBox("Do you really want to delete the whole chain connected with the selected constituent?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.No, MsgBoxResult.Cancel
            ' Show we are busy no longer
            bEdtSpace = False
            ' Return failure
            Return False
        End Select
        ' Go into a loop
        While (Not ndxSrc Is Nothing)
          ' Determine the destination
          ndxDst = CorefDst(ndxSrc)
          ' Delete the source
          DelOneCoref(ndxSrc)
          ' Appropriately show the source
          ShowOneCoref(rtfThis, ndxSrc)
          ' The source becomes the destination
          ndxSrc = ndxDst
        End While
      Else
        ' Delete the source
        DelOneCoref(ndxSrc)
        ' Appropriately show the source
        ShowOneCoref(rtfThis, ndxSrc)
      End If
      ' Return to our first position
      With rtfThis
        .SelectionStart = intPos
        .SelectionLength = 0
      End With
      ' Show we are busy no longer
      bEdtSpace = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/DelCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DelOneCoref
  ' Goal:   Definitely delete the coreference relation of the currently selected element
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DelOneCoref(ByRef ndxThis As XmlNode) As Boolean
    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Delete the <ref> child (if it is there)
      DelXmlNodeChild(ndxThis, "ref", "")
      ' Delete the <fs type="coref"> child with all its children
      DelXmlNodeChild(ndxThis, "fs", "type;coref")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/DelOneCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DelXmlNodeChild
  ' Goal:   Delete the node [strChildName] with attributes defined in [strChildAttributes]
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DelXmlNodeChild(ByRef ndxThis As XmlNode, ByVal strChildName As String, _
                                   ByVal strChildAttributes As String) As Boolean
    Dim ndxChild As XmlNode ' The child to be deleted
    Dim strSelect As String ' Selection of the appropriate node
    Dim arAttr() As String  ' Array of attribute names and values
    Dim intI As Integer     ' COunter

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Determine the selection string
      strSelect = "./" & strChildName
      ' Are there any attributes?
      If (strChildAttributes <> "") Then
        ' Get the array of attributes
        arAttr = Split(strChildAttributes, ";")
        ' Start with [
        strSelect &= "["
        ' Walk through the array
        For intI = 0 To UBound(arAttr) Step 2
          ' Need to append "and"?
          If (intI > 0) Then strSelect &= " and "
          ' Append this attribute name and value
          strSelect &= "@" & arAttr(intI) & "='" & arAttr(intI + 1) & "'"
        Next intI
        ' Supply closing bracket
        strSelect &= "]"
      End If
      ' Determine the child to be deleted
      ndxChild = ndxThis.SelectSingleNode(strSelect)
      ' Remove it
      If (Not ndxChild Is Nothing) Then
        ' Remove all children of the child
        ndxChild.RemoveAll()
        ' Remove the child itself
        ndxThis.RemoveChild(ndxChild)
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/DelXmlNodeChild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  AddCoref
  ' Goal :  Add the selected coreference information to the PSD window
  '         The value [strType] is the kind of coreference we are adding
  ' History:
  ' 02-01-2009  ERK Created
  ' 10-04-2010  ERK Completely redone for Cesax, which works differently internally
  ' ----------------------------------------------------------------------------------------------------------
  Public Function AddCoref(ByVal strType As String) As Boolean
    Dim ndxSrc As XmlNode       ' Source node
    Dim ndxDst As XmlNode       ' Destination node
    Dim intNum As Integer       ' Number of selected constituents
    Dim intI As Integer         ' Index of currently selected source ID
    Dim strStatFile As String   ' The AutoStatistics file name

    Try
      ' Are there enough selected?
      intNum = loc_arCoref.Count
      ' Initialise the statistics report for this file/section combination
      strStatFile = AutoStatInit(strCurrentFile, intCurrentSection)
      ' See how many selections there are
      Select Case intNum
        Case 0
          ' Show that there are too little selected
          Status("First select a source (and an antecedent) constituent!")
          ' Finish statistics
          AutoStatFinish(strStatFile)
          ' Failure, so return without success
          Return False
        Case 1
          ' This is only possible for "Assumed"
          If (Not DoLike(strType, strRefOneArg)) Then
            ' Show that there are too little selected
            Status("First select a source constituent!")
            ' Finish statistics
            AutoStatFinish(strStatFile)
            ' Failure, so return without success
            Return False
          End If
          ' Get the source node
          ndxSrc = loc_arCoref.Item(0) : ndxDst = Nothing
          ' Check whether this can be used as a source
          If LabelType(ndxSrc) = "MayDst" Then
            ' Warn user
            MsgBox("According to your settings " & ndxSrc.Attributes("Label").Value & _
                   " is not allowed as source for a link")
            ' Finish statistics
            AutoStatFinish(strStatFile)
            ' Failure, so return without success
            Return False
          Else
            ' Otherwise: make the correct connection
            If (Not MakeUnlinkedNPnode(ndxSrc, "User-Manual", "Manual", strType)) Then
              ' Finish statistics
              AutoStatFinish(strStatFile)
              ' Return failure
              Return False
            End If
            ' WAS: CorefFromTo(ndxSrc, ndxDst, strType)
          End If
        Case Else
          ' This is only possible for two argument referente types
          If (Not DoLike(strType, strRefTwoArg)) Then
            ' Show that there are too little selected
            Status("First select a source AND an antecedent constituent!")
            ' Finish statistics
            AutoStatFinish(strStatFile)
            ' Failure, so return without success
            Return False
          End If
          ' Loop through source/destination stuff
          For intI = 0 To loc_arCoref.Count - 2
            ' Determine the ID of the currently selected SOURCE node
            ndxSrc = loc_arCoref.Item(intI)
            ' Determine the ID of the selected DESTINATION node
            ndxDst = loc_arCoref.Item(intI + 1)
            ' Establish a connection from Source to Destination 
            If (Not AddCorefLink(ndxSrc, ndxDst, "User-Manual", strType, "", Nothing, "", "Manual")) Then
              ' WAS: If (Not CorefFromTo(ndxSrc, ndxDst, strType)) Then
              ' Finish statistics
              AutoStatFinish(strStatFile)
              ' Failure, so return without success
              Return False
            End If
          Next intI
      End Select
      ' Finish statistics
      AutoStatFinish(strStatFile)
      ' Return well
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/AddCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CorefCount
  ' Goal:   Return the number of corefs selected
  ' History:
  ' 02-11-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CorefCount() As Integer
    Try
      Return loc_arCoref.Count
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/AddCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' -------------------------------------------------------------------------------------------------------
  ' Name:   CorefChange
  ' Goal:   Change the antecedent of [ndxSrc] to node [ndxDst] and make the relation of type [strType]
  ' History:
  ' 10-02-2012  ERK Created
  ' -------------------------------------------------------------------------------------------------------
  Public Function CorefChange(ByRef ndxSrc As XmlNode, ByRef ndxDst As XmlNode, ByVal strType As String) As Boolean
    Dim ndxAntOld As XmlNode  ' The old antecedent
    Dim ndxRef As XmlNode     ' My own <ref> node
    Dim ndxRefType As XmlNode ' My own <f name="RefType"> node

    Try
      ' Check validity of input
      If (ndxSrc Is Nothing) Or ((ndxDst Is Nothing) And _
        Not (strType = strRefAssumed OrElse strType = strRefNew OrElse strType = strRefInert _
             OrElse strType = strRefNewVar)) _
        OrElse (strType = "") Then
        ' Failure...
        Return False
      End If
      ' Get the existing antecedent
      ndxAntOld = CorefDst(ndxSrc)
      ' If we don't have an antecedent, there is nothing to change!
      If (ndxAntOld Is Nothing) Then Return False
      ' Get to the antecedent node
      ndxRef = ndxSrc.SelectSingleNode("./child::ref")
      ndxRefType = ndxSrc.SelectSingleNode("./child::fs[@type='coref']/child::f[@name='RefType']")
      ' Validate
      If (ndxRef Is Nothing) OrElse (ndxRefType Is Nothing) Then Return False
      ' Make the changes
      ndxRef.Attributes("target").Value = ndxDst.Attributes("Id").Value
      ndxRefType.Attributes("value").Value = strType
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/CorefChange error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CorefFromTo
  ' Goal:   Establish a coreference relation from Src to Dst of type strType
  ' History:
  ' 11-05-2010  ERK Created
  ' 10-03-2011  ERK Aanroep kleur toevoegen
  ' ------------------------------------------------------------------------------------
  Public Function CorefFromTo(ByRef ndxSrc As XmlNode, ByRef ndxDst As XmlNode, ByVal strType As String, _
                              Optional ByVal bAsk As Boolean = True) As Boolean
    Dim intId As Integer      ' The ID of the target node
    Dim intNdDist As Integer  ' Nodal distance
    Dim ndxChild As XmlNode   ' The <fs> child of the source node
    Dim ndxAntOld As XmlNode  ' The old antecedent
    Dim strCorefOld As String ' The existing (old) coref antecedent
    Dim strCorefNew As String ' The new coref antecedent
    Dim strSource As String   ' The source's text

    Try
      ' Check validity of input
      If (ndxSrc Is Nothing) Or ((ndxDst Is Nothing) And _
        Not (strType = strRefAssumed OrElse strType = strRefNew OrElse strType = strRefInert _
             OrElse strType = strRefNewVar)) _
        OrElse (strType = "") Then
        ' Failure...
        Return False
      End If
      ' Input should be valid, but check that there is no connection from Dst to Src already (via via)
      If (ExistsCoref(ndxDst, ndxSrc)) Then
        ' Make no relation, because that would create a circle...
        Return False
      End If
      ' Does [ndxSrc] already have a reference to something
      ndxAntOld = CorefDst(ndxSrc)
      If (Not ndxAntOld Is Nothing) AndAlso (bAsk) Then
        ' If we are in fullyauto mode, we are not going to change anything
        If (bIsFullyAuto) Then Return True
        ' Check if this has been labelled as "NoReplace"
        If (CorefAttr(ndxSrc, "NoReplace") Like "True") Then
          ' Return success, but don't make a link
          Return True
        End If
        ' What is the existing (old) coreference relation?
        strCorefOld = CorefAttr(ndxSrc, "RefType") & "[" & ndxAntOld.Attributes("Label").Value & " " & _
            NodeText(ndxAntOld) & "]"
        ' What constituent will be the new coreference relation?
        If (ndxDst Is Nothing) Then
          strCorefNew = strType & "(no antecedent)"
        Else
          strCorefNew = strType & "[" & ndxDst.Attributes("Label").Value & " " & NodeText(ndxDst) & "]"
        End If
        ' What is the source's text?
        strSource = NodeLocation(ndxSrc) & " - [" & ndxSrc.Attributes("Label").Value & " " & NodeText(ndxSrc) & "]"
        ' Warn the user?
        Select Case MsgBox("This new coreference relation is going to replace an existing one." & vbCrLf & _
               "Source: " & strSource & vbCrLf & _
               "Old antecedent: " & strCorefOld & vbCrLf & _
               "New antecedent: " & strCorefNew & vbCrLf & _
               "Replace?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.No
            ' Store the fact that this does not need replacement
            If (Not AddFeature(pdxCurrentFile, ndxSrc, "coref", "NoReplace", "True")) Then Return False
            ' Exit safely (without interrupt)
            Return False
          Case MsgBoxResult.Cancel
            ' Exit with interrupt
            bInterrupt = True
            Return False
        End Select
      End If
      ' Is there a destination node?
      If (ndxDst Is Nothing) Then
        ' The second child has to be the <fs> element
        ndxChild = SetXmlNodeChild(pdxCurrentFile, ndxSrc, "fs", "type;coref", "type", "coref")
        ' The next child of this <fs> element should be the <NdDist>
        intNdDist = 0
        If (SetXmlNodeChild(pdxCurrentFile, ndxChild, "f", "name;NdDist", "value", intNdDist) Is Nothing) Then Return False
        ' The first child of this <fs> element should be the <RefType>
        If (SetXmlNodeChild(pdxCurrentFile, ndxChild, "f", "name;RefType", "value", strType) Is Nothing) Then Return False
      Else
        ' Get the ID of the target node
        intId = ndxDst.Attributes("Id").Value
        ' The first child has to be the <ref> element
        If (SetXmlNodeChild(pdxCurrentFile, ndxSrc, "ref", "", "target", intId) Is Nothing) Then Return False
        ' The second child has to be the <fs> element
        ndxChild = SetXmlNodeChild(pdxCurrentFile, ndxSrc, "fs", "type;coref", "type", "coref")
        ' The next child of this <fs> element should be the <NdDist>
        intNdDist = ndxSrc.Attributes("Id").Value - intId
        If (SetXmlNodeChild(pdxCurrentFile, ndxChild, "f", "name;NdDist", "value", intNdDist) Is Nothing) Then Return False
        ' The first child of this <fs> element should be the <RefType>
        If (SetXmlNodeChild(pdxCurrentFile, ndxChild, "f", "name;RefType", "value", strType) Is Nothing) Then Return False
        ' We should also add the IP distance
        If (GetIpDist(ndxSrc, intNdDist)) Then
          ' Validate
          If (intNdDist >= -1) Then
            ' Add this distance
            If (SetXmlNodeChild(pdxCurrentFile, ndxChild, "f", "name;IPdist", "value", intNdDist) Is Nothing) Then Return False
          End If
        End If
      End If
      ' Are we fully automatic?
      If (Not bIsFullyAuto) Then
        ' Make sure the correct color for the source node is applied
        ShowOneCoref(frmMain.tbEdtMain, ndxSrc)
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/CorefFromTo error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddRevDesc
  ' Goal:   Add a revision description
  ' History:
  ' 05-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddRevDesc(ByRef pdxThisFile As XmlDocument, ByVal strWho As String, _
                             ByVal strWhen As String, ByVal strComment As String) As Boolean
    Dim ndxThis As XmlNode  ' Working node

    Try
      ' Validate
      If (pdxThisFile Is Nothing) Then Return False
      ' Get the correct master node
      ndxThis = SetNodePath(pdxThisFile, "teiHeader/revisionDesc")
      ' Get a good return?
      If (ndxThis Is Nothing) Then
        Logging("AddRevDesc: Could not create path teiHeader/revisionDesc")
        Return False
      End If
      ' Set the different attributes
      If (SetXmlNodeChild(pdxThisFile, ndxThis, "change", "who;" & strWho & ";when;" & strWhen, "who", strWho) Is Nothing) Then Return False
      If (SetXmlNodeChild(pdxThisFile, ndxThis, "change", "who;" & strWho & ";when;" & strWhen, "when", strWhen) Is Nothing) Then Return False
      If (SetXmlNodeChild(pdxThisFile, ndxThis, "change", "who;" & strWho & ";when;" & strWhen, "comment", strComment) Is Nothing) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/AddRevDesc error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetFileDesc
  ' Goal:   Get an item of <fileDesc>
  ' History:
  ' 05-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetFileDesc(ByRef pdxThisFile As XmlDocument, ByVal strName As String, Optional ByVal strSub As String = "") As String
    Dim ndxThis As XmlNode  ' Working node
    ' Dim strSub As String    ' Subpart of xml path

    Try
      ' Validate
      If (pdxThisFile Is Nothing) Then Return ""
      If (strSub = "") Then
        ' Get the sub parth
        Select Case strName
          Case "title", "author", "editor"
            strSub = "titleStmt"
          Case "distributor"
            strSub = "publicationStmt"
          Case "bibl"
            strSub = "sourceDesc"
          Case "manuscript", "original", "subtype", "genre"
            strSub = "creation"
          Case Else
            ' THis is not possible!
            Return ""
        End Select
      End If
      ' Get the correct master node
      ndxThis = pdxThisFile.SelectSingleNode("./descendant::teiHeader")
      ' Check if there is a header
      If (ndxThis Is Nothing) Then
        ' Return empty
        Return ""
      End If
      ndxThis = ndxThis.SelectSingleNode("./descendant::" & strSub)
      ' Get any result?
      If (ndxThis Is Nothing) Then Return ""
      ' Does the attribute exist?
      If (Not ndxThis.Attributes(strName) Is Nothing) Then
        ' Return the value of the attribute
        Return ndxThis.Attributes(strName).Value
      Else
        Return ""
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetFileDesc error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddFileDesc
  ' Goal:   Add an item of <fileDesc>
  ' History:
  ' 05-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddFileDesc(ByRef pdxThisFile As XmlDocument, ByVal strName As String, ByVal strValue As String) As Boolean
    Dim ndxThis As XmlNode  ' Working node

    Try
      ' Validate
      If (pdxThisFile Is Nothing) Then Return False
      ' Get the correct master node
      ndxThis = SetNodePath(pdxThisFile, "teiHeader/fileDesc")
      ' Get a good return?
      If (ndxThis Is Nothing) Then Return False
      ' Now the action depends on the specific element we are adding
      Select Case strName
        Case "title"
          ' Add the appropriate child and attribute
          If (SetXmlNodeChild(pdxThisFile, ndxThis, "titleStmt", "", strName, strValue) Is Nothing) Then Return False
        Case "distributor"
          ' Add the appropriate child and attribute
          If (SetXmlNodeChild(pdxThisFile, ndxThis, "publicationStmt", "", strName, strValue) Is Nothing) Then Return False
        Case "bibl"
          ' Add the appropriate child and attribute
          If (SetXmlNodeChild(pdxThisFile, ndxThis, "sourceDesc", "", strName, strValue) Is Nothing) Then Return False
        Case Else
          ' Failure...
          Return False
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/AddFileDesc error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SetNodePath
  ' Goal:   Get the indicated path [strPath] within [pdxThisFile]
  '         Create elements along the way, if they don't exist
  ' History:
  ' 05-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function SetNodePath(ByRef pdxThisFile As XmlDocument, ByVal strPath As String) _
     As XmlNode
    Dim arPath() As String    ' The path broken up in sub-parts
    Dim strSubPath As String  ' Subpath so far
    Dim intI As Integer       ' Counter
    Dim ndxThis As XmlNode    ' One node
    Dim ndxChild As XmlNode   ' Child node

    Try
      ' Validate
      If (pdxThisFile Is Nothing) Then Return Nothing
      If (strPath = "") Then Return Nothing
      ' Break up the path
      arPath = Split(strPath, "/") : strSubPath = "/" : ndxThis = pdxThisFile.SelectSingleNode("//TEI")
      ' Walk all elements of the path
      For intI = 0 To UBound(arPath)
        ' Set the subpath
        strSubPath &= "/" & arPath(intI)
        ' Find the node that belongs to this path
        ndxChild = pdxThisFile.SelectSingleNode(strSubPath)
        ' Does it exist?
        If (ndxChild Is Nothing) Then
          ' Create it
          ndxChild = pdxThisFile.CreateNode(XmlNodeType.Element, arPath(intI), Nothing)
          ' Add it to the correct place
          If (ndxThis Is Nothing) Then
            ' This is a root element
            pdxThisFile.PrependChild(ndxChild)
          Else
            ' This is a non-root element, a child of [ndxThis]
            ndxThis.PrependChild(ndxChild)
          End If
        End If
        ' Set [ndxThis] to the current child
        ndxThis = ndxChild
      Next intI
      ' Return what we've found
      Return ndxThis
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/SetNodeValue error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetAttrValue
  ' Goal:   Get the value of the indicated attribute if it exists!!
  ' History:
  ' 05-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetAttrValue(ByRef ndxThis As XmlNode, ByVal strAttrName As String) As String
    Try
      ' Does it exist?
      If (ndxThis Is Nothing) Then Return ""
      If (strAttrName = "") Then Return ""
      If (ndxThis.Attributes(strAttrName) Is Nothing) Then
        ' Is this a child-node?
        If (ndxThis.SelectSingleNode("./child::" & strAttrName) Is Nothing) Then
          Return ""
        Else
          ' Need the innertext value
          Return ndxThis.SelectSingleNode("./child::" & strAttrName).InnerText
        End If
      Else
        ' Return the attribute value
        Return ndxThis.Attributes(strAttrName).Value
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetAttrValue error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AdaptEtreeId
  ' Goal:   Adapt the @Id values of <eTree> elements starting from [intEtreeId]
  ' History:
  ' 03-01-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AdaptEtreeId(ByVal intEtreeId As Integer) As Boolean
    Dim ndxList As XmlNodeList    ' All <eTree> elements in this one
    Dim intPrevId As Integer = -1 ' Previous Id
    Dim intId As Integer          ' Current Id
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intStart As Integer       ' From where we start

    Try
      ' Need to just start at the beginning?
      If (intEtreeId = 0) Then
        intStart = 0
      Else
        ' Look for the forest that should contain it
        For intI = 0 To arLeftId.Length - 1
          If (arLeftId(intI) > intEtreeId) Then Exit For
        Next intI
        ' Check if we have it, and then go two lines before, if possible
        If (intI > 0) Then intI -= 1
        If (intI > 0) Then intI -= 1
        intStart = intI
      End If
      ' Start adapting from here
      For intI = intStart To pdxList.Count - 1
        ' Get all the <eTree> elements in here
        ndxList = pdxList(intI).SelectNodes(".//eTree")
        For intJ = 0 To ndxList.Count - 1
          ' Get its Id
          If (ndxList(intJ).Attributes("Id").ToString = "") Then
            ' Take the previous id
            intId = intPrevId
          Else
            ' Get the ID we already have
            intId = ndxList(intJ).Attributes("Id").Value
          End If
          ' intId = ndxList(intJ).Attributes("Id").Value
          ' Check if this needs adaptation
          If (intId <= 0) Then
            intId = 1
            ndxList(intJ).Attributes("Id").Value = intId
            ' Show we are dirty
            frmMain.SetDirty(True)
          ElseIf (intPrevId > 0) AndAlso (intId <> intPrevId + 1) Then
            ' Adapt it
            intId = intPrevId + 1
            ndxList(intJ).Attributes("Id").Value = intId
            ' Show we are dirty
            frmMain.SetDirty(True)
          End If
          ' Adapt previous id
          intPrevId = intId
        Next intJ
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/AdaptEtreeId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CreateNewEtree
  ' Goal:   Create a new <eTree> element to the indicated xml document
  ' History:
  ' 03-01-2012  ERK Created
  ' 01-02-2014  ERK Added @b and @e attributes
  ' ------------------------------------------------------------------------------------
  Public Function CreateNewEtree(ByRef pdxThisFile As XmlDocument) As XmlNode
    Dim intI As Integer           ' Counter
    Dim ndxThis As XmlNode        ' Newly to be created node
    Dim atxChild As XmlAttribute  ' The attribute we are looking for
    Dim arAttr() As String = {"Id", "1", "Label", "new", "IPnum", "0", "from", "0", "to", "0"}

    Try
      ' Validate
      If (pdxThisFile Is Nothing) Then Return Nothing
      ' Create a new <eTree> node
      ndxThis = pdxThisFile.CreateNode(XmlNodeType.Element, "eTree", Nothing)
      ' Create all necessary attributes
      For intI = 0 To UBound(arAttr) Step 2
        ' Add this attribute
        atxChild = pdxThisFile.CreateAttribute(arAttr(intI))
        atxChild.Value = arAttr(intI + 1)
        ndxThis.Attributes.Append(atxChild)
      Next intI
      ' Set the Id attribute to a different value
      intMaxEtreeId += 1 : ndxThis.Attributes("Id").Value = intMaxEtreeId
      ' Return the new node
      Return ndxThis
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/CreateNewEtree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CreateNewEleaf
  ' Goal:   Create a new <eLeaf> element to the indicated xml document
  ' History:
  ' 03-01-2012  ERK Created
  ' 01-02-2014  ERK Added @n feature
  ' ------------------------------------------------------------------------------------
  Public Function CreateNewEleaf(ByRef pdxThisFile As XmlDocument) As XmlNode
    Dim intI As Integer           ' Counter
    Dim ndxThis As XmlNode        ' Newly to be created node
    Dim atxChild As XmlAttribute  ' The attribute we are looking for
    Dim arAttr() As String = {"Type", "Vern", "Text", "new", "prob", "0", "from", "0", "to", "0", "n", "0"}

    Try
      ' Validate
      If (pdxThisFile Is Nothing) Then Return Nothing
      ' Create a new <eTree> node
      ndxThis = pdxThisFile.CreateNode(XmlNodeType.Element, "eLeaf", Nothing)
      ' Create all necessary attributes
      For intI = 0 To UBound(arAttr) Step 2
        ' Add this attribute
        atxChild = pdxThisFile.CreateAttribute(arAttr(intI))
        atxChild.Value = arAttr(intI + 1)
        ndxThis.Attributes.Append(atxChild)
      Next intI
      ' Return the new node
      Return ndxThis
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/CreateNewEleaf error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SetXmlNodeChild
  ' Goal:   Set the attribute [strChildAttr] of the child named [strChildName] of 
  '           xml node [ndxThis] with value [strChildValue]
  '         If this child does not exist, then create it as the FIRST child
  ' History:
  ' 12-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function SetXmlNodeChild(ByRef pdxThisFile As XmlDocument, ByRef ndxThis As XmlNode, _
      ByVal strChildName As String, ByVal strAttrSel As String, ByVal strChildAttr As String, _
      ByVal strChildValue As String) As XmlNode
    Dim ndxChild As XmlNode           ' Child being looked for
    Dim atxChild As XmlAttribute      ' The attribute we are looking for
    Dim arAttr() As String = Nothing  ' Array of attribute name + value
    Dim strSelect As String           ' The string to find the appropriate child
    Dim intI As Integer               ' Counter

    Try
      ' Verify the input
      If (ndxThis Is Nothing) OrElse (strChildName = "") OrElse (strChildAttr = "") Then Return Nothing
      ' Start making the selection string
      strSelect = "./" & strChildName
      If (strAttrSel <> "") Then
        ' Extract the list of necessary attributes
        arAttr = Split(strAttrSel, ";")
        ' Append left bracket
        strSelect &= "["
        ' Visit all attributes
        For intI = 0 To UBound(arAttr) Step 2
          ' Should we add the logical "AND"?
          If (intI > 0) Then
            strSelect &= " and "
          End If
          ' Add this attribute to the list
          strSelect &= "@" & arAttr(intI) & "='" & arAttr(intI + 1) & "'"
        Next intI
        ' Append right bracket
        strSelect &= "]"
      End If
      ' Try to get the appropriate child
      ndxChild = ndxThis.SelectSingleNode(strSelect)
      ' Is there a child?
      If (ndxChild Is Nothing) Then
        ' Create a new child node
        ndxChild = pdxThisFile.CreateNode(XmlNodeType.Element, strChildName, Nothing)
        ' Should first other attributes be created?
        If (strAttrSel <> "") Then
          ' Create all necessary attributes
          For intI = 0 To UBound(arAttr) Step 2
            ' Add this attribute
            atxChild = pdxThisFile.CreateAttribute(arAttr(intI))
            atxChild.Value = arAttr(intI + 1)
            ndxChild.Attributes.Append(atxChild)
          Next intI
        End If
        ' Create a new attribute
        atxChild = pdxThisFile.CreateAttribute(strChildAttr)
        ' Add the attribute to the node
        ndxChild.Attributes.Append(atxChild)
        ' Make this new node the FIRST child of [ndxThis]
        ndxThis.PrependChild(ndxChild)
      End If
      ' Find the attribute of the child
      atxChild = ndxChild.Attributes(strChildAttr)
      If (atxChild) Is Nothing Then
        ' The attribute does not exist, so create it
        atxChild = pdxThisFile.CreateAttribute(strChildAttr)
        ' Add the attribute to the node
        ndxChild.Attributes.Append(atxChild)
      End If
      ' Set the attribute of the child
      atxChild.Value = strChildValue
      ' Return the child
      Return ndxChild
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/SetXmlNodeChild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ExistsCoref
  ' Goal:   Check if [ndxSrc] connects somehow to [ndxDst]
  ' History:
  ' 11-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function ExistsCoref(ByRef ndxSrc As XmlNode, ByRef ndxDst As XmlNode) As Boolean
    Dim intId As Integer    ' Id of the destination
    Dim ndxNext As XmlNode  ' Next node
    Dim bFound As Boolean   ' Flag to indicate success

    Try
      ' Check validity
      If (ndxSrc Is Nothing) OrElse (ndxDst Is Nothing) Then Return False
      ' Get the Id of the destination
      intId = ndxDst.Attributes("Id").Value
      ' Start with the source
      ndxNext = ndxSrc : bFound = False
      ' Go into a loop
      While (Not ndxNext Is Nothing) AndAlso (Not bFound)
        ' Get the next one
        ndxNext = CorefDst(ndxNext)
        ' Get the flag
        If (Not ndxNext Is Nothing) Then bFound = (ndxNext.Attributes("Id").Value = intId)
      End While
      ' Return the result
      If (Not bFound) Then
        ' No clear result
        Return Nothing
      Else
        ' Return the result
        Return (ndxNext.Attributes("Id").Value = intId)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/ExistsCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetFeatVect
  ' Goal:   Get all features into a |-separated feature vector string
  ' History:
  ' 24-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetFeatVect(ByRef ndxThis As XmlNode, ByVal ParamArray arExclude() As String) As String
    Dim ndxList As XmlNodeList  ' List of features
    Dim strBack As String = ""  ' What we return
    Dim intI As Integer         ' Counter
    ' Dim intJ As Integer         ' Exclude features

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (ndxThis.Name <> "eTree") Then Return ""
      ' Get all possible features
      ndxList = ndxThis.SelectNodes("./child::fs/child::f")
      For intI = 0 To ndxList.Count - 1
        ' Include this or not?
        If (Not Array.Exists(arExclude, Function(strIn As String) strIn = ndxList(intI).Attributes("name").Value)) Then
          ' Need to add separator?
          If (strBack <> "") Then strBack &= "|"
          ' Add this feature
          strBack &= ndxList(intI).ParentNode.Attributes("type").Value & "::" & _
            ndxList(intI).Attributes("name").Value & "=" & ndxList(intI).Attributes("value").Value
        End If
      Next intI
      ' Return the result
      Return strBack
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetFeatVect error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetFeature
  ' Goal:   Get the value of the feature having <fs type='strType'> and name <f name='strName'>
  ' History:
  ' 03-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetFeature(ByRef ndxSrc As XmlNode, ByVal strType As String, ByVal strName As String) As String
    Dim ndxThis As XmlNode  ' Result of select statement

    Try
      ' Determine validity
      If (ndxSrc Is Nothing) Then Return ""
      ' Valid input... -- see if we can get a child...
      ndxThis = ndxSrc.SelectSingleNode("./fs[@type='" & strType & "']")
      ' Is this correct?
      If (Not ndxThis Is Nothing) Then
        ' Get the necessary <f> element
        ndxThis = ndxThis.SelectSingleNode("./f[@name='" & strName & "']")
        ' Does it exist?
        If (Not ndxThis Is Nothing) Then
          ' Return the value of this attribute
          Return Trim(ndxThis.Attributes("value").Value)
        End If
      End If
      ' Return failure
      Return ""
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetFeature error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DelFeature
  ' Goal:   If existant, remove feature with <fs type='strType'> and name <f name='strName'>
  ' History:
  ' 27-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DelFeature(ByRef ndxSrc As XmlNode, ByVal strType As String, ByVal strName As String) As Boolean
    Dim ndxThis As XmlNode  ' Result of select statement
    Dim ndxPar As XmlNode   ' Parent

    Try
      ' Determine validity
      If (ndxSrc Is Nothing) Then Logging("DelFeature: Cannot delete") : Return False
      ' Valid input... -- see if we can get a child...
      ndxThis = ndxSrc.SelectSingleNode("./fs[@type='" & strType & "']")
      ' Does the feature type exist?
      If (ndxThis Is Nothing) Then Return True
      ' Get the necessary <f> element
      ndxThis = ndxThis.SelectSingleNode("./f[@name='" & strName & "']")
      ' Does the feature with this name exist?
      If (ndxThis IsNot Nothing) Then
        ' Remove the feature
        ndxPar = ndxThis.ParentNode
        ndxThis.RemoveAll()
        ndxPar.RemoveChild(ndxThis)
        ' Check if there are any children left
        If (ndxPar.ChildNodes.Count = 0) Then
          ' Remove the [fs] node too
          ndxThis = ndxPar
          ndxPar = ndxThis.ParentNode
          ndxThis.RemoveAll()
          ndxPar.RemoveChild(ndxThis)
        End If
      End If
      ' Return positively 
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/DelFeature error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   HasFeature
  ' Goal:   Check existence of feature with <fs type='strType'> and name <f name='strName'>
  ' History:
  ' 20-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function HasFeature(ByRef ndxSrc As XmlNode, ByVal strType As String, ByVal strName As String) As Boolean
    Dim ndxThis As XmlNode  ' Result of select statement

    Try
      ' Determine validity
      If (ndxSrc Is Nothing) Then Return False
      ' Valid input... -- see if we can get a child...
      ndxThis = ndxSrc.SelectSingleNode("./fs[@type='" & strType & "']")
      ' Does the feature type exist?
      If (ndxThis Is Nothing) Then Return False
      ' Get the necessary <f> element
      ndxThis = ndxThis.SelectSingleNode("./f[@name='" & strName & "']")
      ' Does the feature with this name exist?
      Return (Not ndxThis Is Nothing)
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/HasFeature error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CorefAttr
  ' Goal:   Get the value of the coreference attribute of [ndxSrc] if it exists
  ' History:
  ' 12-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CorefAttr(ByRef ndxSrc As XmlNode, ByVal strAttrName As String) As String
    Dim ndxThis As XmlNode  ' Result of select statement

    Try
      ' Determine validity
      If (ndxSrc Is Nothing) Then Return Nothing
      ' Valid input... -- see if we can get a child...
      ndxThis = ndxSrc.SelectSingleNode("./fs[@type='coref']")
      ' Any luck?
      If (ndxThis Is Nothing) Then
        ' No coref  node!
        Return ""
      Else
        ' Determine the Id of the target
        ndxThis = ndxThis.SelectSingleNode("./f[@name='" & strAttrName & "']")
        ' Any result?
        If (ndxThis Is Nothing) Then
          ' Return failre
          Return ""
        Else
          ' Return the value of this attribute
          Return ndxThis.Attributes("value").Value
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/CorefAttr error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CorefDstSetRef
  ' Goal:   Set the reference ID of the destination node of [ndxSrc]
  ' History:
  ' 11-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CorefDstSetRef(ByRef ndxSrc As XmlNode, ByVal intDstId As Integer) As Boolean
    Dim ndxThis As XmlNode  ' Result of select statement

    Try
      ' Determine validity
      If (ndxSrc Is Nothing) OrElse (intDstId < 0) Then Return False
      ' Valid input... -- see if we can get a child...
      ndxThis = ndxSrc.SelectSingleNode("./fs[@type='coref']")
      ' Any luck?
      If (Not ndxThis Is Nothing) Then
        ' Get to the antecedent reference
        ndxThis = ndxSrc.SelectSingleNode("./ref")
        ' Did we get a return?
        If (Not ndxThis Is Nothing) Then
          ' Set the target reference
          ndxThis.Attributes("target").Value = intDstId
          ' Return success
          Return True
        End If
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/CorefDstSetRef error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CorefDstId
  ' Goal:   Get the ID of the destination node of [ndxSrc] if it exists
  ' History:
  ' 11-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CorefDstId(ByRef ndxSrc As XmlNode) As Integer
    Dim ndxThis As XmlNode  ' Result of select statement
    Dim intDstId As Integer ' Id of the target
    Dim intSrcId As Integer ' Id of the source

    Try
      ' Determine validity
      If (ndxSrc Is Nothing) Then Return Nothing
      ' Get the source's Id
      intSrcId = ndxSrc.Attributes("Id").Value
      ' Valid input... -- see if we can get a child...
      ndxThis = ndxSrc.SelectSingleNode("./fs[@type='coref']")
      ' Any luck?
      If (ndxThis Is Nothing) Then
        ' No coref  node!
        Return -1
      Else
        ' Determine the Id of the target
        ndxThis = ndxSrc.SelectSingleNode("./ref")
        ' Did we get a return?
        If (ndxThis Is Nothing) Then
          ' Return empty
          Return -1
        Else
          intDstId = ndxThis.Attributes("target").Value
          ' Return the correct XmlNode from this Id
          Return intDstId
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/CorefDstId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CorefDst
  ' Goal:   Get the destination node of [ndxSrc] if it exists
  ' History:
  ' 11-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CorefDst(ByRef ndxSrc As XmlNode) As XmlNode
    Dim ndxThis As XmlNode  ' Result of select statement
    Dim intDstId As Integer ' Id of the target
    Dim intSrcId As Integer ' Id of the source

    Try
      ' Determine validity
      If (ndxSrc Is Nothing) Then Return Nothing
      ' Get the source's Id
      intSrcId = ndxSrc.Attributes("Id").Value
      ' Valid input... -- see if we can get a child...
      ndxThis = ndxSrc.SelectSingleNode("./fs[@type='coref']")
      ' Any luck?
      If (ndxThis Is Nothing) Then
        ' No coref  node!
        Return Nothing
      Else
        ' Determine the Id of the target
        ndxThis = ndxSrc.SelectSingleNode("./ref")
        ' Did we get a return?
        If (ndxThis Is Nothing) Then
          ' Return empty
          Return Nothing
        Else
          intDstId = ndxThis.Attributes("target").Value
          ' Return the correct XmlNode from this Id
          Return IdToNode(intDstId)
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/CorefDst error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetSectionLineNum
  ' Goal:   Get the number of the line where [intPos] finds itself 
  '         This uses recursive bisection
  ' History:
  ' 18-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetSectionLineNum(ByVal intPos As Integer, ByRef intMin As Integer, _
                                    ByRef intMax As Integer) As Integer
    Dim intHalf As Integer    ' Division between Min and Max

    Try
      ' Are we down to 1 line?
      If (intMin = intMax) Then
        ' Return this line
        Return intMin
        ' Debug.Print(arLeftId(16) & " " & arLeftId(17))
      End If
      ' Get the half of the section
      intHalf = intMin + (intMax - intMin) \ 2
      ' Try the first half
      If (intPos >= arLeft(intMin)) AndAlso (intPos < arLeft(intHalf + 1)) Then
        ' The answer must be here
        Return GetSectionLineNum(intPos, intMin, intHalf)
      Else
        ' The answer must be higher
        Return GetSectionLineNum(intPos, intHalf + 1, intMax)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetSectionLineNum error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetLineIdx
  ' Goal:   Get the index of the line that should contain the <eTree> with [intId]
  ' History:
  ' 18-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetLineIdx(ByVal intId As Integer, ByRef intMin As Integer, ByRef intMax As Integer) As Integer
    Dim intHalf As Integer    ' Division between Min and Max

    Try
      ' Are we down to 1 line?
      If (intMin = intMax) Then
        ' Return this line
        Return intMin
        ' Debug.Print(arLeftId(16) & " " & arLeftId(17))
      End If
      ' Get the half of the section
      intHalf = intMin + (intMax - intMin) \ 2
      ' Try the first half
      If (intId >= arLeftId(intMin)) AndAlso (intId < arLeftId(intHalf + 1)) Then
        ' The answer must be here
        Return GetLineIdx(intId, intMin, intHalf)
      Else
        ' The answer must be higher
        Return GetLineIdx(intId, intHalf + 1, intMax)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetLineIdx error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   TravSel
  ' Goal:   Walk [ndxThis] and return node with eTree/Id=[intId] in [ndxSel]
  ' Note:   Walking is done breadth-first, depth-last
  ' History:
  ' 19-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function TravSel(ByRef ndxThis As XmlNode, ByRef ndxSel As XmlNode, ByVal intId As Integer) As Boolean
    Dim ndxNext As XmlNode    ' Next node in series

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' (1) Breadth first: check myself and my sisters
      ndxNext = ndxThis.FirstChild
      While (Not ndxNext Is Nothing)
        ' Only evaluate <eTree> types
        If (ndxNext.Name = "eTree") Then
          ' Check this node's ID
          If (ndxNext.Attributes("Id").Value = intId) Then
            ndxSel = ndxNext
            Return True
          End If
          ' Check its children
          If (TravSel(ndxNext, ndxSel, intId)) Then Return True
        End If
        ' Try next sibling
        ndxNext = ndxNext.NextSibling
      End While
      ' Return failure if we get here
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/TravSel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IdToNode
  ' Goal:   Determine the [XmlNode] of this [intId]
  ' Note:   There is a problem in that preceding nodes are not always found relative to the existing node
  '         I don't know why...
  ' History:
  ' 11-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function IdToNode(ByVal intId As Integer) As XmlNode
    Dim ndxThis As XmlNode    ' Result of select statement
    Dim ndxForest As XmlNode  ' This line's forest
    Dim intI As Integer       ' Counter for the [arLeft] array
    Dim intMin As Integer = 0 ' Minimum index of [arLeftId]
    Dim intMax As Integer = 0 ' Maximum index of [arLeftId]

    Try
      ' Validate
      If (pdxList Is Nothing) Then Return Nothing
      ' Set maximum
      intMax = pdxList.Count - 1
      ' Get the index of the <forest> line we must pursue
      intI = GetLineIdx(intId, intMin, intMax)
      ' Verify the answer
      If (intI < 0) Then Return Nothing
      ' Look in this line for the correct node
      ndxForest = pdxList(intI)
      ' Try to get the node in this domain
      'ndxThis = ndxForest.SelectSingleNode(".//eTree[@Id=" & intId & "]")
      ' Alternative solution: same effect, but less time
      ndxThis = Nothing
      If (Not TravSel(ndxForest, ndxThis, intId)) Then
        ' Failure
        Return Nothing
      End If
      ' Did we find it?
      If (Not ndxThis Is Nothing) Then
        ' Return the result
        Return ndxThis
      End If
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/IdToNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   HasStartSection
  ' Goal:   Check if the start of the document, the first <forest>, has a section
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function HasStartSection() As Boolean
    Dim ndxForest As XmlNode    ' The <forest> element starting the section

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Return False
      ' Get the first forest line
      ndxForest = pdxCurrentFile.SelectSingleNode("//forest")
      ' Does this have a section
      Return (Not ndxForest.Attributes("Section") Is Nothing)
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/HasStartSection error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddStartSection
  ' Goal:   Add a Section attribute at the first <forest> element
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddStartSection() As Boolean
    Dim ndxForest As XmlNode    ' The <forest> element starting the section

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Return False
      ' Get the first forest line
      ndxForest = pdxCurrentFile.SelectSingleNode("//forest")
      ' Does this have a section
      If (ndxForest.Attributes("Section") Is Nothing) Then
        ' Add a section here
        If (Not AddSectionAttribute(ndxForest)) Then Return False
      End If
      ' Return succes
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/AddStartSection error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   HasSection
  ' Goal:   See if there is a section at the indicated line in the XML document
  ' History:
  ' 13-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function HasSection(ByVal intLine As Integer) As Boolean
    Dim ndxForest As XmlNode    ' The <forest> element starting the section

    Try
      ' Are we anywhere?
      If (intLine < 0) OrElse (intLine > intSectSize - 1) Then Exit Function
      ' Get where we are
      ndxForest = GetForest(intLine)
      If (ndxForest Is Nothing) Then Return False
      ' Return the presence of section or not
      Return (Not ndxForest.Attributes("Section") Is Nothing)
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/HasSection error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DeriveSections
  ' Goal:   Make sections afresh
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DeriveSections() As Boolean
    Try
      ' Validate
      If (pdxSection Is Nothing) Then SectionCalculate()
      Return AutoAddSections()
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/DeriveSections error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DelAllSections
  ' Goal:   Remove all sections
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DelAllSections() As Boolean
    Dim ndxForest As XmlNode = Nothing  ' The forest we are considering
    Dim intI As Integer   ' COunter

    Try
      ' Validate
      If (pdxSection Is Nothing) Then SectionCalculate()
      ' Go through all sections
      For intI = 0 To pdxSection.Count - 1
        ' Get the appropriate forest node
        If (Not GetForestNode(pdxSection(intI), ndxForest)) Then Return False
        ' Check to see if this has a section
        If (Not ndxForest.Attributes("Section") Is Nothing) Then
          ' There is a section, so delete it
          ndxForest.Attributes.Remove(ndxForest.Attributes("Section"))
        End If
      Next intI
      ' Recalculate sections
      SectionCalculate()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/DelAllSections error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetTorontoId
  ' Goal:   Automatically try to add sections to this pdx document
  ' History:
  ' 20-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetTorontoId(ByRef ndxThis As XmlNode) As String
    Dim ndxChild As XmlNode ' The <eLeaf> child
    Dim strText As String   ' Text of the node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      ' Is this a forest or an eTree?
      If (ndxThis.Name = "forest") Then
        ' Get to the first <eTree>
        ndxChild = ndxThis.Item("eTree")
        If (ndxChild Is Nothing) Then Return ""
        ' It should have label CODE
        If (ndxChild.Attributes("Label").Value <> "CODE") Then Return ""
        ' Get its first child
        ndxChild = ndxChild.Item("eLeaf")
      Else
        ' Get the <eLeaf> child
        ndxChild = ndxThis.Item("eLeaf")
      End If
      ' Double check
      If (ndxChild Is Nothing) Then Return ""
      If (ndxChild.Attributes("Text") Is Nothing) Then Return ""
      ' Get the correct text
      strText = ndxChild.Attributes("Text").Value
      ' Double check
      If (Left(strText, 2) <> "<T") Then Return ""
      ' Okay, now return the Id
      strText = Left(strText, 7)
      ' Double check the numeric part of this id
      If (IsNumeric(Mid(strText, 3))) Then
        ' Okay, this is a toronto id
        Return strText
      Else
        ' No, this is just something starting with (CODE <T...), but not a toronto Id
        Return ""
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetTorontoId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsYcoe
  ' Goal:   If this is YCOE, then it has a (CODE <T...) start
  ' History:
  ' 20-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function IsYcoe(ByRef pdxFile As XmlDocument) As Boolean
    Dim ndxNext As XmlNode  ' One node
    Dim strText As String   ' Text value of the <eLeaf>

    Try
      ' Validate
      If (pdxFile Is Nothing) Then Return False
      ' Get the first <TEI> child
      ndxNext = pdxFile.Item("TEI")
      If (Not ndxNext Is Nothing) Then
        ' Get the first <forestGrp> child
        ndxNext = ndxNext.Item("forestGrp")
        If (Not ndxNext Is Nothing) Then
          ' Get the first <forest> child
          ndxNext = ndxNext.Item("forest")
          If (Not ndxNext Is Nothing) Then
            ' Get the first <eTree> child
            ndxNext = ndxNext.Item("eTree")
            If (Not ndxNext Is Nothing) Then
              ' This <eTree> should have a label CODE
              If (ndxNext.Attributes("Label").Value = "CODE") Then
                ' Get the first <eLeaf> child
                ndxNext = ndxNext.Item("eLeaf")
                If (Not ndxNext Is Nothing) Then
                  ' This <eLeaf> should have a particular text value
                  strText = ndxNext.Attributes("Text").Value
                  Return (Left(strText, 2) = "<T")
                End If
              End If
            End If
          End If
        End If
      End If
      ' In all other cases: no match
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/IsYcoe error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetFirstForest
  ' Goal:   Get the first <forest> node in this Xml document
  ' History:
  ' 20-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetFirstForest(ByRef pdxFile As XmlDocument, ByRef ndxForest As XmlNode) As Boolean
    ' Dim ndxNext As XmlNode  ' One node

    Try
      ' Validate
      If (pdxFile Is Nothing) Then Return False
      ' Get the first forest node in an easier way...
      ndxForest = pdxFile.SelectSingleNode("./descendant::forest[1]")
      '' Get the first <TEI> child
      'ndxNext = pdxFile.Item("TEI")
      'If (Not ndxNext Is Nothing) Then
      '  ' Get the first <forestGrp> child
      '  ndxNext = ndxNext.Item("forestGrp")
      '  If (Not ndxNext Is Nothing) Then
      '    ' Get the first <forest> child
      '    ndxNext = ndxNext.Item("forest")
      '    If (Not ndxNext Is Nothing) Then
      '      ' Define the value of this forest node
      '      ndxForest = ndxNext
      '      ' Return success
      '      Return True
      '    End If
      '  End If
      'End If
      ' Check the result
      If (ndxForest Is Nothing) Then
        ' In all other cases: no match
        Logging("GetFirstForest: could not find a <forest> node in this XmlDocument")
        Return False
      Else
        Return True
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetFirstForest error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   HasHeading
  ' Goal:   Find out whether this <forest> has an <eTree> child with CODE <heading>
  ' History:
  ' 20-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function HasHeading(ByRef ndxForest As XmlNode) As Boolean
    Dim ndxNext As XmlNode  ' Working node
    Dim strText As String   ' Text of <eLeaf>

    Try
      ' Validate
      If (ndxForest Is Nothing) Then Return False
      ' Get the first <eTree> child
      ndxNext = ndxForest.Item("eTree")
      If (Not ndxNext Is Nothing) Then
        ' The label of this child should be CODE
        If (ndxNext.Attributes("Label").Value = "CODE") Then
          ' Get the <eLeaf> child of it
          ndxNext = ndxNext.Item("eLeaf")
          If (Not ndxNext Is Nothing) Then
            ' Get the text of this <eLeaf>
            strText = ndxNext.Attributes("Text").Value
            ' Is this a heading?
            Return (strText Like "<heading>")
          End If
        End If
      End If
      ' No heading found
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/HasHeading error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AutoAddSections
  ' Goal:   Automatically try to add sections to this pdx document
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AutoAddSections() As Boolean
    Dim ndxForest As XmlNode = Nothing  ' The forest we are considering
    Dim strToronto As String = ""       ' The <Tnnnnn identifier indicating the Toronto file
    Dim strThisId As String = ""        ' The current id
    Dim bUseToronto As Boolean = False  ' Whether we need to use the Toronto <Tnnnnn> indications or not
    Dim bMakeBreak As Boolean = False   ' Whether we need to make a break or not
    Dim intMinForestId As Integer = 5   ' The minimum forestId after which we allow another heading
    Dim intNum As Integer               ' Number of forests
    Dim intPtc As Integer               ' Percentage

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Return False
      ' Show what we are doing
      Status("Looking for sections....")
      ' Get the first forest node
      If (Not GetFirstForest(pdxCurrentFile, ndxForest)) Then Return False
      ' Add a section break here at the beginning
      If (Not AddSectionAttribute(ndxForest)) Then Return False
      ' Determine whether we are toronto or not
      bUseToronto = IsYcoe(pdxCurrentFile)
      ' For toronto we need to get a string
      If (bUseToronto) Then
        ' Note the 5-numbers after <T...>
        strToronto = GetTorontoId(ndxForest)
      End If
      ' Get number of forests
      intNum = ndxForest.ParentNode.ChildNodes.Count
      ' Go to the next forest for the loop
      ndxForest = ndxForest.NextSibling
      ' Go into a loop
      While (Not ndxForest Is Nothing)
        ' Can we show where we are?
        intPtc = (ndxForest.Attributes("forestId").Value * 100) \ intNum
        Status("Looking for sections " & intPtc & "%", intPtc)
        ' Initially we assume no break has to be made here
        bMakeBreak = False
        ' Check whether we need to add a section break here
        If (bUseToronto) Then
          ' Check if the Toronto Id is changing
          strThisId = GetTorontoId(ndxForest)
          ' ============= DEBUG ============
          'If (strThisId <> "") Then Stop
          ' ================================
          If (strThisId <> "") AndAlso (strThisId <> strToronto) Then
            ' Make this the toronto id
            strToronto = strThisId
            ' Make sure we can make a section break
            bMakeBreak = True
          End If
        Else
          ' Check if a <heading> node occurs here, and if it is not too close to the start
          If (HasHeading(ndxForest)) AndAlso (ndxForest.Attributes("forestId").Value > intMinForestId) Then
            ' Okay, we need to make a break here
            bMakeBreak = True
          End If
        End If
        ' Do we need to make a break?
        If (bMakeBreak) Then
          ' Add the Section attribute to this forest element
          If (Not AddSectionAttribute(ndxForest)) Then Return False
        End If
        ' Go to the next forest
        ndxForest = ndxForest.NextSibling
      End While


      'If (IsYcoe(pdxCurrentFile)) Then
      '  ' We are toronto!

      'Else

      'End If
      '' Look for all (CODE <heading>) nodes
      'ndxList = pdxCurrentFile.SelectNodes("//eTree[@Label='CODE' and contains(./eLeaf/@Text, '<heading>')]")
      '' ANy results?
      'If (ndxList.Count = 0) Then
      '  ' We don't have (CODE <heading>) nodes, but we might have (CODE <Tnnnnn....>) nodes for toronto
      '  If (IsYcoe(pdxCurrentFile)) Then
      '    Status("Looking for Toronto sections (OE)...")
      '    ' Yes, there are probably toronto nodes, so we need to do derive section breaks differently
      '    ndxList = pdxCurrentFile.SelectNodes("//eTree[@Label='CODE' and starts-with(./eLeaf/@Text, '<T')]")
      '    bUseToronto = True
      '  Else
      '    ' No, so we are just not able to add section breaks based on the content
      '    Return False
      '  End If
      'End If
      '' There are results - at least add a section at the first <forest> element
      'ndxForest = pdxCurrentFile.SelectSingleNode("//forest")
      'If (Not AddSectionAttribute(ndxForest)) Then Return False
      '' Are we in Toronto mode or not?
      'If (bUseToronto) Then
      '  ' Note the 5-numbers after <T...>
      '  strToronto = GetTorontoId(ndxList(0))
      '  ' Set the start alright
      '  intStart = 1
      'Else
      '  ' See if we should take over the first <heading> or not...
      '  If (ndxList(0).Attributes("Id").Value < 100) Then
      '    intStart = 1
      '  Else
      '    intStart = 0
      '  End If
      'End If
      '' Visit all the CODE nodes...
      'For intI = intStart To ndxList.Count - 1
      '  ' Assume that we don't have to make a section break here
      '  bMakeBreak = False
      '  ' Are we in toronto mode?
      '  If (bUseToronto) Then
      '    ' Check if the Toronto Id is changing
      '    strThisId = GetTorontoId(ndxList(intI))
      '    If (strThisId <> "") AndAlso (strThisId <> strToronto) Then
      '      ' Make this the toronto id
      '      strToronto = strThisId
      '      ' Make sure we can make a section break
      '      bMakeBreak = True
      '    End If
      '  End If
      '  ' Do we need to make a break?
      '  If (Not bUseToronto) OrElse (bMakeBreak) Then
      '    ' Get the forest belonging to this one
      '    If (Not GetForestNode(ndxList(intI), ndxForest)) Then Return False
      '    ' Add the Section attribute to this forest element
      '    If (Not AddSectionAttribute(ndxForest)) Then Return False
      '  End If
      'Next intI
      ' Calculate sections afresh
      SectionCalculate()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/AutoAddSections error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddSectionAttribute
  ' Goal:   Add a section attribute at the [ndxThis] forest node
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function AddSectionAttribute(ByRef ndxForest As XmlNode) As Boolean
    Dim atxSect As XmlAttribute ' The new attribute to add

    Try
      ' Validate
      If (ndxForest Is Nothing) Then Return False
      If (ndxForest.Name <> "forest") Then Return False
      ' Check if the appropriate attribute already exists
      If (ndxForest.Attributes("Section") Is Nothing) Then
        ' No section yet, so make it
        atxSect = pdxCurrentFile.CreateAttribute("Section")
        ' Add this attribute to the appropriate <forest> node
        ndxForest.Attributes.Append(atxSect)
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/AddSectionAttribute error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddSection
  ' Goal:   Add a section at the indicated line in the XML document
  ' History:
  ' 13-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddSection(ByVal intLine As Integer) As Boolean
    Dim ndxForest As XmlNode    ' The <forest> element starting the section

    Try
      ' Are we anywhere?
      If (intLine < 0) OrElse (intSectFirst + intLine > pdxList.Count - 1) Then Exit Function
      ' Get where we are
      ndxForest = GetForest(intLine)
      If (ndxForest Is Nothing) Then Return False
      ' Add the section attribute
      If (Not AddSectionAttribute(ndxForest)) Then Return False
      ' Calculate sections afresh
      SectionCalculate()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/AddSection error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   SectionInsert
  ' Goal:   Insert a section before this line
  ' History:
  ' 13-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SectionInsert(ByRef rtfThis As RichTextBox)
    Try
      ' Are we anywhere?
      If (loc_intLineNum < 0) OrElse (loc_intLineNum > arLeft.Length - 1) Then Exit Sub
      ' Add the section
      If (AddSection(loc_intLineNum)) Then
        ' Set this line to Italic
        LineItalic(rtfThis, loc_intLineNum, True)
        ' Calculate sections afresh
        SectionCalculate()
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/SectionInsert error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   SectionCalculate
  ' Goal:   Calculate all the sections afresh
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SectionCalculate()
    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Exit Sub
      ' Make a list of all <forest> elements starting a section
      pdxSection = pdxCurrentFile.SelectNodes("//forest[@Section]")
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/SectionCalculate error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   LineItalic
  ' Goal:   Set a line to italic or reset from Italic
  ' History:
  ' 13-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub LineItalic(ByRef rtfThis As RichTextBox, ByVal intLine As Integer, ByVal bItalic As Boolean)
    Dim intPos As Integer       ' Original seleciton start
    Dim intFrom As Integer      ' New selection start
    Dim intTo As Integer        ' New selection end

    Try
      ' Remember the original selection start
      intPos = rtfThis.SelectionStart
      ' Retrieve the new start + end
      intFrom = arLeft(intLine)
      If (intLine < UBound(arLeft)) Then
        intTo = arLeft(intLine + 1) - 1
      Else
        intTo = rtfThis.TextLength - 1
      End If
      ' Set the text of this line to Italic
      With rtfThis
        .SelectionStart = intFrom
        .SelectionLength = intTo - intFrom
        ' Italicize or not?
        If (bItalic) Then
          ' SetItalic
          ' .SelectionFont = New Font(.SelectionFont, .SelectionFont.Style Or FontStyle.Italic)
          .SelectionFont = New Font(.Font, .Font.Style Or FontStyle.Italic)
        Else
          ' Reset Italic
          '.SelectionFont = New Font(.SelectionFont, .SelectionFont.Style And Not FontStyle.Italic)
          .SelectionFont = New Font(.Font, .Font.Style And Not FontStyle.Italic)
        End If
        ' Reset selection length
        .SelectionLength = 0
        .SelectionStart = intPos
        ' Reset Italic at any rate
        .SelectionFont = New Font(.SelectionFont, .SelectionFont.Style And Not FontStyle.Italic)
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/LineItalic error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   SectionDelete
  ' Goal:   Delete the section that is present before this line (if any)
  ' History:
  ' 13-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SectionDelete(ByRef rtfThis As RichTextBox)
    Dim ndxForest As XmlNode  ' The <forest> element starting the section

    Try
      ' Are we anywhere?
      If (loc_intLineNum < 0) OrElse (loc_intLineNum > arLeft.Length - 1) Then Exit Sub
      ' Get where we are
      ndxForest = GetForest(loc_intLineNum)
      If (ndxForest Is Nothing) Then Exit Sub
      ' Check if the appropriate attribute already exists
      If (ndxForest.Attributes("Section") Is Nothing) Then
        ' No section yet, so deleting is not possible
        Status("No section to be deleted")
      Else
        'There is a section, so delete it
        ndxForest.Attributes.Remove(ndxForest.Attributes("Section"))
        ' Reset this line's to Italic
        LineItalic(rtfThis, loc_intLineNum, False)
        ' Recalculate the sections
        SectionCalculate()
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/SectionDelete error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   GetClauseMood
  ' Goal:   Assuming [ndxThis] is a clause, get the mood of this clause
  ' History:
  ' 13-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetClauseMood(ByRef ndxThis As XmlNode) As String
    Dim ndxVfin As XmlNode = Nothing  ' Finite verb
    Dim ndxParent As XmlNode          ' Parent

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      If (ndxThis.Name <> "eTree") Then Return ""
      If (Not ndxThis.Attributes("Label").Value Like "*IP*") Then Return ""
      ' Check for presence of CP parent
      ndxParent = ndxThis.ParentNode
      If (ndxParent IsNot Nothing) Then
        If (ndxParent.Name = "eTree") Then
          ' Is this a CP question parent?
          If (ndxParent.Attributes("Label").Value Like "CP-QUE*") Then Return "Wh"
        End If
      End If
      ' Try get the main verb
      ndxVfin = ndxThis.SelectSingleNode("./child::eTree[tb:matches(@Label, 'AX*|BE*|MD*|HV*|V*')]", conTb)
      ' If there is no finite verb, the mood is unknown
      If (ndxVfin Is Nothing) Then Return "x"
      ' Check if the verb is an imperative
      If (DoLike(ndxVfin.Attributes("Label").Value, strVerbImv)) Then Return "Imv"
      ' Check if the verb is a subjunctive
      If (DoLike(ndxVfin.Attributes("Label").Value, strVerbSubj)) Then Return "Subj"
      ' Default: declarative mood
      Return "Decl"
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetClauseMood error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetIpAncestor
  ' Goal:   Get the IP ancestor of me
  ' History:
  ' 22-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetIpAncestor(ByRef ndxThis As XmlNode) As XmlNode
    Dim ndxNext As XmlNode  ' Working node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return Nothing
      ' Check if I am an <eTree>
      If (ndxThis.Name <> "eTree") Then Return Nothing
      ' Start with myself
      ndxNext = ndxThis
      ' Loop up
      While (Not ndxNext Is Nothing)
        ' Check this possibility
        If (ndxNext.Name = "eTree") Then
          ' Check its label
          If (ndxNext.Attributes("Label").Value Like "IP*") Then
            ' Return what we've found
            Return ndxNext
          End If
        End If
        ' Go one up
        ndxNext = ndxNext.ParentNode
      End While
      ' We have not found it
      Return Nothing
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/GetIpAncestor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   eTreeHandling
  ' Goal:   Process the handling of the forest passed on to me
  ' History:
  ' 12-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function eTreeHandling(ByRef ndxForest As XmlNode) As Boolean
    Dim ndxThis As XmlNode    ' Working node

    Try
      ' (3) The node should be the first under the forest
      ndxThis = ndxForest.SelectSingleNode("./descendant::eTree[1]")
      ' Double check
      If (ndxThis Is Nothing) Then Status("Could not find constituent under forest") : Return False
      '' (4) Adapt the <eTree>/@Id values from here on
      'AdaptEtreeId(ndxThis.Attributes("Id").Value)
      ' (7) Show this is handled properly
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modEditor/eTreeHandling error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   eTreeInsertLevel
  ' Goal:   Insert a node:
  '           eTree, eLeaf - between me and my parent
  '           forest       - between forest and remainder
  ' History:
  ' 03-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function eTreeInsertLevel(ByRef ndxThis As XmlNode, ByRef ndxNew As XmlNode) As Boolean
    Dim ndxChild As XmlNode = Nothing ' Working node
    Dim ndxList As XmlNodeList        ' List of children
    Dim intI As Integer               ' Counter

    Try
      ' Validate something is selected
      If (ndxThis IsNot Nothing) Then
        Select Case ndxThis.Name
          Case "eLeaf", "eTree"
            ' (1) Create a new <eTree> element
            ndxNew = CreateNewEtree(pdxCurrentFile)
            ' (2) Replace the parent's child with the new one
            ndxThis.ParentNode.ReplaceChild(ndxNew, ndxThis)
            ' (3) Set its (only) child
            ndxNew.PrependChild(ndxThis)
            ' (5) Get the appropriate values for @from and @to
            ndxNew.Attributes("from").Value = ndxThis.Attributes("from").Value
            ndxNew.Attributes("to").Value = ndxThis.Attributes("to").Value
            ' Return success
            Return True
          Case "forest"
            ' Insert a level between the <forest> node and the remainder
            ' (1) Create a new <eTree> element
            ndxNew = CreateNewEtree(pdxCurrentFile)
            ' (2) Prepare all the children of the forest parent
            ndxList = ndxThis.SelectNodes("./eTree")
            ' (3) Make sure the new node is the child of the forest
            ndxThis.PrependChild(ndxNew)
            For intI = 0 To ndxList.Count - 1
              ndxChild = ndxList(intI)
              ' Check this is not the new one
              If (ndxChild IsNot ndxNew) Then
                ' Replace the parent of this item
                ndxNew.AppendChild(ndxChild)
              End If
            Next intI
            ' Return success
            Return True
        End Select
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/eTreeInsertLevel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   eTreeMoveLeft
  ' Goal:   Move current node one sibling to the left
  ' History:
  ' 05-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function eTreeMoveLeft(ByRef ndxThis As XmlNode, ByRef ndxNew As XmlNode, _
                                ByRef litThis As LithiumControl) As Boolean
    Dim ndxParent As XmlNode  ' My parent
    Dim ndxSibling As XmlNode ' My sibling to the left

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Can only be <eTree>
      If (ndxThis.Name <> "eTree") Then Return False
      ' Try get parent
      ndxParent = ndxThis.ParentNode
      ' Validate
      If (ndxParent Is Nothing) Then Return False
      ' Get my own preceding sibling
      ndxSibling = ndxThis.SelectSingleNode("./preceding-sibling::eTree[1]")
      ' Validate
      If (ndxSibling Is Nothing) Then Return False
      ' Move myself
      ndxParent.InsertBefore(ndxThis, ndxSibling)
      ' Set the "new" node
      ndxNew = ndxThis
      ' Adapt the sentence, but don't re-calculate "org"
      eTreeSentence(ndxNew, ndxNew, bDoOrg:=False)
      '' Perform the action for the lithium control
      'If (Not litMove(litThis, ndxThis, "left")) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/eTreeMoveLeft error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   eTreeMoveRight
  ' Goal:   Move current node one sibling to the right
  ' History:
  ' 05-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function eTreeMoveRight(ByRef ndxThis As XmlNode, ByRef ndxNew As XmlNode, _
                                 ByRef litThis As LithiumControl) As Boolean
    Dim ndxParent As XmlNode  ' My parent
    Dim ndxSibling As XmlNode ' My sibling to the right

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Can only be <eTree>
      If (ndxThis.Name <> "eTree") Then Return False
      ' Try get parent
      ndxParent = ndxThis.ParentNode
      ' Validate
      If (ndxParent Is Nothing) Then Return False
      ' Get my own following sibling
      ndxSibling = ndxThis.SelectSingleNode("./following-sibling::eTree[1]")
      ' Validate
      If (ndxSibling Is Nothing) Then Return False
      ' Move myself
      ndxParent.InsertAfter(ndxThis, ndxSibling)
      ' Set the "new" node
      ndxNew = ndxThis
      ' Adapt the sentence, but don't re-calculate "org"
      eTreeSentence(ndxNew, ndxNew, bDoOrg:=False)
      '' Perform the action for the lithium control
      'If (Not litMove(litThis, ndxThis, "right")) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/eTreeMoveRight error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   eTreeMoveUp
  ' Goal:   Move current node one position up
  ' History:
  ' 05-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function eTreeMoveUp(ByRef ndxThis As XmlNode, ByRef ndxNew As XmlNode, _
                              ByRef litThis As LithiumControl) As Boolean
    Dim ndxParent As XmlNode  ' My parent
    Dim ndxGrandP As XmlNode  ' Grandparent
    Dim ndxLeafSelfLeft As XmlNode  ' Leftmost leaf under me
    Dim ndxLeafParRight As XmlNode  ' Rightmost leaf of my parent
    Dim intEtreeId As Integer       ' ID of <eTree> element

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' A forest can not move upwards
      If (ndxThis.Name = "forest") Then Return False
      ' Get my ID
      If (ndxThis.Name = "eTree") Then
        intEtreeId = ndxThis.Attributes("Id").Value
      Else
        intEtreeId = -1
      End If
      ' Try get parent
      ndxParent = ndxThis.ParentNode
      ' Validate
      If (ndxParent Is Nothing) Then Return False
      If (ndxParent.Name = "forest") Then Status("Cannot move constituent to become a <forest> sister!") : Return False
      ' Need to go one more upwards
      ndxGrandP = ndxParent.ParentNode
      ' Validate
      If (ndxGrandP Is Nothing) Then Return False
      ' Now check...
      Select Case ndxParent.Name
        'Case "eTree", "forest"
        Case "eTree"
          ' FInd out where I should go...
          ndxLeafSelfLeft = ndxThis.SelectSingleNode("./descendant-or-self::eLeaf[1]")
          ndxLeafParRight = ndxParent.SelectSingleNode("./descendant-or-self::eLeaf[parent::eTree/@Id != " & intEtreeId & " and @to !=0][last()]")
          If (ndxLeafParRight IsNot Nothing) AndAlso (ndxLeafSelfLeft IsNot Nothing) AndAlso _
             (ndxLeafSelfLeft.Attributes("from").Value >= ndxLeafParRight.Attributes("to").Value) Then
            ' Insert myself as child of grandparent, but before my parent
            ndxGrandP.InsertAfter(ndxThis, ndxParent)
          Else
            ' Insert myself as child of grandparent, but before my parent
            ndxGrandP.InsertBefore(ndxThis, ndxParent)
          End If
          '' Make myself the last child of this node
          'ndxParent.AppendChild(ndxThis)
          ndxNew = ndxThis
          ' Adapt the sentence, but don't re-calculate "org"
          eTreeSentence(ndxNew, ndxNew, bDoOrg:=False)
          '' Perform the action for the lithium control
          'If (Not litMove(litThis, ndxThis, "up")) Then Return False
          Return True
      End Select
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/eTreeMoveUp error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   eTreeJoinUnder
  ' Goal:   Join a node as first child under the one right (or left) to me:
  '           eTree   - I become the first child of my following-sibling
  '           eLeaf   - not possible: an <eLeaf> should only have one <eTree> parent
  '           forest  - not possible
  ' History:
  ' 03-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function eTreeJoinUnder(ByRef ndxThis As XmlNode, ByRef ndxNew As XmlNode, _
                                 ByVal bRight As Boolean) As Boolean

    Try
      ' --> I become the first child before any other children of a sibling node
      Select Case ndxThis.Name
        Case "eLeaf", "forest"
          ' This is not possible
          Return False
        Case "eTree"
          ' Get my next sibling
          If (bRight) Then
            ndxNew = ndxThis.SelectSingleNode("./following-sibling::eTree[1]")
          Else
            ndxNew = ndxThis.SelectSingleNode("./preceding-sibling::eTree[1]")
          End If
          ' Found anything?
          If (ndxNew IsNot Nothing) Then
            ' Attach me under there
            If (bRight) Then
              ' I become the first child
              ndxNew.PrependChild(ndxThis)
            Else
              ' I become the last child
              ndxNew.AppendChild(ndxThis)
            End If
            ' Adapt the sentence, but don't re-calculate "org"
            eTreeSentence(ndxNew, ndxNew, bDoOrg:=False)
            ' Return success
            Return True
          End If
      End Select
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/eTreeJoinUnderRight error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   eTreeSentence
  ' Goal:   Re-analyze a whole sentence in the following way:
  '         1. Based on the content of the [eLeaf] nodes:
  '            a. Determine the <seg> text
  '            b. Determine @from and @to for the [eLeaf] nodes
  '         2. Determine @from and @to for all the [eTree] nodes again
  ' History:
  ' 03-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function eTreeSentence(ByRef ndxThis As XmlNode, ByRef ndxNew As XmlNode, Optional ByVal bVerbose As Boolean = True, _
                                Optional ByVal bOldEnglish As Boolean = False, Optional ByVal bDoOrg As Boolean = True) As Boolean
    Dim ndxFor As XmlNode             ' My parent forest node
    Dim ndxChild As XmlNode = Nothing ' Working node
    Dim ndxList As XmlNodeList        ' List of children
    Dim ndxVern As XmlNode            ' Vernacular text line
    Dim ndxLeaf As XmlNode = Nothing  ' One working leaf
    Dim intI As Integer               ' Counter
    Dim intFrom As Integer = 0        ' Word starting point
    Dim intTo As Integer = 0          ' End of word
    Dim bNeedSpace As Boolean = False ' No space needed after this word
    Dim bChanged As Boolean = False   ' Whether anythying has in fact changed
    Dim strLine As String = ""        ' Text of this line

    Try
      ' Validate something is selected
      If (ndxThis Is Nothing) Then Return False
      ' Determine the parent forest node
      ndxFor = ndxThis.SelectSingleNode("./ancestor-or-self::forest[1]")
      If (ndxFor Is Nothing) Then Return False
      ' Need to recalculate the "org" text?
      If (bDoOrg) Then
        ' Get the vernacular text line
        ndxVern = ndxFor.SelectSingleNode("./child::div[@lang='org']/seg")
        If (ndxVern Is Nothing) Then Return False
        ' Get all the [eLeaf] children, but only if they have no CODE nor METADATA ancestor
        ndxList = ndxFor.SelectNodes(".//descendant::eLeaf[count(ancestor::eTree[tb:matches(@Label, '" & strNoText & "')])=0]", conTb)
        ' Walk all the children
        For intI = 0 To ndxList.Count - 1
          ' ============ DEBUG =========
          ' If (intI = 11) Then Stop
          ' ============================
          ' Process this <eLeaf>
          With ndxList(intI)
            ' Check if this <eLeaf> has the correct type
            If (.Attributes("Type").Value = "Punct") AndAlso (.Attributes("Text").Value Like "*[a-zA-Z]*") Then
              ' It must be of type "Vern" instead
              .Attributes("Type").Value = "Vern"
            End If
            Select Case .Attributes("Type").Value
              Case "Vern"
                ' Need to add a space?
                If (bNeedSpace) Then strLine &= " "
                ' Get the starting point of the word
                intFrom = strLine.Length
                ' Add word to the text of this line
                If (bOldEnglish) Then
                  strLine &= VernToEnglish(.Attributes("Text").Value)
                Else
                  strLine &= .Attributes("Text").Value
                End If
                ' Get the correct ending point of the word
                intTo = strLine.Length
                ' Normally each word should be followed by a space
                bNeedSpace = True
              Case "Punct"
                ' Are we supposed to add a space?
                If (bNeedSpace) Then
                  ' Check if this punctuation should be PRECEDED by a space
                  Select Case .Attributes("Text").Value
                    Case ":", ",", ".", "!", "?", ";", ">>"
                      ' A space may NOT precede this punctuation
                    Case "»"
                      ' A space may NOT precede this punctuation
                    Case "«", "<<"
                      ' A space must precede this punctuation
                      strLine &= " "
                    Case "'", """"
                      ' Check if a word is preceding or not
                      If (intI > 0) Then
                        ' We are not at the beginning...
                        If (ndxList(intI - 1).Attributes("Type").Value <> "Vern") Then
                          ' There is NO word preceding, so DO add a space
                          strLine &= " "
                        End If
                      End If
                    Case Else
                      ' In all other cases a space has to be added
                      strLine &= " "
                  End Select
                End If
                ' Get the starting point of the word
                intFrom = strLine.Length
                ' ========== DEBUG ============
                'If (.Attributes("Text").Value = "-m") Then
                '  Stop
                'End If
                ' =============================
                ' Add word to the text of this line
                strLine &= .Attributes("Text").Value
                ' Get the correct ending point of the word
                intTo = strLine.Length
                ' Check if this punctuation should be FOLLOWED by a space
                Select Case .Attributes("Text").Value
                  Case ":", ",", ".", "!", "?", ";", ">>"
                    ' A space must follow
                    bNeedSpace = True
                  Case "»"
                    ' A space should follow this punctuation
                    bNeedSpace = True
                  Case "«"
                    ' A space should not follow
                    bNeedSpace = False
                  Case "'", """"
                    ' Check if a word is preceding or not
                    If (intI > 0) Then
                      ' We are not at the beginning...
                      If (ndxList(intI - 1).Attributes("Type").Value = "Vern") Then
                        ' There is a word preceding, so DO add a space
                        bNeedSpace = True
                      End If
                    End If
                  Case Else
                    ' Reset spacing
                    bNeedSpace = False
                End Select
              Case "Star"
                ' A star item must contain at least a space
                intFrom = strLine.Length
                ' Add this space
                strLine &= " " : bNeedSpace = False
                ' Get the correct ending point of the word
                intTo = strLine.Length
              Case "Zero"
                ' Get the starting point of the word
                intFrom = strLine.Length : intTo = intFrom
            End Select
            ' Validate existence of from and to
            If (.Attributes("from") Is Nothing) Then AddAttribute(ndxList(intI), "from", "0")
            If (.Attributes("to") Is Nothing) Then AddAttribute(ndxList(intI), "to", "0")
            ' Adapt the start and end of the word
            intFrom += 1
            If (.Attributes("from").Value <> intFrom) Then .Attributes("from").Value = intFrom : bChanged = True
            If (.Attributes("to").Value <> intTo) Then .Attributes("to").Value = intTo : bChanged = True
          End With
        Next intI
        ' Adapt the sentence in the vernacular
        ndxVern.InnerText = strLine
        ' Make sure editor is set to dirty
        bEdtDirty = True
      End If
      ' Get all the <eTree> nodes
      ndxList = ndxFor.SelectNodes("./descendant::eTree")
      ' Treat them all
      For intI = 0 To ndxList.Count - 1
        ' Access this one
        With ndxList(intI)
          ' Determine their @from and @to values
          ndxLeaf = .SelectSingleNode("./descendant::eLeaf[1]")
          If (ndxLeaf IsNot Nothing) Then
            ' Double check
            If (ndxLeaf.Attributes("from") Is Nothing) Then AddXmlAttribute(pdxCurrentFile, ndxLeaf, "from", "0")
            ' Get the value
            intFrom = ndxLeaf.Attributes("from").Value
            ' Validate
            If (.Attributes("from") Is Nothing) Then
              AddAttribute(ndxList(intI), "from", intFrom)
            Else
              ' See if we need changing
              If (.Attributes("from").Value <> intFrom) Then .Attributes("from").Value = intFrom : bChanged = True
            End If
          End If
          ndxLeaf = .SelectSingleNode("./descendant::eLeaf[last()]")
          If (ndxLeaf IsNot Nothing) Then
            ' Double check
            If (ndxLeaf.Attributes("to") Is Nothing) Then AddXmlAttribute(pdxCurrentFile, ndxLeaf, "to", "0")
            ' Get the value
            intTo = ndxLeaf.Attributes("to").Value
            ' Validate
            If (.Attributes("to") Is Nothing) Then
              AddAttribute(ndxList(intI), "to", intTo)
            Else
              ' See if we need changing
              If (.Attributes("to").Value <> intTo) Then .Attributes("to").Value = intTo : bChanged = True
            End If
          End If
        End With
      Next intI
      ' We end with the same node we started with
      ' NO!! then we change it... ndxNew = ndxThis
      ' Give message to user
      If (bChanged) AndAlso (bVerbose) Then
        Status("Word positions in line " & ndxFor.Attributes("forestId").Value)
        'Else
        '  Logging("No changes were needed")
      End If
      ' Return success
      Return bChanged
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/eTreeSentence error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   eTreeGlueToPrec
  ' Goal:   Glue current sentence to the preceding one
  ' History:
  ' 05-03-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function eTreeGlueToPrec(ByRef ndxThFor As XmlNode, ByRef ndxNew As XmlNode) As Boolean
    Dim ndxAnc As XmlNode       ' Parent of <forest>
    Dim ndxPrFor As XmlNode     ' Preceding forest
    Dim ndxList As XmlNodeList  ' List of nodes to be transported
    Dim ndxWork As XmlNode      ' Working node
    Dim strLang As String       ' One div
    Dim strLocPre As String     ' Location prefix
    Dim intForId As Integer     ' ID of <prForest>
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (ndxThFor Is Nothing) OrElse (ndxThFor.Name <> "forest") Then Return False
      ' Ask permission
      Select Case MsgBox("There is no <restore>, so you might want to save first" & vbCrLf & _
                         "Do you really want to glue the current sentence to the preceding one?" & vbCrLf, _
                         MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.No, MsgBoxResult.Cancel
          Status("Aborted")
          Return False
      End Select
      ' Get the preceding forest
      ndxPrFor = ndxThFor.SelectSingleNode("./preceding-sibling::forest[1]")
      If (ndxPrFor Is Nothing) Then Return False
      intForId = ndxPrFor.Attributes("forestId").Value
      ' Get location prefix
      strLocPre = IO.Path.GetFileNameWithoutExtension(ndxPrFor.Attributes("Location").Value)
      ' Get forestgroup
      ndxAnc = ndxThFor.ParentNode : If (ndxAnc Is Nothing) OrElse (ndxAnc.Name <> "forestGrp") Then Status("No <forestGrp> ancestor found") : Return False
      ' Collect all that needs to be transported
      ndxList = ndxThFor.SelectNodes("./child::eTree")
      For intI = 0 To ndxList.Count - 1
        ' Move this node from current forest to preceding forest
        ndxPrFor.AppendChild(ndxList(intI))
      Next intI
      ' Treat all the <div> parts
      ndxList = ndxThFor.SelectNodes("./child::div")
      For intI = 0 To ndxList.Count - 1
        ' Identify language
        strLang = ndxList(intI).Attributes("lang").Value
        ' Find target in preceding forest
        ndxWork = ndxPrFor.SelectSingleNode("./child::div[@lang='" & strLang & "']")
        If (ndxWork IsNot Nothing) Then
          ' Glue the descriptions
          ndxWork.FirstChild.InnerText = ndxWork.FirstChild.InnerText & " " & ndxList(intI).FirstChild.InnerText
        End If
      Next intI
      ' Process the current and the following sentence again
      eTreeSentence(ndxPrFor, ndxNew)
      ' Remove the current forest
      ndxThFor.RemoveAll()
      ndxAnc.RemoveChild(ndxThFor)
      ' Repare @forestId values
      ndxWork = ndxAnc.SelectSingleNode("./child::forest[@forestId='" & intForId & "']")
      ndxWork = ndxWork.NextSibling
      While (ndxWork IsNot Nothing)
        intForId += 1
        ndxWork.Attributes("forestId").Value = intForId
        ndxWork.Attributes("Location").Value = strLocPre & "." & intForId
        ndxWork = ndxWork.SelectSingleNode("./following-sibling::forest[1]")
      End While
      ' Redo initialisation
      If (Not InitCurrentFile()) Then Return False
      SectionCalculate()
      ' Load this section again
      frmMain.GoShowSection(intCurrentSection)
      ' Go to the correct forest id
      GotoNode(ndxPrFor)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/eTreeGlueToPrec error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   eTreeDoubleForest
  ' Goal:   Add a copy of the current forest *after* this one
  ' History:
  ' 23-12-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function eTreeDoubleForest(ByRef ndxThFor As XmlNode, _
                                  ByRef ndxNew As XmlNode) As Boolean
    Dim ndxThis As XmlNode      ' <div> section of the new <forest>
    Dim ndxAnc As XmlNode       ' Parent of <forest>
    Dim ndxNxFor As XmlNode     ' Following (new) <forest>
    Dim strLocPre As String     ' Location prefix
    Dim bChanged As Boolean     ' Changes
    Dim intForId As Integer     ' ID of forest

    Try
      ' Validate
      If (ndxThFor Is Nothing) OrElse (ndxThFor.Name <> "forest") Then Status("Could not identify current forest") : Return False
      ' Get the current forest id
      intForId = ndxThFor.Attributes("forestId").Value
      ' Get forestgroup
      ndxAnc = ndxThFor.ParentNode : If (ndxAnc Is Nothing) OrElse (ndxAnc.Name <> "forestGrp") Then Status("No <forestGrp> ancestor found") : Return False
      ' Create a new forest node
      SetXmlDocument(pdxCurrentFile)
      ndxNxFor = AddXmlChildAfter(ndxAnc, ndxThFor, "forest", "forestId", intForId + 1, "attribute", _
                                  "TextId", ndxThFor.Attributes("TextId").Value, "attribute", _
                                  "File", ndxThFor.Attributes("File").Value, "attribute", _
                                  "Location", ndxThFor.Attributes("TextId").Value & "." & intForId + 1, "attribute")
      ' Copy the content of the current forest to the new one
      ndxNxFor.InnerXml = ndxThFor.InnerXml
      ' Process the current and the following sentence again
      bChanged = eTreeSentence(ndxThFor, ndxNew)
      bChanged = bChanged Or eTreeSentence(ndxNxFor, ndxNew)
      ' Get location prefix
      strLocPre = IO.Path.GetFileNameWithoutExtension(ndxThFor.Attributes("Location").Value)
      ' Do the @forestId counts from ndxNxFor onwards
      ndxThis = ndxNxFor
      While (ndxThis IsNot Nothing)
        intForId += 1
        ndxThis.Attributes("forestId").Value = intForId
        ndxThis.Attributes("Location").Value = strLocPre & "." & intForId
        ndxThis = ndxThis.SelectSingleNode("./following-sibling::forest")
        bChanged = bChanged Or eTreeSentence(ndxThis, ndxNew)
      End While
      ' Redo initialisation
      If (Not InitCurrentFile()) Then Status("Could not re-initialise file") : Return False
      SectionCalculate()
      ' Load this section again
      frmMain.GoShowSection(intCurrentSection)
      ' Go to the correct forest id
      GotoNode(ndxThFor)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/eTreeDoubleForest error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   eTreeCutForest
  ' Goal:   Cut forest starting from <ndxThis>
  ' History:
  ' 22-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function eTreeCutForest(ByRef ndxThis As XmlNode, ByRef ndxThFor As XmlNode, _
                                  ByRef ndxNew As XmlNode) As Boolean
    Dim ndxAnc As XmlNode       ' Parent of <forest>
    Dim ndxNxFor As XmlNode     ' Following (new) <forest>
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim ndxDiv As XmlNode       ' <div> section of the new <forest>
    Dim strSent As String       ' Sentence to be split
    Dim strLang As String       ' Language
    Dim strLocPre As String     ' Location prefix
    Dim colDiv As New StringColl ' Split collection
    Dim bChanged As Boolean     ' Changes
    Dim intForId As Integer     ' ID of forest
    Dim intCutFrom As Integer   ' value of @from attribute to start cutting from
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (ndxThis.Name <> "eTree") Then Status("An <eTree> needs to be selected") : Return False
      If (ndxThFor Is Nothing) OrElse (ndxThFor.Name <> "forest") Then Status("Could not identify current forest") : Return False
      ' Get the cutfrom value
      intCutFrom = ndxThis.Attributes("from").Value
      ' Ask permission
      Select Case MsgBox("There is no <restore>, so you might want to save first" & vbCrLf & _
                         "Do you really want to cut the sentence here?" & vbCrLf, _
                         MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.No, MsgBoxResult.Cancel
          Status("Aborted")
          Return False
      End Select
      ' Get the current forest id
      intForId = ndxThFor.Attributes("forestId").Value
      ' Get forestgroup
      ndxAnc = ndxThFor.ParentNode : If (ndxAnc Is Nothing) OrElse (ndxAnc.Name <> "forestGrp") Then Status("No <forestGrp> ancestor found") : Return False
      ' Create a new forest node
      SetXmlDocument(pdxCurrentFile)
      ndxNxFor = AddXmlChildAfter(ndxAnc, ndxThFor, "forest", "forestId", intForId + 1, "attribute", _
                                  "TextId", ndxThFor.Attributes("TextId").Value, "attribute", _
                                  "File", ndxThFor.Attributes("File").Value, "attribute", _
                                  "Location", ndxThFor.Attributes("TextId").Value & "." & intForId + 1, "attribute")
      ' Process the different <div> sections under the forest
      ndxList = ndxThFor.SelectNodes("./child::div")
      For intI = 0 To ndxList.Count - 1
        ' Do not ask for the "org" section -- that will be handled automatically
        strLang = ndxList(intI).Attributes("lang").Value
        strSent = ndxList(intI).InnerText
        ' colDiv.Add(strLang, strSent)
        ' Add appropriate <div> section under the new forest
        ndxDiv = AddXmlChild(ndxNxFor, "div", "lang", strLang, "attribute", _
                                              "seg", strSent, "child")
      Next intI
      ' Get the next node
      ' nd() ' xNext = ndxThis.SelectSingleNode("./following-sibling::eTree")
      ' If (ndxNext Is Nothing) Then Return False
      ' Find all the highest <eTree> nodes that need to move from the current to the next forest
      ndxList = ndxThFor.SelectNodes("./descendant::eTree[@from >= " & intCutFrom & _
                                     " and not(descendant::eTree[@from < " & intCutFrom & "])" & _
                                     " and ( (parent::forest) or (parent::eTree[@from < " & intCutFrom & " ]))]")
      ' Move these nodes straight under the following forest
      For intI = 0 To ndxList.Count - 1
        ndxNxFor.AppendChild(ndxList(intI))
      Next intI
      ' Process the current and the following sentence again
      bChanged = eTreeSentence(ndxThFor, ndxNew)
      bChanged = bChanged Or eTreeSentence(ndxNxFor, ndxNew)
      ' Get location prefix
      strLocPre = IO.Path.GetFileNameWithoutExtension(ndxThFor.Attributes("Location").Value)
      ' Do the @forestId counts from ndxNxFor onwards
      ndxThis = ndxNxFor
      While (ndxThis IsNot Nothing)
        intForId += 1
        ndxThis.Attributes("forestId").Value = intForId
        ndxThis.Attributes("Location").Value = strLocPre & "." & intForId
        ndxThis = ndxThis.SelectSingleNode("./following-sibling::forest")
      End While
      ' Redo initialisation
      If (Not InitCurrentFile()) Then Status("Could not re-initialise file") : Return False
      SectionCalculate()
      ' Load this section again
      frmMain.GoShowSection(intCurrentSection)
      ' Go to the correct forest id
      GotoNode(ndxThFor)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/eTreeCutForest error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   eTreeDelete
  ' Goal:   Delete a node:
  '           eTree         - replace me as child by all my children
  '           forest, eLeaf - not possible
  ' History:
  ' 03-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function eTreeDelete(ByRef ndxThis As XmlNode, ByRef ndxNew As XmlNode, Optional ByVal bAsk As Boolean = True) As Boolean
    Dim ndxChild As XmlNode = Nothing ' Working node
    Dim ndxList As XmlNodeList        ' List of children
    Dim intI As Integer               ' Counter

    Try
      ' Validate something is selected
      If (ndxThis IsNot Nothing) Then
        Select Case ndxThis.Name
          Case "forest"
            ' Impossible
            Return False
          Case "eLeaf"
            If (bAsk) Then
              ' Ask permission
              Select Case MsgBox("Do you really want to delete this endnode?", MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.No, MsgBoxResult.Cancel
                  Return False
              End Select
            End If
            ' (1) Get my parent
            ndxNew = ndxThis.ParentNode
            ' (2) Remove child of parent
            ndxNew.RemoveChild(ndxThis)
            ' (3) Remove myself
            ndxThis.RemoveAll()
            ' Adapt the sentence
            eTreeSentence(ndxNew, ndxNew)
            ' Return success
            Return True
          Case "eTree"
            If (bAsk) Then
              ' Ask permission
              Select Case MsgBox("Do you really want to delete this constituent?", MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.No, MsgBoxResult.Cancel
                  Return False
              End Select
            End If
            ' (1) Take note of all my current children
            ndxList = ndxThis.ChildNodes
            ' (2) Take note of who is going to be the new parent
            ndxNew = ndxThis.ParentNode
            ' (3) Are there any children to take care of?
            If (ndxList.Count = 0) Then
              ' If I have no children, then i can still be deleted
              ndxNew.RemoveChild(ndxThis)
              ndxThis.RemoveAll()
              Return True
            Else
              ' I have children: get the first one
              ndxChild = ndxList(0)
              ' If this child is an <eLeaf>, then stop here!
              If (ndxChild.Name = "eLeaf") Then Return False
              ' Replace me in my parent's place
              ndxThis.ParentNode.ReplaceChild(ndxChild, ndxThis)
              ' Work the remaining children
              For intI = 1 To ndxList.Count - 1
                ' Insert this node after [ndxChild]
                ndxChild.ParentNode.InsertAfter(ndxList(intI), ndxList(intI - 1))
              Next intI
            End If
            ' (4) Remove the dis-entangled me
            ndxThis.RemoveAll()
            ' Adapt the sentence, but don't re-calculate "org"
            eTreeSentence(ndxNew, ndxNew, bDoOrg:=False)
            ' Return success
            Return True
        End Select
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/eTreeDelete error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   eTreeAdd
  ' Goal:   Add a node:
  '           eTree         - replace me as child by all my children
  '           forest, eLeaf - not possible
  '         strType can be:
  '           right         - Add a sibling to my right
  '           left          - Add a sibling to my left
  '           child         - Create/add a child <eTree> node
  '           eLeaf         - Create an <eLeaf> node under me (if there is none)
  ' History:
  ' 03-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function eTreeAdd(ByRef ndxThis As XmlNode, ByRef ndxNew As XmlNode, ByVal strType As String) As Boolean
    Dim ndxWork As XmlNode = Nothing  ' Working node

    Try
      ' Validate something is selected
      If (ndxThis IsNot Nothing) Then
        Select Case ndxThis.Name
          Case "eLeaf", "forest"
            ' Impossible
            Return False
          Case "eTree"
            ' Check what is our task
            Select Case strType
              Case "right", "Right" ' Add a sibling <eTree> to my right
                ' (1) Create a new <eTree> element
                ndxNew = CreateNewEtree(pdxCurrentFile)
                ' (2) Get my parent
                ndxWork = ndxThis.ParentNode
                If (ndxWork IsNot Nothing) Then
                  ' Insert the new node as child after me
                  ndxWork.InsertAfter(ndxNew, ndxThis)
                  '' Adapt the <eTree>/@Id values from [ndxThis] onwards
                  'AdaptEtreeId(ndxThis.Attributes("Id").Value)
                End If
                ' Adapt the sentence, but don't re-calculate "org"
                eTreeSentence(ndxNew, ndxNew, bDoOrg:=False)
                ' Return success
                Return True
              Case "left", "Left"   ' Add a sibling <eTree> to my left
                ' (1) Create a new <eTree> element
                ndxNew = CreateNewEtree(pdxCurrentFile)
                ' (2) Get my parent
                ndxWork = ndxThis.ParentNode
                If (ndxWork IsNot Nothing) Then
                  ' Insert the new node as child before me
                  ndxWork.InsertBefore(ndxNew, ndxThis)
                  '' Go to the first <eTree> child of [ndxWork]
                  'ndxWork = ndxWork.SelectSingleNode("./child::eTree[1]")
                  'If (ndxWork IsNot Nothing) Then
                  '  ' Adapt the <eTree>/@Id values from [ndxWork] onwards
                  '  AdaptEtreeId(ndxWork.Attributes("Id").Value)
                  'End If
                End If
                ' Adapt the sentence, but don't re-calculate "org"
                eTreeSentence(ndxNew, ndxNew, bDoOrg:=False)
                ' Return success
                Return True
              Case "child", "Child", "childlast" ' Add a child <eTree> under me
                ' (1) Create a new <eTree> element
                ndxNew = CreateNewEtree(pdxCurrentFile)
                ' (2) Add the <eTree> child under me
                ndxThis.AppendChild(ndxNew)
                '' Adapt the <eTree>/@Id values from [ndxThis] onwards
                'AdaptEtreeId(ndxThis.Attributes("Id").Value)
                ' Adapt the sentence, but don't re-calculate "org"
                eTreeSentence(ndxNew, ndxNew, bDoOrg:=False)
                ' Return success
                Return True
              Case "firstchild", "childfirst" ' Add a child <eTree> under me as the first one
                ' (1) Create a new <eTree> element
                ndxNew = CreateNewEtree(pdxCurrentFile)
                ' (2) Add the <eTree> child under me
                ndxThis.PrependChild(ndxNew)
                '' Adapt the <eTree>/@Id values from [ndxThis] onwards
                'AdaptEtreeId(ndxThis.Attributes("Id").Value)
                ' Adapt the sentence, but don't re-calculate "org"
                eTreeSentence(ndxNew, ndxNew, bDoOrg:=False)
                ' Return success
                Return True
              Case "eLeaf", "eleaf", "leaf", "endnode"
                ' (1) Check: do I have any <eLeaf> or <eTree> children?
                If (ndxThis.SelectSingleNode("./child::eLeaf") Is Nothing) AndAlso _
                   (ndxThis.SelectSingleNode("./child::eTree") Is Nothing) Then
                  ' (1) Create a new <eLeaf> element
                  ndxNew = CreateNewEleaf(pdxCurrentFile)
                  ' (2) Add it under me
                  ndxThis.AppendChild(ndxNew)
                  ' Re-do the sentence, including "org"
                  eTreeSentence(ndxThis, ndxNew)
                  ' Return success
                  Return True
                End If
              Case Else
                Return False
            End Select
        End Select
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/eTreeAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   eTreeDoLabel
  ' Goal:   Get a new label for me
  ' History:
  ' 05-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function eTreeDoLabel(ByRef ndxThis As XmlNode, ByVal strLabel As String, ByRef bHandled As Boolean) As Boolean
    Try
      ' Validate and initialize
      If (ndxThis Is Nothing) Then Return False
      If (strLabel = "") Then Return False
      bHandled = False
      ' Add the new label at this point in the tree
      Select Case ndxThis.Name
        Case "eTree"
          ' Change the label and show we handled it
          ndxThis.Attributes("Label").Value = strLabel : bHandled = True
        Case "eLeaf"
          ' Change the label and show we handled it
          ndxThis.Attributes("Text").Value = strLabel : bHandled = True
          ' Check the type of node
          If (Strings.Left(strLabel, 1) = "*") Then
            ndxThis.Attributes("Type").Value = "Star"
          ElseIf (strLabel = "0") Then
            ndxThis.Attributes("Type").Value = "Zero"
            'ElseIf (DoLike(strLabel, ",|.|<|>|;|:|'|\|[[]|]|{|}|=|+|-|_|(|)|[*]|&|[^]|%|$|#|@|!")) Then
          ElseIf (InStr(",.<>;:'\[]{}=+-_()*&^%$#@?!«»`", strLabel) > 0) Then
            ndxThis.Attributes("Type").Value = "Punct"
          Else
            ndxThis.Attributes("Type").Value = "Vern"
          End If
        Case "forest"
          ' Change the label and show we handled it
          ndxThis.Attributes("Location").Value = strLabel : bHandled = True
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modEditor/eTreeDoLabel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SelectedTreeNode
  ' Goal:   Get the currently selected node in the litTree
  ' History:
  ' 14-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function SelectedTreeNode(ByRef litThis As Lithium.LithiumControl, ByRef strLabel As String, _
                                   ByRef pntThis As Point) As XmlNode
    Dim ndxThis As XmlNode = Nothing    ' one node
    Dim ndxForest As XmlNode = Nothing  ' One forest
    Dim ndxNew As XmlNode = Nothing     ' one node
    Dim strNodeName As String           ' Name of selected node
    Dim strNodeId As String             ' id of selected node
    Dim intId As Integer                ' Selected id

    Try
      ' Validate
      intId = litThis.SelectedId
      If (intId < 0) OrElse (intId >= litThis.Shapes.Count) Then Status("No tree element selected") : Return Nothing
      ' Where are we?
      With litThis.Shapes(intId)
        strNodeName = .NodeName
        strNodeId = .NodeId
        pntThis = .Location
      End With
      ' Initialise
      strLabel = ""
      ' Get the current forest
      If (Not GetCurrentForest(ndxForest, ndxThis)) Then Return Nothing
      Select Case strNodeName
        Case "eTree"
          ndxThis = ndxForest.SelectSingleNode("./descendant::eTree[@Id='" & strNodeId & "']")
          If (Not ndxThis Is Nothing) Then strLabel = ndxThis.Attributes("Label").Value
        Case "eLeaf"
          ndxThis = ndxForest.SelectSingleNode("./descendant::eTree[@Id='" & strNodeId & "']/child::eLeaf[1]")
          If (Not ndxThis Is Nothing) Then strLabel = ndxThis.Attributes("Text").Value
        Case "forest"
          ndxThis = ndxForest
          If (Not ndxThis Is Nothing) Then strLabel = ndxThis.Attributes("Location").Value
        Case Else
          ' Cannot handle this kind of node
          Status("Cannot handle nodes of type: " & strNodeName)
          ' Return positively anyway
          Return Nothing
      End Select
      ' Return the node
      Return ndxThis
    Catch ex As Exception
      ' Warn the user
      HandleErr("modEditor/SelectedTreeNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   InitTreeStack
  ' Goal:   Initialise the undo treestack
  ' History:
  ' 21-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function InitTreeStack() As Boolean
    Try
      loc_colTreeEdit.Clear()
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modEditor/InitTreeStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TreeStackSize
  ' Goal:   Get the size of the tree stack
  ' History:
  ' 21-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TreeStackSize() As Integer
    Return loc_colTreeEdit.Count
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   AddTreeStack
  ' Goal:   Add the current FOREST node to the tree-stack
  ' History:
  ' 21-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function AddTreeStack(Optional ByVal strOperation As String = "-") As Boolean
    Dim ndxFor As XmlNode = Nothing ' Current forest
    Dim ndxCns As XmlNode = Nothing ' Current node
    Dim strText As String           ' Text of current forest
    Dim strSel As String            ' Label of selected node
    Dim intForId As Integer         ' ID of forest

    Try
      ' Get current forest
      If (Not GetCurrentForest(ndxFor, ndxCns)) Then Return "(no forest)"
      ' Validate
      If (ndxFor Is Nothing) Then Return False
      ' Get text of forest
      strText = ndxFor.OuterXml
      intForId = ndxFor.Attributes("forestId").Value
      ' Get currently selected tree node
      ndxCns = ndxCurrentTreeNode
      strSel = ""
      If (ndxCns IsNot Nothing) Then
        ' Action depends on the type of tree
        Select Case ndxCns.Name
          Case "forest"
            strSel = "[forest]"
          Case "eTree"
            strSel = "[" & ndxCns.Attributes("Label").Value & "]"
          Case "eLeaf"
            strSel = "[endnode]"
        End Select
      End If
      ' Add text to stack
      loc_colTreeEdit.Add(strText, intForId & ": " & strSel & strOperation)
      ' Update combobox
      UpdateTreeStack()
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modEditor/AddTreeStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   UpdateTreeStack
  ' Goal:   Update the combobox showing the treestack
  ' History:
  ' 21-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function UpdateTreeStack() As Boolean
    Dim arItems(0) As String        ' List of items for combobox

    Try
      ' Check if tree combobox should be shown
      If (frmMain.TabControl1.SelectedTab Is frmMain.tpTree) Then
        ' Show we are busy
        bTreeBusy = True
        ' Is there anything?
        If (loc_colTreeEdit.Count = 0) Then
          frmMain.cboSyntStack.Visible = False
        Else
          ' Adapt and show the combobox
          frmMain.cboSyntStack.Visible = True
          ' Fill the CBO afresh
          With frmMain.cboSyntStack
            .Items.Clear()
            If (Not ShowTreeStack(arItems)) Then bTreeBusy = False : Return False
            ' Is there anything to add?
            .Items.AddRange(arItems)
            .SelectedIndex = .Items.Count - 1
          End With
        End If
        ' We are no longer busy
        bTreeBusy = False
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modEditor/UpdateTreeStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   ShowTreeStack
  ' Goal:   Show what is inside the treestack - reversed, with the last operation in [0]
  ' History:
  ' 21-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function ShowTreeStack(ByRef arList() As String) As Boolean
    Dim intSize As Integer = loc_colTreeEdit.Count
    Dim intI As Integer   ' Counter

    Try
      ' Validate
      If (intSize = 0) Then ReDim arList(0) : Return True
      ' Make room
      ReDim arList(0 To intSize - 1)
      ' Add elements
      For intI = intSize - 1 To 0 Step -1
        arList(intI) = loc_colTreeEdit.Exmp(intI)
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modEditor/ShowTreeStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   PopTreeStack
  ' Goal:   Pop the last element from the tree stack and let it replace the correct forest
  ' History:
  ' 21-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function PopTreeStack(Optional ByVal intItem As Integer = -1) As Boolean
    Dim ndxFor As XmlNode = Nothing ' Current forest
    Dim ndxCns As XmlNode = Nothing ' Current node
    Dim strText As String = ""      ' Text from stack
    Dim pdxThis As New XmlDocument  ' For reading purposes
    Dim intForestId As Integer      ' ID of the forest we have in the stack
    Dim intNum As Integer = 1       ' Number of items I want to have deleted
    Dim intI As Integer             ' Counter

    Try
      ' Busy?
      If (bTreeBusy) Then Return True
      ' Validate
      If (loc_colTreeEdit.Count = 0) Then Status("Nothing can be undone") : Return False
      ' Get the current forest 
      If (Not GetCurrentForest(ndxFor, ndxCns)) Then Return "(no forest)"
      ' Calculate the number of items to be undone
      If (intItem >= 0) Then
        intNum = loc_colTreeEdit.Count - intItem
      End If
      ' Perform all undo's
      For intI = 1 To intNum
        ' Get the last element from the stack
        strText = loc_colTreeEdit.Item(loc_colTreeEdit.Count - 1)
        ' Check the forestId of this element
        pdxThis.LoadXml(strText)
        intForestId = pdxThis.SelectSingleNode("./descendant::forest").Attributes("forestId").Value
        If (intForestId <> ndxFor.Attributes("forestId").Value) Then
          ' Ask user permission to continue
          Select Case MsgBox("You are currently in line " & ndxFor.Attributes("forestId").Value & _
                             ", but the undo operation concerns line " & intForestId & "." & vbCrLf & _
                             "Would you like to continue this undo operation?", vbYesNoCancel)
            Case MsgBoxResult.No, MsgBoxResult.Cancel
              Status("Undo has been aborted")
              Return False
          End Select
          ' Find the forest that needs to be undone
          ndxFor = GetForestWithId(intForestId)
          ' Validate
          If (ndxFor Is Nothing) Then Logging("Could not find forest [" & intForestId & "]") : Return False
        End If
        ' Replace the forest text
        ndxFor.InnerXml = pdxThis.SelectSingleNode("./descendant::forest").InnerXml
        ' Delete the element from the stack
        loc_colTreeEdit.DelItem(loc_colTreeEdit.Count - 1)
      Next intI
      ' Update combobox
      UpdateTreeStack()
      ' Make sure the correct node is being shown again
      ShowTree(frmMain.litTree)
      ' Make sure events are passed
      Application.DoEvents()
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modEditor/PopTreeStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
End Module
