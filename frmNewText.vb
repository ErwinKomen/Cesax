Public Class frmNewText
  ' ======================== LOCAL VARIABLES ===========================================
  Private strFileName As String = ""    ' Name of the file that has been created
  Private strDirName As String = ""     ' Directory where to store the file
  Private loc_strTitle As String = ""   ' Meta info: title
  Private loc_strDistr As String = ""   ' Meta info: distributor
  Private loc_strSource As String = ""  ' Meta info: source
  Private loc_strAuthor As String = ""  ' Meta info: author
  Private loc_strCreaOrg As String = "" ' Meta info: creation of original
  Private loc_strCreaMan As String = "" ' Meta info: creation of manuscript
  Private loc_strSubType As String = "" ' Meta info: title
  Private loc_strGenre As String = ""   ' Meta info: title
  Private loc_strEditor As String = ""  ' Meta info: title
  Private loc_strLngId As String = ""   ' Meta info: title
  Private loc_strLngName As String = "" ' Meta info: title
  Private loc_strText As String = ""    ' Meta info: title
  Private arHdType() As HeaderInfo
  ' ======================== PROPERTIES ================================================
  ' Allow getting and setting of the filename
  Public Property FileName() As String
    Get
      Return strFileName
    End Get
    Set(ByVal value As String)
      strFileName = value
    End Set
  End Property
  Public Property Directory() As String
    Get
      Return strDirName
    End Get
    Set(ByVal value As String)
      strDirName = value
      ' Also set the value in the textbox
      Me.tbGenDirName.Text = strDirName
    End Set
  End Property

  '---------------------------------------------------------------------------------------------------------
  ' Name:       cmdOk_Click()
  ' Goal:       Create the document and return to the caller
  ' History:
  ' 22-05-2015  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub cmdOk_Click(sender As System.Object, e As System.EventArgs) Handles cmdOk.Click
    Dim strFile As String     ' File name
    Dim strTempFile As String ' Temporary file
    Dim strValue As String    ' Value
    Dim intI As Integer       ' Counter
    Dim pdxConv As Xml.XmlDocument

    Try
      ' Validate file name
      If (Me.tbGenFileName.Text.Trim = "") Then Status("Please provide a name for the text first") : Exit Sub
      ' Get the file name
      strFile = Me.strFileName
      ' Check if the destination file exists already
      If (IO.File.Exists(strFile)) Then
        ' Ask user whether we may overwrite or not
        Dim bOkay As Boolean = False
        While (Not bOkay)
          Select Case MsgBox("The chosen filename already exists: [" & strFile & "]" & vbCrLf & _
                             "Would you like to overwrite it?", MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.Yes
              bOkay = True
            Case MsgBoxResult.No
              ' Exit gracefully and allow user to change the file name
              Status("Please change the file name")
              Exit Sub
            Case MsgBoxResult.Cancel
              ' We leave this box
              Me.DialogResult = Windows.Forms.DialogResult.Cancel
              Me.Close()
          End Select
        End While
      End If
      ' Check if there is a text
      If (Me.tbGenText.Text.Trim = "") Then Status("Please provide a text") : Exit Sub
      ' Store the text as a file preliminarily
      strTempFile = Me.strDirName & "\temp.txt"
      IO.File.WriteAllText(strTempFile, Me.tbGenText.Text, System.Text.Encoding.UTF8)
      ' Convert the text
      If (Not ConvertOneTxtToPsdx(strTempFile, strFile, loc_strLngId)) Then Status("Could not convert this text") : Exit Sub
      ' Signal creation
      Logging("Psdx text has been created: " & strFile)
      ' Add the header information
      pdxConv = New Xml.XmlDocument
      pdxConv.Load(strFile)
      For intI = 0 To arHdType.Length - 1
        With arHdType(intI)
          ' Debug.Print(arHdType(intI).Box Is Nothing)
          If (.Box Is Nothing) Then
            strValue = .Rich.Text
          Else
            strValue = .Box.Text
          End If
          Debug.Print("Header add: [" & .Type & ", " & strValue & ", " & .SubType & "]")
          If (Not DoAddFileDesc(pdxConv, .Type, strValue, .SubType)) Then Status("Unable to add [" & .Type & .SubType & "]") : Exit Sub
        End With
      Next intI
      pdxConv.Save(strFile)
      ' Signal header
      Logging("Header info added: " & strFile)

      ' Perform tokenization
      If (Not OneTokenizePsdx(strFile)) Then Status("Could not tokenize this text") : Exit Sub
      ' Signal completion
      Logging("Psdx text has been created and tokenized: " & strFile)
      ' Give the correct dialogue result
      Me.DialogResult = Windows.Forms.DialogResult.OK
      Me.Close()
    Catch ex As Exception
      ' Show error
      HandleErr("frmNewText/cmdOk error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       cmdCancel_Click()
  ' Goal:       Return to caller and do not save results
  ' History:
  ' 22-05-2015  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       cmdCancel_Click()
  ' Goal:       Return to caller and do not save results
  ' History:
  ' 22-05-2015  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub cmdGenDir_Click(sender As System.Object, e As System.EventArgs) Handles cmdGenDir.Click
    Dim strSelDir As String = ""  ' Selected directory

    Try
      ' Retrieve what we have right now
      strSelDir = strDirName
      ' Try get directory
      If (GetDirName(Me.FolderBrowserDialog1, strSelDir, _
                     "Select the directory where the PSDX text should be stored", strSelDir)) Then
        ' Replace what is shown
        Me.tbGenDirName.Text = strSelDir
        Me.strDirName = strSelDir
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmNewText/GenDir error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       tbGenFileName_TextChanged() ... and others!
  ' Goal:       Process changes in the information
  ' History:
  ' 22-05-2015  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub tbGenFileName_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbGenFileName.TextChanged
    Dim strFile As String

    Try
      ' Make sure no spaces are allowed
      strFile = Trim(Me.tbGenFileName.Text.Replace(" ", ""))
      ' Possibly append extension
      If (InStr(strFile, ".") = 0) Then strFile &= ".psdx"
      ' Prepend directory
      strFile = Me.tbGenDirName.Text & "\" & strFile
      Me.strFileName = strFile
    Catch ex As Exception
      ' Show error
      HandleErr("frmNewText/GenFileName error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub tbGenTitle_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbGenTitle.TextChanged
    Me.loc_strTitle = Me.tbGenTitle.Text
  End Sub
  Private Sub tbGenDistributor_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbGenDistributor.TextChanged
    Me.loc_strDistr = Me.tbGenDistributor.Text
  End Sub
  Private Sub tbGenSource_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbGenSource.TextChanged
    Me.loc_strSource = Me.tbGenSource.Text
  End Sub
  Private Sub tbGenAuthor_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbGenAuthor.TextChanged
    Me.loc_strAuthor = Me.tbGenAuthor.Text
  End Sub
  Private Sub tbGenCreaOrig_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbGenCreaOrig.TextChanged
    Me.loc_strCreaOrg = Me.tbGenCreaOrig.Text
  End Sub
  Private Sub tbGenCreaManu_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbGenCreaManu.TextChanged
    Me.loc_strCreaMan = Me.tbGenCreaManu.Text
  End Sub
  Private Sub tbGenSubType_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbGenSubType.TextChanged
    Me.loc_strSubType = Me.tbGenSubType.Text
  End Sub
  Private Sub tbGenGenre_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbGenGenre.TextChanged
    Me.loc_strGenre = Me.tbGenGenre.Text
  End Sub
  Private Sub tbGenEditor_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbGenEditor.TextChanged
    Me.loc_strEditor = Me.tbGenEditor.Text
  End Sub
  Private Sub tbGenLngId_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbGenLngId.TextChanged
    Me.loc_strLngId = Me.tbGenLngId.Text
  End Sub
  Private Sub tbGenLngName_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbGenLngName.TextChanged
    Me.loc_strLngName = Me.tbGenLngName.Text
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       frmNewText_VisibleChanged
  ' Goal:       Welcome the user when we come into the picture
  ' History:
  ' 22-05-2015  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub frmNewText_VisibleChanged(sender As Object, e As System.EventArgs) Handles Me.VisibleChanged
    If (Me.Visible) Then
      DoInit()
    End If
  End Sub

  Private Sub frmNewText_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
    ' Set timer
    Me.Timer1.Enabled = True
  End Sub
  Private Sub DoInit()
    Try
      ' Make sure I am visible
      While (Not Me.Visible)
        Application.DoEvents()
      End While
      ' Give a welcome message to the user
      Status("Please provide details of the text and copy-paste the text itself")
      ' Perhaps get details myself...
      If (strCurrentFile <> "") Then
        ' Yes, copy the details of the current file
        Me.tbGenTitle.Text = GetFileDesc(pdxCurrentFile, "title")
        Me.tbGenDistributor.Text = GetFileDesc(pdxCurrentFile, "distributor")
        Me.tbGenSource.Text = GetFileDesc(pdxCurrentFile, "bibl")
        Me.tbGenAuthor.Text = GetFileDesc(pdxCurrentFile, "author")
        Me.tbGenCreaManu.Text = GetFileDesc(pdxCurrentFile, "manuscript")
        Me.tbGenCreaOrig.Text = GetFileDesc(pdxCurrentFile, "original")
        Me.tbGenSubType.Text = GetFileDesc(pdxCurrentFile, "subtype")
        Me.tbGenEditor.Text = GetFileDesc(pdxCurrentFile, "editor")
        Me.tbGenGenre.Text = GetFileDesc(pdxCurrentFile, "genre")
        Me.tbGenLngId.Text = GetFileDesc(pdxCurrentFile, "ident", "language")
        Me.tbGenLngName.Text = GetFileDesc(pdxCurrentFile, "name", "language")
        ' Set the current etho language
        strCurrentEthno = GetFileDesc(pdxCurrentFile, "ident", "language")
        If (strCurrentEthno = "") Then
          ' Do we have a period?
          If (strCurrentPeriod <> "") Then
            If (DoLike(strCurrentPeriod, "O[1-4]|O[1-4][1-4]|M[1-4]|M[1-4][1-4]|E[1-3]|B[1-3]")) Then
              strCurrentEthno = "eng_hist_" & strCurrentPeriod
            End If
          End If
        End If

      End If
      arHdType = New HeaderInfo() { _
        New HeaderInfo("title", "", Me.tbGenTitle, Nothing), _
        New HeaderInfo("distributor", "", Me.tbGenDistributor, Nothing), _
        New HeaderInfo("bibl", "", Nothing, Me.tbGenSource), _
        New HeaderInfo("original", "", Me.tbGenCreaOrig, Nothing), _
        New HeaderInfo("manuscript", "", Me.tbGenCreaManu, Nothing), _
        New HeaderInfo("author", "", Me.tbGenAuthor, Nothing), _
        New HeaderInfo("subtype", "", Me.tbGenSubType, Nothing), _
        New HeaderInfo("genre", "", Me.tbGenGenre, Nothing), _
        New HeaderInfo("editor", "", Nothing, Me.tbGenEditor), _
        New HeaderInfo("ident", "language", Me.tbGenLngId, Nothing), _
        New HeaderInfo("name", "language", Me.tbGenLngName, Nothing) _
      }
      ' Set focus to the correct start: the file name
      Me.tbGenFileName.Focus()
    Catch ex As Exception
      ' Show error
      HandleErr("frmNewText/DoInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off timer
      Me.Timer1.Enabled = False
      ' Perform initialisation
      DoInit()
    Catch ex As Exception
      ' Show error
      HandleErr("frmNewText/Timer1 error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class