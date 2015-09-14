Imports System.Xml
Public Class PosTag
  ' ===================================================================================
  ' Name: PosTag
  ' Goal: Class to facilitate agglutinative Part-Of-Speech tagging.
  '       We are making use of the following information
  '       1 - a <liftpos> dictionary   (XML file containing finite forms)
  '       2 - an affix dictionary      (shoebox file)
  '
  '       The relevant information from the LIFT dictionary is transformed into 
  '         a plain ASCII dictionary with each line containing all available variants
  '         of a particular form, together with their grammatical class
  ' History:
  ' 24-03-2011  ERK Derived from AggluSpell for Chechen
  ' 24/jul/2015 ERK This implementation (within Cesax) has restrictions:
  '             the 'cat' will be held equal to 'pos'
  ' ===================================================================================
  ' ====================== PUBLIC TYPES ==========================================
  ' Structure containing all information per suffix 
  Public Structure SfxEntry
    Dim Name As String      ' English name of the affix
    Dim Gloss As String     ' ENglish gloss of the affix
    Dim Src As String       ' Source category
    Dim Dst As String       ' Destination category
    Dim Suffix As String    ' The actual affix
    Dim Rewrite As String   ' How it is rewritten
    Dim EnvPre As String    ' Conditioning environment: preceding
    Dim EnvFol As String    ' Conditioning environment: following
    Dim Type As String      ' What is the type? (finite, suffix)
    Dim Pos As Integer      ' Position within dictionary (if finite)
  End Structure
  Public Structure DictEntry
    Dim Line As String      ' the actual line in the dictionary
    Dim Word As String      ' The word in this entry we were looking for
    Dim Cat As String       ' The category of the word we were looking for
    Dim Pos As String       ' Part-of-speech of this entry
  End Structure
  ' All language specific information: See below in class Language
  ' ====================== PUBLIC VARIABLES =======================================
  Public aggLang As Language = Nothing  ' Language description (set in INIT)
  Public bUseCyrillic As Boolean        ' Are we using cyrillic or not?
  Public colChange As New StringColl    ' Changes that need to be made
  Public tblSkip As DataTable           ' Contains the skipped parses
  ' ====================== GET/SET VARIABLES =======================================
  Private ndxCurrentForest As XmlNode = Nothing ' pointer to current forest
  Private ndxCurrentWord As XmlNode = Nothing   ' pointer to current word
  Private loc_strLanguage As String = ""        ' Needed to browse through Afx dictionary
  Private loc_strLangValue As String = "Lat"    ' Affix dictionary language field value
  Private loc_strDictFile As String = ""        ' Name of the dictionary file in <liftpos> format
  Private loc_strAffxFile As String = ""        ' Name of affix dictionary
  Private loc_strLexField As String = "lex"     ' Which field to use
  Private loc_strAlloField As String = "a"      ' Which allophone field to use in Affix dic
  ' ====================== GET only VARIABLES =======================================
  Private tdlDict As DataSet = Nothing          ' 
  Private lstResult As New List(Of Parse)
  ' ====================== LOCAL VARIABLES =======================================
  Private strError As String = ""   ' If there is an error, this has the text
  Private strResult As String = ""  ' Possible result (only the FIRST one!!)
  Private strPartOfSp As String = ""  ' Part of speech
  Private strDentry As String = ""  ' Dictionary entry found by InDictionary()
  Private strDict As String         ' Contains the whole dictionary
  Private arStack() As SfxEntry     ' Stack for keeping track of suffixes
  Private colStack As New NodeColl  ' Collection for suffixes
  Private intStckNum As Integer     ' Counter for the stack of suffixes
  Private arSfx() As SfxEntry       ' Array of suffixes
  Private intSfxNum As Integer      ' Number of suffixes
  Private intArSfxStart As Integer  ' Starting point for search in suffix array
  Private intDictStart As Integer   ' Starting point for search in dictionary
  Private bInit As Boolean = False  ' Have we initialised spell checking already?
  Private intLevel As Integer       ' Parsing depth level for recursive function
  Private strTrace As String        ' Keep track of spell checking trace
  Private dteSfxDictStamp As Date   ' Date stamp of suffix file
  Private arTrace() As String       ' Array of lines of trace
  Private intTrace As Integer       ' counter for trace
  Private tblParse As DataTable     ' Contains the parses
  Private prsThis As New Parse      ' Room for one parse only!
  Private intParseId As Integer = 1 ' Unique parse id
  Private colTrace As New StringColl ' Here we keep the traces
  Private strPunct As String = ".,<>/?';:[{]}\|-_=+()*&^%$#@!"
  ' ====================== Property get/set ======================================
  Public Property currentForest() As XmlNode
    Get
      Return ndxCurrentForest
    End Get
    Set(ByVal value As XmlNode)
      ndxCurrentForest = value    ' Pointer to currently selected <forest>
    End Set
  End Property
  Public Property currentWord() As XmlNode
    Get
      Return ndxCurrentWord
    End Get
    Set(ByVal value As XmlNode)
      ndxCurrentWord = value      ' Pointer to currently selected word
    End Set
  End Property
  Public Property affixLanguage() As String
    Get
      Return loc_strLanguage
    End Get
    Set(ByVal value As String)
      loc_strLanguage = value     ' Language name to be used in the affix file
      If (loc_strLanguage.ToLower.Contains("cyr")) Then
        loc_strLangValue = "Cyr"
      End If
    End Set
  End Property
  Public Property affixFile() As String
    Get
      Return loc_strAffxFile
    End Get
    Set(ByVal value As String)
      loc_strAffxFile = value
    End Set
  End Property
  Public Property liftposDic() As String
    Get
      Return loc_strDictFile
    End Get
    Set(ByVal value As String)
      loc_strDictFile = value     ' Location of the <liftpos> organized dictionary
    End Set
  End Property
  Public ReadOnly Property dict() As DataSet
    Get
      Return tdlDict
    End Get
  End Property
  Public ReadOnly Property parses() As List(Of Parse)
    Get
      Return Me.lstResult
    End Get
  End Property
  Public WriteOnly Property lexField() As String
    Set(ByVal strName As String)
      loc_strLexField = strName
    End Set
  End Property
  Public WriteOnly Property alloField() As String
    Set(ByVal strName As String)
      loc_strAlloField = strName
    End Set
  End Property


  ' ====================== ACTUAL CODE ===========================================
  ' ----------------------------------------------------------------------------------
  ' Name :      New
  ' Goal :      Initialise using as much default settings as possible
  ' History:
  ' 24-07-2015  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Sub New(ByVal strLangName As String)
    Dim strLang As String
    Dim strDictFile As String
    Dim strAffxFile As String

    Try
      ' Check for the presence of a <listpos> dictionary
      ' -- only use the THREE letter code preceding -
      strLang = Split(strLangName, "-")(0)
      ' Try get default <liftpos> dictionary
      strDictFile = GetDocDir() & "\LiftPos_" & strLang & ".dic"
      If (Not IO.File.Exists(strDictFile)) Then
        Status("PosTag/new: could not find <liftpos> file: " & strDictFile)
      End If
      ' Try get default affix dictionary
      strAffxFile = GetDocDir() & "\Affix_" & strLang & ".dic"
      If (Not IO.File.Exists(strAffxFile)) Then
        Status("PosTag/new: could not find affix file: " & strAffxFile)
      End If
      ' We have everything: go on and initialize further
      InitPosTag(strLangName, strDictFile, strAffxFile)
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/New error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ----------------------------------------------------------------------------------
  ' Name :      InitPosTag
  ' Goal :      Initialise parsing and/or spell checking
  ' History:
  ' 14-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Sub InitPosTag(ByVal strLangName As String, ByVal strDictFile As String, _
                        ByVal strAffxFile As String)
    Try
      ' reset error string
      strError = ""
      ' Reset results string
      strResult = ""
      ' Do we need to Initialise further?
      If (bInit) And (Not NeedUpdate()) Then Exit Sub
      ' set local language, dict and affix dict names
      Me.loc_strAffxFile = strAffxFile
      Me.loc_strDictFile = strDictFile
      Me.loc_strLanguage = strLangName
      ' determine language specific information
      aggLang = New Language(loc_strLanguage, loc_strDictFile, loc_strAffxFile, True)
      ' Initialise what is needed
      ReDim arSfx(0)
      arSfx = Nothing
      ' Check existence of affix dict
      If (IO.File.Exists(aggLang.SfxFile)) Then
        ' Read suffix dictionary
        If (Not ReadSfxDict(aggLang.SfxFile)) Then Status("PosTag/InitPosTag: cannot read affixes") : Exit Sub
        Status("Number of suffixes read: " & UBound(arSfx) + 1)
        Logging("Number of suffixes read: " & UBound(arSfx) + 1)
        ' Fix date/time stamp of the dictionary file of suffixes
        dteSfxDictStamp = IO.File.GetLastWriteTime(aggLang.SfxFile)
      End If
      ' is the <liftpos> dict readable?
      If (IO.File.Exists(aggLang.DictFile)) Then
        ' Read the <liftpos> dictionary
        If (Not ReadDataset("LiftPos.xsd", aggLang.DictFile, tdlDict)) Then Exit Sub
        ' Convert the main dictionary for fast searching
        If (Not InitDict()) Then Exit Sub
      End If
      ' Set init to true
      bInit = (tdlDict IsNot Nothing)
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/InitPosTag error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  IsNumber
  ' Goal :  Check if this is a number, possibly containing -,.
  ' History:
  ' 04-08-2011  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Function IsNumber(ByVal strIn As String) As Boolean
    Dim strTmp As String = ""

    Try
      ' Try simple way
      If (IsNumeric(strIn)) Then Return True
      ' Take out unnecessary elements
      strTmp = strIn.Replace("-", "")
      strTmp = strTmp.Replace(",", "")
      strTmp = strTmp.Replace(".", "")
      ' Return the result
      Return IsNumeric(strTmp)
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/IsNumber error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  IsPunct
  ' Goal :  Check if this is punctuation only
  ' History:
  ' 15-09-2011  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Function IsPunct(ByVal strIn As String) As Boolean
    Dim intPos As Integer ' Position in string

    Try
      ' Check all elements
      For intPos = 1 To strIn.Length
        If (InStr(strPunct, Mid(strIn, intPos, 1)) = 0) Then Return False
      Next intPos
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/IsPunct error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  DoParse
  ' Goal :  Parse a word, returning combinations of Category + Feature
  ' History:
  ' 31-03-2011  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function DoParse(ByVal strWord As String, ByVal intId As Integer, ByRef strAlt As String, _
      ByVal strPrev As String, Optional ByVal strParseTypes As String = "dict|rewrite|carla") As Boolean
    Dim strCat As String          ' Category
    Dim strPos As String          ' POS
    Dim strUword As String        ' Variant of the word with first letter upper case
    Dim strInit As String         ' A possible initial
    Dim ndxWork As XmlNode        ' Working node
    Dim ndxPrev As XmlNode        ' Previous node
    Dim intNum As Integer         ' Number of successful parsings
    Dim bUcase As Boolean         ' Whether the first letter starts with an uppercase
    Dim bFound As Boolean         ' Did we find something?
    Dim strSpeechIntro As String = """" & ":«"

    Try
      ' Validate
      If (strWord = "") Then Return False
      If (Not bInit) Then Return False
      ' Check for interrupt
      If (bInterrupt) Then Return False
      ' Initialisations
      TraceReset()        ' Tracing
      'colResult.Clear()   ' Result stack
      lstResult.Clear()     ' result stack
      strAlt = ""         ' Alternative writing for this word
      bFound = False      ' We have not (yet) found something
      ' Is this word a numeral?
      If (IsNumber(strWord)) Then
        ' Numerals can be parsed straightforwardly
        If (Not AddParse(strWord, strWord, strWord, "QN", "NUM")) Then Return False
      ElseIf (IsPunct(strWord)) Then
        ' This is not a word, but punctuation
        ' The kind of parse we give here depends on the actual punctuation
        Select Case strWord
          Case ".", "!", "?"
            If (Not AddParse(strWord, strWord, strWord, "Punct", ".")) Then Return False
          Case ",", ";", ":"
            If (Not AddParse(strWord, strWord, strWord, "Punct", ",")) Then Return False
          Case """", "«", "»"
            If (Not AddParse(strWord, strWord, strWord, "Punct", "-")) Then Return False
          Case Else
            If (Not AddParse(strWord, strWord, strWord, "Punct", "-")) Then Return False
        End Select
      Else
        ' Perform language-dependant adaptations of the word
        strWord = aggLang.AdaptOrthography(strWord)
        ' Check first letter upper case
        bUcase = (AscW(Left(strWord, 1)) = AscW(UCase(Left(strWord, 1))))
        ' ============== DEBUG
        ' If (strWord Like "хІай*") OrElse (strWord Like "ХІай*") Then Stop
        ' ===========================
        ' Initially we have not found the word
        strWord = aggLang.GetLcase(strWord)
        ' strWord = strWord.ToLower ' GetLcase(strWord)
        ' ============== DEBUG
        ' If (strWord Like "кулпатраву") Then Stop
        ' ===========================
        ' Make alternative variant of the word
        strUword = UCase(Left(strWord, 1)) & Mid(strWord, 2)
        ' The initial category is EMPTY
        strCat = ""
        ' Initialise the parsing depth level
        intLevel = 0
        ' Try to parse all possibilities for this word
        intNum = GetParses(strWord, strParseTypes)
        ' Did we get NO parses at all?
        If (intNum = 0) Then
          ' ========= DEBUG ==========
          ' If (strWord = "10chu") Then Stop
          ' ==========================
          ' Check if this is a derived numeral, e.g: 1-ra, 2-chu etc.
          strCat = DerivedNumber(strWord)
          If (strCat <> "") Then
            ' Determine the POS: it must be equal to the CAT
            ' strPos = CatToPOS(strWord, "(no definition available)", strCat, "")
            strPos = strCat
            ' File the answer away
            If (Not AddParse("", strWord, strWord, strCat, strPos)) Then Return False
            ' Indicate we have it...
            bFound = True
          End If
          ' If this did not work, try Upper case beginning
          If (Not bFound) Then
            intNum = GetParses(strUword, strParseTypes)
            ' ======== DEBUG =========
            If (intNum > 0) Then Stop
            ' ========================
            ' Found something?
            If (intNum = 0) Then
              ' Attempt other strategies...
              If (strCat = "") AndAlso (InStr(strWord, "1") > 0) AndAlso (Not DoLike(strWord, "*sheran|*bettan")) Then
                ' Interpret the number "1" as the letter "w" (Cyrillic influence)
                intNum = GetParses(strWord.Replace("1", "w"), strParseTypes)
              ElseIf (strCat = "") AndAlso (InStr(strWord, "3") > 0) AndAlso (Not DoLike(strWord, "*sheran|*bettan")) Then
                ' Interpret the number "3" as the letter "z" (Cyrillic influence)
                intNum = GetParses(strWord.Replace("3", "z"), strParseTypes)
                ' Thoroughly check if this could be a name (starting with a capital then)
              ElseIf (strCat = "") AndAlso (Left(strWord, 1) <> "w") AndAlso (bUcase) AndAlso _
                (GetForestId(ndxCurrentForest) > 2 OrElse GetForestId(ndxCurrentForest) = 0) _
                AndAlso (strWord.Length > 5) _
                AndAlso (intId > 0) AndAlso (strPrev = "" OrElse InStr(strSpeechIntro, strPrev) = 0) Then
                ' Criteria for common noun:
                ' (1) start with upper case
                ' (2) Not the first word in a text
                ' (3) Not the first word in a sentence
                ' (4) Not after a direct speech introducer
                ' Do the name check
                strCat = GuessCat(strWord, "Capital")
                If (strCat = "") Then
                  ' Regard this as a common name anyway
                  strCat = "NPR" : strPos = "NPR-NOM"
                  ' If we say THIS is a name, then perhaps the preceding word is a name too:
                  ' (1) It must be 1 capital letter
                  ' (2) It must be followed by a period
                  ' (3) The capital letter must be of category "X"
                  ndxPrev = ndxCurrentWord
                  ' See if this is a period
                  If (Not ndxPrev Is Nothing) AndAlso (ndxPrev.FirstChild.Attributes("Text").Value = ".") Then
                    ' The previous sibling should be the capital letter
                    ndxWork = ndxPrev.PreviousSibling
                    If (Not ndxWork Is Nothing) AndAlso ((ndxWork.Attributes("Label").Value = "X") OrElse (ndxWork.FirstChild.Attributes("Text").Value = "A")) Then
                      ' See if this is 1 capital letter
                      strInit = ndxWork.FirstChild.Attributes("Text").Value
                      If (strInit.Length = 1) AndAlso (AscW(Left(strInit, 1)) = AscW(UCase(Left(strInit, 1)))) Then
                        ' This is one unrecognized capital letter --> treat as a name
                        ' (4) add the period to the initial
                        ndxWork.FirstChild.Attributes("Text").Value &= "."
                        ' (5) Set the category correctly
                        ndxWork.Attributes("Label").Value = "NPR-NOM"
                        ' (6) Delete the [prev] node and its child
                        ndxPrev.RemoveAll()
                        ndxCurrentForest.RemoveChild(ndxPrev)
                      End If
                    End If
                  End If
                Else
                  ' Determine the POS
                  ' strPos = CatToPOS(strWord, "(no definition available)", strCat, "")
                  strPos = strCat
                End If
                If (Not AddParse("", strWord, strWord, strCat, strPos)) Then Return False
                ' NOTE: the following ELSE is skipped for this version
                'Else
                '  ' We did not get parses in a "normal" way - try going cyrillic and back to latin again
                '  strAlt = CyrToLat(LatToCyr(strWord))
                '  intNum = GetParses(strAlt, strParseTypes)
                '  If (intNum > 0) Then
                '    ' Add change to collection
                '    colChange.AddUnique(strWord & ";" & strAlt)
                '  Else
                '    ' Reset the alt reading
                '    strAlt = ""
                '    ' There is one last option: try to guess the word category from the form of the word
                '    strCat = GuessCat(strWord, "Any")
                '    ' Got something?
                '    If (strCat <> "") Then
                '      ' Determine the POS
                '      strPos = CatToPOS(strWord, "(no definition available)", strCat, "")
                '      ' File the answer away
                '      If (Not AddParse("", strWord, strWord, strCat, strPos)) Then Return False
                '    End If
                '  End If
              End If
            End If
          End If
        End If
      End If
      ' Check for interrupt
      Return (intNum > 0)
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/DoParse error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      DerivedNumber
  ' Goal :      Check if this is a derived number
  ' History:
  ' 31-08-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Private Function DerivedNumber(ByVal strWord As String) As String
    Dim strVariant As String    ' One variant
    Dim arVariant() As String   ' Array of variants
    Dim dtrGuess() As DataRow   ' Selection from the <Guess> table 
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter

    Try
      ' validate
      If (tdlSettings.Tables("Guess") Is Nothing) Then Return ""
      ' Okay we have a Guess table - use it!
      dtrGuess = tdlSettings.Tables("Guess").Select("Type='NumSfx'")
      For intI = 0 To dtrGuess.Length - 1
        ' Get all variants
        arVariant = Split(dtrGuess(intI).Item("Match").ToString, "|")
        For intJ = 0 To arVariant.Length - 1
          ' Get this variant
          strVariant = arVariant(intJ)
          ' Check this suffix
          If (Right(strWord, strVariant.Length) = strVariant) AndAlso _
             (IsNumber(Left(strWord, strWord.Length - strVariant.Length))) Then
            ' This is an inflected number - return the corresponding CAT
            Return dtrGuess(intI).Item("Cat")
          End If
        Next intJ
      Next intI
      ' Return failure
      Return ""
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/DerivedNumber error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      GuessCat
  ' Goal :      Guess the category of the word by looking at its form
  ' History:
  ' 31-08-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Private Function GuessCat(ByVal strWord As String, ByVal strType As String) As String
    Dim strVariant As String    ' One variant
    Dim arVariant() As String   ' Array of variants
    Dim strExclude As String    ' Words that need to be excluded
    Dim dtrGuess() As DataRow   ' Selection from the <Guess> table 
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter

    Try
      ' validate
      If (tdlSettings.Tables("Guess") Is Nothing) Then Return ""
      ' Okay we have a Guess table - use it!
      dtrGuess = tdlSettings.Tables("Guess").Select("Type='" & strType & "'")
      For intI = 0 To dtrGuess.Length - 1
        ' Get all variants
        arVariant = Split(dtrGuess(intI).Item("Match").ToString, "|")
        For intJ = 0 To arVariant.Length - 1
          ' Get this variant
          strVariant = arVariant(intJ)
          ' Check this suffix
          If (strWord Like strVariant) Then
            ' Check if this word is not in the list of exceptions
            strExclude = dtrGuess(intI).Item("Exclude").ToString
            If (Not (DoLike(strWord, strExclude))) Then
              ' Return the category we found
              Return dtrGuess(intI).Item("Cat")
            End If
          End If
        Next intJ
      Next intI
      ' Return failure
      Return ""
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/GuessCat error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      GetForestId
  ' Goal :      Get the ID for this forest
  ' History:
  ' 29-08-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Private Function GetForestId(ByRef ndxForest As XmlNode) As Integer
    Try
      If (ndxForest Is Nothing) Then
        ' Return failure, indicated by zero
        Return 0
      Else
        ' Return the correct forest id
        Return ndxForest.Attributes("forestId").Value
      End If
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/GetForestId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      AddParse
  ' Goal :      Try add the parse
  ' History:
  ' 22-07-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Private Function AddParse(ByVal strLemma As String, ByVal strWord As String, ByVal strGloss As String, _
                            ByVal strCat As String, ByVal strPos As String) As Boolean
    Try
      ' Make a new parse
      prsThis = New Parse : prsThis.ParseId = intParseId : intParseId += 1
      ' Access the module parse item
      With prsThis
        ' Store the whole word and the lexeme
        .SetWord(strWord, strWord, strCat, strPos)
        .SetLex(strLemma, strWord, strWord, strCat, strPos)
      End With
      ' Add this new parse to the result collection
      lstResult.Add(prsThis)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/AddParse error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      InitSkipStack
  ' Goal :      Initialise the table in which skipped parses will be kept
  ' History:
  ' 06-09-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function InitSkipStack() As Boolean
    Dim intI As Integer   ' Counter

    Try
      ' Check if the table is there already
      If (tblSkip Is Nothing) Then
        tblSkip = New DataTable("Skip")
        With tblSkip
          .Columns.Add("SkipId", System.Type.GetType("System.Int32"))
          .Columns.Add("Lex", System.Type.GetType("System.String"))
          .Columns.Add("Count", System.Type.GetType("System.Int32"))
          .Columns.Add("Loc", System.Type.GetType("System.String"))
        End With
      Else
        ' Clear all rows
        For intI = tblSkip.Rows.Count - 1 To 0 Step -1
          ' Clear this row
          tblSkip.Rows(intI).Delete()
        Next intI
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/InitSkipStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      AddSkipStack
  ' Goal :      Add one row to a Skip stack
  ' History:
  ' 06-09-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function AddSkipStack(ByVal strLex As String, ByVal strLoc As String) As DataRow
    Dim dtrThis As DataRow = Nothing  ' New datarow
    Dim dtrFound() As DataRow         ' Result of SELECT

    Try
      ' Validate
      If (tblSkip Is Nothing) Then Return Nothing
      ' Check if we already have a row with this loc
      dtrFound = tblSkip.Select("Lex='" & strLex.Replace("'", "''") & "'")
      If (dtrFound.Length = 0) Then
        ' Make a new row
        dtrThis = tblSkip.NewRow
        ' Fill it
        With dtrThis
          .Item("SkipId") = GetNewId(tblSkip, "SkipId")
          .Item("Lex") = strLex
          .Item("Count") = 1
          .Item("Loc") = strLoc
        End With
        ' Add this row
        tblSkip.Rows.Add(dtrThis)
      Else
        ' Add to this row
        dtrFound(0).Item("Count") += 1
        dtrThis = dtrFound(0)
      End If
      ' Return this row
      Return dtrThis
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/AddSkipStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      SkipStackSize
  ' Goal :      Give the number of Skips in the stack
  ' History:
  ' 06-09-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function SkipStackSize() As Integer
    Try
      ' Validate
      If (tblSkip Is Nothing) Then Return 0
      ' Return the size
      Return tblSkip.Rows.Count
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/SkipStackSize error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function

  ' ----------------------------------------------------------------------------------
  ' Name :      InitDict
  ' Goal :      Initialise fast dictionary searching
  ' History:
  ' 29-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function InitDict() As Boolean
    Try
      ' Validate
      If (aggLang.Name = "") Then Return False
      ' Sort all main dictionary entries
      aggLang.Main = tdlDict.Tables("entry")
      aggLang.Pdg = tdlDict.Tables("pdg")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/InitDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      TryDict
  ' Goal :      Find all entries of strWord in dictionary and store as results
  ' History:
  ' 02-01-2007  ERK Created
  ' ----------------------------------------------------------------------------------
  Private Function TryDict(ByVal strWord As String) As Boolean
    Dim strCat As String          ' Category of the word
    Dim strMorph As String        ' Morphological break-down
    Dim strPos As String          ' POS of this particular word
    Dim strLemma As String        ' Current lemma
    Dim strDef As String          ' Definition
    Dim intI As Integer           ' Counter
    Dim dtrFound() As DataRow     ' Result of selection
    'Dim objRes As NodeColl        ' Collection of resulting nodes

    Try
      ' Initialisations
      strCat = "" : strLemma = strWord
      ' ============= DEBUG ==============
      ' If (strWord = "lar") Then Stop
      ' ==================================
      dtrFound = aggLang.Main.Select(loc_strLexField & "='" & strWord.Replace("'", "''") & "'")
      ' Add the results
      For intI = 0 To dtrFound.Length - 1
        ' Get the values for this result
        With dtrFound(intI)
          strCat = .Item("rpos") : strMorph = strCat : strPos = .Item("pos")
          strDef = .Item("def")
          ' No spaces
          strDef = strDef.Replace(",", "|")
          strDef = strDef.Replace(" ", "")
        End With
        ' Make a new element
        prsThis = New Parse : prsThis.ParseId = intParseId : intParseId += 1
        ' Add this result
        With prsThis
          .Clear()
          .SetWord(strWord, strMorph, strCat, strPos)
          ' .SetLex(strLemma, strWord, strWord & "/" & strMorph, strCat, strPos)
          .SetLex(strLemma, strWord, strDef, strCat, strPos)
        End With
        ' colResult.Add(prsThis)
        lstResult.Add(prsThis)
        ' Add tracing
        TraceAdd(Space(intLevel * 2) & "Dictionary.Main(lev=" & intLevel & "): " & _
                 strWord & "/" & dtrFound(intI).Item("rpos"))
      Next intI
      ' Check for subentries
      dtrFound = aggLang.Pdg.Select(loc_strLexField & "='" & strWord.Replace("'", "''") & "'")
      strDef = "-"
      ' Add the results
      For intI = 0 To dtrFound.Length - 1
        ' Get the values for this result
        With dtrFound(intI)
          strCat = .Item("rpos") : strMorph = strCat : strPos = .Item("pos")
          ' Get the paradigm definition,not the lemma...
          strDef = .GetParentRow("entry_pdg").Item("def").ToString
          ' No spaces
          strDef = strDef.Replace(",", "|")
          strDef = strDef.Replace(" ", "")
          ' Get the lemma from the paradigm
          strLemma = .GetParentRow("entry_pdg").Item(loc_strLexField).ToString
        End With
        ' Make a new element
        prsThis = New Parse : prsThis.ParseId = intParseId : intParseId += 1
        ' Add this result
        With prsThis
          .Clear()
          .SetWord(strWord, strMorph, strCat, strPos)
          .SetLex(strLemma, strWord, strDef & "/" & strMorph, strCat, strPos)
        End With
        ' colResult.Add(prsThis)
        lstResult.Add(prsThis)
        ' Add tracing
        TraceAdd(Space(intLevel * 2) & "Dictionary(lev=" & intLevel & "): " & strWord & "/" & dtrFound(intI).Item("rpos"))
      Next intI

      ' Received success
      Return True
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/TryDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ----------------------------------------------------------------------------------
  ' Name :      TryWordCat
  ' Goal :      Test the combination of word/category
  ' Note :      
  ' History:
  ' 14-07-2006  ERK Created
  ' 20-07-2006  ERK Adapted for new scheme
  ' 31-12-2007  ERK Not allow finite form on same line with matching category
  ' 08-04-2011  ERK Adapted for CheCorpus, for POS tagging
  ' 16-04-2011  ERK Probleem: hoe voorkom ik te VEEL parses (die niet goed zijn)?
  ' ----------------------------------------------------------------------------------
  Private Function TryWordCat(ByVal strWord As String, ByVal strCat As String) As Boolean
    Dim intPos As Integer         ' Position within string strWord
    Dim intResult As Integer      ' Number of results found
    Dim intI As Integer           ' Counter for looping through suffixes found
    Dim intJ As Integer           ' counter
    Dim intNum As Integer         ' Number of entries in dictionary
    Dim intEntry As Integer       ' Number of the current entry
    Dim intCurrentPos As Integer  ' Position in string
    Dim intResultPos As Integer   ' Current position in the colResult stack
    Dim strEntry As String        ' Text of the current dictionary entry
    Dim strFinite As String       ' Finite form we are looking at
    Dim strSuffix As String       ' Suffix we are considering
    Dim strWordNew As String      ' New word to be tested
    Dim strCatNew As String       ' Category of new word to be tested
    Dim strCatOrg As String       ' Original category to work with
    Dim strSrcCat As String       ' Source category
    Dim strDstCat As String       ' Destination category
    Dim strEnvPre As String       ' Preceding environment
    Dim strEnvFol As String       ' Following environment
    Dim strPos As String          ' The part of speech
    Dim strCurrentPos As String   ' POS we are currently working on
    Dim strGloss As String        ' My gloss
    Dim arDictLine() As DictEntry ' Array of entry lines from dictionary
    Dim arResult() As SfxEntry    ' Array of resulting suffix objects
    ' Dim objEntry As SfxEntry      ' One entry found in result array
    Dim dtrSuffix() As DataRow    ' Result of looking for the correct suffixes
    Dim dtrEntry As DataRow       ' One affix match we found
    Dim dtrParent() As DataRow    ' My parent
    Dim prsLocal As New Parse     ' Local parse 
    Dim bFound As Boolean         ' Flag denoting success or failure

    Try
      ' Register the level we are in
      intLevel += 1
      ' Initialise the position within the word -string
      intPos = 1
      ' Initialise original category
      strCatOrg = strCat
      ' Loop through the whole word, dividing it in two parts
      Do
        ' Re-initialise category
        strCat = strCatOrg
        ' Go to the next position in the word
        intPos += 1
        ' Divide the word in a finite and a suffix part
        strFinite = Left(strWord, intPos - 1)
        strSuffix = Mid(strWord, intPos)
        ' ========= DEBUGGING ===============
        ' If (strFinite = "шин") AndAlso (strSuffix = "ал") AndAlso (intLevel = 1) Then Stop
        ' ===================================
        ' Initialise this possible parse
        prsLocal.Clear()
        ' Re initialise results array 
        ReDim arResult(0)
        arResult = Nothing
        ' Get all suffixes from SfxDict matching strSuffix
        dtrSuffix = tdlSettings.Tables("AfxEl").Select("Lang='" & loc_strLangValue & _
              "' AND Rewrite='" & strSuffix.Replace("'", "''") & "'")

        intResult = dtrSuffix.Length
        ' Are there any results?
        If (intResult > 0) Then
          ' Add this finite form
          ' Show that we are pursuing this finite form
          TraceAdd(Space(intLevel * 2) & "[1] finite(lev=" & intLevel & "): " & strFinite & " + " & Sfx(strSuffix))
        End If
        ' Go one level up for the suffixes
        intLevel += 1
        ' Loop through all results...
        For intI = 1 To intResult
          ' Re-initialise category
          strCat = strCatOrg
          ' ============= DEBUG ===============
          'If (intI = 11) Then Stop
          ' ===================================
          ' Retrieve the next result from the array
          ' objEntry = arResult(intI - 1)
          dtrEntry = dtrSuffix(intI - 1)
          ' Get the source and destination categories
          strSrcCat = dtrEntry.Item("CatSrc").ToString
          strDstCat = dtrEntry.Item("CatDst").ToString
          strEnvPre = dtrEntry.Item("EnvPre").ToString
          strEnvFol = dtrEntry.Item("EnvFol").ToString
          ' Get my parent
          dtrParent = tdlSettings.Tables("Affix").Select("AffixId=" & dtrEntry.Item("AffixId"))
          If (dtrParent Is Nothing) Then
            ' This is not possible
            MsgBox("modPosTag/TryWordCat: cannot find a parent row")
            Return False
          Else
            ' Get my "gloss" from the parent
            strGloss = dtrParent(0).Item("Gloss").ToString
            ' Do we have a gloss?
            If (strGloss = "") Then
              ' There's no gloss, so take the name instead
              strGloss = dtrParent(0).Item("Name").ToString
            End If
          End If
          ' Adapt tracing
          TraceAdd(Space(intLevel * 2) & "sfx: " & Sfx(strSuffix) & " " & strGloss & " " & _
                   Sfx(strDstCat) & " < " & dtrEntry.Item("Suffix") & " " & _
                   Sfx(strSrcCat) & " / " & strEnvPre & "_" & strEnvFol)
          ' Check the category: the category of the TOTAL word should match
          '                     the destination category of the suffix
          ' ============ DEBUG
          'If (objEntry.Src = "QAdj") AndAlso (strCat = "") Then Stop
          ' ====================
          If (strCat = "") Or (strCat = strDstCat) Then
            ' Check the environment
            If (CheckEnv(strFinite, strSuffix, strEnvPre, strEnvFol)) Then
              ' Make up a new word and category
              strWordNew = strFinite & dtrEntry.Item("Suffix")
              strCatNew = strSrcCat
              'If (strCat = "") Or (strCat = objEntry.Dst) Then
              '  ' Check the environment
              '  If (CheckEnv(strFinite, strSuffix, objEntry.EnvPre, objEntry.EnvFol)) Then
              '    ' Make up a new word and category
              '    strWordNew = strFinite & objEntry.Suffix
              '    strCatNew = objEntry.Src
              ' Determine the resulting category
              ' strPos = CatToPOS(strWord, "(no definition)", objEntry.Dst.Replace("-", "/"), "")
              strPos = ""
              ' Reset the DictLine stack for this combination
              ReDim arDictLine(0)
              ' See if this NEW combination is in dictionary
              intNum = GetDictLines(strWordNew, strCatNew, arDictLine)
              ' If there is, then store results
              If (intNum > 0) Then
                ' Store all matched dictionary entries
                For intEntry = 1 To intNum
                  Dim strDef As String = ""
                  Dim prsSave As New Parse

                  bFound = True
                  strCat = arDictLine(intEntry - 1).Cat
                  strEntry = arDictLine(intEntry - 1).Line
                  strCurrentPos = arDictLine(intEntry - 1).Pos
                  ' Make a new parse
                  prsThis = New Parse : prsThis.ParseId = intParseId : intParseId += 1
                  ' Add this new parse to the result collection
                  lstResult.Add(prsThis)
                  ' Find the correct definition by applying TryDict
                  prsSave = prsThis
                  If (TryDict(strWordNew)) Then
                    ' The answer must be in the last 'result'
                    Dim prsLex As Parse = lstResult.Last
                    strDef = prsLex.Def
                    ' Delete this last one
                    lstResult.RemoveAt(lstResult.Count - 1)
                    prsThis = prsSave
                  Else
                    strDef = prsThis.Def
                  End If

                  ' Access the module parse item
                  With prsThis
                    ' Store the whole word and the lexeme
                    .SetWord(strWord, "", strCat, "")
                    ' Set the [objLex] in the class [Parse]
                    '.SetLex("", strWordNew, strEntry, strCatNew, strCurrentPos)
                    .SetLex(strEntry, strWordNew, strDef, strCatNew, strCurrentPos)
                    .PushSfx(dtrEntry.Item("Suffix") & ">" & dtrEntry.Item("Rewrite"), _
                             strGloss, strSrcCat, strDstCat, strPos)
                  End With
                  ' Adapt tracing
                  TraceAdd(Space(intLevel * 2) & "[3] FOUND(lev=" & intLevel & "): " & _
                           strWordNew & "/" & strCat & _
                           " " & arDictLine(intEntry - 1).Pos & " (" & strEntry & "+" & _
                           strGloss & "/" & strDstCat & ")")
                  'End If
                Next intEntry
              Else  ' There are no dictionary entries
                ' Note where we are
                ' intResultPos = colResult.Count
                intResultPos = lstResult.Count
                ' Try recursively with this combination
                If (TryWordCat(strWordNew, strCatNew)) Then
                  ' Adapt all the entries in [colResult] starting from [intResultPos]
                  ' For intJ = intResultPos To colResult.Count - 1
                  For intJ = intResultPos To lstResult.Count - 1
                    ' Access the module parse item
                    With lstResult.Item(intJ)
                      .PushSfx(dtrEntry.Item("Suffix") & ">" & dtrEntry.Item("Rewrite"), _
                               strGloss, strSrcCat, strDstCat, strPos)
                    End With
                  Next intJ
                  ' Adapt tracing
                  TraceAdd(Space(intLevel * 2) & "[4] Trying(lev=" & intLevel & "): " & strWordNew & "/" & strCatNew)
                  ' Make sure result gets evaluated positively
                  bFound = True
                Else  ' TryWordCat
                  ' Add trace comment
                  TraceAdd(Space(intLevel * 2) & vbTab & _
                           "[5] combination failed(lev=" & intLevel & "): " & strWordNew & "/" & strCatNew)
                  ' If this is level 1, then pop the whole combination
                  ' If (intLevel = 1) Then colResult.Pop()
                  If (intLevel = 1) Then lstResult.RemoveAt(lstResult.Count - 1)
                End If ' TryWordCat
              End If   ' intNum > 0
            Else       ' CheckEnv
              ' Add trace comment
              TraceAdd(Space(intLevel * 2) & vbTab & "[6] environment failed(lev=" & intLevel & "): / " & _
                strEnvPre & "_" & strEnvFol)
            End If
          Else         ' Category matching
            ' Add trace comment
            TraceAdd(Space(intLevel * 2) & vbTab & "[7] category failed(lev=" & intLevel & "): " & _
                     strCat & "(" & strDstCat & ")")
          End If       ' Category matching
        Next intI
        ' Adapt level
        intLevel -= 1
        ' Continue until the suffix is empty (we had all possibilities then!)
      Loop Until (strSuffix = "")
      'End If
      ' Adapt level
      intLevel -= 1
      ' Pass back ok
      Return bFound
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/TryWordCat error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function


  ' ----------------------------------------------------------------------------------
  ' Name :      GetParses
  ' Goal :      Try to parse the [strWord], and return the result in [lstResult]
  ' History:
  ' 08-04-2011  ERK Derived from GetParses
  ' ----------------------------------------------------------------------------------
  Private Function GetParses(ByVal strWord As String, Optional ByVal strType As String = "dict|rewrite|carla") As Integer
    Dim strPOS As String = ""   ' The part of speech
    Dim strCat As String = ""   ' The category
    Dim arType() As String = Split(strType, "|")
    Dim intNum As Integer = 0   ' Result
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intK As Integer         ' Counter

    Try
      ' Initialise the parse stack and initialise the parses
      lstResult.Clear()
      ' =============== DEBUG ==============
      ' If (strWord = "baax") Then Stop
      ' ====================================
      ' Start search other forms with empty category
      strCat = ""
      For intK = 0 To arType.Length - 1
        ' Watch out for interrupt
        If (bInterrupt) Then Return 0
        Select Case arType(intK)
          Case "dict"
            ' Store results that can be found straight in dictionary
            If (Not TryDict(strWord)) Then Return 0
            intNum += lstResult.Count
          Case "rewrite"
            If (TryWordCat(strWord, strCat)) Then
              ' No need to keep track of the number of parses, but this does show we have at least some result
              ' intNum = colResult.Count
              intNum += lstResult.Count
            End If
          Case "carla"
            ' DO NOT USE IN THIS IMPLEMENTATION!!
            'If (TryWordCarla(strWord)) Then
            '  ' No need to keep track of the number of parses, but this does show we have at least some result
            '  intNum = colResult.Count
            'End If
        End Select
      Next intK
      ' Tidy up
      intI = 0
      While (intI < lstResult.Count - 1)
        ' Get the POS of this outcome
        strPOS = lstResult.Item(intI).POS
        ' See if it occurs elsewhere
        For intJ = intI + 1 To lstResult.Count - 1
          ' Check co-occurrance
          If (lstResult.Item(intJ).POS = strPOS) Then
            ' Throw out item [intI]
            lstResult.RemoveAt(intJ)
            ' Leave this for loop
            Exit For
          End If
        Next intJ
        intI += 1
      End While
      'For intI = 0 To lstResult.Count - 1
      '  ' Get the POS of this outcome
      '  strPOS = lstResult.Item(intI).POS
      '  ' See if it occurs elsewhere
      '  For intJ = intI + 1 To lstResult.Count - 1
      '    ' Check co-occurrance
      '    If (lstResult.Item(intJ).POS = strPOS) Then
      '      ' Throw out item [intI]
      '      lstResult.RemoveAt(intJ)
      '      ' Leave this for loop
      '      Exit For
      '    End If
      '  Next intJ
      'Next intI
      ' Return the result
      Return intNum
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/GetParses error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      NumLines
  ' Goal :      Determine the amount of lines in string strIn
  ' History:
  ' 21-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Private Function NumLines(ByVal strIn As String) As Integer
    Dim arThis() As String

    Try
      If (strIn = "") Then
        NumLines = 0
      Else
        arThis = Split(strIn, vbCrLf)
        NumLines = UBound(arThis) + 1
      End If
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/NumLines error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function
 
  ' ----------------------------------------------------------------------------------
  ' Name :      CheckEnv
  ' Goal :      See if strSuffix when attached to strfinite fulfills the 
  '               environment conditions set out by strPre and strFol
  ' History:
  ' 20-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Private Function CheckEnv(ByVal strFinite As String, ByVal strSuffix As String, _
    ByVal strPre As String, ByVal strFol As String) As Boolean
    Dim strLet As String = ""
    Dim strFiniteEnding As String
    Dim intPos As Integer
    Dim bFound As Boolean

    Try
      ' TODO: Implement checking...
      strPre = Trim(strPre)
      bFound = True
      ' Determine where to put the pointer
      If (Right(strPre, 1) = "]") Then
        intPos = InStrRev(strPre, "[")
      Else
        intPos = Len(strPre) - 1
      End If
      While (intPos > 0) And (strFinite <> "")
        ' Extract letter or environment indicator we are looking at
        strFiniteEnding = Mid(strPre, intPos)
        'If (Right(strFiniteEnding, 1) <> "]") Then
        '  MsgBox("Problem checking environment!")
        '  Stop
        'End If
        ' Adjust the remainder of the environment
        strPre = Trim(Left(strPre, intPos - 1))
        Select Case strFiniteEnding
          Case "[C]"
            ' Finite form should end on consonant
            bFound = aggLang.EndsOnCons(strFinite, strLet)
            ' Adapt Finite for further environment checking
            strFinite = Left(strFinite, Len(strFinite) - Len(strLet))
          Case "[V]"
            ' Finite form should end on consonant
            bFound = aggLang.EndsOnVowel(strFinite, strLet)
            ' Adapt Finite for further environment checking
            strFinite = Left(strFinite, Len(strFinite) - Len(strLet))
          Case "[D]"
            ' Finite should end on class marker
            bFound = aggLang.EndsOnClass(strFinite, strLet)
            ' Adapt Finite for further environment checking
            strFinite = Left(strFinite, Len(strFinite) - Len(strLet))
        End Select
        ' Determine the new position where the next environment indicator starts
        If (Right(strPre, 1) = "]") Then
          intPos = InStrRev(strPre, "[")
        Else
          intPos = Len(strPre) - 1
        End If
      End While
      CheckEnv = bFound
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/CheckEnv error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      SpellError
  ' Goal :      Give possible spelling error back
  ' History:
  ' 14-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function SpellError() As String
    SpellError = strError
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      SpellResult
  ' Goal :      Give the results of the spelling of the words
  ' History:
  ' 14-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function SpellResult() As String
    SpellResult = IIf(strResult = "", strDentry & vbCrLf, strResult)
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      NeedUpdate
  ' Goal :      Find out whether an update in initialisation is needed
  ' History:
  ' 31-12-2007  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function NeedUpdate() As Boolean
    If (aggLang Is Nothing) Then Return True
    If (aggLang.SfxFile Is Nothing) Then
      NeedUpdate = True
    Else
      NeedUpdate = (IO.File.GetLastWriteTime(aggLang.SfxFile) <> dteSfxDictStamp)
    End If
  End Function

  ' ----------------------------------------------------------------------------------
  ' Name :      ReadSfxDict
  ' Goal :      Read suffix dictionary into structure array from indicated file
  ' History:
  ' 14-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Private Function ReadSfxDict(ByVal strFile As String) As Boolean
    Dim strText As String       ' Text of the suffix file
    Dim strEntry As String      ' One entry
    Dim strLine As String       ' One line in an entry
    Dim strName As String       ' Name of the suffix
    Dim strGloss As String      ' English gloss of the suffix
    Dim strType As String       ' Type of affix (suffix, prefix etc)
    Dim strCat As String        ' Category
    Dim strEnv As String        ' Conditioning environment
    Dim strEnvPre As String     ' Preceding environment (before _)
    Dim strEnvFol As String     ' Following environment (after _)
    Dim strSrc As String        ' Source category
    Dim strDst As String        ' Destination category
    Dim strSuffix As String     ' Source suffix
    Dim strRewrite As String    ' Destination suffix
    Dim strSfm As String        ' Standard format marker (command)
    Dim strRest As String       ' Remainder of the line
    Dim strTable As String      ' Name of one table
    Dim strLang As String       ' Language abbreviation
    Dim strComm As String       ' Comment
    Dim arEntry() As String     ' Array of entries
    Dim arLine() As String      ' Lines within each entry
    Dim arCat() As String       ' array of categories
    Dim arSrc() As String
    Dim arDst() As String
    Dim arTable() As String = {"AfxEl", "Affix"}
    Dim intAffixId As Integer   ' ID within <Affix>
    Dim intAfxElId As Integer   ' ID within <AfxEl>
    Dim intPos As Integer       ' Position within a string
    Dim intI As Integer         ' stepping through array of entries
    Dim intJ As Integer         ' Stepping through lines in an entry
    Dim intK As Integer         ' Stepping through all categories
    Dim dtrAffix As DataRow     ' One \q entry stowed away in a datarow
    Dim dtrAfxEl As DataRow     ' One \c, \a, \ac Affix element stowed away in a datarow
    Dim dtrParent As DataRow    ' The parent row in [AffixList]

    Try
      If (Not IO.File.Exists(strFile)) Then
        ' Give error and return
        strError &= "Cannot open suffix dictionary " & strFile
        ReadSfxDict = False
        Exit Function
      End If
      ' Initialise needed strings
      strName = "" : strGloss = "" : strSrc = "" : strDst = ""
      strSuffix = "" : strRewrite = "" : strEnvPre = "" : strEnvFol = ""
      ReDim arSfx(0) : intSfxNum = 0
      ' Initialise string arrays
      arSrc = Nothing : arDst = Nothing : arCat = Nothing
      dtrAffix = Nothing : dtrAfxEl = Nothing
      ' Make sure there is an [AffixList] element
      If (tdlSettings.Tables("AffixList") Is Nothing) OrElse (tdlSettings.Tables("AffixList").Rows.Count = 0) Then
        ' Make one row
        dtrParent = Nothing
        If (Not CreateNewRow(tdlSettings, "AffixList", "", -1, dtrParent)) Then Return False
      Else
        ' Get the parent row
        dtrParent = tdlSettings.Tables("AffixList").Rows(0)
        ' Go through all necessary tables
        For intK = 0 To UBound(arTable)
          ' Get name of this table
          strTable = arTable(intK)
          ' Delete any previous rows
          If (Not tdlSettings.Tables(strTable) Is Nothing) AndAlso (tdlSettings.Tables(strTable).Rows.Count > 0) Then
            ' Delete existing rows
            With tdlSettings.Tables(strTable)
              For intI = .Rows.Count - 1 To 0 Step -1
                .Rows(intI).Delete()
              Next intI
            End With
          End If
        Next intK
        ' Accept changes
        tdlSettings.AcceptChanges()
      End If
      ' Read file into text
      strText = IO.File.ReadAllText(strFile)
      ' Convert into array of suffix entries
      arEntry = Split(strText, "\g ")
      ' Go through array of entries
      For intI = 0 To UBound(arEntry)
        ' Get the text of this entry
        strEntry = arEntry(intI)
        ' Convert into lines
        arLine = Split(strEntry, vbCrLf)
        ' The first line should contain the name of this suffix
        strName = Trim(arLine(0))
        ' Initialise other matters
        strGloss = ""
        ' Make a new entry in <Affix>
        If (Not CreateNewRow(tdlSettings, "Affix", "AffixId", intAffixId, dtrAffix)) Then Return False
        dtrAffix.Item("Name") = strName
        dtrAffix.SetParentRow(dtrParent)
        ' Go through the remaining lines
        For intJ = 1 To UBound(arLine)
          strLine = Trim(arLine(intJ))
          If (strLine <> "") Then
            intPos = InStr(strLine, " ")
            If (intPos > 0) Then
              ' Split into command and remainder
              strSfm = Left(strLine, intPos - 1)
              strRest = Trim(Mid(strLine, intPos + 1))
              Select Case strSfm
                Case "\t"     ' Type of affix
                  strType = strRest
                  dtrAffix.Item("Type") = strType
                Case "\ge"   ' English gloss of suffix
                  strGloss = strRest
                  dtrAffix.Item("Gloss") = strGloss
                Case "\de"    ' Definition of affix
                  dtrAffix.Item("Def") = strRest
                Case "\c"    ' Category
                  ' Strip off everything after the character "//"
                  intPos = InStr(strRest, "//")
                  If (intPos > 0) Then
                    strRest = Trim(Left(strRest, intPos - 1))
                  End If
                  ' Divide the category line up into array of categories
                  arCat = Split(strRest, ",")
                  ' Make space for source and destination categories
                  ReDim arSrc(UBound(arCat))
                  ReDim arDst(UBound(arCat))
                  ' Step through array of categories
                  For intK = 0 To UBound(arCat)
                    strCat = arCat(intK)
                    ' There should be a line between the Source and Destination category
                    intPos = InStr(strCat, "/")
                    If (intPos > 0) Then
                      ' Divide into source and destination
                      strSrc = Trim(Left(strCat, intPos - 1))
                      strDst = Trim(Mid(strCat, intPos + 1))
                      ' Process into array
                      arSrc(intK) = strSrc
                      arDst(intK) = strDst
                    End If
                  Next
                Case "\a", "\ac"    ' Allophone -- these are the real entries!!!
                  ' NOTE: take \a for latin and \ac for cyrillic
                  strLang = IIf(strSfm = "\a", "Lat", "Cyr")
                  ' Strip off everything after the character "//"
                  intPos = InStr(strRest, "//")
                  If (intPos > 0) Then
                    strComm = Trim(Mid(strRest, intPos + 2))
                    strRest = Trim(Left(strRest, intPos - 1))
                  Else
                    strComm = ""
                  End If
                  ' Initialise environment
                  strEnvPre = "" : strEnvFol = ""
                  ' See if there is a conditioning environment following
                  intPos = InStr(strRest, "/")
                  If (intPos > 0) Then
                    ' Also get the conditioning environment
                    strEnv = Trim(Mid(strRest, intPos + 1))
                    ' Get remainder of strRest
                    strRest = Trim(Left(strRest, intPos - 1))
                    ' Divide environment in PRECEDING and FOLLOWING part by looking _
                    intPos = InStr(strEnv, "_")
                    If (intPos > 0) Then
                      strEnvPre = Trim(Left(strEnv, intPos - 1))
                      strEnvFol = Trim(Mid(strEnv, intPos + 1))
                    End If
                  End If
                  ' Get the source and destination string
                  intPos = InStr(strRest, ">")
                  If (intPos > 0) Then
                    strSuffix = Trim(Left(strRest, intPos - 1))
                    ' Adapt for zero suffix
                    If (strSuffix = "0") Then strSuffix = ""
                    strRewrite = Trim(Mid(strRest, intPos + 1))
                    ' Adapt for zero rewrite (don't know if this should be possible though)
                    If (strRewrite = "0") Then strRewrite = ""
                    ' If there are ANY categories...
                    If (Not arCat Is Nothing) Then
                      ' Step through all categories
                      For intK = 0 To UBound(arSrc)
                        ' Is this a latin one?
                        If (strLang = "Lat") Then
                          ' Process the suffix into the array
                          ReDim Preserve arSfx(intSfxNum)
                          With arSfx(intSfxNum)
                            .Name = strName
                            .Gloss = strGloss
                            .Src = arSrc(intK)
                            .Dst = arDst(intK)
                            .Suffix = strSuffix
                            .Rewrite = strRewrite
                            .EnvPre = strEnvPre
                            .EnvFol = strEnvFol
                            .Type = "suffix"
                          End With
                          intSfxNum += 1
                        End If
                        ' Process the affix into one <AfxEl> entry
                        If (Not CreateNewRow(tdlSettings, "AfxEl", "AfxElId", intAfxElId, dtrAfxEl, 0, intAffixId, _
                            arSrc(intK), arDst(intK), strLang, strSuffix, strRewrite, strEnvPre, strEnvFol, strComm)) Then Return False
                        ' Set the correct parent
                        dtrAfxEl.SetParentRow(dtrAffix)
                      Next intK
                    End If
                  End If
                Case "\dte"  ' ENtry date -- skip
              End Select
            End If
          End If
        Next intJ
      Next intI
      ' Save the updated settings file
      XmlSaveSettings(strSetFile)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/ReadSfxDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
 
  ' ----------------------------------------------------------------------------------
  ' Name :      GetDictLines
  ' Goal :      Get an array of strings containing dictionary lines with the correct
  '               combination of strWord/strCat.
  '             If strCat is empty, then entries with any category are given
  ' History:
  ' 01-01-2008  ERK Created
  ' ----------------------------------------------------------------------------------
  Private Function GetDictLines(ByVal strWord As String, ByRef strCat As String, _
    ByRef arDictLine() As DictEntry) As Integer
    Dim intSize As Integer        ' The current size of the array
    Dim intI As Integer           ' Counter
    Dim dtrFound() As DataRow     ' Result of SELECT function

    Try
      ' Initialise array
      intSize = 0 : ReDim arDictLine(0)
      ' First check if it is a main entry - depending on specified [strCat] or not
      If (strCat = "") Then
        ' Any category will do
        dtrFound = aggLang.Main.Select(loc_strLexField & "='" & strWord.Replace("'", "''") & "'")
      Else
        ' Look for combination with specific category
        dtrFound = aggLang.Main.Select(loc_strLexField & "='" & strWord.Replace("'", "''") & "' AND " & _
                                       "rpos='" & strCat & "'")
      End If

      ' Add the results
      For intI = 0 To UBound(dtrFound)
        ' Add this to the stack
        AddDictLine(arDictLine, strWord, dtrFound(intI).Item("rpos"), dtrFound(intI).Item("pos"), strWord)
        intSize += 1
        ' Add tracing
        TraceAdd(Space(intLevel * 2) & "Dictionary.Main(lev=" & intLevel & "): " & strWord & "/" & dtrFound(intI).Item("rpos"))
      Next intI
      ' Check for subentries - again depending on specific [strCat] or not
      If (strCat = "") Then
        ' Any category will do
        dtrFound = aggLang.Pdg.Select(loc_strLexField & "='" & strWord.Replace("'", "''") & "'")
      Else
        ' Look for combination with specific category
        dtrFound = aggLang.Pdg.Select(loc_strLexField & "='" & strWord.Replace("'", "''") & "' AND " & _
                                       "rpos='" & strCat & "'")
      End If
      ' Add the results
      For intI = 0 To UBound(dtrFound)
        ' Add this to the stack
        AddDictLine(arDictLine, dtrFound(intI).GetParentRow("entry_pdg").Item(loc_strLexField), _
                    dtrFound(intI).Item("rpos"), dtrFound(intI).Item("pos"), strWord)
        intSize += 1
        ' Add tracing
        TraceAdd(Space(intLevel * 2) & "Dictionary.Sub(lev=" & intLevel & "): " & strWord & "/" & dtrFound(intI).Item("rpos"))
      Next intI
      ' Return the size
      Return intSize
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/GetDictLines error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      AddDictLine
  ' Goal :      Add one line into the [arDictLine] array
  ' History:
  ' 05-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Private Function AddDictLine(ByRef arDictLine() As DictEntry, ByVal strMain As String, ByVal strCat As String, _
                               ByVal strPos As String, ByVal strWord As String) As Boolean
    Dim intSize As Integer        ' The current size of the array
    Dim objLine As DictEntry      ' One new line into the dictionary

    Try
      ' Store line information into structure
      objLine = New DictEntry
      With objLine
        .Line = strMain   ' Main entry
        .Cat = strCat     ' Output category
        .Word = strWord   ' This entry
        .Pos = strPos     ' POS of this line
      End With
      ' Push line into stack of resulting lines
      If (arDictLine.Length = 1) Then
        ' Check if there is any content
        If (arDictLine(0).Cat Is Nothing) Then
          ' There is no content yet
          intSize = 0
        Else
          intSize = 1
        End If
      Else
        intSize = arDictLine.Length
      End If
      ReDim Preserve arDictLine(0 To intSize)
      arDictLine(intSize) = objLine
      intSize += 1
      Return True
    Catch ex As Exception
      ' Show error
      MsgBox("modPosTag/AddDictLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
 
  Private Function Sfx(ByVal strIn As String) As String
    Sfx = IIf(strIn = "", ".", strIn)
  End Function


  ' ----------------------------------------------------------------------------------
  ' Name :      TraceReset
  ' Goal :      Clear the string that keeps the trace of the spelling process
  ' History:
  ' 29-12-2007  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Sub TraceReset()
    colTrace.Clear()
  End Sub
  ' ----------------------------------------------------------------------------------
  ' Name :      TraceAdd
  ' Goal :      Add to the string that keeps the trace of the spelling process
  '             First add a newline
  ' History:
  ' 29-12-2007  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Sub TraceAdd(ByVal strIn As String)
    colTrace.Add(strIn)
  End Sub
  ' ----------------------------------------------------------------------------------
  ' Name :      TraceGet
  ' Goal :      Retrieve the string that keeps the trace of the spelling process
  ' History:
  ' 29-12-2007  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function TraceGet() As String
    ' Return the result
    Return colTrace.Text
  End Function

End Class

' All language specific information
Public Class Language
  ' ===================================================================================
  ' Name: Language
  ' Goal: Define a language for usage in POS tagging.
  '       We need to specify the following:
  '       1 - a LIFT dictionary   (XML file containing finite forms)
  '       2 - an affix dictionary (shoebox file)
  ' History:
  ' 24-03-2011  ERK Derived from AggluSpell for Chechen
  ' ===================================================================================
  ' ====================== Publically accessible ======================================
  Public Name As String      ' Name of this language
  Public Ucase As String     ' Upper case characters
  Public Lcase As String     ' Lower case characters
  Public NoCase As String    ' Any letters that don't have a case
  Public CatCase As String   ' Letters used to describe categories
  Public Symbol As Collection ' Contains definition of C, V and other symbols
  Public SfxFile As String   ' Name/location of suffix file
  Public DictFile As String  ' Name/location of dictionary file
  Public Main As DataTable   ' main dictionary entries
  Public Pdg As DataTable    ' paradigm dictionary entries
  ' ********************** LOCAL CONSTANTS ***************************************
  Private Const SUFFIX_FILE = "d:\data files\carla\Spelling\CecAffix.dic"
  Private Const DICT_FILE = "d:\data files\carla\spelling\CecLatPos.dic"
  ' Cyrillic alphabeth
  Private Const WORD_CHAR_CYR_LCASE As String = "абвгдеёжзийклмнопрстуфхцчшщъыьэюяІ"
  Private Const WORD_CHAR_CYR_UCASE As String = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯІ"
  ' Chechen latin alphabeth
  Private Const WORD_CHAR_LAT As String = "abcdefghiìjklmnopqrstuvwxyz'ABCDEFGHIÌJKLMNOPQRSTUVWXYZ-"
  Private Const WORD_CHAR_LAT_LCASE As String = "abcdefghiìjklmnopqrstuvwxyz'-"
  Private Const WORD_CHAR_LAT_UCASE As String = "ABCDEFGHIÌJKLMNOPQRSTUVWXYZ'-"
  ' Characters that can be in a category description
  Private Const WORD_CHAR_CAT As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ/-"
  ' Definition of symbols -- (this is very difficult in Cyrillic)
  Private Const SYMBOL_C_PRE As String = "b,c,c',ch,ch',d,g,gh,hw,h,j,k,k',l,m,n,p,p',r,s,sh,t,t',v,w,x,',z,zh" & _
    ",б,в,г,гІ,д,ж,з,й,к,кІ,л,м,н,п,пІ,р,с,т,тІ,х,хь,хІ,ц,цІ,ч,чІ,ш,ъ,ь,І"
  Private Const SYMBOL_C_AFT As String = "b,c,c',ch,ch',d,g,gh,hw,h,j,k,k',l,m,n,p,p',r,s,sh,t,t',v,w,x,',z,zh" & _
    ",б,в,г,гІ,д,ж,з,й,к,кІ,л,м,н,п,пІ,р,с,т,тІ,х,хь,хІ,ц,цІ,ч,чІ,ш,ъ,ь,І,ю,юь,юьй,я,яь"
  Private Const SYMBOL_V_PRE As String = "yy,ae,ii,ie,ye,uo,y,aa,a,ee,e,i,oo,o,uu,u" & _
    ",уьй,юьй,аь,ий,иэ,оь,уо,уь,юь,яь,ā,а,ē,е,и,ō,о,у,ӯ,э̄,э,ю̄,ю,я̄,я"
  Private Const SYMBOL_V_AFT As String = "a,ae,e,i,ii,o,u,uo,ye,e,ie,y,yy" & _
    ",а,аь,е,и,ий,о,у,уо,оь,э,иэ,уь,уьй"
  Private Const SYMBOL_D_PRE As String = "b,v,d,j"
  Private Const SYMBOL_D_AFT As String = "b,v,d,j"

  ' Definition of symbols -- (this is very difficult in Cyrillic)
  'Private Const SYMBOL_C_PRE As String = "б,в,г,гІ,д,ж,з,й,к,кІ,л,м,н,п,пІ,р,с,т,тІ,х,хь,хІ,ц,цІ,ч,чІ,ш,ъ,ь,І"
  'Private Const SYMBOL_C_AFT As String = "б,в,г,гІ,д,ж,з,й,к,кІ,л,м,н,п,пІ,р,с,т,тІ,х,хь,хІ,ц,цІ,ч,чІ,ш,ъ,ь,І,ю,юь,юьй,я,яь"
  'Private Const SYMBOL_V_PRE As String = "уьй,юьй,аь,ий,иэ,оь,уо,уь,юь,яь,ā,а,ē,е,и,ō,о,у,ӯ,э̄,э,ю̄,ю,я̄,я"
  'Private Const SYMBOL_V_AFT As String = "а,аь,е,и,ий,о,у,уо,оь,э,иэ,уь,уьй"

  Private bInit As Boolean = False
  Private bIsCyrillic As Boolean = False
  ' ----------------------------------------------------------------------------------
  ' Name :      New
  ' Goal :      Read language specific information from initialisation file
  ' Note :      For test purposes use constants...
  ' History:
  ' 20-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Sub New(ByVal strFullName As String, ByVal strDictFile As String,
                      ByVal strAffxFile As String, Optional ByVal bIsCyrillic As Boolean = False)
    Dim arThis() As String

    Try
      With Me
        .Name = strFullName
        .DictFile = strDictFile
        .SfxFile = strAffxFile
        .bIsCyrillic = bIsCyrillic
        If (bIsCyrillic) Then
          .Lcase = WORD_CHAR_CYR_LCASE
          .Ucase = WORD_CHAR_CYR_UCASE
        Else
          .Lcase = WORD_CHAR_LAT_LCASE
          .Ucase = WORD_CHAR_LAT_UCASE
        End If
        .NoCase = ""
        .CatCase = WORD_CHAR_CAT
        ' Determine the symbols
        .Symbol = New Collection
        arThis = Split(SYMBOL_C_PRE, ",")
        .Symbol.Add(arThis, "C1")
        arThis = Split(SYMBOL_C_AFT, ",")
        .Symbol.Add(arThis, "C2")
        arThis = Split(SYMBOL_V_PRE, ",")
        .Symbol.Add(arThis, "V1")
        arThis = Split(SYMBOL_V_AFT, ",")
        .Symbol.Add(arThis, "V2")
        arThis = Split(SYMBOL_D_PRE, ",")
        .Symbol.Add(arThis, "D1")
        arThis = Split(SYMBOL_D_AFT, ",")
        .Symbol.Add(arThis, "D2")
      End With
      bInit = True
    Catch ex As Exception
      ' Show error
      MsgBox("modLanguage/LangInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ----------------------------------------------------------------------------------
  ' Name :      IsVowel
  ' Goal :      Determine of the ONE character passed on is a vowel
  ' History:
  ' 15-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function IsVowel(ByVal strIn As String) As Boolean
    Dim intI As Integer
    Dim arVowel() As String

    Try
      arVowel = Me.Symbol("V1")
      For intI = 0 To UBound(arVowel)
        If (arVowel(intI) = Left(strIn, 1)) Then
          IsVowel = True
          Exit Function
        End If
      Next
      IsVowel = False
    Catch ex As Exception
      ' Show error
      MsgBox("modLanguage/IsVowel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      EndsOnVowel
  ' Goal :      See if the last character(s) of strIn are a vowel
  ' History:
  ' 21-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function EndsOnVowel(ByVal strIn As String, Optional ByRef strLet As String = "") As Boolean
    Dim intI As Integer
    Dim intLen As Integer
    Dim arVowel() As String

    Try
      ' Take the "preceding vowel"  category
      arVowel = Me.Symbol("V1")
      For intI = 0 To UBound(arVowel)
        intLen = Len(arVowel(intI))
        If (arVowel(intI) = Right(strIn, intLen)) Then
          strLet = arVowel(intI)
          EndsOnVowel = True
          Exit Function
        End If
      Next
      EndsOnVowel = False
    Catch ex As Exception
      ' Show error
      MsgBox("modLanguage/EndsOnVowel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      EndsOnCons
  ' Goal :      See if the last character(s) of strIn are a consonant
  ' History:
  ' 21-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function EndsOnCons(ByVal strIn As String, Optional ByRef strLet As String = "") As Boolean
    Dim intI As Integer
    Dim intLen As Integer
    Dim arCons() As String

    Try
      ' Take the "preceding consonant"  category
      arCons = Me.Symbol("C1")
      For intI = 0 To UBound(arCons)
        intLen = Len(arCons(intI))
        If (arCons(intI) = Right(strIn, intLen)) Then
          strLet = arCons(intI)
          EndsOnCons = True
          Exit Function
        End If
      Next
      EndsOnCons = False
    Catch ex As Exception
      ' Show error
      MsgBox("modLanguage/EndsOnCons error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      EndsOnClass
  ' Goal :      See if the last character(s) of strIn are a class marker
  ' History:
  ' 21-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function EndsOnClass(ByVal strIn As String, Optional ByRef strLet As String = "") As Boolean
    Dim intI As Integer
    Dim intLen As Integer
    Dim arClass() As String

    Try
      ' Take the "preceding class marker"  category
      arClass = Me.Symbol("D1")
      For intI = 0 To UBound(arClass)
        intLen = Len(arClass(intI))
        If (arClass(intI) = Right(strIn, intLen)) Then
          strLet = arClass(intI)
          EndsOnClass = True
          Exit Function
        End If
      Next
      EndsOnClass = False
    Catch ex As Exception
      ' Show error
      MsgBox("modLanguage/EndsOnClass error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ----------------------------------------------------------------------------------
  ' Name :      IsWordChar
  ' Goal :      Is the first character of the string strIn a word-character?
  ' History:
  ' 14-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function IsWordChar(ByVal strIn As String) As Boolean
    IsWordChar = (InStr(Me.Lcase & Me.Ucase & Me.NoCase, Left(strIn, 1)) > 0)
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      IsCatChar
  ' Goal :      Is the first character of the string strIn a word-character?
  ' History:
  ' 14-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function IsCatChar(ByVal strIn As String) As Boolean
    IsCatChar = (InStr(Me.CatCase, Left(strIn, 1)) > 0)
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      AdaptOrthography
  ' Goal :      Perform language-dependant orthography adaptations
  ' History:
  ' 04-09-2015  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function AdaptOrthography(ByVal strIn As String) As String
    Try
      Select Case Me.Name
        Case "lbe_Cyr", "lbe-Cyr"
          strIn = strIn.Replace("I", "І")
      End Select
      ' Return the result
      Return strIn
    Catch ex As Exception
      ' Show error
      MsgBox("Language/AdaptOrthography error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      GetLcase
  ' Goal :      Convert every character from whatever case into lower case
  '               taking into account the specifications of the language!
  ' History:
  ' 20-07-2006  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function GetLcase(ByVal strIn As String) As String
    Dim intPos As Integer
    Dim strLet As String
    Dim intCase As Integer
    Dim strOut As String = ""

    Try
      'Examine every letter
      For intPos = 1 To Len(strIn)
        ' Get the letter we are looking at
        strLet = Mid(strIn, intPos, 1)
        ' Convert this one letter
        intCase = InStr(Me.Lcase, strLet)
        ' Is it already a lower case letter?
        If (intCase = 0) Then
          'Is it an upper case letter?
          intCase = InStr(Me.Ucase, strLet)
          If (intCase <> 0) Then
            strLet = Mid(Me.Lcase, intCase, 1)
          End If
        End If
        ' Build up output string
        strOut &= strLet
      Next
      ' Return output string
      GetLcase = strOut
    Catch ex As Exception
      ' Show error
      MsgBox("Language/GetLcase error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function

End Class
