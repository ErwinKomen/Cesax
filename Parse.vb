Imports System.Xml
Public Class Parse
  ' ============================================ CLASS INFORMATION ===========================================
  ' Name :  Parse
  ' Goal :  Stores all the information pertaining to one particular parse
  '         A parse consists of:
  '         - The total resulting word
  '         - The net result category of the word (e.g. VbPrsN)
  '         - The net result Part-of-Speech (e.g. N-GEN)
  '         - A starting lexeme
  '         - Category of the lexeme
  '         - A list of suffixes ordered from left to right
  '         - Source and destination categories of these suffixes
  ' History:
  ' 08-04-2011  ERK Created
  ' ==========================================================================================================

  ' ====================== LOCAL TYPES =======================================================================
  Private Structure WordInfo
    Dim Word As String    ' The actual word
    Dim Lemma As String   ' The lemma
    Dim Gloss As String   ' English gloss (if available)
    Dim CatSrc As String  ' The input category (eg VbPrsN)
    Dim Cat As String     ' The output category
    Dim POS As String     ' The part-of-speech (e.g. N-GEN)
    Dim Feat As String    ' Feature(s) associated with this word
  End Structure
  ' ====================== LOCAL CONSTANTS ===================================================================
  Private MAX_SFX As Integer = 100
  ' ====================== LOCAL VARIABLES ===================================================================
  Private objWord As New WordInfo         ' The word we are considering in this parse
  Private objLex As New WordInfo          ' The starting lexeme, after which suffixes are glued
  Private intSize As Integer = 0          ' Size of suffix array
  Private arSfx(0 To MAX_SFX) As WordInfo ' Array of suffixes
  Private bHasParse As Boolean = False    ' Whether we have at least some basic parse information
  Private intParseId As Integer = 0       ' ID of this parse
  ' ==========================================================================================================
  Public Property ParseId() As Integer
    Get
      ' Return the local id
      Return intParseId
    End Get
    Set(ByVal value As Integer)
      ' Set the local id
      intParseId = value
    End Set
  End Property
  ' ----------------------------------------------------------------------------------
  ' Name :      HasParse
  ' Goal :      Check if this contains some parse
  ' History:
  ' 08-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public ReadOnly Property HasParse() As Boolean
    Get
      Return bHasParse
    End Get
  End Property
  ' ----------------------------------------------------------------------------------
  ' Name :      POS
  ' Goal :      Return the resulting POS of the total word
  ' History:
  ' 08-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public ReadOnly Property POS() As String
    Get
      Dim intLexPos As Integer    ' Position within string

      ' Is there a total result?
      If (objWord.POS = "") Then
        ' Calculate it
        If (intSize > 0) Then
          ' Get the POS from the category of the last suffix
          ' objWord.POS = CatToPOS(objWord.Word, "(no definition available)", arSfx(intSize - 1).Cat, "")
          objWord.POS = arSfx(intSize - 1).Cat
        End If
      End If
      ' Check for the presence of a + sign in the lexpos
      intLexPos = InStr(objLex.POS, "+")
      If (intLexPos > 0) AndAlso (InStr(objWord.POS, "+") = 0) Then
        ' Adapt my own POS
        objWord.POS = Left(objLex.POS, intLexPos) & objWord.POS
      End If
      ' Return the total resulting POS
      Return objWord.POS
    End Get
  End Property
  ' ----------------------------------------------------------------------------------
  ' Name :      LexPos
  ' Goal :      Return the POS of the lexeme
  ' History:
  ' 17-05-2012  ERK Created
  ' ----------------------------------------------------------------------------------
  Public ReadOnly Property LexPos() As String
    Get
      Return objLex.POS
    End Get
  End Property
  ' ----------------------------------------------------------------------------------
  ' Name :      Cat
  ' Goal :      Return the resulting Category of the total word
  ' History:
  ' 09-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public ReadOnly Property Cat() As String
    Get
      ' Is there a total result?
      If (objWord.Cat = "") Then
        ' Calculate it
        If (intSize > 0) Then
          ' Get the POS from the category of the last suffix
          objWord.Cat = arSfx(intSize - 1).Cat
        End If
      End If
      ' Return the total resulting POS
      Return objWord.Cat
    End Get
  End Property
  Public ReadOnly Property Word() As String
    Get
      Return objWord.Word
    End Get
  End Property
  Public ReadOnly Property Lex() As String
    Get
      Return objLex.Word
    End Get
  End Property
  Public ReadOnly Property Lemma() As String
    Get
      Return objLex.Lemma
    End Get
  End Property
  ' ----------------------------------------------------------------------------------
  ' Name :      Feat
  ' Goal :      Get or set features for this word
  ' History:
  ' 08-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Property Feat() As String
    Get
      ' Get the feature
      Return objWord.Feat
    End Get
    Set(ByVal value As String)
      objWord.Feat = value
    End Set
  End Property
  ' ----------------------------------------------------------------------------------
  ' Name :      Morph
  ' Goal :      The break-up into morphemes of the result of this parse
  ' History:
  ' 08-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public ReadOnly Property Morph() As String
    Get
      Dim strOut As String = "" ' Resulting morphemic representation
      Dim intI As Integer       ' Counter

      Try
        ' Start with the <Lex>
        strOut = objLex.Word & "/" & objLex.Cat
        ' Visit all suffixe
        For intI = 0 To intSize - 1
          ' Add the information for this suffix
          strOut &= "-" & arSfx(intI).Word & "/" & arSfx(intI).Cat
        Next intI
        ' Return the result
        Return strOut
      Catch ex As Exception
        ' Show error
        MsgBox("Parse/Morph error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
        ' Return failure
        Return ""
      End Try
    End Get
  End Property
  ' ----------------------------------------------------------------------------------
  ' Name :      Def
  ' Goal :      The dictionary definition of the main entry
  ' History:
  ' 24-07-2015  ERK Created
  ' ----------------------------------------------------------------------------------
  Public ReadOnly Property Def() As String
    Get
      Return objLex.Gloss
    End Get
  End Property
  ' ----------------------------------------------------------------------------------
  ' Name :      SetWord
  ' Goal :      Set the total word
  ' History:
  ' 08-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function SetWord(ByVal strWord As String, ByVal strGloss As String, ByVal strCat As String, _
                           ByVal strPOS As String) As Boolean
    Try
      ' Set the features
      With objWord
        .Word = strWord : .Gloss = strGloss : .CatSrc = "" : .Cat = strCat : .POS = strPOS
        .Feat = ""
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      MsgBox("Parse/SetWord error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      SetLex
  ' Goal :      Set the starting lexeme
  ' History:
  ' 08-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function SetLex(ByVal strLemma As String, ByVal strWord As String, ByVal strGloss As String, _
                         ByVal strCat As String, ByVal strPOS As String) As Boolean
    Try
      ' Set the features
      With objLex
        .Lemma = strLemma : .Word = strWord : .Gloss = strGloss : .CatSrc = "" : .Cat = strCat : .POS = strPOS
      End With
      ' We now have at least something
      bHasParse = True
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      MsgBox("Parse/SetLex error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      PushSfx
  ' Goal :      Add one suffix to the stack
  ' History:
  ' 08-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function PushSfx(ByVal strSuffix As String, ByVal strGloss As String, ByVal strCatSrc As String, _
                     ByVal strCat As String, ByVal strPOS As String) As Boolean
    Try
      ' Validate
      If (intSize >= MAX_SFX - 1) Then Return False
      ' Add the suffix
      With arSfx(intSize)
        .Word = strSuffix
        .Gloss = strGloss
        .CatSrc = strCatSrc
        .Cat = strCat
        .POS = strPOS
      End With
      ' Make room
      intSize += 1
      ' Adapt the total resulting Cat and POS of the total word
      With objWord
        If (strCat <> "") Then .Cat = strCat
        If (strPOS <> "") Then .POS = strPOS
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      MsgBox("Parse/PushSfx error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      PopSfx
  ' Goal :      Delete the suffix from the stack
  ' History:
  ' 08-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Function PopSfx() As Boolean
    Try
      ' Validate
      If (intSize < 1) Then
        ' Reset to zero
        intSize = 0
      Else
        ' Just decrement the counter
        intSize -= 1
      End If
      Return True
    Catch ex As Exception
      ' Show error
      MsgBox("Parse/PopSfx error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------
  ' Name :      Clear
  ' Goal :      Clear whatever needs clearing
  ' History:
  ' 08-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------
  Public Sub Clear()
    Try
      intSize = 0
      bHasParse = False
    Catch ex As Exception
      ' Show error
      MsgBox("Parse/Clear error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

End Class
