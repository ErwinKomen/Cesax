Imports System.Xml
Module modStatIrat
  ' =========================== LOCAL TYPES ====================================
  Private Structure ValFreq
    Dim Value As Integer  ' The value
    Dim Freq As Integer   ' The frequency
  End Structure
  ' =========================== LOCAL VARIABLES ================================
  Dim arV1() As ValFreq   ' Values for rater #1
  Dim arV2() As ValFreq   ' Values for rater #2
  Dim intV1Num As Integer ' Number of values in [arV1]
  Dim intV2Num As Integer ' Number of values in [arV2]
  Dim colRefType As New StringColl  ' Holds the different referential states
  ' =============================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   DoIrat
  ' Goal:   Compare the current text with a second version of this text and provide measures
  '           for interrater agreement, including Cohen's Kappa
  ' History:
  ' 12-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DoIrat(ByRef pdxFileOne As XmlDocument, ByRef pdxFileTwo As XmlDocument, ByRef strHtml As String) As Boolean
    Dim colHtml As New StringColl ' Here we gather the report
    Dim ndxList1 As XmlNodeList   ' NPs in first file
    Dim ndxList2 As XmlNodeList   ' NPs in second file
    Dim arData() As String        ' Data 
    Dim strState1 As String       ' Referential status in file 1
    Dim strState2 As String       ' Referential status in file 2
    Dim strAnt1 As String         ' Antecedent number in file 1
    Dim strAnt2 As String         ' Antecedent number in file 2
    Dim intAnt1 As Integer        ' Antecedent ID of 1
    Dim intAnt2 As Integer        ' Antecedent ID of 2
    Dim intI As Integer           ' Counter
    Dim intEtreeId As Integer     ' Current ID
    Dim intPtc As Integer         ' Percentage
    Dim intEnd As Integer         ' The actual end of the array
    Dim dblKappa As Double = 0    ' Kappa value
    Dim dblAgr As Double = 0      ' Agreement value

    Try
      ' Validate
      If (pdxFileOne Is Nothing) OrElse (pdxFileTwo Is Nothing) Then Return False
      ' Initialise
      colHtml.Add("<html><body><h2>Interrater agreement</h2><p>File = " & _
              IO.Path.GetFileNameWithoutExtension(strCurrentFile) & "<p>")
      ' Find first and second file information
      ndxList1 = pdxFileOne.SelectNodes("//eTree[tb:matches(@Label, '" & strNPsourceTypes & "')]", conTb)
      ndxList2 = pdxFileTwo.SelectNodes("//eTree[tb:matches(@Label, '" & strNPsourceTypes & "')]", conTb)
      ' Check sizes
      If (ndxList1.Count <> ndxList2.Count) Then Logging("DoIrat: texts have different sizes") : Return False
      ' Make room for the data
      ReDim arData(0 To ndxList1.Count - 1)
      ' Gather the data
      For intI = 0 To ndxList1.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ ndxList1.Count
        Status("Irat " & intPtc & "%", intPtc)
        ' Collect the information
        intEtreeId = ndxList1(intI).Attributes("Id").Value
        strState1 = RefStateId(CorefAttr(ndxList1(intI), "RefType"))
        strState2 = RefStateId(CorefAttr(ndxList2(intI), "RefType"))
        intAnt1 = CorefDstId(ndxList1(intI))
        intAnt2 = CorefDstId(ndxList2(intI))
        strAnt1 = IIf(intAnt1 < 0, "0", Math.Abs(intAnt1 - intEtreeId))
        strAnt2 = IIf(intAnt2 < 0, "0", Math.Abs(intAnt2 - intEtreeId))
        ' Add the information to the data array
        arData(intI) = intEtreeId & ";" & strState1 & ";" & strAnt1 & ";" & strState2 & ";" & strAnt2
      Next intI
      ' Eat the array backwards
      intEnd = CommonEnd(arData)
      ReDim Preserve arData(0 To intEnd)
      ' Start output table
      colHtml.Add("<table><tr><td>Type</td><td>Agreement</td><td>Cohen's Kappa</td></tr>")
      ' Calculate the kappa value for value #1
      If (CohensKappa(arData, 1, 3, dblKappa, dblAgr)) Then
        ' Add kappa and agreement information
        colHtml.Add("<tr><td>Antecedent</td><td>" & Format(dblAgr, "0.0000") & "</td><td>" & Format(dblKappa, "0.0000") & "</td></tr>")
      End If
      ' Calculate the kappa value for value #1
      If (CohensKappa(arData, 2, 4, dblKappa, dblAgr)) Then
        ' Add kappa and agreement information
        colHtml.Add("<tr><td>Referential type</td><td>" & Format(dblAgr, "0.0000") & "</td><td>" & Format(dblKappa, "0.0000") & "</td></tr>")
      End If
      ' Give other information
      colHtml.Add("</table>")
      colHtml.Add("<p>Other statistics</p>")
      colHtml.Add("<table>")
      colHtml.Add("<tr><td>Total number of NPs evaluated:</td><td>" & intEnd + 1 & "</td></tr>")
      ' Finish report
      colHtml.Add("</table></body></html>")
      strHtml = colHtml.Text
      ' Return success
      Return True
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("modStatIrat/DoIrat error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return empty...
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   RefStateId
  ' Goal:   Return a numerical ID of the current state
  ' History:
  ' 12-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function RefStateId(ByVal strState As String) As String
    Dim intId As Integer  ' ID of this state

    Try
      ' make sure we add the "zero" type
      If (colRefType.Count = 0) Then colRefType.AddUnique("(none)")
      ' Check what we get
      If (strState = "") Then
        Return "0"
      Else
        intId = colRefType.Find(strState)
        If (intId < 0) Then
          ' Add it
          colRefType.AddUnique(strState)
        End If
        Return colRefType.Find(strState)
      End If
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("modStatIrat/RefStateId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return empty...
      Return "0"
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CohensKappa
  ' Goal:   Look for the line with first value [strNode] in [arIn]
  '         Delimiter within one line is [strDelim]
  ' Input:  arData() is a string array where each line contains integer data 
  '           separated by ";" signs
  '         intCol1 and intCol2 define which columns in [arData] belong to 
  '           rater #1 and rater #2
  ' Return: dblKappa is Cohen's kappa (0 ... 1)
  '         dblAgr is the agreement (0 ... 1)
  ' History:
  ' 24-06-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CohensKappa(ByRef arData() As String, ByVal intCol1 As Integer, _
      ByVal intCol2 As Integer, ByRef dblKappa As Double, ByRef dblAgr As Double) As Boolean
    Dim arLine() As String  ' The values of one line
    Dim intI As Integer     ' Counter
    Dim intJ As Integer     ' Counter
    Dim intN As Integer     ' Number of data points
    Dim intM As Integer     ' Intermediate number
    Dim intVal1 As Integer  ' Value for rater #1
    Dim intVal2 As Integer  ' Value for rater #2
    Dim dblP_0 As Double    ' P0 from the Kappa formula = percentage agreement
    Dim dblP_c As Double    ' Pc from the Kappa formula: k=(P0-Pc)/(1-Pc)

    Try
      ' Initialise the Value sets
      intN = arData.Length
      ' Adapt [intN] for empty elements at the end of the data array
      While (arData(intN - 1) = "") OrElse (InStr(arData(intN - 1), ";") = 0)
        intN -= 1
      End While
      ' Initialise valFreq sets
      ValFreqClear(intN)
      intM = 0
      ' Visit all points
      For intI = 1 To intN
        ' Get the data for this line
        arLine = Split(arData(intI - 1), ";")
        ' Get the values for rater #1 and rater #2
        intVal1 = CInt(arLine(intCol1)) : intVal2 = CInt(arLine(intCol2))
        ' Put these values in the respective arrays
        IncrVal(arV1, intVal1) : IncrVal(arV2, intVal2)
        ' Keep track of agreement
        If (intVal1 = intVal2) Then intM += 1
      Next intI
      ' Calculate P0
      dblP_0 = intM / intN
      ' Calculate Sum for Pc
      intM = 0
      For intI = 1 To intN
        intJ = GetIndex(arV2, arV1(intI).Value, intN)
        If (intJ > 0) Then
          ' Keep track of sum
          intM += arV1(intI).Freq * arV2(intJ).Freq
        End If
      Next intI
      ' Calculate total Pc
      dblP_c = intM / (intN * intN)
      ' Pass back Kappa and Agreement
      dblKappa = (dblP_0 - dblP_c) / (1 - dblP_c)
      dblAgr = dblP_0
      ' Return success
      Return True
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("modStatIrat/CohensKappa error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return empty...
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetIndex
  ' Goal:   Get the index of [intVal] within [arV]
  ' History:
  ' 25-06-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetIndex(ByRef arV() As ValFreq, ByVal intVal As Integer, _
                            ByVal intN As Integer) As Integer
    Dim intI As Integer     ' Counter

    Try
      ' Go through all values
      For intI = 1 To intN
        ' Is this the correct one?
        If (arV(intI).Value = intVal) Then
          ' Return success
          Return intI
        End If
      Next intI
      ' Return failure
      GetIndex = -1
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("modStatIrat/GetIndex error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return empty...
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   ValFreqClear
  ' Goal:   Clear and initialise the ValFreq arrays
  ' History:
  ' 24-06-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub ValFreqClear(ByVal intN As Integer)
    Dim intI As Integer     ' Counter

    Try
      ' Set sizes
      ReDim arV1(0 To intN)
      ReDim arV2(0 To intN)
      ' Visit all members
      For intI = 0 To intN
        ' Clear this member
        arV1(intI).Freq = 0 : arV1(intI).Value = 0
        arV2(intI).Freq = 0 : arV2(intI).Value = 0
      Next intI
      ' Reset numbers
      intV1Num = 0 : intV2Num = 0
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("modStatIrat/ValFreqClear error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   IncrVal
  ' Goal:   Increment the frequency for the value [intVal] in array [arV]
  '         Array [arV] right now has [intNum] members (from 1 to intNum)
  ' History:
  ' 24-06-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub IncrVal(ByRef arV() As ValFreq, ByVal intVal As Integer)
    Dim intI As Integer   ' Counter

    Try
      ' Check all members of [arV]
      For intI = 0 To intV1Num - 1
        ' Does this member contain the value?
        If (arV(intI).Value = intVal) Then
          ' Increment its frequency
          arV(intI).Freq += 1
          ' Exit
          Exit Sub
        End If
      Next intI
      ' Member was not found, so add to [arV]
      With arV(intV1Num)
        .Value = intVal
        .Freq = 1
      End With
      intV1Num += 1
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("modStatIrat/IncrVal error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   CommonEnd
  ' Goal:   Determine where the two files have a common end
  ' History:
  ' 15-06-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CommonEnd(ByVal arThis() As String) As Integer
    Dim arLine() As String        ' One line broken up
    'Dim strEmpty As String = ";0;0;0;0" ' String that indicates an empty line
    Dim intI As Integer           ' Counter
    Dim bBase As Boolean = False  ' The base array ends on ;0;0;0;0
    Dim bAdded As Boolean = False ' The added array ends on ;0;0

    Try
      ' Step backwards from the end onwards
      For intI = arThis.Length - 1 To 0 Step -1
        ' Only process content lines
        If (arThis(intI) <> "") Then
          ' Unpack this line
          arLine = Split(arThis(intI), ";")
          ' Has [bBase] already been set?
          If (Not bBase) Then
            ' Try to set it...
            bBase = (arLine(1) <> "0" OrElse arLine(2) <> "0")
          End If
          ' Has [bAdded] already been set?
          If (Not bAdded) Then
            ' Try to set it...
            bAdded = (arLine(3) <> "0" OrElse arLine(4) <> "0")
          End If
          ' Have both been set?
          If (bBase And bAdded) Then
            ' This is the common end!
            Return intI + 1
          End If
        End If
      Next intI
      ' No common end was found - indicate failure
      Return -1
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("modStatIrat/CommonEnd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return empty...
      Return -1
    End Try
  End Function
End Module
