'Imports SharpEntropy
Imports System.Xml
Module modMaxEnt
  ' ------------------------------------------------------------------------------------
  ' Name:   InitTrain
  ' Goal:   Initialize maximum entropy training
  ' History:
  ' 07-11-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function InitTrain() As Boolean
    Try

      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMaxEnt/InitTrain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DoTrainCoref
  ' Goal:   Train the given [strFile] for coreference
  ' History:
  ' 07-11-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DoTrainCoref(ByVal strInFile As String) As Boolean
    'Dim ndxThis As XmlNode    ' Current NP under investigation
    'Dim ndxForest As XmlNode  ' The forest for this line
    'Dim ndxTarget As XmlNode  ' Destination NP
    'Dim strRefType As String  ' The kind of coreference relation (if any)
    'Dim strGrRole As String   ' The grammatical role of the source and destination
    'Dim strNpType As String   ' The NP type of the source and the destination
    'Dim strRelation As String ' NP or GrRole relation between source and target
    'Dim intPtc As Integer     ' Percentage
    'Dim intNp As Integer      ' Counter for the NP's within one <forest> line
    'Dim intF As Integer       ' One forest item

    Try
      ' Validate
      If (Not System.IO.File.Exists(strInFile)) Then
        Status("Could not find file " & strInFile)
        Return False
      End If
      ' Show we are reading
      Status("Reading file " & strInFile)
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxCurrentFile)) Then
        ' Initialise [pdxList] and [arLeftId()]
        If (Not InitCurrentFile()) Then Return False
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMaxEnt/DoTrainCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

End Module
