Imports System.Xml
Module modXmlNode
  ' =========================================================================================================
  ' Name: modXmlNode
  ' Goal: This module implements additional features for working with an XmlDocument and XmlNode elements
  ' History:
  ' 22/sep/2010 ERK Created
  ' ========================================== LOCAL VARIABLES ==============================================
  Private pdxDoc As XmlDocument = Nothing   ' The XML document serving as basis
  Private cmbThis As New ComboBox           ' Combobox for SFM
  Private strSpace As String = " ,.<>?/;:\|[{]}=+-_)(*&^%$#@!~`" & """" & vbTab & vbCrLf
  Private strNs As String = ""              ' Possible namespace URI
  ' =========================================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   SetXmlDocument
  ' Goal:   Set the local copy of the xml document
  ' History:
  ' 22-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SetXmlDocument(ByRef pdxThis As XmlDocument, Optional ByVal nsThis As String = "")
    ' Set the document to the indicated one
    pdxDoc = pdxThis
    ' If (nsThis <> "") Then strNs = nsThis
    strNs = nsThis
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   AddXmlChild
  ' Goal:   Make a new XmlNode element of type [strTag] using the [arValue] values
  '         These values consist of:
  '         (a) itemname
  '         (b) itemvalue
  '         (c) itemtype: "attribute" or "child"
  '         Append this node as child under [ndxParent]
  ' Return: The XmlNode element that has been made is returned
  ' History:
  ' 22-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddXmlChild(ByRef ndxParent As XmlNode, ByVal strTag As String, _
                            ByVal ParamArray arValue() As String) As XmlNode
    Dim ndxThis As XmlNode        ' Working node
    Dim ndxChild As XmlNode       ' Child node
    'Dim ndxNext As XmlNode        ' Working node for finding the correct sibling
    Dim atxChild As XmlAttribute  ' The attribute we are looking for
    Dim intI As Integer           ' Counter

    Try
      ' Validate (NB: we DO allow empty parents)
      If (strTag = "") OrElse (pdxDoc Is Nothing) Then Return Nothing
      ' Make a new XmlNode in the local XML document
      If (strNs = "") Then
        ndxThis = pdxDoc.CreateNode(XmlNodeType.Element, strTag, Nothing)
      Else
        ndxThis = pdxDoc.CreateNode(XmlNodeType.Element, strTag, strNs)
      End If
      ' Validate
      If (ndxThis Is Nothing) Then Return Nothing
      ' Do we have a parent?
      If (ndxParent Is Nothing) Then
        ' Take the document as starting point
        pdxDoc.AppendChild(ndxThis)
      Else
        ' Just append it
        ndxParent.AppendChild(ndxThis)
      End If
      ' Walk through the values
      For intI = 0 To UBound(arValue) Step 3
        ' Action depends on the type of value
        Select Case arValue(intI + 2)
          Case "attribute"
            ' Create attribute
            atxChild = pdxDoc.CreateAttribute(arValue(intI))
            ' Fillin value of this attribute
            atxChild.Value = arValue(intI + 1)
            ' Append attribute to this node
            ndxThis.Attributes.Append(atxChild)
          Case "child"
            ' Create this node
            If (strNs = "") Then
              ndxChild = pdxDoc.CreateNode(XmlNodeType.Element, arValue(intI), Nothing)
            Else
              ndxChild = pdxDoc.CreateNode(XmlNodeType.Element, arValue(intI), strNs)
            End If
            ' Fill in the value of this node
            ndxChild.InnerText = arValue(intI + 1)
            ' Append this node as child
            ndxThis.AppendChild(ndxChild)
          Case Else
            ' There is no other option yet, so return failure
            Return Nothing
        End Select
      Next intI
      ' Return the new node
      Return ndxThis
    Catch ex As Exception
      ' Warn user
      MsgBox("modXmlNode/AddXmlChild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddXmlAttribute
  ' Goal:   Add an attribute to the indicated node
  ' History:
  ' 23-02-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddXmlAttribute(ByRef pdxThis As XmlDocument, ByRef ndxThis As XmlNode, _
                  ByVal strAttrName As String, Optional ByVal strAttrValue As String = "") As Boolean
    Dim atxChild As XmlAttribute  ' The attribute we are looking for

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Check if hte attribute is already there
      If (ndxThis.Attributes(strAttrName) Is Nothing) Then
        ' It is not there, so add it
        atxChild = pdxThis.CreateAttribute(strAttrName)
        ' Append attribute to this node
        ndxThis.Attributes.Append(atxChild)
      End If
      ' Optionally give the value
      If (strAttrValue <> "") Then
        ndxThis.Attributes(strAttrName).Value = strAttrValue
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn user
      MsgBox("modXmlNode/AddXmlAttribute error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddXmlChildSorted
  ' Goal:   Make a new XmlNode element of type [strTag] using the [arValue] values
  '         These values consist of:
  '         (a) itemname
  '         (b) itemvalue
  '         (c) itemtype: "attribute" or "child"
  '         Append this node as child under [ndxParent]
  ' Return: The XmlNode element that has been made is returned
  ' Note:   The new element is appended in an alphabetical way
  ' History:
  ' 22-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddXmlChildSorted(ByRef ndxParent As XmlNode, ByVal strTag As String, _
            ByVal ParamArray arValue() As String) As XmlNode
    Dim ndxThis As XmlNode          ' Working node
    Dim ndxChild As XmlNode         ' Child node
    Dim atxChild As XmlAttribute    ' The attribute we are looking for
    Dim ndxNext As XmlNode          ' Working node
    Dim intI As Integer             ' Counter
    Dim strAttrName As String = ""  ' Name of attribute for sorting
    Dim strAttrVal As String = ""   ' Value of attribute for sorting

    Try
      ' Validate (NB: we DO allow empty parents)
      If (strTag = "") OrElse (pdxDoc Is Nothing) Then Return Nothing
      ' Make a new XmlNode in the local XML document
      ndxThis = pdxDoc.CreateNode(XmlNodeType.Element, strTag, Nothing)
      ' Validate
      If (ndxThis Is Nothing) Then Return Nothing
      ' Walk through the values
      For intI = 0 To UBound(arValue) Step 3
        ' Action depends on the type of value
        Select Case arValue(intI + 2)
          Case "attribute"
            ' Create attribute
            atxChild = pdxDoc.CreateAttribute(arValue(intI))
            ' Fillin value of this attribute
            atxChild.Value = arValue(intI + 1)
            If (strAttrName = "") Then
              strAttrName = arValue(intI)
              strAttrVal = arValue(intI + 1)
            End If
            ' Append attribute to this node
            ndxThis.Attributes.Append(atxChild)
          Case "child"
            ' Create this node
            ndxChild = pdxDoc.CreateNode(XmlNodeType.Element, arValue(intI), Nothing)
            ' Fill in the value of this node
            ndxChild.InnerText = arValue(intI + 1)
            ' Append this node as child
            ndxThis.AppendChild(ndxChild)
          Case Else
            ' There is no other option yet, so return failure
            Return Nothing
        End Select
      Next intI
      ' Do we have a parent?
      If (ndxParent Is Nothing) Then
        ' Take the document as starting point
        pdxDoc.AppendChild(ndxThis)
      Else
        ' Do we have an attribute name & value?
        If (strAttrName = "") Then
          ' Just append it
          ndxParent.AppendChild(ndxThis)
        Else
          ' Start and look for the correct place
          ndxNext = ndxParent.FirstChild
          While (Not ndxNext Is Nothing)
            ' Check if this is the correct one
            If (Not ndxNext.Attributes(strAttrName) Is Nothing) Then
              If (strAttrVal < ndxNext.Attributes(strAttrName).Value) Then
                ' Insert before [ndxNext]
                ndxParent.InsertBefore(ndxThis, ndxNext)
                ' Leave the While
                Exit While
              End If
            End If
            ' Go to the next sibling
            ndxNext = ndxNext.NextSibling
          End While
          ' Check if it got inserted somewhere
          If (ndxNext Is Nothing) Then
            ' Just add it
            ndxParent.AppendChild(ndxThis)
          End If
        End If
      End If
      ' Return the new node
      Return ndxThis
    Catch ex As Exception
      ' Warn user
      MsgBox("modXmlNode/AddXmlChildSorted error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddXmlChildAfter
  ' Goal:   Make a new XmlNode element of type [strTag] using the [arValue] values
  '         These values consist of:
  '         (a) itemname
  '         (b) itemvalue
  '         (c) itemtype: "attribute" or "child"
  '         Append this node as child under [ndxParent]
  '         If possible add it AFTER [ndxAfter]
  ' Return: The XmlNode element that has been made is returned
  ' Note:   The new element is appended in an alphabetical way
  ' History:
  ' 22-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddXmlChildAfter(ByRef ndxParent As XmlNode, ByRef ndxAfter As XmlNode, _
        ByVal strTag As String, ByVal ParamArray arValue() As String) As XmlNode
    Dim ndxThis As XmlNode        ' Working node
    Dim ndxChild As XmlNode       ' Child node
    'Dim ndxNext As XmlNode        ' Working node for finding the correct sibling
    Dim atxChild As XmlAttribute  ' The attribute we are looking for
    Dim intI As Integer           ' Counter

    Try
      ' Validate (NB: we DO allow empty parents)
      If (pdxDoc Is Nothing) Then Return Nothing
      ' Action depends on tag
      If (strTag = "") Then
        ' Make a new text node
        ndxThis = pdxDoc.CreateTextNode("")
      Else
        ' Make a new XmlNode in the local XML document
        ndxThis = pdxDoc.CreateNode(XmlNodeType.Element, strTag, Nothing)
      End If
      ' Validate
      If (ndxThis Is Nothing) Then Return Nothing
      ' Do we have a parent?
      If (ndxParent Is Nothing) Then
        ' Take the document as starting point
        pdxDoc.AppendChild(ndxThis)
      ElseIf (Not ndxAfter Is Nothing) Then
        ' Add it after the indicated node
        ndxParent.InsertAfter(ndxThis, ndxAfter)
      Else
        ' Just append it
        ndxParent.AppendChild(ndxThis)
      End If
      ' Walk through the values
      For intI = 0 To UBound(arValue) Step 3
        ' Action depends on the type of value
        Select Case arValue(intI + 2)
          Case "attribute"
            ' Create attribute
            atxChild = pdxDoc.CreateAttribute(arValue(intI))
            ' Fillin value of this attribute
            atxChild.Value = arValue(intI + 1)
            ' Append attribute to this node
            ndxThis.Attributes.Append(atxChild)
          Case "child"
            ' Create this node
            ndxChild = pdxDoc.CreateNode(XmlNodeType.Element, arValue(intI), Nothing)
            ' Fill in the value of this node
            ndxChild.InnerText = arValue(intI + 1)
            ' Append this node as child
            ndxThis.AppendChild(ndxChild)
          Case Else
            ' There is no other option yet, so return failure
            Return Nothing
        End Select
      Next intI
      ' Return the new node
      Return ndxThis
    Catch ex As Exception
      ' Warn user
      MsgBox("modXmlNode/AddXmlChildAfter error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddXmlChildBefore
  ' Goal:   Make a new XmlNode element of type [strTag] using the [arValue] values
  '         These values consist of:
  '         (a) itemname
  '         (b) itemvalue
  '         (c) itemtype: "attribute" or "child"
  '         Append this node as child under [ndxParent]
  '         If possible add it BEFORE [ndxBefore]
  ' Return: The XmlNode element that has been made is returned
  ' History:
  ' 22-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddXmlChildBefore(ByRef ndxParent As XmlNode, ByRef ndxBefore As XmlNode, _
        ByVal strTag As String, ByVal ParamArray arValue() As String) As XmlNode
    Dim ndxThis As XmlNode        ' Working node
    Dim ndxChild As XmlNode       ' Child node
    'Dim ndxNext As XmlNode        ' Working node for finding the correct sibling
    Dim atxChild As XmlAttribute  ' The attribute we are looking for
    Dim intI As Integer           ' Counter

    Try
      ' Validate (NB: we DO allow empty parents)
      If (strTag = "") OrElse (pdxDoc Is Nothing) Then Return Nothing
      ' Make a new XmlNode in the local XML document
      ndxThis = pdxDoc.CreateNode(XmlNodeType.Element, strTag, Nothing)
      ' Validate
      If (ndxThis Is Nothing) Then Return Nothing
      ' Do we have a parent?
      If (ndxParent Is Nothing) Then
        ' Take the document as starting point
        pdxDoc.AppendChild(ndxThis)
      ElseIf (Not ndxBefore Is Nothing) Then
        ' Add it before the indicated node
        ndxParent.InsertBefore(ndxThis, ndxBefore)
      Else
        ' Just append it
        ndxParent.AppendChild(ndxThis)
      End If
      ' Walk through the values
      For intI = 0 To UBound(arValue) Step 3
        ' Action depends on the type of value
        Select Case arValue(intI + 2)
          Case "attribute"
            ' Create attribute
            atxChild = pdxDoc.CreateAttribute(arValue(intI))
            ' Fillin value of this attribute
            atxChild.Value = arValue(intI + 1)
            ' Append attribute to this node
            ndxThis.Attributes.Append(atxChild)
          Case "child"
            ' Create this node
            ndxChild = pdxDoc.CreateNode(XmlNodeType.Element, arValue(intI), Nothing)
            ' Fill in the value of this node
            ndxChild.InnerText = arValue(intI + 1)
            ' Append this node as child
            ndxThis.AppendChild(ndxChild)
          Case Else
            ' There is no other option yet, so return failure
            Return Nothing
        End Select
      Next intI
      ' Return the new node
      Return ndxThis
    Catch ex As Exception
      ' Warn user
      MsgBox("modXmlNode/AddXmlChildBefore error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   FirstChildWithAttr
  ' Goal:   Get the first child of <strTag> with attribute [strParName] = [strParValue]
  ' History:
  ' 22-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function FirstChildWithAttr(ByRef ndxParent As XmlNode, ByVal strTag As String, _
                                ByVal strParName As String, ByVal strParValue As String) As XmlNode
    Dim ndxWork As XmlNode  ' Working node

    Try
      ' Validate
      If (ndxParent Is Nothing) OrElse (strTag = "") OrElse (strParName = "") Then Return Nothing
      ' Go to the first child
      ndxWork = ndxParent.FirstChild
      ' Loop until we have our result
      While (Not ndxWork Is Nothing)
        ' Examine this child: the tagname
        If (ndxWork.Name = strTag) Then
          ' Does it have the parameter?
          If (Not ndxWork.Attributes(strParName) Is Nothing) Then
            ' Check the value of this parameter
            If (ndxWork.Attributes(strParName).Value = strParValue) Then
              ' Found it
              Return ndxWork
            End If
          End If
        End If
        ' Go to the next sibling
        ndxWork = ndxWork.NextSibling
      End While
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Warn user
      MsgBox("modXmlNode/FirstChildWithAttr error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LastChildWithAttr
  ' Goal:   Get the first child of <strTag> with attribute [strParName] = [strParValue]
  ' Note:   If [bOrdered] is set to TRUE, then we only look backwards while there are
  '           entries with attribute [strParName] being larger than [strParValue]
  ' History:
  ' 22-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function LastChildWithAttr(ByRef ndxParent As XmlNode, ByVal strTag As String, _
      ByVal strParName As String, ByVal strParValue As String) As XmlNode
    Dim ndxWork As XmlNode          ' Working node
    Dim strAttrValue As String = "" ' The value of the attribute

    Try
      ' Validate
      If (ndxParent Is Nothing) OrElse (strTag = "") OrElse (strParName = "") Then Return Nothing
      ' Go to the first child
      ndxWork = ndxParent.LastChild
      ' Loop until we have our result
      While (Not ndxWork Is Nothing)
        ' Examine this child: the tagname
        If (ndxWork.Name = strTag) Then
          ' Does it have the parameter?
          If (Not ndxWork.Attributes(strParName) Is Nothing) Then
            ' Get the value of this attribute
            strAttrValue = ndxWork.Attributes(strParName).Value
            ' Check the value of this parameter
            If (strAttrValue = strParValue) Then
              ' Found it
              Return ndxWork
            End If
          End If
        End If
        ' Go to the next sibling
        ndxWork = ndxWork.PreviousSibling
      End While
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Warn user
      MsgBox("modXmlNode/LastChildWithAttr error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LastChild
  ' Goal:   Get the last child of <strTag> 
  ' History:
  ' 29-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function LastChild(ByRef ndxParent As XmlNode, ByVal strTag As String) As XmlNode
    Dim ndxWork As XmlNode          ' Working node
    Dim strAttrValue As String = "" ' The value of the attribute

    Try
      ' Validate
      If (ndxParent Is Nothing) OrElse (strTag = "") Then Return Nothing
      ' Go to the first child
      ndxWork = ndxParent.LastChild
      ' Loop until we have our result
      While (Not ndxWork Is Nothing)
        ' Examine this child: the tagname
        If (ndxWork.Name = strTag) Then
          ' Found it
          Return ndxWork
        End If
        ' Go to the next sibling
        ndxWork = ndxWork.PreviousSibling
      End While
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Warn user
      MsgBox("modXmlNode/LastChild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LastSiblingWithAttr
  ' Goal:   Get the first sibling of <strTag> with attribute > [strParValue] (working backwards)
  ' Note:   If [bFirstChar] is set to TRUE, then we only look backwards while there are
  '           entries starting with the first character of [strParValue]
  ' History:
  ' 22-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function LastSiblingWithAttr(ByRef ndxNext As XmlNode, ByVal strTag As String, _
      ByVal strParName As String, ByVal strParValue As String) As XmlNode
    Dim strAttrValue As String = "" ' The value of the attribute

    Try
      ' Validate
      If (ndxNext Is Nothing) OrElse (strTag = "") OrElse (strParName = "") Then Return Nothing
      ' Loop until we have our result
      While (Not ndxNext Is Nothing)
        ' Examine this child: the tagname
        If (ndxNext.Name = strTag) Then
          ' Does it have the parameter?
          If (Not ndxNext.Attributes(strParName) Is Nothing) Then
            ' Get the value of this attribute
            strAttrValue = ndxNext.Attributes(strParName).Value
            ' Check the value of this parameter
            If (String.Compare(strAttrValue, strParValue, True) < 0) Then
              ' Found it
              Return ndxNext
            End If
            'If (StrComp(strAttrValue, strParValue, CompareMethod.Text) < 0) Then
            '  ' Found it
            '  Return ndxNext
            'End If
          End If
        End If
        ' Go to the next sibling (working backwards)
        ndxNext = ndxNext.PreviousSibling
      End While
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Warn user
      MsgBox("modXmlNode/LastSiblingWithAttr error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SomeChild
  ' Goal:   Get the first child with the indicated [strName]
  ' History:
  ' 06-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function SomeChild(ByRef ndxParent As XmlNode, ByVal strName As String) As XmlNode
    Dim ndxNext As XmlNode        ' Working node

    Try
      ' Validate
      If (ndxParent Is Nothing) Then Return Nothing
      ' Start with the first child
      ndxNext = ndxParent.FirstChild
      ' Check all children
      While (Not ndxNext Is Nothing)
        ' Check condition
        If (ndxNext.Name = strName) Then
          ' Return this
          Return ndxNext
        End If
        ' Go to next sibling
        ndxNext = ndxNext.NextSibling
      End While
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Warn the user
      MsgBox("modXmlNode/SomeChild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   NextSibling
  ' Goal:   Get the first following sibling with the indicated [strName]
  ' History:
  ' 06-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function NextSibling(ByRef ndxParent As XmlNode, ByVal strName As String) As XmlNode
    Dim ndxNext As XmlNode        ' Working node

    Try
      ' Validate
      If (ndxParent Is Nothing) Then Return Nothing
      ' Start with the first sibling
      ndxNext = ndxParent.NextSibling
      ' Check all children
      While (Not ndxNext Is Nothing)
        ' Check condition
        If (ndxNext.Name = strName) Then
          ' Return this
          Return ndxNext
        End If
        ' Go to next sibling
        ndxNext = ndxNext.NextSibling
      End While
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Warn the user
      MsgBox("modXmlNode/NextSibling error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  AddEtreeChild
  ' Goal :  Add an <eTree> child under [ndxParent]
  ' History:
  ' 26-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function AddEtreeChild(ByRef ndxParent As XmlNode, ByVal intId As Integer, ByVal strLabel As String, _
                                  ByVal intSt As Integer, ByVal intEn As Integer) As XmlNode
    Try
      ' Add the child
      Return AddXmlChild(ndxParent, "eTree", "Id", intId, "attribute", _
                                "Label", strLabel, "attribute", _
                                "from", intSt, "attribute", "to", intEn, "attribute")
    Catch ex As Exception
      ' Warn the user
      MsgBox("modXmlNode/AddEtreeChild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  AddFeature
  ' Goal :  Add a <fs> child under [ndxParent], and an <f> child under the <fs> element
  ' History:
  ' 17-05-2012  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function AddFeatureXml(ByRef ndxParent As XmlNode, ByVal strType As String, ByVal strName As String, _
                               ByVal strValue As String) As Boolean
    Dim ndxFs As XmlNode  ' The <fs> child
    Dim ndxF As XmlNode   ' The <f> child

    Try
      ' Validate
      If (ndxParent Is Nothing) Then Return False
      If (strType = "" OrElse strName = "" OrElse strValue = "") Then Return False
      ' Do we already have an <fs> child?
      ndxFs = ndxParent.SelectSingleNode("./child::fs[@type='" & strType & "']")
      If (ndxFs Is Nothing) Then
        ' Add such a child
        ndxFs = AddXmlChild(ndxParent, "fs", "type", strType, "attribute")
      End If
      ' Add an appropriate <f> child
      ndxF = ndxFs.SelectSingleNode("./child::f[@name='" & strName & "']")
      If (ndxF Is Nothing) Then
        ' Create a new child
        If (AddXmlChild(ndxFs, "f", "name", strName, "attribute", "value", strValue, "attribute") Is Nothing) Then Return False
      Else
        ' Adapt the current child
        ndxF.Attributes("value").Value = strValue
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      MsgBox("modXmlNode/AddEtreeChild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  AddEleafChild
  ' Goal :  Add an <eLeaf> child under [ndxParent]
  ' History:
  ' 26-04-2011  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function AddEleafChild(ByRef ndxParent As XmlNode, ByVal strType As String, ByVal strText As String, _
                                  ByVal intSt As Integer, ByVal intEn As Integer) As XmlNode
    Try
      ' Add the child
      Return AddXmlChild(ndxParent, "eLeaf", "Type", strType, "attribute", _
                                "Text", strText, "attribute", _
                                "from", intSt, "attribute", _
                                "to", intEn, "attribute", _
                                "n", 0, "attribute")
    Catch ex As Exception
      ' Warn the user
      MsgBox("modXmlNode/AddEleafChild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       ForestSet()
  ' Goal:       Set the line in <forest> in the indicated language
  ' History:
  ' 16-04-2011  ERK Created
  '---------------------------------------------------------------
  Public Function ForestSet(ByRef ndxThis As XmlNode, ByVal strLang As String, ByVal strValue As String) As Boolean
    Dim ndxWork As XmlNode    ' Working node

    Try
      ' Validate 
      If (ndxThis Is Nothing) OrElse (strLang = "") Then Return ""
      ' Get to the right node
      ndxWork = ndxThis.FirstChild
      While (Not ndxWork Is Nothing)
        ' Is this the one?
        If (ndxWork.Name = "div") Then
          ' Is it the right language?
          If (ndxWork.Attributes("lang").Value = strLang) Then
            ' Get the first child, that should be <seg>
            ndxWork = ndxWork.FirstChild
            ' Test
            If (Not ndxWork Is Nothing) Then
              ' Okay, got you!
              If (ndxWork.InnerText <> strValue) Then
                ' Only change text when needed
                ndxWork.InnerText = strValue
              End If
              Return True
            End If
          End If
        End If
        ' Try next one
        ndxWork = ndxWork.NextSibling
      End While
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      MsgBox("modXmlNode/ForestSet error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   XmlReaderGetNextNode
  ' Goal:   Try to read a next chunk using XmlReader and return the specified node from it
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function XmlReaderGetNextNode(ByRef rdThis As XmlReader, ByRef xmlDoc As XmlDocument, ByVal strTag As String, _
                                       ByRef ndxThis As XmlNode) As Boolean
    Dim strText As String
    Dim bEnd As Boolean

    Try
      ' Validate
      If (rdThis Is Nothing) OrElse (rdThis.EOF) OrElse (xmlDoc Is Nothing) OrElse (strTag = "") Then Return False
      Do
        ' Read until we get the correct element
        strText = rdThis.ReadOuterXml
        If (strText = "") Then Return False
        ' Transform the text into XmlDocument
        xmlDoc.LoadXml(strText)
        ' Get the node with the identified tag from this document
        ndxThis = xmlDoc.SelectSingleNode("./descendant::" & strTag)
        bEnd = (rdThis.EOF) OrElse (ndxThis IsNot Nothing)
      Loop Until bEnd
      ' Return the result
      Return bEnd
    Catch ex As Exception
      ' Show error
      HandleErr("modXmlNode/XmlReaderGetNextNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   XmlRelation
  ' Goal:   Find out what the relation of [ndxThis] is with respect to [ndxBase]
  ' History:
  ' 01-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function XmlRelation(ByRef ndxBase As XmlNode, ByRef ndxThis As XmlNode) As String
    Dim ndxWork As XmlNode  ' Working node
    Dim bIsLeaf As Boolean = False
    Dim intId As Integer    ' ID of node

    Try
      ' Validate
      If (ndxBase Is Nothing) OrElse (ndxThis Is Nothing) Then Return ""
      ' What type do we have?
      Select Case ndxThis.Name
        Case "eTree"
          ' Okay, continue
          ndxWork = ndxThis
        Case "eLeaf"
          ' Get the <eTree> node above it
          ndxWork = ndxThis.ParentNode
          ' Quick check
          If (ndxBase Is ndxWork) Then
            Return "child::eLeaf"
          End If
          bIsLeaf = True
        Case Else
          ' Cannot process this
          Return ""
      End Select
      ' Get the id
      intId = ndxWork.Attributes("Id").Value
      ' Is [ndxThis] under [ndxBase]?
      If (ndxBase.SelectSingleNode("./descendant::eTree[@Id=" & intId & "]") IsNot Nothing) Then
        ' Is it a child node?
        If (ndxBase.SelectSingleNode("./child::eTree[@Id=" & intId & "]") IsNot Nothing) Then
          ' It is a child
          Return IIf(bIsLeaf, "child::eTree/child::eLeaf", "child")
        Else
          Return IIf(bIsLeaf, "descendant::eTree/child::eLeaf", "descendant")
        End If
      ElseIf (ndxBase.SelectSingleNode("./ancestor::eTree[@Id=" & intId & "]") IsNot Nothing) Then
        ' It is an ancestor
        Return "ancestor"
      End If
      ' Is it a sibling?
      If (ndxBase.SelectSingleNode("./following-sibling::eTree[@Id=" & intId & "]") IsNot Nothing) Then
        ' Yes
        Return "following-sibling"
      ElseIf (ndxBase.SelectSingleNode("./preceding-sibling::eTree[@Id=" & intId & "]") IsNot Nothing) Then
        ' It is a preceding sibling
        Return "preceding-sibling"
      ElseIf (ndxBase.SelectSingleNode("./following::eTree[@Id=" & intId & "]") IsNot Nothing) Then
        ' It is a following node
        Return "following"
      ElseIf (ndxBase.SelectSingleNode("./preceding::eTree[@Id=" & intId & "]") IsNot Nothing) Then
        ' It is a preceding node
        Return "preceding"
      Else
        ' Do not know what it is
        Return ""
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("modXmlNode/XmlRelation error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   getXmlPath
  ' Goal:   Get the path from myself to the root
  ' History:
  ' 26-06-2015  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function getXmlPath(ByRef ndThis As XmlNode) As String
    Try
      ' Validate
      If (ndThis Is Nothing OrElse ndThis.NodeType <> XmlNodeType.Element) Then Return ""
      ' This is an element, so return my current name
      Return getXmlPath(ndThis.ParentNode) & "/" & ndThis.Name
    Catch ex As Exception
      ' Warn the user
      HandleErr("modXmlNode/getXmlPath error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   getCommonAncestor
  ' Goal:   Get the nearest common ancestor of two nodes
  ' History:
  ' 26-06-2015  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function getCommonAncestor(ByRef ndxBef As XmlNode, ByRef ndxAft As XmlNode, ByVal strCond As String, _
                                    ByRef ndxLeft As XmlNode, ByRef ndxRight As XmlNode) As XmlNode
    Dim ndxMyBef As XmlNode ' My copy of ndxBef
    Dim ndxWork As XmlNode  ' working node
    Dim strType As String   ' Kind of nodes

    Try
      ' Validate
      If (ndxBef Is Nothing) OrElse (ndxAft Is Nothing) Then Return Nothing
      If (ndxBef Is ndxAft) Then Return ndxBef
      ' Determine node type
      strType = ndxBef.Name
      ' Initialize left and right
      ndxLeft = ndxBef : ndxRight = ndxAft
      ndxMyBef = ndxBef
      ' Outer loop: before
      While (ndxMyBef IsNot Nothing) AndAlso (ndxMyBef.Name = strType)
        ' Find out if there is an ancestor of ndxAft equal to ndxBef
        ndxWork = ndxAft
        While (ndxWork IsNot Nothing) AndAlso (ndxWork.Name = strType)
          ' Test
          If (ndxMyBef Is ndxWork) Then
            ' Found it
            Return ndxMyBef
          End If
          ' adjust right
          ndxRight = ndxWork
          ' Go higher
          ndxWork = ndxWork.ParentNode
        End While
        ' Adjust left
        ndxLeft = ndxMyBef
        ' Try parent
        ndxMyBef = ndxMyBef.ParentNode
      End While

      ' Not found
      Return Nothing
    Catch ex As Exception
      ' Warn the user
      HandleErr("modXmlNode/getCommonAncestor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
End Module
