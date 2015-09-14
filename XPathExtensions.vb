Imports System.Xml
Imports System.Xml.XPath
Imports System.Xml.Xsl
Public Class XPathFunctions
  ' ------------------------------------------------------------------------------------
  ' Name:   DepType
  ' Goal:   Return the type of node [ndxThis] for dependency conversion processing
  ' Note:   The types that are returned:
  '           cnode       - non-end node, but only descendant is *con*
  '           code        - code node
  '           dnode       - node in its dislocated position
  '           droot       - location where the dislocated node should actually appear
  '           empty       - zero endnode or [ndxThis] itself is nothing or *con* node
  '           endnode     - 'normal' endnode with a vernacular eleaf child
  '           node        - normal non-end node
  '           punct       - punctuation endnode
  '           star        - unspecified endnode with * value
  '           tnode       - original node to which a trace points
  '           troot       - trace (empty) node, pointing to traceNd
  '           (nothing)   - error
  ' History:
  ' 30-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DepType(ByRef ndxThis As XmlNode) As String
    Dim ndxFor As XmlNode   ' Forest
    Dim ndxWork As XmlNode  ' Working node
    Dim ndxLeaf As XmlNode  ' Leaf
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim strValue As String  ' Value
    Dim intExt As Integer   ' Label extension number
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return "empty"
      If (ndxThis.Name <> "eTree") Then Return "empty"
      ' Give room for interrupt
      Application.DoEvents()
      ' =========== DEBUG ============
      ' If (ndxThis.Attributes("Id").Value = 169) Then Stop
      ' ==============================
      ' First check: do I have any 'Vern|Star|Punct' descendants?
      If (ndxThis.SelectSingleNode("./descendant::eLeaf[tb:matches(@Type,'Vern|Star|Punct')]", conTb) Is Nothing) Then Return "empty"
      ' Check if I am a code node
      If (ndxThis.Attributes("Label").Value = "CODE") Then Return "code"
      ' Check for cnode
      ndxList = ndxThis.SelectNodes("./descendant::eLeaf")
      If (ndxList.Count = 1) Then
        ' Check for [cnode]
        If (ndxList(0).Attributes("Text").Value = "*con*") Then Return "cnode"
      End If
      ' Get possible label number
      intExt = LabelExtNum(ndxThis)
      ' See if this is an endnode
      ndxLeaf = ndxThis.SelectSingleNode("./child::eLeaf")
      If (ndxLeaf IsNot Nothing) Then
        ' Type depends on the type of <eLeaf>
        Select Case ndxLeaf.Attributes("Type").Value
          Case "Vern"
            ' Action depends on existence of label number
            If (intExt < 0) Then Return "endnode"
            ' So if there IS a label number, then we continue
            ' Example: (WPRO-5 wes)
          Case "Zero"
            Return "empty"
          Case "Star"
            ' There are several possibilities...
            strValue = ndxLeaf.Attributes("Text").Value
            If (strValue = "*con*") OrElse (strValue = "*pro*") Then
              Return "empty"
            ElseIf (strValue = "*exp*") Then
              ' Extraposed clause
              Return "droot"
            ElseIf (InStr(strValue, "*ICH*") = 1) Then
              ' Dislocated: node with label -n may exist
              Return "droot"
            ElseIf (InStr(strValue, "*T*") = 1) Then
              ' Trace: node with label -n MAY exist
              Return "troot"
            Else
              ' This should not happen, but: who knows?
              Return "star"
            End If
          Case "Punct"
            Return "punct"
          Case Else
            ' This should not happen, but return the name of the type itself
            Return ndxLeaf.Attributes("Type").Value
        End Select
      End If
      ' Check if we have a number
      If (intExt >= 0) Then
        ' There is a label extension number (e.g: NP_SBJ-1)
        ' We need to determine what type of coreference goes on:
        ' a - Trace source          NP_SBJ-1 >> (NP_SBJ *T*-1)
        ' b - Trace source          WPRO-2   >> (NP_OB1-2 *T*)
        ' c - Dislocation pointer   PP-2     >> (NP *ICH*-2)
        ' d - Dislocation pointer   CP_REL-1 >> (CP_REL-1 *ICH*)
        ' e - Extraposition         CP_THT-5 >> (NP_SBJ-5 *exp*)
        ' f - Coreference           NP_LFD-1 >> (PRO-1 sy)... (PRO-1 sy)
        ' We need to find if there is a node this points to, and if so, which one
        ' (1) get the forest
        ndxFor = ndxThis.SelectSingleNode("./ancestor::forest")
        ' (2) Check for type [a]
        ndxWork = ndxFor.SelectSingleNode("./descendant::eTree[child::eLeaf[@Text='*T*-" & intExt & "']]")
        If (ndxWork IsNot Nothing) Then Return "tnode"
        ' (3) Check for type [c]
        ndxWork = ndxFor.SelectSingleNode("./descendant::eTree[child::eLeaf[@Text='*ICH*-" & intExt & "']]")
        If (ndxWork IsNot Nothing) Then Return "dnode"
        ' (4) Check for type [b] 
        ndxList = ndxFor.SelectNodes("./descendant::eTree[child::eLeaf[@Text='*T*']]")
        ' Walk the results
        For intI = 0 To ndxList.Count - 1
          ' Is this our node?
          If (ndxList(intI).Attributes("Label").Value Like "*-" & intExt) Then Return "tnode"
        Next intI
        ' (5) Check for type [d] 
        ndxList = ndxFor.SelectNodes("./descendant::eTree[child::eLeaf[@Text='*ICH*']]")
        ' Walk the results
        For intI = 0 To ndxList.Count - 1
          ' Is this our node?
          If (ndxList(intI).Attributes("Label").Value Like "*-" & intExt) Then Return "dnode"
        Next intI
        ' (6) Check for type [e] 
        ndxList = ndxFor.SelectNodes("./descendant::eTree[child::eLeaf[@Text='*exp*']]")
        ' Walk the results
        For intI = 0 To ndxList.Count - 1
          ' Is this our node?
          If (ndxList(intI).Attributes("Label").Value Like "*-" & intExt) Then Return "dnode"
        Next intI
        ' (7) we do not have to check for type [f], because such nodes are regarded as regular, and will come out with the default type.
        ' (8) It may be a normal endnode
        If (ndxLeaf IsNot Nothing) AndAlso (ndxLeaf.Attributes("Type").Value = "Vern") Then Return "endnode"
      End If
      ' Default: this is a normal node
      Return "node"
    Catch ex As Exception
      ' Show error
      HandleErr("XpathExt/DepType error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LabelExtNum
  ' Goal:   Get the extension-number after the @Label feature of [ndxThis]
  ' History:
  ' 30-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function LabelExtNum(ByRef ndxThis As XmlNode) As Integer
    Dim strLabel As String    ' Label
    Dim strLast As String     ' String follownig the last hyphen
    Dim intPos As Integer     ' Position in string

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return -1
      If (ndxThis.Name <> "eTree") Then Return -1
      ' Get label
      If (ndxThis.Attributes("Label") Is Nothing) Then Return -1
      strLabel = ndxThis.Attributes("Label").Value
      ' Look for the last hyphen
      intPos = InStrRev(strLabel, "-")
      If (intPos > 0) Then
        ' Get the string following the last hyphen
        strLast = Mid(strLabel, intPos + 1)
        ' Is this numeric?
        If (IsNumeric(strLast)) Then
          ' Return the number
          Return CInt(strLast)
        End If
      End If
      ' Return failure
      Return -1
    Catch ex As Exception
      ' Show error
      HandleErr("XpathExt/LabelExtNum error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
End Class

' ===============================================================================================================
' Name :  XPathExtensions
' Goal :  The interface that resolves and executes a specified user-defined function. 
' History:
' 17-06-2010  ERK Taken from http://msdn.microsoft.com/en-us/library/dd567715.aspx
' ===============================================================================================================
Public Class XPathExtensions
  Implements IXsltContextFunction

  ' The data types of the arguments passed to XPath extension function.
  Private m_ArgTypes() As XPathResultType
  ' The minimum number of arguments that can be passed to function.
  Private m_MinArgs As Integer
  ' The maximum number of arguments that can be passed to function.
  Private m_MaxArgs As Integer
  ' The data type returned by extension function.
  Private m_ReturnType As XPathResultType
  ' The name of the extension function.
  Private m_FunctionName As String
  ' Make sure we have a link to our own Xpath functions
  Private m_XpFun As New XPathFunctions

  ' Constructor used in the ResolveFunction method of the custom XsltContext 
  ' class to return an instance of IXsltContextFunction at run time.
  Public Sub New(ByVal MinArgs As Integer, ByVal MaxArgs As Integer, ByVal ReturnType As XPathResultType, _
                 ByVal ArgTypes() As XPathResultType, ByVal FunctionName As String)
    m_MinArgs = MinArgs
    m_MaxArgs = MaxArgs
    m_ReturnType = ReturnType
    m_ArgTypes = ArgTypes
    m_FunctionName = FunctionName
  End Sub

  ' Readonly property methods to access private fields.
  Public ReadOnly Property ArgTypes() As XPathResultType() Implements IXsltContextFunction.ArgTypes
    Get
      Return m_ArgTypes
    End Get
  End Property

  Public ReadOnly Property MaxArgs() As Integer Implements IXsltContextFunction.Maxargs
    Get
      Return m_MaxArgs
    End Get
  End Property

  Public ReadOnly Property MinArgs() As Integer Implements IXsltContextFunction.Minargs
    Get
      Return m_MinArgs
    End Get
  End Property

  Public ReadOnly Property ReturnType() As XPathResultType Implements IXsltContextFunction.ReturnType
    Get
      Return m_ReturnType
    End Get
  End Property

  ' Function to execute a specified user-defined XPath 
  ' extension function at run time.
  Public Function Invoke(ByVal Context As XsltContext, ByVal Args() As Object, ByVal DocContext As XPathNavigator) As Object Implements IXsltContextFunction.Invoke
    Dim strOne As String
    Dim strTwo As String
    Dim Node As XPathNodeIterator
    Dim objThis As IHasXmlNode
    Dim ndxThis As XmlNode        ' Working node

    Select Case m_FunctionName
      Case "CountChar"
        Return CountChar(DirectCast(Args(0), XPathNodeIterator), CChar(Args(1)))
      Case "FindTaskBy"
        Return FindTaskBy(DirectCast(Args(0), XPathNodeIterator), CStr(Args(1).ToString()))
      Case "Left"
        Return Left(CStr(Args(0)), CInt(Args(1)))
      Case "Right"
        Return Right(CStr(Args(0)), CInt(Args(1)))
      Case "Like", "matches"
        If (Args(0).GetType Is System.Type.GetType("System.String")) Then
          strOne = CStr(Args(0).ToString)
        Else
          Node = DirectCast(Args(0), XPathNodeIterator)
          ' Move to the first selected node
          ' See "XpathNodeIterator Class":
          '     An XPathNodeIterator object returned by the XPathNavigator class is not positioned on the first node 
          '     in a selected set of nodes. A call to the MoveNext method of the XPathNodeIterator class must be made 
          '     to position the XPathNodeIterator object on the first node in the selected set of nodes. 
          Node.MoveNext()
          ' Get the value of this attribute or node
          strOne = Node.Current.Value
          '' Check the kind of node we have
          'Select Case Node.Current.Name
          '  Case "eTree"
          '    ' Get the @Label
          '    strOne = Node.Current.GetAttribute("Label", "")
          '  Case "eLeaf"
          '    Node.MoveNext()
          '    ' Get the @Text
          '    strOne = Node.Current.GetAttribute("Text", "")
          '    strOne = Node.CurrentPosition
          '  Case "Feature"
          '    ' Get the @Value
          '    strOne = Node.Current.GetAttribute("Value", "")
          '  Case "Result"
          '    ' Since nothing is specified, I really do not know which of the features to take...
          '    strOne = ""
          '  Case Else
          '    strOne = ""
          'End Select
        End If
        ' If (InStr(strOne, "some") > 0) Then Stop

        strTwo = CStr(Args(1))
        Return DoLike(strOne, strTwo)
      Case "DepType", "deptype"
        ' ndxThis = DirectCast(Args(0), XmlNode)
        Node = DirectCast(Args(0), XPathNodeIterator)
        Node.MoveNext()
        objThis = DirectCast(Node.Current, IHasXmlNode)
        ndxThis = objThis.GetNode
        'ndxThis = DirectCast(Node.Current, XmlNode)
        Return DepType(ndxThis)
      Case "LabelExtNum", "labelnum"
        ' ndxThis = DirectCast(Args(0), XmlNode)
        Node = DirectCast(Args(0), XPathNodeIterator)
        Node.MoveNext()
        objThis = DirectCast(Node.Current, IHasXmlNode)
        ndxThis = objThis.GetNode
        'ndxThis = DirectCast(Node.Current, XmlNode)
        Return LabelExtNum(ndxThis)
      Case Else

    End Select
    ' Return Nothing for unknown function name.
    Return Nothing

  End Function

  ' XPath extension functions.
  Private Function CountChar(ByVal Node As XPathNodeIterator, ByVal CharToCount As Char) As Integer

    Dim CharCount As Integer = 0

    For CharIndex As Integer = 0 To Node.Current.Value.Length - 1
      If Node.Current.Value(CharIndex) = CharToCount Then
        CharCount += 1
      End If
    Next

    Return CharCount

  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DoLike
  ' Goal:   Perform the "Like" function using the pattern (or patterns) stored in [strPattern]
  '         There can be more than 1 pattern in [strPattern], which must be separated
  '         by a vertical bar: |
  ' History:
  ' 17-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function DoLike(ByVal strText As String, ByVal strPattern As String) As Boolean
    Dim arPattern() As String   ' Array of patterns
    Dim intI As Integer         ' Counter

    Try
      ' Reduce the [strPattern]
      strPattern = Trim(strPattern)
      ' ============== DEBUG ==============
      ' If (strPattern = "a") Then Stop
      ' ===================================
      ' SPlit the [strPattern] into different ones
      arPattern = Split(strPattern, "|")
      ' Perform the "Like" operation for all needed patterns
      For intI = 0 To UBound(arPattern)
        ' See if something positive comes out of this comparison
        If (strText Like arPattern(intI)) Then
          ' Return success
          Return True
        End If
      Next intI
      ' No match has happened, so return false
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("XpathExt/DoLike error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DepType
  ' Goal:   Return the type of node [ndxThis] for dependency conversion processing
  ' History:
  ' 30-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function DepType(ByRef ndxThis As XmlNode) As String
    Return m_XpFun.DepType(ndxThis)
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LabelExtNum
  ' Goal:   Get the extension-number after the @Label feature of [ndxThis]
  ' History:
  ' 30-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function LabelExtNum(ByRef ndxThis As XmlNode) As Integer
    Return m_XpFun.LabelExtNum(ndxThis)
  End Function

  ' This overload will not force the user 
  ' to cast to string in the xpath expression
  Private Function FindTaskBy(ByVal Node As XPathNodeIterator, ByVal Text As String) As String

    If (Node.Current.Value.Contains(Text)) Then
      Return Node.Current.Value
    Else
      Return ""
    End If

  End Function

End Class
' ===============================================================================================================
' Name :  CustomContext
' Goal :  Provide a custom context
' History:
' 17-06-2010  ERK Taken from http://msdn.microsoft.com/en-us/library/dd567715.aspx
' ===============================================================================================================

Class CustomContext
  Inherits XsltContext

  Private Const ExtensionsNamespaceUri As String = TREEBANK_EXTENSIONS

  ' XsltArgumentList to store names and values of user-defined variables.
  Private m_ArgList As XsltArgumentList

  Public Sub New()

  End Sub

  Public Sub New(ByVal NT As NameTable, ByVal Args As XsltArgumentList)
    MyBase.New(NT)
    m_ArgList = Args
  End Sub

  ' Empty implementation, returns 0.
  Public Overrides Function CompareDocument(ByVal BaseUri As String, ByVal NextBaseUri As String) As Integer
    Return 0
  End Function

  ' Empty implementation, returns false.
  Public Overrides Function PreserveWhitespace(ByVal Node As XPathNavigator) As Boolean
    Return False
  End Function

  Public Overrides Function ResolveFunction(ByVal Prefix As String, ByVal Name As String, _
                                            ByVal ArgTypes() As XPathResultType) As IXsltContextFunction

    If LookupNamespace(Prefix) = ExtensionsNamespaceUri Then
      Select Case Name
        Case "CountChar"
          Return New XPathExtensions(2, 2, _
                          XPathResultType.Number, ArgTypes, "CountChar")

        Case "FindTaskBy" ' Implemented but not called.
          Return New XPathExtensions(2, 2, _
                          XPathResultType.String, ArgTypes, "FindTaskBy")

        Case "Right"  ' Implemented but not called.
          Return New XPathExtensions(2, 2, _
                          XPathResultType.String, ArgTypes, "Right")

        Case "Left"   ' Implemented but not called.
          Return New XPathExtensions(2, 2, _
                          XPathResultType.String, ArgTypes, "Left")

        Case "Like", "matches"   ' Implemented but not called.
          Return New XPathExtensions(2, 2, _
                          XPathResultType.Boolean, ArgTypes, "Like")
        Case "DepType", "deptype"
          Return New XPathExtensions(1, 1, XPathResultType.String, ArgTypes, "DepType")
        Case "LabelExtNum", "labelnum"
          Return New XPathExtensions(1, 1, XPathResultType.Number, ArgTypes, "LabelExtNum")
        Case Else

      End Select
    End If
    ' Return Nothing if none of the functions match name.
    Return Nothing

  End Function

  ' Function to resolve references to user-defined XPath 
  ' extension variables in XPath query.
  Public Overrides Function ResolveVariable(ByVal Prefix As String, ByVal Name As String) As IXsltContextVariable
    If LookupNamespace(Prefix) = ExtensionsNamespaceUri OrElse Len(Prefix) > 0 Then
      Throw New XPathException(String.Format("Variable '{0}:{1}' is not defined.", Prefix, Name))
    End If

    Select Case Name
      Case "charToCount", "left", "right", "text"
        ' Create an instance of an XPathExtensionVariable 
        ' (custom IXsltContextVariable implementation) object 
        ' by supplying the name of the user-defined variable to resolve.
        Return New XPathExtensionVariable(Prefix, Name)

        ' The Evaluate method of the returned object will be used at run time
        ' to resolve the user-defined variable that is referenced in the XPath
        ' query expression. 
      Case Else

    End Select
    ' Return Nothing if none of the variables match name.
    Return Nothing

  End Function

  Public Overrides ReadOnly Property Whitespace() As Boolean
    Get
      Return True
    End Get
  End Property

  ' The XsltArgumentList property is accessed by the Evaluate method of the 
  ' XPathExtensionVariable object that the ResolveVariable method returns. 
  ' It is used to resolve references to user-defined variables in XPath query 
  ' expressions. 
  Public ReadOnly Property ArgList() As XsltArgumentList
    Get
      Return m_ArgList
    End Get
  End Property
End Class

' ===============================================================================================================
' Name :  XPathExtensions
' Goal :  The interface used to resolve references to user-defined variables
'           in XPath query expressions at run time. An instance of this class 
'           is returned by the overridden ResolveVariable function of the 
'           custom XsltContext class. 
' History:
' 17-06-2010  ERK Taken from http://msdn.microsoft.com/en-us/library/dd567715.aspx
' ===============================================================================================================
Public Class XPathExtensionVariable
  Implements IXsltContextVariable

  ' Namespace of user-defined variable.
  Private m_Prefix As String
  ' The name of the user-defined variable.
  Private m_VarName As String

  ' Constructor used in the overridden ResolveVariable function of custom XsltContext.
  Public Sub New(ByVal Prefix As String, ByVal VarName As String)
    m_Prefix = Prefix
    m_VarName = VarName
  End Sub

  ' Function to return the value of the specified user-defined variable.
  ' The GetParam method of the XsltArgumentList property of the active
  ' XsltContext object returns value assigned to the specified variable.
  Public Function Evaluate(ByVal Context As XsltContext) As Object Implements IXsltContextVariable.Evaluate
    Dim vars As XsltArgumentList = DirectCast(Context, CustomContext).ArgList
    Return vars.GetParam(m_VarName, m_Prefix)
  End Function

  ' Determines whether this variable is a local XSLT variable.
  ' Needed only when using a style sheet.
  Public ReadOnly Property IsLocal() As Boolean Implements IXsltContextVariable.IsLocal
    Get
      Return False
    End Get
  End Property

  ' Determines whether this parameter is an XSLT parameter.
  ' Needed only when using a style sheet.
  Public ReadOnly Property IsParam() As Boolean Implements IXsltContextVariable.IsParam
    Get
      Return False
    End Get
  End Property

  Public ReadOnly Property VariableType() As XPathResultType Implements IXsltContextVariable.VariableType
    Get
      Return XPathResultType.Any
    End Get
  End Property
End Class
