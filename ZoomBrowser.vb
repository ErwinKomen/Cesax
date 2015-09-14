Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Namespace ZoomBrowser
  Partial Public Class MyBrowser
    Inherits WebBrowser
#Region "enums"
    Public Enum OLECMDID
      ' ...
      OLECMDID_OPTICAL_ZOOM = 63
      OLECMDID_OPTICAL_GETZOOMRANGE = 64
      ' ...
    End Enum

    Public Enum OLECMDEXECOPT
      ' ...
      OLECMDEXECOPT_DONTPROMPTUSER
      ' ...
    End Enum

    Public Enum OLECMDF
      ' ...
      OLECMDF_SUPPORTED = 1
    End Enum
#End Region

#Region "IWebBrowser2"
    'SuppressUnmanagedCodeSecurity,
    <ComImport(), TypeLibType(TypeLibTypeFlags.FOleAutomation Or TypeLibTypeFlags.FDual Or TypeLibTypeFlags.FHidden), Guid("D30C1661-CDAF-11d0-8A3E-00C04FC9E26E")> _
    Public Interface IWebBrowser2
      <DispId(100)> _
      Sub GoBack()
      <DispId(&H65)> _
      Sub GoForward()
      <DispId(&H66)> _
      Sub GoHome()
      <DispId(&H67)> _
      Sub GoSearch()
      <DispId(&H68)> _
      Sub Navigate(<[In]()> ByVal Url As String, <[In]()> ByRef flags As Object, <[In]()> ByRef targetFrameName As Object, <[In]()> ByRef postData As Object, <[In]()> ByRef headers As Object)
      <DispId(-550)> _
      Sub Refresh()
      <DispId(&H69)> _
      Sub Refresh2(<[In]()> ByRef level As Object)
      <DispId(&H6A)> _
      Sub [Stop]()
      <DispId(200)> _
      ReadOnly Property Application() As Object
      <DispId(&HC9)> _
      ReadOnly Property Parent() As Object
      <DispId(&HCA)> _
      ReadOnly Property Container() As Object
      <DispId(&HCB)> _
      ReadOnly Property Document() As Object
      <DispId(&HCC)> _
      ReadOnly Property TopLevelContainer() As Boolean
      <DispId(&HCD)> _
      ReadOnly Property Type() As String
      <DispId(&HCE)> _
      Property Left() As Integer
      <DispId(&HCF)> _
      Property Top() As Integer
      <DispId(&HD0)> _
      Property Width() As Integer
      <DispId(&HD1)> _
      Property Height() As Integer
      <DispId(210)> _
      ReadOnly Property LocationName() As String
      <DispId(&HD3)> _
      ReadOnly Property LocationURL() As String
      <DispId(&HD4)> _
      ReadOnly Property Busy() As Boolean
      <DispId(300)> _
      Sub Quit()
      <DispId(&H12D)> _
      Sub ClientToWindow(ByRef pcx As Integer, ByRef pcy As Integer)
      <DispId(&H12E)> _
      Sub PutProperty(<[In]()> ByVal [property] As String, <[In]()> ByVal vtValue As Object)
      <DispId(&H12F)> _
      Function GetProperty(<[In]()> ByVal [property] As String) As Object
      <DispId(0)> _
      ReadOnly Property Name() As String
      <DispId(-515)> _
      ReadOnly Property HWND() As Integer
      <DispId(400)> _
      ReadOnly Property FullName() As String
      <DispId(&H191)> _
      ReadOnly Property Path() As String
      <DispId(&H192)> _
      Property Visible() As Boolean
      <DispId(&H193)> _
      Property StatusBar() As Boolean
      <DispId(&H194)> _
      Property StatusText() As String
      <DispId(&H195)> _
      Property ToolBar() As Integer
      <DispId(&H196)> _
      Property MenuBar() As Boolean
      <DispId(&H197)> _
      Property FullScreen() As Boolean
      <DispId(500)> _
      Sub Navigate2(<[In]()> ByRef URL As Object, <[In]()> ByRef flags As Object, <[In]()> ByRef targetFrameName As Object, <[In]()> ByRef postData As Object, <[In]()> ByRef headers As Object)
      <DispId(&H1F5)> _
      Function QueryStatusWB(<[In]()> ByVal cmdID As OLECMDID) As OLECMDF
      <DispId(&H1F6)> _
      Sub ExecWB(<[In]()> ByVal cmdID As OLECMDID, <[In]()> ByVal cmdexecopt As OLECMDEXECOPT, ByRef pvaIn As Object, ByVal pvaOut As IntPtr)
      <DispId(&H1F7)> _
      Sub ShowBrowserBar(<[In]()> ByRef pvaClsid As Object, <[In]()> ByRef pvarShow As Object, <[In]()> ByRef pvarSize As Object)
      <DispId(-525)> _
      ReadOnly Property ReadyState() As WebBrowserReadyState
      <DispId(550)> _
      Property Offline() As Boolean
      <DispId(&H227)> _
      Property Silent() As Boolean
      <DispId(&H228)> _
      Property RegisterAsBrowser() As Boolean
      <DispId(&H229)> _
      Property RegisterAsDropTarget() As Boolean
      <DispId(&H22A)> _
      Property TheaterMode() As Boolean
      <DispId(&H22B)> _
      Property AddressBar() As Boolean
      <DispId(&H22C)> _
      Property Resizable() As Boolean
    End Interface
#End Region

    Private axIWebBrowser2 As IWebBrowser2

    Public Sub New()
    End Sub

    Protected Overrides Sub AttachInterfaces(ByVal nativeActiveXObject As Object)
      MyBase.AttachInterfaces(nativeActiveXObject)
      Me.axIWebBrowser2 = DirectCast(nativeActiveXObject, IWebBrowser2)
    End Sub

    Protected Overrides Sub DetachInterfaces()
      MyBase.DetachInterfaces()
      Me.axIWebBrowser2 = Nothing
    End Sub

    Public Sub Zoom(ByVal factor As Integer)
      Dim pvaIn As Object = factor
      Try
        Me.axIWebBrowser2.ExecWB(OLECMDID.OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER, pvaIn, IntPtr.Zero)
      Catch generatedExceptionName As Exception
        Throw
      End Try
    End Sub
  End Class
End Namespace
