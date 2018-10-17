Imports System.Web
Imports System.Web.Mail
Imports System.Collections.Specialized
Public Class ErrorHandler
  Inherits System.Web.UI.Page


  Public sConnString As String = ""
  Public sErrorEmailAddress As String = ""
  Public sErrorRedirectPage As String = ""
  Public oUtil As SharedUtilities
  Public bRunningAsDLL As Boolean = False

  Private Function CNullS(ByVal oInp As Object, Optional ByVal sDefault As String = "") As String
    CNullS = sDefault
    If Not IsDBNull(oInp) Then
      If Not (oInp Is Nothing) Then CNullS = CStr(oInp)
    End If
  End Function
  Public Sub HandleError(ByVal Ex As Exception, Optional ByVal bRedirectToErrorPage As Boolean = True, Optional ByVal bSendEmail As Boolean = True)

    On Error GoTo EH
    Dim sHTML As String = "", iErrID As Int32 = 0

    ' Get Error information
    sHTML = GetError(Ex)
    If Not bRunningAsDLL Then
      sHTML += GetHTMLError()

      iErrID = LogWebError(Ex.Message, Ex.Source, True, CNullS(Ex.TargetSite.ToString), sHTML, , Ex.StackTrace)

      ' Redirect to Error page if needed
      If bRedirectToErrorPage Then
        Err.Clear()
        HandleRedirect(bRedirectToErrorPage, iErrID)
      End If
    End If
EH:
    Err.Clear()
  End Sub

  Public Function LogWebError(ByVal sMessage As String, Optional ByVal sSource As String = "", Optional ByVal bSendEmail As Boolean = True, Optional ByVal sTargetSite As String = "", Optional ByVal sErrorHTML As String = "", Optional ByVal iNumber As Integer = 0, Optional ByVal sStackTrace As String = "", Optional ByVal sUserInput As String = "") As Integer

    Dim oError As New TableWebErrorLog(sConnString), iErrID As Int32 = 0
    Dim sEmailSubject As String = "", sLastPage As String = ""

    If sErrorHTML.ToString = "" Then
      Dim Heading As String
      Heading = "<TABLE BORDER=""0"" WIDTH=""100%"" CELLPADDING=""1"" CELLSPACING=""0""><TR><TD bgcolor=""black"" COLSPAN=""2""><FONT face=""Arial"" color=""white""><B> <!--HEADER--></B></FONT></TD></TR></TABLE>"
      sErrorHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Message: " & sMessage)
      sErrorHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Source: " & sSource)
      sErrorHTML += GetHTMLError()

    End If
    '        sHTML = sErrorHTML

    ' Save in DB
    sErrorHTML = sErrorHTML
    oError.cErrorHTML__String = sErrorHTML.ToString
    oError.cErrorMessage__String = Left(sMessage.ToString, 5900)
    oError.cErrorSource__String = sSource.ToString
    oError.cErrorStackTrace__String = sStackTrace.ToString
    oError.cErrorTargetSite__String = sTargetSite.ToString
    oError.dtErrorDate__Date = Now
    LogWebError = oError.Insert
    oError = Nothing

    If bSendEmail Then

      ' Email it
      sEmailSubject = "Web Error from CStay Resrv Website - " & sMessage.Substring(0, Math.Min(30, sMessage.Length))

      If CStr(sErrorEmailAddress) <> "" Then
        oUtil.SendMail(sErrorEmailAddress, sErrorEmailAddress, sEmailSubject, sMessage.ToString & "<BR/><BR/>" & sErrorHTML, sErrorEmailAddress, sErrorEmailAddress, True, , True)
      End If
    End If

  End Function

  Public Function GetError(ByVal Ex As Exception) As String
    'Returns HTML an formatted error message.
    Dim Heading As String
    Dim MyHTML As String
    Dim Error_Info As New NameValueCollection
    Heading = "<TABLE BORDER=""0"" WIDTH=""100%"" CELLPADDING=""1"" CELLSPACING=""0""><TR><TD bgcolor=""black"" COLSPAN=""2""><FONT face=""Arial"" color=""white""><B> <!--HEADER--></B></FONT></TD></TR></TABLE>"
    MyHTML = "<FONT face=""Arial"" size=""4"" color=""red"">Error - " & Ex.Message & "</FONT><BR/><BR/>"
    Error_Info.Add("Message", CleanHTML(Ex.Message))
    Error_Info.Add("Source", CleanHTML(Ex.Source))
    Error_Info.Add("TargetSite", CleanHTML(CNullS(Ex.TargetSite.ToString)))
    Error_Info.Add("StackTrace", CleanHTML(Replace(Ex.StackTrace, " at ", "<BR/>")))
    If Ex.InnerException IsNot Nothing Then
      Error_Info.Add("Message", CleanHTML(Ex.Message))
      Error_Info.Add("Source", CleanHTML(Ex.Source))
      Error_Info.Add("TargetSite", CleanHTML(CNullS(Ex.TargetSite.ToString)))
      Error_Info.Add("StackTrace", CleanHTML(Replace(Ex.StackTrace, " at ", "<BR/>")))
    End If
    MyHTML += Heading.Replace("<!--HEADER-->", "Error Information")
    MyHTML += CollectionToHtmlTable(Error_Info)
    Return MyHTML
  End Function



  Public Function CollectionToHtmlTable(ByRef Collection As NameValueCollection) As String
    Dim TDName As String, TDValue As String
    Dim MyHTML As String
    Dim i As Integer
    TDName = "<TD width=""170"" ><FONT face=""Arial"" size=""2""><!--NAME--></FONT></TD>"
    TDValue = "<TD ><FONT face=""Arial"" size=""2""><!--VALUE--></FONT></TD>"
    MyHTML = "<TABLE width=""100%"">" & _
    " <TR bgcolor=""#C0C0C0"">" & _
    TDName.Replace("<!--NAME-->", " <B>Name</B>") & _
    "Value" & TDValue.Replace("<!--VALUE-->", " <B>Value</B>") & "</TR>"
    'No Body? -> N/A
    If (Collection.Count <= 0) Then
      Collection = New NameValueCollection
      Collection.Add("N/A", "")
    Else
      'Table Body
      For i = 0 To Collection.Count - 1
        On Error Resume Next
        If Collection.Keys(i) <> "__VIEWSTATE" Then
          MyHTML += "<TR valign=""top"" bgcolor=""#EEEEEE"">" & _
          TDName.Replace("<!--NAME-->", Collection.Keys(i)) & " " & _
          TDValue.Replace("<!--VALUE-->", Collection(i)) & "</TR> "
        End If
      Next i
    End If
    'Table Footer
    Return MyHTML & "</TABLE>"
  End Function
  Private Function CollectionToHtmlTable(ByVal Collection As HttpCookieCollection) As String
    'Converts HttpCookieCollection to NameValueCollection
    Dim NVC As New NameValueCollection
    Dim i As Integer
    Dim Value As String
    CollectionToHtmlTable = ""
    Try
      If Collection.Count > 0 Then
        For i = 0 To Collection.Count - 1
          NVC.Add(Collection.Keys(i), Collection(i).Value)
        Next i
      End If
      Value = CollectionToHtmlTable(NVC)
      Return Value
    Catch MyError As Exception
      Dim s As String = MyError.Message
    End Try
  End Function
  Private Function CollectionToHtmlTable(ByVal Collection As System.Web.SessionState.HttpSessionState) As String
    'Converts HttpSessionState to NameValueCollection
    Dim NVC As New NameValueCollection
    Dim i As Integer
    Dim Value As String
    If Collection.Count > 0 Then
      On Error Resume Next
      For i = 0 To Collection.Count - 1
        NVC.Add(Collection.Keys(i), CNullS(Collection.Item(i)))
        Err.Clear()
      Next i
    End If
    Value = CollectionToHtmlTable(NVC)
    Return Value
  End Function
  'Private Function CollectionToHtmlTable(ByRef oViewState As System.Web.UI.StateBag) As String
  '  'Converts HttpSessionState to NameValueCollection
  '  Dim NVC As New NameValueCollection
  '  Dim i As Integer
  '  Dim Value As String

  '  '    Select Case sNamedCollection
  '  '      Case "ViewState"
  '  If Not oViewState Is Nothing Then
  '    For Each sKey As String In ViewState.Keys
  '      NVC.Add(sKey, ViewState.Item(sKey).ToString)
  '    Next
  '  End If
  '  '    End Select
  '  Value = CollectionToHtmlTable(NVC)
  '  Return Value
  'End Function

  Private Function CleanHTML(ByVal HTML As String) As String
    If HTML.Length <> 0 Then
      HTML.Replace("<", "<").Replace("\r\n", "<BR/>").Replace("&", "&").Replace(" ", " ")
    Else
      HTML = ""
    End If
    Return HTML
  End Function

  Public Sub New(ByRef oShared As SharedUtilities, ByVal bDontUseAppSettings As Boolean)
    sConnString = oShared.sConnString
    sErrorEmailAddress = oShared.sErrorEmailAddress
    sErrorRedirectPage = oShared.sErrorRedirectPage
    oUtil = oShared
  End Sub


  Public Sub New()

  End Sub
  Public Sub New(ByRef oShared As SharedUtilities)
    oUtil = oShared
    sErrorEmailAddress = System.Configuration.ConfigurationSettings.AppSettings("ErrorEmailAddress")
    sErrorRedirectPage = System.Configuration.ConfigurationSettings.AppSettings("ErrorRedirectPage")
  End Sub

  Public Function GetHTMLError() As String
    GetHTMLError = ""
    If Not bRunningAsDLL Then

      'Returns HTML an formatted error message.
      Dim Heading As String
      Dim MyHTML As String = ""
      Dim Error_Info As New NameValueCollection
      Heading = "<TABLE BORDER=""0"" WIDTH=""100%"" CELLPADDING=""1"" CELLSPACING=""0""><TR><TD bgcolor=""black"" COLSPAN=""2""><FONT face=""Arial"" color=""white""><B> <!--HEADER--></B></FONT></TD></TR></TABLE>"
      '// QueryString Collection
      MyHTML += "<BR/><BR/>" + Heading.Replace("<!--HEADER-->", "QueryString Collection")
      MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.QueryString)
      '// Form Collection
      MyHTML += "<BR/><BR/>" + Heading.Replace("<!--HEADER-->", "Form Collection")
      MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.Form)
      ''// View State
      'MyHTML += "<BR/><BR/>" + Heading.Replace("<!--HEADER-->", "Cookies Collection")
      'MyHTML += CollectionToHtmlTable(oViewState)
      '// Cookies Collection
      MyHTML += "<BR/><BR/>" + Heading.Replace("<!--HEADER-->", "Cookies Collection")
      MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.Cookies)
      '// Session Variables
      MyHTML += "<BR/><BR/>" + Heading.Replace("<!--HEADER-->", "Session Variables")
      MyHTML += CollectionToHtmlTable(HttpContext.Current.Session)
      '// Server Variables
      MyHTML += "<BR/><BR/>" + Heading.Replace("<!--HEADER-->", "Server Variables")
      MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.ServerVariables)
      GetHTMLError = MyHTML
    End If
  End Function
  Public Function GetIISInfo() As String

    Dim Heading As String
    Dim MyHTML As String
    Heading = "<TABLE BORDER=""0"" WIDTH=""100%"" CELLPADDING=""1"" CELLSPACING=""0""><TR><TD bgcolor=""black"" COLSPAN=""2""><FONT face=""Arial"" color=""white""><B> <!--HEADER--></B></FONT></TD></TR></TABLE>"
    MyHTML = "<FONT face=""Arial"" size=""4"" color=""red"">IIS Information</FONT><BR><BR>"
    '// QueryString Collection
    MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "QueryString Collection")
    MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.QueryString)
    '// Form Collection
    MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Form Collection")
    MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.Form)
    ''// View State
    'MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Cookies Collection")
    'MyHTML += CollectionToHtmlTable(oViewState)
    '// Cookies Collection
    MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Cookies Collection")
    MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.Cookies)
    '// Session Variables
    MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Session Variables")
    MyHTML += CollectionToHtmlTable(HttpContext.Current.Session)
    '// Server Variables
    MyHTML += "<BR><BR>" + Heading.Replace("<!--HEADER-->", "Server Variables")
    MyHTML += CollectionToHtmlTable(HttpContext.Current.Request.ServerVariables)
    GetIISInfo = MyHTML
  End Function


  Private Sub HandleRedirect(ByVal bRedirectToErrorPage As Boolean, ByVal iErrID As Integer)

    Dim sErrRedirect As String = "", sLastPage As String = ""

    sErrRedirect = sErrorRedirectPage

    If sErrRedirect.ToString <> "" Then
      sLastPage = HttpContext.Current.Request.ServerVariables("SCRIPT_NAME")
      HttpContext.Current.Response.Redirect(sErrRedirect & "?ErrID=" & iErrID & "&LastPage=" & HttpContext.Current.Server.UrlEncode(sLastPage.ToString), True)
    End If

  End Sub

  Protected Overrides Sub Finalize()
    oUtil = Nothing
    MyBase.Finalize()
  End Sub
End Class
