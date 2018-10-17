Imports System.Runtime.CompilerServices

Public Class MyUtilities

    Public Sub New()

    End Sub
    Public Sub New(ByVal oShared As SharedUtilities)

    End Sub

  Public Function GetLine(<CallerMemberName> Optional sCallerName As String = "", <CallerLineNumber> Optional iCallerLineNumber As Integer = 0) As String
    Return Now.ToString & " - " & sCallerName & " - " & iCallerLineNumber & vbCrLf
  End Function


End Class
