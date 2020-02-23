Imports System.ComponentModel
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Timers

Public Class cTranslator
    Public Event Finished() ' Callback

    Private webClient As WebClient
    Private WithEvents timerOut As Timer
    Private cell As cellArguments

    ' Constructor
    Public Sub New(ByVal sheet As Object, ByVal cellAddr As String, ByVal cellValue As String)
        cell = New cellArguments
        cell.sheet = sheet
        cell.address = cellAddr
        cell.value = cellValue
    End Sub

    Private Sub timeOut()
        On Error Resume Next
        ' When Timeout occur
        webClient.Dispose()
        RaiseEvent Finished() ' Notify completed
    End Sub

#Region "Translate"
    ' Translate and get result
    Public Sub startTranslate(ByVal fromLang As String, ByVal toLang As String, ByVal strCharset As String)
        On Error GoTo NULL
        ' Check cell condition
        If (cell.value.Chars(0) = "=" Or cell.value = "") Then
            RaiseEvent Finished()
            Exit Sub
        End If
        ' Base url
        Dim url As String = "https://translate.google.pl/m?hl={0}&sl={1}&tl={2}&ie=UTF-8&prev=_m&q={3}"
        ' Encode URL to escape special characters
        Dim str As String = EncodeUrl(cell.value)
        ' Combine with params
        url = String.Format(url, fromLang, fromLang, toLang, str)
        ' Send get request
        webClient = New WebClient
        ' Set approviate encoding for webclient
        webClient.Encoding = Encoding.GetEncoding(strCharset)
        ' Start webclient request
        AddHandler webClient.DownloadStringCompleted, AddressOf webClient_DownloadStringCompleted
        ' Start timer to make timeout event 5 seconds
        timerOut = New Timer(5000)
        timerOut.AutoReset = False
        timerOut.Enabled = False
        ' Start async tasks
        webClient.DownloadStringAsync(New Uri(url))
NULL:
    End Sub

    Private Sub webClient_DownloadStringCompleted(ByVal sender As Object, ByVal e As DownloadStringCompletedEventArgs)
        On Error GoTo NULL
        Dim result As String = WebUtility.HtmlDecode(ParsingHTMLResult(e.Result))
        ' Put translated string to cell
        cell.sheet.range(cell.address).value = result
NULL:
        webClient.Dispose()
        RaiseEvent Finished()
    End Sub
#End Region

#Region "Parsing HTML function"

    ' Get translated value and remove HTML tags
    Private Function ParsingHTMLResult(str As String) As Object
        On Error Resume Next
        Dim matches As MatchCollection
        ' Check if include translated string
        matches = Regex.Matches(str, "<div dir=""ltr"" class=""t0"">(.*?)</div>", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
        If matches.Count = 1 Then
            ' Remove html tag
            Dim regex As New Regex("^<.*?>|(<\/[^<>/]*>)(?!.*\1)", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
            Dim result As String = regex.Replace(matches.Item(0).ToString, String.Empty)
            Return result
        End If
        ' No matches
        Return ""
    End Function

    ' Escape special url character from string
    Private Function EncodeUrl(ByVal str As String) As String
        On Error Resume Next
        Dim pattern As String = "[$&+,/:;=?@]"
        ' Check for each occurence of special character and replace one by one
        Dim match As Match = Regex.Match(str, pattern, RegexOptions.Singleline Or RegexOptions.IgnoreCase)
        Do While (match.Success)
            str = str.Replace(match.Value, Uri.EscapeDataString(match.Value))
            match = Regex.Match(str, pattern)
        Loop
        Return str
    End Function

    Private Sub timerOut_Elapsed(sender As Object, e As ElapsedEventArgs) Handles timerOut.Elapsed
        timeOut()
    End Sub

#End Region

End Class
