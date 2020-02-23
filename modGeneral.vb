Imports System.Net
Imports Microsoft.Office.Interop.Excel

Module modGeneral
    ' Array of full language name (separated by ; Fullname ; LanguageCode ; Charset)
    Public languageListNames As New List(Of String) _
    From {
        "Vietnamese;vi;ISO-8859-1",
        "English;en;ISO-8859-1",
        "Chinese (Simplified);zh-CN;GB2312",
        "Chinese (Traditional);zh-TW;Big5",
        "Korean;ko;EUC-KR",
        "Japanese;ja;Shift-JIS"
    }

    ' Store selection range and value
    Public Class cellProperties
        ' Properties
        Public Property sheet As Object                    ' Sheet to interact with
        Public Property addresses As New List(Of String)   ' List of selected cells
        Public Property values As New List(Of String)      ' List of selected cell values
    End Class

    ' For worker agruments
    Public Class cellArguments
        ' Properties
        Public Property sheet As Object
        Public Property address As String
        Public Property value As String
        ' Constructor
        Public Sub New(Optional sSheet As Object = Nothing,
                       Optional sAddr As String = "",
                       Optional sValue As String = "")
            sheet = sSheet
            address = sAddr
            value = sValue
        End Sub
    End Class

    ' Store Language properties
    Public Class languageProperties
        ' Properties
        Public Property name As String
        Public Property code As String
        Public Property charset As String
        ' Constructor
        Public Sub New(ByVal sName As String, ByVal sCode As String, ByVal sCharset As String)
            name = sName
            code = sCode
            charset = sCharset
        End Sub
    End Class

    Public languageList As List(Of languageProperties)
    Public translatedLst As List(Of cellArguments) ' Store translated cells

    ' Check for internet connection
    Public Function hasInternet() As Boolean
        Try
            Using client = New WebClient()
                Using stream = client.OpenRead("https://www.google.com/")
                    Return True
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function

End Module
