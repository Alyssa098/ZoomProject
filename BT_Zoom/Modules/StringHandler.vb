Option Strict On
Option Explicit On

Imports System
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Collections.Generic
Imports System.Web.HttpContext

Public Class StringHandler

    Public Enum ConcatenationMethod
        Html = 1
        Textbox = 2
    End Enum

#Region "String Manipulation/Parsing and Computations"

    Public Shared Function GetActivationCode() As String
        Static rand As New Random
        Dim sb As New System.Text.StringBuilder(15)
        For i As Integer = 1 To 15
            Dim charIndex As Integer
            ' allow only digits and letters
            Do
                charIndex = rand.Next(48, 123)
            Loop Until (charIndex >= 48 AndAlso charIndex <= 57) OrElse (charIndex >= 65 AndAlso charIndex <= 90) OrElse (charIndex >= 97 AndAlso charIndex <= 122)
            ' add the random char to the code being built
            sb.Append(Convert.ToChar(charIndex))
        Next
        Return sb.ToString
    End Function

    Public Shared Function Compute_Size(ByVal bytes As Object) As String
        If IsDBNull(bytes) Then
            Return "0"
        Else
            Return (CDbl(CLng(bytes) / 1024).ToString("0.#"))
        End If
    End Function

    Public Shared Function AbbreviateString(ByVal instr As String, ByVal maxlength As Integer, Optional ByVal strEnd As String = "") As String
        If String.IsNullOrEmpty(instr) Then
            Return (instr)
        End If
        If instr.Length > maxlength Then
            Return (instr.Substring(0, Math.Min(instr.Length, maxlength - strEnd.Length)) & strEnd)
        Else
            Return (instr)  'not longer than maxlength just return
        End If
    End Function

    Public Shared Function InjectionProofListOfIntsCheck(ByVal instr As String) As Boolean
        Dim retValue As Boolean = True
        If instr.Length > 0 Then
            Dim arrIDs As String() = instr.Split(","c)
            Dim checkInt As Integer
            Try
                For i As Integer = 0 To arrIDs.Length - 1
                    checkInt = CInt(arrIDs(i))  'for injection security, we need to verify all the incoming values are integers before allowing them passed to dynamic sql
                Next
            Catch
                retValue = False
            End Try
        End If
        Return (retValue)
    End Function

    ''' <summary>
    ''' Try to parse a string as a list of integers
    ''' </summary>
    ''' <param name="instr">Comma delimated list of integers in string form</param>
    ''' <param name="intList">IntList to return if instr is valid</param>
    ''' <returns>False if instr is invalid, else intList is filled with valid integers</returns>
    ''' <remarks>near copy of InjectProofListOfIntsCheck, but exits function on failure and does not rely on exceptions to process</remarks>
    Public Shared Function TryParseListOfInts(ByVal instr As String, ByRef intList As List(Of Integer)) As Boolean
        Dim goodList As Boolean = True
        intList = New List(Of Integer) ''Recreate the list just in case someone passes it in with values
        If instr.Length > 0 Then
            Dim arrIDs As String() = instr.Split(","c)
            For Each s As String In arrIDs
                Dim checkInt As Integer
                If Integer.TryParse(s, checkInt) Then
                    intList.Add(checkInt)
                Else
                    intList.Clear()
                    Return False ''Exit function if the string contains invalid characters
                End If
            Next
        End If
        Return goodList
    End Function

    Public Shared Function InjectionProof(ByVal userInputSQL As String) As String
        'Remove ;, --, and ' from user input
        Return (userInputSQL.Replace(";", "").Replace("--", "").Replace("'", ""))
    End Function

    Public Shared Function InjectionProofEscapeSingleQuote(ByVal userInputSQL As String) As String
        'Remove ;, --, and ' from user input
        Return (userInputSQL.Replace(";", "").Replace("--", "").Replace("'", "''"))
    End Function

    Public Shared Function EscapeSingleQuote(ByVal userInputSQL As String) As String
        Return (userInputSQL.Replace("'", "''"))
    End Function

    ''' <summary>
    ''' Pass in a dictionary and key/values and recieve back an string in the format of key1=value1 AMPERSAND key2=value2
    ''' </summary>
    ''' <param name="dict"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetStringFromDictionary(ByVal dict As Dictionary(Of String, String)) As String
        Dim sb As New StringBuilder
        Dim ampersand As String = ""
        For Each item As KeyValuePair(Of String, String) In dict
            sb.Append(String.Format("{2}{0}={1}", item.Key, item.Value, ampersand))
            ampersand = "&"
        Next

        Return sb.ToString
    End Function

    Public Shared Function GetDictionaryFromString(ByVal sString As String, Optional ByVal doUrlDecode As Boolean = False) As Dictionary(Of String, String)
        Dim sDictItems() As String
        Dim dictItem() As String
        Dim dict As New Dictionary(Of String, String)

        sDictItems = sString.Split(CChar("&"))
        For Each sKVP As String In sDictItems
            dictItem = sKVP.Split(CChar("="))
            If dictItem.Length > 1 Then
                If doUrlDecode Then
                    dict.Add(dictItem(0), Current.Server.UrlDecode(dictItem(1)))
                Else
                    dict.Add(dictItem(0), dictItem(1))
                End If
            Else ''Add the string to the dict with nothing as a value. JIC a dependant method requires some value to work
                dict.Add(dictItem(0), String.Empty)
            End If
        Next

        Return dict
    End Function

    Public Shared Function ParseLinksAndText(ByVal sComment As String) As String

        '<div><a href="http://www.google.com">www.google.com</a></div>

        Dim curIndex1 As Integer
        Dim curIndex2 As Integer
        Dim RemainingString As String = sComment
        Dim sb As New StringBuilder
        Dim UrlEndIndex As Integer
        Dim UrlString As String
        Dim iCurrentLinkNumber As Integer = 1
        If sComment.Contains("<div>") Then
            Return sComment
        End If

        If RemainingString.ToLower.Contains("www.") Or RemainingString.ToLower.Contains("http://") Then

            Do While RemainingString.ToLower.Contains("www.") Or RemainingString.ToLower.Contains("http://")
                curIndex1 = RemainingString.IndexOf("http://")
                curIndex2 = RemainingString.IndexOf("www.")

                If curIndex1 > -1 AndAlso curIndex2 > -1 Then
                    curIndex1 = Math.Min(curIndex1, curIndex2)
                Else
                    curIndex1 = Math.Max(curIndex1, curIndex2)
                End If

                If curIndex1 = -1 Then
                    sb.Append(RemainingString)
                    Exit Do
                End If

                UrlEndIndex = RemainingString.IndexOfAny(New Char() {" "c, ControlChars.Lf, ","c}, curIndex1 + 2)
                If UrlEndIndex = -1 Then
                    UrlString = RemainingString.Substring(curIndex1)
                Else
                    UrlString = RemainingString.Substring(curIndex1, UrlEndIndex - curIndex1)
                End If

                sb.Append(RemainingString.Substring(0, curIndex1))
                sb.Append("<div><a target=""_blank"" href=""")

                If UrlString.IndexOf("http://") = -1 Then
                    sb.Append("http://")
                End If

                'add the url
                sb.Append(UrlString)
                sb.Append(""">")
                If UrlString.Trim.Length > 30 Then
                    sb.Append("Link - " + iCurrentLinkNumber.ToString)
                Else
                    sb.Append(UrlString)
                End If

                sb.Append("</a></div> ")

                If UrlEndIndex = -1 Then
                    RemainingString = String.Empty
                Else
                    RemainingString = RemainingString.Substring(UrlEndIndex)
                End If
                iCurrentLinkNumber += 1
            Loop

            sb.Append(RemainingString)

            Return sb.ToString
        Else
            Return sComment
        End If
    End Function

    Public Shared Function GetConcatenationMethodString(ByVal concatMethod As ConcatenationMethod) As String
        Select Case concatMethod
            Case ConcatenationMethod.Html
                Return String.Format("{0}{0}", "<br />")
            Case ConcatenationMethod.Textbox
                Return String.Format("{0}{0}", vbCrLf)
            Case Else
                Return String.Empty
        End Select
    End Function

    Public Shared Function ConcatenateStringsForDisplay(ByVal concatMethod As ConcatenationMethod, ByVal ParamArray strings() As String) As String
        Dim sb As New StringBuilder()

        Dim concatMethodStr As String = GetConcatenationMethodString(concatMethod)

        For i As Integer = 0 To UBound(strings)
            If Not String.IsNullOrEmpty(strings(i)) Then
                If Not String.IsNullOrEmpty(sb.ToString()) Then
                    sb.Append(concatMethodStr)
                End If
                'NOTE: we could take this a step further by calling strings(i).FormatPlainTextForHTML() if the concatMethod is Html... take the burden off the caller
                sb.Append(strings(i))
            End If
        Next

        Return sb.ToString()
    End Function

    Public Shared Function ReplaceStringIgnoreCase(ByVal origString As String, ByVal findString As String, ByVal replaceWith As String) As String
        Return Regex.Replace(origString, findString, replaceWith, RegexOptions.IgnoreCase)
    End Function

    Public Shared Function RandomNumberString(ByVal n As Integer) As String
        Dim s As String = ""
        Dim r As Random = New Random()
        For i As Integer = 0 To n - 1
            s = s & r.Next(10).ToString
        Next
        Return (s)
    End Function

    Public Shared Function ReplaceUnallowedCharactersInFileName(ByVal sourceStr As String) As String
        Return sourceStr.Replace(Chr(34), "").Replace(":", "").Replace("\", "").Replace("/", "").Replace("|", "").Replace("?", "").Replace("*", "").Replace("<", "").Replace(">", "")
    End Function

    Public Shared Function ReplaceInvincibleCharactersWithWhiteSpace(ByVal sourceStr As String) As String
        Dim result As String = sourceStr
        result = result.Replace(Chr(8), Chr(32)) 'replace BackSpace character with space
        result = result.Replace(Chr(9), Chr(32)) 'replace Horizontal TAB character with space
        result = result.Replace(Chr(10), Chr(32)) 'replace New Line (or LineFeed) character with space
        result = result.Replace(Chr(11), Chr(32)) 'replace Vertical TAB character with space
        result = result.Replace(Chr(12), Chr(32)) 'replace Form Feed character with space
        result = result.Replace(Chr(13), Chr(32)) 'replace Carriage Return character with space

        Return result
    End Function

    Public Shared Function ReplaceNewLineCharacters(ByVal sourceStr As String, ByVal replacementStr As String) As String
        If Not String.IsNullOrEmpty(sourceStr) Then
            Return sourceStr.Replace(Chr(10), replacementStr).Replace(Chr(13), replacementStr) ' replace New Line (or LineFeed) or Carriage Return characters with replacement string
        End If
        Return sourceStr
    End Function

    Public Shared Function AddPrefixToID(ByVal BuilderPrefix As System.Data.SqlTypes.SqlString, ByVal JobsitePrefix As System.Data.SqlTypes.SqlString, ByVal ID As String) As String
        Dim prefix As New StringBuilder()
        prefix.Append(BuilderPrefix.GetTextIfNotNullOrWhiteSpace(Function(s) String.Format("{0}-", s)))
        prefix.Append(JobsitePrefix.GetTextIfNotNullOrWhiteSpace(Function(s) String.Format("{0}-", s)))
        Return String.Format("{0}{1}", prefix.ToString(), ID.ToString())
    End Function

    Public Shared Function GetEnumFromString(Of T)(value As String) As T
        If [Enum].IsDefined(GetType(T), value) Then
            Return DirectCast([Enum].Parse(GetType(T), value, True), T)
        Else
            Dim enumNames As String() = [Enum].GetNames(GetType(T))
            For Each enumName As String In enumNames
                Dim e As Object = [Enum].Parse(GetType(T), enumName)
                If value = GetDescription(DirectCast(e, [Enum])) Then
                    Return DirectCast(e, T)
                End If
            Next
        End If
        Throw New ArgumentException(String.Format("The value '{0}' does not match a valid enum name or description.'", value))
    End Function

#End Region

    'Add Other constants sub classes...
    Public Class BTRichTextEditor
        Public Shared _beginAmmendValue As String = "begin_ammend"
        Public Shared Property BeginAmmendValue() As String
            Get
                Return _beginAmmendValue
            End Get
            Set(ByVal value As String)
                _beginAmmendValue = value
            End Set
        End Property

        Public Shared _endAmmendValue As String = "end_ammend"
        Public Shared Property EndAmmendValue() As String
            Get
                Return _endAmmendValue
            End Get
            Set(ByVal value As String)
                _endAmmendValue = value
            End Set
        End Property

        Public Shared ReadOnly Property EndAmmendTag() As String
            Get
                Return String.Format("<!-- {0} -->", _endAmmendValue)
            End Get
        End Property

        Public Shared ReadOnly Property BeginAmmendTag() As String
            Get
                Return String.Format("<!-- {0} -->", _beginAmmendValue)
            End Get
        End Property

        ''' <summary>
        ''' Add a line break between the text and the ammended text
        ''' </summary> 
        Public Shared Function TextLineBreakAmmendedText(ByVal descText As String) As String
            Return descText.Replace(BeginAmmendTag(), "<br />" & BeginAmmendTag())
        End Function
    End Class

End Class
