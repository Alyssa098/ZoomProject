Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports System.Globalization
Imports System.Collections.Generic
Imports System.Data.SqlTypes
Imports System.ComponentModel
Imports System.Data
Imports System.Text.RegularExpressions
Imports System.Web.UI.WebControls
Imports System.Linq
Imports System.Web.UI.HtmlControls
Imports System.Data.SqlClient
Imports System.Text
Imports System.Web.UI
Imports BT_Zoom.Interfaces

Public Module modExtensions
    
    <Extension()> _
    Public Function GetTextIfNotNullOrWhiteSpace(sqlStr As SqlString, ByVal toStringMethod As Func(Of String, String)) As String
        If sqlStr.IsNullOrWhiteSpace Then
            Return String.Empty
        Else
            Return toStringMethod(sqlStr.Value)
        End If
    End Function

    <Extension()> _
    Public Function GetTextIfNotNullOrWhiteSpace(sqlStr As SqlString) As String
        Return sqlStr.GetTextIfNotNullOrWhiteSpace(Function(x) x)
    End Function

    <Extension()> _
    Public Function IsNullOrWhiteSpace(sqlStr As SqlString) As Boolean
        If Not sqlStr.IsNull Then
            Return String.IsNullOrWhitespace(sqlStr.Value)
        Else
            Return True
        End If
    End Function

    <Extension()>
    Public Function IsNullOrEmpty(Of T)(ByVal list As IEnumerable(Of T)) As Boolean
        Return list Is Nothing OrElse Not list.Any()
    End Function

    <Extension()> _
    Public Function GetDefaultIfNull(sqlStr As SqlString, Optional ByVal defaultValue As String = "") As String
        If sqlStr.IsNull Then
            Return defaultValue
        Else
            Return sqlStr.Value
        End If
    End Function

    <Extension()> _
    Public Function GetDefaultIfNull(sqlInt As SqlInt32, Optional ByVal defaultValue As Integer = 0) As Integer
        If sqlInt.IsNull Then
            Return defaultValue
        Else
            Return sqlInt.Value
        End If
    End Function

    <Extension()>
    Public Function GetDefaultIfNull(sqlInt As IBTSqlInt32, Optional defaultValue As Integer = 0) As Integer
        Return GetDefaultIfNull(sqlInt.Obj, defaultValue)
    End Function

    <Extension()>
    Public Function GetDefaultIfNull(sqlString As IBTSqlString, Optional defaultValue As String = "") As String
        Return GetDefaultIfNull(sqlString.Obj, defaultValue)
    End Function

    <Extension()> _
    Public Function GetDefaultIfNull(sqlMoney As SqlMoney, Optional ByVal defaultValue As Integer = 0) As Double
        If sqlMoney.IsNull Then
            Return defaultValue
        Else
            Return sqlMoney.Value
        End If
    End Function

    <Extension()> _
    Public Function GetDefaultIfNull(sqlMoney As SqlMoney, ByVal defaultValue As Decimal) As Decimal
        If sqlMoney.IsNull Then
            Return defaultValue
        Else
            Return sqlMoney.Value
        End If
    End Function

    <Extension()> _
    Public Function GetDefaultIfNull(sqlBool As SqlBoolean, Optional ByVal defaultValue As Boolean = False) As Boolean
        If sqlBool.IsNull Then
            Return defaultValue
        Else
            Return sqlBool.Value
        End If
    End Function

    <Extension()>
    Public Function GetDefaultIfNull(sqlbool As IBTSqlBoolean, Optional defaultValue As Boolean = False) As Boolean
        Return GetDefaultIfNull(sqlbool.Obj, defaultValue)
    End Function

    <Extension()> _
    Public Function GetDefaultIfNull(sqlDate As SqlDateTime, Optional ByVal defaultValue As DateTime? = Nothing) As DateTime?
        If sqlDate.IsNull Then
            Return defaultValue
        Else
            Return sqlDate.Value
        End If
    End Function

#Region "String Methods"

    ''' <summary>
    ''' Right now, just replacing vbcrlf with a br tag and quotes with &quot;
    ''' </summary>
    ''' <param name="sString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function FormatPlainTextForHTML(ByVal sString As String) As String
        Return sString.Replace(vbCrLf, "<br>")
    End Function

    ''' <summary>
    ''' Replace any characters in a string that could cause problems displaying in HTML. Replaces quotes(") and apostrophes(')
    ''' </summary>
    ''' <param name="s"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function EscapeStringsForHTML(ByVal s As String) As String
        Return s.Replace("""", "&quot;").Replace("'", "&#39;")
    End Function

    ''' <summary>
    ''' Remove HTML tags and return plain text
    ''' </summary>
    ''' <param name="Text"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function FormatHTMLAsPlainText(ByVal Text As String, Optional ByVal regexOptions As RegexOptions = RegexOptions.None) As String
        Text = System.Web.HttpUtility.HtmlDecode(Text)
        Text = Text.Replace("</div>", vbCrLf)
        Text = Text.Replace("<br>", vbCrLf)
        Text = Text.Replace("<br />", vbCrLf)
        Return Regex.Replace(Text, "<.+?>", String.Empty, regexOptions)
    End Function

    ''' <summary>
    ''' For Javascript Strings that can't use HTML tags and need escaped for quotes
    ''' </summary>
    ''' <param name="text"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function FormatStringForJS(text As String) As String
        Return text.Replace("\", "\\").Replace("'", "\'").Replace("""", "\""").Replace(vbCrLf, "\n").Replace("<br>", "\n").Replace("<br />", "\n")
    End Function

    <Extension()> _
    Public Function RemoveHeadTagsFromHTML(ByVal Text As String) As String
        Return Regex.Replace(Text, "<head>.+?</head>", String.Empty, RegexOptions.Singleline)
    End Function

    <Extension>
    Public Function RemoveUnnecessaryEmailBrackets(ByVal Text As String) As String
        If String.IsNullOrEmpty(Text) Then
            Return Text
        End If

        Dim email As String = Text.Trim()
        If email.First() = "<"c AndAlso email.Last() = ">"c Then
            email = email.Substring(1, email.Length - 2)
        End If

        Return email
    End Function

    <Extension>
    Public Function IsNullOrWhitespace(ByVal s As String) As Boolean
        Return String.IsNullOrWhitespace(s)
    End Function

    ''' <summary>
    ''' Convert a string to camelcase using TextInfo.ToTitleCase()
    ''' </summary>
    ''' <param name="s"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToCamelCase(s As String) As String
        ''see below link for text translation
        ''http://stackoverflow.com/questions/1206019/converting-string-to-title-case-in-c-sharp
        Dim textInfo As TextInfo = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo
        Dim titleCase As String = textInfo.ToTitleCase(s.ToLower())
        titleCase = Regex.Replace(titleCase, "\s+", "")
        titleCase = Regex.Replace(titleCase, "\/", "")
        s = Char.ToLowerInvariant(titleCase(0)) + titleCase.Substring(1)
        Return s
    End Function

    ''' <summary>
    ''' Check to see if a string has HTML inside of it
    ''' </summary>
    ''' <param name="s">string to check</param>
    ''' <remarks>
    ''' Using Regex. HttpUtility.HtmlEncode() does not work if you string contains apostrophes, ampersands, etc. 
    ''' See http://stackoverflow.com/questions/204646/how-to-validate-that-a-string-doesnt-contain-html-using-c-sharp for other options
    ''' </remarks>
    <Extension()>
    Public Function ContainsHTML(s As String) As Boolean
        Dim fullTagRegex As Regex = New Regex("<\s*([^ >]+)[^>]*>.*?<\s*/\s*\1\s*>")
        Dim singleTagRegex As Regex = New Regex("<[^>]+>") ''every now and then we get single tags. Check for those JIC
        Return fullTagRegex.IsMatch(s) OrElse singleTagRegex.IsMatch(s)
    End Function

    ''' <summary>
    ''' Decode any special HTML characters in the string
    ''' </summary>
    ''' <returns>decoded string</returns>
    ''' <remarks>relies on HttpUtility.HtmlDecode .NET method</remarks>
    <Extension()>
    Public Function DecodeHTMLCharacters(s As String) As String
        Return System.Web.HttpUtility.HtmlDecode(s)
    End Function

#End Region

    ''' <summary>
    ''' This Extension Method allows us to get the description off of any enum
    ''' http://msmvps.com/blogs/deborahk/archive/2009/07/10/enum-binding-to-the-description-attribute.aspx
    ''' </summary>
    ''' <param name="currentEnum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <ExtensionAttribute()> _
    Public Function GetDescription(ByVal currentEnum As [Enum]) As String
        Dim description As String = String.Empty

        Dim fi As Reflection.FieldInfo = currentEnum.GetType.GetField(currentEnum.ToString)
        Dim da As DescriptionAttribute = DirectCast(Attribute.GetCustomAttribute(fi, GetType(DescriptionAttribute)), DescriptionAttribute)
        If da IsNot Nothing Then
            description = da.Description
        Else
            description = currentEnum.ToString
        End If

        Return description
    End Function

    ''' <summary>
    ''' Get the value of an enum given a passed in description
    ''' taken from http://stackoverflow.com/questions/4367723/get-enum-from-description-attribute
    ''' </summary>
    ''' <typeparam name="T">Type of Enum</typeparam>
    ''' <param name="description">Description of enum to get</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <ExtensionAttribute()> _
    Public Function GetValueFromDescription(Of T)(description As String) As T
        Dim type As Type = GetType(T)
        If Not type.IsEnum Then
            Throw New InvalidOperationException()
        End If

        For Each field As Reflection.FieldInfo In type.GetFields()
            Dim da As DescriptionAttribute = DirectCast(Attribute.GetCustomAttribute(field, GetType(DescriptionAttribute)), DescriptionAttribute)

            If da IsNot Nothing Then
                If da.Description.Equals(description) Then
                    Return CType(field.GetValue(Nothing), T)
                Else
                    If field.Name.Equals(description) Then
                        Return CType(field.GetValue(Nothing), T)
                    End If
                End If
            End If
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Return Json object name for an enumeration. Converts first character to lower case.
    ''' </summary>
    ''' <param name="currentEnum">Enum to convert</param>
    <ExtensionAttribute()>
    Public Function GetJsonName(ByVal currentEnum As [Enum]) As String
        Dim name As String = currentEnum.ToString()
        Return Char.ToLower(name(0)) & name.Substring(1)
    End Function

    <Extension>
    Public Function GetIntegerValue(ByVal currentEnum As [Enum]) As Integer
        Return Convert.ToInt32(currentEnum)
    End Function

    ''' <summary>
    ''' Shortcut for .GetIntegerValue().ToString()
    ''' </summary>
    ''' <param name="currentEnum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <ExtensionAttribute()> _
    Public Function GetIntegerStringValue(ByVal currentEnum As [Enum]) As String
        Return currentEnum.GetIntegerValue().ToString()
    End Function


    ''' <summary>
    ''' Gets the attribue of a given attribute class
    ''' </summary>
    ''' <param name="value">The value of the enum to retreive the attribute for</param>
    <Extension()> _
    Public Function GetAttribute(Of TAttribute As Attribute)(value As [Enum]) As TAttribute
        Dim type As System.Type = value.[GetType]()
        Dim name As String = [Enum].GetName(type, value)
        Return type.GetField(name).GetCustomAttributes(False).OfType(Of TAttribute)().SingleOrDefault()
    End Function

    ''' <summary>
    ''' Function to determine whether the current string is a valid integer value and greater than the passed in minValue
    ''' </summary>
    ''' <param name="currentString">String to try and parse</param>
    ''' <param name="minValue">Min value to check against</param>
    ''' <returns>True if current string is a valid value</returns>
    <Extension()>
    Public Function IsValidIntegerValue(ByVal currentString As String, Optional minValue As Integer = 0) As Boolean
        If currentString Is Nothing Then
            Return False
        End If

        Dim parsedValue As Integer
        If Integer.TryParse(currentString, parsedValue) Then
            Return parsedValue > minValue
        Else
            Return False
        End If
    End Function

    <ExtensionAttribute()> _
    Public Function ToList(Of T)(ByVal dt As DataTable, ByVal toListMethod As Func(Of DataRow, T)) As List(Of T)
        Dim result As New List(Of T)
        For Each dr As DataRow In dt.Rows
            result.Add(toListMethod(dr))
        Next
        Return result
    End Function

    <ExtensionAttribute()> _
    Public Function ToList(Of T)(dt As DataTable,
                                     toListMethod As Func(Of DataRow, T),
                                     defaultOption As T) As List(Of T)
        Dim result As New List(Of T)
        result.Add(defaultOption)
        result.AddRange(dt.ToList(toListMethod))

        Return result
    End Function

    <ExtensionAttribute()> _
    Public Function ToDictionary(Of T, W)(dt As DataTable,
                                              toDictionaryKeyMethod As Func(Of DataRow, T),
                                              toDictionaryValueMethod As Func(Of DataRow, W)) As Dictionary(Of T, W)
        Dim result As New Dictionary(Of T, W)
        For Each dr As DataRow In dt.Rows
            Dim key As T = toDictionaryKeyMethod(dr)
            If result.ContainsKey(key) Then
                Continue For
            End If

            result.Add(key, toDictionaryValueMethod(dr))
        Next
        Return result
    End Function

    <Extension()>
    Public Function ToHashSet(Of T)(ByVal source As IEnumerable(Of T)) As HashSet(Of T)
        Return New HashSet(Of T)(source)
    End Function

    <Extension>
    Public Function ContainsAny(Of T)(ByVal collection As IEnumerable(Of T), ByVal items As IEnumerable(Of T)) As Boolean
        Return items.Any(Function(i) collection.Contains(i))
    End Function

    <Extension()> _
    Public Sub Clear(ByVal sw As IO.StringWriter)
        Dim sb As Text.StringBuilder = sw.GetStringBuilder
        sb.Remove(0, sb.Length)
    End Sub

    <Extension()> _
    Public Function ReplaceFirstOccurence(ByVal sourceStr As String, ByVal searchStr As String, ByVal replaceStr As String) As String
        Dim pos As Integer = sourceStr.IndexOf(searchStr)
        If pos < 0 Then
            Return sourceStr
        End If
        Return String.Format("{0}{1}{2}", sourceStr.Substring(0, pos), replaceStr, sourceStr.Substring(pos + searchStr.Length))
    End Function

#Region "GridSort"

    Private Function ApplyDefaultsAndGetTd(e As EventArgs, headerId As String, isNumeric As Boolean, getSortableValueJS As String) As HtmlTableCell
        Dim td As HtmlTableCell

        Select Case e.GetType()
            Case GetType(RepeaterItemEventArgs)
                td = DirectCast(DirectCast(e, RepeaterItemEventArgs).Item.FindControl(headerId), HtmlTableCell)
            Case GetType(DataGridItemEventArgs)
                td = DirectCast(DirectCast(e, DataGridItemEventArgs).Item.FindControl(headerId), HtmlTableCell)
            Case Else
                Throw New NotImplementedException(String.Format("EventArgs type '{0}' not implemented for GetTd().", e.GetType().ToString()))
        End Select

        If td Is Nothing Then
            Throw New Exception(String.Format("Column '{0}' could not be found or is not marked runat=server", headerId))
        Else
            If String.IsNullOrWhitespace(td.Attributes("class")) Then
                td.Attributes("class") = If(isNumeric, "gsColNum", "gsCol")
            Else
                td.Attributes("class") += If(isNumeric, " gsColNum", " gsCol")
            End If

            If Not String.IsNullOrWhitespace(getSortableValueJS) Then
                td.Attributes.Add("data-getvaluejsfunc", getSortableValueJS)
            End If
        End If

        Return td
    End Function

    <Extension()> _
    Public Sub AddSortableDateAttribute(td As HtmlTableCell, dt As SqlDateTime)
        If td Is Nothing Then
            Throw New Exception("Column not found or not declared as runat=server.")
        End If
        If Not dt.IsNull Then
            td.Attributes.Add("dt", dt.Value.ToString("yyyyMMddHHmm"))
        End If
    End Sub

    <Extension()> _
    Public Sub AddSortableDateAttribute(td As HtmlTableCell, dt As DateTime)
        If td Is Nothing Then
            Throw New Exception("Column not found or not declared as runat=server.")
        End If
        td.Attributes.Add("dt", dt.ToString("yyyyMMddHHmm"))
    End Sub

    <Extension()> _
    Public Sub AddSortableCurrencyAttribute(td As HtmlTableCell, d As SqlDecimal)
        If td Is Nothing Then
            Throw New Exception("Column not found or not declared as runat=server.")
        End If
        If Not d.IsNull Then
            td.Attributes.Add("cur", d.Value.ToString("0.##"))
        End If
    End Sub

    <Extension()> _
    Public Sub AddSortableCurrencyAttribute(td As HtmlTableCell, d As Decimal)
        If td Is Nothing Then
            Throw New Exception("Column not found or not declared as runat=server.")
        End If
        td.Attributes.Add("cur", d.ToString("0.##"))
    End Sub

    <Extension()> _
    Public Sub AddSortableCurrencyAttribute(td As HtmlTableCell, d As Double)
        If td Is Nothing Then
            Throw New Exception("Column not found or not declared as runat=server.")
        End If
        td.Attributes.Add("cur", d.ToString("0.##"))
    End Sub

    <Extension()>
    Public Sub AddSortableValueAttribute(td As HtmlTableCell, ByVal v As [Enum])
        td.AddSortableValueAttribute(v.GetIntegerValue.ToString)
    End Sub

    <Extension()>
    Public Sub AddSortableValueAttribute(td As HtmlTableCell, ByVal v As String)
        If td Is Nothing Then
            Throw New Exception("Column not found or not declared as runat=server.")
        End If
        td.Attributes.Add("v", v)
    End Sub

    <Extension()> _
    Public Sub MakeHeaderSortableByText(e As EventArgs, headerId As String, Optional getSortableValueJS As String = "")
        ApplyDefaultsAndGetTd(e, headerId, False, getSortableValueJS)
    End Sub

    <Extension()> _
    Public Sub MakeHeaderSortableByNumber(e As EventArgs, headerId As String, Optional getSortableValueJS As String = "")
        ApplyDefaultsAndGetTd(e, headerId, True, getSortableValueJS)
    End Sub

    <Extension()> _
    Public Sub MakeHeaderSortableByCurrency(e As EventArgs, headerId As String, Optional getSortableValueJS As String = "")
        Dim td As HtmlTableCell = ApplyDefaultsAndGetTd(e, headerId, True, getSortableValueJS)

        If String.IsNullOrWhitespace(getSortableValueJS) Then
            td.Attributes.Add("data-getvaluejsfunc", "getSortValCurrency")
        End If
    End Sub

    <Extension()> _
    Public Sub MakeHeaderSortableByDate(e As EventArgs, headerId As String, Optional getSortableValueJS As String = "")
        Dim td As HtmlTableCell = ApplyDefaultsAndGetTd(e, headerId, False, getSortableValueJS)

        If String.IsNullOrWhitespace(getSortableValueJS) Then
            td.Attributes.Add("data-getvaluejsfunc", "getSortValDate")
        End If
    End Sub

#End Region

#Region "Data Access"

    <Extension()>
    Public Function ToSqlString(ByVal val As String) As SqlString
        Return New SqlString(val)
    End Function

    <Extension()>
    Public Function ToSqlInt32(ByVal val As Int32) As SqlInt32
        Return New SqlInt32(val)
    End Function

    <Extension()>
    Public Function ToSqlInt32(ByVal val As Integer?) As SqlInt32
        If val.HasValue Then
            Return New SqlInt32(val.Value)
        End If
        Return SqlInt32.Null
    End Function

    <Extension()>
    Public Function ToSqlByte(ByVal val As Byte) As SqlByte
        Return New SqlByte(val)
    End Function

    <Extension()>
    Public Function ToSqlInt64(ByVal val As Int64) As SqlInt64
        Return New SqlInt64(val)
    End Function

    <Extension()>
    Public Function ToSqlInt64(ByVal val As Int32) As SqlInt64
        Return New SqlInt64(val)
    End Function

    <Extension()>
    Public Function ToSqlDateTime(ByVal val As DateTime) As SqlDateTime
        Return New SqlDateTime(val)
    End Function

    <Extension()>
    Public Function ToSqlDateTime(ByVal val As DateTime?) As SqlDateTime
        If val.HasValue Then
            If val.Value.Year < SqlDateTime.MinValue.Value.Year Then
                Return SqlDateTime.Null ''Check if value is below SqlDateTime.Min, if so return null
            Else
                Return val.Value.ToSqlDateTime()
            End If
        Else
            Return SqlDateTime.Null
        End If
    End Function

    <Extension()>
    Public Function ToSqlBoolean(ByVal val As Boolean) As SqlBoolean
        Return New SqlBoolean(val)
    End Function

    <Extension()>
    Public Function ToSqlDecimal(ByVal val As Decimal) As SqlDecimal
        Return New SqlDecimal(val)
    End Function

    <Extension()>
    Public Function ToSqlDecimal(ByVal val As Decimal?) As SqlDecimal
        If val.HasValue Then
            Return New SqlDecimal(val.Value)
        End If
        Return SqlDecimal.Null
    End Function

    <Extension()>
    Public Function ToSqlDouble(ByVal val As Double) As SqlDouble
        Return New SqlDouble(val)
    End Function

    ''' <summary>
    ''' Gets a value from the data reader by column name as the type specified, or if null, the default of the type specified. (Sql type will return null as default).
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="reader"></param>
    ''' <param name="columnName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function GetValueOrDefault(Of T)(ByVal reader As IDataReader, ByVal columnName As String) As T
        Return reader.GetValueOrDefault(Of T)(reader.GetOrdinal(columnName))
    End Function

    ''' <summary>
    ''' Gets a value from the data reader by column index as the type specified, or if null, the default of the type specified. (Sql type will return null as default).
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="reader"></param>
    ''' <param name="columnIndex"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension>
    Public Function GetValueOrDefault(Of T)(ByVal reader As IDataReader, ByVal columnIndex As Integer) As T
        Dim defaultValue As T
        Dim val As Object = reader.GetValue(columnIndex)

        If val Is Nothing OrElse IsDBNull(val) Then
            Return defaultValue
        Else
            Select Case GetType(T)
                Case GetType(SqlInt32)
                    Dim i32 As New SqlInt32(CInt(val))
                    Return CType(DirectCast(i32, Object), T)
                Case GetType(SqlInt64)
                    Dim i64 As New SqlInt64(CLng(val))
                    Return CType(DirectCast(i64, Object), T)
                Case GetType(SqlBoolean)
                    Dim b As New SqlBoolean(CBool(val))
                    Return CType(DirectCast(b, Object), T)
                Case GetType(SqlDateTime)
                    Dim d As New SqlDateTime(CDate(val))
                    Return CType(DirectCast(d, Object), T)
                Case GetType(SqlString)
                    Dim s As New SqlString(CStr(val))
                    Return CType(DirectCast(s, Object), T)
                Case Else
                    Return CType(val, T)
            End Select
        End If
    End Function

    <Extension()>
    Public Function ToSqlMoney(ByVal val As Decimal) As SqlMoney
        Return New SqlMoney(val)
    End Function

    <Extension()>
    Public Function TryParseSqlInt32(ByVal s As String, ByRef result As SqlInt32, Optional ByVal defaultToZero As Boolean = True) As Boolean
        Dim i As Integer
        Dim parsed As Boolean = Integer.TryParse(s, i)
        If parsed Then
            result = i.ToSqlInt32
        ElseIf defaultToZero Then
            result = 0.ToSqlInt32
        End If
        Return parsed
    End Function

    <Extension()>
    Public Function TryParseSqlDateTime(ByVal d As DateTime, ByRef result As SqlDateTime) As Boolean
        Dim isValid As Boolean = False

        Try
            result = New SqlDateTime(d)
            isValid = True
        Catch ste As SqlTypeException
            ' The exception is raised when the parsed date is not a valid SqlDateTime.
        End Try

        Return isValid
    End Function

    <Extension()>
    Public Function BT_CSqlDateTime(ByRef dr As DataRow, ByVal columnName As String, Optional ByVal convertEmptyStringToNull As Boolean = False, Optional ByVal useUSFormat As Boolean = False) As SqlDateTime
        If dr.IsNull(columnName) Then
            Return SqlDateTime.Null
        End If
        If convertEmptyStringToNull Then
            If String.IsNullOrWhitespace(dr(columnName).ToString()) Then
                Return SqlDateTime.Null
            End If
        End If
        If useUSFormat Then
            Return Date.Parse(dr(columnName).ToString(), New CultureInfo("en-US")).ToSqlDateTime
        End If
        Return CDate(dr(columnName)).ToSqlDateTime
    End Function

    <Extension()>
    Public Function BT_CSqlInt32(ByRef dr As DataRow, ByVal columnName As String) As SqlInt32
        If dr.IsNull(columnName) Then
            Return SqlInt32.Null
        End If
        Return CInt(dr(columnName)).ToSqlInt32
    End Function

    <Extension()>
    Public Function BT_CSqlInt32(ByRef dr As DataRow, ByVal column As IBTSqlColumnBase) As SqlInt32
        Return BT_CSqlInt32(dr, column.GetDataRowColumnName())
    End Function

    <Extension()>
    Public Function BT_CSqlInt64(ByRef dr As DataRow, ByVal columnName As String) As SqlInt64
        If dr.IsNull(columnName) Then
            Return SqlInt64.Null
        End If
        Return CLng(dr(columnName)).ToSqlInt64
    End Function

    <Extension()>
    Public Function BT_CSqlDecimal(ByRef dr As DataRow, ByVal columnName As String) As SqlDecimal
        If dr.IsNull(columnName) Then
            Return SqlDecimal.Null
        End If
        Return CDec(dr(columnName)).ToSqlDecimal
    End Function

    <Extension()>
    Public Function BT_CSqlMoney(ByRef dr As DataRow, ByVal columnName As String) As SqlMoney
        If dr.IsNull(columnName) Then
            Return SqlMoney.Null
        End If
        Return CDec(dr(columnName)).ToSqlMoney
    End Function

    <Extension()>
    Public Function BT_CSqlDouble(ByRef dr As DataRow, ByVal columnName As String) As SqlDouble
        If dr.IsNull(columnName) Then
            Return SqlDouble.Null
        End If
        Return CDbl(dr(columnName)).ToSqlDouble
    End Function

    <Extension()>
    Public Function BT_CSqlInt16(ByRef dr As DataRow, ByVal columnName As String) As SqlInt16
        If dr.IsNull(columnName) Then
            Return SqlInt16.Null
        End If
        Return New SqlInt16(CShort(dr(columnName)))
    End Function

    <Extension()>
    Public Function BT_CSqlByte(ByRef dr As DataRow, ByVal columnName As String) As SqlByte
        If dr.IsNull(columnName) Then
            Return SqlByte.Null
        End If
        Return New SqlByte(CByte(dr(columnName)))
    End Function

    <Extension()>
    Public Function BT_CSqlGuid(ByRef dr As DataRow, ByVal columnName As String) As SqlGuid
        If dr.IsNull(columnName) Then
            Return SqlGuid.Null
        End If
        Return New SqlGuid(dr(columnName).ToString())
    End Function

    <Extension()>
    Public Function BT_CSqlString(ByRef dr As DataRow, ByVal columnName As String) As SqlString
        If dr.IsNull(columnName) Then
            Return SqlString.Null
        End If
        Return dr(columnName).ToString.ToSqlString
    End Function

    <Extension()>
    Public Function BT_CSqlBoolean(ByRef dr As DataRow, ByVal columnName As String) As SqlBoolean
        If dr.IsNull(columnName) Then
            Return SqlBoolean.Null
        End If
        Return CBool(dr(columnName)).ToSqlBoolean
    End Function

    <Extension()>
    Public Function BT_CTimeSpan(ByRef dr As DataRow, ByVal columnName As String) As TimeSpan?
        If dr.IsNull(columnName) Then
            Return Nothing
        End If
        Return CType(dr(columnName), TimeSpan)
    End Function

    <Extension()>
    Public Function BT_CDate(ByRef dr As DataRow, ByVal columnName As String) As DateTime?
        If dr.IsNull(columnName) Then
            Return Nothing
        End If
        Return CDate(dr(columnName))
    End Function

    <Extension()>
    Public Function BT_CDate(ByRef dr As DataRow, ByVal column As IBTSqlColumn) As DateTime?
        Return dr.BT_CDate(column.GetDataRowColumnName())
    End Function

    <Extension()>
    Public Function BT_CDate(ByRef dr As DataRow, ByVal columnName As String, Optional ByVal defaultIfNull As Date = Nothing) As Date
        If dr.IsNull(columnName) Then
            Return defaultIfNull
        End If
        Return CDate(dr(columnName))
    End Function

    <Extension()>
    Public Function BT_CString(ByRef dr As DataRow, ByVal columnName As String) As String
        If dr.IsNull(columnName) Then
            Return String.Empty
        End If
        Return dr(columnName).ToString
    End Function

    <Extension()>
    Public Function BT_CBool(ByRef dr As DataRow, ByVal columnName As String, Optional defaultIfNull As Boolean = False) As Boolean
        If dr.IsNull(columnName) Then
            Return defaultIfNull
        End If
        Return CBool(dr(columnName))
    End Function

    <Extension()>
    Public Function BT_CBool(ByRef dr As DataRow, ByVal column As ISqlSelectExpression, Optional defaultIfNull As Boolean = False) As Boolean
        Return dr.BT_CBool(column.GetDataRowColumnName(), defaultIfNull)
    End Function

    <Extension()>
    Public Function BT_CInt(ByRef dr As DataRow, ByVal columnName As String, Optional defaultIfNull As Integer = 0) As Integer
        If dr.IsNull(columnName) Then
            Return defaultIfNull
        End If
        Return CInt(dr(columnName))
    End Function

    <Extension()>
    Public Function BT_CInt(ByRef dr As DataRow, ByVal column As IBTSqlColumnBase, Optional defaultIfNull As Integer = 0) As Integer
        Return BT_CInt(dr, column.GetDataRowColumnName(), defaultIfNull)
    End Function

    <Extension()>
    Public Function BT_CType(Of T)(ByRef dr As DataRow, ByVal column As IBTSqlColumnBase, Optional defaultIfNull As T = Nothing) As T
        Return dr.BT_CType(Of T)(column.GetDataRowColumnName(), defaultIfNull)
    End Function

    <Extension()>
    Public Function BT_CType(Of T)(ByRef dr As DataRow, ByVal columnName As String, Optional defaultIfNull As T = Nothing) As T
        If dr.IsNull(columnName) Then
            Return defaultIfNull
        End If
        Return CType(dr(columnName), T)
    End Function

    <Extension()>
    Public Function BT_CLong(ByRef dr As DataRow, ByVal columnName As String, Optional defaultIfNull As Integer = 0) As Long
        If dr.IsNull(columnName) Then
            Return defaultIfNull
        End If
        Return CLng(dr(columnName))
    End Function

    <Extension()>
    Public Function BT_CDbl(ByRef dr As DataRow, ByVal columnName As String, Optional defaultIfNull As Double = 0) As Double
        If dr.IsNull(columnName) Then
            Return defaultIfNull
        End If
        Return CDbl(dr(columnName))
    End Function

    <Extension()>
    Public Function BT_CDec(ByRef dr As DataRow, ByVal columnName As String, Optional defaultIfNull As Decimal = 0) As Decimal
        If dr.IsNull(columnName) Then
            Return defaultIfNull
        End If
        Return CDec(dr(columnName))
    End Function

    <Extension()>
    Public Function BT_CDec(ByRef dr As DataRow, ByVal column As ISqlSelectExpression) As Decimal
        Return BT_CDec(dr, column.GetDataRowColumnName())
    End Function

    <Extension()>
    Public Function BT_CDbl(ByRef dr As DataRowView, ByVal columnName As String, Optional defaultIfNull As Double = 0) As Double
        If IsDBNull(columnName) Then
            Return defaultIfNull
        End If
        Return CDbl(dr(columnName))
    End Function

    <Extension()>
    Public Function BT_CInt(ByRef dr As DataRowView, ByVal columnName As String, Optional defaultIfNull As Integer = 0) As Integer
        Dim v As Integer = defaultIfNull
        Integer.TryParse(CStr(dr(columnName)), v)
        Return v
    End Function

    <Extension()>
    Public Function BT_CSqlInt32(ByRef dr As DataRowView, ByVal columnName As String) As SqlInt32
        Return dr.Row.BT_CSqlInt32(columnName)
    End Function

    <Extension()>
    Public Function BT_CString(ByRef dr As DataRowView, ByVal columnName As String) As String
        If dr.Row.IsNull(columnName) Then
            Return ""
        End If
        Return CStr(dr(columnName))
    End Function

    <Extension()>
    Public Function BT_CString(ByRef dr As DataRow, ByVal column As IBTSqlColumnBase) As String
        Return BT_CString(dr, column.GetDataRowColumnName())
    End Function

    <Extension()>
    Public Function BT_CBool(ByRef dr As DataRowView, ByVal columnName As String, Optional defaultIfNull As Boolean = False) As Boolean
        Dim v As Boolean = defaultIfNull
        Boolean.TryParse(CStr(dr(columnName)), v)
        Return v
    End Function

    <Extension()>
    Public Function BT_CDate(ByRef dr As DataRowView, ByVal columnName As String) As DateTime?
        If dr.Row.IsNull(columnName) Then
            Return Nothing
        End If
        Return CDate(dr(columnName))
    End Function
#End Region

    <Extension()> _
    Public Sub ReplaceColumnValues(ByVal dataTable As DataTable, ByVal columnName As String, ByVal oldValue As String, ByVal newValue As String)
        If Not dataTable.Columns.Contains(columnName) Then
            Exit Sub
        End If

        For Each row As DataRow In dataTable.Rows
            row(columnName) = row(columnName).ToString().Replace(oldValue, newValue)
        Next
    End Sub

    <Extension()>
    Public Function TryParseFromDescription(Of T)(ByVal e As [Enum], ByVal description As String, ByRef type As T) As Boolean
        Dim enumType As Type = e.GetType()
        For Each fieldInfo As Reflection.FieldInfo In enumType.GetFields()
            If fieldInfo.CustomAttributes.Count > 0 Then
                Dim descAttribute As DescriptionAttribute = DirectCast(Attribute.GetCustomAttribute(fieldInfo, GetType(DescriptionAttribute)), DescriptionAttribute)
                If descAttribute.Description = description Then
                    type = CType([Enum].Parse(fieldInfo.DeclaringType, fieldInfo.Name.ToString()), T)
                    Return True
                End If
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Adds elements from parameter IDictionary into base IDictionary. A duplicate key will use the value from base dictionary without throwing an error.
    ''' </summary>
    <Extension()>
    Public Function Merge(Of TKey, TVal)(ByVal dictA As Dictionary(Of TKey, TVal), ByVal dictB As Dictionary(Of TKey, TVal)) As Dictionary(Of TKey, TVal)
        For Each kvp As KeyValuePair(Of TKey, TVal) In dictB
            If Not dictA.ContainsKey(kvp.Key) Then
                dictA.Add(kvp.Key, kvp.Value)
            End If
        Next

        Return dictA
    End Function

    ''' <summary>
    ''' Returns Dictionary, to be used for Controller Response, for the DataTable's Row
    ''' 
    ''' Uses Column Names as Keys, and row values as Value
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="index"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function GetDictionaryByIndex(dt As DataTable,
                                             index As Integer) As Dictionary(Of String, Object)
        Dim returnDict As New Dictionary(Of String, Object)

        Dim dr As DataRow = Nothing
        Dim rowExists As Boolean = (index >= 0 AndAlso dt.Rows.Count > index)
        If rowExists Then
            dr = dt.Rows(index)
        End If

        For Each col As DataColumn In dt.Columns
            returnDict.Add(col.ColumnName, If(rowExists, dr(col), Nothing))
        Next

        Return returnDict
    End Function

    ''' <summary>
    ''' Returns an enumerable collection with all objects cloned
    ''' </summary>
    ''' <typeparam name="T">Type of source collection</typeparam>
    ''' <param name="source">Source collection to clone</param>
    <Extension()>
    Public Function Clone(Of T As {ICloneable})(ByVal source As IEnumerable(Of T)) As IEnumerable(Of T)
        Return source.Select(Function(s) DirectCast(s.Clone(), T))
    End Function

    ''' <summary>
    ''' Returns the root master page
    ''' </summary>
    ''' <param name="target"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function GetRootMaster(ByVal target As MasterPage) As MasterPage
        Dim tmp As MasterPage = target
        While (tmp.Master IsNot Nothing)
            tmp = tmp.Master
        End While
        Return tmp
    End Function

    <Extension()>
    Public Sub TrimSeconds(ByVal dt As IBTSqlDateTime)
        If dt.IsPopulated AndAlso Not dt.IsNull AndAlso (Not dt.IsTrackingIsDirty OrElse dt.IsDirty) Then
            dt.Utc = dt.Utc.Value.TrimSeconds.ToSqlDateTime
        End If
    End Sub

    <Extension()>
    Public Function TrimSeconds(ByVal dt As DateTime) As DateTime
        Return New DateTime(dt.Ticks - dt.Ticks Mod TimeSpan.TicksPerMinute)
    End Function

    <Extension()>
    Public Function TrimMilliSeconds(ByVal dt As DateTime) As DateTime
        If dt.Millisecond <> 0 Then
            Return dt.AddMilliseconds(-dt.Millisecond)
        Else
            Return dt
        End If
    End Function

    ''' <summary>
    ''' Trim the time values from your DateTime. Creates a new DateTime object
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function TrimTime(ByVal dt As DateTime) As DateTime
        Return New DateTime(dt.Year, dt.Month, dt.Day)
    End Function

    <Extension()>
    Public Function ToUnixTime(ByVal d As DateTime) As Long
        'http://stackoverflow.com/questions/7983441/unix-time-conversions-in-c-sharp
        ' We specifically want the total seconds portion of this. Wepay didn't like it when we were sending the milliseconds.
        Return CType((DateTime.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds, Long)
    End Function

    ''' <summary>
    ''' Checks to see if value is a valid enum value
    ''' </summary>
    ''' <typeparam name="T">Type of Enum to convert to</typeparam>
    ''' <param name="obj">value to translate</param>
    ''' <param name="value">translated value if successful</param>
    ''' <returns>successful</returns>
    ''' <remarks>This was moved from WebAPI\BTValidate to enable WebAPI to be broken into it's own project (bt-14653)</remarks>
    <Extension()>
    Public Function IsEnum(Of T As {Structure, IConvertible, IComparable, IFormattable})(ByVal obj As Object, ByRef value As T) As Boolean
        Dim success As Boolean = False
        Dim tempValue As T
        If [Enum].TryParse(Of T)(obj.ToString(), tempValue) Then
            If [Enum].IsDefined(GetType(T), tempValue) Then
                value = tempValue
                success = True
            End If
        End If
        Return success
    End Function

    <Extension()>
    Public Function [In](Of T)(ByVal val As T, ByVal objs As IEnumerable(Of T)) As Boolean
        Return objs.Contains(val)
    End Function

    <Extension()>
    Public Function [In](Of T)(ByVal val As T, ByVal ParamArray args() As T) As Boolean
        Return args.Contains(val)
    End Function

End Module
