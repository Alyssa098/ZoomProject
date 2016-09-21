Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Linq
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Text

Public Class Utility

    Public Shared Function ToDelimitedString(Of T)(ByVal items As List(Of T), ByVal toStringMethod As Func(Of T, String), Optional ByVal delimiter As String = ", ") As String
        If items Is Nothing Then
            Return String.Empty
        End If
        Return String.Join(delimiter, items.Select(toStringMethod).ToArray())
    End Function

    ' Uses the default to string implementation of the object
    Public Shared Function ToDelimitedString(Of T)(ByVal items As List(Of T), Optional ByVal delimiter As String = ",") As String
        Return ToDelimitedString(items, Function(item As T) item.ToString(), delimiter)
    End Function
    
    Public Shared Function FromDelimitedString(Of T)(ByVal str As String, ByVal fromStringMethod As Func(Of String, T), Optional ByVal delimiter As String = ", ") As List(Of T)
        Dim result As New List(Of T)
        If String.IsNullOrWhiteSpace(str) Then
            Return result
        End If
        Dim tokens As String() = str.Split(New String() {delimiter}, StringSplitOptions.RemoveEmptyEntries)
        For Each token As String In tokens
            result.Add(fromStringMethod(token))
        Next
        Return result
    End Function

    Public Shared Function ListOfIntegersFromDelimitedString(ByVal str As String, Optional ByVal delimiter As String = ",") As List(Of Integer)
        Dim result As New List(Of Integer)
        If String.IsNullOrWhiteSpace(str) Then
            Return result
        End If
        Dim tokens As String() = str.Split(New String() {delimiter}, StringSplitOptions.RemoveEmptyEntries)
        For Each token As String In tokens
            Dim i As Integer

            If Integer.TryParse(token, i) Then
                result.Add(i)
            End If
        Next
        Return result
    End Function

    Public Shared Function ObjectToByteArray(Of T)(obj As T) As Byte()
        If obj Is Nothing Then
            Throw New ArgumentNullException("obj")
        End If
        Using ms As New MemoryStream()
            Dim bf As New BinaryFormatter()
            bf.Serialize(ms, obj)
            Return ms.ToArray()
        End Using
    End Function

    Public Shared Function ByteArrayToObject(Of T)(bytes() As Byte) As T
        Using ms As New MemoryStream()
            ms.Write(bytes, 0, bytes.Length)
            ms.Seek(0, SeekOrigin.Begin)
            Dim bf As New BinaryFormatter()
            Dim obj As T = CType(bf.Deserialize(ms), T)
            Return obj
        End Using
    End Function

    Public Shared Function IsACommonType(Of T)() As Boolean
        Return GetType(T).IsPrimitive OrElse
               Type.Equals(GetType(T), GetType(String)) OrElse
               Type.Equals(GetType(T), GetType(DateTime)) OrElse
               Type.Equals(GetType(T), GetType(Decimal)) OrElse
               Nullable.GetUnderlyingType(GetType(T)) IsNot Nothing
    End Function

    Public Shared Function TryParseNullableDate(ByVal inputString As String, ByRef output As DateTime?) As Boolean
        If inputString Is Nothing Then
            output = Nothing
            Return True
        End If

        Dim inputNotNull As String = inputString.ToString()
        Dim dateNotNull As DateTime
        If (Date.TryParse(inputNotNull, dateNotNull)) Then
            output = dateNotNull
            Return True
        Else
            Return False
        End If
    End Function

End Class
