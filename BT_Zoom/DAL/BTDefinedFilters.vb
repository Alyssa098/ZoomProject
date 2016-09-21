Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Linq

Public Class BTDefinedFilters
    Private _filterFunctions As Dictionary(Of Integer, Action(Of String))

    Public Sub New(functions As Dictionary(Of Integer, Action(Of String)))
        _filterFunctions = functions
    End Sub

    Public Sub ApplyFilters(filters As Dictionary(Of String, String))
        If filters Is Nothing Then
            Exit Sub
        End If

        For Each filterItem As KeyValuePair(Of String, String) In filters.AsEnumerable
            Dim k As String = filterItem.Key
            Dim v As String = filterItem.Value

            If String.IsNullOrWhiteSpace(v) Then
                Continue For 'Blank value, skip to next filter
            End If

            Dim filterType As Integer
            If Integer.TryParse(k, filterType) AndAlso
               _filterFunctions.ContainsKey(filterType) Then

                _filterFunctions(filterType).Invoke(v)
            End If
        Next
    End Sub
End Class