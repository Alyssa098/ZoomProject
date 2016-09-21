Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Text
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Linq
Imports System.Text.RegularExpressions
Imports BT_Zoom.Enums.BTDebug
Imports BT_Zoom.Interfaces

Public Class BTSqlUpdateBuilder(Of TTable As {ISqlWritableTableExpression, New})
    Inherits BTSqlBuilderBase(Of TTable)
    Implements ISqlUpdateBuilder, IBTSqlUpdateBuilder(Of TTable)

    Public Sub New(ByVal fromTable As TTable, Optional ByVal useLineBreaks As Boolean = True)
        MyBase.New(fromTable, useLineBreaks)
        If String.IsNullOrWhitespace(fromTable.Alias) Then
            Throw New ArgumentException("An alias is required for the from table in an update", "fromTable")
        End If
    End Sub

    Public Sub New(Optional ByVal useLineBreaks As Boolean = True)
        MyBase.New(New TTable, useLineBreaks)
    End Sub

    Public Sub Execute() Implements ISqlUpdateBuilder.Execute
        DataAccessHandler.ExecuteNonQuery(Me.Render(), Me.Parameters.ToArray())
    End Sub

#Region "Constants"

    Private Const Sql_UPDATE As String = "UPDATE"
    Private Const Sql_SET As String = "SET"

#End Region

#Region "Set"

    Private _setList As New List(Of Tuple(Of IBTSqlColumn, ISqlAssignable))
    Public ReadOnly Property SetList As List(Of Tuple(Of IBTSqlColumn, ISqlAssignable)) Implements ISqlUpdateBuilder.SetList
        Get
            Return _setList
        End Get
    End Property
    Public Sub AddSet(column As IBTSqlColumn, rightExpression As ISqlAssignable) Implements ISqlUpdateBuilder.AddSet
        For i As Integer = 0 To _setList.Count - 1
            Dim expression As Tuple(Of IBTSqlColumn, ISqlAssignable) = _setList(i)
            If expression.Item1.RenderForAssignment() = column.RenderForAssignment() Then
                Throw New ArgumentException(String.Format("The column '{0}' was already added to the update statement", column.RenderForAssignment()), "column")
            End If
        Next
        _setList.Add(Tuple.Create(column, rightExpression))
    End Sub
    Private Function GetSqlSetClause() As String
        Dim sb As New StringBuilder()
        For i As Integer = 0 To _setList.Count - 1
            Dim expression As Tuple(Of IBTSqlColumn, ISqlAssignable) = _setList(i)
            If i > 0 Then
                If UseLineBreaks Then
                    sb.AppendLine()
                    sb.Append("  ")
                End If
                sb.Append(", ")
            End If
            sb.AppendFormat("{0} = {1}", expression.Item1.RenderForAssignment(), expression.Item2.RenderForAssignment())
        Next
        Return sb.ToString()
    End Function

#End Region

#Region "From"

    Public ReadOnly Property From As TTable Implements IBTSqlUpdateBuilder(Of TTable).From
        Get
            Return _from
        End Get
    End Property

    Private Function GetSqlFromClause() As String
        Return From.Render()
    End Function

#End Region

#Region "BuildSql"

    ''' <summary>
    ''' If true, suppress no filters error (i.e. Will not throw an error if the where clause is empty), only update this if you are sure that's what you need
    ''' </summary>
    Public Property AllowNoFilters As Boolean = False Implements ISqlUpdateBuilder.AllowNoFilters

    Private Sub BuildSql(ByRef sb As StringBuilder)

        Dim sqlSet As String = GetSqlSetClause()
        Dim sqlFrom As String = GetSqlFromClause()
        Dim sqlJoin As String = GetSqlJoinClause()
        Dim sqlWhere As String = GetSqlWhereClause()

        sb.Append(Sql_UPDATE)

        If TopNum.HasValue Then
            If TopNum.Value <= 0 OrElse TopNum.Value > Sql_TopMax Then
                Throw New ArgumentOutOfRangeException("TopNum", TopNum.Value, String.Format("0 <= TopNum <= {0}", Sql_TopMax))
            End If
            sb.AppendFormat(" {0} {1}", Sql_TOP, TopNum.Value.ToString())
        End If
        sb.AppendFormat(" {0}", From.Alias)
        AppendLineOrSpace(sb)

        If String.IsNullOrWhitespace(sqlSet) Then
            Throw New BTSqlException(String.Format("{0} clause is empty", Sql_SET))
        End If
        sb.AppendFormat("{0} {1}", Sql_SET, sqlSet)
        AppendLineOrSpace(sb)

        If String.IsNullOrWhitespace(sqlFrom) Then
            Throw New BTSqlException(String.Format("{0} clause is empty", Sql_FROM))
        End If
        sb.AppendFormat("{0} {1}", Sql_FROM, sqlFrom)
        AppendLineOrSpace(sb)

        If Not String.IsNullOrWhitespace(sqlJoin) Then
            sb.Append(sqlJoin)
        End If

        If String.IsNullOrWhitespace(sqlWhere) AndAlso Not AllowNoFilters Then
            Throw New BTSqlException(String.Format("{0} clause is empty", Sql_WHERE))
        End If
        If Not String.IsNullOrWhitespace(sqlWhere) Then
            sb.AppendFormat("{0} {1}", Sql_WHERE, sqlWhere)
        End If

        'for debugging
        Dim sb2 As New StringBuilder()
        sb2.AppendLine("--------------------------")
        sb2.AppendLine(sb.ToString())
        sb2.AppendLine("--------------------------")
        BTDebug.WriteLine(BTDebugOutputTypes.Sql, sb2.ToString())
    End Sub

    ''' <summary>
    ''' Do everything that the base builder didn't already take care of...
    ''' </summary>
    Public Overrides Sub Render(ByRef sb As StringBuilder)
        BuildSql(sb)
    End Sub

#End Region

End Class
