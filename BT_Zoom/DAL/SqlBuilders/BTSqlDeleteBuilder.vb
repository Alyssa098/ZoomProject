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

Public Class BTSqlDeleteBuilder(Of TTable As {ISqlWritableTableExpression, New})
    Inherits BTSqlBuilderBase(Of TTable)
    Implements ISqlDeleteBuilder, IBTSqlDeleteBuilder(Of TTable)

    Public Sub New(ByVal fromTable As TTable, Optional ByVal useLineBreaks As Boolean = True)
        MyBase.New(fromTable, useLineBreaks)
    End Sub

    Public Sub New(Optional ByVal useLineBreaks As Boolean = True)
        MyBase.New(New TTable, useLineBreaks)
    End Sub

    Public Function Execute() As Integer Implements ISqlDeleteBuilder.Execute
        Return DataAccessHandler.ExecuteNonQuery(Me.Render(), Me.Parameters.ToArray())
    End Function

#Region "Constants"

    Private Const Sql_DELETE As String = "DELETE"

#End Region

#Region "From"

    Public ReadOnly Property From As TTable Implements IBTSqlDeleteBuilder(Of TTable).From
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
    Public Property AllowNoFilters As Boolean = False Implements ISqlDeleteBuilder.AllowNoFilters

    Private Sub BuildSql(ByRef sb As StringBuilder)

        Dim sqlFrom As String = GetSqlFromClause()
        Dim sqlJoin As String = GetSqlJoinClause()
        Dim sqlWhere As String = GetSqlWhereClause()

        sb.Append(Sql_DELETE)

        If TopNum.HasValue Then
            If TopNum.Value <= 0 OrElse TopNum.Value > Sql_TopMax Then
                Throw New ArgumentOutOfRangeException("TopNum", TopNum.Value, String.Format("0 <= TopNum <= {0}", Sql_TopMax))
            End If
            sb.AppendFormat(" {0} ({1})", Sql_TOP, TopNum.Value.ToString())
        End If
        sb.Append(" ")

        If String.IsNullOrWhitespace(sqlFrom) Then
            Throw New BTSqlException(String.Format("{0} clause is empty", Sql_FROM))
        End If

        If Not String.IsNullOrEmpty(From.Alias) Then
            sb.AppendFormat("{0} ", From.Alias)
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
