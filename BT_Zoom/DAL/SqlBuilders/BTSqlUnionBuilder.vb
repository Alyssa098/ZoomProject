Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Text
Imports System.Linq
Imports BT_Zoom.BTSql
Imports BT_Zoom.Interfaces

Public Class BTSqlUnionBuilder(Of TTable As {IBaseTableExpression, New})
    Inherits BTSqlBuilderBase(Of TTable)
    Implements IBTSqlUnionBuilder(Of TTable)

    Private ReadOnly _sqlQuery As IBTSqlSelectQuery
    Private ReadOnly _unionList As New List(Of ISqlUnionable)
    Private ReadOnly _baseSqlBuilder As IBTSqlSelectBuilder(Of TTable)
    Public ReadOnly Property BaseSelectList As List(Of ISqlSelectExpression) Implements IBTSqlUnionBuilder(Of TTable).BaseSelectList
        Get
            Return _baseSqlBuilder.SelectList
        End Get
    End Property

    ''' <param name="sqlBuilder">Sql builder representing the base table or CTE for the union</param>
    ''' <param name="unionQueries">Queries to union with the base From table, coupled with UnionType</param>
    Public Sub New(ByVal sqlBuilder As IBTSqlSelectBuilder(Of TTable), ByVal useLineBreaks As Boolean, ByVal ParamArray unionQueries As ISqlUnionable())
        MyBase.New(sqlBuilder.From, useLineBreaks)

        _sqlQuery = sqlBuilder.ToSelectQuery
        _baseSqlBuilder = sqlBuilder

        If unionQueries IsNot Nothing Then
            For Each q As ISqlUnionable In unionQueries
                AddUnion(q)
            Next
        End If
    End Sub

    ''' <param name="sqlBuilder">Sql builder representing the base table or CTE for the union</param>
    ''' <param name="unionQueries">Queries to union with the base From table, coupled with UnionType</param>
    Public Sub New(ByVal sqlBuilder As IBTSqlSelectBuilder(Of TTable), ByVal ParamArray unionQueries As ISqlUnionable())
        Me.New(sqlBuilder, True, unionQueries)
    End Sub

    Public ReadOnly Property From As TTable Implements IBTSqlUnionBuilder(Of TTable).From
        Get
            Return _from
        End Get
    End Property

#Region "Methods"

    Public Sub AddUnion(unionQuery As ISqlUnionable) Implements IBTSqlUnionBuilder(Of TTable).AddUnion
        _unionList.Add(unionQuery)
    End Sub

    Public Sub AddUnions(ByVal unionQueries As ISqlUnionable()) Implements IBTSqlUnionBuilder(Of TTable).AddUnions
        For Each q As ISqlUnionable In unionQueries
            AddUnion(q)
        Next
    End Sub

    Public Sub ClearUnionList() Implements IBTSqlUnionBuilder(Of TTable).ClearUnionList
        _unionList.Clear()
    End Sub

#End Region

#Region "Overrides and Implements"

    Public Overrides Sub Render(ByRef sb As StringBuilder)
        sb.Append(_sqlQuery.RenderForUnion())
        For Each q As ISqlUnionable In _unionList
            AppendLineOrSpace(sb)
            sb.Append(q.RenderForUnion())
        Next
    End Sub

#End Region

End Class
