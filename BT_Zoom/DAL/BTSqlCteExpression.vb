Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Text
Imports System.Data.SqlClient
Imports System.Linq
Imports BT_Zoom.Interfaces

Public Class BtSqlCteExpression(Of TTable As {ISqlBaseTableExpression, New})
    Implements ISqlCteExpression, IBTSqlCteExpression(Of TTable)

    Private Const Sql_WITH As String = ";WITH"

#Region "Private fields"

    Private _alias As String
    Private ReadOnly _cteName As String
    Private _sqlBuilder As IBTSqlBuilderBase(Of TTable)

#End Region

    Public Sub New()

    End Sub

    Public Sub New(ByVal sqlBuilder As IBTSqlBuilderBase(Of TTable), ByVal cteName As String, ByVal [alias] As String)
        If String.IsNullOrWhitespace(cteName) OrElse sqlBuilder Is Nothing Then
            Throw New BTSqlException("This action would result in the BtSqlCteExpression having no representation")
        End If

        _sqlBuilder = sqlBuilder
        _cteName = cteName
        _alias = [alias]
    End Sub

#Region "Properties"

    Public ReadOnly Property CteName As String Implements ISqlCteExpression.CteName
        Get
            Return _cteName
        End Get
    End Property

    Public ReadOnly Property From As TTable Implements IBTSqlCteExpression(Of TTable).From
        Get
            If TypeOf _sqlBuilder Is BTSqlSelectBuilder(Of TTable) Then
                Return DirectCast(_sqlBuilder, BTSqlSelectBuilder(Of TTable)).From
            ElseIf TypeOf _sqlBuilder Is BTSqlUnionBuilder(Of TTable) Then
                Return DirectCast(_sqlBuilder, BTSqlUnionBuilder(Of TTable)).From
            Else
                Throw New NotImplementedException("Only Select and Union builders are supported for a CTE Expression")
            End If
        End Get
    End Property

    Private ReadOnly Property SelectList As List(Of ISqlSelectExpression) Implements IBTSqlCteExpression(Of TTable).SelectList
        Get
            If TypeOf _sqlBuilder Is BTSqlSelectBuilder(Of TTable) Then
                Return DirectCast(_sqlBuilder, BTSqlSelectBuilder(Of TTable)).SelectList
            ElseIf TypeOf _sqlBuilder Is BTSqlUnionBuilder(Of TTable) Then
                Return DirectCast(_sqlBuilder, BTSqlUnionBuilder(Of TTable)).BaseSelectList
            Else
                Throw New NotImplementedException("Only Select and Union builders are supported for a CTE Expression")
            End If
        End Get
    End Property

    Public ReadOnly Property Parameters As IEnumerable(Of SqlParameter) Implements IBTSqlCteExpression(Of TTable).Parameters
        Get
            Return _sqlBuilder.Parameters
        End Get
    End Property

#End Region

#Region "Implements and overrides"

    Public ReadOnly Property [Alias] As String Implements ISqlJoinable.Alias
        Get
            Return _alias
        End Get
    End Property

    Public Function GetDependencyIdentifiers() As List(Of String) Implements IHasDependencies.GetDependencyIdentifiers
        Return New List(Of String)
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements IHasDependencies.GetDependenciesByIdentifier
        Return New List(Of ISqlSelectExpression)
    End Function

    ''' <summary>
    ''' Render for From / Join Expression (not the CTE itself)
    ''' </summary>
    Public Function Render() As String Implements ISqlExpression.Render
        Return ToString()
    End Function

    Public Overrides Function ToString() As String
        If String.IsNullOrWhitespace([Alias]) Then
            Return CteName
        Else
            Return String.Format("{0} {1}", CteName, [Alias])
        End If
    End Function

    ''' <summary>
    ''' Renders the sql for the CTE itself (i.e. ;WITH clause)
    ''' </summary>
    ''' <param name="isFirst">If it's the first CTE in the list, render with keyword ;WITH else render with comma</param>
    Public Function RenderForCte(ByVal isFirst As Boolean) As String Implements ISqlCteExpression.RenderForCte
        Dim sb As New StringBuilder
        If isFirst Then
            sb.AppendFormat("{0} {1} AS", Sql_WITH, CteName)
        Else
            sb.AppendLine(",")
            sb.AppendFormat("{0} AS", CteName)
        End If
        sb.AppendLine()
        sb.AppendLine("(")
        sb.Append(_sqlBuilder.RenderForCte())
        sb.Append(")")
        Return sb.ToString
    End Function

    Public Function CteColumn(ByVal column As ISqlSelectExpression) As IBTSqlCteColumn Implements ISqlCteExpression.CteColumn
        Return New BTSqlCteColumn(Me, column.GetDataRowColumnName)
    End Function

    Public Function Star(ParamArray columnsToExlude As ISqlSelectExpression()) As ISqlSelectExpression() Implements IBTSqlCteExpression(Of TTable).Star
        Dim result As List(Of ISqlSelectExpression) = SelectList

        If columnsToExlude IsNot Nothing Then
            Dim excludeList As New List(Of ISqlSelectExpression)(columnsToExlude)
            result.RemoveAll(Function(expr As ISqlSelectExpression) excludeList.Contains(expr))
        End If

        Return result.Select(Function(cteReturnedColumn As ISqlSelectExpression) CteColumn(cteReturnedColumn)).ToArray()
    End Function


    Public Sub ChangeAlias(ByVal [alias] As String) Implements ISqlJoinable.ChangeAlias
        If String.IsNullOrWhitespace(CteName) AndAlso String.IsNullOrWhitespace([alias]) Then
            Throw New BTSqlException("This action would result in the BtSqlCteExpression having no representation")
        End If
        _alias = [alias]
    End Sub

    Public Property UseDirtyRead As Boolean = False Implements ISqlBaseTableExpression.UseDirtyRead

#End Region

#Region "Methods"

    Public Sub AddUnion(ParamArray unionQueries As ISqlUnionable()) Implements IBTSqlCteExpression(Of TTable).AddUnion
        Dim unionBuilder As BTSqlUnionBuilder(Of TTable) = TryCast(_sqlBuilder, BTSqlUnionBuilder(Of TTable))
        If unionBuilder IsNot Nothing Then
            unionBuilder.AddUnions(unionQueries)
        Else
            Dim selectBuilder As BTSqlSelectBuilder(Of TTable) = TryCast(_sqlBuilder, BTSqlSelectBuilder(Of TTable))
            If selectBuilder Is Nothing Then
                Throw New BTSqlException("Cannot add a union to this sql builder.")
            Else
                _sqlBuilder = New BTSqlUnionBuilder(Of TTable)(selectBuilder, unionQueries)
            End If
        End If
    End Sub

#End Region
End Class