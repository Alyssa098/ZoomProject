Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Linq
Imports BT_Zoom.Interfaces

Public Class BTSqlGenericFromTable
    Inherits BTSqlTable
    Implements ICanUpdateOwnerForPaging, ISqlUnionable, IBTSqlGenericFromTable

#Region "Constructors"

    Public Sub New()
        MyBase.New("", "", "gtbl")
    End Sub

    Public Sub New(ByVal sql As ISqlSelectQuery, ByVal [alias] As String)
        MyBase.New("", "", [alias])
        SubQuery = sql
    End Sub

    Public Sub New(ByVal subQuerySql As ISqlSelectBuilder, ByVal [alias] As String)
        Me.New(New BTSqlSelectQuery(subQuerySql), [alias])
    End Sub

#End Region

#Region "Properties"

    Public Property SubQuery As ISqlSelectQuery Implements IBTSqlGenericFromTable.SubQuery

    Private Property SelectableColumns As New Dictionary(Of String, ISqlSelectExpression)

#End Region

#Region "Overrides and Implements"

    ''' <summary>
    ''' Overrides the BTSqlTable ToString to render the sub query in place of the FROM table
    ''' </summary>
    Public Overrides Function ToString() As String
        If SubQuery Is Nothing Then
            Throw New BTSqlException("Cannot render nothing.  Please define the sub query.")
        ElseIf String.IsNullOrWhiteSpace([Alias]) Then
            Throw New BTSqlException("Sub query must have an alias.")
        End If
        Return String.Format("({0}) {1}", SubQuery.RenderForUnion(), [Alias])
    End Function

    Public Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String)) Implements ICanUpdateOwnerForPaging.UpdateOwnerForPaging
        SubQuery.UpdateOwnerForPaging(pagingAlias, ownersToChange)
    End Sub

    ''' <summary>
    ''' Renders without the Parentheses and table Alias
    ''' </summary>
    Public Function RenderForUnion() As String Implements ISqlUnionable.RenderForUnion
        Return SubQuery.RenderForUnion()
    End Function

#End Region

#Region "Methods"

    ''' <summary>
    ''' Adds a column to the table that will be returned as a select expression in SelectStar
    ''' </summary>
    ''' <param name="column">Select expression that will be used to get the column name in select list of sub query</param>
    ''' <param name="alias">Alias given to the column in output select list</param>
    ''' <returns>Returns a BTSqlColumn with the name and alias given</returns>
    ''' <remarks>Assumes the Subquery will output a column with the given columnName</remarks>
    Public Function AddSelectableColumn(ByVal column As ISqlSelectExpression, Optional ByVal [alias] As String = "") As IBTSqlColumn Implements IBTSqlGenericFromTable.AddSelectableColumn
        Dim columnName As String = column.GetDataRowColumnName()
        If String.IsNullOrEmpty([alias]) AndAlso SelectableColumns.ContainsKey(columnName) OrElse SelectableColumns.ContainsKey([alias]) Then
            Throw New BTSqlException("Column names must be unique.")
        End If
        Dim col As IBTSqlColumn = TryCast(column, BTSqlColumn)
        If col Is Nothing OrElse Not col.Table.Equals(Me) Then
            col = NewColumn(columnName, [alias])
        End If
        SelectableColumns.Add(col.GetDataRowColumnName, col)
        Return col
    End Function

    Public Function NewColumn(ByVal columnName As String, Optional ByVal [alias] As String = "") As IBTSqlColumn Implements IBTSqlGenericFromTable.NewColumn
        Return New BTSqlColumn(Me, columnName) With {.Alias = [alias]}
    End Function

    Public Sub AddSelectableColumns(ByVal ParamArray columns As ISqlSelectExpression()) Implements IBTSqlGenericFromTable.AddSelectableColumns
        For Each c As ISqlSelectExpression In columns
            AddSelectableColumn(c)
        Next
    End Sub

    Public Function SelectableColumn(ByVal column As ISqlSelectExpression) As ISqlSelectExpression Implements IBTSqlGenericFromTable.SelectableColumn
        Return SelectableColumns(column.GetDataRowColumnName)
    End Function

    ''' <summary>
    ''' Returns a column that will not be output in the select list but could be used in filters
    ''' </summary>
    ''' <param name="columnName">Column name in select list of sub query</param>
    ''' <returns>Returns a BTSqlColumn with the name given and Me as the owner Table</returns>
    ''' <remarks>Assumes the Subquery will output a column with this name</remarks>
    Public Function NonSelectableColumn(ByVal columnName As String) As IBTSqlColumn Implements IBTSqlGenericFromTable.NonSelectableColumn
        Return NewColumn(columnName)
    End Function

    Public Function NonSelectableColumn(ByVal column As ISqlSelectExpression) As IBTSqlColumn Implements IBTSqlGenericFromTable.NonSelectableColumn
        Return NonSelectableColumn(column.GetDataRowColumnName)
    End Function

    Public Function SelectStar() As IEnumerable(Of ISqlSelectExpression) Implements IBTSqlGenericFromTable.SelectStar
        Return SelectableColumns.Values.ToArray()
    End Function

    ''' <summary>
    ''' Returns a list of BTSqlColumns from the SelectableColumns dictionary
    ''' </summary>
    ''' <param name="tableAlias">Columns in returned list will use this this as the table alias</param>
    ''' <remarks>If we do a select star on BTSqlTable, this could be an override</remarks>
    Public Function SelectStar(ByVal tableAlias As String) As IEnumerable(Of ISqlSelectExpression) Implements IBTSqlGenericFromTable.SelectStar
        Dim outerTable As New BTSqlTable("", "", tableAlias)
        Return (From c In SelectableColumns.Values
                Select New BTSqlColumn(outerTable, c.GetDataRowColumnName))
    End Function

#End Region

End Class
