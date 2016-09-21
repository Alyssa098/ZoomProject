Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Text
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Linq
Imports System.Text.RegularExpressions
Imports System.Data
Imports BT_Zoom.Enums.BTDebug
Imports BT_Zoom.Interfaces

Public MustInherit Class BTSqlInsertBuilderBase(Of TTable As {ISqlWritableTableExpression, New})
    Inherits BTSqlBuilderBase(Of TTable)
    Implements ISqlInsertBuilder, IBTSqlInsertBuilderBase(Of TTable)

    Public Sub New(ByVal intoTable As TTable, Optional ByVal useLineBreaks As Boolean = True)
        MyBase.New(intoTable, useLineBreaks)
    End Sub

    Public Sub New(Optional ByVal useLineBreaks As Boolean = True)
        MyBase.New(New TTable, useLineBreaks)
    End Sub

    Public Sub Execute() Implements ISqlInsertBuilder.Execute
        DataAccessHandler.ExecuteNonQuery(Me.Render(), Me.Parameters.ToArray())
    End Sub

#Region "Constants"

    Private Const Sql_INSERT As String = "INSERT"
    Protected Const Sql_INTO As String = "INTO"

#End Region

#Region "Into"

    Public ReadOnly Property Into As TTable Implements IBTSqlInsertBuilderBase(Of TTable).Into
        Get
            Return _from
        End Get
    End Property

    Private Function GetSqlIntoClause() As String
        Return Into.Render()
    End Function

#End Region

#Region "Columns"

    Private ReadOnly _columnList As New List(Of IBTSqlColumn)
    Public ReadOnly Property ColumnList As List(Of IBTSqlColumn) Implements ISqlInsertBuilder.ColumnList
        Get
            Return _columnList
        End Get
    End Property

    Public Sub AddColumns(ByVal ParamArray columns() As IBTSqlColumn) Implements ISqlInsertBuilder.AddColumns
        If columns IsNot Nothing AndAlso columns.Length > 0 Then
            _columnList.AddRange(columns)
        End If
    End Sub

    Private Function GetSqlColumnClause() As String
        Dim sb As New StringBuilder()
        For i As Integer = 0 To ColumnList.Count - 1
            Dim column As IBTSqlColumn = ColumnList(i)
            If i > 0 Then
                sb.Append(", ")
            End If
            sb.Append(column.Render())
        Next
        Return sb.ToString()
    End Function

#End Region

#Region "BuildSql"

    ''' <summary>
    ''' Do everything that the base builder didn't already take care of...
    ''' </summary>
    Public Overrides Sub Render(ByRef sb As StringBuilder)
        BuildSql(sb)
    End Sub

    Private Sub BuildSql(ByRef sb As StringBuilder)

        Dim sqlColumns As String = GetSqlColumnClause()
        Dim sqlInto As String = GetSqlIntoClause()

        If String.IsNullOrWhitespace(sqlColumns) Then
            Throw New BTSqlException("BuildSql requires at least one column in the insert list")
        End If
        If String.IsNullOrWhitespace(sqlInto) Then
            Throw New BTSqlException(String.Format("{0} clause is empty", Sql_INTO))
        End If

        sb.Append(Sql_INSERT)

        If TopNum.HasValue Then
            If TopNum.Value <= 0 OrElse TopNum.Value > Sql_TopMax Then
                Throw New ArgumentOutOfRangeException("TopNum", TopNum.Value, String.Format("0 <= TopNum <= {0}", Sql_TopMax))
            End If
            sb.AppendFormat(" {0} {1}", Sql_TOP, TopNum.Value.ToString())
        End If

        sb.AppendFormat(" {0} {1} ({2})", Sql_INTO, sqlInto, sqlColumns)

        AppendLineOrSpace(sb)

        BuildSqlInternal(sb) 'NOTE: defer to derived classes to render its specific portion of the insert statement

        'for debugging
        Dim sb2 As New StringBuilder()
        sb2.AppendLine("--------------------------")
        sb2.AppendLine(sb.ToString())
        sb2.AppendLine("--------------------------")
        BTDebug.WriteLine(BTDebugOutputTypes.Sql, sb2.ToString())
    End Sub

    Protected MustOverride Sub BuildSqlInternal(ByRef sb As StringBuilder)

#End Region

End Class

Public Class BTSqlInsertSelectBuilder(Of TIntoTable As {ISqlTableExpression, New}, TSelectTable As {ISqlTableExpression, New})
    Inherits BTSqlInsertBuilderBase(Of TIntoTable)
    Implements IBTSqlInsertSelectBuilder(Of TIntoTable, TSelectTable)

    Public Sub New(ByVal intoTable As TIntoTable, ByVal selectTable As TSelectTable, Optional ByVal useLineBreaks As Boolean = True)
        MyBase.New(intoTable, useLineBreaks)

        _selectBuilder = New BTSqlSelectBuilder(Of TSelectTable)(selectTable)

    End Sub

    Public Sub New(ByVal intoTable As TIntoTable, ByVal selectBuilder As BTSqlSelectBuilder(Of TSelectTable), Optional ByVal useLineBreaks As Boolean = True)
        MyBase.New(intoTable, useLineBreaks)

        _selectBuilder = selectBuilder

    End Sub

    Private ReadOnly _selectBuilder As IBTSqlSelectBuilder(Of TSelectTable)
    Public ReadOnly Property SelectQuery As IBTSqlSelectBuilder(Of TSelectTable) Implements IBTSqlInsertSelectBuilder(Of TIntoTable, TSelectTable).SelectQuery
        Get
            Return _selectBuilder
        End Get
    End Property

    Protected Overrides Sub BuildSqlInternal(ByRef sb As StringBuilder)
        SelectQuery.Render(sb)
    End Sub

End Class

Public Class BTSqlInsertValuesBuilder(Of TIntoTable As {ISqlWritableTableExpression, New})
    Inherits BTSqlInsertBuilderBase(Of TIntoTable)
    Implements IBTSqlInsertValuesBuilder(Of TIntoTable)

    Public Sub New(ByVal intoTable As TIntoTable, Optional ByVal useLineBreaks As Boolean = True)
        MyBase.New(intoTable, useLineBreaks)
        intoTable.ChangeAlias("") 'no alias allowed on an insert values table expression
    End Sub

    Public Sub New(Optional ByVal useLineBreaks As Boolean = True)
        MyBase.New(New TIntoTable, useLineBreaks)
        Into.ChangeAlias("") 'no alias allowed on an insert values table expression
    End Sub

#Region "Constants"

    Private Const Sql_VALUES As String = "VALUES"
    Private Const Sql_OUTPUT As String = "OUTPUT"
    Private Const Sql_INSERTED As String = "INSERTED"
    Private Const Sql_DECLARE As String = "DECLARE"

#End Region

#Region "Values"

    Private ReadOnly _values As New List(Of List(Of ISqlSelectExpression))()
    Public ReadOnly Property Values As List(Of List(Of ISqlSelectExpression)) Implements IBTSqlInsertValuesBuilder(Of TIntoTable).Values
        Get
            Return _values
        End Get
    End Property

    Public Sub AddValues(ByVal ParamArray expressions() As ISqlSelectExpression) Implements IBTSqlInsertValuesBuilder(Of TIntoTable).AddValues
        If expressions IsNot Nothing AndAlso expressions.Length > 0 Then
            _values.Add(New List(Of ISqlSelectExpression)(expressions))
        End If
    End Sub

    Private Function GetSqlValuesClause() As String
        Dim sb As New StringBuilder()
        For i As Integer = 0 To Values.Count - 1
            If i > 0 Then
                If UseLineBreaks Then
                    sb.AppendLine()
                End If
                sb.Append(", ")
            End If
            sb.Append("(")

            Dim lst As List(Of ISqlSelectExpression) = Values(i)

            For j As Integer = 0 To lst.Count - 1
                Dim expression As ISqlSelectExpression = lst(j)
                If j > 0 Then
                    sb.Append(", ")
                End If
                sb.Append(expression.Render())
            Next

            sb.Append(")")
        Next
        If UseLineBreaks Then
            sb.AppendLine()
        End If
        Return sb.ToString()
    End Function

#End Region

#Region "Output for PKs"

    Private ReadOnly _outputParams As New List(Of Tuple(Of IBTSqlParameter, IBTSqlColumn))
    Private _outputParamsReturnAsTable As Boolean = False

    Public Sub AddOutputFromInsert(ByVal outputParams As List(Of Tuple(Of IBTSqlParameter, IBTSqlColumn))) Implements IBTSqlInsertValuesBuilder(Of TIntoTable).AddOutputFromInsert
        For Each pc As Tuple(Of IBTSqlParameter, IBTSqlColumn) In outputParams
            If pc.Item1.Parameter.Direction = ParameterDirection.ReturnValue Then
                Throw New BTSqlException("All parameters supplied for PK output must be declare as Output.")
            ElseIf pc.Item1.Parameter.Direction = ParameterDirection.Input Then
                pc.Item1.Parameter.Direction = ParameterDirection.InputOutput
            End If
            _outputParams.Add(pc)
        Next
    End Sub

    Public Sub AddOutputFromInsert(ByVal outputParams As List(Of Tuple(Of IBTSqlParameter, IBTSqlColumn)), ByVal asTable As Boolean) Implements IBTSqlInsertValuesBuilder(Of TIntoTable).AddOutputFromInsert
        _outputParamsReturnAsTable = asTable
        AddOutputFromInsert(outputParams)
    End Sub

    Private Function GetOutputTableDeclaration() As String
        If _outputParams.Any Then
            Dim sb As New StringBuilder
            sb.AppendFormat("{0} {1} {2} (", Sql_DECLARE, BTOutputTable.Sql_OUTPUT_TABLE, BTSqlUtility.Sql_TABLE)
            If UseLineBreaks Then
                sb.AppendLine()
            Else
                sb.Append(" ")
            End If
            sb.AppendLine(Utility.ToDelimitedString(_outputParams, Function(pc) pc.Item2.RenderForTable(), CommaDelimiter))
            sb.AppendLine(")")
            Return sb.ToString
        End If
        Return ""
    End Function

    Private Function CommaDelimiter() As String
        Return String.Format(", {0}", If(UseLineBreaks, vbCrLf, ""))
    End Function

    Private Sub AddOutputClause(ByRef sb As StringBuilder)
        If _outputParams.Any Then
            sb.Append(Sql_OUTPUT)
            sb.Append(" ")
            sb.AppendLine(Utility.ToDelimitedString(_outputParams, Function(pc) String.Format("{0}.{1}", Sql_INSERTED, pc.Item2.GetDataRowColumnName), CommaDelimiter))
            sb.AppendFormat("{0} {1}", Sql_INTO, BTOutputTable.Sql_OUTPUT_TABLE)
            If UseLineBreaks Then
                sb.AppendLine()
            Else
                sb.Append(" ")
            End If
        End If
    End Sub

    Private Sub AddOutputSelect(ByRef sb As StringBuilder)
        If _outputParams.Any Then
            sb.AppendLine(";")

            Dim s As New BTSqlSelectBuilder(Of BTOutputTable)(New BTOutputTable())
            s.AllowNoFilters = True
            If _outputParamsReturnAsTable Then
                s.AddSelect(s.From.Star())
            Else
                For Each o As Tuple(Of IBTSqlParameter, IBTSqlColumn) In _outputParams
                    s.AddSelect(New BTSqlTextSelectExpression("{0} = {1}", "", o.Item1, New BTSqlColumn(s.From, o.Item2.Name)))
                Next
            End If
            s.Render(sb)

            sb.Append(";")
        End If
    End Sub

#End Region

    Protected Overrides Sub BuildSqlInternal(ByRef sb As StringBuilder)
        Dim sqlValues As String = GetSqlValuesClause()
        If String.IsNullOrWhitespace(sqlValues) Then
            Throw New BTSqlException("BuildSql requires at least one expression in the insert values clause")
        End If

        sb.Insert(0, GetOutputTableDeclaration()) 'add the table declaration at the beginning
        AddOutputClause(sb)

        sb.AppendFormat("{0} {1}", Sql_VALUES, sqlValues)

        AddOutputSelect(sb)
    End Sub

    Private Class BTOutputTable
        Inherits BTSqlTable

        Public Const Sql_OUTPUT_TABLE As String = "@outputTbl"

        Public Sub New()
            MyBase.New(Sql_OUTPUT_TABLE, "", "")
        End Sub

        Public Function Star() As BTSqlTextSelectExpression
            Return New BTSqlTextSelectExpression("*", "")
        End Function

    End Class

End Class