Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports System.Text
Imports System.Data
Imports BT_Zoom.Interfaces

<Serializable>
Public MustInherit Class BTSqlColumnBase
    Implements ISqlSelectExpression, IBTSqlColumnBase

    Protected _name As String

    Public Property [Alias] As String Implements ISqlSelectExpression.Alias

    Public ReadOnly Property Name As String Implements IBTSqlColumnBase.Name
        Get
            Return _name
        End Get
    End Property

    Public MustOverride ReadOnly Property OwnerTable As IBaseTableExpression Implements IBTSqlColumnBase.OwnerTable

    Public Function WithAlias(ByVal [alias] As String) As ISqlSelectExpression Implements ISqlSelectExpression.WithAlias
        Dim result As BTSqlColumnBase = CType(Clone(), BTSqlColumnBase)
        result.Alias = [alias]
        Return result
    End Function

    Public Function GetDataRowColumnName() As String Implements ISqlSelectExpression.GetDataRowColumnName
        If Not String.IsNullOrWhiteSpace(Me.Alias) Then
            Return Me.Alias.Replace("[", String.Empty).Replace("]", String.Empty)
        Else
            Return Name
        End If
    End Function

#Region "ToString / Render"

    Public Overrides Function ToString() As String Implements IBTSqlColumnBase.ToString
        Dim useAlias As Boolean = Not String.IsNullOrWhiteSpace(Me.Alias)
        Dim sb As New StringBuilder()
        If useAlias Then
            sb.Append("(")
        End If
        Dim owner As String = RenderOwner()
        If Not String.IsNullOrWhiteSpace(owner) Then
            sb.AppendFormat("{0}.", owner)
        End If
        sb.Append(Name)
        If useAlias Then
            sb.AppendFormat(") AS {0}", Me.Alias)
        End If
        Return sb.ToString()
    End Function

    Public Function Render() As String Implements ISqlSelectExpression.Render
        Return ToString()
    End Function

    Public Overridable Function RenderForGroupBy() As String Implements ISqlSelectExpression.RenderForGroupBy
        Return RenderForOrderBy()
    End Function

    Public Overridable Function RenderForOrderBy() As String Implements ISqlSelectExpression.RenderForOrderBy
        Dim sb As New StringBuilder()

        Dim owner As String = RenderOwner()
        If Not String.IsNullOrWhiteSpace(owner) Then
            sb.AppendFormat("{0}.", owner)
        End If
        sb.Append(Name)

        Return sb.ToString()
    End Function

    Public Overridable Function RenderForFilter() As String Implements ISqlSelectExpression.RenderForFilter
        Dim useAlias As Boolean = Not String.IsNullOrWhiteSpace(Me.Alias)
        Dim sb As New StringBuilder()

        Dim owner As String = RenderOwner()
        If Not String.IsNullOrWhiteSpace(owner) Then
            sb.AppendFormat("{0}.", owner)
        End If
        If useAlias Then
            sb.Append(Me.Alias)
        Else
            sb.Append(Name)
        End If

        Return sb.ToString()
    End Function

    Public Overridable Function RenderForFunction() As String Implements ISqlSelectExpression.RenderForFunction
        Return RenderForFilter()
    End Function

    Public Overridable Function RenderOwner() As String Implements IBTSqlColumnBase.RenderOwner
        If Not String.IsNullOrWhiteSpace(_pagingAlias) Then
            Return _pagingAlias
        ElseIf Not String.IsNullOrWhiteSpace(OwnerTable.Alias) Then
            Return OwnerTable.Alias
        Else
            Return String.Empty
        End If
    End Function

    Public Overridable Function RenderForJoin() As String Implements ISqlSelectExpression.RenderForJoin
        Return RenderForOrderBy()
    End Function

    Public Overridable Function RenderForAssignment() As String Implements ISqlSelectExpression.RenderForAssignment
        Return RenderForOrderBy()
    End Function

#End Region

    Public Function GetParameters() As List(Of IBTSqlParameter) Implements ISqlSelectExpression.GetParameters
        Return New List(Of IBTSqlParameter)
    End Function

    Public ReadOnly Property Parameters As IEnumerable(Of SqlParameter) Implements IHasParameters.Parameters
        Get
            Return New List(Of SqlParameter)()
        End Get
    End Property

    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlSelectExpression.GetDependencyIdentifiers
        Dim result As New List(Of String)
        If Not String.IsNullOrWhiteSpace(OwnerTable.Alias) Then
            result.Add(OwnerTable.Alias)
        End If
        Return result
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlSelectExpression.GetDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)
        If Not String.IsNullOrWhiteSpace(OwnerTable.Alias) AndAlso owner = OwnerTable.Alias Then
            result.Add(Me)
        End If
        Return result
    End Function

    Private _pagingAlias As String = String.Empty
    Public Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String)) Implements ISqlSelectExpression.UpdateOwnerForPaging
        If Not String.IsNullOrWhiteSpace(OwnerTable.Alias) Then
            If ownersToChange.Contains(OwnerTable.Alias) Then
                _pagingAlias = pagingAlias
            End If
        End If
    End Sub

    Public MustOverride Function Clone() As Object Implements ISqlSelectExpression.Clone

    Public Function CloneWithoutAlias() As Object Implements IBTSqlColumnBase.CloneWithoutAlias
        Dim c As BTSqlColumnBase = DirectCast(Clone(), BTSqlColumnBase)
        c.Alias = String.Empty
        Return c
    End Function

End Class

<Serializable>
Public Class BTSqlColumn
    Inherits BTSqlColumnBase
    Implements IBTSqlColumn

    'cmdToExecute.Parameters.Add(New SqlParameter("@ileadID", SqlDbType.Int, 4, ParameterDirection.Output, True, 10, 0, "", DataRowVersion.Proposed, _leadID))

    Public Sub New(ByVal table As IBaseTableExpression, ByVal name As String, ByVal sqlDbType As SqlDbType, ByVal size As Integer, ByVal isNullable As Boolean, ByVal precision As Byte, ByVal scale As Byte, ByVal isPrimaryKey As Boolean, ByVal isIdentity As Boolean)
        _table = table
        _name = name
        _sqlDbType = sqlDbType
        _size = size
        _isNullable = isNullable
        _precision = precision
        _scale = scale
        _isPrimaryKey = isPrimaryKey
        _isIdentity = isIdentity
    End Sub

    Public Sub New(ByVal table As IBaseTableExpression, ByVal name As String)
        _table = table
        _name = name
    End Sub

#Region "Properties"

    Public Overrides ReadOnly Property OwnerTable As IBaseTableExpression
        Get
            Return Table
        End Get
    End Property

    Private _table As IBaseTableExpression
    Public ReadOnly Property Table As IBaseTableExpression Implements IBTSqlColumn.Table
        Get
            Return _table
        End Get
    End Property

    Private _sqlDbType As SqlDbType
    Public ReadOnly Property SqlDbType As SqlDbType Implements IBTSqlColumn.SqlDbType
        Get
            Return _sqlDbType
        End Get
    End Property

    Private _size As Integer
    Public ReadOnly Property Size As Integer Implements IBTSqlColumn.Size
        Get
            Return _size
        End Get
    End Property

    Private _isNullable As Boolean
    Public ReadOnly Property IsNullable As Boolean Implements IBTSqlColumn.IsNullable
        Get
            Return _isNullable
        End Get
    End Property

    Private _precision As Byte
    Public ReadOnly Property Precision As Byte Implements IBTSqlColumn.Precision
        Get
            Return _precision
        End Get
    End Property

    Private _scale As Byte
    Public ReadOnly Property Scale As Byte Implements IBTSqlColumn.Scale
        Get
            Return _scale
        End Get
    End Property

    Private _isPrimaryKey As Boolean
    Public ReadOnly Property IsPrimaryKey As Boolean Implements IBTSqlColumn.IsPrimaryKey
        Get
            Return _isPrimaryKey
        End Get
    End Property

    Private _isIdentity As Boolean
    Public ReadOnly Property IsIdentity As Boolean Implements IBTSqlColumn.IsIdentity
        Get
            Return _isIdentity
        End Get
    End Property

#End Region

    Public Function CreateSqlParameter(ByVal direction As ParameterDirection) As SqlParameter Implements IBTSqlColumn.CreateSqlParameter
        Dim p As New SqlParameter()
        p.ParameterName = String.Format("@{0}", Name)
        p.SqlDbType = SqlDbType
        p.Size = Size
        p.Direction = direction
        p.IsNullable = IsNullable
        p.Precision = Precision
        p.Scale = Scale
        p.SourceColumn = String.Empty
        p.SourceVersion = DataRowVersion.Proposed
        Return p

        'Return New SqlParameter("@ileadID", SqlDbType.Int, 4, ParameterDirection.Output, True, 10, 0, "", DataRowVersion.Proposed, _leadID)
    End Function

    Public Function CreateSqlParameterForInsert() As SqlParameter Implements IBTSqlColumn.CreateSqlParameterForInsert
        If IsIdentity Then
            Return CreateSqlParameter(ParameterDirection.Output)
        Else
            Return CreateSqlParameter()
        End If
    End Function

    Public Function CreateSqlParameter() As SqlParameter Implements IBTSqlColumn.CreateSqlParameter
        Return CreateSqlParameter(ParameterDirection.Input)
    End Function

    Public Function RenderForTable() As String Implements IBTSqlColumn.RenderForTable
        'like  "GlobalUserId INT NOT NULL"
        Return String.Format("{0} {1} {2}NULL", Name, ColumnDataType, If(IsNullable, "", "NOT "))
    End Function

    Private Function ColumnDataType() As String Implements IBTSqlColumn.ColumnDataType
        Select Case SqlDbType
            Case SqlDbType.NChar, SqlDbType.NVarChar, SqlDbType.VarBinary, SqlDbType.VarChar, SqlDbType.Binary, SqlDbType.Char
                Return String.Format("{0}({1})", SqlDbType.ToString, If(Size > 0, Size.ToString, "MAX"))
            Case SqlDbType.Decimal
                Return String.Format("{0}({1},{2})", SqlDbType.ToString, Precision, Scale)
            Case SqlDbType.Time
                Return String.Format("{0}({1})", SqlDbType.ToString, Scale)
            Case Else
                Return SqlDbType.ToString
        End Select
    End Function

    Public Overrides Function Clone() As Object
        'NOTE: this performs a shallow copy
        Dim c As New BTSqlColumn(Me.Table, Me.Name, Me.SqlDbType, Me.Size, Me.IsNullable, Me.Precision, Me.Scale, Me.IsPrimaryKey, Me.IsIdentity)
        c.Alias = Me.Alias
        Return c
    End Function

End Class