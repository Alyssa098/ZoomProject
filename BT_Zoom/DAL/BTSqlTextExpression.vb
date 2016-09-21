Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Data.SqlClient
Imports BT_Zoom.Enums.BTSql
Imports BT_Zoom.Interfaces

<Serializable>
Public MustInherit Class BTSqlTextExpression

    Public Property Text As String

    Public Sub New(ByVal text As String)
        Me.Text = text
    End Sub

    Public Overrides Function ToString() As String
        Return Text
    End Function

End Class

<Serializable>
Public Class BTSqlTextSelectExpression
    Inherits BTSqlTextExpression
    Implements ISqlSelectExpression

    Public Sub New(ByVal text As String, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        MyBase.New(text)

        Me.Alias = [alias]
        _dependencies = New List(Of ISqlSelectExpression)()
        For Each dep As ISqlSelectExpression In dependencies
            _dependencies.Add(DirectCast(dep.Clone(), ISqlSelectExpression))
        Next
    End Sub

    Public Property [Alias] As String Implements ISqlSelectExpression.Alias

    Private _dependencies As List(Of ISqlSelectExpression)

    Public Function WithAlias(ByVal [alias] As String) As ISqlSelectExpression Implements ISqlSelectExpression.WithAlias
        Dim result As BTSqlTextSelectExpression = CType(Clone(), BTSqlTextSelectExpression)
        result.Alias = [alias]
        Return result
    End Function

    Public Function GetDataRowColumnName() As String Implements ISqlSelectExpression.GetDataRowColumnName
        If Not String.IsNullOrWhitespace(Me.Alias) Then
            Return Me.Alias.Replace("[", String.Empty).Replace("]", String.Empty)
        Else
            Return String.Format(MyBase.ToString(), _dependencies.ToArray())
        End If
    End Function

#Region "ToString / Render"

    Public Overrides Function ToString() As String
        Dim useAlias As Boolean = Not String.IsNullOrWhitespace(Me.Alias)
        Dim sb As New StringBuilder()
        If useAlias Then
            sb.Append("(")
        End If
        sb.AppendFormat(MyBase.ToString(), _dependencies.ToArray())
        If useAlias Then
            sb.AppendFormat(") AS {0}", Me.Alias)
        End If
        Return sb.ToString()
    End Function

    Public Function Render() As String Implements ISqlSelectExpression.Render
        Return ToString()
    End Function

    Public Function RenderForGroupBy() As String Implements ISqlSelectExpression.RenderForGroupBy
        Return RenderForOrderBy()
    End Function

    Public Function RenderForOrderBy() As String Implements ISqlSelectExpression.RenderForOrderBy
        Dim sb As New StringBuilder()
        sb.Append("(")
        sb.Append(String.Format(MyBase.ToString(), _dependencies.ToArray()))
        sb.Append(")")
        Return sb.ToString()
    End Function

    Public Function RenderForFilter() As String Implements ISqlSelectExpression.RenderForFilter
        Dim sb As New StringBuilder()
        sb.Append("(")
        sb.Append(String.Format(MyBase.ToString(), _dependencies.ToArray()))
        sb.Append(")")
        Return sb.ToString()
    End Function

    Public Function RenderForFunction() As String Implements ISqlSelectExpression.RenderForFunction
        Return String.Format(MyBase.ToString(), _dependencies.ToArray())
    End Function

    Public Function RenderForJoin() As String Implements ISqlSelectExpression.RenderForJoin
        Throw New BTSqlException("Text expression cannot be rendered for join ON clause.")
    End Function

    Public Overridable Function RenderForAssignment() As String Implements ISqlSelectExpression.RenderForAssignment
        Return RenderForOrderBy()
    End Function

#End Region

    Public Function GetParameters() As List(Of IBTSqlParameter) Implements ISqlSelectExpression.GetParameters
        Dim result As New List(Of IBTSqlParameter)
        For Each dependency As ISqlSelectExpression In _dependencies
            Dim param As BTSqlParameter = TryCast(dependency, BTSqlParameter)
            If param IsNot Nothing Then
                result.Add(param)
            End If
        Next
        Return result
    End Function

    Public ReadOnly Property Parameters As IEnumerable(Of SqlParameter) Implements IHasParameters.Parameters
        Get
            Dim result As New List(Of SqlParameter)
            For Each dependency As ISqlSelectExpression In _dependencies
                If dependency.Parameters.Count > 0 Then
                    result.AddRange(dependency.Parameters)
                End If
            Next
            Return result
        End Get
    End Property

    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlSelectExpression.GetDependencyIdentifiers
        Dim result As New List(Of String)
        For Each dependency As ISqlSelectExpression In _dependencies
            BTSqlUtility.AddDependencyIdentifiers(result, dependency)
        Next
        Return result
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlSelectExpression.GetDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)
        For Each dependency As ISqlSelectExpression In _dependencies
            BTSqlUtility.AddDependenciesByIdentifier(result, dependency, owner)
        Next
        Return result
    End Function

    Public Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String)) Implements ISqlSelectExpression.UpdateOwnerForPaging
        If _dependencies IsNot Nothing AndAlso _dependencies.Count > 0 Then
            For Each d As ISqlSelectExpression In _dependencies
                d.UpdateOwnerForPaging(pagingAlias, ownersToChange)
            Next
        End If
    End Sub

    Public Function Clone() As Object Implements ISqlSelectExpression.Clone
        'NOTE: this performs a shallow copy
        Return New BTSqlTextSelectExpression(Me.Text, Me.Alias, _dependencies.ToArray())
    End Function

End Class

<Serializable>
Public Class BTSqlTextFilterExpression
    Inherits BTSqlTextExpression
    Implements ISqlFilterExpression

    Public Sub New(ByVal text As String, ByVal booleanOperator As BooleanOperatorTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        MyBase.New(text)
        Me.BooleanOperator = booleanOperator
        Me.IsFirstFilter = False
        _dependencies = New List(Of ISqlSelectExpression)()
        For Each dep As ISqlSelectExpression In dependencies
            _dependencies.Add(DirectCast(dep.Clone(), ISqlSelectExpression))
        Next
    End Sub

    Public Property IsFirstFilter As Boolean Implements ISqlFilterExpression.IsFirstFilter

    Public Property BooleanOperator As BooleanOperatorTypes Implements ISqlFilterExpression.BooleanOperator

    Private _dependencies As List(Of ISqlSelectExpression)

    Public Overrides Function ToString() As String
        Dim sb As New StringBuilder()
        If Not IsFirstFilter Then
            sb.AppendFormat("  {0} ", BooleanOperator.GetDescription())
        End If
        sb.AppendFormat(MyBase.ToString(), _dependencies.ToArray())
        Return sb.ToString()
    End Function

    Public Function Render() As String Implements ISqlFilterExpression.Render
        Return ToString()
    End Function

    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlFilterExpression.GetDependencyIdentifiers
        Dim result As New List(Of String)
        For Each dependency As ISqlSelectExpression In _dependencies
            BTSqlUtility.AddDependencyIdentifiers(result, dependency)
        Next
        Return result
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlFilterExpression.GetDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)
        For Each dependency As ISqlSelectExpression In _dependencies
            BTSqlUtility.AddDependenciesByIdentifier(result, dependency, owner)
        Next
        Return result
    End Function

    Public Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String)) Implements ISqlFilterExpression.UpdateOwnerForPaging
        If _dependencies IsNot Nothing AndAlso _dependencies.Count > 0 Then
            For Each d As ISqlSelectExpression In _dependencies
                d.UpdateOwnerForPaging(pagingAlias, ownersToChange)
            Next
        End If
    End Sub

End Class

<Serializable>
Public Class BTSqlTextOrderByExpression
    Inherits BTSqlTextExpression
    Implements ISqlOrderByExpression

    Public Sub New(ByVal text As String, ByVal direction As DirectionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        MyBase.New(text)
        Me.Direction = direction
        _dependencies = New List(Of ISqlSelectExpression)()
        For Each dep As ISqlSelectExpression In dependencies
            _dependencies.Add(DirectCast(dep.Clone(), ISqlSelectExpression))
        Next
    End Sub

    Public Property Direction As DirectionTypes Implements ISqlOrderByExpression.Direction

    Private _dependencies As List(Of ISqlSelectExpression)

    Public Overrides Function ToString() As String
        Dim sb As New StringBuilder()
        sb.Append("(")
        sb.AppendFormat(MyBase.ToString(), _dependencies.ToArray())
        sb.Append(")")
        sb.AppendFormat(" {0}", Direction.GetDescription())
        Return sb.ToString()
    End Function

    Public Function Render() As String Implements ISqlExpression.Render
        Return ToString()
    End Function

    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlOrderByExpression.GetDependencyIdentifiers
        Dim result As New List(Of String)
        For Each dependency As ISqlSelectExpression In _dependencies
            BTSqlUtility.AddDependencyIdentifiers(result, dependency)
        Next
        Return result
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlOrderByExpression.GetDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)
        For Each dependency As ISqlSelectExpression In _dependencies
            BTSqlUtility.AddDependenciesByIdentifier(result, dependency, owner)
        Next
        Return result
    End Function

End Class

<Diagnostics.DebuggerDisplay("Alias={Alias}")> _
<Serializable>
Public Class BTSqlTextJoinExpression
    Inherits BTSqlTextExpression
    Implements ISqlJoinExpression

    Public Sub New(ByVal joinType As JoinTypes, ByVal queryText As String, ByVal [alias] As String, ByVal column As ISqlSelectExpression, ByVal otherColumn As ISqlSelectExpression, ByVal ParamArray filters() As ISqlFilterExpression)
        MyBase.New(queryText)

        If joinType = JoinTypes.OuterApply OrElse joinType = JoinTypes.CrossApply Then
            Throw New BTSqlException(String.Format("{0} and {1} are not allowed with BTSqlTextJoinExpression.  Please use BTSqlTextApplyExpression instead.", JoinTypes.OuterApply.GetDescription(), JoinTypes.CrossApply.GetDescription()))
        End If

        Me.JoinType = joinType
        _alias = [alias]
        _column = column.WithAlias(_alias)
        _otherColumn = DirectCast(otherColumn.Clone(), ISqlSelectExpression)
        _filters = New List(Of ISqlFilterExpression)
        _filters.AddRange(filters)
    End Sub

    Private _column As ISqlSelectExpression
    Private _otherColumn As ISqlSelectExpression
    Private _filters As List(Of ISqlFilterExpression)

    Private _alias As String
    Public ReadOnly Property [Alias] As String Implements ISqlJoinExpression.Alias
        Get
            Return _alias
        End Get
    End Property

    Public Sub ChangeAlias(ByVal [alias] As String) Implements ISqlJoinable.ChangeAlias
        _alias = [alias]
        _column.Alias = [alias]
    End Sub

    Public Property JoinType As JoinTypes Implements ISqlJoinExpression.JoinType

    ''' <summary>
    ''' When false, if the join is not needed by any dependencies (e.g. select statements, filters, etc), it will be removed from the query
    ''' </summary>
    ''' <value>Default is false.</value>
    Public Property ShouldNotRemove As Boolean Implements ISqlJoinExpression.ShouldNotRemove

    Public Overrides Function ToString() As String
        Dim useAlias As Boolean = Not String.IsNullOrWhitespace(Me.Alias)
        Dim sb As New StringBuilder()
        sb.AppendFormat("{0} ", JoinType.GetDescription())
        If useAlias Then
            sb.Append("(")
        End If
        sb.Append(MyBase.ToString())
        If useAlias Then
            sb.AppendFormat(") {0}", Me.Alias)
        End If
        sb.AppendFormat(" ON {0} = {1}", _column, _otherColumn)
        For Each filter As ISqlFilterExpression In _filters
            sb.AppendFormat(" {0}", filter.Render())
        Next
        Return sb.ToString()
    End Function

    Public Function Render() As String Implements ISqlJoinExpression.Render
        Return ToString()
    End Function

    Public Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String)) Implements ISqlJoinExpression.UpdateOwnerForPaging
        _otherColumn.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        If _filters IsNot Nothing AndAlso _filters.Count > 0 Then
            For Each f As ISqlFilterExpression In _filters
                f.UpdateOwnerForPaging(pagingAlias, ownersToChange)
            Next
        End If
    End Sub

    Public ReadOnly Property OtherColumn As ISqlSelectExpression Implements ISqlJoinExpression.OtherColumn
        Get
            Return _otherColumn
        End Get
    End Property

    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlJoinExpression.GetDependencyIdentifiers
        Dim result As New List(Of String)

        BTSqlUtility.AddDependencyIdentifiers(result, _otherColumn)

        If _filters IsNot Nothing Then
            For Each f As ISqlFilterExpression In _filters
                BTSqlUtility.AddDependencyIdentifiers(result, f)
            Next
        End If

        Return result
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlJoinExpression.GetDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)

        BTSqlUtility.AddDependenciesByIdentifier(result, _otherColumn, owner)

        If _filters IsNot Nothing Then
            For Each f As ISqlFilterExpression In _filters
                BTSqlUtility.AddDependenciesByIdentifier(result, f, owner)
            Next
        End If

        Return result
    End Function

End Class

<Serializable>
Public Class BTSqlTextInsertValueExpression
    Inherits BTSqlTextExpression
    Implements ISqlSelectExpression

    Public Sub New(ByVal text As String)
        MyBase.New(text)

    End Sub

    Public Property [Alias]() As String Implements ISqlSelectExpression.[Alias]
        Get
            Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
        End Get
        Set(value As String)
            Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
        End Set
    End Property

    Public Function WithAlias(ByVal [alias] As String) As ISqlSelectExpression Implements ISqlSelectExpression.WithAlias
        Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
    End Function

    Public Function GetDataRowColumnName() As String Implements ISqlSelectExpression.GetDataRowColumnName
        Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
    End Function

#Region "ToString / Render"

    Public Overrides Function ToString() As String
        Return Text
    End Function

    Public Function Render() As String Implements ISqlSelectExpression.Render
        Return ToString()
    End Function

    Public Function RenderForGroupBy() As String Implements ISqlSelectExpression.RenderForGroupBy
        Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
    End Function

    Public Function RenderForOrderBy() As String Implements ISqlSelectExpression.RenderForOrderBy
        Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
    End Function

    Public Function RenderForFilter() As String Implements ISqlSelectExpression.RenderForFilter
        Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
    End Function

    Public Function RenderForFunction() As String Implements ISqlSelectExpression.RenderForFunction
        Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
    End Function

    Public Function RenderForJoin() As String Implements ISqlSelectExpression.RenderForJoin
        Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
    End Function

    Public Overridable Function RenderForAssignment() As String Implements ISqlSelectExpression.RenderForAssignment
        Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
    End Function

#End Region

    Public Function GetParameters() As List(Of IBTSqlParameter) Implements ISqlSelectExpression.GetParameters
        Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
    End Function

    Public ReadOnly Property Parameters As IEnumerable(Of SqlParameter) Implements IHasParameters.Parameters
        Get
            Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
        End Get
    End Property


    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlSelectExpression.GetDependencyIdentifiers
        Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlSelectExpression.GetDependenciesByIdentifier
        Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
    End Function

    Public Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String)) Implements ISqlSelectExpression.UpdateOwnerForPaging
        Throw New NotImplementedException("This operation is not supported on BTSqlTextInsertValueExpression")
    End Sub

    Public Function Clone() As Object Implements ISqlSelectExpression.Clone
        'NOTE: this performs a shallow copy
        Return New BTSqlTextInsertValueExpression(Me.Text)
    End Function

End Class
