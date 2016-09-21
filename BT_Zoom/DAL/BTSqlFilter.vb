Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Text
Imports System.Linq
Imports BT_Zoom.BTSql
Imports BT_Zoom.Enums.BTSql
Imports BT_Zoom.Interfaces

<Serializable>
Public MustInherit Class BTSqlFilter
    Implements ISqlFilterExpression, IBTSqlFilter

    Public Property IsFirstFilter As Boolean Implements ISqlFilterExpression.IsFirstFilter
    Public Property BooleanOperator As BooleanOperatorTypes Implements ISqlFilterExpression.BooleanOperator

    Public Sub New(ByVal booleanOperator As BooleanOperatorTypes)
        Me.BooleanOperator = booleanOperator
        Me.IsFirstFilter = False
    End Sub

    Protected MustOverride Function RenderInternal() As String

    Public Function Render() As String Implements ISqlFilterExpression.Render
        Return ToString()
    End Function

    Public Overrides Function ToString() As String
        Dim sb As New StringBuilder()
        sb.AppendFormat(" {0} ", If(IsFirstFilter, BooleanOperator.GetDescription().Replace(BooleanOperatorTypes.AND.GetDescription, "").Replace(BooleanOperatorTypes.OR.GetDescription, ""), BooleanOperator.GetDescription()))
        sb.Append(RenderInternal())
        Return sb.ToString()
    End Function

    Public MustOverride Function GetDependencyIdentifiers() As List(Of String) Implements ISqlFilterExpression.GetDependencyIdentifiers

    Public MustOverride Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlFilterExpression.GetDependenciesByIdentifier

    Public MustOverride Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String)) Implements ISqlFilterExpression.UpdateOwnerForPaging

End Class

<Serializable>
Public Class BTSqlComparisonFilter
    Inherits BTSqlFilter
    Implements IBTSqlComparisonFilter

    Public Property LeftExpression As ISqlSelectExpression Implements IBTSqlComparisonFilter.LeftExpression
    Public Property RightExpression As ISqlSelectExpression Implements IBTSqlComparisonFilter.RightExpression
    Public Property RightClause As String Implements IBTSqlComparisonFilter.RightClause
    Public Property ComparisonOperator As ComparisonOperatorTypes Implements IBTSqlComparisonFilter.ComparisonOperator

    Public Sub New(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightClause As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        MyBase.New(booleanOperator)

        Me.LeftExpression = DirectCast(leftExpression.Clone(), ISqlSelectExpression)
        Me.RightExpression = Nothing
        Me.ComparisonOperator = comparisonOperator
        Me.RightClause = rightClause
    End Sub

    Public Sub New(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        MyBase.New(booleanOperator)

        Me.LeftExpression = DirectCast(leftExpression.Clone(), ISqlSelectExpression)
        Me.RightExpression = DirectCast(rightExpression.Clone(), ISqlSelectExpression)
        Me.ComparisonOperator = comparisonOperator
        Me.RightClause = Nothing
    End Sub

    Public Sub New(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        MyBase.New(booleanOperator)

        Me.LeftExpression = DirectCast(leftExpression.Clone(), ISqlSelectExpression)
        Me.RightExpression = New BTSqlTextSelectExpression(rightExpression.Parameter.ParameterName, Nothing)
        Me.ComparisonOperator = comparisonOperator
        Me.RightClause = Nothing
    End Sub

    Protected Overrides Function RenderInternal() As String
        Dim sb As New StringBuilder()
        If LeftExpression IsNot Nothing Then
            sb.Append(LeftExpression.RenderForFilter())
        End If
        sb.AppendFormat(" {0} ", ComparisonOperator.GetDescription())
        If RightExpression IsNot Nothing Then
            sb.Append(RightExpression.RenderForFilter())
        Else
            sb.Append(RightClause)
        End If
        Return sb.ToString()
    End Function

    Public Overrides Function GetDependencyIdentifiers() As List(Of String)
        Dim result As New List(Of String)

        BTSqlUtility.AddDependencyIdentifiers(result, LeftExpression)
        BTSqlUtility.AddDependencyIdentifiers(result, RightExpression)

        'NOTE: assumption is that rightClause will never contain dependencies.  It is intended to be used for simple values like "5", which should be parameterized, but Jeff wants this capability for now.

        Return result
    End Function

    Public Overrides Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression)
        Dim result As New List(Of ISqlSelectExpression)

        BTSqlUtility.AddDependenciesByIdentifier(result, LeftExpression, owner)
        BTSqlUtility.AddDependenciesByIdentifier(result, RightExpression, owner)

        'NOTE: assumption is that rightClause will never contain dependencies.  It is intended to be used for simple values like "5", which should be parameterized, but Jeff wants this capability for now.

        Return result
    End Function

    Public Overrides Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String))
        If LeftExpression IsNot Nothing Then
            LeftExpression.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        End If
        If RightExpression IsNot Nothing Then
            RightExpression.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        End If
    End Sub

End Class

<Serializable>
Public Class BTSqlLogicalFilter
    Inherits BTSqlFilter
    Implements IBTSqlLogicalFilter

    Public Property LeftExpression As ISqlSelectExpression Implements IBTSqlLogicalFilter.LeftExpression
    Public Property RightClause As String Implements IBTSqlLogicalFilter.RightClause
    Public Property RightExpression As ISqlSelectExpression Implements IBTSqlLogicalFilter.RightExpression
    Public Property LogicalOperator As LogicalOperatorTypes Implements IBTSqlLogicalFilter.LogicalOperator

    Public Sub New(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        MyBase.New(booleanOperator)

        If logicalOperator <> LogicalOperatorTypes.IsNull AndAlso logicalOperator <> LogicalOperatorTypes.IsNotNull Then
            Throw New ArgumentException(String.Format("Please supply the correct arguments for {0}", logicalOperator.GetDescription()), "logicalOperator")
        End If

        Me.LeftExpression = DirectCast(leftExpression.Clone(), ISqlSelectExpression)
        Me.LogicalOperator = logicalOperator
        Me.RightClause = String.Empty

    End Sub

    Public Sub New(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal logicalExpression As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        MyBase.New(booleanOperator)

        Me.LeftExpression = DirectCast(leftExpression.Clone(), ISqlSelectExpression)
        Me.LogicalOperator = logicalOperator

        SetRightClause(logicalExpression)
    End Sub

    Public Sub New(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND, Optional ByVal useLeadingParen As Boolean = False, Optional ByVal useTrailingParen As Boolean = False)
        MyBase.New(booleanOperator)

        Me.LeftExpression = DirectCast(leftExpression.Clone(), ISqlSelectExpression)
        Me.LogicalOperator = logicalOperator

        If rightExpression.Parameter.SqlDbType = Data.SqlDbType.Structured AndAlso Not String.IsNullOrEmpty(rightExpression.Parameter.TypeName) Then
            If logicalOperator = LogicalOperatorTypes.In OrElse logicalOperator = LogicalOperatorTypes.NotIn Then
                'table parameter
                Me.RightExpression = BTSqlUtility.CreateSqlSelectBuilderFrom(rightExpression.Parameter).ToSelectQuery(True)
            Else
                Throw New ArgumentException(String.Format("Please supply the correct arguments for {0} with a table parameter", logicalOperator.GetDescription()), "logicalOperator")
            End If
        Else
            SetRightClause(rightExpression.Parameter.ParameterName)
        End If
    End Sub

    Public Sub New(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        MyBase.New(booleanOperator)

        If TypeOf rightExpression Is ISqlSelectQuery Then
            Select Case logicalOperator
                Case LogicalOperatorTypes.In, LogicalOperatorTypes.NotIn, LogicalOperatorTypes.Exists, LogicalOperatorTypes.NotExists
                    Me.RightExpression = DirectCast(rightExpression.Clone(), ISqlSelectExpression)
                Case Else
                    Throw New ArgumentException(String.Format("Please supply the correct arguments for {0}", logicalOperator.GetDescription()), "logicalOperator")
            End Select
        Else
            Me.RightExpression = DirectCast(rightExpression.Clone(), ISqlSelectExpression)
        End If

        Me.LeftExpression = DirectCast(leftExpression.Clone(), ISqlSelectExpression)
        Me.LogicalOperator = logicalOperator
    End Sub

    Public Sub New(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal leftBetween As IBTSqlParameter, ByVal rightBetween As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        MyBase.New(booleanOperator)

        If logicalOperator <> LogicalOperatorTypes.Between AndAlso logicalOperator <> LogicalOperatorTypes.NotBetween Then
            Throw New ArgumentException(String.Format("Please supply the correct arguments for {0}", logicalOperator.GetDescription()), "logicalOperator")
        End If

        Me.LeftExpression = DirectCast(leftExpression.Clone(), ISqlSelectExpression)
        Me.LogicalOperator = logicalOperator
        Me.RightClause = String.Format("{0} AND {1}", leftBetween.Parameter.ParameterName, rightBetween.Parameter.ParameterName)

    End Sub

    Public Sub New(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal leftBetween As ISqlSelectExpression, ByVal rightBetween As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        MyBase.New(booleanOperator)

        If logicalOperator <> LogicalOperatorTypes.Between AndAlso logicalOperator <> LogicalOperatorTypes.NotBetween Then
            Throw New ArgumentException(String.Format("Please supply the correct arguments for {0}", logicalOperator.GetDescription()), "logicalOperator")
        End If

        Me.LeftExpression = DirectCast(leftExpression.Clone(), ISqlSelectExpression)
        Me.LogicalOperator = logicalOperator
        Me.RightClause = String.Format("{0} AND {1}", leftBetween.RenderForFilter(), rightBetween.RenderForFilter())
    End Sub

    Public Sub New(ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectQuery, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        MyBase.New(booleanOperator)

        If logicalOperator <> LogicalOperatorTypes.Exists AndAlso logicalOperator <> LogicalOperatorTypes.NotExists Then
            Throw New ArgumentException(String.Format("Please supply the correct arguments for {0}", logicalOperator.GetDescription()), "logicalOperator")
        End If

        Me.RightExpression = DirectCast(rightExpression.Clone(), ISqlSelectExpression)
        Me.LogicalOperator = logicalOperator
    End Sub

    Private Sub SetRightClause(ByVal logicalExpression As String)
        Select Case LogicalOperator
            Case LogicalOperatorTypes.In, LogicalOperatorTypes.NotIn, LogicalOperatorTypes.Exists, LogicalOperatorTypes.NotExists
                Me.RightClause = String.Format("({0})", logicalExpression)
            Case LogicalOperatorTypes.Like, LogicalOperatorTypes.NotLike
                Me.RightClause = String.Format("'{0}'", logicalExpression)
            Case LogicalOperatorTypes.Contains
                Me.RightClause = logicalExpression
            Case Else
                Throw New ArgumentException(String.Format("Please supply the correct arguments for {0}", LogicalOperator.GetDescription()), "logicalOperator")
        End Select
    End Sub

    Protected Overrides Function RenderInternal() As String
        Dim sb As New StringBuilder()
        If LeftExpression IsNot Nothing AndAlso LogicalOperator <> LogicalOperatorTypes.Contains Then 'could be nothing for Exists clauses
            sb.Append(LeftExpression.RenderForFilter())
        End If
        sb.AppendFormat(" {0} ", LogicalOperator.GetDescription())
        Dim rightQuery As ISqlSelectQuery = TryCast(RightExpression, ISqlSelectQuery)
        If rightQuery IsNot Nothing Then
            If Not rightQuery.IsTableParameterTable Then
                sb.Append("(")
            End If
            sb.AppendFormat("{0}", rightQuery.RenderForFilter())
            If Not rightQuery.IsTableParameterTable Then
                sb.Append(")")
            End If
        ElseIf RightExpression IsNot Nothing Then
            sb.AppendFormat(RightExpression.RenderForFilter)
        Else
            If LogicalOperator = LogicalOperatorTypes.Contains Then
                sb.AppendFormat("({0}, {1})", LeftExpression.RenderForFilter(), RightClause)
            Else
                sb.Append(RightClause)
            End If
        End If
        Return sb.ToString()
    End Function

    Public Overrides Function GetDependencyIdentifiers() As List(Of String)
        Dim result As New List(Of String)

        BTSqlUtility.AddDependencyIdentifiers(result, LeftExpression)
        BTSqlUtility.AddDependencyIdentifiers(result, RightExpression)

        If _rightDependencies.Count > 0 Then
            For Each dependency As ISqlSelectExpression In _rightDependencies
                BTSqlUtility.AddDependencyIdentifiers(result, dependency)
            Next
        End If

        Return result
    End Function

    Public Overrides Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression)
        Dim result As New List(Of ISqlSelectExpression)

        BTSqlUtility.AddDependenciesByIdentifier(result, LeftExpression, owner)
        BTSqlUtility.AddDependenciesByIdentifier(result, RightExpression, owner)

        If _rightDependencies.Count > 0 Then
            For Each dependency As ISqlSelectExpression In _rightDependencies
                'NOTE: not sure about this...
                BTSqlUtility.AddDependenciesByIdentifier(result, dependency, owner)
            Next
        End If

        Return result
    End Function

    Private _rightDependencies As New List(Of ISqlSelectExpression)

    Public Sub AddDependencies(ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTSqlLogicalFilter.AddDependencies 'NOTE: this is a hack in case the logical expression for in,exists,like contain dependencies for joins
        If dependencies IsNot Nothing AndAlso dependencies.Length > 0 Then
            For Each dep As ISqlSelectExpression In dependencies
                _rightDependencies.Add(DirectCast(dep.Clone(), ISqlSelectExpression))
            Next
        End If
    End Sub

    Public Overrides Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String))
        If LeftExpression IsNot Nothing Then
            LeftExpression.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        End If
        If RightExpression IsNot Nothing Then
            RightExpression.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        End If
    End Sub

End Class

''' <summary>
''' Groups a list of filters with opening and closing parentheses
''' </summary>
<Serializable>
Public Class BTSqlFilterGroup
    Implements ISqlFilterExpression, IBTSqlFilterGroup

    Private ReadOnly _filters As New List(Of ISqlFilterExpression)

    Public ReadOnly Property Count As Integer Implements IBTSqlFilterGroup.Count
        Get
            Return _filters.Count
        End Get
    End Property

    Public Sub New()
        Me.New(BooleanOperatorTypes.AND)
    End Sub

    Public Sub New(ByVal booleanOperator As BooleanOperatorTypes, ParamArray filters As ISqlFilterExpression())
        Me.BooleanOperator = booleanOperator
        _filters.AddRange(filters)
    End Sub

#Region "Interface"

    Public Sub UpdateOwnerForPaging(pagingAlias As String, ownersToChange As List(Of String)) Implements ICanUpdateOwnerForPaging.UpdateOwnerForPaging
        For Each f As ISqlFilterExpression In _filters
            f.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        Next
    End Sub

    Public Function GetDependenciesByIdentifier(owner As String) As List(Of ISqlSelectExpression) Implements IHasDependencies.GetDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)
        For Each f As ISqlFilterExpression In _filters
            BTSqlUtility.AddDependenciesByIdentifier(result, f, owner)
        Next
        Return result
    End Function

    Public Function GetDependencyIdentifiers() As List(Of String) Implements IHasDependencies.GetDependencyIdentifiers
        Dim result As New List(Of String)
        For Each f As ISqlFilterExpression In _filters
            BTSqlUtility.AddDependencyIdentifiers(result, f)
        Next
        Return result
    End Function

    Public Function Render() As String Implements ISqlExpression.Render
        If Not _filters.Any Then
            Throw New BTSqlException("You must have at least one filter in the filter group.")
        End If
        Dim sb As New StringBuilder()
        If IsFirstFilter Then
            If BooleanOperator = BooleanOperatorTypes.AndNot OrElse BooleanOperator = BooleanOperatorTypes.OrNot Then
                sb.Append(" NOT ")
            End If
        Else
            sb.AppendFormat(" {0} ", BooleanOperator.GetDescription())
        End If
        sb.Append("(")
        For i As Integer = 0 To _filters.Count - 1
            Dim filter As ISqlFilterExpression = _filters(i)
            filter.IsFirstFilter = (i = 0)
            sb.AppendLine(filter.Render())
        Next
        sb.Append(")")
        Return sb.ToString()
    End Function

    Public Property BooleanOperator As BooleanOperatorTypes Implements ISqlFilterExpression.BooleanOperator

    Public Property IsFirstFilter As Boolean Implements ISqlFilterExpression.IsFirstFilter

#End Region

#Region "Add Filters"

    Public Sub AddFilter(ByVal filter As ISqlFilterExpression) Implements IBTSqlFilterGroup.AddFilter
        _filters.Add(filter)
    End Sub

    Public Function AddFilterGroup(ByVal booleanOperator As BooleanOperatorTypes, ByVal ParamArray filters As ISqlFilterExpression()) As IBTSqlFilterGroup Implements IBTSqlFilterGroup.AddFilterGroup
        Dim fg As New BTSqlFilterGroup(booleanOperator, filters)
        _filters.Add(fg)
        Return fg
    End Function

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightClause As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTSqlFilterGroup.AddFilter
        _filters.Add(New BTSqlComparisonFilter(leftExpression, comparisonOperator, rightClause, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTSqlFilterGroup.AddFilter
        _filters.Add(New BTSqlComparisonFilter(leftExpression, comparisonOperator, rightExpression, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOpertator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTSqlFilterGroup.AddFilter
        _filters.Add(New BTSqlLogicalFilter(leftExpression, logicalOpertator, rightExpression, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectQuery, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTSqlFilterGroup.AddFilter
        _filters.Add(New BTSqlLogicalFilter(logicalOperator, rightExpression, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTSqlFilterGroup.AddFilter
        _filters.Add(New BTSqlLogicalFilter(leftExpression, logicalOperator, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal logicalExpression As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTSqlFilterGroup.AddFilter
        _filters.Add(New BTSqlLogicalFilter(leftExpression, logicalOperator, logicalExpression, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal leftBetween As IBTSqlParameter, ByVal rightBetween As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTSqlFilterGroup.AddFilter
        _filters.Add(New BTSqlLogicalFilter(leftExpression, logicalOperator, leftBetween, rightBetween, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal leftBetween As ISqlSelectExpression, ByVal rightBetween As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTSqlFilterGroup.AddFilter
        _filters.Add(New BTSqlLogicalFilter(leftExpression, logicalOperator, leftBetween, rightBetween, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTSqlFilterGroup.AddFilter
        _filters.Add(New BTSqlComparisonFilter(leftExpression, comparisonOperator, rightExpression, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTSqlFilterGroup.AddFilter
        _filters.Add(New BTSqlLogicalFilter(leftExpression, logicalOperator, rightExpression, booleanOperator))
    End Sub

#End Region

End Class
