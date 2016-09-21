Option Strict On
Option Explicit On

Imports System.Text
Imports System.Collections.Generic
Imports System.Data
Imports System.Linq
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Collections
Imports BT_Zoom.Delegates
Imports BT_Zoom.Interfaces

<Serializable>
Public MustInherit Class BTSqlTransformBase
    Implements ISqlSelectExpression, IBTSqlTransformBase

    Public Sub New(ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal type As Type, Optional ByVal sortExpressions As IBTSqlOrderByGroup = Nothing)
        Me.Alias = [alias]
        _sortExpressionGroup = sortExpressions
        _expressionList = New List(Of ISqlSelectExpression)
        For Each expr As ISqlSelectExpression In expressionList
            _expressionList.Add(DirectCast(expr.Clone(), ISqlSelectExpression))
        Next
        _type = type
    End Sub

    Public MustOverride Function ApplyTransform(ByVal dr As DataRow, ByVal instance As IBTList) As String Implements IBTSqlTransformBase.ApplyTransform
    Public MustOverride Function ApplyTransform(ByVal dr As DataRow) As String Implements IBTSqlTransformBase.ApplyTransform

    Private _expressionList As List(Of ISqlSelectExpression)
    Private _sortExpressionGroup As IBTSqlOrderByGroup
    Private _type As Type

    Public ReadOnly Property ExpressionList As List(Of ISqlSelectExpression) Implements IBTSqlTransformBase.ExpressionList
        Get
            Return _expressionList
        End Get
    End Property

    Public ReadOnly Property Type As Type Implements IBTSqlTransformBase.Type
        Get
            Return _type
        End Get
    End Property

    Public ReadOnly Property SortExpressionGroup As IBTSqlOrderByGroup Implements IBTSqlTransformBase.SortExpressionGroup
        Get
            Return _sortExpressionGroup
        End Get
    End Property

    Public Property [Alias] As String Implements ISqlSelectExpression.Alias

    Public Function WithAlias(ByVal [alias] As String) As ISqlSelectExpression Implements ISqlSelectExpression.WithAlias
        Dim result As BTSqlTransformBase = CType(Clone(), BTSqlTransformBase)
        result.Alias = [alias]
        Return result
    End Function

    Public Function GetDataRowColumnName() As String Implements ISqlSelectExpression.GetDataRowColumnName
        If Not String.IsNullOrWhitespace(Me.Alias) Then
            Return Me.Alias.Replace("[", String.Empty).Replace("]", String.Empty)
        Else
            Throw New BTSqlException("An alias is required")
        End If
    End Function

#Region "ToString / Render"

    Public Overrides Function ToString() As String Implements IBTSqlTransformBase.ToString
        Throw New BTSqlException("BTSqlTransformBase does not render itself")
    End Function

    Public Overridable Function Render() As String Implements ISqlSelectExpression.Render
        Return ToString()
    End Function

    Public Overridable Function RenderForGroupBy() As String Implements ISqlSelectExpression.RenderForGroupBy
        Return ToString()
    End Function

    Public Overridable Function RenderForOrderBy() As String Implements ISqlSelectExpression.RenderForOrderBy
        If _sortExpressionGroup Is Nothing Then
            Return ToString()
        Else
            Return _sortExpressionGroup.Render
        End If
    End Function

    Public Overridable Function RenderForFilter() As String Implements ISqlSelectExpression.RenderForFilter
        Return ToString()
    End Function

    Public Overridable Function RenderForFunction() As String Implements ISqlSelectExpression.RenderForFunction
        Return ToString()
    End Function

    Public Overridable Function RenderForJoin() As String Implements ISqlSelectExpression.RenderForJoin
        Throw New BTSqlException("Transform expression cannot be rendered for join ON clause.")
    End Function

    Public Overridable Function RenderForAssignment() As String Implements ISqlSelectExpression.RenderForAssignment
        Throw New BTSqlException("Transform expression cannot be rendered for assignment.")
    End Function

#End Region

    Public Function GetParameters() As List(Of IBTSqlParameter) Implements ISqlSelectExpression.GetParameters
        Dim result As New List(Of IBTSqlParameter)
        For Each dependency As ISqlSelectExpression In ExpressionList
            result.AddRange(dependency.GetParameters)
        Next
        Return result
    End Function

    Public ReadOnly Property Parameters As IEnumerable(Of SqlParameter) Implements IHasParameters.Parameters
        Get
            Dim result As New List(Of SqlParameter)
            For Each dependency As ISqlSelectExpression In ExpressionList
                If dependency.Parameters.Count > 0 Then
                    result.AddRange(dependency.Parameters)
                End If
            Next
            Return result
        End Get
    End Property

    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlSelectExpression.GetDependencyIdentifiers
        Dim result As New List(Of String)
        For Each dependency As ISqlSelectExpression In ExpressionList
            BTSqlUtility.AddDependencyIdentifiers(result, dependency)
        Next
        Return result
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlSelectExpression.GetDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)
        For Each dependency As ISqlSelectExpression In ExpressionList
            BTSqlUtility.AddDependenciesByIdentifier(result, dependency, owner)
        Next
        Return result
    End Function

    Public Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String)) Implements ISqlSelectExpression.UpdateOwnerForPaging
        If _expressionList IsNot Nothing AndAlso _expressionList.Count > 0 Then
            For Each s As ISqlSelectExpression In _expressionList
                s.UpdateOwnerForPaging(pagingAlias, ownersToChange)
            Next
        End If
    End Sub

    Public MustOverride Function Clone() As Object Implements ISqlSelectExpression.Clone

End Class

Public Class BTSqlTransform
    Inherits BTSqlTransformBase
    Implements IBTSqlTransform

    Public Sub New(ByVal transform As ApplySqlTransformationDelegate, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), Optional ByVal sortExpressions As IBTSqlOrderByGroup = Nothing)
        Me.New(transform, [alias], expressionList, GetType(String), sortExpressions)
    End Sub

    Public Sub New(ByVal transform As ApplySqlTransformationDelegate, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal type As Type, Optional ByVal sortExpressions As IBTSqlOrderByGroup = Nothing)
        MyBase.New([alias], expressionList, type, sortExpressions)
        _transform = transform
    End Sub

    Public Overrides Function ApplyTransform(ByVal dr As DataRow, ByVal instance As IBTList) As String
        Return ApplyTransform(dr)
    End Function

    Public Overrides Function ApplyTransform(dr As DataRow) As String
        Return _transform(dr, ExpressionList)
    End Function

    Private _transform As ApplySqlTransformationDelegate

    Public Overrides Function Clone() As Object
        'NOTE: this performs a shallow copy
        Return New BTSqlTransform(_transform, Me.Alias, ExpressionList, Type, SortExpressionGroup)
    End Function
End Class

Public Class BTSqlRelatedDataTransform
    Inherits BTSqlTransformBase
    Implements IBTSqlRelatedDataTransform

    Public Sub New(ByVal transform As ApplySqlTransformationRelatedDataDelegate, ByVal loadData As LoadRelatedDataForTransform, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), Optional ByVal sortExpressions As IBTSqlOrderByGroup = Nothing)
        Me.New(transform, loadData, [alias], expressionList, GetType(String), sortExpressions)
    End Sub

    Public Sub New(ByVal transform As ApplySqlTransformationRelatedDataDelegate, ByVal loadData As LoadRelatedDataForTransform, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal type As Type, Optional ByVal sortExpressions As IBTSqlOrderByGroup = Nothing)
        MyBase.New([alias], expressionList, type, sortExpressions)
        _transform = transform
        _loadData = loadData
    End Sub

    Public Overrides Function ApplyTransform(ByVal dr As DataRow, ByVal instance As IBTList) As String
        RelatedData = _loadData(instance)
        Return ApplyTransform(dr)
    End Function

    Public Overrides Function ApplyTransform(dr As DataRow) As String
        If _relatedData Is Nothing Then
            Throw New BTSqlException("You must provide the related data before attempting the transform.")
        End If
        Return _transform(dr, ExpressionList, _relatedData)
    End Function

    Private _transform As ApplySqlTransformationRelatedDataDelegate
    Private _loadData As LoadRelatedDataForTransform
    Private _relatedData As Dictionary(Of String, IBTList)

    Public WriteOnly Property RelatedData As Dictionary(Of String, IBTList) Implements IBTSqlRelatedDataTransform.RelatedData
        Set(value As Dictionary(Of String, IBTList))
            _relatedData = value
        End Set
    End Property

    Public Overrides Function Clone() As Object
        Return New BTSqlRelatedDataTransform(_transform, _loadData, [Alias], ExpressionList, Type, SortExpressionGroup)
    End Function
End Class

Public Class BTSqlNewColumnTransform
    Inherits BTSqlTransform
    Implements IBTSqlNewColumnTransform

    ''' <summary>
    ''' Creates a transform for a new column with the default sort being the first column in <paramref name="expressionList"></paramref> unless not <paramref name="useFirstColAsSortExpression"></paramref>
    ''' </summary>
    Public Sub New(ByVal transform As ApplySqlTransformationDelegate, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), Optional ByVal useFirstColAsSortExpression As Boolean = True)
        MyClass.New(transform, [alias], expressionList, If(useFirstColAsSortExpression, New BTSqlOrderByGroup(expressionList(0)), Nothing))
    End Sub

    Public Sub New(ByVal transform As ApplySqlTransformationDelegate, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal sortExpressions As IBTSqlOrderByGroup)
        MyBase.New(transform, [alias], expressionList, sortExpressions)
    End Sub

    Public Sub New(ByVal transform As ApplySqlTransformationDelegate, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal sortExpression As ISqlSelectExpression)
        MyClass.New(transform, [alias], expressionList, New BTSqlOrderByGroup(sortExpression))
    End Sub

    Public Sub New(ByVal transform As ApplySqlTransformationDelegate, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal type As Type, ByVal sortExpressions As IBTSqlOrderByGroup)
        MyBase.New(transform, [alias], expressionList, type, sortExpressions)
    End Sub

    Public Sub New(ByVal transform As ApplySqlTransformationDelegate, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal type As Type, ByVal sortExpression As ISqlSelectExpression)
        MyClass.New(transform, [alias], expressionList, type, New BTSqlOrderByGroup(sortExpression))
    End Sub

    ''' <summary>
    ''' Creates a transform for a new column with the default sort being the first column in <paramref name="expressionList"></paramref>,
    ''' the output column object will be of <paramref name="type"></paramref> unless not <paramref name="useFirstColAsSortExpression"></paramref>
    ''' </summary>
    Public Sub New(ByVal transform As ApplySqlTransformationDelegate, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal type As Type, Optional ByVal useFirstColAsSortExpression As Boolean = True)
        MyClass.New(transform, [alias], expressionList, type, If(useFirstColAsSortExpression, New BTSqlOrderByGroup(expressionList(0)), Nothing))
    End Sub

End Class

Public Class BTSqlNewRelatedDataColumnTransform
    Inherits BTSqlRelatedDataTransform

    ''' <summary>
    ''' Creates a transform for a new column using related data with the default sort being the first column in <paramref name="expressionList"></paramref> unless not <paramref name="useFirstColAsSortExpression"></paramref>
    ''' </summary>
    Public Sub New(ByVal transform As ApplySqlTransformationRelatedDataDelegate, ByVal loadData As LoadRelatedDataForTransform, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), Optional ByVal useFirstColAsSortExpression As Boolean = True)
        MyClass.New(transform, loadData, [alias], expressionList, If(useFirstColAsSortExpression, New BTSqlOrderByGroup(expressionList(0)), Nothing))
    End Sub

    Public Sub New(ByVal transform As ApplySqlTransformationRelatedDataDelegate, ByVal loadData As LoadRelatedDataForTransform, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal sortExpressions As IBTSqlOrderByGroup)
        MyBase.New(transform, loadData, [alias], expressionList, sortExpressions)
    End Sub

    Public Sub New(ByVal transform As ApplySqlTransformationRelatedDataDelegate, ByVal loadData As LoadRelatedDataForTransform, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal sortExpression As ISqlSelectExpression)
        MyClass.New(transform, loadData, [alias], expressionList, New BTSqlOrderByGroup(sortExpression))
    End Sub

    Public Sub New(ByVal transform As ApplySqlTransformationRelatedDataDelegate, ByVal loadData As LoadRelatedDataForTransform, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal type As Type, ByVal sortExpressions As IBTSqlOrderByGroup)
        MyBase.New(transform, loadData, [alias], expressionList, type, sortExpressions)
    End Sub

    ''' <summary>
    ''' Creates a transform for a new column using related data with the default sort being the first column in <paramref name="expressionList"></paramref>,
    ''' the output column object will be of <paramref name="type"></paramref> unless not <paramref name="useFirstColAsSortExpression"></paramref>
    ''' </summary>
    Public Sub New(ByVal transform As ApplySqlTransformationRelatedDataDelegate, ByVal loadData As LoadRelatedDataForTransform, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal type As Type, Optional ByVal useFirstColAsSortExpression As Boolean = True)
        MyClass.New(transform, loadData, [alias], expressionList, type, If(useFirstColAsSortExpression, New BTSqlOrderByGroup(expressionList(0)), Nothing))
    End Sub

    Public Sub New(ByVal transform As ApplySqlTransformationRelatedDataDelegate, ByVal loadData As LoadRelatedDataForTransform, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal type As Type, ByVal sortExpression As ISqlSelectExpression)
        MyBase.New(transform, loadData, [alias], expressionList, type, New BTSqlOrderByGroup(sortExpression))
    End Sub
End Class
