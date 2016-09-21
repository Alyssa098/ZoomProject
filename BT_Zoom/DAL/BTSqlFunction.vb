Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Text
Imports System.Linq
Imports System.Data.SqlClient
Imports BT_Zoom.Enums.BTSql
Imports BT_Zoom.Interfaces


'NOTE: we may want to augment this with the return type for scalar functions and the parameters passed to the functions

Public MustInherit Class BTSqlFunction
    Implements IBTSqlFunction

    Public Sub New(ByVal name As String, ByVal owner As String, ByVal isScalar As Boolean)
        _name = name
        _owner = owner
        _isScalar = isScalar
    End Sub

    Private _name As String
    Public ReadOnly Property Name As String Implements IBTSqlFunction.Name
        Get
            Return _name
        End Get
    End Property

    Private _owner As String
    Public ReadOnly Property Owner As String Implements IBTSqlFunction.Owner
        Get
            Return _owner
        End Get
    End Property

    Private _isScalar As Boolean
    Public ReadOnly Property IsScalar As Boolean Implements IBTSqlFunction.IsScalar
        Get
            Return _isScalar
        End Get
    End Property

End Class

Public Class BTSqlScalarFunction
    Inherits BTSqlFunction

    Public Sub New(ByVal name As String, ByVal owner As String)
        MyBase.New(name, owner, True)
    End Sub

End Class

Public Class BTSqlTableFunction
    Inherits BTSqlFunction

    Public Sub New(ByVal name As String, ByVal owner As String)
        MyBase.New(name, owner, False)
    End Sub

End Class

Public Class BTSqlFunctionExpression
    Implements ISqlSelectExpression

    Public Sub New(ByVal fn As IBTSqlFunction, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        _fn = fn
        Me.Alias = [alias]
        _dependencies = New List(Of ISqlSelectExpression)()
        For Each dep As ISqlSelectExpression In dependencies
            _dependencies.Add(DirectCast(dep.Clone(), ISqlSelectExpression))
        Next
    End Sub

    Public Sub New(ByVal fn As IBTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Me.New(fn, String.Empty, dependencies)
    End Sub

    Public Sub New(ByVal fnType As FunctionTypes, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        _fnType = fnType
        Me.Alias = [alias]
        _dependencies = New List(Of ISqlSelectExpression)()
        For Each dep As ISqlSelectExpression In dependencies
            _dependencies.Add(DirectCast(dep.Clone(), ISqlSelectExpression))
        Next
    End Sub

    Public Sub New(ByVal fnType As FunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Me.New(fnType, String.Empty, dependencies)
    End Sub

    Public Sub New(ByVal datePartFunctionType As DatePartFunctionTypes, ByVal datePartType As DatePartTypes, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        _isDatePart = True
        _datePartFunctionType = datePartFunctionType
        _datePartType = datePartType
        Me.Alias = [alias]
        _dependencies = New List(Of ISqlSelectExpression)()
        For Each dep As ISqlSelectExpression In dependencies
            _dependencies.Add(DirectCast(dep.Clone(), ISqlSelectExpression))
        Next
    End Sub

    Public Sub New(ByVal datePartFunctionType As DatePartFunctionTypes, ByVal datePartType As DatePartTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Me.New(datePartFunctionType, datePartType, String.Empty, dependencies)
    End Sub

    Public Sub New(ByVal convertFunctionType As DataTypeFunctionTypes, ByVal dataType As DataTypes, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        _isDataType = True
        _dataTypeFunctionType = convertFunctionType
        _dataType = dataType
        Me.Alias = [alias]
        _dependencies = New List(Of ISqlSelectExpression)()
        For Each dep As ISqlSelectExpression In dependencies
            _dependencies.Add(DirectCast(dep.Clone(), ISqlSelectExpression))
        Next
    End Sub

    Public Sub New(ByVal convertFunctionType As DataTypeFunctionTypes, ByVal dataType As DataTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Me.New(convertFunctionType, dataType, String.Empty, dependencies)
    End Sub

    Public Sub New(ByVal arithmeticFunctionType As ArithmeticFunctionTypes, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        _isArithmeticOperator = True
        _arithmeticFunctionType = arithmeticFunctionType
        Me.Alias = [alias]
        _dependencies = New List(Of ISqlSelectExpression)()
        For Each dep As ISqlSelectExpression In dependencies
            _dependencies.Add(DirectCast(dep.Clone(), ISqlSelectExpression))
        Next
    End Sub

    Public Sub New(ByVal arithmeticFunctionType As ArithmeticFunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Me.New(arithmeticFunctionType, String.Empty, dependencies)
    End Sub

    Private _fn As IBTSqlFunction
    Private _fnType As FunctionTypes

    Private _datePartFunctionType As DatePartFunctionTypes
    Private _datePartType As DatePartTypes
    Private _isDatePart As Boolean = False

    Private _dataTypeFunctionType As DataTypeFunctionTypes
    Private _dataType As DataTypes
    Private _isDataType As Boolean = False

    Private _arithmeticFunctionType As ArithmeticFunctionTypes
    Private _isArithmeticOperator As Boolean = False

    Public Property [Alias] As String Implements ISqlSelectExpression.Alias

    Private _dependencies As List(Of ISqlSelectExpression)

    Public ReadOnly Property Dependencies As List(Of ISqlSelectExpression)
        Get
            Return _dependencies
        End Get
    End Property

    Public Function WithAlias([alias] As String) As ISqlSelectExpression Implements ISqlSelectExpression.WithAlias
        Dim result As BTSqlFunctionExpression = CType(Clone(), BTSqlFunctionExpression)
        result.Alias = [alias]
        Return result
    End Function

    Public Function GetDataRowColumnName() As String Implements ISqlSelectExpression.GetDataRowColumnName
        Return GetName().Replace("[", String.Empty).Replace("]", String.Empty)
    End Function

#Region "ToString / Render"

    Public Overrides Function ToString() As String
        Dim sb As New StringBuilder()
        Dim dependencySeparator As String = ", "

        If _fn IsNot Nothing Then
            sb.AppendFormat("{0}{1}(", If(Not String.IsNullOrWhitespace(_fn.Owner), String.Format("{0}.", _fn.Owner), String.Empty), _fn.Name)
        ElseIf _isDatePart Then
            sb.AppendFormat("{0}({1}, ", _datePartFunctionType.GetDescription(), _datePartType.GetDescription())
        ElseIf _isDataType Then
            sb.AppendFormat("{0}({1}, ", _dataTypeFunctionType.GetDescription(), _dataType.GetDescription())
        ElseIf _isArithmeticOperator Then
            sb.AppendFormat("(")
            dependencySeparator = _arithmeticFunctionType.GetDescription()
        Else
            sb.AppendFormat("{0}(", _fnType.GetDescription())
            If _fnType = FunctionTypes.Count AndAlso Not _dependencies.Any Then 'no columns, use *
                sb.Append("1")
            End If
        End If

        For i As Integer = 0 To _dependencies.Count - 1
            Dim expr As ISqlSelectExpression = _dependencies(i)
            If i > 0 Then
                sb.Append(dependencySeparator)
            End If
            sb.Append(expr.RenderForFunction())
        Next
        sb.Append(")")
        Return sb.ToString()
    End Function

    Private Function GetName() As String
        Dim name As String = String.Empty
        If Not String.IsNullOrWhitespace(Me.Alias) Then
            name = Me.Alias
        Else
            If _fn IsNot Nothing Then
                Return _fn.Name
            ElseIf _isDatePart Then
                Return _datePartFunctionType.GetDescription()
            ElseIf _isDataType Then
                Return _dataType.GetDescription()
            Else
                Return _fnType.GetDescription()
            End If
        End If
        Return name
    End Function

    Public Function Render() As String Implements ISqlExpression.Render
        Return String.Format("{0} AS {1}", ToString(), GetName())
    End Function

    Public Function RenderForFilter() As String Implements ISqlSelectExpression.RenderForFilter
        Return ToString()
    End Function

    Public Function RenderForGroupBy() As String Implements ISqlSelectExpression.RenderForGroupBy
        Return ToString()
    End Function

    Public Function RenderForOrderBy() As String Implements ISqlSelectExpression.RenderForOrderBy
        Return ToString()
    End Function

    Public Function RenderForFunction() As String Implements ISqlSelectExpression.RenderForFunction
        Return ToString()
    End Function

    Public Function RenderForJoin() As String Implements ISqlSelectExpression.RenderForJoin
        Return ToString()
    End Function

    Public Overridable Function RenderForAssignment() As String Implements ISqlSelectExpression.RenderForAssignment
        Return ToString()
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

    Public Function GetDependencyIdentifiers() As List(Of String) Implements IHasDependencies.GetDependencyIdentifiers
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

    Public Sub UpdateOwnerForPaging(pagingAlias As String, ownersToChange As Collections.Generic.List(Of String)) Implements ISqlSelectExpression.UpdateOwnerForPaging
        If _dependencies IsNot Nothing AndAlso _dependencies.Count > 0 Then
            For Each d As ISqlSelectExpression In _dependencies
                d.UpdateOwnerForPaging(pagingAlias, ownersToChange)
            Next
        End If
    End Sub

    Public Function Clone() As Object Implements ICloneable.Clone
        'NOTE: this performs a shallow copy
        If _fn IsNot Nothing Then
            Return New BTSqlFunctionExpression(_fn, Me.Alias, _dependencies.ToArray())
        ElseIf _isDatePart Then
            Return New BTSqlFunctionExpression(_datePartFunctionType, _datePartType, _dependencies.ToArray())
        ElseIf _isDataType Then
            Return New BTSqlFunctionExpression(_dataTypeFunctionType, _dataType, _dependencies.ToArray())
        ElseIf _isArithmeticOperator Then
            Return New BTSqlFunctionExpression(_arithmeticFunctionType, _dependencies.ToArray())
        Else
            Return New BTSqlFunctionExpression(_fnType, Me.Alias, _dependencies.ToArray())
        End If
    End Function

End Class