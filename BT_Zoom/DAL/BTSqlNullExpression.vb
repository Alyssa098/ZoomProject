Imports BT_Zoom.Interfaces

Public Class BTSqlNullExpression
    Implements ISqlSelectExpression

    Public Sub New(ByVal [alias] As String)
        Me.Alias = [alias]
    End Sub

    Public Sub UpdateOwnerForPaging(pagingAlias As String, ownersToChange As Collections.Generic.List(Of String)) Implements ICanUpdateOwnerForPaging.UpdateOwnerForPaging

    End Sub

    Public Function GetDependenciesByIdentifier(owner As String) As Collections.Generic.List(Of ISqlSelectExpression) Implements IHasDependencies.GetDependenciesByIdentifier
        Return Nothing
    End Function

    Public Function GetDependencyIdentifiers() As Collections.Generic.List(Of String) Implements IHasDependencies.GetDependencyIdentifiers
        Return New System.Collections.Generic.List(Of String)
    End Function

    Public Function GetParameters() As Collections.Generic.List(Of IBTSqlParameter) Implements IHasParameters.GetParameters
        Return Nothing
    End Function

    Public ReadOnly Property Parameters As Collections.Generic.IEnumerable(Of Data.SqlClient.SqlParameter) Implements IHasParameters.Parameters
        Get
            Return Nothing
        End Get
    End Property

    Public Function RenderForAssignment() As String Implements ISqlAssignable.RenderForAssignment
        Return Render()
    End Function

    Public Function Render() As String Implements ISqlExpression.Render
        If String.IsNullOrEmpty(Me.Alias) Then
            Return "NULL"
        Else
            Return String.Format("NULL AS {0}", Me.Alias)
        End If
    End Function

    Public Property [Alias] As String Implements ISqlSelectExpression.Alias

    Public Function GetDataRowColumnName() As String Implements ISqlSelectExpression.GetDataRowColumnName
        Return [Alias]
    End Function

    Public Function RenderForFilter() As String Implements ISqlSelectExpression.RenderForFilter
        Return Render()
    End Function

    Public Function RenderForFunction() As String Implements ISqlSelectExpression.RenderForFunction
        Return Render()
    End Function

    Public Function RenderForGroupBy() As String Implements ISqlSelectExpression.RenderForGroupBy
        Return Render()
    End Function

    Public Function RenderForJoin() As String Implements ISqlSelectExpression.RenderForJoin
        Return Render()
    End Function

    Public Function RenderForOrderBy() As String Implements ISqlSelectExpression.RenderForOrderBy
        Return Render()
    End Function

    Public Function WithAlias([alias] As String) As ISqlSelectExpression Implements ISqlSelectExpression.WithAlias
        Dim result As BTSqlNullExpression = DirectCast(Clone(), BTSqlNullExpression)
        result.Alias = [alias]

        Return result
    End Function

    Public Function Clone() As Object Implements ICloneable.Clone
        Return New BTSqlNullExpression(Me.Alias)
    End Function
End Class
