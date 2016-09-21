Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient
Imports BT_Zoom.Interfaces

Namespace BTSql

    Public Module BtSqlExtensions

        <Extension()>
        Public Function ToBTSqlParameter(ByVal parameter As SqlParameter) As IBTSqlParameter
            Return New BTSqlParameter(parameter)
        End Function

        <Extension()>
        Public Function ToExpression(ByVal fn As BTSqlFunction, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression) As BTSqlFunctionExpression
            Return New BTSqlFunctionExpression(fn, [alias], dependencies)
        End Function

        <Extension()>
        Public Function ToExpression(ByVal fn As BTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression) As BTSqlFunctionExpression
            Return New BTSqlFunctionExpression(fn, String.Empty, dependencies)
        End Function

        <Extension()>
        Public Function ToSelectQuery(ByVal sqlBuilder As ISqlSelectBuilder,
                                      Optional ByVal isTableParameterTable As Boolean = False,
                                      Optional ByVal [alias] As String = "") As BTSqlSelectQuery
            Return New BTSqlSelectQuery(sqlBuilder, isTableParameterTable, [alias])
        End Function

#Region "ToInsertValue"

        <Extension()>
        Public Function ToInsertValue(ByVal obj As String) As BTSqlTextInsertValueExpression
            Return New BTSqlTextInsertValueExpression(String.Format("'{0}'", obj))
        End Function

        <Extension()>
        Public Function ToInsertValue(ByVal obj As Integer) As BTSqlTextInsertValueExpression
            Return New BTSqlTextInsertValueExpression(obj.ToString())
        End Function

        <Extension()>
        Public Function ToInsertValue(ByVal obj As Long) As BTSqlTextInsertValueExpression
            Return New BTSqlTextInsertValueExpression(obj.ToString())
        End Function

        <Extension()>
        Public Function ToInsertValue(ByVal obj As Double) As BTSqlTextInsertValueExpression
            Return New BTSqlTextInsertValueExpression(obj.ToString())
        End Function

        <Extension()>
        Public Function ToInsertValue(ByVal obj As Decimal) As BTSqlTextInsertValueExpression
            Return New BTSqlTextInsertValueExpression(obj.ToString())
        End Function

        <Extension()>
        Public Function ToInsertValue(ByVal obj As DateTime) As BTSqlTextInsertValueExpression
            Return New BTSqlTextInsertValueExpression(String.Format("'{0}'", obj.ToString()))
        End Function

#End Region

#Region "ToSelectValue"

        <Extension()>
        Public Function ToSelectValue(ByVal obj As Integer,
                                      Optional [alias] As String = "") As BTSqlTextSelectExpression
            Return New BTSqlTextSelectExpression(obj.ToString(), [alias])
        End Function

        <Extension()>
        Public Function ToSelectValue(ByVal obj As String,
                                      Optional [alias] As String = "") As BTSqlTextSelectExpression
            Return New BTSqlTextSelectExpression(String.Format("'{0}'", obj), [alias])
        End Function

#End Region

    End Module

End Namespace
