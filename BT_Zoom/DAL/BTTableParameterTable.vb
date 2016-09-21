Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports BT_Zoom.BTSql
Imports BT_Zoom.Interfaces

Public Class BTTableParameterTable
    Inherits BTSqlTable
    Implements IBTTableParameterTable, ICanUpdateOwnerForPaging

    Private _tableParameter As IBTSqlParameter

    Public Sub New()
        MyBase.New("", "", "", "", True)
    End Sub

    Public Sub New(ByVal tableParam As IBTSqlParameter)
        Me.New()

        TableParameter = tableParam
    End Sub

    Public Sub New(ByVal tableParam As SqlParameter)
        Me.New()

        TableParameter = tableParam.ToBTSqlParameter
    End Sub

    Public Property TableParameter As IBTSqlParameter Implements IBTTableParameterTable.TableParameter
        Get
            Return _tableParameter
        End Get
        Set(ByVal value As IBTSqlParameter)
            If value.Parameter.SqlDbType <> Data.SqlDbType.Structured OrElse String.IsNullOrEmpty(value.Parameter.TypeName) Then
                Throw New BTSqlException("The parameter provided is not a table parameter.")
            End If
            _tableParameter = value
        End Set
    End Property

    Private ReadOnly _Value As IBTSqlColumn = New BTSqlColumn(Me, "value")
    Public ReadOnly Property Value As IBTSqlColumn Implements IBTTableParameterTable.Value
        Get
            Return _Value
        End Get
    End Property

#Region "Overrides and Implements"

    ''' <summary>
    ''' Overrides the BTSqlTable ToString to render the sub query in place of the FROM table
    ''' </summary>
    Public Overrides Function ToString() As String
        If TableParameter Is Nothing Then
            Throw New BTSqlException("Cannot render nothing.  Please define the table parameter.")
        End If

        Return String.Format("{0}{1}", TableParameter.RenderForFunction, If(Not String.IsNullOrWhiteSpace([Alias]), String.Format(" {0}", [Alias]), String.Empty))
    End Function

    Public Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String)) Implements ICanUpdateOwnerForPaging.UpdateOwnerForPaging
        If TableParameter IsNot Nothing Then
            TableParameter.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        End If
    End Sub

#End Region

End Class
