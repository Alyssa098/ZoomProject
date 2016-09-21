Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports BT_Zoom.Interfaces

Public Class BTSqlCteColumn
    Inherits BTSqlColumnBase
    Implements IBTSqlCteColumn

    Private ReadOnly _cte As ISqlCteExpression

    Public Sub New(ByVal cte As ISqlCteExpression, ByVal name As String)
        _cte = cte
        _name = name
    End Sub

    Public ReadOnly Property Cte As ISqlCteExpression
        Get
            Return _cte
        End Get
    End Property
    
    Public Overrides ReadOnly Property OwnerTable() As IBaseTableExpression
        Get
            Return Cte
        End Get
    End Property

    Public Overrides Function Clone() As Object
        'NOTE: this performs a shallow copy
        Dim c As New BTSqlCteColumn(Cte, Name)
        c.Alias = [Alias]
        Return c
    End Function

End Class