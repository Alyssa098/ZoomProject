Option Explicit On
Option Strict On

'NOTE: we may want to augment this with the parameters passed to the stored procedures

Public Class BTSqlStoredProcedure

    Public Sub New(ByVal name As String, ByVal owner As String)
        _name = name
        _owner = owner
    End Sub

    Private _name As String
    Public ReadOnly Property Name As String
        Get
            Return _name
        End Get
    End Property

    Private _owner As String
    Public ReadOnly Property Owner As String
        Get
            Return _owner
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return String.Format("{0}.{1}", Owner, Name)
    End Function
End Class
