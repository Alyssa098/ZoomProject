Option Strict On
Option Explicit On

Imports System.Web
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Threading
Imports BT_Zoom.builderTrendLLBL

Public Class RequestHandler

    Public Enum RequestItemTypes
        <Description("ConnectionStack")> ConnectionStack = 1
    End Enum

#Region "Helper"

    Private Shared Function GetItem(itemType As RequestItemTypes) As Object
        Return GetItem(itemType.ToString())
    End Function

    Private Shared Function GetItem(key As String) As Object
        Return HttpContext.Current.Items(key)
    End Function

    Private Shared Function GetItem(Of T)(itemType As RequestItemTypes) As T
        Return GetItem(Of T)(itemType.ToString())
    End Function

    Private Shared Function GetItem(Of T)(key As String) As T
        Return DirectCast(GetItem(key), T)
    End Function

    Private Shared Sub SetItem(itemType As RequestItemTypes, o As Object)
        SetItem(itemType.ToString(), o)
    End Sub

    Private Shared Sub SetItem(key As String, o As Object)
        HttpContext.Current.Items(key) = o
    End Sub

    Private Shared Function ContainsItem(itemType As RequestItemTypes) As Boolean
        Return ContainsItem(itemType.ToString())
    End Function

    Private Shared Function ContainsItem(key As String) As Boolean
        Return HttpContext.Current.Items.Contains(key)
    End Function

    Private Shared Sub RemoveItem(key As String)
        HttpContext.Current.Items.Remove(key)
    End Sub

    Private Shared Sub RemoveItem(itemType As RequestItemTypes)
        RemoveItem(itemType.ToString())
    End Sub

#End Region
#Region "ConnectionProvider - do not use outside of BTConnectionProvider"
    ''' <summary>
    ''' Return the current request's stack of <see cref="builderTrendLLBL.BTConnectionProvider"></see> objects.
    ''' DO NOT USE outside of <see cref="builderTrendLLBL.BTConnectionProvider"></see>
    ''' </summary>
    Public Shared ReadOnly Property ConnectionStack As Stack(Of BTConnectionProvider)
        Get
            If Not ContainsItem(RequestItemTypes.ConnectionStack) Then
                SetItem(RequestItemTypes.ConnectionStack, New Stack(Of BTConnectionProvider))
            End If
            Return GetItem(Of Stack(Of BTConnectionProvider))(RequestItemTypes.ConnectionStack)
        End Get
    End Property

#End Region

End Class
