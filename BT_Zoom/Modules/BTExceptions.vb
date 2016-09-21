Option Strict On
Option Explicit On

Imports System.Diagnostics
Imports System.Runtime.Serialization

Public Class BTMissingEntityExeption
    Inherits Exception

    Public Sub New(errorMessage As String)
        MyBase.New(errorMessage)
    End Sub

    Public Sub New(errorMessage As String, e As Exception)
        MyBase.New(errorMessage, e)
    End Sub
End Class

''' <summary>
''' This base class for exceptions thrown by the BuilderTrend project.
''' </summary>
Public Class BTException
    Inherits Exception

    ''' <summary>
    ''' Enumeration describing the type of custom exception that has occurred.
    ''' For all but AccessDenied, any exception message set can be considered user-friendly for display. AccessDenied should be given a generic error message.
    ''' </summary>
    Public Enum BTExceptionType
        AccessDenied
        EntityDoesntExist
        Validation
        Other
        ConnectionProviderFailure
        BadServiceRequest
        TooManyRows
        KeyNotFound
    End Enum

    Public ExceptionType As BTExceptionType = BTExceptionType.Other

    ''' <summary>
    ''' When handled by our debug emails, this will override the message sender.
    ''' </summary>
    ''' <remarks>Can be used to make particular exception types stand out.</remarks>
    Public ReadOnly Property EmailErrorFrom As String
        Get
            Select Case ExceptionType
                Case BTExceptionType.TooManyRows
                    Return "tooManyItemsReturned@buildertrend.net"
                Case Else
                    Return ""
            End Select
        End Get
    End Property

    Private _emailErrorTo As String = String.Empty
    Public Property EmailErrorTo As String
        Get
            Return _emailErrorTo
        End Get
        Set(value As String)
            _emailErrorTo = value
        End Set
    End Property

    ''' <summary>
    ''' Returns if exception contains a user-friendly error message.
    ''' </summary>
    Public ReadOnly Property IsFriendlyMessage As Boolean
        Get
            Return ExceptionType <> BTExceptionType.AccessDenied AndAlso ExceptionType <> BTExceptionType.BadServiceRequest
        End Get
    End Property

    ''' <summary>
    ''' Initializes a new instance of the <see cref="BTException"/> class.
    ''' </summary>
    Public Sub New()
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="BTException"/> class with a specified exception type
    ''' </summary>
    ''' <param name="exType">Specific type of exception for clarification</param>
    Public Sub New(exType As BTExceptionType)
        Me.New()
        ExceptionType = exType
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="BTException"/> class with a specified error message.
    ''' </summary>
    ''' <param name="message">The message that describes the error.</param>
    Public Sub New(message As String)
        MyBase.New(message)
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="BTException"/> class with a specified error message and specified exception type.
    ''' </summary>
    ''' <param name="message">The message that describes the error.</param>
    ''' <param name="exType">Specific type of exception for clarification</param>
    Public Sub New(message As String, exType As BTExceptionType)
        Me.New(message)
        ExceptionType = exType
    End Sub


    ''' <summary>
    ''' Initializes a new instance of the <see cref="BTException"/> class with a specified formatted error message.
    ''' </summary>
    ''' <param name="message">The message that describes the errors, with placeholders in the same format as <see cref="String.Format"/>.</param>
    ''' <param name="args">The values to be subsituted as in the message property's parameters.</param>
    Public Sub New(message As String, ParamArray args As Object())
        MyBase.New(String.Format(message, args))
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="BTException"/> class with a specified error message and a reference to the inner exception that is the cause of this exception.
    ''' </summary>
    ''' <param name="message">The error message that explains the reason for the exception.</param>
    ''' <param name="innerException">The exception that is the cause of the current exception, or a null reference if no inner exception is specified.</param>
    Public Sub New(message As String, innerException As Exception)
        MyBase.New(message, innerException)
    End Sub

    Public Sub New(message As String, innerException As Exception, ByVal exType As BTExceptionType)
        Me.New(message, innerException)
        ExceptionType = exType
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="BTException"/> class with serialized data.
    ''' </summary>
    ''' <param name="info">The <see cref="SerializationInfo"/> that holds the serialized object data about the exception being thrown.</param>
    ''' <param name="context">The <see cref="StreamingContext"/> that contains contextual information about the source or destination.</param>
    Public Sub New(info As SerializationInfo, context As StreamingContext)
        MyBase.New(info, context)
    End Sub
End Class

Public Class BTSqlException
    Inherits Exception

    Public Sub New(ByVal message As String)
        MyBase.New(message)
    End Sub

    Public Sub New(ByVal message As String, ByVal innerException As Exception)
        MyBase.New(message, innerException)
    End Sub

End Class

