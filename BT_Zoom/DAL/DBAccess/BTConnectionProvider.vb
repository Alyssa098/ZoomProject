Imports System.Collections.Generic
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Diagnostics

Namespace builderTrendLLBL

    Public Class BTConnectionProvider

        Private Const MaxBaseTransactionNameLength As Integer = 15

        Public Enum Status
            ''' <summary>
            ''' Data Access was successful, commit transaction and close the connection (if applicable)
            ''' </summary>
            Success

            ''' <summary>
            ''' Data Access was unsuccessful for some reason, rollback entire transaction and close the connection (if applicable)
            ''' </summary>
            Fail
        End Enum

        ''' <summary>
        ''' A new instance of this class will be created publicly by the shared methods
        ''' </summary>
        Private Sub New(ByVal isNewConnection As Boolean)
            If isNewConnection Then
                Try
                    SetupNewProvider()
                Catch ex As BTException
                    If ex.ExceptionType = BTException.BTExceptionType.ConnectionProviderFailure Then
                        'the DB must be unavailable
                        'push it on to the stack anyway and rethrow...we'll check that the connection is open before we handle the Failure
                        RequestHandler.ConnectionStack.Push(Me)
                        Throw
                    End If
                End Try
            Else
                'share the connection provider 
                ' do this before pushing Me to the stack so we get the shared provider 
                Provider = Current.Provider
            End If

            RequestHandler.ConnectionStack.Push(Me)
        End Sub

        Private Sub New(ByVal isNewConnection As Boolean, ByVal transactionName As String)
            Me.New(isNewConnection)

            If IsConnectionCreatedLocally Then
                SetupNewTransaction(transactionName)
            Else
                SetupTransaction(transactionName)
            End If
        End Sub

#Region "Properties"

        Public Shared ReadOnly Property Current As BTConnectionProvider
            Get
                Return GetCurrentProvider(False)
            End Get
        End Property

        Public ReadOnly Property CurrentStats As ConnectionProviderStats
            Get
                Return Current.Provider.Stats
            End Get
        End Property

        Public ReadOnly Property Connection As SqlConnection
            Get
                Return Provider.DBConnection
            End Get
        End Property

        Public ReadOnly Property Transaction As SqlTransaction
            Get
                Return Provider.CurrentTransaction
            End Get
        End Property

        Public ReadOnly Property IsTransactionPending As Boolean
            Get
                Return Provider.IsTransactionPending
            End Get
        End Property

        ''' <summary>
        ''' When true, the <see cref="ConnectionProvider"></see> was created as part of this instance
        ''' </summary>
        Private Property IsConnectionCreatedLocally As Boolean

        ''' <summary>
        ''' When true, the <see cref="ConnectionProvider.CurrentTransaction"></see> was created as part of this instance
        ''' </summary>
        Private Property IsTransactionCreatedLocally As Boolean

        Private Property Provider As ConnectionProvider

#End Region

#Region "Shared Methods"

        Public Shared Sub Create()
            Create(Nothing)
        End Sub

        ''' <summary>
        ''' Create a new connection provider, open the connection and add it to the stack
        ''' </summary>
        ''' <param name="transactionName">If provided, begin a new transaction</param>
        Private Shared Sub Create(ByVal transactionName As String)
            Dim cp As New BTConnectionProvider(True, transactionName)
        End Sub

        ''' <summary>
        ''' If the stack is empty, create a new connection provider, open the connection and add it to the stack
        ''' </summary>
        Public Shared Sub UseOrCreate()
            Dim cp As BTConnectionProvider = [Get]()
            If cp Is Nothing Then
                'create a new one, no transaction
                Create()
            End If
        End Sub

        ''' <summary>
        ''' If the stack is empty or no transaction has been started, create a new connection, open it, start a transaction and add it to the stack
        ''' </summary>
        ''' <param name="transactionName">Name of the transaction to start in the event the top of the stack is not a transaction.</param>
        ''' <param name="newConnectionIfNoTransaction">If true, a new connection will be created if a transaction has not already been started on the current connection.</param>
        ''' <remarks>If the connection provider on the top of the stack already has a transaction started, we'll return that connection, regardless of transaction name.</remarks>
        Public Shared Sub UseOrCreate(ByVal transactionName As String, Optional ByVal newConnectionIfNoTransaction As Boolean = True)
            Dim cp As BTConnectionProvider = [Get]()
            If cp Is Nothing Then
                'there's nothing on the stack, create a new one with a new transaction
                Create(transactionName)
            Else
                'the top doesn't have a transaction running, create a new provider and start a new transaction (don't use the previous connection)
                cp.SetupTransaction(transactionName, newConnectionIfNoTransaction)
            End If
        End Sub

        ''' <summary>
        ''' Returns nothing when the <see cref="RequestHandler.ConnectionStack"></see> is empty, otherwise adds a new item to the top of the stack
        ''' that shares the <see cref="ConnectionProvider"></see> with the previous top item
        ''' </summary>
        Private Shared Function [Get]() As BTConnectionProvider
            If RequestHandler.ConnectionStack.Count = 0 Then
                Return Nothing
            End If
            Return New BTConnectionProvider(False)
        End Function

        ''' <summary>
        ''' Pop the current transaction off the stack and rolls back. If the <see cref="ConnectionProvider"></see> was created locally, clean up.
        ''' </summary>
        ''' <remarks>This should always be the first action taken inside of a catch block if BTConnectionProvider has been used in the try</remarks>
        Public Shared Sub Failure()
            Done(Status.Fail)
        End Sub

        ''' <summary>
        ''' Pop the current transaction off the stack and commits. If the <see cref="ConnectionProvider"></see> was created locally, clean up.
        ''' </summary>
        ''' <remarks>This should always be the last action taken inside of a try block if BTConnectionProvider has been used in the try. 
        ''' The only exception to that rule is the return statement. A return statement can follow this method call.</remarks>
        Public Shared Sub Success()
            Done(Status.Success)
        End Sub

        ''' <summary>
        ''' Pop the current transaction off the stack. If the <see cref="ConnectionProvider"></see> was created locally, clean up.
        ''' </summary>
        ''' <param name="status">The action taken was a Success or Fail</param>
        Private Shared Sub Done(ByVal status As Status)
            ' in case of error in following try block, do not immediately pop off stack
            Dim cp As BTConnectionProvider = GetCurrentProvider(False)
            If cp.IsConnectionCreatedLocally OrElse cp.IsTransactionCreatedLocally Then
                Try
                    If cp.IsTransactionCreatedLocally AndAlso cp.Provider.IsTransactionPending Then
                        Select Case status
                            Case status.Success
                                cp.Provider.CommitTransaction()
                            Case Else 'Fail
                                cp.Provider.RollbackTransaction()
                        End Select
                    End If
                    ' pop off on success
                    GetCurrentProvider(True)
                Catch ex As Exception
                    Throw
                Finally
                    If cp.IsConnectionCreatedLocally Then
                        cp.Provider.CloseConnection(False)
                        cp.Provider.Dispose()
                    End If
                End Try
            Else
                ' no try, pop off
                GetCurrentProvider(True)
            End If
        End Sub

        ''' <summary>
        ''' This function checks to see if the given exception is a BTException. If it is, it determines if the
        ''' type of the BTException is <see cref="BTException.BTExceptionType.ConnectionProviderFailure"></see>.
        ''' If it is not, it will return true.
        ''' </summary>
        ''' <param name="ex"></param>
        ''' <returns></returns>
        ''' <remarks><see cref="BTException.BTExceptionType.ConnectionProviderFailure"></see> is used to indicate that
        ''' we do not need to rethrow the exception. This method should only be called in cases where <see cref="BTException.BTExceptionType.ConnectionProviderFailure"></see>
        ''' is used AND we are rethrowing exceptions. If you are not rethrowing an exception, do not call this method.</remarks>
        Public Shared Function ShouldRethrowException(ex As Exception) As Boolean
            If ex.GetType().IsAssignableFrom(GetType(BTException)) Then
                Dim btEx As BTException = TryCast(ex, BTException)
                If btEx IsNot Nothing AndAlso
                   btEx.ExceptionType = BTException.BTExceptionType.ConnectionProviderFailure Then
                    Return False
                End If
            End If

            Return True
        End Function

        ''' <summary>
        ''' Returns a new transaction name
        ''' </summary>
        ''' <param name="baseTransactionName">A name given to the transaction that will help identify the type of transaction.</param>
        Private Shared Function BuildTransactionName(ByVal baseTransactionName As String) As String
            If baseTransactionName.Length > MaxBaseTransactionNameLength Then
                Throw New Exception(String.Format("baseTransactionName must be less than {0} characters.", MaxBaseTransactionNameLength))
            End If

            Return String.Format("conTrans_{0}", baseTransactionName)
        End Function

        Private Shared Function GetCurrentProvider(ByVal popFromStack As Boolean) As BTConnectionProvider
            If RequestHandler.ConnectionStack.Count = 0 Then
                Throw New BTSqlException("No ConnectionProviders available.")
            End If
            If popFromStack Then
                Dim prov As BTConnectionProvider = RequestHandler.ConnectionStack.Pop()
                Return prov
            Else
                Return RequestHandler.ConnectionStack.Peek()
            End If
        End Function

#End Region

#Region "Instance Methods"

        Private Sub SetupNewProvider()
            IsConnectionCreatedLocally = True
            Provider = New ConnectionProvider()
            Provider.OpenConnection()
        End Sub

        Private Sub SetupNewTransaction(ByVal transactionName As String)
            If Not String.IsNullOrWhiteSpace(transactionName) Then
                IsTransactionCreatedLocally = True
                Provider.BeginTransaction(BuildTransactionName(transactionName))
            End If
        End Sub

        Private Sub SetupTransaction(ByVal transactionName As String, Optional ByVal newConnectionIfNoTransaction As Boolean = True)
            If Not Provider.IsTransactionPending Then
                If newConnectionIfNoTransaction Then
                    SetupNewProvider()
                End If
                SetupNewTransaction(transactionName)
            End If
        End Sub

#End Region

        ' /// <summary>
        ' /// Purpose: provides a SqlConnection object which can be shared among data-access tier objects
        ' /// to provide a way to do ADO.NET transaction coding without the hassling with SqlConnection objects
        ' /// on a high level.
        ' /// </summary>
        <Serializable()>
        Private Class ConnectionProvider
            Implements IDisposable

#Region " Class Member Declarations "

            <NonSerialized()> Private _dBConnection As SqlConnection
            Private _isTransactionPending, _isDisposed As Boolean
            Private _currentTransaction As SqlTransaction
            Private _connectionWasProvided As Boolean = False

#End Region

#Region "Properties"

            Public Property Stats As New ConnectionProviderStats

#End Region

            Public Sub New(Optional conn As SqlConnection = Nothing)
                InitClass(conn)
            End Sub

            Public Sub New(transactionName As String, Optional startTransactionOnInitialization As Boolean = True)
                Me.New()
                If OpenConnection() Then
                    If startTransactionOnInitialization Then
                        BeginTransaction(transactionName)
                    End If
                End If
            End Sub

            ' /// <summary>
            ' /// Purpose: Implements the IDispose' method Dispose.
            ' /// </summary>
            Public Overloads Sub Dispose() Implements IDisposable.Dispose
                Dispose(True)
                GC.SuppressFinalize(Me)
            End Sub

            ' /// <summary>
            ' /// Purpose: Implements the Dispose functionality.
            ' /// </summary>
            Protected Overridable Overloads Sub Dispose(ByVal isDisposing As Boolean)
                ' // Check to see if Dispose has already been called.
                If Not _isDisposed Then
                    If isDisposing Then
                        ' // Dispose managed resources.
                        If Not (_currentTransaction Is Nothing) Then
                            _currentTransaction.Dispose()
                            _currentTransaction = Nothing
                        End If
                        If _dBConnection IsNot Nothing AndAlso Not _connectionWasProvided Then
                            ' // closing the connection will abort (rollback) any pending transactions
                            _dBConnection.Close()
                            _dBConnection.Dispose()
                            _dBConnection = Nothing
                        End If
                    End If
                End If
                _isDisposed = True
            End Sub

            ' /// <summary>
            ' /// Purpose: Initializes class members.
            ' /// </summary>
            Private Sub InitClass(Optional conn As SqlConnection = Nothing)
                If conn Is Nothing Then
                    _dBConnection = New SqlConnection()
                    Dim configReader As AppSettingsReader = New AppSettingsReader()
                    _dBConnection.ConnectionString = configReader.GetValue("Main.ConnectionString", "".GetType()).ToString()
                Else
                    _dBConnection = conn
                End If
                _connectionWasProvided = conn IsNot Nothing
                _isDisposed = False
                _currentTransaction = Nothing
                _isTransactionPending = False
            End Sub

            ' /// <summary>
            ' /// Purpose: Opens the connection object.
            ' /// </summary>
            ' /// <returns>True, if succeeded, otherwise an Exception exception is thrown.</returns>
            Public Function OpenConnection() As Boolean
                If IsConnectionOpen Then
                    Throw New Exception("OpenConnection::Connection is already open.")
                End If
                If _connectionWasProvided AndAlso Not IsConnectionOpen Then
                    InitClass()
                End If
                Dim openSuccess As Boolean
                Try
                    _dBConnection.Open()
                    openSuccess = True
                    Stats.ConnectionOpenedDateUtc = Date.UtcNow
                Catch ex As Exception
                    Throw New BTException("Error opening DB Connection.", ex, BTException.BTExceptionType.ConnectionProviderFailure)
                End Try
                _isTransactionPending = False
                Return openSuccess
            End Function

            ' /// <summary>
            ' /// Purpose: Starts a new ADO.NET transaction using the open connection object of this class.
            ' /// </summary>
            ' /// <param name="transactionName">Name of the transaction to start</param>
            ' /// <returns>True, if transaction is started correctly, otherwise an Exception exception is thrown</returns>
            Public Function BeginTransaction(ByVal transactionName As String) As Boolean
                Return BeginTransaction(transactionName, IsolationLevel.ReadCommitted)
            End Function

            ' /// <summary>
            ' /// Purpose: Starts a new ADO.NET transaction using the open connection object of this class.
            ' /// </summary>
            ' /// <param name="transactionName">Name of the transaction to start</param>
            ' /// <param name="IsolationLevel">if you want to set it to read uncommitted records without locks</param>
            ' /// <returns>True, if transaction is started correctly, otherwise an Exception exception is thrown</returns>
            Public Function BeginTransaction(ByVal transactionName As String, ByVal isolationLevel As IsolationLevel) As Boolean
                If _isTransactionPending Then
                    Throw New Exception("BeginTransaction::Already transaction pending. Nesting not allowed")
                End If
                If Not IsConnectionOpen Then
                    Throw New Exception("BeginTransaction::Connection is not open.")
                End If

                _currentTransaction = _dBConnection.BeginTransaction(isolationLevel, transactionName)
                Stats.TransactionOpenedDateUtc = Date.UtcNow
                _isTransactionPending = True
                
                Return True
            End Function

            ' /// <summary>
            ' /// Purpose: Commits a pending transaction on the open connection object of this class.
            ' /// </summary>
            ' /// <returns>True, if commit was succesful, or an Exception exception is thrown</returns>
            Public Function CommitTransaction() As Boolean
                If Not _isTransactionPending Then
                    Throw New Exception("CommitTransaction::No transaction pending.")
                End If
                If Not IsConnectionOpen Then
                    Throw New Exception("CommitTransaction::Connection is not open.")
                End If

                CommitTransactionInternal()

                Stats.TransactionCommittedStatus = ConnectionProviderStats.TransactionStatus.Committed
                Stats.TransactionClosedDateUtc = Date.UtcNow
                Return True
            End Function

            ' /// <summary>
            ' /// Purpose: Rolls back a pending transaction on the open connection object of this class, 
            ' /// or rolls back to the savepoint with the given name. Savepoints are created with SaveTransaction().
            ' /// </summary>
            ' /// <param name="transactionToRollback">Name of transaction to roll back. Can be name of savepoint</param>
            ' /// <returns>True, if rollback was succesful, or an Exception exception is thrown</returns>
            Public Function RollbackTransaction() As Boolean
                If Not _isTransactionPending Then
                    Throw New Exception("RollbackTransaction::No transaction pending.")
                End If
                If Not IsConnectionOpen Then
                    Throw New Exception("RollbackTransaction::Connection is not open.")
                End If

                RollbackTransactionInternal()

                Stats.TransactionCommittedStatus = ConnectionProviderStats.TransactionStatus.RolledBack
                Stats.TransactionClosedDateUtc = Date.UtcNow
                Return True
            End Function

            ' /// <summary>
            ' /// Purpose: Closes the open connection. Depending on bCommitPendingTransactions, a pending
            ' /// transaction is commited, or aborted. 
            ' /// </summary>
            ' /// <param name="commitPendingTransaction">Flag for what to do when a transaction is still pending. True
            ' /// will commit the current transaction, False will abort (rollback) the complete current transaction.</param>
            ' /// <returns>True, if close was succesful, False if connection was already closed, or an Exception exception is thrown when
            ' /// an error occurs</returns>
            Public Function CloseConnection(ByVal commitPendingTransaction As Boolean) As Boolean
                If _connectionWasProvided Then
                    Return False
                End If
                If Not IsConnectionOpen Then
                    Return False
                End If
                If _isTransactionPending Then
                    If commitPendingTransaction Then
                        CommitTransactionInternal()
                    Else
                        RollbackTransactionInternal()
                    End If
                End If
                _dBConnection.Close()
                Stats.ConnectionClosedDateUtc = Date.UtcNow
                Return True
            End Function

            Private Sub CommitTransactionInternal()
                _currentTransaction.Commit()

                _isTransactionPending = False
                _currentTransaction.Dispose()
                _currentTransaction = Nothing

            End Sub

            Private Sub RollbackTransactionInternal()
                _currentTransaction.Rollback()

                _isTransactionPending = False
                _currentTransaction.Dispose()
                _currentTransaction = Nothing
            End Sub

            Private ReadOnly Property IsConnectionOpen As Boolean
                Get
                    Return _dBConnection IsNot Nothing AndAlso ((_dBConnection.State And ConnectionState.Open) > 0)
                End Get
            End Property

#Region " Class Property Declarations "

            Public ReadOnly Property CurrentTransaction() As SqlTransaction
                Get
                    Return _currentTransaction
                End Get
            End Property

            Public ReadOnly Property IsTransactionPending() As Boolean
                Get
                    Return _isTransactionPending
                End Get
            End Property

            Public ReadOnly Property DBConnection() As SqlConnection
                Get
                    Return _dBConnection
                End Get
            End Property

#End Region

        End Class

        Public Class ConnectionProviderStats

            Public Enum TransactionStatus
                Committed
                RolledBack
            End Enum

            Public Property ConnectionOpenedDateUtc As DateTime?
            Public Property ConnectionClosedDateUtc As DateTime?
            Public Property TransactionOpenedDateUtc As DateTime?
            Public Property TransactionClosedDateUtc As DateTime?
            Public Property TransactionCommittedStatus As TransactionStatus?

        End Class

    End Class

End Namespace