Option Strict On
Option Explicit On

Imports System.Data
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Text
Imports System.Linq
Imports BT_Zoom.builderTrendLLBL
Imports BT_Zoom.Interfaces

Public Class BTSqlUtility

    Public Const Sql_TABLE As String = "TABLE"
    Public Const Sql_DROP As String = "DROP"
    Public Const Sql_CREATE As String = "CREATE"
    
    Private Const MaxNumberOfInsertValuesAllowed As Integer = 1000
    Public Const MaxNumberOfParametersAllowed As Integer = 2099 'the limit should be 2100 but seems if fails then, so go 1 less

    Public Shared Function GetColumnIndexFromSelectExpression(ByVal dt As DataTable, ByVal expr As ISqlSelectExpression) As Integer
        Return dt.Columns(expr.GetDataRowColumnName()).Ordinal
    End Function

    Public Shared Function GetRelatedDataTable(Of TList As IBTList)(ByVal listName As String, ByVal relatedData As Dictionary(Of String, IBTList)) As TList
        If relatedData.ContainsKey(listName) Then
            Dim l As IBTList = relatedData(listName)
            If TypeOf l Is TList Then
                Return DirectCast(l, TList)
            Else
                Throw New BTSqlException(String.Format("{0} was expected to be of type {1} but it is {2}", listName, GetType(TList), relatedData(listName).GetType))
            End If
        End If
        Return Nothing
    End Function

    Public Shared Sub AddDependencyIdentifiers(ByRef result As List(Of String), ByVal exp As IHasDependencies)
        If exp Is Nothing Then
            Exit Sub
        End If
        Dim identifiers As List(Of String) = exp.GetDependencyIdentifiers()
        For Each identifier As String In identifiers
            If Not String.IsNullOrWhitespace(identifier) AndAlso Not result.Contains(identifier) Then
                result.Add(identifier)
            End If
        Next
    End Sub

    Public Shared Sub AddDependenciesByIdentifier(ByRef result As List(Of ISqlSelectExpression), ByVal exp As IHasDependencies, ByVal owner As String, Optional ByVal doClone As Boolean = False)
        If exp Is Nothing Then
            Exit Sub
        End If
        Dim dependencies As List(Of ISqlSelectExpression) = exp.GetDependenciesByIdentifier(owner)
        For Each dependency As ISqlSelectExpression In dependencies
            If doClone Then
                result.Add(DirectCast(dependency.Clone(), ISqlSelectExpression))
            Else
                result.Add(dependency)
            End If
        Next
    End Sub


    ''' <summary>
    ''' Creates a new instance of BTSqlBuilder to generate "SELECT value FROM @sqlParamName"
    ''' </summary>
    ''' <param name="sqlTableParameter">Must be a sql table parameter</param>
    ''' <param name="suppressNoFilterWarning">If true, you can leave the where clause empty without getting a runtime exception</param>
    ''' <returns>A BTSqlBuilder created from the sqlTableParameter with the value field in the select list</returns>
    ''' <remarks>Often used in filters to filter by a list of items, consumer can add filters if necessary</remarks>
    Public Shared Function CreateSqlSelectBuilderFrom(ByVal sqlTableParameter As SqlParameter, Optional suppressNoFilterWarning As Boolean = True) As BTSqlSelectBuilder(Of BTTableParameterTable)
        Dim sb As New BTSqlSelectBuilder(Of BTTableParameterTable)(New BTTableParameterTable(sqlTableParameter), False)
        sb.AddSelect(sb.From.Value)
        sb.AllowNoFilters = suppressNoFilterWarning
        Return sb
    End Function

    ''' <summary>
    ''' Creates a new <see cref="BTSqlColumn"></see> that has no table representation for temporary column uses
    ''' </summary>
    ''' <param name="name">Column name</param>
    ''' <returns>Generated <see cref="BTSqlColumn"></see></returns>
    ''' <remarks></remarks>
    Public Shared Function CreateNewTempColumn(ByVal name As String) As ISqlSelectExpression
        Return New BTSqlColumn(New BTSqlTable("", "", "", allowNoRepresentation:=True), name)
    End Function

    Private Const NullParameterExceptionMessage As String = "Unable to determine {0} because value({1}) is null"
    Public Shared Sub CheckForNulls(functionName As String,
                                     ParamArray values() As SqlTypes.INullable)
        For i As Integer = 1 To values.Length
            If values(i - 1).IsNull Then
                Throw New BTSqlException(String.Format(NullParameterExceptionMessage, functionName, i))
            End If
        Next
    End Sub

    Public Shared Function ResetJoins(sqlJoinExpressions As List(Of ISqlJoinExpression),
                                        joinsAppliedToPageQuery As List(Of ISqlJoinExpression)) As List(Of ISqlJoinExpression)
        Dim returnExpression As New List(Of ISqlJoinExpression)
        If sqlJoinExpressions IsNot Nothing Then
            returnExpression.AddRange(sqlJoinExpressions)
        End If

        If joinsAppliedToPageQuery IsNot Nothing Then
            For Each j As ISqlJoinExpression In joinsAppliedToPageQuery
                Dim canAdd As Boolean = True
                For Each otherJoin As ISqlJoinExpression In sqlJoinExpressions
                    If j.Alias = otherJoin.Alias Then
                        canAdd = False
                    End If
                Next
                If canAdd Then
                    returnExpression.Add(j)
                End If
            Next
        End If

        Return returnExpression
    End Function

    ''' <summary>
    ''' Execute a list of ISqlExpressions
    ''' </summary>
    ''' <param name="statementList">Statement list to execute</param>
    ''' <remarks></remarks>
    Public Shared Sub ExecuteStatementGroup(statementList As IEnumerable(Of ISqlBuilder), Optional extraParams As IEnumerable(Of IBTSqlParameter) = Nothing)
        Dim parameters As New List(Of SqlParameter)()
        If extraParams IsNot Nothing AndAlso extraParams.Any() Then
            parameters.AddRange(extraParams.Select(Function(x) x.Parameter)) ''Get the extra params into our optional list
        End If
        Dim finalStatementList As New List(Of String)()
        For Each statement As ISqlBuilder In statementList
            finalStatementList.Add(statement.Render())
            parameters.AddRange(statement.Parameters)
        Next

        DataAccessHandler.ExecuteNonQuery(finalStatementList, parameters.ToArray())
    End Sub

    ''' <summary>
    ''' Execute a list of ISqlExpressions independently within a transaction.
    ''' </summary>
    ''' <param name="statementList">Statement list to execute.</param>
    ''' <remarks></remarks>
    Public Shared Sub ExecuteStatementsWithTransaction(ByVal statementList As IEnumerable(Of ISqlBuilder))
        Try
            BTConnectionProvider.UseOrCreate("ExecStmtGrp")

            For Each statement As ISqlBuilder In statementList
                DataAccessHandler.ExecuteNonQuery(statement.Render(), statement.Parameters.ToArray())
            Next

            BTConnectionProvider.Success()
        Catch exc As Exception
            BTConnectionProvider.Failure()
            Throw New BTSqlException("Error executing non-query", exc)
        End Try
    End Sub

    ''' <summary>
    ''' Combine a list of ISqlExpressions into one statement and execute it
    ''' </summary>
    ''' <param name="statementList">Statement list to combine</param>
    ''' <remarks></remarks>
    Public Shared Sub ExecuteStatementsInSingleTrip(ByVal statementList As IEnumerable(Of ISqlBuilder), Optional addDbTrans As Boolean = False)
        Dim sb As New StringBuilder()
        Dim params As New List(Of SqlParameter)

        For Each statement As ISqlBuilder In statementList
            If statement Is Nothing Then
                Continue For
            End If
            sb.Append(statement.Render().Trim())
            sb.Append(";")
            params.AddRange(statement.Parameters)
        Next

        If sb.Length > 0 Then
            If addDbTrans Then
                sb.Insert(0, "BEGIN TRAN ")
                sb.Append(" COMMIT TRAN")
            End If
            DataAccessHandler.ExecuteNonQuery(sb.ToString(), params.Distinct().ToArray())
        End If
    End Sub

    ''' <summary>
    ''' Performs a batch insert on a specified table given a column list and a collection of parameter lists.
    ''' </summary>
    ''' <typeparam name="T">The type of table that the insert will be performed on</typeparam>
    ''' <param name="table">The table to perform the insert on</param>
    ''' <param name="columnList">List of columns to perform insert on</param>
    ''' <param name="parameterLists">Lists Of IBTSqlParameters</param>
    ''' <remarks></remarks>
    Public Shared Sub PerformBatchInsert(Of T As {ISqlTableExpression, New})(ByVal table As T, ByVal columnList As List(Of IBTSqlColumn), ParamArray parameterLists() As List(Of IBTSqlParameter))
        If Not IsBatchable(columnList, parameterLists) Then
            Exit Sub
        End If

        Dim statementList As New List(Of ISqlBuilder)()

        Dim batchCountAndSize As Tuple(Of Integer, Integer) = GetBatchCountAndSize(parameterLists(0).Count, columnList.Count)
        Dim batchCount As Integer = batchCountAndSize.Item1
        Dim batchSize As Integer = batchCountAndSize.Item2

        Dim toSkip As Integer
        Dim toTake As Integer

        For i As Integer = 0 To batchCount - 1
            Dim values As New List(Of ISqlSelectExpression)
            Dim tableInsert As New BTSqlInsertValuesBuilder(Of T)(table)

            columnList.ForEach(Sub(c) tableInsert.AddColumns(c))

            toSkip = i * batchSize
            toTake = batchSize

            Dim firstParameterList As List(Of IBTSqlParameter) = parameterLists(0).Skip(toSkip).Take(toTake).ToList()

            For j As Integer = 0 To firstParameterList.Count - 1
                values.Add(tableInsert.AddParameter(firstParameterList(j)))

                For k As Integer = 1 To parameterLists.Count() - 1
                    Dim nextParameterList As List(Of IBTSqlParameter) = parameterLists(k).Skip(toSkip).Take(toTake).ToList()
                    values.Add(tableInsert.AddParameter(nextParameterList(j)))
                Next

                tableInsert.AddValues(values.ToArray())
                values.Clear()
            Next

            statementList.Add(tableInsert)
        Next

        ExecuteStatementsWithTransaction(statementList)
    End Sub

    Public Shared Sub PerformBatchInsert(Of T As {ISqlTableExpression, New})(insertBuilder As BTSqlInsertValuesBuilder(Of T))
        If insertBuilder.ColumnList.Count = 0 OrElse insertBuilder.Values.Count = 0 Then
            Exit Sub
        End If

        Dim statementList As New List(Of ISqlBuilder)()

        Dim batchCountAndSize As Tuple(Of Integer, Integer) = GetBatchCountAndSize(insertBuilder.Values.Count, insertBuilder.ColumnList.Count)
        Dim batchCount As Integer = batchCountAndSize.Item1
        Dim batchSize As Integer = batchCountAndSize.Item2

        Dim toSkip As Integer
        Dim toTake As Integer

        For i As Integer = 0 To batchCount - 1
            Dim values As New List(Of ISqlSelectExpression)
            Dim tmpInsertBuilder As New BTSqlInsertValuesBuilder(Of T)(insertBuilder.Into)

            insertBuilder.ColumnList.ForEach(Sub(c) tmpInsertBuilder.AddColumns(c))


            toSkip = i * batchSize
            toTake = batchSize

            For Each v As IEnumerable(Of ISqlSelectExpression) In insertBuilder.Values.Skip(toSkip).Take(toTake)
                tmpInsertBuilder.AddValues(v.ToArray())
                For Each ex As ISqlSelectExpression In v
                    tmpInsertBuilder.AddParameters(ex.GetParameters().Cast(Of IBTSqlParameter).ToList())
                Next
            Next

            statementList.Add(tmpInsertBuilder)
        Next

        ExecuteStatementsWithTransaction(statementList)
    End Sub

    ''' <summary>
    ''' Gets the batch count and size for a given amount of records and columns.
    ''' </summary>
    ''' <param name="recordCount">The amount of records being used</param>
    ''' <param name="columnCount">The amount of columns being used</param>
    ''' <param name="extraParams">The number of parameters that should be excluded from the max count as not available for use here</param>
    ''' <returns>Tuple containing the batch count and batch size</returns>
    ''' <remarks></remarks>
    Public Shared Function GetBatchCountAndSize(ByVal recordCount As Integer, ByVal columnCount As Integer, Optional ByVal extraParams As Integer = 0) As Tuple(Of Integer, Integer)
        Dim batchSize As Integer = MaxNumberOfInsertValuesAllowed
        Dim batchCount As Integer = CInt(Math.Ceiling(recordCount / MaxNumberOfInsertValuesAllowed))
        Dim parameterCount As Integer = recordCount * columnCount
        Dim parameterCountPerBatch As Integer = CInt(parameterCount / batchCount)

        Dim paramsAllowed As Integer = (MaxNumberOfParametersAllowed - extraParams)
        If parameterCountPerBatch > paramsAllowed Then
            batchSize = CInt(Math.Floor(paramsAllowed / columnCount)) 'get the largest number of items I can fit in a batch
            batchCount = CInt(Math.Ceiling(recordCount / batchSize)) 'then figure how many batches of that I need to complete them all
        End If

        Return New Tuple(Of Integer, Integer)(batchCount, batchSize)
    End Function

    ''' <summary>
    ''' Determines if a set of columns and parameter lists is able to be batched.
    ''' </summary>
    ''' <param name="columnList">List of columns to perform insert on</param>
    ''' <param name="parameterLists">Lists Of IBTSqlParameters</param>
    ''' <returns>Boolean value representing if batching is possible</returns>
    ''' <remarks></remarks>
    Private Shared Function IsBatchable(ByVal columnList As List(Of IBTSqlColumn), ByVal parameterLists() As List(Of IBTSqlParameter)) As Boolean
        If (columnList Is Nothing OrElse columnList.Count = 0) OrElse
           (parameterLists Is Nothing OrElse parameterLists.Count() = 0 OrElse parameterLists.All(Function(p) p.Count = 0)) Then
            Return False
        End If

        ' The number of columns we're inserting into must match the number of parameter lists.
        If columnList.Count <> parameterLists.Count() Then
            Return False
        End If

        ' All parameter lists must have the same amount of items.
        If Not parameterLists.All(Function(p) parameterLists(0).Count = p.Count) Then
            Return False
        End If

        Return True
    End Function

End Class