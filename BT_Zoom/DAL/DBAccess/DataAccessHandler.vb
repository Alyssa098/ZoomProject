Option Strict On
Option Explicit On

Imports System.Configuration
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Collections.Generic
Imports Microsoft.SqlServer.Server
Imports System.Text
Imports System.Web
Imports BT_Zoom.builderTrendLLBL
Imports BT_Zoom.Interfaces

Public Class DataAccessHandler

    Public Const Schema As String = "dbo"
    Public Const LogTableRowsOver As Integer = 5000

    Private Shared Sub LogTableIfTooBig(ByRef tbl As DataTable, cmdText As String, ByVal maxDataRowsToLog As Integer, ByVal ParamArray params() As SqlParameter)
        If tbl.Rows.Count > maxDataRowsToLog Then
            Dim msg As New StringBuilder()
            msg.AppendLine()
            msg.AppendFormat("Table {0} was returned with {1} rows and {2} columns.", If(String.IsNullOrWhitespace(tbl.TableName), "Un-named", tbl.TableName), tbl.Rows.Count, tbl.Columns.Count)
            msg.AppendLine()
            msg.AppendLine()
            msg.AppendLine(cmdText)
            If params.Length > 0 Then
                msg.AppendLine("Parameters:")
                WriteOutParameters(msg, params, vbCrLf)
            End If
        End If
    End Sub

#Region "AddDataTableToDataset"

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")>
    Public Shared Sub AddDataTableToDataset(ByRef ds As DataSet, ByVal cmdText As String, ByVal tableName As String, ByVal ParamArray params() As SqlParameter)

        Dim cmd As New SqlCommand(AppendMetaDataToQuery(cmdText))
        cmd.CommandType = CommandType.Text
        Dim da As New SqlDataAdapter(cmd)

        Dim bHasErrorCode As Boolean = AddParameters(cmd, params)

        Try
            BTConnectionProvider.UseOrCreate() 'if no current connection exists, create a new one

            cmd.Connection = BTConnectionProvider.Current.Connection
            cmd.Transaction = BTConnectionProvider.Current.Transaction

            da.Fill(ds, tableName)


            HandleErrorCode(bHasErrorCode, cmd)
            cmd.Parameters.Clear()

            BTConnectionProvider.Success()
        Catch ex As Exception
            BTConnectionProvider.Failure()
            HandleException(ex, cmdText, params)
            Throw New BTSqlException("Error adding DataTable to DataSet.", ex)
        Finally
            cmd.Dispose()
            da.Dispose()
        End Try

    End Sub

#End Region

#Region "GetDataSet"

    Public Shared Function GetDataSet(ByVal cmdText As String, ByVal cmdType As CommandType, ByVal ParamArray params() As SqlParameter) As DataSet
        Return GetDataSet(cmdText, cmdType, Nothing, params)
    End Function

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")>
    Public Shared Function GetDataSet(ByVal cmdText As String, ByVal cmdType As CommandType, tableMappings As List(Of Tuple(Of String, String)), ByVal ParamArray params() As SqlParameter) As DataSet
        Dim ds As New DataSet()

        Dim cmd As New SqlCommand(AppendMetaDataToQuery(cmdText, cmdType))
        cmd.CommandType = cmdType

        Dim da As New SqlDataAdapter(cmd)
        If tableMappings IsNot Nothing Then
            For Each tableMapping As Tuple(Of String, String) In tableMappings
                da.TableMappings.Add(tableMapping.Item1, tableMapping.Item2)
            Next
        End If

        Dim bHasErrorCode As Boolean = AddParameters(cmd, params)

        Try

            BTConnectionProvider.UseOrCreate() 'if no current connection exists, create a new one

            cmd.Connection = BTConnectionProvider.Current.Connection
            cmd.Transaction = BTConnectionProvider.Current.Transaction

            da.Fill(ds)


            HandleErrorCode(bHasErrorCode, cmd)
            cmd.Parameters.Clear()

            BTConnectionProvider.Success()
        Catch ex As Exception
            BTConnectionProvider.Failure()
            HandleException(ex, cmdText, params)
            Throw New BTSqlException("Error Getting DataSet", ex)
        Finally
            cmd.Dispose()
            da.Dispose()
        End Try

        Return ds
    End Function

#End Region

#Region "ExecuteNonQuery"

    Public Shared Function ExecuteNonQuery(ByVal sqlBuilder As ISqlBuilder) As Integer
        Return ExecuteNonQuery(sqlBuilder.Render, sqlBuilder.Parameters.ToArray)
    End Function

    Public Shared Function ExecuteNonQuery(ByVal cmdText As String, ByVal ParamArray params As SqlParameter()) As Integer
        Return ExecuteNonQuery(New String() {cmdText}, params)
    End Function

    Public Shared Function ExecuteNonQuery(ByVal cmdText As IEnumerable(Of String), ByVal ParamArray params As SqlParameter()) As Integer

        Dim rowsAffected As Integer = 0

        Dim cmd As SqlCommand = Nothing
        Dim sql As String = String.Empty

        Try

            BTConnectionProvider.UseOrCreate() 'if no current connection exists, create a new one

            For i As Integer = 0 To cmdText.Count() - 1

                sql = AppendMetaDataToQuery(cmdText(i))

                cmd = New SqlCommand(sql, BTConnectionProvider.Current.Connection, BTConnectionProvider.Current.Transaction)

                Dim bHasErrorCode As Boolean = AddParameters(cmd, params)

                rowsAffected += cmd.ExecuteNonQuery()

                HandleErrorCode(bHasErrorCode, cmd)
                cmd.Parameters.Clear()
                cmd.Dispose()

            Next

            BTConnectionProvider.Success()

        Catch exc As Exception
            BTConnectionProvider.Failure()
            HandleException(exc, sql, params)
            Throw New BTSqlException("Error executing non-query", exc)
        Finally
            If cmd IsNot Nothing Then
                cmd.Dispose()
            End If
        End Try

        Return rowsAffected
    End Function

#End Region

#Region "ExecuteParameters"

    Public Shared Sub ExecuteParameters(ByVal cmdText As String, ByVal cmdType As CommandType, ByVal ParamArray params() As SqlParameter)
        ExecuteParameters(cmdText, cmdType, 0, params)
    End Sub

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")>
    Public Shared Sub ExecuteParameters(ByVal cmdText As String, ByVal cmdType As CommandType, ByRef rowsAffected As Integer, ByVal ParamArray params() As SqlParameter)

        Dim cmd As New SqlCommand(AppendMetaDataToQuery(cmdText, cmdType))
        cmd.CommandType = cmdType
        cmd.CommandTimeout = 45    '45 seconds instead of default 20

        Dim bHasErrorCode As Boolean = AddParameters(cmd, params)

        Try
            BTConnectionProvider.UseOrCreate() 'if no current connection exists, create a new one

            cmd.Connection = BTConnectionProvider.Current.Connection
            cmd.Transaction = BTConnectionProvider.Current.Transaction

            rowsAffected = cmd.ExecuteNonQuery()


            HandleErrorCode(bHasErrorCode, cmd)
            cmd.Parameters.Clear()

            BTConnectionProvider.Success()
        Catch ex As Exception
            BTConnectionProvider.Failure()
            HandleException(ex, cmdText, params)
            Throw New BTSqlException("Error executing query with parameters.", ex)
        Finally
            cmd.Dispose()
        End Try

    End Sub

#End Region

#Region "ExecuteScalar"

    Public Shared Function ExecuteScalar(ByVal cmdText As String, cmdType As CommandType) As String
        Return ExecuteScalar(cmdText, cmdType, Nothing)
    End Function

    Public Shared Function ExecuteScalar(cmdText As String) As String
        Return ExecuteScalar(cmdText, CommandType.Text, Nothing)
    End Function

    Public Shared Function ExecuteScalar(builder As ISqlBuilder) As String
        Return ExecuteScalar(builder.Render, CommandType.Text, builder.Parameters.ToArray())
    End Function

    Public Shared Function ExecuteScalar(cmdText As String, ByVal ParamArray params() As SqlParameter) As String
        Return ExecuteScalar(cmdText, CommandType.Text, params)
    End Function

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")>
    Public Shared Function ExecuteScalar(ByVal cmdText As String, cmdType As CommandType, ByVal ParamArray params() As SqlParameter) As String

        Dim sValueToReturn As String

        If cmdType = CommandType.StoredProcedure AndAlso cmdText.Contains(Schema) = False Then
            cmdText = String.Format("{0}.{1}", Schema, cmdText)
        End If

        Dim cmd As SqlCommand = New SqlCommand(AppendMetaDataToQuery(cmdText, cmdType))
        cmd.CommandType = cmdType

        AddParameters(cmd, params)

        Try

            BTConnectionProvider.UseOrCreate() 'if no current connection exists, create a new one

            cmd.Connection = BTConnectionProvider.Current.Connection
            cmd.Transaction = BTConnectionProvider.Current.Transaction

            sValueToReturn = CStr(cmd.ExecuteScalar())


            cmd.Parameters.Clear()

            BTConnectionProvider.Success()
        Catch exc As Exception
            BTConnectionProvider.Failure()
            HandleException(exc, cmdText, params)
            Throw New BTSqlException("Error executing scalar", exc)
        Finally
            cmd.Dispose()
        End Try

        Return sValueToReturn
    End Function

#End Region

#Region "GetDataTable"

    Public Shared Function GetDataTable(ByVal cmdText As String) As DataTable
        Return GetDataTable(cmdText, CommandType.Text, "")
    End Function

    Public Shared Function GetDataTable(ByVal cmdText As String, cmdTimeout As Integer) As DataTable
        Return GetDataTable(cmdText, CommandType.Text, "", LogTableRowsOver, cmdTimeout)
    End Function

    Public Shared Function GetDataTable(ByVal cmdText As String, ByVal cmdType As CommandType, ByVal tableName As String, ByVal ParamArray params() As SqlParameter) As DataTable
        Return GetDataTable(cmdText, cmdType, tableName, LogTableRowsOver, params)
    End Function

    Public Shared Function GetDataTable(ByVal cmdText As String, ByVal cmdType As CommandType, ByVal tableName As String, ByVal maxDataRowsToLog As Integer, ByVal ParamArray params() As SqlParameter) As DataTable
        Return GetDataTable(cmdText, cmdType, tableName, maxDataRowsToLog, 30, params)
    End Function

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")>
    Public Shared Function GetDataTable(ByVal cmdText As String, ByVal cmdType As CommandType, ByVal tableName As String, ByVal maxDataRowsToLog As Integer, cmdTimeout As Integer, ByVal ParamArray params() As SqlParameter) As DataTable

        Dim toReturn As DataTable = New DataTable(tableName)

        Dim cmd As New SqlCommand(AppendMetaDataToQuery(cmdText, cmdType))
        cmd.CommandType = cmdType
        cmd.CommandTimeout = cmdTimeout
        Dim da As New SqlDataAdapter(cmd)

        Dim bHasErrorCode As Boolean = AddParameters(cmd, params)

        Try

            BTConnectionProvider.UseOrCreate() 'if no current connection exists, create a new one

            cmd.Connection = BTConnectionProvider.Current.Connection
            cmd.Transaction = BTConnectionProvider.Current.Transaction

            da.Fill(toReturn)

            HandleErrorCode(bHasErrorCode, cmd)
            cmd.Parameters.Clear()
            LogTableIfTooBig(toReturn, cmdText, maxDataRowsToLog, params)

            BTConnectionProvider.Success()
        Catch ex As Exception
            BTConnectionProvider.Failure()
            HandleException(ex, cmdText, params)
            Throw New BTSqlException("Error getting DataTable", ex)
        Finally
            cmd.Dispose()
            da.Dispose()
        End Try

        Return toReturn
    End Function

#End Region

#Region "CreateSqlParameter"

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As Decimal) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <param name="varcharLength">Optional length to supply for varchar creation</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As String, Optional varcharLength As Integer = 0) As SqlParameter
        Dim param As SqlParameter
        If varcharLength = 0 Then
            param = New SqlParameter(parameterName, If(parameterValue Is Nothing, Convert.DBNull, parameterValue))
            param.SqlDbType = SqlDbType.VarChar
        Else
            param = New SqlParameter(parameterName, SqlDbType.VarChar, varcharLength)
            param.Value = If(parameterValue Is Nothing, Convert.DBNull, parameterValue)
        End If
        Return param
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <param name="varcharLength">Optional length to supply for varchar creation</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As SqlString, Optional varcharLength As Integer = 0) As SqlParameter
        Dim param As SqlParameter
        If varcharLength = 0 Then
            param = New SqlParameter(parameterName, parameterValue)
            param.SqlDbType = SqlDbType.VarChar
        Else
            param = New SqlParameter(parameterName, SqlDbType.VarChar, varcharLength)
            param.Value = parameterValue
        End If
        Return param
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As Int32) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As Long) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As SqlInt64) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As Int32?) As SqlParameter
        Return New SqlParameter(parameterName, If(parameterValue Is Nothing, Convert.DBNull, parameterValue.Value))
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As SqlInt32) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As SqlMoney) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As DateTime) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As SqlDecimal) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As SqlDateTime) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As DateTime?) As SqlParameter
        Return New SqlParameter(parameterName, If(parameterValue Is Nothing, Convert.DBNull, parameterValue.Value))
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As Boolean) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As SqlBoolean) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As Boolean?) As SqlParameter
        Return New SqlParameter(parameterName, If(parameterValue Is Nothing, Convert.DBNull, parameterValue.Value))
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As Double) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> with the specified name and value. 
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter.</param>
    ''' <param name="parameterValue">The value that should be sent to SQL Server for this parameter.</param>
    ''' <returns></returns>
    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As TimeSpan) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> representing a list of integer values. 
    ''' The parameter will be passed to SQL Server using a table parameter.
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter in the stored procedure or SQL query.</param>
    ''' <param name="intList">A list of values to pass to the stored procedure or SQL query. The values in the list must be unique.</param>
    ''' <returns>A <see cref="SqlParameter"/> that can be added to a <see cref="SqlCommand"/> object.</returns>
    Public Shared Function CreateSqlParameter(parameterName As String, intList As IEnumerable(Of Integer)) As SqlParameter
        Dim list As New List(Of SqlDataRecord)

        For Each val As Integer In intList.Distinct()
            Dim record As SqlDataRecord = New SqlDataRecord(New SqlMetaData("value", SqlDbType.Int))
            record.SetInt32(0, val)
            list.Add(record)
        Next

        ' SqlCommand requires empty table parameters to be specified as a null value
        If list.Count = 0 Then
            list = Nothing
        End If

        Dim sqlParameter As New SqlParameter(parameterName, SqlDbType.Structured, Nothing)
        sqlParameter.TypeName = "dimension.BTIntTable"
        sqlParameter.Value = list
        Return sqlParameter
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> representing a list of integer values. 
    ''' The parameter will be passed to SQL Server using a table parameter.
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter in the stored procedure or SQL query.</param>
    ''' <param name="intList">A list of values to pass to the stored procedure or SQL query. The values in the list must be unique.</param>
    ''' <returns>A <see cref="SqlParameter"/> that can be added to a <see cref="SqlCommand"/> object.</returns>
    Public Shared Function CreateSqlParameter(parameterName As String, intList As IEnumerable(Of Long)) As SqlParameter
        Dim list As New List(Of SqlDataRecord)

        For Each val As Integer In intList
            Dim record As SqlDataRecord = New SqlDataRecord(New SqlMetaData("value", SqlDbType.BigInt))
            record.SetInt64(0, val)
            list.Add(record)
        Next

        ' SqlCommand requires empty table parameters to be specified as a null value
        If list.Count = 0 Then
            list = Nothing
        End If

        Dim sqlParameter As New SqlParameter(parameterName, SqlDbType.Structured, Nothing)
        sqlParameter.TypeName = "dimension.BTBigIntTable"
        sqlParameter.Value = list
        Return sqlParameter
    End Function

    ''' <summary>
    ''' Create a <see cref="SqlParameter"/> representing a list of string values. 
    ''' The parameter will be passed to SQL Server using a table parameter.
    ''' </summary>
    ''' <param name="parameterName">The name of the parameter in the stored procedure or SQL query.</param>
    ''' <param name="stringList">A list of values to pass to the stored procedure or SQL query. The values in the list must be unique.</param>
    ''' <returns>A <see cref="SqlParameter"/> that can be added to a <see cref="SqlCommand"/> object.</returns>
    Public Shared Function CreateSqlParameter(parameterName As String, stringList As IEnumerable(Of String)) As SqlParameter
        Dim list As New List(Of SqlDataRecord)

        For Each val As String In stringList
            Dim record As SqlDataRecord = New SqlDataRecord(New SqlMetaData("value", SqlDbType.VarChar, 4000))
            record.SetString(0, val)
            list.Add(record)
        Next

        ' SqlCommand requires empty table parameters to be specified as a null value
        If list.Count = 0 Then
            list = Nothing
        End If

        Dim sqlParameter As New SqlParameter(parameterName, SqlDbType.Structured, Nothing)
        sqlParameter.TypeName = "dimension.BTStringTable"
        sqlParameter.Value = list
        Return sqlParameter
    End Function

    Public Shared Function CreateSqlParameter(parameterName As String, dateList As IEnumerable(Of DateTime)) As SqlParameter
        Dim list As New List(Of SqlDataRecord)

        For Each val As DateTime In dateList
            Dim record As SqlDataRecord = New SqlDataRecord(New SqlMetaData("value", SqlDbType.DateTime))
            record.SetDateTime(0, val)
            list.Add(record)
        Next

        ' SqlCommand requires empty table parameters to be specified as a null value
        If list.Count = 0 Then
            list = Nothing
        End If

        Dim sqlParameter As New SqlParameter(parameterName, SqlDbType.Structured, Nothing)
        sqlParameter.TypeName = "dimension.BTDateTimeTable"
        sqlParameter.Value = list
        Return sqlParameter
    End Function

    Public Shared Function CreateSqlParameter(parameterName As String, dbType As System.Data.SqlDbType) As SqlParameter
        Return New SqlParameter(parameterName, dbType)
    End Function

    Public Shared Function CreateSqlParameter(parameterName As String, dbType As System.Data.SqlDbType, ByVal sourceColumn As IBTSqlColumn) As SqlParameter
        Return New SqlParameter(parameterName, dbType) With {.SourceColumn = sourceColumn.GetDataRowColumnName}
    End Function

    Public Shared Function CreateSqlParameter(parameterName As String, parameterValue As SqlGuid) As SqlParameter
        Return New SqlParameter(parameterName, parameterValue)
    End Function
#End Region

#Region "SetValueFromDB"

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As SqlString)
        prop = dr.BT_CSqlString(columnName)
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As SqlGuid)
        If dr.IsNull(columnName) Then
            prop = SqlGuid.Null
        Else
            prop = New SqlGuid(dr(columnName).ToString)
        End If
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As String)
        prop = dr.BT_CString(columnName)
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As SqlBoolean)
        If IsDBNull(dr(columnName)) Then
            prop = SqlBoolean.Null
        Else
            prop = New SqlBoolean(CBool(dr(columnName).ToString))
        End If
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As Boolean?)
        If IsDBNull(dr(columnName)) Then
            prop = Nothing
        Else
            prop = CBool(dr(columnName).ToString())
        End If
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As Boolean)
        prop = dr.BT_CBool(columnName)
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As SqlByte)
        If IsDBNull(dr(columnName)) Then
            prop = SqlByte.Null
        Else
            prop = New SqlByte(CByte(dr(columnName)))
        End If
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As SqlInt16)
        prop = dr.BT_CSqlInt16(columnName)
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As SqlInt32)
        prop = dr.BT_CSqlInt32(columnName)
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As SqlDecimal)
        If IsDBNull(dr(columnName)) Then
            prop = SqlDecimal.Null
        Else
            prop = New SqlDecimal(CDec(dr(columnName).ToString))
        End If
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As Integer?)
        If IsDBNull(dr(columnName)) Then
            prop = Nothing
        Else
            prop = CInt(dr(columnName).ToString())
        End If
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As Integer)
        prop = dr.BT_CInt(columnName)
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As SqlInt64)
        If IsDBNull(dr(columnName)) Then
            prop = SqlInt64.Null
        Else
            prop = New SqlInt64(CLng(dr(columnName).ToString))
        End If
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As SqlDateTime, Optional ByVal convertEmptyStringToNull As Boolean = False, Optional ByVal useUSFormat As Boolean = False)
        prop = GetSqlDateTimeValue(dr, columnName, convertEmptyStringToNull, useUSFormat)
    End Sub

    Public Shared Function GetSqlDateTimeValue(ByVal dr As DataRow, ByVal columnName As String, Optional ByVal convertEmptyStringToNull As Boolean = False, Optional ByVal useUSFormat As Boolean = False) As SqlDateTime
        Return dr.BT_CSqlDateTime(columnName, convertEmptyStringToNull, useUSFormat)
    End Function

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As DateTime?)
        prop = dr.BT_CDate(columnName)
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As DateTime)
        Dim d As DateTime?
        SetValueFromDB(dr, columnName, d)
        If d.HasValue Then
            prop = d.Value
        Else
            prop = DateTime.MinValue
        End If
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As SqlMoney, Optional ByVal convertEmptyStringToNull As Boolean = False)
        If IsDBNull(dr(columnName)) Then
            prop = SqlMoney.Null
        Else
            Dim str As String = dr(columnName).ToString()
            If String.IsNullOrWhitespace(str) AndAlso convertEmptyStringToNull Then
                prop = SqlMoney.Null
            Else
                prop = New SqlMoney(CDec(str))
            End If
        End If
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As Long?)
        If IsDBNull(dr(columnName)) Then
            prop = Nothing
        Else
            prop = CLng(dr(columnName).ToString())
        End If
    End Sub

    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As Long)
        If IsDBNull(dr(columnName)) Then
            prop = 0
        Else
            prop = CLng(dr(columnName).ToString())
        End If
    End Sub

    ''' <summary>
    ''' Generic method for parsing Enum values out of the DB
    ''' </summary>
    ''' <typeparam name="TEnum">Must be an enum type</typeparam>
    ''' <param name="dr"></param>
    ''' <param name="columnName"></param>
    ''' <param name="prop">Must be an enum value</param>
    ''' <remarks>Enums are structures that implement IConvertible, IComparable, and IFormattable. 
    ''' This is the closest we can come to a generic method constrained to enums in .NET 4.5</remarks>
    Public Shared Sub SetValueFromDB(Of TEnum As {Structure, IConvertible, IComparable, IFormattable})(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As TEnum)
        If Not GetType(TEnum).IsEnum Then
            Throw New ArgumentException("Type must be an Enum")
        End If

        If dr.IsNull(columnName) Then
            prop = Nothing 'what ever it's nothing value is 
        Else
            [Enum].TryParse(dr(columnName).ToString, prop)
        End If
    End Sub

    ''' <summary>
    ''' Generic method for parsing Enum values out of the DB
    ''' </summary>
    ''' <typeparam name="TEnum">Must be an enum type</typeparam>
    ''' <param name="dr"></param>
    ''' <param name="columnName"></param>
    ''' <param name="prop">Must be an enum value</param>
    ''' <remarks>Enums are structures that implement IConvertible, IComparable, and IFormattable. 
    ''' This is the closest we can come to a generic method constrained to enums in .NET 4.5</remarks>
    Public Shared Sub SetValueFromDB(Of TEnum As {Structure, IConvertible, IComparable, IFormattable})(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As TEnum?)
        If Not GetType(TEnum).IsEnum Then
            Throw New ArgumentException("Type must be an Enum")
        End If

        If dr.IsNull(columnName) Then
            prop = Nothing
        Else
            Dim e As TEnum
            [Enum].TryParse(dr(columnName).ToString, e)
            prop = e
        End If
    End Sub

    ''' <summary>
    ''' Assumes the data row column of dr is a comma delimited list of integers and set prop as the corresponding list
    ''' </summary>
    Public Shared Sub SetValueFromDB(ByVal dr As DataRow, ByVal columnName As String, ByRef prop As List(Of Integer), Optional ByVal delimiter As String = ",")
        If dr.IsNull(columnName) Then
            prop = New List(Of Integer) 'empty list
        Else
            prop = Utility.FromDelimitedString(dr(columnName).ToString, Function(i) CInt(i), delimiter)
        End If
    End Sub

#End Region

    Public Shared Function GetConnectionString() As SqlConnectionStringBuilder
        Return New SqlConnectionStringBuilder(ConfigurationManager.AppSettings("Main.ConnectionString"))
    End Function

    Public Shared Function GetBtDbName() As String
        Return GetConnectionString.InitialCatalog()
    End Function

    Private Shared Sub HandleException(ByVal ex As Exception, ByVal sql As String, ByVal ParamArray params() As SqlParameter)

        Dim sb As New StringBuilder()

        sb.Append("SQL<br />")
        sb.Append("=====================<br />")
        sb.AppendFormat("{0}<br /><br />", sql)

        If params IsNot Nothing Then

            sb.Append("Parameters<br />")
            sb.Append("=====================<br />")

            WriteOutParameters(sb, params, "<br />")

        End If

    End Sub

    Public Shared Sub WriteOutParameters(ByRef sb As StringBuilder, ByVal params As SqlParameter(), ByVal newLine As String)
        If sb Is Nothing Then
            sb = New StringBuilder
        End If
        For i As Integer = 0 To params.Count() - 1
            Dim p As SqlParameter = params(i)
            If p IsNot Nothing Then
                Dim pName As String = "???"
                If Not String.IsNullOrWhitespace(p.ParameterName) Then
                    pName = p.ParameterName
                End If
                Dim pValue As String = "???"
                If p.Value IsNot Nothing Then
                    If p.SqlDbType = System.Data.SqlDbType.Structured AndAlso p.TypeName = "dimension.BTIntTable" AndAlso p.Value.GetType().FullName.Contains("System.Collections.Generic.List") Then
                        Dim sb2 As New StringBuilder()
                        Dim lst As List(Of Microsoft.SqlServer.Server.SqlDataRecord) = CType(p.Value, List(Of Microsoft.SqlServer.Server.SqlDataRecord))
                        If lst IsNot Nothing Then
                            For j As Integer = 0 To lst.Count - 1
                                If j > 0 Then
                                    sb2.Append(",")
                                End If
                                sb2.AppendFormat("{0}", lst(j).Item(0).ToString())
                            Next
                            pValue = sb2.ToString()
                        End If
                    Else
                        pValue = p.Value.ToString()
                    End If
                End If

                sb.AppendFormat("{0} = {1}{2}", pName, pValue, newLine)
            End If
        Next
    End Sub

    Private Shared Sub HandleErrorCode(ByVal bHasErrorCode As Boolean, ByVal cmd As SqlCommand)
        If bHasErrorCode Then
            Dim errorCode As New SqlInt32(CType(cmd.Parameters.Item("@iErrorCode").Value.ToString(), Integer))
            If Not errorCode.Equals(New SqlInt32(LLBLError.AllOk)) Then
                Throw New Exception(String.Format("ErrorCode: {0} ", errorCode.ToString()))
            End If
        End If
    End Sub

    Private Shared Function AddParameters(ByVal cmd As SqlCommand, ByVal ParamArray params() As SqlParameter) As Boolean
        Dim bHasErrorCode As Boolean = False

        If params IsNot Nothing Then
            For Each param As SqlParameter In params
                cmd.Parameters.Add(param)
                If param.ParameterName.ToUpper() = "@iErrorCode".ToUpper() Then
                    bHasErrorCode = True
                End If
            Next
        End If

        Return bHasErrorCode
    End Function

    Private Shared Function AppendMetaDataToQuery(sqlText As String, Optional cmdType As CommandType = CommandType.Text) As String
        Return sqlText
    End Function

End Class

Public Enum LLBLError
    AllOk
    ' // Add more here (check the comma's!)
End Enum
