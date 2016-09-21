Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Data
Imports System.Linq
Imports System.Text
Imports BT_Zoom.Delegates
Imports BT_Zoom.Interfaces
Imports System.Runtime.Serialization
Imports BT_Zoom.BTSql
Imports BT_Zoom.builderTrendLLBL
Imports BT_Zoom.Enums.BTSql


#Region "FieldInfo"

<Serializable()>
Public MustInherit Class FieldInfoBase
    Implements IFieldInfoBase

    Public MustOverride ReadOnly Property Field As IBTBaseProperty Implements IFieldInfoBase.Field
    Public Property FieldName As String Implements IFieldInfoBase.FieldName

    Public MustOverride ReadOnly Property Columns As List(Of IBTSqlColumn) Implements IFieldInfoBase.Columns

    Public MustOverride Sub CreateMeIfNecessary() Implements IFieldInfoBase.CreateMeIfNecessary

    Public Overrides Function ToString() As String Implements IFieldInfoBase.ToString
        Return FieldName
    End Function

End Class

<Serializable()>
Public Class SqlFieldInfo
    Inherits FieldInfoBase
    Implements ISqlFieldInfo

    Public Property SqlField As IBTBaseSqlProperty = Nothing Implements ISqlFieldInfo.SqlField
    Public Property Column As IBTSqlColumn Implements ISqlFieldInfo.Column
    Public Property CreateField As CreateFieldDelegate Implements ISqlFieldInfo.CreateField

    Public Overrides ReadOnly Property Field As IBTBaseProperty
        Get
            Return SqlField
        End Get
    End Property

    Public Overrides ReadOnly Property Columns As List(Of IBTSqlColumn)
        Get
            Return New List(Of IBTSqlColumn) From {Column}
        End Get
    End Property

    Public Overrides Sub CreateMeIfNecessary()
        If SqlField Is Nothing Then
            SqlField = CreateField(Column, FieldName)
        End If
    End Sub

    Public Overrides Function ToString() As String
        Dim sb As New StringBuilder()
        sb.Append(MyBase.ToString())
        sb.AppendFormat(", Column={0}, Field={1}", Column, SqlField)
        Return sb.ToString()
    End Function

End Class

<Serializable()>
Public Class DependentFieldInfo(Of TType)
    Inherits FieldInfoBase
    Implements IDependentFieldInfo(Of TType)

    Public Property DependentField As IBTDependentProperty(Of TType) = Nothing Implements IDependentFieldInfo(Of TType).DependentField
    Public Property DependentFields As List(Of IFieldInfoBase) Implements IDependentFieldInfo(Of TType).DependentFields
    Public Property CalculateValue As OfType(Of TType).CalculateValueDelegate Implements IDependentFieldInfo(Of TType).CalculateValue

    Public Overrides ReadOnly Property Field As IBTBaseProperty
        Get
            Return DependentField
        End Get
    End Property

    Public Overrides ReadOnly Property Columns As List(Of IBTSqlColumn)
        Get
            Dim result As New List(Of IBTSqlColumn)()
            For Each field As IFieldInfoBase In DependentFields
                result.AddRange(field.Columns)
            Next
            Return result
        End Get
    End Property

    Public Overrides Sub CreateMeIfNecessary()
        If Field Is Nothing Then
            If DependentFields IsNot Nothing AndAlso DependentFields.Count > 0 Then
                DependentField = New BTDependentProperty(Of TType)(FieldName, DependentFields, CalculateValue)
            Else
                Throw New Exception("Unable to create the field because there is no column and there are no dependent fields.")
            End If
        End If
    End Sub

    Public Overrides Function ToString() As String
        Dim sb As New StringBuilder()
        sb.Append(MyBase.ToString())
        sb.Append(", ")
        Dim cols As List(Of IBTSqlColumn) = Columns
        If cols.Count > 0 Then
            For i As Integer = 0 To cols.Count - 1
                sb.AppendFormat("Column{0}={1}, ", i, cols(0))
            Next
        End If
        sb.AppendFormat("Field={0}", DependentField)
        Return sb.ToString()
    End Function

End Class

#End Region

<Serializable>
Public MustInherit Class BTBaseEntity(Of TTable As {IBTSqlTable, New}, TList As {IBTBaseList(Of TTable)})
    Implements IBTBaseEntity(Of TTable, TList), ISerializable

    <NonSerialized>
    Private _table As TTable
    Public ReadOnly Property Table As TTable Implements IBTBaseEntity(Of TTable, TList).Table
        Get
            Return _table
        End Get
    End Property

    <NonSerialized()>
    Protected _list As TList

    ''' <summary>
    ''' We assume True unless <see cref="SetList"></see> was called explicitly. <see cref="SetList"></see> is called by the 
    ''' enumerator code in BTBaseEntityList and a few other places.  In these cases, the list was created outside of the Entity object.
    ''' </summary>
    ''' <remarks>The True assumption makes it so we don't have to make changes to all of the custom or auto files.</remarks>
    <NonSerialized>
    Private _listCreatedLocally As Boolean = True

    Public Sub New(ByVal table As TTable)
        _table = table
    End Sub

    Public Sub New()
        _table = CreateTable()
    End Sub

    Protected Overridable Sub SetEntityDefaults()
        'Hook to allow you to set any defaults when creating a new object (not loaded from the DB)
    End Sub

    Protected MustOverride Function CreateTable() As TTable

    Protected MustOverride Function CreateList() As TList

#Region "Serialization"

    Protected Sub PopulateFromSerializationInfo(info As SerializationInfo)
        Dim fieldDict As Dictionary(Of String, Reflection.FieldInfo) = GetAllFieldsForSerialization().ToDictionary(Function(f) f.Name)
        For Each entry As SerializationEntry In info
            If fieldDict.ContainsKey(entry.Name) Then
                Dim f As Reflection.FieldInfo = fieldDict(entry.Name)
                Dim sf As ISqlFieldInfo = TryCast(f.GetValue(Me), ISqlFieldInfo)
                If sf IsNot Nothing Then
                    sf.CreateMeIfNecessary()
                    sf.SqlField.SetMyValue(entry.Value, entry.ObjectType)
                Else
                    f.SetValue(Me, entry.Value)
                End If
            End If
        Next
    End Sub

    Public Sub GetObjectData(info As SerializationInfo, context As StreamingContext) Implements ISerializable.GetObjectData
        For Each f As Reflection.FieldInfo In GetAllFieldsForSerialization()
            If f.GetCustomAttributes(GetType(NonSerializedAttribute), True).Length = 0 Then 'skip all fields that are NonSerialized
                Dim o As Object = f.GetValue(Me)
                Dim sf As ISqlFieldInfo = TryCast(o, ISqlFieldInfo)
                If sf IsNot Nothing Then
                    If sf.Field IsNot Nothing AndAlso sf.Field.IsPopulated Then 'if it's not yet been populated, skip it
                        info.AddValue(f.Name, sf.SqlField.ValueForSerialization)
                    End If
                Else
                    info.AddValue(f.Name, o)
                End If
            End If
        Next
    End Sub

    Private Function GetAllFieldsForSerialization() As Reflection.FieldInfo()
        Return Me.GetType().GetFields(Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.Public Or Reflection.BindingFlags.FlattenHierarchy)
    End Function

#End Region

#Region "CRUD"

#Region "Create"

    ''' <summary>
    ''' Construct and execute the sql Insert for the populated fields. All non-identity primary keys are required.  If a field is not populated it will not be included in the insert statement and DB default will be assumed.
    ''' </summary>
    ''' <param name="nonIdentityFieldsToPopulateAfterInsert">Any non-identity columns that should be returned from the inserted row and populated on the object.  Identity columns will automatically be returned.</param>
    ''' <remarks>Override <see cref="PreInsertTransaction"></see> to allow for throwing a NotImplemetedException.</remarks>
    Public Sub Insert(ByVal ParamArray nonIdentityFieldsToPopulateAfterInsert As IBTBaseSqlProperty()) Implements IBTBaseEntity(Of TTable, TList).Insert
        If Not PreInsertTransaction() Then
            Exit Sub
        End If

        Dim i As New BTSqlInsertValuesBuilder(Of TTable)

        Dim vals As New List(Of ISqlSelectExpression)

        Dim outputLookup As New Dictionary(Of String, Tuple(Of IBTSqlParameter, IBTBaseSqlProperty))
        For Each fi As ISqlFieldInfo In GetPrimaryKeyFields()
            fi.CreateMeIfNecessary()
            Dim p As IBTSqlParameter = i.AddParameter(fi.SqlField.CreateSqlParameterForInsert.ToBTSqlParameter)
            If fi.Column.IsIdentity Then
                If Not fi.SqlField.IsNew Then
                    Throw New Exception(String.Format("Identity key '{0}' should be null or its new (e.g. 0) value.", fi.FieldName))
                End If
                'this is an identity, we'll add as an output and update it after we execute the sql
                outputLookup.Add(fi.FieldName, Tuple.Create(p, fi.SqlField))
            Else
                If Not fi.SqlField.IsPopulated OrElse fi.SqlField.IsNull Then
                    Throw New NoNullAllowedException(String.Format("Primary key column '{0}' cannot be null.", fi.Column.GetDataRowColumnName))
                End If
                'else add it
                i.AddColumns(fi.Column)
                vals.Add(p)
                If nonIdentityFieldsToPopulateAfterInsert.Contains(fi.SqlField) Then
                    outputLookup.Add(fi.FieldName, Tuple.Create(p, fi.SqlField))
                End If
            End If
        Next

        For Each fi As ISqlFieldInfo In GetNonPrimaryKeyFields(True)
            Dim p As IBTSqlParameter = fi.SqlField.CreateSqlParameterForInsert.ToBTSqlParameter
            Dim shouldAdd As Boolean = False
            If fi.Field.IsPopulated Then
                If fi.SqlField.IsNull AndAlso Not fi.Column.IsNullable Then
                    Throw New NoNullAllowedException(String.Format("Column '{0}' cannot be null.", fi.Column.GetDataRowColumnName))
                End If
                i.AddColumns(fi.Column)
                vals.Add(p)
                shouldAdd = True
            End If
            If nonIdentityFieldsToPopulateAfterInsert.Contains(fi.SqlField) Then
                outputLookup.Add(fi.FieldName, Tuple.Create(p, fi.SqlField))
                shouldAdd = True
            End If
            If shouldAdd Then
                i.AddParameter(p)
            End If
        Next

        i.AddValues(vals.ToArray)

        i.AddOutputFromInsert(outputLookup.Values.Select(Function(t) Tuple.Create(t.Item1, t.Item2.Column)).ToList)

        Try
            BTConnectionProvider.UseOrCreate(String.Format("Ins{0}", StringHandler.AbbreviateString(_table.TableName, 12))) 'trans name max length is 15 chars

            BeforeInsert()

            DataAccessHandler.ExecuteParameters(i.Render, CommandType.Text, i.Parameters.ToArray)

            AfterInsert()

            BTConnectionProvider.Success()
        Catch ex As Exception
            BTConnectionProvider.Failure()
            Throw New BTSqlException(String.Format("{0}::Insert::Error occurred.", _table.TableName), ex)
        End Try

        For Each t As Tuple(Of IBTSqlParameter, IBTBaseSqlProperty) In outputLookup.Values
            t.Item2.SetMyValue(t.Item1.Parameter)
        Next

        PostInsertTransaction(nonIdentityFieldsToPopulateAfterInsert.Where(Function(f) Not TypeOf f.Column.OwnerTable Is TTable).ToArray)
    End Sub

    ''' <summary>
    ''' Override for any additional processing, validation that should take place before the Insert transaction.
    ''' </summary>
    ''' <returns>Return false to cancel processing the Insert.</returns>
    Protected Overridable Function PreInsertTransaction() As Boolean
        Return True
    End Function


    ''' <summary>
    ''' Override for any additional processing that should take place before the entity insert but still part of the Insert transaction.
    ''' </summary>
    ''' <remarks>We could pass this function the insert builder to allow for updating the insert builder if necessary.</remarks>
    Protected Overridable Sub BeforeInsert()

    End Sub

    ''' <summary>
    ''' Override for any additional processing that should take place as part of the Insert transaction.
    ''' </summary>
    ''' <remarks>Examples include inserting Assigned Users or other linked data.</remarks>
    Protected Overridable Sub AfterInsert()

    End Sub

    ''' <summary>
    ''' Override for any additional processing that should take place after the Insert transaction has completed successfully.
    ''' </summary>
    ''' <param name="nonIdentityFieldsToPopulateAfterInsert">Any fields that should be populated for other related data (maybe as a pass through to the related data's insert call).  Properties/fields initially included that belong to TTable are already populated and are not included.</param>
    Protected Overridable Sub PostInsertTransaction(ByVal ParamArray nonIdentityFieldsToPopulateAfterInsert As IBTBaseSqlProperty())

    End Sub

    ''' <summary>
    ''' Set any AddedBy fields with current information (e.g. AddedById, AddedByDate, AddedByName) prior to inserting
    ''' </summary>
    Public Overridable Sub SetAddedBy(ByVal userId As Guid, ByVal timestamp As DateTime) Implements IBTBaseEntity(Of TTable, TList).SetAddedBy
        Throw New NotImplementedException("This is not implemented.")
    End Sub

#End Region

#Region "Read"

    Protected Overridable Sub InitializeRelatedDataFlags(isLoadPartial As Boolean)

    End Sub

    Public Sub LoadPartial(ParamArray fields() As IBTBaseProperty) Implements IBTBaseEntity(Of TTable, TList).LoadPartial

        If _list Is Nothing Then
            Throw New Exception("In order to LoadPartial, you need to have instantiated this entity using the primary key constructor.")
        End If
        If _list.Data IsNot Nothing Then
            Throw New Exception("Cannot LoadPartial because this entity was already loaded from the DB.")
        End If
        If fields.Length = 0 Then
            Throw New Exception("Cannot LoadPartial unless at least one field is provided.")
        End If

        For Each field As BTBaseProperty In fields
            _list.AddSelect(field.Columns.ToArray())
        Next

        LoadInternal(fields)

    End Sub

    Protected Overridable Sub PreLoadInternal(ParamArray fields() As IBTBaseProperty)

        'Hook to allow the custom implementation of the entity to add joins and corresponding select expressions (e.g. AccountingIDs)

    End Sub

    Protected Sub LoadInternal(ParamArray fields() As IBTBaseProperty)

        PreLoadInternal(fields)

        Dim cnt As Integer = _list.LoadAll()
        If cnt <> 1 Then
            Throw New BTMissingEntityExeption(String.Format("Failed to load the entity.  Query returned {0} records.  SQL => {1}", cnt, _list.Sql))
        End If

        PopulateFromDataRow(_list.Data.Rows(0), _list.ColumnAliasMappings.ToArray())

    End Sub

    Public Sub SetList(Of TTList As IBTBaseList(Of TTable))(ByVal list As TTList) Implements IBTBaseEntity(Of TTable, TList).SetList
        _listCreatedLocally = False
        _list = CType(TryCast(TryCast(list, BTBaseList(Of TTable)), IBTBaseList(Of TTable)), TList)   'wow, this is ugly!!!
    End Sub

    Public Sub PopulateFromDataRow(dr As DataRow, ParamArray mappings() As IBTSqlColumnBase) Implements IBTBaseEntity(Of TTable, TList).PopulateFromDataRow

        'set the alias on any field columns that are provided in the mappings so that GetDataRowColumnName will result
        ' in the proper string to use when indexing the column name of the DataRow
        If mappings.Length > 0 Then
            For Each mapping As IBTSqlColumnBase In mappings
                If Not String.IsNullOrWhitespace(mapping.Alias) AndAlso mapping.Alias <> mapping.Name AndAlso TypeOf mapping.OwnerTable Is TTable Then
                    Dim fi As ISqlFieldInfo = Fields.FirstOrDefault(Function(x) x.Column.Name = mapping.Name)
                    If fi IsNot Nothing Then
                        fi.Column.Alias = mapping.Alias
                    End If
                End If
            Next
        End If

        For i As Integer = 0 To dr.Table.Columns.Count - 1

            Dim columnName As String = dr.Table.Columns(i).ColumnName.ToUpperInvariant()
            If FieldsAsDictionary.ContainsKey(columnName) Then
                Dim fi As ISqlFieldInfo = FieldsAsDictionary(columnName)
                fi.CreateMeIfNecessary()
                fi.SqlField.SetMyValue(dr, columnName)
            End If

        Next

        PostPopulateFromDataRow(dr, mappings)

        If _list IsNot Nothing Then
            If _listCreatedLocally Then
                _list.Dispose()
            End If

            _list = Nothing 'we shouldn't need this anymore, we don't need to keep it around
        End If

    End Sub

    Protected Overridable Sub PostPopulateFromDataRow(dr As DataRow, ParamArray mappings() As IBTSqlColumnBase)

        'Hook to allow the custom implementation of the entity to populate any other entities with this DataRow (e.g. AccountingIDs)

    End Sub

#End Region

#Region "Update"

    ''' <summary>
    ''' Construct and execute the sql update statement for the populated fields (if dirty tracking is on, only the dirty fields will be updated).  The primary keys are required.
    ''' </summary>
    ''' <remarks>Override <see cref="PreUpdateTransaction"></see> to allow for throwing a NotImplemetedException.</remarks>
    Public Sub Update() Implements IBTBaseEntity(Of TTable, TList).Update
        If Not PreUpdateTransaction() Then
            Exit Sub
        End If

        Dim updFields As List(Of ISqlFieldInfo) = GetFieldsToUpdate()
        If Not updFields.Any Then
            'nothing to update
            Exit Sub
        End If

        Dim u As New BTSqlUpdateBuilder(Of TTable)
        AddPrimaryKeyFilter(u, "update")

        For Each fi As ISqlFieldInfo In updFields
            If Not fi.Column.IsNullable AndAlso fi.SqlField.IsNull Then
                Throw New NoNullAllowedException(String.Format("Column '{0}' cannot be null.", fi.Column.GetDataRowColumnName))
            End If
            u.AddSet(fi.Column, u.AddParameter(fi.SqlField.CreateSqlParameter.ToBTSqlParameter))
        Next

        Try
            BTConnectionProvider.UseOrCreate(String.Format("Upd{0}", StringHandler.AbbreviateString(_table.TableName, 12))) 'trans name max length is 15 chars

            BeforeUpdate()

            Dim rowsUpdated As Integer = DataAccessHandler.ExecuteNonQuery(u.Render, u.Parameters.ToArray)
            If rowsUpdated = 0 Then
                Throw New Exception("No rows were updated.")
            End If

            AfterUpdate()

            BTConnectionProvider.Success()
        Catch ex As Exception
            BTConnectionProvider.Failure()
            Throw New BTSqlException(String.Format("{0}::Update::Error occurred.", _table.TableName), ex)
        End Try

        PostUpdateTransaction()
    End Sub

    ''' <summary>
    ''' Override for any additional processing, validation that should take place before the Update transaction.
    ''' </summary>
    ''' <returns>Return false to cancel processing the Update.</returns>
    Protected Overridable Function PreUpdateTransaction() As Boolean
        Return True
    End Function

    ''' <summary>
    ''' Override for any additional processing that should take place before the entity update but still part of the Update transaction.
    ''' </summary>
    ''' <remarks>Examples include linked Photos, etc.
    ''' We could pass this function the update builder to allow for updating the update builder if necessary.</remarks>
    Protected Overridable Sub BeforeUpdate()

    End Sub

    ''' <summary>
    ''' Override for any additional processing that should take place as part of the Update transaction.
    ''' </summary>
    ''' <remarks>Examples include updating Assigned Users or other linked data.</remarks>
    Protected Overridable Sub AfterUpdate()

    End Sub

    ''' <summary>
    ''' Override for any additional processing that should take place after the Update transaction has completed successfully.
    ''' </summary>
    ''' <remarks>Examples include sending emails, etc.</remarks>
    Protected Overridable Sub PostUpdateTransaction()

    End Sub

    ''' <summary>
    ''' Set any LastUpdatedBy fields with current information (e.g. LastUpdatedById, LastUpdatedByDate, LastUpdatedByName) prior to update
    ''' </summary>
    Public Sub SetUpdatedBy(ByVal userID As Guid, ByVal timestamp As DateTime) Implements IBTBaseEntity(Of TTable, TList).SetUpdatedBy
        If IsTrackingIsDirty AndAlso Not IsDirty Then
            Exit Sub
        End If
        'only update the values if the item is already dirty
        SetUpdatedByInternal(userID, timestamp)
    End Sub

    ''' <summary>
    ''' Set any LastUpdatedBy fields with current information (e.g. LastUpdatedById, LastUpdatedByDate, LastUpdatedByName) prior to update
    ''' </summary>
    Public Sub SetUpdatedBy(ByVal userId As Guid) Implements IBTBaseEntity(Of TTable, TList).SetUpdatedBy
        If IsTrackingIsDirty AndAlso Not IsDirty Then
            Exit Sub
        End If
        'only update the values if the item is already dirty
        SetUpdatedByInternal(userId, DateTime.UtcNow)
    End Sub

    Protected Overridable Sub SetUpdatedByInternal(ByVal userId As Guid, ByVal timestamp As DateTime)
        Throw New NotImplementedException("This is not implemented.")
    End Sub

#End Region

#Region "Delete"

    ''' <summary>
    ''' Construct and execute the sql delete statement.  Only the primary keys are required to be populated.
    ''' </summary>
    ''' <remarks>Overridable to allow use to override with the List DeleteAll pattern or throwing a NotImplemetedException.</remarks>
    Public Overridable Sub Delete() Implements IBTBaseEntity(Of TTable, TList).Delete
        If Not PreDeleteTransaction() Then
            Exit Sub
        End If

        Dim d As New BTSqlDeleteBuilder(Of TTable)
        AddPrimaryKeyFilter(d, "delete")

        Try
            BTConnectionProvider.UseOrCreate(String.Format("Del{0}", StringHandler.AbbreviateString(_table.TableName, 12))) 'trans name max length is 15 chars

            BeforeDelete()

            DataAccessHandler.ExecuteParameters(d.Render, CommandType.Text, d.Parameters.ToArray)

            AfterDelete()

            BTConnectionProvider.Success()
        Catch ex As Exception
            BTConnectionProvider.Failure()
            Throw New BTSqlException(String.Format("{0}::Delete::Error occurred.", _table.TableName), ex)
        End Try

        PostDeleteTransaction()
    End Sub

    ''' <summary>
    ''' Override for any additional processing, validation that should take place before the Delete transaction.
    ''' </summary>
    ''' <returns>Return false to cancel processing the Delete.</returns>
    Protected Overridable Function PreDeleteTransaction() As Boolean
        Return True
    End Function

    ''' <summary>
    ''' Override for any additional processing that should take place before the entity delete but still part of the Delete transaction.
    ''' </summary>
    ''' <remarks>Examples include deleting foreign-keyed linked items, etc.
    ''' We could pass this function the delete builder to also allow for updating the delete builder if necessary.</remarks>
    Protected Overridable Sub BeforeDelete()

    End Sub

    ''' <summary>
    ''' Override for any additional processing that should take place as part of the Delete transaction.
    ''' </summary>
    ''' <remarks>Examples include deleting non-foreign-keyed but still related entities.</remarks>
    Protected Overridable Sub AfterDelete()

    End Sub

    ''' <summary>
    ''' Override for any additional processing that should take place after the Delete transaction has completed successfully.
    ''' </summary>
    ''' <remarks>Examples include sending emails, etc.</remarks>
    Protected Overridable Sub PostDeleteTransaction()

    End Sub

#End Region

    Private Sub AddPrimaryKeyFilter(ByVal sqlBuilder As BTSqlBuilderBase(Of TTable), ByVal crudFunc As String)
        For Each fi As ISqlFieldInfo In GetPrimaryKeyFields()
            If fi.Field Is Nothing OrElse Not fi.Field.IsPopulated Then
                Throw New Exception(String.Format("You must populate the primary keys before attempting to {0}. Column '{1}' has not been populated.", crudFunc, fi.Column.GetDataRowColumnName))
            ElseIf fi.SqlField.IsNull Then
                Throw New NoNullAllowedException(String.Format("Primary key column '{0}' cannot be null.", fi.Column.GetDataRowColumnName))
            End If
            sqlBuilder.AddFilter(fi.Column, ComparisonOperatorTypes.Equals, sqlBuilder.AddParameter(fi.SqlField.CreateSqlParameter.ToBTSqlParameter))
        Next
    End Sub

    Public Sub Save() Implements IBTBaseEntity(Of TTable, TList).Save
        Dim pkIdentities As IEnumerable(Of ISqlFieldInfo) = GetPrimaryKeyFields.Where(Function(fi) fi.Column.IsIdentity)
        If pkIdentities.Count = 1 Then
            Dim pk As ISqlFieldInfo = pkIdentities.First
            If pk.Field Is Nothing OrElse pk.SqlField.IsNew Then
                Insert()
            Else
                Update()
            End If
        Else
            Throw New NotImplementedException("Save is only implemented for Entities with exactly one PK Identity.  Please use Insert or Update.")
        End If
    End Sub

#End Region

#Region "FieldInfo"

    <NonSerialized>
    Private _fieldsAsDictionary As Dictionary(Of String, ISqlFieldInfo)
    Protected ReadOnly Property FieldsAsDictionary As Dictionary(Of String, ISqlFieldInfo)
        Get
            If _fieldsAsDictionary Is Nothing Then
                Dim columnName As String
                _fieldsAsDictionary = New Dictionary(Of String, ISqlFieldInfo)
                Fields.ForEach(Sub(x)
                                   columnName = x.Column.GetDataRowColumnName.ToUpperInvariant()
                                   If Not _fieldsAsDictionary.ContainsKey(columnName) Then
                                       _fieldsAsDictionary.Add(columnName, x)
                                   End If
                               End Sub)
            End If
            Return _fieldsAsDictionary
        End Get
    End Property

    <NonSerialized>
    Protected ReadOnly Fields As New List(Of ISqlFieldInfo)
    Protected Sub RegisterFields(ByVal ParamArray fields() As ISqlFieldInfo)
        If fields IsNot Nothing AndAlso fields.Length > 0 Then
            Me.Fields.AddRange(fields)
        End If
    End Sub

    Private Function GetCreatedFields() As List(Of ISqlFieldInfo)
        Return Fields.Where(Function(fi As ISqlFieldInfo) fi.Field IsNot Nothing).ToList()
    End Function

    Private Function GetPopulatedFields() As List(Of ISqlFieldInfo)
        Return Fields.Where(Function(fi As ISqlFieldInfo) fi.Field IsNot Nothing AndAlso fi.Field.IsPopulated).ToList()
    End Function

    Private Function GetFieldsToUpdate() As List(Of ISqlFieldInfo)
        If IsTrackingIsDirty Then
            Return GetDirtyProperties()
        Else
            Return GetPopulatedNonPrimaryKeyFields()
        End If
    End Function

    Protected Sub CreateAllFields()
        For Each field As ISqlFieldInfo In Fields
            field.CreateMeIfNecessary()
        Next
    End Sub

    Private Function GetPrimaryKeyFields() As List(Of ISqlFieldInfo)
        Return Fields.Where(Function(x As ISqlFieldInfo) x.Column.IsPrimaryKey).ToList()
    End Function

    Private Function GetNonPrimaryKeyFields(Optional ByVal createdOnly As Boolean = True) As List(Of ISqlFieldInfo)
        Dim f As IEnumerable(Of ISqlFieldInfo) = Fields.Where(Function(x As ISqlFieldInfo) Not x.Column.IsPrimaryKey)
        If Not createdOnly Then
            Return f.ToList()
        Else
            Dim pf As List(Of ISqlFieldInfo) = GetCreatedFields()
            Return f.Where(Function(fi) pf.Contains(fi)).ToList
        End If
    End Function

    Private Function GetPopulatedNonPrimaryKeyFields() As List(Of ISqlFieldInfo)
        Dim pf As List(Of ISqlFieldInfo) = GetPopulatedFields()
        Return GetNonPrimaryKeyFields.Where(Function(fi) pf.Contains(fi)).ToList
    End Function

#Region "CreateFieldDelegate Implementations"

    Protected Function CreateBTSqlByte(column As IBTSqlColumn, fieldName As String) As IBTSqlByte
        Return New BTSqlByte(column, fieldName)
    End Function

    Protected Function CreateBTSqlInt16(column As IBTSqlColumn, fieldName As String) As IBTSqlInt16
        Return New BTSqlInt16(column, fieldName)
    End Function

    Protected Function CreateBTSqlInt32(column As IBTSqlColumn, fieldName As String) As IBTSqlInt32
        Return New BTSqlInt32(column, fieldName)
    End Function

    Protected Function CreateBTSqlInt64(column As IBTSqlColumn, fieldName As String) As IBTSqlInt64
        Return New BTSqlInt64(column, fieldName)
    End Function

    Protected Function CreateBTSqlString(column As IBTSqlColumn, fieldName As String) As IBTSqlString
        Return New BTSqlString(column, fieldName)
    End Function

    Protected Function CreateBTSqlDateTime(column As IBTSqlColumn, fieldName As String) As IBTSqlDateTime
        Return New BTSqlDateTime(column, fieldName)
    End Function

    Protected Function CreateBTSqlTimeSpan(column As IBTSqlColumn, fieldName As String) As IBTSqlTimeSpan
        Return New BTSqlTimeSpan(column, fieldName)
    End Function

    Protected Function CreateBTSqlBoolean(column As IBTSqlColumn, fieldName As String) As IBTSqlBoolean
        Return New BTSqlBoolean(column, fieldName)
    End Function

    Protected Function CreateBTSqlDecimal(column As IBTSqlColumn, fieldName As String) As IBTSqlDecimal
        Return New BTSqlDecimal(column, fieldName)
    End Function

    Protected Function CreateBTSqlMoney(column As IBTSqlColumn, fieldName As String) As IBTSqlMoney
        Return New BTSqlMoney(column, fieldName)
    End Function

    Protected Function CreateBTSqlDouble(column As IBTSqlColumn, fieldName As String) As IBTSqlDouble
        Return New BTSqlDouble(column, fieldName)
    End Function

    Protected Function CreateBTSqlGuid(column As IBTSqlColumn, fieldName As String) As IBTSqlGuid
        Return New BTSqlGuid(column, fieldName)
    End Function

#End Region

#End Region

#Region "Dirty Tracking"

    Public ReadOnly Property IsDirty As Boolean Implements IBTBaseEntity(Of TTable, TList).IsDirty
        Get
            If Not IsTrackingIsDirty Then
                Return False   'we may want to throw an exception here
            End If
            For Each fi As ISqlFieldInfo In GetPopulatedFields()
                If fi.SqlField.IsDirty Then
                    Return True
                End If
            Next
            Return False
        End Get
    End Property

    <NonSerialized>
    Private _isTrackingIsDirty As Boolean = False
    Public ReadOnly Property IsTrackingIsDirty As Boolean Implements IBTBaseEntity(Of TTable, TList).IsTrackingIsDirty
        Get
            Return _isTrackingIsDirty
        End Get
    End Property

    Public Sub StartTrackingIsDirty() Implements IBTBaseEntity(Of TTable, TList).StartTrackingIsDirty
        If Not IsTrackingIsDirty Then
            _isTrackingIsDirty = True
            For Each fi As SqlFieldInfo In GetPopulatedFields()
                fi.SqlField.StartTrackingIsDirty()
            Next
        End If
    End Sub

    Public Sub StopTrackingIsDirty() Implements IBTBaseEntity(Of TTable, TList).StopTrackingIsDirty
        If IsTrackingIsDirty Then
            For Each fi As SqlFieldInfo In GetPopulatedFields()
                fi.SqlField.StopTrackingIsDirty()
            Next
            _isTrackingIsDirty = False
        End If
    End Sub

    Public Function GetDirtyProperties() As List(Of ISqlFieldInfo) Implements IBTBaseEntity(Of TTable, TList).GetDirtyProperties
        Dim result As New List(Of ISqlFieldInfo)
        If IsTrackingIsDirty Then
            result = GetPopulatedFields().Where(Function(fi) fi.SqlField IsNot Nothing AndAlso fi.SqlField.IsDirty).ToList()
        End If
        Return result
    End Function

    Public Function GetDirtyPropertiesNames() As String Implements IBTBaseEntity(Of TTable, TList).GetDirtyPropertiesNames
        Dim result As String = String.Empty
        If IsTrackingIsDirty Then
            result = Utility.ToDelimitedString(GetDirtyProperties(), Function(fi As ISqlFieldInfo) fi.SqlField.PropertyName, ", ")
        End If
        Return result
    End Function

#End Region

#Region "Partial Entities"

    Public Sub RequirePopulatedProperties(ParamArray props As IBTBaseSqlProperty()) Implements IBTBaseEntity(Of TTable, TList).RequirePopulatedProperties
        If props.IsNullOrEmpty Then
            Exit Sub
        End If
        Dim notPopulated As IBTBaseSqlProperty = props.FirstOrDefault(Function(p) Not p.IsPopulated)
        If notPopulated IsNot Nothing Then
            Throw New Exception(String.Format("{0} has not been populated.", notPopulated.PropertyName))
        End If
    End Sub

#End Region

End Class
