Option Strict On
Option Explicit On

Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Text
Imports System.Collections.Generic
Imports System.Linq
Imports BT_Zoom.Interfaces
Imports BT_Zoom.Delegates
Imports System.Runtime.Serialization

<Serializable>
Public MustInherit Class BTBaseProperty
    Implements IBTBaseProperty

    Public Sub New(ByVal propertyName As String)
        _propertyName = propertyName
    End Sub

    Private _propertyName As String
    Public ReadOnly Property PropertyName As String Implements IBTBaseProperty.PropertyName
        Get
            Return _propertyName
        End Get
    End Property

    Public MustOverride ReadOnly Property IsPopulated As Boolean Implements IBTBaseProperty.IsPopulated

    Public MustOverride ReadOnly Property Columns As List(Of IBTSqlColumn) Implements IBTBaseProperty.Columns

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String Implements IBTBaseProperty.ToString
        Return PopulatedString()
    End Function

    Protected Function PopulatedString() As String
        Return If(IsPopulated, "[populated] ", "")
    End Function

End Class

<Serializable>
Public Class BTDependentProperty(Of TType)
    Inherits BTBaseProperty
    Implements IBTDependentProperty(Of TType)

    Public Sub New(ByVal name As String, dependentFields As List(Of IFieldInfoBase), calculateValue As OfType(Of TType).CalculateValueDelegate)
        MyBase.New(name)

        _dependentFields = dependentFields
        _calculateValue = calculateValue
    End Sub

    Private _dependentFields As List(Of IFieldInfoBase)
    Private _calculateValue As OfType(Of TType).CalculateValueDelegate

    Public Overrides ReadOnly Property Columns As List(Of IBTSqlColumn)
        Get
            Dim result As New List(Of IBTSqlColumn)()
            For Each fi As IFieldInfoBase In _dependentFields
                result.AddRange(fi.Columns)
            Next
            Return result
        End Get
    End Property

    Public Overrides ReadOnly Property IsPopulated As Boolean
        Get
            For Each fi As IFieldInfoBase In _dependentFields
                If Not fi.Field.IsPopulated Then
                    Return False
                End If
            Next
            Return True
        End Get
    End Property

    Public ReadOnly Property Value As TType Implements IBTDependentProperty(Of TType).Value
        Get
            Return _calculateValue(_dependentFields)
        End Get
    End Property

    Public Overrides Function ToString() As String
        Dim sb As New StringBuilder()
        sb.Append(MyBase.PopulatedString())
        If IsPopulated Then
            sb.AppendFormat("Value={0} ", Value)
        End If
        Return sb.ToString()
    End Function

End Class

#Region "Sql Types"

<Serializable>
Public MustInherit Class BTBaseSqlProperty
    Inherits BTBaseProperty
    Implements IBTBaseSqlProperty

    Protected MustOverride Function ValueAsString() As String Implements IBTBaseSqlProperty.ValueAsString

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(name)

        _column = column
        _isPopulated = False

    End Sub

    Private _column As IBTSqlColumn
    Public ReadOnly Property Column As IBTSqlColumn Implements IBTBaseSqlProperty.Column
        Get
            Return _column
        End Get
    End Property

    Public Overrides ReadOnly Property Columns As List(Of IBTSqlColumn)
        Get
            Return New List(Of IBTSqlColumn) From {Column}
        End Get
    End Property

    Protected _isPopulated As Boolean
    Public Overrides ReadOnly Property IsPopulated As Boolean
        Get
            Return _isPopulated
        End Get
    End Property

    Protected MustOverride ReadOnly Property ValueForSerialization As Object Implements IBTBaseSqlProperty.ValueForSerialization
    Protected MustOverride Sub SetMyValue(ByVal serializedObj As Object, ByVal serializedType As Type) Implements IBTBaseSqlProperty.SetMyValue

    Public MustOverride Sub SetMyValue(dr As DataRow, columnName As String) Implements IBTBaseSqlProperty.SetMyValue
    Public MustOverride Sub SetMyValue(p As SqlParameter) Implements IBTBaseSqlProperty.SetMyValue
    Public MustOverride ReadOnly Property IsNull As Boolean Implements IBTBaseSqlProperty.IsNull
    Public MustOverride Sub Clear(Optional ByVal showAsPopulated As Boolean = False) Implements IBTBaseSqlProperty.Clear


    Public Overridable ReadOnly Property IsNew As Boolean Implements IBTBaseSqlProperty.IsNew
        Get
            Return Not IsPopulated OrElse IsNull
        End Get
    End Property

    ''' <summary>
    ''' Two properties are considered equal if their names, column names and underlying Table type are the same
    ''' </summary>
    Public Overrides Function Equals(obj As Object) As Boolean Implements IBTBaseSqlProperty.Equals
        Dim compareTo As BTBaseSqlProperty = TryCast(obj, BTBaseSqlProperty)
        If compareTo Is Nothing Then
            Return False
        End If
        Return (PropertyName = compareTo.PropertyName AndAlso Column.Name = compareTo.Column.Name AndAlso Column.OwnerTable.GetType Is compareTo.Column.OwnerTable.GetType)
    End Function

#Region "Dirty Tracking"

    Public MustOverride ReadOnly Property IsDirty As Boolean Implements IBTBaseSqlProperty.IsDirty

    Protected MustOverride Sub TakeSnapshot() Implements IBTBaseSqlProperty.TakeSnapshot
    Protected MustOverride Sub ClearSnapshot() Implements IBTBaseSqlProperty.ClearSnapshot

    <NonSerialized>
    Private _isTrackingIsDirty As Boolean = False
    Public ReadOnly Property IsTrackingIsDirty As Boolean Implements IBTBaseSqlProperty.IsTrackingIsDirty
        Get
            Return _isTrackingIsDirty
        End Get
    End Property

    Public Sub StartTrackingIsDirty() Implements IBTBaseSqlProperty.StartTrackingIsDirty
        _isTrackingIsDirty = True
        TakeSnapshot()
    End Sub

    Public Sub StopTrackingIsDirty() Implements IBTBaseSqlProperty.StopTrackingIsDirty
        ClearSnapshot()
        _isTrackingIsDirty = False
    End Sub

#End Region

#Region "SqlParameter"

    Public MustOverride Function CreateSqlParameterForInsert() As SqlParameter Implements IBTBaseSqlProperty.CreateSqlParameterForInsert
    Public MustOverride Function CreateSqlParameter() As SqlParameter Implements IBTBaseSqlProperty.CreateSqlParameter

#End Region

End Class

<Serializable()>
Public MustInherit Class BTSqlProperty(Of TObjType, TPrimitiveType)
    Inherits BTBaseSqlProperty
    Implements IBTSqlProperty(Of TObjType, TPrimitiveType), IComparable

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(column, name)

        _wasPopulated = False
    End Sub

    <NonSerialized>
    Private _wasPopulated As Boolean
    Private ReadOnly Property WasPopulated As Boolean Implements IBTSqlProperty(Of TObjType, TPrimitiveType).WasPopulated
        Get
            Return _wasPopulated
        End Get
    End Property

    <NonSerialized>
    Private _objOriginal As TObjType
    Protected ReadOnly Property ObjOriginal As TObjType Implements IBTSqlProperty(Of TObjType, TPrimitiveType).ObjOriginal
        Get
            Return _objOriginal
        End Get
    End Property

    Private _obj As TObjType
    Public Overridable Property Obj As TObjType Implements IBTSqlProperty(Of TObjType, TPrimitiveType).Obj
        Get
            If Not IsPopulated Then
                Throw New Exception(String.Format("{0} should be populated before attempting to use its value.", PropertyName))
            End If
            Return _obj
        End Get
        Set(value As TObjType)
            _obj = value
            _isPopulated = True  'Assumption is that if any value (null or otherwise) is assigned to this property, then it is populated
            ValidateNull(Column.IsNullable, Column.GetDataRowColumnName)
        End Set
    End Property

    Protected Overrides ReadOnly Property ValueForSerialization As Object
        Get
            Return Obj
        End Get
    End Property

    Protected Overrides Sub SetMyValue(serializedObj As Object, ByVal serializedType As Type)
        If GetType(TObjType) Is serializedType Then
            Obj = DirectCast(serializedObj, TObjType)
        Else
            Throw New SerializationException(String.Format("Serialized type was '{0}' but property expected '{1}'", serializedType, GetType(TObjType)))
        End If
    End Sub

    Protected Overrides Sub TakeSnapshot()
        _wasPopulated = IsPopulated
        _objOriginal = _obj
    End Sub

    Protected Overrides Sub ClearSnapshot()
        _wasPopulated = False
        _objOriginal = Nothing
    End Sub

    Protected Sub CanCallIsDirty()
        If Not WasPopulated AndAlso IsPopulated Then
            Throw New Exception(String.Format("Cannot call IsDirty on {0} because it didn't have a snapshot value, but it does have a value now.", PropertyName))
        End If
    End Sub

    Public Overrides ReadOnly Property IsDirty As Boolean
        Get
            If Not IsTrackingIsDirty Then
                Throw New Exception(String.Format("Cannot call IsDirty on {0} because it does not have dirty tracking ability turned on.", PropertyName))
            Else
                CanCallIsDirty()
                If IsNull AndAlso OriginalIsNull Then
                    Return False
                ElseIf IsNull AndAlso Not OriginalIsNull Then
                    Return True
                ElseIf Not IsNull AndAlso OriginalIsNull Then
                    Return True
                Else
                    Return IsDirtyInternal
                End If
            End If
        End Get
    End Property

    Protected MustOverride ReadOnly Property IsDirtyInternal As Boolean

    Public MustOverride ReadOnly Property Value As TPrimitiveType Implements IBTSqlProperty(Of TObjType, TPrimitiveType).Value

    Public MustOverride ReadOnly Property OriginalValue As TPrimitiveType Implements IBTSqlProperty(Of TObjType, TPrimitiveType).OriginalValue
    Public MustOverride ReadOnly Property OriginalIsNull As Boolean Implements IBTSqlProperty(Of TObjType, TPrimitiveType).OriginalIsNull

    Protected MustOverride ReadOnly Property NullValue As TObjType Implements IBTSqlProperty(Of TObjType, TPrimitiveType).NullValue

    ''' <summary>
    ''' Set to corresponding Null (New) value without performing a null validity check
    ''' </summary>
    Public Overrides Sub Clear(Optional ByVal showAsPopulated As Boolean = False)
        _obj = NullValue
        _isPopulated = showAsPopulated
    End Sub

    Public ReadOnly Property IsNotNull As Boolean Implements IBTSqlProperty(Of TObjType, TPrimitiveType).IsNotNull
        Get
            Return Not IsNull
        End Get
    End Property

    Private Sub ValidateNull(allowNull As Boolean, itemName As String)
        If Not allowNull AndAlso IsNull Then
            Throw New NoNullAllowedException(String.Format("The value for {0} cannot be null when assigning to the {1} property", itemName, PropertyName))
        End If
    End Sub

    Public Overrides Function ToString() As String
        Dim sb As New StringBuilder()
        sb.Append(MyBase.PopulatedString())
        If Not IsPopulated Then
            sb.Append("[not populated] ")
        ElseIf IsNull Then
            sb.Append("[null] ")
        Else
            sb.AppendFormat("Value={0} ", ValueAsString)
        End If
        If IsTrackingIsDirty AndAlso IsDirty Then
            sb.Append("[dirty] ")
            'If we want to include the original value in the display, then we'll need a MustOverride OriginalValue property
        End If
        Return sb.ToString()
    End Function

    Protected Overrides Function ValueAsString() As String
        Return Value.ToString
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        If Me.GetType() Is obj.GetType() Then
            Return MyBase.Equals(obj)
        End If
        Return False
    End Function

    'Public MustOverride Function CompareTo(otherValue As TObjType) As Integer Implements IComparable.CompareTo

    Public MustOverride Function CompareTo(ByVal otherValue As Object) As Integer Implements IComparable.CompareTo

#Region "SqlParameter"

    Public Overrides Function CreateSqlParameterForInsert() As SqlParameter
        Dim sp As SqlParameter = Column.CreateSqlParameterForInsert()
        SetParameterValue(sp)
        Return sp
    End Function

    Public Overrides Function CreateSqlParameter() As SqlParameter
        Dim sp As SqlParameter = Column.CreateSqlParameter()
        SetParameterValue(sp)
        Return sp
    End Function

    Protected Overridable Sub SetParameterValue(ByRef sp As SqlParameter)
        sp.Value = _obj
    End Sub

#End Region

End Class

<Serializable()>
Public Class BTSqlInt16
    Inherits BTSqlProperty(Of SqlInt16, Short)
    Implements IBTSqlInt16

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(column, name)

    End Sub

    Public Overrides ReadOnly Property Value As Short
        Get
            Return Obj.Value
        End Get
    End Property

    Public Overrides ReadOnly Property IsNull As Boolean
        Get
            Return Obj.IsNull
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalValue As Short
        Get
            Return ObjOriginal.Value
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalIsNull As Boolean
        Get
            Return ObjOriginal.IsNull
        End Get
    End Property

    Protected Overrides ReadOnly Property IsDirtyInternal As Boolean
        Get
            Return Not (Value = OriginalValue)
        End Get
    End Property

    Public Overrides Sub SetMyValue(dr As DataRow, columnName As String)
        Obj = dr.BT_CSqlInt16(columnName)
    End Sub

    Public Overrides Sub SetMyValue(p As SqlParameter)
        Obj = DirectCast(p.Value, SqlInt16)
    End Sub

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String
        Return String.Format("[Int16] {0}", MyBase.ToString())
    End Function

    Protected Overrides ReadOnly Property NullValue As SqlInt16
        Get
            Return SqlInt16.Null
        End Get
    End Property

    Public Overrides ReadOnly Property IsNew As Boolean
        Get
            Return MyBase.IsNew OrElse Value = 0
        End Get
    End Property

    Public Overrides Function CompareTo(otherValue As Object) As Integer
        Dim o As BTSqlInt16 = TryCast(otherValue, BTSqlInt16)
        If o Is Nothing Then
            Return -1
        End If
        Return Obj.CompareTo(o.Obj)
    End Function

End Class

<Serializable()>
Public Class BTSqlInt32
    Inherits BTSqlProperty(Of SqlInt32, Integer)
    Implements IBTSqlInt32

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(column, name)

    End Sub

    Public Overrides ReadOnly Property Value As Integer
        Get
            Return Obj.Value
        End Get
    End Property

    Public Overrides ReadOnly Property IsNull As Boolean
        Get
            Return Obj.IsNull
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalValue As Integer
        Get
            Return ObjOriginal.Value
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalIsNull As Boolean
        Get
            Return ObjOriginal.IsNull
        End Get
    End Property

    Protected Overrides ReadOnly Property IsDirtyInternal As Boolean
        Get
            Return Not (Value = OriginalValue)
        End Get
    End Property

    Public Overrides Sub SetMyValue(dr As DataRow, columnName As String)
        Obj = dr.BT_CSqlInt32(columnName)
    End Sub

    Public Overrides Sub SetMyValue(p As SqlParameter)
        Obj = DirectCast(p.Value, SqlInt32)
    End Sub

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String
        Return String.Format("[Int32] {0}", MyBase.ToString())
    End Function

    Protected Overrides ReadOnly Property NullValue As SqlInt32
        Get
            Return SqlInt32.Null
        End Get
    End Property

    Public Overrides ReadOnly Property IsNew As Boolean
        Get
            Return MyBase.IsNew OrElse Value = 0
        End Get
    End Property

    Public Overrides Function CompareTo(otherValue As Object) As Integer
        Dim o As BTSqlInt32 = TryCast(otherValue, BTSqlInt32)
        If o Is Nothing Then
            Return -1
        End If
        Return Obj.CompareTo(o.Obj)
    End Function


End Class

<Serializable()>
Public Class BTSqlInt64
    Inherits BTSqlProperty(Of SqlInt64, Long)
    Implements IBTSqlInt64

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(column, name)

    End Sub

    Public Overrides ReadOnly Property Value As Long
        Get
            Return Obj.Value
        End Get
    End Property

    Public Overrides ReadOnly Property IsNull As Boolean
        Get
            Return Obj.IsNull
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalValue As Long
        Get
            Return ObjOriginal.Value
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalIsNull As Boolean
        Get
            Return ObjOriginal.IsNull
        End Get
    End Property

    Protected Overrides ReadOnly Property IsDirtyInternal As Boolean
        Get
            Return Not (Value = OriginalValue)
        End Get
    End Property

    Public Overrides Sub SetMyValue(dr As DataRow, columnName As String)
        Obj = dr.BT_CSqlInt64(columnName)
    End Sub

    Public Overrides Sub SetMyValue(p As SqlParameter)
        Obj = DirectCast(p.Value, SqlInt64)
    End Sub

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String
        Return String.Format("[Int64] {0}", MyBase.ToString())
    End Function

    Protected Overrides ReadOnly Property NullValue As SqlInt64
        Get
            Return SqlInt64.Null
        End Get
    End Property

    Public Overrides ReadOnly Property IsNew As Boolean
        Get
            Return MyBase.IsNew OrElse Value = 0
        End Get
    End Property

    Public Overrides Function CompareTo(otherValue As Object) As Integer
        Dim o As BTSqlInt64 = TryCast(otherValue, BTSqlInt64)
        If o Is Nothing Then
            Return -1
        End If
        Return Obj.CompareTo(o.Obj)
    End Function

End Class

<Serializable()>
Public Class BTSqlString
    Inherits BTSqlProperty(Of SqlString, String)
    Implements IBTSqlString

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(column, name)

    End Sub

    Public Overrides ReadOnly Property Value As String
        Get
            Return Obj.Value
        End Get
    End Property

    Public Overrides ReadOnly Property IsNull As Boolean
        Get
            Return Obj.IsNull
        End Get
    End Property

    Public ReadOnly Property IsNullOrEmpty As Boolean Implements IBTSqlString.IsNullOrEmpty
        Get
            Return Obj.IsNull OrElse Obj.Value.IsNullOrEmpty
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalValue As String
        Get
            Return ObjOriginal.Value
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalIsNull As Boolean
        Get
            Return ObjOriginal.IsNull
        End Get
    End Property

    Protected Overrides ReadOnly Property IsDirtyInternal As Boolean
        Get
            Return Not (Value = OriginalValue)
        End Get
    End Property

    Public Overrides Sub SetMyValue(dr As DataRow, columnName As String)
        Obj = dr.BT_CSqlString(columnName)
    End Sub

    Public Overrides Sub SetMyValue(p As SqlParameter)
        Obj = DirectCast(p.Value, SqlString)
    End Sub

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String
        Return String.Format("[String] {0}", MyBase.ToString())
    End Function

    Protected Overrides Sub SetParameterValue(ByRef sp As SqlParameter)
        MyBase.SetParameterValue(sp)
        If sp.SqlDbType = SqlDbType.Text Then
            'not sure if this is necessary but we're doing this in the old entities for Text columns
            Dim length As Integer = 0
            If IsPopulated AndAlso Not IsNull Then
                length = Value.Length
            End If
            sp.Size = length
        End If
    End Sub

    Protected Overrides ReadOnly Property NullValue As SqlString
        Get
            Return SqlString.Null
        End Get
    End Property

    Public Overrides Function CompareTo(otherValue As Object) As Integer
        Dim o As BTSqlString = TryCast(otherValue, BTSqlString)
        If o Is Nothing Then
            Return -1
        End If
        Return Obj.CompareTo(o.Obj)
    End Function

End Class

<Serializable()>
Public Class BTSqlDateTime
    Inherits BTSqlProperty(Of SqlDateTime, DateTime)
    Implements IBTSqlDateTime

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(column, name)

    End Sub

    ''' <summary>
    ''' Please use either the Utc or Local properties instead of Value.
    ''' </summary>
    ''' <remarks>Adding Obsolete attribute so we'll get a build warning if it's explicitly called, EditorBrowsable attribute should remove it from intellisense</remarks>
    <Obsolete("Please use either the Utc or Local properties instead of Value.")>
    <EditorBrowsable(EditorBrowsableState.Never)>
    Public Overrides ReadOnly Property Value As DateTime Implements IBTSqlDateTime.Value
        Get
            Throw New Exception("Please use either the Utc or Local properties instead of Value.")
        End Get
    End Property

    Public Overrides ReadOnly Property IsNull As Boolean
        Get
            Return Utc.IsNull
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalValue As DateTime
        Get
            Return ObjOriginal.Value
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalIsNull As Boolean
        Get
            Return ObjOriginal.IsNull
        End Get
    End Property

    Protected Overrides ReadOnly Property IsDirtyInternal As Boolean
        Get
            Return Not (Utc.Value = OriginalValue)
        End Get
    End Property

    Public Overrides Sub SetMyValue(dr As DataRow, columnName As String)
        Utc = dr.BT_CSqlDateTime(columnName)
    End Sub

    Public Overrides Sub SetMyValue(p As SqlParameter)
        Utc = DirectCast(p.Value, SqlDateTime)
    End Sub

    Protected Overrides ReadOnly Property ValueForSerialization As Object
        Get
            Return Utc
        End Get
    End Property

    Protected Overrides Sub SetMyValue(serializedObj As Object, ByVal serializedType As Type)
        If GetType(SqlDateTime) Is serializedType Then
            Utc = DirectCast(serializedObj, SqlDateTime)
        Else
            Throw New SerializationException(String.Format("Serialized type was '{0}' but property expected 'SqlDateTime'", serializedType))
        End If
    End Sub

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String
        Return String.Format("[DateTime] {0}", MyBase.ToString())
    End Function

    Public Property Utc As SqlDateTime Implements IBTSqlDateTime.Utc
        Get
            Return MyBase.Obj
        End Get
        Set(value As SqlDateTime)
            MyBase.Obj = value
        End Set
    End Property

    <Obsolete("Please use either the Utc or Local properties instead of Obj.")>
    Public Overrides Property Obj As SqlDateTime Implements IBTSqlDateTime.Obj
        Get
            Throw New Exception("Please use either the Utc or Local properties instead of Obj.")
        End Get
        Set(value As SqlDateTime)
            Throw New Exception("Please use either the Utc or Local properties instead of Obj.")
        End Set
    End Property

    Protected Overrides ReadOnly Property NullValue As SqlDateTime
        Get
            Return SqlDateTime.Null
        End Get
    End Property

    Protected Overrides Sub SetParameterValue(ByRef sp As SqlParameter)
        ' We have to do this here because for some reason Date SqlDbTypes don't get
        ' converted correctly.
        If sp.SqlDbType = SqlDbType.Date Then
            If Utc.IsNull Then
                sp.Value = SqlDateTime.Null
            Else
                sp.Value = Utc.Value
            End If
        Else
            MyBase.SetParameterValue(sp)
        End If
    End Sub

    Public Overrides Function CompareTo(otherValue As Object) As Integer
        Dim o As BTSqlDateTime = TryCast(otherValue, BTSqlDateTime)
        If o Is Nothing Then
            Return -1
        End If
        Return Utc.CompareTo(o.Utc)
    End Function
End Class

<Serializable()>
Public Class BTSqlTimeSpan
    Inherits BTSqlProperty(Of TimeSpan?, TimeSpan)
    Implements IBTSqlTimeSpan

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(column, name)

    End Sub

    Public Overrides ReadOnly Property Value As TimeSpan
        Get
            Return Obj.Value
        End Get
    End Property

    Public Overrides ReadOnly Property IsNull As Boolean
        Get
            Return Not Obj.HasValue
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalValue As TimeSpan
        Get
            Return ObjOriginal.Value
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalIsNull As Boolean
        Get
            Return Not ObjOriginal.HasValue
        End Get
    End Property

    Protected Overrides ReadOnly Property IsDirtyInternal As Boolean
        Get
            Return Not (Value = OriginalValue)
        End Get
    End Property

    Public Overrides Sub SetMyValue(dr As DataRow, columnName As String)
        Obj = dr.BT_CTimeSpan(columnName)
    End Sub

    Public Overrides Sub SetMyValue(p As SqlParameter)
        Obj = CType(p.Value, TimeSpan?)
    End Sub

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String
        Return String.Format("[TimeSpan] {0}", MyBase.ToString())
    End Function

    Protected Overrides ReadOnly Property NullValue As TimeSpan?
        Get
            Return New TimeSpan?
        End Get
    End Property

    Public Overrides Function CompareTo(otherValue As Object) As Integer
        Throw New NotImplementedException
    End Function

End Class

<Serializable()>
Public Class BTSqlBoolean
    Inherits BTSqlProperty(Of SqlBoolean, Boolean)
    Implements IBTSqlBoolean

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(column, name)

    End Sub

    Public Overrides ReadOnly Property Value As Boolean
        Get
            Return Obj.Value
        End Get
    End Property

    Public Overrides ReadOnly Property IsNull As Boolean
        Get
            Return Obj.IsNull
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalValue As Boolean
        Get
            Return ObjOriginal.Value
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalIsNull As Boolean
        Get
            Return ObjOriginal.IsNull
        End Get
    End Property

    Protected Overrides ReadOnly Property IsDirtyInternal As Boolean
        Get
            Return Not (Value = OriginalValue)
        End Get
    End Property

    Public ReadOnly Property IsTrue() As Boolean Implements IBTSqlBoolean.IsTrue
        Get
            Return Not IsNull AndAlso Value
        End Get
    End Property

    Public ReadOnly Property IsFalse() As Boolean Implements IBTSqlBoolean.IsFalse
        Get
            Return Not IsNull AndAlso Not Value
        End Get
    End Property

    Public Overrides Sub SetMyValue(dr As DataRow, columnName As String)
        Obj = dr.BT_CSqlBoolean(columnName)
    End Sub

    Public Overrides Sub SetMyValue(p As SqlParameter)
        Obj = DirectCast(p.Value, SqlBoolean)
    End Sub

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String
        Return String.Format("[Boolean] {0}", MyBase.ToString())
    End Function

    Protected Overrides ReadOnly Property NullValue As SqlBoolean
        Get
            Return SqlBoolean.Null
        End Get
    End Property

    Public Overrides Function CompareTo(otherValue As Object) As Integer
        Dim o As BTSqlBoolean = TryCast(otherValue, BTSqlBoolean)
        If o Is Nothing Then
            Return -1
        End If
        Return Obj.CompareTo(o.Obj)
    End Function

End Class

<Serializable()>
Public Class BTSqlDecimal
    Inherits BTSqlProperty(Of SqlDecimal, Decimal)
    Implements IBTSqlDecimal

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(column, name)

    End Sub

    Public Overrides ReadOnly Property Value As Decimal
        Get
            Return Obj.Value
        End Get
    End Property

    Public Overrides ReadOnly Property IsNull As Boolean
        Get
            Return Obj.IsNull
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalValue As Decimal
        Get
            Return ObjOriginal.Value
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalIsNull As Boolean
        Get
            Return ObjOriginal.IsNull
        End Get
    End Property

    Protected Overrides ReadOnly Property IsDirtyInternal As Boolean
        Get
            Return Not (Value = OriginalValue)
        End Get
    End Property

    Public Overrides Sub SetMyValue(dr As DataRow, columnName As String)
        Obj = dr.BT_CSqlDecimal(columnName)
    End Sub

    Public Overrides Sub SetMyValue(p As SqlParameter)
        Obj = DirectCast(p.Value, SqlDecimal)
    End Sub

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String
        Return String.Format("[Decimal] {0}", MyBase.ToString())
    End Function

    Protected Overrides ReadOnly Property NullValue As SqlDecimal
        Get
            Return SqlDecimal.Null
        End Get
    End Property

    Public Overrides Function CompareTo(otherValue As Object) As Integer
        Dim o As BTSqlDecimal = TryCast(otherValue, BTSqlDecimal)
        If o Is Nothing Then
            Return -1
        End If
        Return Obj.CompareTo(o.Obj)
    End Function

End Class

<Serializable()>
Public Class BTSqlMoney
    Inherits BTSqlProperty(Of SqlMoney, Decimal)
    Implements IBTSqlMoney

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(column, name)

    End Sub

    Public Overrides ReadOnly Property Value As Decimal
        Get
            Return Obj.Value
        End Get
    End Property

    Public Overrides ReadOnly Property IsNull As Boolean
        Get
            Return Obj.IsNull
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalValue As Decimal
        Get
            Return ObjOriginal.Value
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalIsNull As Boolean
        Get
            Return ObjOriginal.IsNull
        End Get
    End Property

    Protected Overrides ReadOnly Property IsDirtyInternal As Boolean
        Get
            Return Not (Value = OriginalValue)
        End Get
    End Property

    Public Overrides Sub SetMyValue(dr As DataRow, columnName As String)
        Obj = dr.BT_CSqlMoney(columnName)
    End Sub

    Public Overrides Sub SetMyValue(p As SqlParameter)
        Obj = DirectCast(p.Value, SqlMoney)
    End Sub

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String
        Return String.Format("[Money] {0}", MyBase.ToString())
    End Function

    Protected Overrides ReadOnly Property NullValue As SqlMoney
        Get
            Return SqlMoney.Null
        End Get
    End Property

    Public Overrides Function CompareTo(otherValue As Object) As Integer
        Dim o As BTSqlMoney = TryCast(otherValue, BTSqlMoney)
        If o Is Nothing Then
            Return -1
        End If
        Return Obj.CompareTo(o.Obj)
    End Function

End Class

<Serializable()>
Public Class BTSqlDouble
    Inherits BTSqlProperty(Of SqlDouble, Double)
    Implements IBTSqlDouble

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(column, name)

    End Sub

    Public Overrides ReadOnly Property Value As Double
        Get
            Return Obj.Value
        End Get
    End Property

    Public Overrides ReadOnly Property IsNull As Boolean
        Get
            Return Obj.IsNull
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalValue As Double
        Get
            Return ObjOriginal.Value
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalIsNull As Boolean
        Get
            Return ObjOriginal.IsNull
        End Get
    End Property

    Protected Overrides ReadOnly Property IsDirtyInternal As Boolean
        Get
            Return Not (Value = OriginalValue)
        End Get
    End Property

    Public Overrides Sub SetMyValue(dr As DataRow, columnName As String)
        Obj = dr.BT_CSqlDouble(columnName)
    End Sub

    Public Overrides Sub SetMyValue(p As SqlParameter)
        Obj = DirectCast(p.Value, SqlDouble)
    End Sub

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String
        Return String.Format("[Double] {0}", MyBase.ToString())
    End Function

    Protected Overrides ReadOnly Property NullValue As SqlDouble
        Get
            Return SqlDouble.Null
        End Get
    End Property

    Public Overrides Function CompareTo(otherValue As Object) As Integer
        Dim o As BTSqlDouble = TryCast(otherValue, BTSqlDouble)
        If o Is Nothing Then
            Return -1
        End If
        Return Obj.CompareTo(o.Obj)
    End Function

End Class

<Serializable()>
Public Class BTSqlGuid
    Inherits BTSqlProperty(Of SqlGuid, Guid)
    Implements IBTSqlGuid

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(column, name)

    End Sub

    Public Overrides ReadOnly Property Value As Guid
        Get
            Return Obj.Value
        End Get
    End Property

    Public Overrides ReadOnly Property IsNull As Boolean
        Get
            Return Obj.IsNull
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalValue As Guid
        Get
            Return ObjOriginal.Value
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalIsNull As Boolean
        Get
            Return ObjOriginal.IsNull
        End Get
    End Property

    Protected Overrides ReadOnly Property IsDirtyInternal As Boolean
        Get
            Return Not (Value = OriginalValue)
        End Get
    End Property

    Public Overrides Sub SetMyValue(dr As DataRow, columnName As String)
        Obj = dr.BT_CSqlGuid(columnName)
    End Sub

    Public Overrides Sub SetMyValue(p As SqlParameter)
        Obj = DirectCast(p.Value, SqlGuid)
    End Sub

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String
        Return String.Format("[Guid] {0}", MyBase.ToString())
    End Function

    Protected Overrides ReadOnly Property NullValue As SqlGuid
        Get
            Return SqlGuid.Null
        End Get
    End Property

    Public Overrides Function CompareTo(otherValue As Object) As Integer
        Dim o As BTSqlGuid = TryCast(otherValue, BTSqlGuid)
        If o Is Nothing Then
            Return -1
        End If
        Return Obj.CompareTo(o.Obj)
    End Function

End Class

<Serializable()>
Public Class BTSqlByte
    Inherits BTSqlProperty(Of SqlByte, Byte)
    Implements IBTSqlByte

    Public Sub New(ByVal column As IBTSqlColumn, ByVal name As String)
        MyBase.New(column, name)

    End Sub

    Public Overrides ReadOnly Property Value As Byte
        Get
            Return Obj.Value
        End Get
    End Property

    Public Overrides ReadOnly Property IsNull As Boolean
        Get
            Return Obj.IsNull
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalValue As Byte
        Get
            Return ObjOriginal.Value
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalIsNull As Boolean
        Get
            Return ObjOriginal.IsNull
        End Get
    End Property

    Protected Overrides ReadOnly Property IsDirtyInternal As Boolean
        Get
            Return Not (Value = OriginalValue)
        End Get
    End Property

    Public Overrides Sub SetMyValue(dr As DataRow, columnName As String)
        Obj = dr.BT_CSqlByte(columnName)
    End Sub

    Public Overrides Sub SetMyValue(p As SqlParameter)
        Obj = DirectCast(p.Value, SqlByte)
    End Sub

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String
        Return String.Format("[Byte] {0}", MyBase.ToString())
    End Function

    Protected Overrides ReadOnly Property NullValue As SqlByte
        Get
            Return SqlByte.Null
        End Get
    End Property

    Public Overrides Function CompareTo(otherValue As Object) As Integer
        Dim o As BTSqlByte = TryCast(otherValue, BTSqlByte)
        If o Is Nothing Then
            Return -1
        End If
        Return Obj.CompareTo(o.Obj)
    End Function

End Class

<Serializable()>
Public Class BTSqlEnum(Of TEnum As {Structure, IConvertible, IComparable, IFormattable})
    Inherits BTSqlProperty(Of TEnum?, TEnum)
    Implements IBTSqlEnum(Of TEnum)

    Private ReadOnly _dependentField As ISqlFieldInfo
    Private _dependentProperty As BTSqlInt32

    Public Sub New(ByVal dependentField As ISqlFieldInfo)
        MyBase.New(dependentField.Column, dependentField.FieldName)
        _dependentField = dependentField
    End Sub

    Private ReadOnly Property DependentProperty As BTSqlInt32
        Get
            If _dependentProperty Is Nothing Then
                _dependentField.CreateMeIfNecessary()
                If Not TypeOf _dependentField.SqlField Is BTSqlInt32 Then
                    Throw New SqlTypeException(String.Format("The underlying dependent SqlField must be of type BTSqlInt32 not {0}.", _dependentField.SqlField.GetType))
                End If
                _dependentProperty = DirectCast(_dependentField.SqlField, BTSqlInt32)
            End If
            Return _dependentProperty
        End Get
    End Property

    Public Overrides Property Obj As TEnum?
        Get
            If DependentProperty.IsNull Then
                Return NullValue
            End If
            Return ParseIntegerToTEnum(DependentProperty.Value)
        End Get
        Set(value As TEnum?)
            If value.HasValue Then
                DependentProperty.Obj = CInt(CType(value.Value, Object)).ToSqlInt32
            Else
                DependentProperty.Obj = SqlInt32.Null
            End If
        End Set
    End Property

    Public Overrides ReadOnly Property IsPopulated As Boolean
        Get
            Return DependentProperty.IsPopulated
        End Get
    End Property

    Public Overrides ReadOnly Property Value As TEnum
        Get
            Return Obj.Value
        End Get
    End Property

    Public Overrides ReadOnly Property IsNull As Boolean
        Get
            Return DependentProperty.IsNull
        End Get
    End Property

    Protected Overrides ReadOnly Property IsDirtyInternal As Boolean
        Get
            Return Not DependentProperty.IsDirty
        End Get
    End Property

    Protected Overrides ReadOnly Property NullValue As TEnum?
        Get
            Return Nothing
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalIsNull As Boolean
        Get
            Return Not DependentProperty.OriginalIsNull
        End Get
    End Property

    Public Overrides ReadOnly Property OriginalValue As TEnum
        Get
            Return ParseIntegerToTEnum(DependentProperty.OriginalValue)
        End Get
    End Property

    Public Overloads Overrides Sub SetMyValue(dr As DataRow, columnName As String)
        DependentProperty.SetMyValue(dr, columnName)
    End Sub

    Public Overloads Overrides Sub SetMyValue(p As SqlParameter)
        DependentProperty.SetMyValue(p)
    End Sub

    Private Function ParseIntegerToTEnum(ByVal v As Integer) As TEnum
        Dim eType As Type = GetType(TEnum)
        Dim eVal As TEnum
        If Not [Enum].TryParse(v.ToString, eVal) Then
            Throw New Exception(String.Format("{0} could not be cast to {1}.", v, eType.ToString))
        ElseIf Not ([Enum].IsDefined(eType, DependentProperty.Value) OrElse eType.IsDefined(GetType(FlagsAttribute), False)) Then
            Throw New Exception(String.Format("{0} is not defined as a {1}.", v, eType.ToString))
        End If
        Return eVal
    End Function

    <Obsolete("For debugging only.  Should not be used in code.")>
    Public Overrides Function ToString() As String
        Return String.Format("[BTSqlEnum: {0}] {1}", GetType(TEnum), MyBase.ToString())
    End Function

    Public Overrides Function CompareTo(otherValue As Object) As Integer
        Throw New NotImplementedException
    End Function

End Class

#End Region