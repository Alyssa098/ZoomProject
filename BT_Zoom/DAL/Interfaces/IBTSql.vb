Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Text
Imports BT_Zoom.Delegates
Imports BT_Zoom.Enums.BTDebug
Imports BT_Zoom.Enums.BTSql

Namespace Interfaces

    Public Interface IBTSqlParameter
        Inherits ISqlSelectExpression

        Property Parameter As SqlParameter
        Function ToString() As String
    End Interface

    Public Interface ICanUpdateOwnerForPaging
        Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String))
    End Interface

    Public Interface IHasDependencies
        Function GetDependencyIdentifiers() As List(Of String)
        Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression)
    End Interface

    Public Interface IHasParameters
        Function GetParameters() As List(Of IBTSqlParameter)
        ReadOnly Property Parameters As IEnumerable(Of SqlParameter)
    End Interface

    Public Interface ISqlAssignable
        Inherits ISqlExpression

        Function RenderForAssignment() As String
    End Interface

    Public Interface ISqlExpression
        Function Render() As String
    End Interface

    Public Interface ISqlOrderByExpression
        Inherits ISqlExpression, IHasDependencies

        Property Direction As DirectionTypes
    End Interface

    Public Interface ISqlSelectExpression
        Inherits ISqlAssignable, IHasDependencies, ICloneable, ICanUpdateOwnerForPaging, IHasParameters

        Property [Alias] As String
        Function WithAlias(ByVal [alias] As String) As ISqlSelectExpression
        Function GetDataRowColumnName() As String
        Function RenderForFilter() As String
        Function RenderForGroupBy() As String
        Function RenderForOrderBy() As String
        Function RenderForFunction() As String
        Function RenderForJoin() As String

    End Interface

    Public Interface IBTSqlOrderByGroup
        Inherits ISqlOrderByColumnExpression

        Property AdditionalOrderByList As List(Of ISqlOrderByExpression)
        ReadOnly Property PrimaryOrderByExpression As ISqlOrderByExpression
        Sub AddAdditionalOrderBy(ByVal column As ISqlSelectExpression, Optional ByVal direction As DirectionTypes = DirectionTypes.DefaultOrMatchPrimary)
        Sub AddAdditionalOrderByAlwaysAscending(ByVal column As ISqlSelectExpression)
        Sub AddAdditionalOrderByAlwaysDescending(ByVal column As ISqlSelectExpression)
        Function ToString() As String

    End Interface

    Public Interface IBTSqlTransformBase
        Inherits ISqlSelectExpression

        Function ApplyTransform(ByVal dr As DataRow, ByVal instance As IBTList) As String
        Function ApplyTransform(ByVal dr As DataRow) As String

        ReadOnly Property ExpressionList As List(Of ISqlSelectExpression)
        ReadOnly Property Type As Type
        ReadOnly Property SortExpressionGroup As IBTSqlOrderByGroup
        Function ToString() As String

    End Interface

    Public Interface IBTSqlTransform
        Inherits IBTSqlTransformBase

    End Interface

    Public Interface IBTSqlRelatedDataTransform
        Inherits IBTSqlTransformBase

        WriteOnly Property RelatedData As Dictionary(Of String, IBTList)

    End Interface

    Public Interface IBTSqlNewColumnTransform
        Inherits IBTSqlTransform

    End Interface

    Public Interface IBTSqlNewRelatedDataColumnTransform
        Inherits IBTSqlRelatedDataTransform

    End Interface

    Public Interface IBTList
        Inherits IDisposable

        ReadOnly Property Data As DataTable
        ReadOnly Property RelatedData As Dictionary(Of String, IBTList)
        ReadOnly Property TotalRowsAllPages As Integer
        ReadOnly Property LastRowReturned As Integer
        ReadOnly Property FirstRowReturned As Integer
        ReadOnly Property InfiniteScrollStatus As InfiniteScrollStatusTypes
        ReadOnly Property InfiniteScrollIsLoaded As Boolean
        Property MaxRowsThreshold As Integer?
        Property LogMaxRowsThreshold As Integer?
        Property IsDistinct As Boolean
        Property PagingSortDirection As DirectionTypes
        ReadOnly Property SelectList As List(Of ISqlSelectExpression)
        ReadOnly Property ColumnAliasMappings As List(Of IBTSqlColumnBase)
        ReadOnly Property SelectDictionary As Dictionary(Of String, ISqlSelectExpression)

        ReadOnly Property Parameters As IEnumerable(Of SqlParameter)
        Function AddParameter(ByVal param As IBTSqlParameter) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As Integer) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As Integer?) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlInt32) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As Boolean) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As Long) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of Integer)) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of String)) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of Long)) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As DateTime) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlDateTime) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As String, Optional varcharLength As Integer = 0) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As Decimal) As IBTSqlParameter
        Sub AddParameters(ByVal lstParams As List(Of SqlParameter))
        Sub AddParameters(ByVal lstParams As List(Of IBTSqlParameter))
        Function ContainsParameter(ByVal paramName As String) As Boolean

        Function LoadPage(ByVal pageNumber As Integer, ByVal pageSize As Integer) As Integer
        Function LoadRows(ByVal rowStartNum As Integer, ByVal rowEndNum As Integer) As Integer
        Function LoadRowNum(ByVal entityId As Integer, ByVal entityIdColumn As ISqlSelectExpression) As Integer?
        Function LoadInfiniteScroll(ByVal beginRowNum As Integer, ByVal endRowNum As Integer, ByVal lastDisplayedEntityId As Integer, ByVal entityIdColumn As ISqlSelectExpression) As Integer
        Function LoadAll() As Integer
        Function LoadTop(ByVal topNumToLoad As Integer) As Integer
        Function LoadCount() As Integer
        Sub RemoveInfiniteScrollColumns()

        ReadOnly Property Sql As String
        ReadOnly Property SqlForWeb As String
        ReadOnly Property DependencyIdentifiers As List(Of String)
        Function BuildSql(ByVal isCount As Boolean, ByVal topNumToLoad As Integer?, ByVal pageNumber As Integer?, ByVal pageSize As Integer?, ByVal rowStartNum As Integer?, ByVal rowEndNum As Integer?, ByVal entityId As Integer?, ByVal entityIdExpression As ISqlSelectExpression) As String
        ReadOnly Property OrderedSelectList As List(Of ISqlSelectExpression)
        Sub AddSelect(ByVal ParamArray expressions() As ISqlSelectExpression)
        Sub AddSelect(ByVal fnType As FunctionTypes, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddSelect(ByVal fnType As FunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddSelect(ByVal fn As IBTSqlFunction, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddSelect(ByVal fn As IBTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression)

        ReadOnly Property Transforms As List(Of IBTSqlTransformBase)
        Sub AddTransform(ByVal ParamArray transforms() As IBTSqlTransformBase)
        Sub AddTransform(ByVal transform As ApplySqlTransformationDelegate, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal sortExpressions As IBTSqlOrderByGroup, Optional ByVal type As Type = Nothing)
        Sub AddTransform(ByVal transform As ApplySqlTransformationDelegate, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), Optional ByVal type As Type = Nothing)
        Sub AddTransform(ByVal transform As ApplySqlTransformationRelatedDataDelegate, ByVal loadData As LoadRelatedDataForTransform, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression), ByVal sortExpressions As IBTSqlOrderByGroup)
        Sub AddTransform(ByVal transform As ApplySqlTransformationRelatedDataDelegate, ByVal loadData As LoadRelatedDataForTransform, ByVal [alias] As String, ByVal expressionList As List(Of ISqlSelectExpression))

        Function AddJoin(ByVal join As ISqlJoinExpression) As ISqlJoinExpression
        Function AddJoin(Of TJoinTable As {ISqlBaseTableExpression})(ByVal joinType As JoinTypes, ByVal table As TJoinTable, ByVal column As IBTSqlColumnBase, ByVal otherColumn As ISqlSelectExpression, ByVal ParamArray filters() As ISqlFilterExpression) As ISqlJoinExpression
        Sub AddJoins(ByVal ParamArray joinList As ISqlJoinExpression())
        Function AddJoinOuterApply(ByVal tbl As IBTSqlGenericFromTable) As ISqlJoinExpression

        Property AllowNoFilters As Boolean
        Property EnableOptionRecompile() As Boolean
        Property EnableOptionOptimizeForUnknown() As Boolean
        Property EnableOptionForceOrder() As Boolean
        Sub AddFilter(ByVal filter As ISqlFilterExpression)
        Sub AddFilters(ParamArray filterList As ISqlFilterExpression())
        Function AddFilterGroup(ByVal booleanOperator As BooleanOperatorTypes, ParamArray filterList As ISqlFilterExpression()) As IBTSqlFilterGroup
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightClause As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal logicalExpression As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal leftBetween As IBTSqlParameter, ByVal rightBetween As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectQuery, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Function AddOrFilterGroup(ByVal filterExpressions As IEnumerable(Of IBTSqlFilter), Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) As IBTSqlFilterGroup
        Function AddKeywordSearchFilter(ByVal parameterName As String, ByVal leftExpressions As List(Of ISqlSelectExpression), ByVal rightExpressions As List(Of String), Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) As IBTSqlFilterGroup
        Function AddKeywordSearchFilter(ByVal parameterName As String, ByVal leftExpression As ISqlSelectExpression, ByVal rightExpression As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) As IBTSqlFilterGroup
        Function GenerateKeywordSearchFilterGroup(ByVal parameterName As String, ByVal leftExpressions As List(Of ISqlSelectExpression), ByVal rightExpressions As List(Of String), Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) As IBTSqlFilterGroup
        Function AddFullTextKeywordSearchFilter(ByVal parameterName As String, ByVal leftExpressions As List(Of ISqlSelectExpression), ByVal rightExpressions As List(Of String), Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND, Optional ByVal partialWordSearch As Boolean = False) As IBTSqlFilterGroup
        Function GenerateFullTextKeywordSearchFilterGroup(ByVal parameterName As String, ByVal leftExpressions As List(Of ISqlSelectExpression), ByVal rightExpressions As List(Of String), Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND, Optional ByVal partialWordSearch As Boolean = False) As IBTSqlFilterGroup
        Sub AddJobsiteFilter(leftExpression As ISqlSelectExpression, filter As IJobsiteFilter)

        Sub AddGroupBy(ByVal ParamArray expressions() As ISqlGroupByExpression)
        Sub AddGroupBy(ByVal ParamArray expressions() As ISqlSelectExpression)

        Sub AddHaving(ByVal filter As ISqlFilterExpression)

        Sub AddOrderBy(ByVal orderByExpr As ISqlOrderByExpression)
        Sub AddOrderBys(ByVal ParamArray orderByList As ISqlOrderByExpression())
        Sub AddOrderBy(ByVal text As String, ByVal direction As DirectionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderBy(ByVal column As ISqlSelectExpression, ByVal direction As DirectionTypes)
        Sub AddOrderBy(ByVal fnType As FunctionTypes, ByVal direction As DirectionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderBy(ByVal fn As IBTSqlFunction, ByVal direction As DirectionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderByAscending(ByVal column As ISqlSelectExpression)
        Sub AddOrderByAscending(ByVal text As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderByAscending(ByVal fnType As FunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderByAscending(ByVal fn As IBTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderByDescending(ByVal column As ISqlSelectExpression)
        Sub AddOrderByDescending(ByVal text As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderByDescending(ByVal fnType As FunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderByDescending(ByVal fn As IBTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression)

        Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal cte As IBTSqlCteExpression(Of TOtherTable)) As ISqlCteExpression
        Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal sqlBuilder As IBTSqlSelectBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String, ByVal ParamArray unionQueries As ISqlUnionable()) As ISqlCteExpression
        Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal unionBuilder As IBTSqlUnionBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String) As ISqlCteExpression
        Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal cte As IBTSqlCteExpression(Of TOtherTable), ByVal position As Integer) As ISqlCteExpression
        Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal sqlBuilder As IBTSqlSelectBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String, ByVal position As Integer, ByVal ParamArray unionQueries As ISqlUnionable()) As ISqlCteExpression
        Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal unionBuilder As IBTSqlUnionBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String, ByVal position As Integer) As ISqlCteExpression

        Function AddTempTable(ByVal tempTbl As ISqlTempTableExpression) As ISqlTempTableExpression
        Function AddTempTable(Of T As {ISqlTempTableExpression, New})() As ISqlTempTableExpression

        Function AddAccountingIDs(fkColumn As IBTSqlColumn, tableAlias As String, columnPrefix As String) As List(Of ISqlSelectExpression)

    End Interface

    Public Interface IBTBaseList(Of TTable As {ISqlBaseTableExpression})
        Inherits IBTList

        ReadOnly Property SqlBuilder As IBTSqlSelectBuilder(Of TTable)
        ReadOnly Property From As TTable

    End Interface

    Public Interface IBTBaseEntityList(Of TEntity As {IBTBaseEntity(Of TTable, TList)}, TTable As {IBTSqlTable}, TList As {IBTBaseList(Of TTable)})
        Inherits IBTBaseList(Of TTable), IEnumerable(Of TEntity)

        Property CacheEntityList As Boolean
        Function GetEntity(ByVal idx As Integer) As TEntity

    End Interface

    Public Interface IBTBaseEntity(Of TTable As {IBTSqlTable}, TList As {IBTBaseList(Of TTable)})

        ReadOnly Property Table As TTable

        Sub Insert(ByVal ParamArray nonIdentityFieldsToPopulateAfterInsert As IBTBaseSqlProperty())
        Sub SetAddedBy(ByVal userId As Guid, ByVal timestamp As DateTime)
        Sub LoadPartial(ParamArray fields() As IBTBaseProperty)
        Sub SetList(Of TTList As IBTBaseList(Of TTable))(ByVal list As TTList)
        Sub PopulateFromDataRow(dr As DataRow, ParamArray mappings() As IBTSqlColumnBase)
        Sub Update()
        Sub SetUpdatedBy(ByVal userId As Guid, ByVal timestamp As DateTime)
        Sub SetUpdatedBy(ByVal userId As Guid)
        Sub Delete()
        Sub Save()
        ReadOnly Property IsDirty As Boolean
        ReadOnly Property IsTrackingIsDirty As Boolean
        Sub StartTrackingIsDirty()
        Sub StopTrackingIsDirty()
        Function GetDirtyProperties() As List(Of ISqlFieldInfo)
        Function GetDirtyPropertiesNames() As String
        Sub RequirePopulatedProperties(ParamArray props As IBTBaseSqlProperty())

    End Interface

    Public Interface ISqlFromExpression
        Inherits ISqlExpression, IHasDependencies

    End Interface

    Public Interface ISqlFilterExpression
        Inherits ISqlExpression, IHasDependencies, ICanUpdateOwnerForPaging

        Property IsFirstFilter As Boolean
        Property BooleanOperator As BooleanOperatorTypes

    End Interface

    Public Interface ISqlGroupByExpression
        Inherits ISqlExpression, IHasDependencies

    End Interface

    Public Interface ISqlOrderByColumnExpression
        Inherits ISqlOrderByExpression

        Property Column As ISqlSelectExpression
    End Interface

    Public Interface ISqlJoinable

        ReadOnly Property [Alias] As String
        Sub ChangeAlias(ByVal [alias] As String)
    End Interface

    Public Interface IBaseTableExpression
        Inherits ISqlFromExpression, ISqlJoinable
    End Interface

    Public Interface ISqlBaseTableExpression
        Inherits IBaseTableExpression

        Property UseDirtyRead As Boolean
    End Interface

    Public Interface ISqlWritableTableExpression
        Inherits IBaseTableExpression
    End Interface

    Public Interface ISqlTableExpression
        Inherits ISqlBaseTableExpression, ISqlWritableTableExpression
    End Interface

    Public Interface ISqlViewExpression
        Inherits ISqlBaseTableExpression
    End Interface

    Public Interface ISqlTempTableExpression
        Inherits ISqlWritableTableExpression

        Property RenderType As TempTableRenderTypes

        Function RenderDeclaration() As String
        Function RenderCleanup() As String
    End Interface

    Public Interface IBTSqlView
        Inherits ISqlViewExpression

        ReadOnly Property ViewName As String
        ReadOnly Property Schema As String
        ReadOnly Property DatabaseName As String
        ReadOnly Property AllowNoRepresentation As Boolean
        Function WithAlias(ByVal [alias] As String) As IBTSqlView
        Function AddCustomExpressionsToStar() As List(Of ISqlSelectExpression)
        Sub RemoveAutoExpressionsFromStar(autoExpressions As List(Of ISqlSelectExpression))

    End Interface

    Public Interface ISqlJoinExpression
        Inherits ISqlExpression, ISqlJoinable, IHasDependencies, ICanUpdateOwnerForPaging

        Property JoinType As JoinTypes
        ReadOnly Property OtherColumn As ISqlSelectExpression

        ''' <summary>
        '''  When false, if the join is not needed by any dependencies (i.e. select statements, filters, etc), it will be removed from the query
        ''' </summary>
        Property ShouldNotRemove As Boolean
    End Interface

#Region "SqlBuilder"

    Public Interface ISqlBuilder
        Inherits ISqlExpression, IHasParameters

        Property UseLineBreaks As Boolean
        Function RenderForCte() As String
        Overloads Sub Render(ByRef sb As StringBuilder)
        Property TopNum As Integer?

        Function AddParameter(ByVal param As IBTSqlParameter) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As Integer) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlInt32) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As Boolean) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlBoolean) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As Long) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlInt64) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of Integer)) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of Long)) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of String)) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of DateTime)) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As DateTime) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlDateTime) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As String, Optional varcharLength As Integer = 0) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlString, Optional varcharLength As Integer = 0) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As Decimal) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlDecimal) As IBTSqlParameter
        Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlGuid) As IBTSqlParameter
        Sub AddParameters(ByVal lstParams As List(Of SqlParameter))
        Sub AddParameters(ByVal lstParams As List(Of IBTSqlParameter))
        Function ContainsParameter(ByVal paramName As String) As Boolean

        Function AddCte(ByVal cteExpression As ISqlCteExpression) As ISqlCteExpression
        Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal baseSqlBuilder As IBTSqlSelectBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String, ByVal ParamArray unionQueries As ISqlUnionable()) As ISqlCteExpression
        Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal unionBuilder As IBTSqlUnionBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String) As ISqlCteExpression
        Function AddCte(ByVal cteExpression As ISqlCteExpression, ByVal position As Integer) As ISqlCteExpression
        Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal baseSqlBuilder As IBTSqlSelectBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String, ByVal position As Integer, ByVal ParamArray unionQueries As ISqlUnionable()) As ISqlCteExpression
        Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal unionBuilder As IBTSqlUnionBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String, ByVal position As Integer) As ISqlCteExpression
        Function IsCteWithSameNameAlreadyAdded(ByVal cteExpression As ISqlCteExpression) As Boolean

        Function AddTempTable(Of T As {ISqlTempTableExpression, New})() As ISqlTempTableExpression
        Function AddTempTable(ByVal tempTbl As ISqlTempTableExpression) As ISqlTempTableExpression

        Function AddJoin(ByVal join As ISqlJoinExpression) As ISqlJoinExpression
        Function AddJoinOuterApply(ByVal tbl As IBTSqlGenericFromTable) As ISqlJoinExpression
        Function AddCrossJoin(Of TJoinTable As {IBaseTableExpression})(ByVal tbl As TJoinTable) As ISqlJoinExpression
        Function AddJoin(Of TJoinTable As {IBaseTableExpression})(ByVal joinType As JoinTypes, ByVal table As TJoinTable, ByVal column As IBTSqlColumnBase, ByVal otherColumn As ISqlSelectExpression, ByVal ParamArray filters() As ISqlFilterExpression) As ISqlJoinExpression
        Function IsJoinAlreadyAdded(ByVal join As ISqlJoinExpression) As Boolean


        Sub AddFilter(ByVal filter As ISqlFilterExpression)
        Function AddFilterGroup(ByVal booleanOperator As BooleanOperatorTypes, ParamArray filters As ISqlFilterExpression()) As IBTSqlFilterGroup
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightClause As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOpertator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectQuery, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal logicalExpression As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal leftBetween As IBTSqlParameter, ByVal rightBetween As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddJobsiteFilter(leftExpression As ISqlSelectExpression, filter As IJobsiteFilter)

    End Interface

    Public Interface IBTSqlBuilderBase(Of TTable As {IBaseTableExpression})
        Inherits ISqlBuilder

    End Interface

    Public Interface ISqlSelectBuilder
        Inherits ISqlBuilder, IHasDependencies, ICanUpdateOwnerForPaging, ISqlFilterable

        ReadOnly Property SelectList As List(Of ISqlSelectExpression)
        ReadOnly Property SelectDependentList As List(Of ISqlSelectExpression)
        ReadOnly Property SelectListAppliedToPageQuery As List(Of ISqlSelectExpression)
        Sub AddSelect(ByVal ParamArray expressions() As ISqlSelectExpression)
        Sub AddSelect(ByVal fnType As FunctionTypes, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddSelect(ByVal fnType As FunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddSelect(ByVal fn As IBTSqlFunction, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddSelect(ByVal fn As IBTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddTransform(ByVal newColumnTransform As IBTSqlTransformBase)

        Function AddKeywordSearchFilter(parameterName As String, leftExpressions As List(Of ISqlSelectExpression), rightExpressions As List(Of String), Optional booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) As IBTSqlFilterGroup
        Sub AddKeywordSearchFilter(leftExpression As ISqlSelectExpression, parameter As IBTSqlParameter)
        Function GenerateKeywordSearchFilterGroup(ByVal parameterName As String, ByVal leftExpressions As List(Of ISqlSelectExpression), ByVal rightExpressions As List(Of String), Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND, Optional ByVal partialWordSearch As Boolean = False, Optional fullTextSearch As Boolean = False) As IBTSqlFilterGroup

        Sub AddGroupBy(ByVal ParamArray expressions() As ISqlGroupByExpression)
        Sub AddGroupBy(ByVal ParamArray expressions() As ISqlSelectExpression)
        Sub ClearSqlGroupBy()

        Sub AddHaving(ByVal filter As ISqlFilterExpression)
        Sub ClearSqlHaving()

        Sub AddOrderBy(ByVal orderBy As ISqlOrderByExpression)
        Sub AddOrderBy(ByVal text As String, ByVal direction As DirectionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderBy(ByVal column As ISqlSelectExpression, ByVal direction As DirectionTypes)
        Sub AddOrderBy(ByVal fnType As FunctionTypes, ByVal direction As DirectionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderBy(ByVal fn As IBTSqlFunction, ByVal direction As DirectionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderByAscending(ByVal column As ISqlSelectExpression)
        Sub AddOrderByAscending(ByVal text As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderByAscending(ByVal fnType As FunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderByAscending(ByVal fn As IBTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderByDescending(ByVal column As ISqlSelectExpression)
        Sub AddOrderByDescending(ByVal text As String, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderByDescending(ByVal fnType As FunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub AddOrderByDescending(ByVal fn As IBTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression)
        Sub ClearSqlOrderBy()

        Function GetSelectDependencyIdentifiers() As List(Of String)
        Function GetGroupByDependencyIdentifiers() As List(Of String)
        Function GetHavingDependencyIdentifiers() As List(Of String)
        Function GetOrderByDependencyIdentifiers() As List(Of String)
        Function GetFilterDependencyIdentifiers() As List(Of String)
        Function GetJoinDependencyIdentifiers() As List(Of String)

        Function GetSelectDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression)
        Function GetGroupByDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression)
        Function GetHavingDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression)
        Function GetOrderByDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression)
        Function GetFilterDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression)
        Function GetJoinDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression)

        Property IsCount As Boolean
        Property IsDistinct As Boolean
        Property ParametersNeededByBuilder As List(Of IBTSqlParameter)
        Property EntityId As Integer?
        Property PagingSortDirection As DirectionTypes
        Property EntityIdExpression As ISqlSelectExpression
        Property AllowNoFilters As Boolean
        Property EnableOptionRecompile As Boolean
        Property EnableOptionOptimizeForUnknown As Boolean
        Property EnableOptionForceOrder As Boolean
        Sub SetPagingByNumberAndSize(ByVal pageNumber As Integer?, ByVal pageSize As Integer?)
        Sub SetPagingByRows(ByVal rowNumStart As Integer?, ByVal rowNumEnd As Integer?)

        Function AddAccountingIDs(fkColumn As IBTSqlColumn, tableAlias As String, columnPrefix As String) As List(Of ISqlSelectExpression)

    End Interface

    Public Interface IBTSqlSelectBuilder(Of TTable As {IBaseTableExpression})
        Inherits IBTSqlBuilderBase(Of TTable), ISqlSelectBuilder

        ReadOnly Property From As TTable

    End Interface

    Public Interface ISqlUpdateBuilder
        Inherits ISqlBuilder

        Sub Execute()

        ReadOnly Property SetList As List(Of Tuple(Of IBTSqlColumn, ISqlAssignable))
        Sub AddSet(column As IBTSqlColumn, rightExpression As ISqlAssignable)

        Property AllowNoFilters As Boolean

    End Interface

    Public Interface IBTSqlUpdateBuilder(Of TTable As {IBaseTableExpression})
        Inherits IBTSqlBuilderBase(Of TTable), ISqlUpdateBuilder

        ReadOnly Property From As TTable

    End Interface

    Public Interface ISqlInsertBuilder
        Inherits ISqlBuilder

        Sub Execute()

        ReadOnly Property ColumnList As List(Of IBTSqlColumn)
        Sub AddColumns(ByVal ParamArray columns() As IBTSqlColumn)


    End Interface

    Public Interface IBTSqlInsertBuilderBase(Of TTable As {IBaseTableExpression})
        Inherits ISqlInsertBuilder, IBTSqlBuilderBase(Of TTable)

        ReadOnly Property Into As TTable

    End Interface

    Public Interface IBTSqlInsertSelectBuilder(Of TIntoTable As {IBaseTableExpression}, TSelectTable As {IBaseTableExpression})
        Inherits IBTSqlInsertBuilderBase(Of TIntoTable)

        ReadOnly Property SelectQuery As IBTSqlSelectBuilder(Of TSelectTable)

    End Interface

    Public Interface IBTSqlInsertValuesBuilder(Of TIntoTable As {IBaseTableExpression})
        Inherits IBTSqlInsertBuilderBase(Of TIntoTable)

        ReadOnly Property Values As List(Of List(Of ISqlSelectExpression))
        Sub AddValues(ByVal ParamArray expressions() As ISqlSelectExpression)

        Sub AddOutputFromInsert(ByVal outputParams As List(Of Tuple(Of IBTSqlParameter, IBTSqlColumn)))
        Sub AddOutputFromInsert(ByVal outputParams As List(Of Tuple(Of IBTSqlParameter, IBTSqlColumn)), ByVal asTable As Boolean)

    End Interface

    Public Interface IBTSqlUnionBuilder(Of TTable As {IBaseTableExpression})
        Inherits IBTSqlBuilderBase(Of TTable)

        ReadOnly Property BaseSelectList As List(Of ISqlSelectExpression)

        ReadOnly Property From As TTable

        Sub AddUnion(unionQuery As ISqlUnionable)
        Sub AddUnions(ByVal unionQueries As ISqlUnionable())
        Sub ClearUnionList()

    End Interface

    Public Interface ISqlDeleteBuilder
        Inherits ISqlBuilder

        Function Execute() As Integer

        Property AllowNoFilters As Boolean

    End Interface

    Public Interface IBTSqlDeleteBuilder(Of TTable As {ISqlWritableTableExpression})
        Inherits IBTSqlBuilderBase(Of TTable), ISqlDeleteBuilder

        ReadOnly Property From As TTable

    End Interface

#End Region

    Public Interface ISqlSelectQuery
        Inherits ISqlAssignable, IHasDependencies, ICloneable, ICanUpdateOwnerForPaging, ISqlFilterable, ISqlUnionable

        ReadOnly Property IsTableParameterTable As Boolean

    End Interface

    Public Interface ISqlFilterable
        Function RenderForFilter() As String

    End Interface

    Public Interface ISqlCteExpression
        Inherits ISqlTableExpression, ISqlCteColumnFinder

        ReadOnly Property CteName As String
        Function RenderForCte(ByVal isFirst As Boolean) As String
    End Interface

    Public Interface IBTSqlColumnBase
        Inherits ISqlSelectExpression

        ReadOnly Property Name As String
        ReadOnly Property OwnerTable As IBaseTableExpression
        Function ToString() As String
        Function RenderOwner() As String
        Function CloneWithoutAlias() As Object

    End Interface

    Public Interface IBTSqlColumn
        Inherits IBTSqlColumnBase

        ReadOnly Property Table As IBaseTableExpression
        ReadOnly Property SqlDbType As SqlDbType
        ReadOnly Property Size As Integer
        ReadOnly Property IsNullable As Boolean
        ReadOnly Property Precision As Byte
        ReadOnly Property Scale As Byte
        ReadOnly Property IsPrimaryKey As Boolean
        ReadOnly Property IsIdentity As Boolean
        Function CreateSqlParameter(ByVal direction As ParameterDirection) As SqlParameter
        Function CreateSqlParameterForInsert() As SqlParameter
        Function CreateSqlParameter() As SqlParameter
        Function RenderForTable() As String
        Function ColumnDataType() As String

    End Interface

    Public Interface IBTSqlCteColumn
        Inherits IBTSqlColumnBase

    End Interface

    Public Interface ISqlCteColumnFinder
        Function CteColumn(ByVal column As ISqlSelectExpression) As IBTSqlCteColumn
    End Interface

    Public Interface ISqlUnionable
        Inherits IHasDependencies

        Function RenderForUnion() As String
    End Interface

    Public Interface IBTSqlTableBase
        Inherits IBaseTableExpression

        ReadOnly Property TableName As String
        Function WithAlias(ByVal [alias] As String) As IBTSqlTableBase
        Function ToString() As String
    End Interface

    Public Interface IBTSqlTable
        Inherits IBTSqlTableBase, ISqlTableExpression

        ReadOnly Property Schema As String
        ReadOnly Property DatabaseName As String
        ReadOnly Property AllowNoRepresentation As Boolean
        Property IndexToUse As String

        Sub ChangeDatabase(databaseName As String)

        Function AddCustomExpressionsToStar() As List(Of ISqlSelectExpression)
        Sub RemoveAutoExpressionsFromStar(autoExpressions As List(Of ISqlSelectExpression))
        Sub AddAliasesWithPrefix(expressions As List(Of ISqlSelectExpression))

    End Interface

    Public Interface IBTSqlFunction

        ReadOnly Property Name As String
        ReadOnly Property Owner As String
        ReadOnly Property IsScalar As Boolean

    End Interface

    Public Interface IBTSqlGenericFromTable
        Inherits IBTSqlTable, ICanUpdateOwnerForPaging, ISqlUnionable

        Property SubQuery As ISqlSelectQuery
        Function AddSelectableColumn(ByVal column As ISqlSelectExpression, Optional ByVal [alias] As String = "") As IBTSqlColumn
        Function NewColumn(ByVal columnName As String, Optional ByVal [alias] As String = "") As IBTSqlColumn
        Sub AddSelectableColumns(ByVal ParamArray columns As ISqlSelectExpression())
        Function SelectableColumn(ByVal column As ISqlSelectExpression) As ISqlSelectExpression
        Function NonSelectableColumn(ByVal columnName As String) As IBTSqlColumn
        Function NonSelectableColumn(ByVal column As ISqlSelectExpression) As IBTSqlColumn
        Function SelectStar() As IEnumerable(Of ISqlSelectExpression)
        Function SelectStar(ByVal tableAlias As String) As IEnumerable(Of ISqlSelectExpression)

    End Interface

    Public Interface IBTSqlFilter
        Inherits ISqlFilterExpression


    End Interface

    Public Interface IBTSqlComparisonFilter
        Inherits IBTSqlFilter

        Property LeftExpression As ISqlSelectExpression
        Property RightExpression As ISqlSelectExpression
        Property RightClause As String
        Property ComparisonOperator As ComparisonOperatorTypes

    End Interface

    Public Interface IBTSqlLogicalFilter
        Inherits IBTSqlFilter

        Property LeftExpression As ISqlSelectExpression
        Property RightExpression As ISqlSelectExpression
        Property RightClause As String
        Property LogicalOperator As LogicalOperatorTypes

        Sub AddDependencies(ByVal ParamArray dependencies() As ISqlSelectExpression)

    End Interface

    Public Interface IJobsiteFilter

        Property JobIDs As List(Of Integer)
        Property SubQuery As IBTSqlSelectQuery
        Property SelectColumn As ISqlSelectExpression

        ReadOnly Property ShouldUseSubQuery As Boolean
        Function ToTable(ByRef parameters As List(Of IBTSqlParameter)) As Tuple(Of IBTSqlTable, ISqlSelectExpression)


    End Interface

    Public Interface IBTSqlSelectQuery
        Inherits ISqlSelectQuery, ISqlSelectExpression

        Shadows Function RenderForFilter() As String

    End Interface

    Public Interface IBTSqlCteExpression(Of TTable As {ISqlBaseTableExpression, New})
        Inherits ISqlCteExpression

        ReadOnly Property From As TTable
        ReadOnly Property SelectList As List(Of ISqlSelectExpression)
        ReadOnly Property Parameters As IEnumerable(Of SqlParameter)
        Function Star(ParamArray columnsToExlude As ISqlSelectExpression()) As ISqlSelectExpression()
        Sub AddUnion(ParamArray unionQueries As ISqlUnionable())

    End Interface

    Public Interface IBTSqlFilterGroup
        Inherits ISqlFilterExpression

        ReadOnly Property Count As Integer
        Sub AddFilter(ByVal filter As ISqlFilterExpression)
        Function AddFilterGroup(ByVal booleanOperator As BooleanOperatorTypes, ByVal ParamArray filters As ISqlFilterExpression()) As IBTSqlFilterGroup
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightClause As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOpertator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectQuery, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal logicalExpression As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal leftBetween As IBTSqlParameter, ByVal rightBetween As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal leftBetween As ISqlSelectExpression, ByVal rightBetween As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)
        Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND)

    End Interface

    Public Interface IBTTableParameterTable
        Inherits IBTSqlTable, ICanUpdateOwnerForPaging

        Property TableParameter As IBTSqlParameter
        ReadOnly Property Value As IBTSqlColumn

    End Interface

    Public Interface IBTCteList(Of TTable As {ISqlBaseTableExpression, New})
        Inherits IBTBaseList(Of IBTSqlCteExpression(Of TTable)), ISqlCteColumnFinder

        ReadOnly Property CteTable As IBTSqlCteExpression(Of TTable)
        Sub AddUnion(ByVal sqlBuilder As ISqlSelectBuilder, ByVal unionType As UnionTypes)

    End Interface

#Region "BT Properties"

    Public Interface IBTBaseProperty

        ReadOnly Property PropertyName As String
        ReadOnly Property IsPopulated As Boolean
        ReadOnly Property Columns As List(Of IBTSqlColumn)
        <Obsolete("For debugging only.  Should not be used in code.")> Function ToString() As String

    End Interface

    Public Interface IBTDependentProperty(Of TType)
        Inherits IBTBaseProperty

        ReadOnly Property Value As TType

    End Interface

    Public Interface IBTBaseSqlProperty
        Inherits IBTBaseProperty

        Function ValueAsString() As String
        ReadOnly Property Column As IBTSqlColumn

        Sub SetMyValue(dr As DataRow, columnName As String)
        Sub SetMyValue(p As SqlParameter)
        Sub SetMyValue(serializedObj As Object, ByVal serializedType As Type)

        ReadOnly Property IsNull As Boolean
        Sub Clear(Optional ByVal showAsPopulated As Boolean = False)
        ReadOnly Property IsNew As Boolean
        Function Equals(obj As Object) As Boolean

        ReadOnly Property IsDirty As Boolean
        Sub TakeSnapshot()
        Sub ClearSnapshot()
        ReadOnly Property IsTrackingIsDirty As Boolean
        Sub StartTrackingIsDirty()
        Sub StopTrackingIsDirty()

        Function CreateSqlParameterForInsert() As SqlParameter
        Function CreateSqlParameter() As SqlParameter

        ReadOnly Property ValueForSerialization As Object

    End Interface

    Public Interface IBTSqlProperty(Of TObjType, TPrimitiveType)
        Inherits IBTBaseSqlProperty

        ReadOnly Property WasPopulated As Boolean
        ReadOnly Property ObjOriginal As TObjType
        Property Obj As TObjType

        ReadOnly Property Value As TPrimitiveType
        ReadOnly Property OriginalValue As TPrimitiveType
        ReadOnly Property OriginalIsNull As Boolean
        ReadOnly Property NullValue As TObjType

        ReadOnly Property IsNotNull As Boolean

    End Interface

    Public Interface IBTSqlInt16
        Inherits IBTSqlProperty(Of SqlInt16, Short)

    End Interface

    Public Interface IBTSqlInt32
        Inherits IBTSqlProperty(Of SqlInt32, Integer)

    End Interface

    Public Interface IBTSqlInt64
        Inherits IBTSqlProperty(Of SqlInt64, Long)

    End Interface

    Public Interface IBTSqlString
        Inherits IBTSqlProperty(Of SqlString, String)

        ReadOnly Property IsNullOrEmpty As Boolean

    End Interface

    Public Interface IBTSqlDateTime
        Inherits IBTSqlProperty(Of SqlDateTime, DateTime)

        Property Utc As SqlDateTime

        ''' <remarks>We had to put Shadows on Obj so we could mark it as obsolete.</remarks>
        <Obsolete("Please use either the Utc or Local properties instead of Obj.")>
        Shadows Property Obj As SqlDateTime

        ''' <remarks>We had to put Shadows on Value so we could mark it as obsolete.</remarks>
        <Obsolete("Please use either the Utc or Local properties instead of Value.")>
        Shadows ReadOnly Property Value As DateTime

    End Interface

    Public Interface IBTSqlTimeSpan
        Inherits IBTSqlProperty(Of TimeSpan?, TimeSpan)

    End Interface

    Public Interface IBTSqlBoolean
        Inherits IBTSqlProperty(Of SqlBoolean, Boolean)

        ReadOnly Property IsTrue() As Boolean
        ReadOnly Property IsFalse() As Boolean

    End Interface

    Public Interface IBTSqlDecimal
        Inherits IBTSqlProperty(Of SqlDecimal, Decimal)

    End Interface

    Public Interface IBTSqlMoney
        Inherits IBTSqlProperty(Of SqlMoney, Decimal)

    End Interface

    Public Interface IBTSqlDouble
        Inherits IBTSqlProperty(Of SqlDouble, Double)

    End Interface

    Public Interface IBTSqlGuid
        Inherits IBTSqlProperty(Of SqlGuid, Guid)

    End Interface

    Public Interface IBTSqlByte
        Inherits IBTSqlProperty(Of SqlByte, Byte)

    End Interface

    Public Interface IBTSqlEnum(Of TEnum As {Structure, IConvertible, IComparable, IFormattable})
        Inherits IBTSqlProperty(Of TEnum?, TEnum)

    End Interface

#End Region

#Region "FieldsInfo"

    Public Interface IFieldInfoBase

        ReadOnly Property Field As IBTBaseProperty
        Property FieldName As String
        ReadOnly Property Columns As List(Of IBTSqlColumn)
        Sub CreateMeIfNecessary()
        Function ToString() As String

    End Interface

    Public Interface ISqlFieldInfo
        Inherits IFieldInfoBase

        Property SqlField As IBTBaseSqlProperty
        Property Column As IBTSqlColumn
        Property CreateField As CreateFieldDelegate

    End Interface

    Public Interface IDependentFieldInfo(Of TType)
        Inherits IFieldInfoBase

        Property DependentField As IBTDependentProperty(Of TType)
        Property DependentFields As List(Of IFieldInfoBase)
        Property CalculateValue As OfType(Of TType).CalculateValueDelegate

    End Interface

#End Region

End Namespace