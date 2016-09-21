Option Strict On
Option Explicit On

Imports System.ComponentModel

Namespace Enums

    Namespace BTDebug

        <Flags>
        Public Enum BTDebugOutputTypes
            None = 0
            ConnectionProvider = 1
            Sql = 2
            DataTableSize = 4
            CacheOnlineOffline = 8
            CacheHitMiss = 16
            PageEvents = 32
            TimezoneOffset = 64
            GetFileOriginalPath = 128
            ZipFileDetails = 256
            BTInboxReader = 512
            WriteLogWithInfo = 1024
        End Enum

    End Namespace

    Namespace BTSql

        Public Enum DirectionTypes
            <DescriptionAttribute("ASC")> ASC = 0
            <DescriptionAttribute("DESC")> DESC = 1
            ''' <summary>
            ''' Renders as nothing unless it's part of an BTSqlOrderByGroup
            ''' </summary>
            <DescriptionAttribute("")> [DefaultOrMatchPrimary] = 2
        End Enum

        Public Enum BooleanOperatorTypes
            <DescriptionAttribute("AND")> [AND] = 0
            <DescriptionAttribute("OR")> [OR] = 1
            <DescriptionAttribute("AND NOT")> AndNot = 2
            <DescriptionAttribute("OR NOT")> OrNot = 3
        End Enum

        Public Enum ComparisonOperatorTypes
            <DescriptionAttribute("=")> [Equals] = 0
            <DescriptionAttribute("<>")> [NotEquals] = 1
            <DescriptionAttribute(">")> [GreaterThan] = 2
            <DescriptionAttribute(">=")> [GreaterThanOrEqualTo] = 3
            <DescriptionAttribute("<")> [LessThan] = 4
            <DescriptionAttribute("<=")> [LessThanOrEqualTo] = 5
        End Enum

        Public Enum LogicalOperatorTypes
            <DescriptionAttribute("IS NULL")> [IsNull] = 0
            <DescriptionAttribute("IS NOT NULL")> [IsNotNull] = 1
            <DescriptionAttribute("IN")> [In] = 2
            <DescriptionAttribute("NOT IN")> [NotIn] = 3
            <DescriptionAttribute("LIKE")> [Like] = 4
            <DescriptionAttribute("NOT LIKE")> [NotLike] = 5
            <DescriptionAttribute("BETWEEN")> [Between] = 6
            <DescriptionAttribute("NOT BETWEEN")> [NotBetween] = 7
            <DescriptionAttribute("EXISTS")> [Exists] = 8
            <DescriptionAttribute("NOT EXISTS")> [NotExists] = 9
            <DescriptionAttribute("CONTAINS")> [Contains] = 10
        End Enum

        'NOTE: These are not implemented yet
        Public Enum ComparisonOperatorModifierTypes
            <DescriptionAttribute("ANY")> [Any] = 0  'SOME is equivalent
            <DescriptionAttribute("ALL")> [All] = 1
        End Enum

        Public Enum JoinTypes
            <DescriptionAttribute("INNER JOIN")> [InnerJoin] = 0
            <DescriptionAttribute("LEFT OUTER JOIN")> [LeftOuterJoin] = 1
            <DescriptionAttribute("RIGHT OUTER JOIN")> [RightOuterJoin] = 2
            <DescriptionAttribute("FULL OUTER JOIN")> [FullOuterJoin] = 3
            <DescriptionAttribute("CROSS JOIN")> [CrossJoin] = 4
            <DescriptionAttribute("OUTER APPLY")> [OuterApply] = 5
            <DescriptionAttribute("CROSS APPLY")> [CrossApply] = 6
            <DescriptionAttribute("INNER MERGE JOIN")> [InnerMergeJoin] = 7
        End Enum

        'NOTE: add other functions to this as needed
        Public Enum FunctionTypes
            <DescriptionAttribute("COUNT")> [Count] = 0
            <DescriptionAttribute("SUM")> [Sum] = 1
            <DescriptionAttribute("MAX")> [Max] = 2
            <DescriptionAttribute("MIN")> [Min] = 3
            <DescriptionAttribute("AVG")> [Avg] = 4
            <DescriptionAttribute("ISNULL")> [IsNull] = 5
            <DescriptionAttribute("CAST")> [Cast] = 6
            <DescriptionAttribute("COALESCE")> [Coalesce] = 8
            <DescriptionAttribute("LEN")> [Len] = 9
            <DescriptionAttribute("GETUTCDATE")> [GetUtcDate] = 10
            <DescriptionAttribute("NULLIF")> [NullIf] = 11
            <Description("LTRIM")> [LeftTrim] = 12
            <Description("RTRIM")> [RightTrim] = 13
            <Description("DISTINCT")> [Distinct] = 14
            <Description("STUFF")> [Stuff] = 15
            <Description("SUBSTRING")> [Substring] = 16
            <Description("POWER")> [Power] = 17

        End Enum

        Public Enum DataTypeFunctionTypes
            <DescriptionAttribute("CONVERT")> [Convert] = 1
        End Enum

        Public Enum DataTypes
            <DescriptionAttribute("INT")> [INT] = 0
            <DescriptionAttribute("DECIMAL(9,2)")> [DECIMAL92] = 1
            <DescriptionAttribute("FLOAT")> FLOAT = 2
            <DescriptionAttribute("VARCHAR(max)")> VARCHARMAX = 3
        End Enum

        Public Enum DatePartFunctionTypes
            <DescriptionAttribute("DATEADD")> [DateAdd] = 0
            <DescriptionAttribute("DATEPART")> [DatePart] = 1
            <DescriptionAttribute("DATEDIFF")> [DateDiff] = 2
        End Enum

        Public Enum DatePartTypes
            <DescriptionAttribute("year")> [Year] = 0
            <DescriptionAttribute("quarter")> [Quarter] = 1
            <DescriptionAttribute("month")> [Month] = 2
            <DescriptionAttribute("dayofyear")> [DayOfYear] = 3
            <DescriptionAttribute("day")> [Day] = 4
            <DescriptionAttribute("week")> [Week] = 5
            <DescriptionAttribute("weekday")> [WeekDay] = 6
            <DescriptionAttribute("hour")> [Hour] = 7
            <DescriptionAttribute("minute")> [Minute] = 8
            <DescriptionAttribute("second")> [Second] = 9
            <DescriptionAttribute("millisecond")> [Millisecond] = 10
            <DescriptionAttribute("microsecond")> [Microsecond] = 11
            <DescriptionAttribute("nanosecond")> [Nanosecond] = 12
        End Enum

        Public Enum ArithmeticFunctionTypes
            <DescriptionAttribute("-")> [Subtract] = 0
            <DescriptionAttribute("+")> [Add] = 1
            <DescriptionAttribute("*")> [Multiply] = 2
            <DescriptionAttribute("/")> [Divide] = 3
            <DescriptionAttribute("&")> [BooleanAnd] = 4
            <DescriptionAttribute("|")> [BooleanOr] = 5

        End Enum

        Public Enum UnionTypes
            <Description("UNION")> Union = 0
            <Description("UNION ALL")> UnionAll = 1
        End Enum

        ''' <remarks>If another status is added here, we need to make sure the mobile team knows about it and it is backward compatible</remarks>
        Public Enum InfiniteScrollStatusTypes
            ''' <summary>
            ''' Last identifier found, the page matching the row numbers that were supplied is returned
            ''' </summary>
            DataMatch = 0

            ''' <summary>
            ''' The last identifier cannot be found, the first page is returned
            ''' </summary>
            ResetToStart = 1

            ''' <summary>
            ''' The last identifier sent does not match, the location of the last identifier is found
            ''' and the next page following the new location is returned
            ''' </summary>
            DataChanged = 2
        End Enum

        <Flags>
        Public Enum TempTableRenderTypes
            ''' <summary>
            ''' Renders as a From table only. Doesn't get declared/created or dropped.
            ''' </summary>
            None = 0

            ''' <summary>
            ''' CREATE TABLE #temptable ([...])
            ''' </summary>
            Create = 1

            ''' <summary>
            ''' DROP TABLE #temptable
            ''' </summary>
            ''' <remarks></remarks>
            Drop = 2

            ''' <summary>
            ''' Created and Dropped sql is rendered with the sql builder
            ''' </summary>
            Both = Create Or Drop
        End Enum

    End Namespace

    Namespace Videos
        Public Enum HostTypes
            'Just a sample list. Feel free to add to it.
            Youtube = 1
            Vimeo = 2
        End Enum
    End Namespace

End Namespace
