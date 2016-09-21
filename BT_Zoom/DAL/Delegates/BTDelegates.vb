Imports BT_Zoom.Interfaces

Namespace Delegates

    Public Delegate Function ApplySqlTransformationDelegate(ByVal dr As DataRow, ByVal expressionList As List(Of ISqlSelectExpression)) As String
    Public Delegate Function ApplySqlTransformationRelatedDataDelegate(ByVal dr As DataRow, ByVal expressionList As List(Of ISqlSelectExpression), ByVal relatedData As Dictionary(Of String, IBTList)) As String
    Public Delegate Function LoadRelatedDataForTransform(ByVal instance As IBTList) As Dictionary(Of String, IBTList)

    Public Delegate Function CreateFieldDelegate(column As IBTSqlColumn, fieldName As String) As IBTBaseSqlProperty

    Public Class OfType(Of TType)

        Public Delegate Function CalculateValueDelegate(dependentFields As List(Of IFieldInfoBase)) As TType

    End Class

#Region "Schedules"

    Public Delegate Function NotStartedMarkup(ByVal startDate As DateTime) As String
    Public Delegate Function CompletedMarkup(ByVal msg As String) As String

#End Region

End Namespace
