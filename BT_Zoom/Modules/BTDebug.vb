Imports BT_Zoom.Enums.BTDebug
Imports BT_Zoom.Interfaces

Public Class BTDebug

    Public Shared Sub WriteLine(outputType As BTDebugOutputTypes, message As String)

        'NOTE: You can assign a bit flag using the bitwise OR operator to combine more than one type of output to appear in the Output Debug window
        Dim configuredOutputType As BTDebugOutputTypes = BTDebugOutputTypes.None  'TODO: make this configurable somewhere

        If Not configuredOutputType.HasFlag(outputType) Then
            Exit Sub
        End If

        System.Diagnostics.Debug.WriteLine(message)

    End Sub

End Class