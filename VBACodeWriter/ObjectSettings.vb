Public Class ObjectSettings
    Private stringObjectName As String
    'Dim frmListOfVariables As frmListOfVariables
    Private stringFormOrReport As String
    Private stringObjectType As String
    Private stringOriginalModule As String
    Friend LineNumber As Integer
    Public Property LastTableorQueryObjectType As String
    Public Property LastObjectType As String

    Public Sub SetObjectType(ByRef StringObjectTypeIn As String)
        stringObjectType = StringObjectTypeIn

    End Sub

    Public Sub SetObjectName(ByRef stringObjectNameIn As String)
        stringObjectName = stringObjectNameIn
    End Sub

    Public Function GetObjectname() As String

        Return stringObjectName
    End Function

    Public Sub SetFormOrReport(ByRef stringFormorReportIn As String)
        stringFormOrReport = stringFormorReportIn
    End Sub

    Public Function GetFormorReport() As String

        Return stringFormOrReport
    End Function
    Public Function GetObjectType() As String
        Return stringObjectType
    End Function

    Public Sub SetOriginalModule(ByRef stringOriginalModuleIn As String)
        stringOriginalModule = stringOriginalModuleIn
    End Sub
    Public Function GetOriginalModule() As String
        Return stringOriginalModule
    End Function

End Class
