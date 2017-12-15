Imports ESRI.ArcGIS.Geodatabase
Public Class ProjectAttributes
    Private projectFeature As IFeature
    Public Sub New(ByVal ProjectLine As IFeature)
        projectFeature = ProjectLine



    End Sub
    Public Property PROJRTEID() As Long

        Get
            Dim intPos As Integer
            intPos = projectFeature.Fields.FindField("PROJRTEID")
            PROJRTEID = projectFeature.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = projectFeature.Fields.FindField("PROJRTEID")
            projectFeature.Value(intPos) = value
            projectFeature.Store()
        End Set
    End Property
    Public Property ProjectINode() As Long

        Get
            Dim intPos As Integer
            intPos = projectFeature.Fields.FindField("INode")
            ProjectINode = projectFeature.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = projectFeature.Fields.FindField("INode")
            projectFeature.Value(intPos) = value
            projectFeature.Store()

        End Set

    End Property

    Public Property ProjectJNode() As Long

        Get
            Dim intPos As Integer
            intPos = projectFeature.Fields.FindField("JNode")
            ProjectJNode = projectFeature.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = projectFeature.Fields.FindField("JNode")
            projectFeature.Value(intPos) = value
            projectFeature.Store()

        End Set

    End Property
    Public ReadOnly Property HasINode() As Boolean
        Get
            If Me.ProjectINode = 0 Then
                Return False
            Else
                Return True
            End If
        End Get

    End Property

    Public ReadOnly Property HasJNode() As Boolean
        Get
            If Me.ProjectJNode = 0 Then
                Return False
            Else
                Return True
            End If
        End Get

    End Property
End Class
