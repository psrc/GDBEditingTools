Imports ESRI.ArcGIS.Geodatabase
Public Class EdgeAttributes
    Private edgeFeature As IFeature
    Public Sub New(ByVal TransRefEdge As IFeature)
        edgeFeature = TransRefEdge



    End Sub

    Public Property PSRCEdgeID() As Long

        Get
            Dim intPos As Integer
            intPos = edgeFeature.Fields.FindField("PSRCEdgeID")
            PSRCEdgeID = edgeFeature.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = edgeFeature.Fields.FindField("PSRCEdgeID")
            edgeFeature.Value(intPos) = value
            edgeFeature.Store()

        End Set
    End Property

End Class
