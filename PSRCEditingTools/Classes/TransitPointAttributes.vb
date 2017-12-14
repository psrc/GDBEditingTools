Imports ESRI.ArcGIS.Geodatabase
Public Class TransitPointAttributes
    Private transitPointFeature As IFeature
    Public Sub New(ByVal TransitPoint As IFeature)
        transitPointFeature = TransitPoint



    End Sub
    Public Property PSRCJunctionID() As Long

        Get
            Dim intPos As Integer
            intPos = transitPointFeature.Fields.FindField("PSRCJunctID")
            PSRCJunctionID = transitPointFeature.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = transitPointFeature.Fields.FindField("PSRCJunctID")
            transitPointFeature.Value(intPos) = value
            transitPointFeature.Store()

        End Set
    End Property

End Class

