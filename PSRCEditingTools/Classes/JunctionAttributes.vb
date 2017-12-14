Imports ESRI.ArcGIS.Geodatabase


Public Class JunctionAttributes

    Private junctFeature As IFeature
    Public Sub New(ByVal JunctionFeature As IFeature)
        junctFeature = JunctionFeature



    End Sub

    Public Property PSRCJunctID() As Long

        Get
            Dim intPos As Integer
            intPos = junctFeature.Fields.FindField("PSRCJunctID")
            PSRCJunctID = junctFeature.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = junctFeature.Fields.FindField("PSRCJunctID")
            junctFeature.Value(intPos) = value
            junctFeature.Store()

        End Set
    End Property

    Public ReadOnly Property Emme2nodeID() As Long

        Get
            Dim intPos As Integer
            intPos = junctFeature.Fields.FindField("Emme2nodeID")
            PSRCJunctID = junctFeature.Value(intPos)
        End Get
       
    End Property
    Public ReadOnly Property Emme2nodeLabel() As String

        Get
            Dim intPos As Integer
            intPos = junctFeature.Fields.FindField("Emme2nodeLabel")
            PSRCJunctID = junctFeature.Value(intPos)
        End Get

    End Property

    Public ReadOnly Property EMME2Dir() As String


        Get
            Dim intPos As Integer
            intPos = junctFeature.Fields.FindField("Emme2Dir")
            PSRCJunctID = junctFeature.Value(intPos)
        End Get

    End Property

    Public ReadOnly Property EMME2HOV() As String


        Get
            Dim intPos As Integer
            intPos = junctFeature.Fields.FindField("EMME2HOV")
            PSRCJunctID = junctFeature.Value(intPos)
        End Get

    End Property

End Class
