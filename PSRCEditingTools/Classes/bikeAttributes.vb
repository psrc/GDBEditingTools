Imports ESRI.ArcGIS.Geodatabase
Public Class bikeAttributes
    Private bikeAttributeRow As IRow
    Public Sub New(ByVal row As Irow)
        bikeAttributeRow = row
    End Sub


    Public ReadOnly Property OID() As Long
        Get
            'Dim intPos As Integer
            'intPos = modeAttributeRow.Fields.FindField("PSRCEDGEID")
            OID = bikeAttributeRow.Value(0)
        End Get
    End Property


    Public Property PSRCEDGEID() As Long
        Get
            Dim intPos As Integer
            intPos = bikeAttributeRow.Fields.FindField("PSRCEDGEID")
            PSRCEDGEID = bikeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = bikeAttributeRow.Fields.FindField("PSRCEDGEID")
            bikeAttributeRow.Value(intPos) = value
            bikeAttributeRow.Store()
        End Set
    End Property
    Public Property SurfaceType() As Long
        Get
            Dim intPos As Integer
            intPos = bikeAttributeRow.Fields.FindField("SurfaceType")
            SurfaceType = bikeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = bikeAttributeRow.Fields.FindField("SurfaceType")
            bikeAttributeRow.Value(intPos) = value
            bikeAttributeRow.Store()
        End Set
    End Property
    Public Property Width() As Long
        Get
            Dim intPos As Integer
            intPos = bikeAttributeRow.Fields.FindField("Width")
            Width = bikeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = bikeAttributeRow.Fields.FindField("Width")
            bikeAttributeRow.Value(intPos) = value
            bikeAttributeRow.Store()
        End Set
    End Property
    Public Property Status() As Long
        Get
            Dim intPos As Integer
            intPos = bikeAttributeRow.Fields.FindField("Status")
            Width = bikeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = bikeAttributeRow.Fields.FindField("Status")
            bikeAttributeRow.Value(intPos) = value
            bikeAttributeRow.Store()
        End Set
    End Property
End Class
