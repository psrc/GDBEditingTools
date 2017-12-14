Imports ESRI.ArcGIS.Geodatabase
Public Class TurnAttributes
    Private turnFeature As IFeature

    Public Sub New(ByVal TurnMovementFeature As IFeature)

        turnFeature = TurnMovementFeature

    End Sub

    Public ReadOnly Property TurnID() As Long
        Get
            Dim intTurnID As Integer
            intTurnID = turnFeature.Fields.FindField("TurnID")
            TurnID = turnFeature.Value(intTurnID)
        End Get

    End Property

    Public Property PSRCJunctID() As Long
        Get
            Dim intPosJunctID As Integer
            intPosJunctID = turnFeature.Fields.FindField("PSRCJunctID")
            PSRCJunctID = turnFeature.Value(intPosJunctID)
        End Get
        Set(ByVal value As Long)
            Dim intPosJunctID As Integer
            intPosJunctID = turnFeature.Fields.FindField("PSRCJunctID")
            turnFeature.Value(intPosJunctID) = value
            turnFeature.Store()
        End Set
    End Property
    Public Property FrEdgeID() As Long
        Get
            Dim intPosFrEdgeID As Integer
            intPosFrEdgeID = turnFeature.Fields.FindField("FrEdgeID")
            FrEdgeID = turnFeature.Value(intPosFrEdgeID)
        End Get
        Set(ByVal value As Long)
            Dim intPosFrEdgeID As Integer
            intPosFrEdgeID = turnFeature.Fields.FindField("FrEdgeID")
            turnFeature.Value(intPosFrEdgeID) = value
            turnFeature.Store()
        End Set
    End Property
    Public Property ToEdgeID() As Long
        Get
            Dim intPosToEdgeID As Integer
            intPosToEdgeID = turnFeature.Fields.FindField("T0EdgeID")
            ToEdgeID = turnFeature.Value(intPosToEdgeID)
        End Get
        Set(ByVal value As Long)
            Dim intPosToEdgeID As Integer
            intPosToEdgeID = turnFeature.Fields.FindField("ToEdgeID")
            turnFeature.Value(intPosToEdgeID) = value
            turnFeature.Store()
        End Set
    End Property
End Class
