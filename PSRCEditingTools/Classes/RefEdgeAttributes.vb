Imports ESRI.ArcGIS.Geodatabase
Public Class RefEdgeAttributes

    Private refEdgeFeature As IFeature

    Public Sub New(ByVal ReferenceEdgeFeature As IFeature)

        refEdgeFeature = ReferenceEdgeFeature

    End Sub

    Public ReadOnly Property EdgeID() As Long
        Get
            Dim intPosEdgeID As Integer
            intPosEdgeID = refEdgeFeature.Fields.FindField("PSRCEdgeID")
            EdgeID = refEdgeFeature.Value(intPosEdgeID)
        End Get
        
    End Property

    Public Property INode() As Long
        Get
            Dim intPosInode As Integer
            intPosInode = refEdgeFeature.Fields.FindField("INode")
            INode = refEdgeFeature.Value(intPosInode)
        End Get
        Set(ByVal value As Long)
            Dim intPosInode As Integer
            intPosInode = refEdgeFeature.Fields.FindField("INode")
            refEdgeFeature.Value(intPosInode) = value
            refEdgeFeature.Store()
        End Set
    End Property

    Public Property JNode() As Long
        Get
            Dim intPosJnode As Integer
            intPosJnode = refEdgeFeature.Fields.FindField("JNode")
            JNode = refEdgeFeature.Value(intPosJnode)
        End Get
        Set(ByVal value As Long)
            Dim intPosJnode As Integer
            intPosjnode = refEdgeFeature.Fields.FindField("JNode")
            refEdgeFeature.Value(intPosJnode) = value
            refEdgeFeature.Store()
        End Set
    End Property
    Public ReadOnly Property HasINode() As Boolean
        Get
            If Me.INode = 0 Then
                Return False
            Else
                Return True
            End If
        End Get

    End Property

    Public ReadOnly Property HasJNode() As Boolean
        Get
            If Me.JNode = 0 Then
                Return False
            Else
                Return True
            End If
        End Get

    End Property

   

End Class
