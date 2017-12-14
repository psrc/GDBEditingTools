Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI



Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.DataSourcesGDB
Imports Microsoft.VisualBasic

Public Class GetEndPoints
    Private m_InGeom As IGeometry
    Private m_InFC As IFeatureClass
    Private m_EdgeFC As IFeatureClass
    Private xValue As Double
    Private yValue As Double
    Private m_Distance As Double



    Public Sub New(ByVal InGeom As IGeometry, ByVal InFC As IFeatureClass)
        ' Set the property value.
        Me.m_InGeom = InGeom
        Me.m_InFC = InFC

    End Sub

    Public ReadOnly Property x() As Double
        Get
            ' Gets the property value.
            Return xValue
        End Get

    End Property
    Public ReadOnly Property y() As Double
        Get
            ' Gets the property value.
            Return yValue
        End Get

    End Property
    Public ReadOnly Property Distance() As Double
        Get
            ' Gets the property value.
            Return m_Distance
        End Get

    End Property

    

    Public Sub GetNearestFeature()
        Static pFtrIdx As IFeatureIndex2
        Static pIdxQry As IIndexQuery2
        Dim lFtrID As Long
        Dim dDis As Double
        Dim pPoint As IPoint
        Dim pFeature As IFeature



        ' Rebuild the index if necessarry
        'If pFtrIdx Is Nothing Or bRebuildIdx Then
        pIdxQry = Nothing
        pFtrIdx = Nothing
        pFtrIdx = New FeatureIndex
        With pFtrIdx
            .FeatureClass = m_InFC
            .FeatureCursor = Nothing
            .Index(Nothing, Nothing)
        End With
        pIdxQry = pFtrIdx
        'End If

        ' Get the ObjectID and Distance of the nearest feature
        pIdxQry.NearestFeature(m_InGeom, lFtrID, m_Distance)
        If lFtrID > 0 Then

            ' Set the output parameters
            pFeature = m_InFC.GetFeature(lFtrID)
            pPoint = pFeature.Shape
            xValue = pPoint.X
            yValue = pPoint.Y
            'dOutDis = dDis
            pIdxQry = Nothing
            pFtrIdx = Nothing
            GC.Collect()
        Else
            xValue = 0
            yValue = 0
        End If

    End Sub
    Public Sub GetNearestEdge()
        Static pFtrIdx As IFeatureIndex2
        Static pIdxQry As IIndexQuery2
        Dim lFtrID As Long
        Dim dDis As Double
        Dim pPoint As IPoint
        Dim pFeature As IFeature
        Dim indexINode As Long
        Dim indexJNode As Long
        indexINode = m_EdgeFC.FindField("Inode")
        indexINode = m_EdgeFC.FindField("Jnode")

        Dim lngInode As Long
        Dim lngJnode As Long
        Dim pCurve As ICurve
        Dim pGeomBag As IGeometryBag



        ' Rebuild the index if necessarry
        'If pFtrIdx Is Nothing Or bRebuildIdx Then
        pIdxQry = Nothing
        pFtrIdx = Nothing
        pFtrIdx = New FeatureIndex
        With pFtrIdx
            .FeatureClass = m_EdgeFC
            .FeatureCursor = Nothing
            .Index(Nothing, Nothing)
        End With
        pIdxQry = pFtrIdx
        'End If

        ' Get the ObjectID and Distance of the nearest feature (edge in this case)
        pIdxQry.NearestFeature(m_InGeom, lFtrID, m_Distance)
        'get the feature
        Dim pEdgeFeature As IFeature
        pEdgeFeature = m_EdgeFC.GetFeature(lFtrID)



    End Sub


End Class
