
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

Public Class GetNearestEdge
    Private m_InGeom As IGeometry
    Private m_EdgeFC As IFeatureClass
    Private m_Distance As Long

   
    Private m_NearestEdge As IFeature


    Public Sub New(ByVal InGeom As IGeometry, ByVal InEdgeFC As IFeatureClass)
        ' Set the property value.

        Me.m_EdgeFC = InEdgeFC
        Me.m_InGeom = InGeom
    End Sub

   
   

    Public ReadOnly Property NearestEdge() As IFeature
        Get
            ' Gets the property value.
            Return m_NearestEdge
        End Get

    End Property

    
         
    Public Sub getNearestEdge()
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
        m_NearestEdge = pEdgeFeature


    End Sub


End Class
