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
Imports ESRI.ArcGIS.ADF
Public Class frmNodeCount
    Public m_application As IApplication
    Private m_mxDoc As IMxDocument
    Private m_activeView As IActiveView
    Public m_map As IMap
    Public m_PointFeatureClass As IFeatureLayer
    Public m_LineFeatureClass As IFeatureLayer
    
    Public m_Workspace As Workspace

    Private Sub frmNodeCount_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        m_mxDoc = m_application.Document
        m_map = m_mxDoc.FocusMap
        m_activeView = m_mxDoc.FocusMap
        Dim I As Integer
        For I = 0 To m_map.LayerCount - 1

            ComboBox1.Items.Add(m_map.Layer(I).Name)
            ComboBox2.Items.Add(m_map.Layer(I).Name)
            
        Next I
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Using pComReleaser As ComReleaser = New ComReleaser
            Dim I As Integer
            For I = 0 To m_map.LayerCount - 1


                If m_map.Layer(I).Name = ComboBox1.Text Then
                    m_PointFeatureClass = m_map.Layer(I)
                    pComReleaser.ManageLifetime(m_PointFeatureClass)

                End If
                If m_map.Layer(I).Name = ComboBox2.Text Then
                    m_LineFeatureClass = m_map.Layer(I)
                    pComReleaser.ManageLifetime(m_LineFeatureClass)
                End If
            Next


            Dim pWorkSpace As IWorkspace
            'Set pWorkSpace = pPointFeatureclass.FeatureDataset
            Dim pSpatialFilter As ISpatialFilter
            pSpatialFilter = New SpatialFilter
            Dim pSpatialFilter2 As ISpatialFilter
            'Set pSpatialFilter2 = New SpatialFilter
            'Set the join type
            pSpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            'pSpatialFilter2.SpatialRel = esriSpatialRelTouches

            Dim pSpatialCursor As IFeatureCursor
            Dim pSpatialFeature As IFeature

            'Get a cursor pointing to the features in the first layer

            Dim pFeatureCursor As IFeatureCursor
            pFeatureCursor = m_PointFeatureClass.Search(Nothing, False)
            Dim pFeature As IFeature
            Dim pPrevFeature As IFeature
            Dim pLineFeature As IFeature
            Dim pCurve As ICurve
            Dim pTopoOp As ITopologicalOperator
            Dim pPoint As IPoint
            Dim pHitTest As IHitTest
            Dim pPolyLine As IPolyline
            Dim phit As IPoint, bhit As Boolean, dDist As Double, lPartIndex As Long, lSegIndex As Long, bright As Boolean, penv As IEnvelope
            lPartIndex = 0
            lSegIndex = 0
            bright = False
            Dim pPolygon As IPolygon


            Dim lFeatureCount As Long
            Dim x As Long


            Dim lRow As Long
            lRow = 1
            'Loop through the features in layer 1
            lFeatureCount = 0
            x = 0
            pFeature = pFeatureCursor.NextFeature

            While Not pFeature Is Nothing

                pTopoOp = pFeature.Shape
                pPolygon = pTopoOp.Buffer(0.3)
                With pSpatialFilter
                    .Geometry = pPolygon
                    .GeometryField = m_PointFeatureClass.FeatureClass.ShapeFieldName
                    .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

                End With

                'Assign the shape of the current point feature to the spatial query
                'Set pSpatialFilter.Geometry = pFeature.Shape
                'Get a cursor of all polygons intersecting the point
                pSpatialCursor = m_LineFeatureClass.Search(pSpatialFilter, False)
                pLineFeature = pSpatialCursor.NextFeature
                lFeatureCount = 0
                While Not pLineFeature Is Nothing
                    phit = New ESRI.ArcGIS.Geometry.Point

                    'Set pCurve = pLineFeature
                    pPoint = pFeature.Shape
                    pPolyLine = pLineFeature.Shape
                    pHitTest = pPolyLine
                    bhit = pHitTest.HitTest(pPoint, 0.3, esriGeometryHitPartType.esriGeometryPartEndpoint, phit, dDist, lPartIndex, lSegIndex, bright)

                    If bhit Then
                        lFeatureCount = lFeatureCount + 1
                    End If

                    pLineFeature = pSpatialCursor.NextFeature

                End While

                pFeature.Value(pFeature.Fields.FindField("LinkCount")) = lFeatureCount
                pFeature.Store()





                'Set pTopoOp = pPoint
                'Set pPolygon = pTopoOp.Buffer(0.3)
                '
                'With pSpatialFilter2
                '.Geometry = pPolygon
                '.SpatialRel = esriSpatialRelIntersects
                'End With








                ' Dim pData As IDataStatistics
                'Set pData = New DataStatistics
                'pData.Field = "res"
                'Set pData.Cursor = pSpatialCursor

                'Dim pField As IField
                'Set pField = pPolyFeatureclass.Fields.FindField("test")

                'pFeature.Value(pPointFeatureclass.FindField("test")) = pData.Statistics.Sum
                'pFeature.Store




                'Dim pFeatSelection As ISelectionSet
                'Set pFeatSelection = pPolyFeatureclass.Select(pSpatialFilter, esriSelectionTypeHybrid, esriSelectionOptionNormal, pWorkSpace)

                'Do some work here that suits your needs  This will count the features.



                'lFeatureCount = lFeatureCount + 1
                'If x + 5000 < lFeatureCount Then
                'WriteLogLine CStr(lFeatureCount)
                'x = lFeatureCount
                'End If




                'Debug.Print lFeatureCount


                pFeature = pFeatureCursor.NextFeature
            End While
        End Using

    End Sub

    Private Sub Label1_Click(sender As System.Object, e As System.EventArgs) Handles Label1.Click

    End Sub
End Class