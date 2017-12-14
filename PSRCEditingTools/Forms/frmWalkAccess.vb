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
Public Class frmWalkAccess
    Public m_application As IApplication
    Private m_mxDoc As IMxDocument
    Private m_activeView As IActiveView
    Public m_map As IMap
    Public pFLTaz As IFeatureLayer
    Public pFLBlocks As IFeatureLayer
    Public pFLBusStops As IFeatureLayer
    Public pFLTempStops As IFeatureLayer
    Public g_FWS As IFeatureWorkspace
    Public g_workSpaceEdit As IWorkspaceEdit

    Public m_Workspace As Workspace


    
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim i As Integer
        For i = 0 To m_map.LayerCount - 1

            If m_map.Layer(i).Name = cboTAZ.Text Then
                pFLTaz = m_map.Layer(i)
            End If
            If m_map.Layer(i).Name = cboBlocks.Text Then
                pFLBlocks = m_map.Layer(i)
            End If
            If m_map.Layer(i).Name = cboBusStops.Text Then
                pFLBusStops = m_map.Layer(i)
            End If
            If m_map.Layer(i).Name = cboTempStops.Text Then
                pFLTempStops = m_map.Layer(i)
            End If
        Next i
        Dim pDataset As IDataset
        pDataset = pFLTaz.FeatureClass


        m_Workspace = pDataset.Workspace
        g_workSpaceEdit = m_Workspace



        PercentWalkToTransit(500)
    End Sub

    Private Sub frmWalkAccess_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        m_mxDoc = m_application.Document
        m_map = m_mxDoc.FocusMap
        m_activeView = m_mxDoc.FocusMap
        Dim I As Integer
        For I = 0 To m_map.LayerCount - 1

            cboBusStops.Items.Add(m_map.Layer(I).Name)
            cboBlocks.Items.Add(m_map.Layer(I).Name)
            cboTAZ.Items.Add(m_map.Layer(I).Name)
            cboTempStops.Items.Add(m_map.Layer(I).Name)

        Next I
    End Sub
    Private Sub PercentWalkToTransit(ByVal dist As Double)
        Try

        

            Dim pTAZfeat As IFeature
            Dim pFeatCursor As IFeatureCursor
            Dim pFeatCursor2 As IFeatureCursor
            Dim pFeatCursor3 As IFeatureCursor
            Dim pFeatCursor4 As IFeatureCursor
            'Dim pFeatCursor As IFeatureCursor
            Dim pFeature As IFeature
            Dim pFSel As IFeatureSelection
            pFSel = pFLTaz

            Dim pGeometryCollection As IGeometryCollection

            Dim pFilter As IQueryFilter
            pFilter = New QueryFilter
            'pFilter.WhereClause = "intBlkGrpID = " & PointID
            'Set pFeatCursor = TAZ.FeatureClass.Search(pFilter, False)

            'Set pTAZfeat = pFeatCursor.NextFeature

            Dim pPolygon As IPolygon
            Dim pSpatialFilter As ISpatialFilter
            Dim pWorkspaceEdit As IWorkspaceEdit
            pWorkspaceEdit = m_Workspace
            Dim pTopoOp As ITopologicalOperator
            'Set pTopoOp = pTAZfeat.Shape
            'pTopoOp.Buffer (dist)
            'Set pPolygon = New Polygon
            'Set pPolygon = pTopoOp
            'Set pSpatialFilter = New SpatialFilter
            'With pSpatialFilter
            ' Set .Geometry = pPolygon
            '.GeometryField = pFLBusStops.FeatureClass.ShapeFieldName
            '.SpatialRel = esriSpatialRelIntersects
            'End With
            Dim indexTAZID As Long
            Dim indexTransitWalk As Long
            Dim curTAZ As Integer
            Dim pQFilter As IQueryFilter
            Dim pBlockFeat As IFeature

            Dim OID As Long







            If pFSel.SelectionSet.Count = 0 Then
                ' do all of 'em

                pFeatCursor = pFLTaz.FeatureClass.Search(Nothing, False)
            Else
                ' just do selected ones
                'lCount = Centroids.FeatureClass.FeatureCount(Nothing)
                pFSel.SelectionSet.Search(Nothing, False, pFeatCursor)
            End If
            pTAZfeat = pFeatCursor.NextFeature

            'loop through each Buffered TAZ
            Do Until pTAZfeat Is Nothing
                '
                'Set pTopoOp = pTAZfeat.Shape

                'pTopoOp.Buffer (dist)
                'Set pPolygon = New Polygon
                'Set pPolygon = pTopoOp
                pSpatialFilter = New SpatialFilter
                With pSpatialFilter
                    .Geometry = pTAZfeat.Shape
                    .GeometryField = pFLBusStops.FeatureClass.ShapeFieldName
                    .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects

                End With
                indexTAZID = pTAZfeat.Fields.FindField("TAZ")
                curTAZ = pTAZfeat.Value(indexTAZID)
                OID = pTAZfeat.Value(0)

                'select bus stops within buffered TAZ, if there are no bus stops, move on to next TAZ
                If pFLBusStops.FeatureClass.FeatureCount(pSpatialFilter) > 0 Then

                    pFeatCursor2 = pFLBusStops.FeatureClass.Search(pSpatialFilter, False)
                    'deleteFeatures pFLTempStops.FeatureClass
                    deleteAllRows(g_workSpaceEdit, pFLTempStops.FeatureClass)
                    LoadBufferedTransitPoints(pFLTempStops.FeatureClass, pFeatCursor2, CDbl(txtBusBuffer.Text))

                    'Now Select the Blocks from the Current TAZ that intersect with the buffered bus stops


                    pQFilter = New QueryFilter
                    pQFilter.WhereClause = "New_TAZ = " & curTAZ
                    'get all the buffered stops
                    pFeatCursor3 = pFLTempStops.Search(Nothing, False)
                    'add them to a geometry bag
                    pFeature = pFeatCursor3.NextFeature

                    'use the gometry bag for the spatial query
                    pGeometryCollection = New GeometryBag
                    Do Until pFeature Is Nothing
                        pGeometryCollection.AddGeometry(pFeature.Shape)
                        pFeature = pFeatCursor3.NextFeature
                    Loop

                    'new spatial filter
                    pSpatialFilter = New SpatialFilter
                    With pSpatialFilter
                        .Geometry = pGeometryCollection
                        .GeometryField = pFLBlocks.FeatureClass.ShapeFieldName
                        '   only want to select from the current TAZ
                        .WhereClause = "New_TAZ = " & curTAZ
                        .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                    End With

                    pFeatCursor4 = pFLBlocks.FeatureClass.Search(pSpatialFilter, False)
                    pBlockFeat = pFeatCursor4.NextFeature
                    Dim x As Integer
                    Do While Not pBlockFeat Is Nothing
                        indexTransitWalk = pFLBlocks.FeatureClass.FindField("TransitWalk")
                        pBlockFeat.Value(indexTransitWalk) = 1
                        pBlockFeat.Store()

                        pBlockFeat = pFeatCursor4.NextFeature

                    Loop
                    pBlockFeat = Nothing
                    pSpatialFilter = Nothing
                    pQFilter = Nothing
                    pGeometryCollection = Nothing
                    pFeatCursor3 = Nothing
                    pFeatCursor4 = Nothing
                  

                   
                    GC.Collect()

                    x = x + 1

                End If
                Debug.Print("finished TAZ " & curTAZ)
                pTAZfeat = pFeatCursor.NextFeature




                'MsgBox x
                'Exit Sub

                WriteLogLine(curTAZ & " " & OID)
            Loop




        Catch ex As Exception
            MessageBox.Show(ex.ToString)


        End Try











    End Sub
    Public Sub deleteFeatures(ByVal fc As IFeatureClass)
        Dim pFCurs As IFeatureCursor
        pFCurs = fc.Search(Nothing, False)
        Dim pFeat As IFeature
        pFeat = pFCurs.NextFeature
        Do While Not pFeat Is Nothing
            pFeat.Delete()


            pFeat = pFCurs.NextFeature
        Loop

    End Sub
    Public Function deleteAllRows(ByVal pWSedit As IWorkspaceEdit, ByVal pFCls As IFeatureClass)

        'WriteLogLine "start deleting all rows: " & Now()
        Debug.Print("start deleting all rows: " & Now())
        Dim pFCS As IFeatureCursor

        Dim pFtSet As ESRI.ArcGIS.esriSystem.Set
        Dim I As Integer
        Dim pFtEdit As IFeatureEdit, pFt As IFeature
        pFtSet = New ESRI.ArcGIS.esriSystem.Set
        I = 0
        pFCS = pFCls.Search(Nothing, False)

        Dim pDaStats As IDataStatistics
        Dim pStats As IStatisticsResults
        Dim lMax As Long, lMin As Long, l As Long
        Dim pFilt As IQueryFilter2

        pDaStats = New DataStatistics
        pDaStats.Field = pFCls.OIDFieldName
        pDaStats.Cursor = pFCS
        pDaStats.SimpleStats = True
        pStats = pDaStats.Statistics
        If pFCls.FeatureCount(Nothing) = 0 Then Exit Function
        lMax = pStats.Maximum
        lMin = pStats.Minimum
        l = lMax
        pFilt = New QueryFilter

        'deleting rows in desceding order of OID
       
        GC.Collect()


        Do Until lMax <= lMin



            If lMax Mod 10000 > 0 Then
                'the first round, usually the largest OID is not a multiplication of 10000
                If lMax > 10000 Then
                    l = Val(Microsoft.VisualBasic.Left(CStr(lMax), Len(CStr(lMax)) - 4) & "0000")
                Else
                    l = 0
                End If
            Else
                l = lMax - 10000
            End If

            pFilt.WhereClause = pFCls.OIDFieldName & ">=" & CStr(l)

            If pWSedit.IsBeingEdited = False Then pWSedit.StartEditing(True)
            pWSedit.StartEditOperation()
            pFCS = pFCls.Search(pFilt, False)

            pFt = pFCS.NextFeature
            Do Until pFt Is Nothing
                pFtEdit = pFt
                pFtSet.Add(pFtEdit)
                I = I + 1
                If I = 100 Then
                    pFtEdit.DeleteSet(pFtSet)
                    pFtSet = New ESRI.ArcGIS.esriSystem.Set
                    I = 0
                    GC.Collect()

                End If
                pFt = pFCS.NextFeature
            Loop
            Marshal.FinalReleaseComObject(pFCS)

            If pFtSet.Count > 0 Then pFtEdit.DeleteSet(pFtSet)
            lMax = l

            pWSedit.StopEditOperation()
            pWSedit.StopEditing(True)
        Loop
        pStats = Nothing
        pDaStats = Nothing

        pFtSet = Nothing
        pFtEdit = Nothing

        'WriteLogLine "finish deleting all rows: " & Now()
        Debug.Print("end deleting all rows: " & Now())
        pFCS = Nothing
        pWSedit.StopEditOperation()
        pWSedit.StopEditing(True)
        Debug.Print("finish saving deleting: " & Now())



    End Function
    Public Sub LoadBufferedTransitPoints(ByVal BufferedStopsFC As IFeatureClass, ByVal FeatCurs As IFeatureCursor, ByVal buffersize As Double)
        Dim pPolygon As IPolygon
        Dim pTopoOp As ITopologicalOperator
        Dim pFields As IFields
        Dim lSFld As Long
        Dim pFeatureCursor As IFeatureCursor
        Dim pOldFeature As IFeature
        Dim pNewFeature As IFeature
        'delete all features
        Dim pFeat2 As IFeature
        Dim pFilter As ISpatialFilter
        Dim i As Integer
        pFeat2 = FeatCurs.NextFeature
        Do Until pFeat2 Is Nothing
            'buffer point
            pPolygon = New Polygon
            pTopoOp = pFeat2.Shape
            m_map.MapUnits = esriUnits.esriFeet
            pPolygon = pTopoOp.Buffer(buffersize)
            pPolygon = New Polygon
            'setpTopoOp = TransitLine.Shape
            m_map.MapUnits = esriUnits.esriFeet
            pPolygon = pTopoOp.Buffer(buffersize)
            pNewFeature = BufferedStopsFC.CreateFeature
            pNewFeature.Shape = pPolygon
            pNewFeature.Store()
            pFeat2 = FeatCurs.NextFeature
            pNewFeature.Store()
        Loop
        Marshal.FinalReleaseComObject(FeatCurs)




        GC.Collect()




        'pFeatureCursor = FullNetwork.Search(pFilter, False)
        'pOldFeature = pFeatureCursor.NextFeature

        'begin copying to PartialNetwork, one by one
        ' Do Until pOldFeature Is Nothing
        'pNewFeature = PartialNetwork.CreateFeature
        'pFields = PartialNetwork.Fields
        'pNewFeature.Shape = pOldFeature.ShapeCopy
        'pNewFeature.Store()

        'For I = 0 To pFields.FieldCount - 1
        'If (pFields.Field(I).Editable) Then
        'lSFld = pOldFeature.Fields.FindField(pFields.Field(I).Name)
        'If (lSFld <> -1) Then
        'If IsDBNull(pOldFeature.Value(lSFld)) = False Then

        'pNewFeature.Value(I) = pOldFeature.Value(lSFld)

        'End If

        'End If

        'End If
        'GC.Collect()
        'Next I
        'pOldFeature = pFeatureCursor.NextFeature
        'Loop

        'Marshal.FinalReleaseComObject (pFeatureCursor)












    End Sub

End Class