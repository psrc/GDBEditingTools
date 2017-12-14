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

Public Class frmSnapStopsToEndPoints
    Public m_application As IApplication
    Private m_mxDoc As IMxDocument
    Private m_activeView As IActiveView
    Public m_map As IMap
    Public m_PartialNetworkFL As IFeatureLayer
    Public m_FullNetworkFL As IFeatureLayer
    Public m_TransitRouteFL As IFeatureLayer
    Public m_EndPointsFL As IFeatureLayer
    Public m_Workspace As Workspace
    Public m_BusStopsFL As IFeatureLayer

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub


    Private Sub frmSnapStopsToEndPoints_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        m_mxDoc = m_application.Document
        m_map = m_mxDoc.FocusMap
        m_activeView = m_mxDoc.FocusMap
        Dim I As Integer
        For I = 0 To m_map.LayerCount - 1

            ComboBox1.Items.Add(m_map.Layer(I).Name)
            ComboBox2.Items.Add(m_map.Layer(I).Name)
            ComboBox3.Items.Add(m_map.Layer(I).Name)
            ComboBox4.Items.Add(m_map.Layer(I).Name)
            ComboBox5.Items.Add(m_map.Layer(I).Name)
        Next I

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Using pComReleaser As ComReleaser = New ComReleaser
            Dim I As Integer
            For I = 0 To m_map.LayerCount - 1
                If m_map.Layer(I).Name = ComboBox1.Text Then
                    m_PartialNetworkFL = m_map.Layer(I)
                    pComReleaser.ManageLifetime(m_PartialNetworkFL)

                End If
                If m_map.Layer(I).Name = ComboBox2.Text Then
                    m_FullNetworkFL = m_map.Layer(I)
                    pComReleaser.ManageLifetime(m_FullNetworkFL)
                End If
                'If m_pMap.Layer(i).name = ComboBox3.Text Then
                'Set m_pEdgesFL = m_pMap.Layer(i)
                'End If
                If m_map.Layer(I).Name = ComboBox3.Text Then
                    m_TransitRouteFL = m_map.Layer(I)
                    pComReleaser.ManageLifetime(m_TransitRouteFL)
                End If

                If m_map.Layer(I).Name = ComboBox4.Text Then
                    m_EndPointsFL = m_map.Layer(I)
                    pComReleaser.ManageLifetime(m_EndPointsFL)
                End If

                If m_map.Layer(I).Name = ComboBox5.Text Then
                    m_BusStopsFL = m_map.Layer(I)
                    pComReleaser.ManageLifetime(m_BusStopsFL)
                End If

            Next I
            Dim pDataset As IDataset
            pDataset = m_TransitRouteFL.FeatureClass


            m_Workspace = pDataset.Workspace

            Dim x As Integer
            Dim pFilter As IQueryFilter
            Dim pTransitFeature As IFeature
            Dim pBusStopFeature As IFeature
            Dim pNearFeat As IFeature
            Dim pDist As Double
            'Dim pFilter As IQueryFilter
            Dim pFeatureCursor As IFeatureCursor
            Dim pFeatureCursor2 As IFeatureCursor
            Dim pGeom As IGeometry
            'get all transit routes
            Dim pFilt As IQueryFilter
            Dim counter As Integer
            Dim pFSel As IFeatureSelection
            pFSel = m_TransitRouteFL

            counter = 59
            'pFilt = New QueryFilter
            'pFilt.WhereClause = "NewID >= " & counter

            'pFeatureCursor = m_TransitRouteFL.FeatureClass.Search(pFilt, False)


            Dim indexRouteID As Long
            Dim indexX As Long
            Dim indexY As Long
            Dim IndexNewOrder As Integer
            Dim indexDist As Long
            Dim lngRouteID As Long
            Dim pPoint As IPoint
            Dim xcord As Double
            Dim ycord As Double
            Dim prevX As Double
            Dim prevY As Double
            Dim a As Integer
            Dim b As Integer
            Dim NearestEdge As IFeature


            If pFSel.SelectionSet.Count = 0 Then
                ' do all of 'em
                'lCount = pFLayer.FeatureClass.FeatureCount(Nothing)
                pFeatureCursor = m_TransitRouteFL.FeatureClass.Search(Nothing, False)
                pComReleaser.ManageLifetime(pFeatureCursor)
            Else
                ' just do selected ones
                'lCount = pFSel.SelectionSet.Count
                pFSel.SelectionSet.Search(Nothing, False, pFeatureCursor)
                pComReleaser.ManageLifetime(pFeatureCursor)
            End If

            pTransitFeature = pFeatureCursor.NextFeature
            'loop through all transit routes      



            Do Until pTransitFeature Is Nothing

                'a new id to link routes to potential stops 
                indexRouteID = pTransitFeature.Fields.FindField("NewID")
                lngRouteID = pTransitFeature.Value(indexRouteID)





                Dim pWorkspaceEdit As IWorkspaceEdit
                pWorkspaceEdit = m_Workspace
                pComReleaser.ManageLifetime(pWorkspaceEdit)
                If m_PartialNetworkFL.FeatureClass.FeatureCount(Nothing) > 0 Then
                    deleteAllRows0(pWorkspaceEdit, m_PartialNetworkFL.FeatureClass)
                End If
                'deleteAllRows0(pWorkspaceEdit, m_PartialNetworkFL.FeatureClass)
                'deleteAllRows0(pWorkspaceEdit, m_EndPointsFL.FeatureClass)

                Try


                    'find the underlying edges and write them to a Feature Class
                    LoadNetworkEdges2(m_PartialNetworkFL.FeatureClass, m_Workspace, pTransitFeature, m_FullNetworkFL.FeatureClass)
                    'get the endpoints of each individual edge, put them into m_EndPointsFL
                    'LoadEndPoints(m_PartialNetworkFL.FeatureClass, m_EndPointsFL.FeatureClass, m_Workspace)
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString)


                End Try

                'start locating bus stops
                pFilter = New QueryFilter
                pFilter.WhereClause = "NewID = " & lngRouteID
                'pFeatureCursor2 = m_BusStopsFL.FeatureClass.Update(pFilter, True)

                Dim pTableSort As ITableSort
                pTableSort = New TableSort
                pComReleaser.ManageLifetime(pTableSort)
                'Dim pQueryFilter As IQueryFilter
                'pQueryFilter = New QueryFilter

                'pQueryFilter.WhereClause = "[STATE_NAME] like  'A*'"

                With pTableSort
                    .Fields = "ORDER_"
                    .Ascending("ORDER_") = True
                    '.Ascending("count_name ") = True
                    '.CaseSensitive("state_name") = True
                    '.CaseSensitive("count_name ") = True
                    .QueryFilter = pFilter
                    .Table = m_BusStopsFL.FeatureClass
                End With

                pTableSort.Sort(Nothing)

                Dim pCursor As ICursor

                pCursor = pTableSort.Rows
                pComReleaser.ManageLifetime(pCursor)

                Dim pRow As IRow
                pRow = pCursor.NextRow
                a = 1
                b = 1
                'pBusStopFeature = pFeatureCursor2.NextFeature
                Dim nearEdge As IFeature
                While Not pRow Is Nothing

                    Try
                        Dim intOID As Integer
                        Dim OID As Long
                        Dim pCurve As ICurve
                        intOID = pRow.Fields.FindField("OBJECTID")
                        OID = pRow.Value(intOID)


                        'pGeom = pBusStopFeature.Shape
                        'want to get nearest edge to bus stop


                        'then want to get the nearest end point of that edge
                        pGeom = m_BusStopsFL.FeatureClass.GetFeature(OID).Shape
                        Dim myNearEdge As New GetNearestEdge(pGeom, m_PartialNetworkFL.FeatureClass)
                        pComReleaser.ManageLifetime(myNearEdge)
                        myNearEdge.getNearestEdge()
                        nearEdge = myNearEdge.NearestEdge
                        'we have the nearest edge to the bus stop. Get the from and to points from this edge
                        'LoadEndPoints2(nearEdge, m_EndPointsFL.FeatureClass, m_Workspace)
                        Dim endPoints As ESRI.ArcGIS.Geometry.IPoint
                        endPoints = New ESRI.ArcGIS.Geometry.Point



                        ClosestEdgeEndPoint(nearEdge, pGeom, m_Workspace, endPoints)

                        'Dim endPoints As New GetEndPoints(pGeom, m_BusStopsFL.FeatureClass)
                        'endPoints.GetNearestFeature()
                        'GetNearestFeature(pGeom, m_EndPointsFL.FeatureClass, Nothing, True, pNearFeat, xcord, ycord)


                        indexX = pRow.Fields.FindField("x")
                        indexY = pRow.Fields.FindField("y")
                        indexDist = pRow.Fields.FindField("Distance")
                        IndexNewOrder = pRow.Fields.FindField("NewOrder")

                        'If pWorkspaceEdit.IsBeingEdited = False Then pWorkspaceEdit.StartEditing(True)
                        'pWorkspaceEdit.StartEditOperation()

                        'pBusStopFeature.Value(indexX) = endPoints.x
                        'pFeatureCursor2.UpdateFeature(pBusStopFeature)
                        'pBusStopFeature.Value(indexY) = endPoints.y
                        'pFeatureCursor2.UpdateFeature(pBusStopFeature)

                        'Assign new Bus Stop order
                        '   pRow.Value(IndexNewOrder) = a


                        'check to see if stop is 1st or last 
                        pCurve = pTransitFeature.Shape
                        '1st stop
                        If a = 1 Then
                            pRow.Value(IndexNewOrder) = b
                            pRow.Store()
                            pCurve = pTransitFeature.Shape
                            pRow.Value(indexX) = pCurve.FromPoint.X
                            pRow.Store()
                            pRow.Value(indexY) = pCurve.FromPoint.Y
                            pRow.Store()
                            prevX = pCurve.FromPoint.X
                            prevY = pCurve.FromPoint.Y

                            'last stop
                        ElseIf a = pTableSort.Table.RowCount(pFilter) Then
                            'check to see if the last stop has the same coords

                            If prevX = pCurve.ToPoint.X And prevY = pCurve.ToPoint.Y Then
                                b = b - 1
                                pRow.Value(IndexNewOrder) = -2

                                pRow.Store()
                                ' b = a - 1


                            Else

                                pRow.Value(indexX) = pCurve.ToPoint.X
                                pRow.Store()
                                pRow.Value(indexY) = pCurve.ToPoint.Y
                                pRow.Store()
                                pRow.Value(IndexNewOrder) = b
                                pRow.Store()

                            End If
                            'Other stops
                        Else
                            If prevX = endPoints.X And prevY = endPoints.Y Then
                                b = b - 1
                                pRow.Value(IndexNewOrder) = -1
                                pRow.Store()
                                'pRow.Value(indexDist) = endPoints.Distance
                                pRow.Store()
                            Else
                                pRow.Value(indexX) = endPoints.X
                                pRow.Store()
                                pRow.Value(indexY) = endPoints.Y
                                pRow.Store()
                                pRow.Value(IndexNewOrder) = b
                                pRow.Store()
                                prevX = endPoints.X
                                prevY = endPoints.Y
                                'pRow.Value(indexDist) = endPoints.Distance
                                pRow.Store()
                            End If


                            'pWorkspaceEdit.StopEditOperation()
                            'pWorkspaceEdit.StopEditing(True)


                            endPoints = Nothing

                        End If
                        'GC.Collect()

                        a = a + 1
                        b = b + 1
                        'pBusStopFeature = pFeatureCursor2.NextFeature
                        pRow = pCursor.NextRow

                    Catch ex As Exception
                        MessageBox.Show(ex.Message.ToString)

                    End Try
                End While
                Marshal.FinalReleaseComObject(pCursor)


                pTransitFeature = pFeatureCursor.NextFeature
                counter = counter + 1
                'GC.Collect()


            Loop
        End Using
    End Sub
    Public Sub LoadNetworkEdges(ByVal PartialNetwork As IFeatureClass, ByVal featWorkspace As IWorkspace, ByVal TransitLine As IFeature, ByVal FullNetwork As IFeatureClass)
        'delete all features


        'get underlying edges under TransitLine
        Dim pFields As IFields
        Dim lSFld As Long
        Dim pFeatureCursor As IFeatureCursor
        Dim pOldFeature As IFeature
        Dim pNewFeature As IFeature
        Dim pPolygon As IPolygon
        Dim pTopoOp As ITopologicalOperator
        Dim pFilter As ISpatialFilter
        Dim I As Integer
        pPolygon = New Polygon
        pTopoOp = TransitLine.Shape
        m_map.MapUnits = esriUnits.esriFeet
        pPolygon = pTopoOp.Buffer(1)
        pFilter = New SpatialFilter
        With pFilter
            .Geometry = pPolygon
            .GeometryField = FullNetwork.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelContains
        End With

        pFeatureCursor = FullNetwork.Search(pFilter, False)
        pOldFeature = pFeatureCursor.NextFeature

        'begin copying to PartialNetwork, one by one
        Do Until pOldFeature Is Nothing
            pNewFeature = PartialNetwork.CreateFeature
            pFields = PartialNetwork.Fields
            pNewFeature.Shape = pOldFeature.ShapeCopy
            pNewFeature.Store()

            For I = 0 To pFields.FieldCount - 1
                If (pFields.Field(I).Editable) Then
                    lSFld = pOldFeature.Fields.FindField(pFields.Field(I).Name)
                    If (lSFld <> -1) Then
                        If IsDBNull(pOldFeature.Value(lSFld)) = False Then

                            pNewFeature.Value(I) = pOldFeature.Value(lSFld)

                        End If

                    End If
                    pNewFeature.Store()
                End If
                GC.Collect()
            Next I
            pOldFeature = pFeatureCursor.NextFeature
        Loop

        Marshal.FinalReleaseComObject(pFeatureCursor)












    End Sub
    Public Function LoadEndPoints(ByVal PartialNetwork As IFeatureClass, ByVal EndPoints As IFeatureClass, ByVal featWorkspace As IWorkspace)




        Dim pline As IPolyline
        Dim pGeom As IGeometry
        Dim pPoint1 As IPoint
        Dim pPoint2 As IPoint
        Dim pLineFeature As IFeature
        Dim pNewFeature As IFeature
        Dim pFeatCursor As IFeatureCursor
        pFeatCursor = PartialNetwork.Search(Nothing, False)
        pLineFeature = pFeatCursor.NextFeature



        Do Until pLineFeature Is Nothing
            pGeom = pLineFeature.Shape
            pline = pLineFeature.Shape

            pPoint1 = pline.FromPoint
            pPoint2 = pline.ToPoint

            'begin copying to PartialNetwork, one by one

            pNewFeature = EndPoints.CreateFeature
            'Set pFields = EndPoints.Fields
            pNewFeature.Shape = pPoint1
            pNewFeature.Store()


            pNewFeature = EndPoints.CreateFeature
            pNewFeature.Shape = pPoint2
            pNewFeature.Store()


            pLineFeature = pFeatCursor.NextFeature
            GC.Collect()
        Loop
        Marshal.FinalReleaseComObject(pFeatCursor)



    End Function


    Public Function deleteAllRows0(ByVal pWSedit As IWorkspaceEdit, ByVal pFCls As IFeatureClass)
        Using pComReleaser As ComReleaser = New ComReleaser
            'WriteLogLine "start deleting all rows: " & Now()
            Debug.Print("start deleting all rows: " & Now())
            Dim pFCS As IFeatureCursor

            Dim pFtSet As ESRI.ArcGIS.esriSystem.Set
            Dim I As Integer
            Dim pFtEdit As IFeatureEdit, pFt As IFeature
            pFtSet = New ESRI.ArcGIS.esriSystem.Set
            I = 0
            pFCS = pFCls.Search(Nothing, False)
            pComReleaser.ManageLifetime(pFCS)
            

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

        End Using

    End Function

    Public Function LoadEndPoints2(ByVal NearestEdge As IFeature, ByVal EndPoints As IFeatureClass, ByVal featWorkspace As IWorkspace)




        Dim pline As IPolyline
        Dim pBusStopGeom As IGeometry

        Dim pGeom As IGeometry
        Dim pPoint1 As IPoint
        Dim pPoint2 As IPoint
        Dim pLineFeature As IFeature
        Dim pNewFeature As IFeature
        Dim pFeatCursor As IFeatureCursor
        ' pFeatCursor = PartialNetwork.Search(Nothing, False)
        pLineFeature = NearestEdge




        pGeom = pLineFeature.Shape
        pline = pLineFeature.Shape

        pPoint1 = pline.FromPoint
        pPoint2 = pline.ToPoint

        'begin copying to PartialNetwork, one by one

        pNewFeature = EndPoints.CreateFeature
        'Set pFields = EndPoints.Fields
        pNewFeature.Shape = pPoint1
        pNewFeature.Store()


        pNewFeature = EndPoints.CreateFeature
        pNewFeature.Shape = pPoint2
        pNewFeature.Store()


        ' pLineFeature = pFeatCursor.NextFeature
        'GC.Collect()

        'Marshal.FinalReleaseComObject(pFeatCursor)



    End Function
    Public Function ClosestEdgeEndPoint(ByVal NearestEdge As IPolyline, ByVal BusStop As IGeometry, ByVal featWorkspace As IWorkspace, ByRef ClosestStop As IPoint)
        Dim pline As IPolyline
        Dim pBusStopGeom As IGeometry

        Dim pGeom As IGeometry
        Dim pPoint1 As IPoint
        Dim pPoint2 As IPoint
        Dim pLineFeature As IFeature
        Dim pNewFeature As IFeature
        Dim pFeatCursor As IFeatureCursor
        Dim myline As New Line
        Dim myline2 As New Line

        Dim busPoint As IPoint
        busPoint = BusStop
        Dim dist1 As Double
        Dim dist2 As Double


        ' pFeatCursor = PartialNetwork.Search(Nothing, False)
        'pLineFeature = NearestEdge




        'pGeom = pLineFeature.Shape
        'pline = pLineFeature.Shape
        pline = NearestEdge

        pPoint1 = pline.FromPoint
        pPoint2 = pline.ToPoint

        myline.FromPoint = busPoint
        myline.ToPoint = pPoint1
        dist1 = myline.Length

        myline2.FromPoint = busPoint
        myline2.ToPoint = pPoint2
        dist2 = myline2.Length

        If dist1 <= dist2 Then
            ClosestStop = pPoint1
        Else
            ClosestStop = pPoint2

        End If



        'begin copying to PartialNetwork, one by one



        ' pLineFeature = pFeatCursor.NextFeature
        'GC.Collect()

        'Marshal.FinalReleaseComObject(pFeatCursor)



    End Function
    Public Sub LoadNetworkEdges2(ByVal PartialNetwork As IFeatureClass, ByVal featWorkspace As IWorkspace, ByVal TransitLine As IFeature, ByVal FullNetwork As IFeatureClass)
        'delete all features


        'get underlying edges under TransitLine
        Dim pFields As IFields
        Dim lSFld As Long
        Dim featureBuffer As IFeatureBuffer = PartialNetwork.CreateFeatureBuffer
        Dim pFeatureCursor As IFeatureCursor
        Dim pFeatureCursor2 As IFeatureCursor = PartialNetwork.Insert(True)

        Dim pOldFeature As IFeature
        Dim pNewFeature As IFeature
        Dim pPolygon As IPolygon
        Dim pTopoOp As ITopologicalOperator
        Dim pFilter As ISpatialFilter
        Dim I As Integer
        pPolygon = New Polygon
        pTopoOp = TransitLine.Shape
        m_map.MapUnits = esriUnits.esriFeet
        pPolygon = pTopoOp.Buffer(1)
        pFilter = New SpatialFilter
        With pFilter
            .Geometry = pPolygon
            .GeometryField = FullNetwork.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelContains

        End With

        pFeatureCursor = FullNetwork.Search(pFilter, False)
        Dim geomList As List(Of IGeometry) = New List(Of IGeometry)



        pOldFeature = pFeatureCursor.NextFeature

        'begin copying to PartialNetwork, one by one
        Do Until pOldFeature Is Nothing
            geomList.Add(pOldFeature.Shape)

            pOldFeature = pFeatureCursor.NextFeature
        Loop

        For Each geometry As IGeometry In geomList
            'Set the feature buffer's shape and insert it.
            featureBuffer.Shape = geometry
            pFeatureCursor2.InsertFeature(featureBuffer)
        Next geometry

        ' Attempt to flush the buffer.
        pFeatureCursor2.Flush()

        Marshal.FinalReleaseComObject(pFeatureCursor)












    End Sub
    Public Sub LoadNetworkEdges3(ByVal geomList As List(Of IPolyline), ByVal TransitLine As IFeature, ByVal FullNetwork As IFeatureClass)
        'delete all features


        'get underlying edges under TransitLine
        Dim pFields As IFields
        Dim lSFld As Long

        Dim pFeatureCursor As IFeatureCursor


        Dim pOldFeature As IFeature
        Dim pNewFeature As IFeature
        Dim pPolygon As IPolygon
        Dim pTopoOp As ITopologicalOperator
        Dim pFilter As ISpatialFilter
        Dim I As Integer
        pPolygon = New Polygon
        pTopoOp = TransitLine.Shape
        m_map.MapUnits = esriUnits.esriFeet
        pPolygon = pTopoOp.Buffer(1)
        pFilter = New SpatialFilter
        With pFilter
            .Geometry = pPolygon
            .GeometryField = FullNetwork.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelContains

        End With

        pFeatureCursor = FullNetwork.Search(pFilter, False)




        pOldFeature = pFeatureCursor.NextFeature

        'begin copying to PartialNetwork, one by one
        Do Until pOldFeature Is Nothing
            geomList.Add(pOldFeature.Shape)

            pOldFeature = pFeatureCursor.NextFeature
        Loop


        ' Attempt to flush the buffer.

        Marshal.FinalReleaseComObject(pFeatureCursor)












    End Sub
    Public Function getNearestEdge(ByVal partialNet As List(Of IPolyline), ByVal busStop As IGeometry) As IPolyline
        Dim pProxOp As IProximityOperator
        Dim edgeGeom As IPolyline = New polyline
        pProxOp = busStop
        Dim x As Integer
        Dim dist As Double
        Dim y As Double
        y = 0
        For x = 0 To partialNet.Count - 1
            dist = pProxOp.ReturnDistance(partialNet.Item(x))
            If y = 0 Then
                edgeGeom = partialNet.Item(x)
                y = dist
            ElseIf dist < y Then
                edgeGeom = partialNet.Item(x)
                y = dist
            End If




        Next
        Return edgeGeom

    End Function
    Private Sub frmSnapStopsToEndPoints_MinimumSizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MinimumSizeChanged

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Using pComReleaser As ComReleaser = New ComReleaser
            Dim I As Integer
            For I = 0 To m_map.LayerCount - 1
                If m_map.Layer(I).Name = ComboBox1.Text Then
                    m_PartialNetworkFL = m_map.Layer(I)
                    pComReleaser.ManageLifetime(m_PartialNetworkFL)

                End If
                If m_map.Layer(I).Name = ComboBox2.Text Then
                    m_FullNetworkFL = m_map.Layer(I)
                    pComReleaser.ManageLifetime(m_FullNetworkFL)
                End If
                'If m_pMap.Layer(i).name = ComboBox3.Text Then
                'Set m_pEdgesFL = m_pMap.Layer(i)
                'End If
                If m_map.Layer(I).Name = ComboBox3.Text Then
                    m_TransitRouteFL = m_map.Layer(I)
                    pComReleaser.ManageLifetime(m_TransitRouteFL)
                End If

                If m_map.Layer(I).Name = ComboBox4.Text Then
                    m_EndPointsFL = m_map.Layer(I)
                    pComReleaser.ManageLifetime(m_EndPointsFL)
                End If

                If m_map.Layer(I).Name = ComboBox5.Text Then
                    m_BusStopsFL = m_map.Layer(I)
                    pComReleaser.ManageLifetime(m_BusStopsFL)
                End If

            Next I
            Dim pDataset As IDataset
            pDataset = m_TransitRouteFL.FeatureClass


            m_Workspace = pDataset.Workspace

            Dim x As Integer
            Dim pFilter As IQueryFilter
            Dim pTransitFeature As IFeature
            Dim pBusStopFeature As IFeature
            Dim pNearFeat As IFeature
            Dim pDist As Double
            'Dim pFilter As IQueryFilter
            Dim pFeatureCursor As IFeatureCursor
            Dim pFeatureCursor2 As IFeatureCursor
            Dim pGeom As IGeometry
            'get all transit routes
            Dim pFilt As IQueryFilter
            Dim counter As Integer
            Dim pFSel As IFeatureSelection
            pFSel = m_TransitRouteFL

            counter = 59
            'pFilt = New QueryFilter
            'pFilt.WhereClause = "NewID >= " & counter

            'pFeatureCursor = m_TransitRouteFL.FeatureClass.Search(pFilt, False)


            Dim indexRouteID As Long
            Dim indexX As Long
            Dim indexY As Long
            Dim IndexNewOrder As Integer
            Dim indexDist As Long
            Dim lngRouteID As Long
            Dim pPoint As IPoint
            Dim xcord As Double
            Dim ycord As Double
            Dim prevX As Double
            Dim prevY As Double
            Dim a As Integer
            Dim b As Integer
            Dim NearestEdge As IFeature


            If pFSel.SelectionSet.Count = 0 Then
                ' do all of 'em
                'lCount = pFLayer.FeatureClass.FeatureCount(Nothing)
                pFeatureCursor = m_TransitRouteFL.FeatureClass.Search(Nothing, False)
                pComReleaser.ManageLifetime(pFeatureCursor)
            Else
                ' just do selected ones
                'lCount = pFSel.SelectionSet.Count
                pFSel.SelectionSet.Search(Nothing, False, pFeatureCursor)
                pComReleaser.ManageLifetime(pFeatureCursor)
            End If

            pTransitFeature = pFeatureCursor.NextFeature
            'loop through all transit routes      



            Do Until pTransitFeature Is Nothing

                'a new id to link routes to potential stops 
                indexRouteID = pTransitFeature.Fields.FindField("NewID")
                lngRouteID = pTransitFeature.Value(indexRouteID)





                Dim pWorkspaceEdit As IWorkspaceEdit
                pWorkspaceEdit = m_Workspace
                pComReleaser.ManageLifetime(pWorkspaceEdit)
                ' If m_PartialNetworkFL.FeatureClass.FeatureCount(Nothing) > 0 Then
                'deleteAllRows0(pWorkspaceEdit, m_PartialNetworkFL.FeatureClass)
                'End If
                'deleteAllRows0(pWorkspaceEdit, m_PartialNetworkFL.FeatureClass)
                'deleteAllRows0(pWorkspaceEdit, m_EndPointsFL.FeatureClass)
                Dim underlyingEdges As List(Of IPolyline) = New List(Of IPolyline)
                Dim nearEdge As IPolyline = New Polyline

                Try


                    'find the underlying edges and write them to a Feature Class


                    LoadNetworkEdges3(underlyingEdges, pTransitFeature, m_FullNetworkFL.FeatureClass)
                    'get the endpoints of each individual edge, put them into m_EndPointsFL
                    'LoadEndPoints(m_PartialNetworkFL.FeatureClass, m_EndPointsFL.FeatureClass, m_Workspace)
                Catch ex As Exception
                    MessageBox.Show(ex.Message.ToString)


                End Try

                'start locating bus stops
                pFilter = New QueryFilter
                pFilter.WhereClause = "NewID = " & lngRouteID
                'pFeatureCursor2 = m_BusStopsFL.FeatureClass.Update(pFilter, True)

                Dim pTableSort As ITableSort
                pTableSort = New TableSort
                pComReleaser.ManageLifetime(pTableSort)
                'Dim pQueryFilter As IQueryFilter
                'pQueryFilter = New QueryFilter

                'pQueryFilter.WhereClause = "[STATE_NAME] like  'A*'"

                With pTableSort
                    .Fields = "ORDER_"
                    .Ascending("ORDER_") = True
                    '.Ascending("count_name ") = True
                    '.CaseSensitive("state_name") = True
                    '.CaseSensitive("count_name ") = True
                    .QueryFilter = pFilter
                    .Table = m_BusStopsFL.FeatureClass
                End With

                pTableSort.Sort(Nothing)

                Dim pCursor As ICursor

                pCursor = pTableSort.Rows
                pComReleaser.ManageLifetime(pCursor)

                Dim pRow As IRow
                pRow = pCursor.NextRow
                a = 1
                b = 1
                'pBusStopFeature = pFeatureCursor2.NextFeature


                While Not pRow Is Nothing

                    Try
                        Dim intOID As Integer
                        Dim OID As Long
                        Dim pCurve As ICurve
                        intOID = pRow.Fields.FindField("OBJECTID")
                        OID = pRow.Value(intOID)


                        'pGeom = pBusStopFeature.Shape
                        'want to get nearest edge to bus stop


                        'then want to get the nearest end point of that edge
                        pGeom = m_BusStopsFL.FeatureClass.GetFeature(OID).Shape

                        pComReleaser.ManageLifetime(nearEdge)
                        nearEdge = getNearestEdge(underlyingEdges, pGeom)
                        'we have the nearest edge to the bus stop. Get the from and to points from this edge
                        'LoadEndPoints2(nearEdge, m_EndPointsFL.FeatureClass, m_Workspace)
                        Dim endPoints As ESRI.ArcGIS.Geometry.IPoint
                        endPoints = New ESRI.ArcGIS.Geometry.Point



                        ClosestEdgeEndPoint(nearEdge, pGeom, m_Workspace, endPoints)

                        'Dim endPoints As New GetEndPoints(pGeom, m_BusStopsFL.FeatureClass)
                        'endPoints.GetNearestFeature()
                        'GetNearestFeature(pGeom, m_EndPointsFL.FeatureClass, Nothing, True, pNearFeat, xcord, ycord)


                        indexX = pRow.Fields.FindField("x")
                        indexY = pRow.Fields.FindField("y")
                        indexDist = pRow.Fields.FindField("Distance")
                        IndexNewOrder = pRow.Fields.FindField("NewOrder")

                        'If pWorkspaceEdit.IsBeingEdited = False Then pWorkspaceEdit.StartEditing(True)
                        'pWorkspaceEdit.StartEditOperation()

                        'pBusStopFeature.Value(indexX) = endPoints.x
                        'pFeatureCursor2.UpdateFeature(pBusStopFeature)
                        'pBusStopFeature.Value(indexY) = endPoints.y
                        'pFeatureCursor2.UpdateFeature(pBusStopFeature)

                        'Assign new Bus Stop order
                        '   pRow.Value(IndexNewOrder) = a


                        'check to see if stop is 1st or last 
                        pCurve = pTransitFeature.Shape
                        '1st stop
                        If a = 1 Then
                            pRow.Value(IndexNewOrder) = b
                            pRow.Store()
                            pCurve = pTransitFeature.Shape
                            pRow.Value(indexX) = pCurve.FromPoint.X
                            pRow.Store()
                            pRow.Value(indexY) = pCurve.FromPoint.Y
                            pRow.Store()
                            prevX = pCurve.FromPoint.X
                            prevY = pCurve.FromPoint.Y

                            'last stop
                        ElseIf a = pTableSort.Table.RowCount(pFilter) Then
                            'check to see if the last stop has the same coords

                            If prevX = pCurve.ToPoint.X And prevY = pCurve.ToPoint.Y Then
                                b = b - 1
                                pRow.Value(IndexNewOrder) = -2

                                pRow.Store()
                                ' b = a - 1


                            Else

                                pRow.Value(indexX) = pCurve.ToPoint.X
                                pRow.Store()
                                pRow.Value(indexY) = pCurve.ToPoint.Y
                                pRow.Store()
                                pRow.Value(IndexNewOrder) = b
                                pRow.Store()

                            End If
                            'Other stops
                        Else
                            If prevX = endPoints.X And prevY = endPoints.Y Then
                                b = b - 1
                                pRow.Value(IndexNewOrder) = -1
                                pRow.Store()
                                'pRow.Value(indexDist) = endPoints.Distance
                                pRow.Store()
                            Else
                                pRow.Value(indexX) = endPoints.X
                                pRow.Store()
                                pRow.Value(indexY) = endPoints.Y
                                pRow.Store()
                                pRow.Value(IndexNewOrder) = b
                                pRow.Store()
                                prevX = endPoints.X
                                prevY = endPoints.Y
                                'pRow.Value(indexDist) = endPoints.Distance
                                pRow.Store()
                            End If


                            'pWorkspaceEdit.StopEditOperation()
                            'pWorkspaceEdit.StopEditing(True)


                            endPoints = Nothing

                        End If
                        'GC.Collect()

                        a = a + 1
                        b = b + 1
                        'pBusStopFeature = pFeatureCursor2.NextFeature
                        pRow = pCursor.NextRow

                    Catch ex As Exception
                        MessageBox.Show(ex.Message.ToString)

                    End Try
                End While
                Marshal.FinalReleaseComObject(pCursor)


                pTransitFeature = pFeatureCursor.NextFeature
                counter = counter + 1
                'GC.Collect()


            Loop
        End Using
        MessageBox.Show("Finshed!")
    End Sub
End Class