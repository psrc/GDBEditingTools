Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Controls
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports Microsoft.SqlServer
Imports System.Windows.Forms
Module functions


    Public Sub StartSDEWorkspaceEditorOperation(ByVal Workspace As IWorkspace, ByVal app As IApplication)
        Dim pEditor As IEditor
        'Dim pEditLayers As IEditLayers
        'Dim pDataset As IDataset
        'Dim pWrkspc As IWorkspace
        Dim pID As New UID

        Dim pFeatureWorkspace As IFeatureWorkspace
        Dim pWorkspaceEdit As IWorkspaceEdit2
        Dim pMWorkspaceEdit As IMultiuserWorkspaceEdit

        pFeatureWorkspace = Workspace
        pWorkspaceEdit = pFeatureWorkspace
        pMWorkspaceEdit = pFeatureWorkspace
        pID.Value = "esriEditor.Editor"
        pEditor = app.FindExtensionByCLSID(pID)

        If pEditor.EditState = esriEditState.esriStateNotEditing Then
            pEditor.StartEditing(Workspace)

            'If pMWorkspaceEdit.SupportsMultiuserEditSessionMode(esriMultiuserEditSessionMode.esriMESMVersioned) Then

            'pMWorkspaceEdit.StartMultiuserEditing(esriMultiuserEditSessionMode.esriMESMVersioned)
            'this is an example of how the MultiuserEditSessionMode property is implimented        
            'you would want to place edit operations in this space        
            'MsgBox(pMWorkspaceEdit.MultiuserEditSessionMode)
            ' pWorkspaceEdit.StopEditing(True)
            pWorkspaceEdit.StartEditOperation()
            pWorkspaceEdit.EnableUndoRedo()
        End If

        If pWorkspaceEdit.IsInEditOperation Then
            pWorkspaceEdit.StopEditOperation()
        End If





        'pDataset = TargetLayer.FeatureClass
        'pWrkspc = pDataset.Workspace


        pWorkspaceEdit.StartEditOperation()
    End Sub




    Public Sub StartWorkspaceEditorOperation(ByVal Workspace As IWorkspace, ByVal app As IApplication)
        Dim pEditor As IEditor
        Dim pEditLayers As IEditLayers
        'Dim pDataset As IDataset
        Dim pWrkspc As IWorkspace
        Dim pID As New UID



        'pDataset = TargetLayer.FeatureClass
        'pWrkspc = pDataset.Workspace
        pWrkspc = Workspace

        'get a reference to the editor
        pID.Value = "esriEditor.Editor"
        pEditor = app.FindExtensionByCLSID(pID)

        If pEditor.EditState = esriEditState.esriStateNotEditing Then
            pEditor.StartEditing(pWrkspc)
        End If


        ' Otherwise check we're editing in the same workspace
        'ElseIf Not (pEditor.EditWorkspace Is pWrkspc) Then
        'MsgBox "You Must End The Current Edit Session First", vbExclamation
        'Exit Sub
        'End If

        pEditLayers = pEditor
        'If Not pEditLyrs.IsEditable(TargetLayer) Then
        'MsgBox "This Layer Cannot Be Edited", vbExclamation
        'Exit Sub
        'End If

        ' Finally, set the current edit layer
        'pEditLayers.SetCurrentLayer(TargetLayer, 0)
        Dim pWorkspaceEdit As IWorkspaceEdit2
        'Dim pMWorkspaceEdit As IMultiuserWorkspaceEdit
        pWorkspaceEdit = pWrkspc
        'pMWorkspaceEdit = pWrkspc
        ' If pMWorkspaceEdit.SupportsMultiuserEditSessionMode(esriMultiuserEditSessionMode.esriMESMNonVersioned) Then
        'pMWorkspaceEdit.StartMultiuserEditing(esriMultiuserEditSessionMode.esriMESMNonVersioned)
        'End If
        If pWorkspaceEdit.IsInEditOperation Then
            pWorkspaceEdit.StopEditOperation()
        End If
        pWorkspaceEdit.StartEditOperation()





    End Sub
    Public Sub SetEditlayer(ByVal TargetLayer As IFeatureLayer, ByVal app As IApplication)


        Dim pEditor As IEditor
        Dim pEditLayers As IEditLayers
        Dim pDataset As IDataset
        Dim pWrkspc As IWorkspace
        Dim pID As New UID



        pDataset = TargetLayer.FeatureClass
        pWrkspc = pDataset.Workspace

        'get a reference to the editor
        pID.Value = "esriEditor.Editor"
        pEditor = app.FindExtensionByCLSID(pID)

        If pEditor.EditState = esriEditState.esriStateNotEditing Then
            pEditor.StartEditing(pWrkspc)
        End If


        ' Otherwise check we're editing in the same workspace
        'ElseIf Not (pEditor.EditWorkspace Is pWrkspc) Then
        'MsgBox "You Must End The Current Edit Session First", vbExclamation
        'Exit Sub
        'End If

        pEditLayers = pEditor
        'If Not pEditLyrs.IsEditable(TargetLayer) Then
        'MsgBox "This Layer Cannot Be Edited", vbExclamation
        'Exit Sub
        'End If

        ' Finally, set the current edit layer
        If TargetLayer.Name = "sde.SDE.TransRefEdges" Then
            pEditLayers.SetCurrentLayer(TargetLayer, 1)
        ElseIf TargetLayer.Name = "sde.SDE.TransRefJunctions" Then
            pEditLayers.SetCurrentLayer(TargetLayer, 4)
        Else
            pEditLayers.SetCurrentLayer(TargetLayer, 0)
        End If


    End Sub
    Public Sub StopWorkspaceEditOperation(ByVal app As IApplication, ByVal Workspace As IWorkspace)
        Dim pEditor As IEditor
        'Dim bHasEdits As Boolean
        'Dim pEditLayers As IEditLayers


        Dim pID As New UID
        Dim pWorkspaceEdit As IWorkspaceEdit2

        'get a reference to the editor
        pID.Value = "esriEditor.Editor"
        pEditor = app.FindExtensionByCLSID(pID)


        If pEditor.EditState = esriEditState.esriStateNotEditing Then
            Exit Sub
        End If

        Try

       
            pWorkspaceEdit = Workspace
            'pWorkspaceEdit.DisableUndoRedo()
            ' pWorkspaceEdit.StopEditOperation()
            If pWorkspaceEdit.IsInEditOperation = True Then
                pWorkspaceEdit.StopEditOperation()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try



    End Sub

    Public Sub StopEditor(ByVal SaveEdits As Boolean, ByVal app As IApplication)
        Dim pEditor As IEditor
        Dim pEditLayers As IEditLayers
        Dim pDataset As IDataset
        Dim pWrkspc As IWorkspace
        Dim pID As New UID



        'get a reference to the editor
        pID.Value = "esriEditor.Editor"
        pEditor = app.FindExtensionByCLSID(pID)


        pEditLayers = pEditor
        pDataset = pEditLayers.CurrentLayer
        pWrkspc = pDataset.Workspace
        Dim pWorkspaceEdit As IWorkspaceEdit
        pWorkspaceEdit = pWrkspc
        pWorkspaceEdit.StopEditOperation()
        'pWorkspaceEdit.StopEditing(SaveEdits)
        pEditor.StopEditing(SaveEdits)
    End Sub

    Public Function ProjectsMustCoverRefEdgesTopologyFix(ByVal RefEdges As IFeatureLayer, ByVal topoError As IFeature, ByVal app As IApplication, ByVal Workspace As IWorkspace) As IObject




        Dim topoUiD As New UID
        Dim fClass As IFeatureClass
        'Dim pPolyLine As IPolyline


        Dim tempEdge As IFeature
        Dim pRefEdge As IFeature


        SetEditlayer(RefEdges, app)
        fClass = RefEdges.FeatureClass



        'get the topology extension

        'MsgBox "1 " & topoError.TopologyRuleType
        'If topoError Is Nothing Then

        'End If


        'set the tempEdge to the topoError- QI to get the geometry of the error
        tempEdge = topoError

        pRefEdge = fClass.CreateFeature

        pRefEdge.Shape = tempEdge.Shape
        StartSDEWorkspaceEditorOperation(Workspace, app)
        pRefEdge.Store()

        ProjectsMustCoverRefEdgesTopologyFix = pRefEdge


        'new code
        'Dim mRefAttributes As New RefEdgeAttributes(pRefEdge)

        'Dim RefEdgeTable As ITable
        'Set RefEdgeTable = RefEdges.FeatureClass
        'Dim lngEdgeID As Long

        'mRefAttributes.EdgeID = NewRefEdgeID
        'pRefEdge.Store()
        pRefEdge = Nothing
        'StopWorkspaceEditOperation(app)



        'stop editor and save edits
        ' StopEditor True

    End Function

    Public Function EdgesEndPointMustBeCoveredByJunctionsTopologyFix(ByVal Junctions As IFeatureLayer, ByVal topoError As IFeature, ByVal app As IApplication) As IObject

        Dim topoUiD As New UID
        Dim fClass As IFeatureClass
        'Dim pPolyLine As IPolyline


        Dim tempJunction As IFeature
        Dim pJunction As IFeature

        SetEditlayer(Junctions, app)

        fClass = Junctions.FeatureClass



        'get the topology extension


        'set the tempEdge to the topoError- QI to get the geometry of the error
        tempJunction = topoError

        pJunction = fClass.CreateFeature

        pJunction.Shape = tempJunction.Shape

        pJunction.Store()


        'new code
        ' Dim mJunctionAttributes As New JunctionAttributes(pJunction)

        'Dim RefEdgeTable As ITable
        'Set RefEdgeTable = RefEdges.FeatureClass
        'Dim lngEdgeID As Long
        'mJunctionAttributes.PSRCJunctID = NewJunctionID
        'pJunction.Store()
        EdgesEndPointMustBeCoveredByJunctionsTopologyFix = pJunction
        'pJunction = Nothing

        'StopWorkspaceEditOperation(app)


        'stop editor and save edits
        ' StopEditor True

    End Function


    Public Sub GetLargestID(ByVal pTable As ITable, ByVal AttField As String, ByRef LargestAtt As Long)

        'on error GoTo eh

        Dim pSort As ITableSort
        Dim pCs As ICursor, pRow As IRow
        pSort = New TableSort

        'hyu - this block seems redundant, so is comment out.
        With pSort
            .QueryFilter = Nothing
            .Fields = pTable.OIDFieldName
            .Ascending(pTable.OIDFieldName) = False
            .Table = pTable
            .Sort(Nothing)
        End With

        pCs = pSort.Rows
        pRow = pCs.NextRow

        With pSort
            .QueryFilter = Nothing
            .Fields = AttField
            .Ascending(AttField) = False
            .Table = pTable

            .Sort(Nothing)
        End With

        pCs = pSort.Rows
        pRow = pCs.NextRow
        LargestAtt = pRow.Value(pTable.FindField(AttField))

        Exit Sub

eh:
        'MsgBox("Error: GetLargestID " & vbNewLine & Err.Description)

    End Sub


    Public Sub OpenFeatureInspector(ByVal app As IApplication)
        app.Document.activeView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, Nothing, Nothing)
        Dim pUID As New UID
        pUID.Value = "{44276914-98C1-11D1-8464-0000F875B9C6}"


        Dim commandItem As ICommandItem
        commandItem = app.Document.CommandBars.Find(pUID)
        commandItem.Refresh()

        commandItem.Execute()


    End Sub

    Public Sub CopyAttributes(ByVal pSourceFeature As IFeature, ByVal pDestinationFeature As IFeature) ', projRte As String)
        'on error GoTo eh
        'MsgBox "copying attrib"
        Dim pField As IField
        Dim pFields As IFields
        'Dim pRow As IRow
        Dim FieldCount As Integer
        ' Dim index As Long, Tindex As Long

        pFields = pSourceFeature.Fields
        For FieldCount = 0 To pFields.FieldCount - 1
            pField = pFields.Field(FieldCount)
            'If pField.Editable Then
            If Not pField.Type = esriFieldType.esriFieldTypeOID And Not pField.Type = esriFieldType.esriFieldTypeGeometry Then
                '[033006] hyu: check the editible property before assigning value
                If pDestinationFeature.Fields.Field(FieldCount).Editable Then pDestinationFeature.Value(FieldCount) = pSourceFeature.Value(FieldCount)
            End If
            'End If
        Next FieldCount

        Exit Sub

eh:

        'CloseLogFile "PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.CopyAttributes"
        'MsgBox "PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.CopyAttributes"

    End Sub

    Private Function BUILDPOLYLINE(ByVal pSegColl As ISegmentCollection) As IPolyline
        'on error GoTo eh

        'If fVerboseLog Then WriteLogLine "called GlobalMod.buildpolyline"
        'If fVerboseLog Then WriteLogLine "------------------------------"


        Dim pPolyline As IGeometryCollection
        'Dim pGeometry As IGeometry
        pPolyline = New Polyline
        pPolyline.AddGeometries(1, pSegColl)
        BUILDPOLYLINE = pPolyline

        Exit Function

        'eh:
        'CloseLogFile "PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.BUILDPOLYLINE"
        'MsgBox "PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.BUILDPOLYLINE"
        'MsgBox BUILDPOLYLINE.length
    End Function
    Public Function SplitEdgeTopologyFix2(ByVal RefEdges As IFeatureLayer, ByVal Junctions As IFeatureLayer, ByVal env As IEnvelope, ByVal app As IApplication, ByVal workspace As IWorkspace) As Collection





        Dim topoUiD As New UID
        Dim fClass As IFeatureClass
        'Dim pPolyline As IPolyline
        Dim topologyExt As ESRI.ArcGIS.EditorExt.ITopologyExtension
        Dim topology As ITopology
        Dim geoDS As IGeoDataset
        Dim errorContainer As IErrorFeatureContainer
        Dim eErrorFeat As IEnumTopologyErrorFeature
        Dim topoError As ITopologyErrorFeature
        Dim pFeature As IFeature
        Dim pJunction As IFeature
        Dim pRefEdge As IFeature
        Dim pFilter As ISpatialFilter
        Dim pFeatureCursor As IFeatureCursor
        Dim pQFilter As IQueryFilter
        Dim pRefCollection As New Collection


        SetEditlayer(RefEdges, app)
        fClass = RefEdges.FeatureClass



        'get the topology extension
        topoUiD.Value = "esriEditorExt.TopologyExtension"

        topologyExt = app.FindExtensionByCLSID(topoUiD)

        'get the current topology
        topology = topologyExt.CurrentTopology

        geoDS = topology

        errorContainer = topology

        eErrorFeat = errorContainer.ErrorFeaturesByRuleType(geoDS.SpatialReference, esriTopologyRuleType.esriTRTPointCoveredByLineEndpoint, env, True, False)


        topoError = eErrorFeat.Next

        'MsgBox "1 " & topoError.TopologyRuleType
        If topoError Is Nothing Then
            Exit Function
        End If


        'set the tempJunction to the topoError- QI to get the geometry of the error
        pFeature = topoError
        'MsgBox topoError.OriginOID
        'MsgBox topoError.DestinationOID
        'Dim intcounter As Integer
        'incounter = pFeature.Fields.FindField("SHAPE")
        'MsgBox incounter

        'Need to select the junction under the topoerror, which is has the geometry of the Junction:
        pQFilter = New QueryFilter
        pQFilter.WhereClause = "OBJECTID = " & topoError.OriginOID


        pFeatureCursor = Junctions.Search(pQFilter, False)

        pJunction = pFeatureCursor.NextFeature
        'MsgBox(pJunction.Fields.Field(4).AliasName)


        'Set up a spatial filter. esriSpatialContains will get the underlying edges and nothing else- what we want
        Dim pPolygon As IPolygon
        Dim pTopoOp As ITopologicalOperator
        pPolygon = New Polygon
        pTopoOp = pJunction.Shape 'make a buffer object
        pPolygon = pTopoOp.Buffer(1)  'map units only which are in feet and need points here
        pFilter = New SpatialFilter

        With pFilter
            .Geometry = pPolygon
            .GeometryField = RefEdges.FeatureClass.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
        End With



        'Set pFilter = New SpatialFilter
        'With pFilter
        '   Set .Geometry = pJunction.Shape
        '  .GeometryField = "SHAPE"
        ' .SpatialRel = esriSpatialRelIntersects

        'End With

        'Execute the filter:
        pFeatureCursor = RefEdges.Search(pFilter, False)
        'this should be the edge under the junction:
        pRefEdge = pFeatureCursor.NextFeature
        If pRefEdge Is Nothing Then
            'MsgBox("nothing")
        End If


        '*********Junction = tempJunction, Edge = pRefEdge
        'Dim pEnumVertex As IEnumVertex
        Dim pGeoColl As IGeometryCollection
        Dim pPolyCurve As IPolycurve2
        'Dim pEnumSplitPoint As IEnumSplitPoint
        Dim pNewFeature As IFeature
        'Dim PartCount As Integer
        Dim index As Long, indexJ As Long

        Dim pFeatLayer As IFeatureLayer
        pFeatLayer = New FeatureLayer
        pFeatLayer.FeatureClass = Junctions.FeatureClass
        Dim pPoint As IPoint

        'Dim pFilter As ISpatialFilter
        Dim pFC As IFeatureCursor
        Dim pJFeat As IFeature
        Dim indexEdgeID As Long

        'Dim lngNewEdgeID1 As Long
        'Dim lngNewEdgeID2 As Long
        Dim pTable As ITable
        pTable = RefEdges.FeatureClass
        'GetLargestID(pTable, "PSRCEdgeID", lngNewEdgeID1)
        'lngNewEdgeID1 = lngNewEdgeID1 + 1
        'lngNewEdgeID2 = lngNewEdgeID1 + 1


        Dim JunctionAtt As New JunctionAttributes(pJunction)

        'JunctionAtt.JunctionFeature = pJunction
        'Dim firstEdgeAtt As New RefEdgesAttributes
        'Set firstEdgeAtt.RefEdgeFeature = pRefEdge
        Dim SplitID As Long
        SplitID = JunctionAtt.PSRCJunctID

        pPolyCurve = pRefEdge.Shape
        pPoint = pJunction.Shape
        Dim bSplit As Boolean, lPart As Long, lSeg As Long
        pPolyCurve.SplitAtPoint(pPoint, False, True, bSplit, lPart, lSeg)
        ' MsgBox bSplit
        pGeoColl = pPolyCurve

        Dim originalI As Long
        Dim originalJ As Long
        index = pRefEdge.Fields.FindField("INode")
        originalI = pRefEdge.Value(index)
        index = pRefEdge.Fields.FindField("JNode")
        originalJ = pRefEdge.Value(index)
        indexEdgeID = pRefEdge.Fields.FindField("PSRCEdgeID")



        Dim i As Integer
        For i = 0 To 1
            If i = 0 Then
                'StartSDEWorkspaceEditorOperation(workspace, app)
                pRefEdge.Shape = BUILDPOLYLINE(pGeoColl.Geometry(i)) 'orig feature becomes first part
                'pRefEdge.Value(indexEdgeID) = lngNewEdgeID1
                StartSDEWorkspaceEditorOperation(workspace, app)
                pRefEdge.Store()
                'StopWorkspaceEditOperation(workspace, app)

                'Marshal.ReleaseComObject(RefEdges)

               

                'StopWorkspaceEditOperation(app, workspace)

                pNewFeature = pRefEdge
                pRefCollection.Add(pNewFeature)
                'Marshal.ReleaseComObject(pRefEdge)
            Else
                'StartSDEWorkspaceEditorOperation(workspace, app)
                pNewFeature = RefEdges.FeatureClass.CreateFeature
                pNewFeature.Shape = BUILDPOLYLINE(pGeoColl.Geometry(i))
                CopyAttributes(pRefEdge, pNewFeature)
                'pNewFeature.Value(indexEdgeID) = lngNewEdgeID2
                StartSDEWorkspaceEditorOperation(workspace, app)
                pNewFeature.Store()
                'StopWorkspaceEditOperation(workspace, app)


                'StopWorkspaceEditOperation(app, workspace)
                pRefCollection.Add(pNewFeature)
                'Marshal.ReleaseComObject(pNewFeature)
            End If
            pFilter = New SpatialFilter
            With pFilter
                .Geometry = pNewFeature.ShapeCopy
                .GeometryField = Junctions.FeatureClass.ShapeFieldName
                .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            End With

            If Junctions.FeatureClass.FeatureCount(pFilter) > 0 Then
                'MsgBox m_junctShp.featurecount(pFilter)
                pFC = pFeatLayer.Search(pFilter, False)
                pJFeat = pFC.NextFeature
                Do Until pJFeat Is Nothing
                    Dim nodecnt As Long
                    nodecnt = 0
                    'keep original direction
                    'index = pRefEdge.Fields.FindField("UseEmmeN")

                    'If pNewFeature.Value(index) = 0 Then
                    indexJ = pJFeat.Fields.FindField("PSRCJunctID")
                    'Else
                    'indexJ = pJFeat.Fields.FindField("Emme2nodeID")
                    'End If
                    If (pJFeat.Value(indexJ)) > 0 Then
                        If (originalI = pJFeat.Value(indexJ)) Then

                            index = pNewFeature.Fields.FindField("INode")
                            pNewFeature.Value(index) = pJFeat.Value(indexJ)
                            index = pNewFeature.Fields.FindField("JNode")
                            pNewFeature.Value(index) = SplitID

                            'index = pNewFeature.Fields.FindField("UseEmmeN")

                            'If pNewFeature.Value(index) > 0 Then pNewFeature.Value(index) = 3
                            'If pNewFeature.Value(index) = 0 Then pNewFeature.Value(index) = 2
                        ElseIf (originalJ = pJFeat.Value(indexJ)) Then

                            index = pNewFeature.Fields.FindField("JNode")
                            pNewFeature.Value(index) = pJFeat.Value(indexJ)
                            index = pNewFeature.Fields.FindField("INode")
                            pNewFeature.Value(index) = SplitID
                            'index = pNewFeature.Fields.FindField("UseEmmeN")

                            'If pNewFeature.Value(index) > 0 Then pNewFeature.Value(index) = 3
                            'If pNewFeature.Value(index) = 0 Then pNewFeature.Value(index) = 1
                        End If
                    Else 'its Null
                        'If fVerboseLog Then WriteLogLine "PSRCJunctionID is null"
                    End If

                    pJFeat = pFC.NextFeature
                Loop
                pNewFeature.Store()
            End If
            pJFeat = Nothing
        Next i
        SplitEdgeTopologyFix2 = pRefCollection

        pFeatLayer = Nothing
        'StopWorkspaceEditOperation(app)
        If Not pNewFeature Is Nothing Then
            Marshal.ReleaseComObject(pNewFeature)
        End If
        Marshal.ReleaseComObject(pRefEdge)
        GC.Collect()
        Exit Function


        'CloseLogFile "PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.SplitFeature"
        'MsgBox "PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.SplitFeature"

        'Print #1, Err.Description, vbInformation, "SplitFeatures"
        pFeatLayer = Nothing
    End Function
    Public Function SplitEdgeTopologyFix3(ByVal RefEdges As IFeatureLayer, ByVal Junctions As IFeatureLayer, ByVal fTopoError As ITopologyErrorFeature, ByVal app As IApplication, ByVal workspace As IWorkspace) As Collection





        Dim topoUiD As New UID
        Dim fClass As IFeatureClass
        'Dim pPolyline As IPolyline
        'Dim topologyExt As ESRI.ArcGIS.EditorExt.ITopologyExtension
        'Dim topology As ITopology
        'Dim geoDS As IGeoDataset
        'Dim errorContainer As IErrorFeatureContainer
        'Dim eErrorFeat As IEnumTopologyErrorFeature
        Dim topoError As ITopologyErrorFeature
        Dim pFeature As IFeature
        Dim pJunction As IFeature
        Dim pRefEdge As IFeature
        Dim pFilter As ISpatialFilter
        Dim pFeatureCursor As IFeatureCursor
        Dim pQFilter As IQueryFilter
        Dim pRefCollection As New Collection

        topoError = fTopoError
        SetEditlayer(RefEdges, app)
        fClass = RefEdges.FeatureClass



        'get the topology extension
        'topoUiD.Value = "esriEditorExt.TopologyExtension"

        'topologyExt = app.FindExtensionByCLSID(topoUiD)

        'get the current topology
        'topology = topologyExt.CurrentTopology

        'geoDS = topology

        'errorContainer = topology

        'eErrorFeat = errorContainer.ErrorFeaturesByRuleType(geoDS.SpatialReference, esriTopologyRuleType.esriTRTPointCoveredByLineEndpoint, env, True, False)


        'topoError = eErrorFeat.Next

        'MsgBox "1 " & topoError.TopologyRuleType
        ' If topoError Is Nothing Then
        ' Function
        ' End If


        'set the tempJunction to the topoError- QI to get the geometry of the error
        pFeature = topoError
        'MsgBox topoError.OriginOID
        'MsgBox topoError.DestinationOID
        'Dim intcounter As Integer
        'incounter = pFeature.Fields.FindField("SHAPE")
        'MsgBox incounter

        'Need to select the junction under the topoerror, which is has the geometry of the Junction:
        pQFilter = New QueryFilter
        pQFilter.WhereClause = "OBJECTID = " & topoError.OriginOID


        pFeatureCursor = Junctions.Search(pQFilter, False)

        pJunction = pFeatureCursor.NextFeature
        'MsgBox(pJunction.Fields.Field(4).AliasName)


        'Set up a spatial filter. esriSpatialContains will get the underlying edges and nothing else- what we want
        Dim pPolygon As IPolygon
        Dim pTopoOp As ITopologicalOperator
        pPolygon = New Polygon
        pTopoOp = pFeature.Shape  'make a buffer object
        pPolygon = pTopoOp.Buffer(1)  'map units only which are in feet and need points here
        pFilter = New SpatialFilter

        With pFilter
            .Geometry = pPolygon
            .GeometryField = RefEdges.FeatureClass.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
        End With



        'Set pFilter = New SpatialFilter
        'With pFilter
        '   Set .Geometry = pJunction.Shape
        '  .GeometryField = "SHAPE"
        ' .SpatialRel = esriSpatialRelIntersects

        'End With

        'Execute the filter:
        pFeatureCursor = RefEdges.Search(pFilter, False)
        'this should be the edge under the junction:
        pRefEdge = pFeatureCursor.NextFeature
        If pRefEdge Is Nothing Then
            'MsgBox("nothing")
        End If


        '*********Junction = tempJunction, Edge = pRefEdge
        'Dim pEnumVertex As IEnumVertex
        Dim pGeoColl As IGeometryCollection
        Dim pPolyCurve As IPolycurve2
        'Dim pEnumSplitPoint As IEnumSplitPoint
        Dim pNewFeature As IFeature
        'Dim PartCount As Integer
        Dim index As Long, indexJ As Long

        Dim pFeatLayer As IFeatureLayer
        pFeatLayer = New FeatureLayer
        pFeatLayer.FeatureClass = Junctions.FeatureClass
        Dim pPoint As IPoint

        'Dim pFilter As ISpatialFilter
        Dim pFC As IFeatureCursor
        Dim pJFeat As IFeature
        Dim indexEdgeID As Long

        'Dim lngNewEdgeID1 As Long
        'Dim lngNewEdgeID2 As Long
        Dim pTable As ITable
        pTable = RefEdges.FeatureClass
        'GetLargestID(pTable, "PSRCEdgeID", lngNewEdgeID1)
        'lngNewEdgeID1 = lngNewEdgeID1 + 1
        'lngNewEdgeID2 = lngNewEdgeID1 + 1


        Dim JunctionAtt As New JunctionAttributes(pJunction)

        'JunctionAtt.JunctionFeature = pJunction
        'Dim firstEdgeAtt As New RefEdgesAttributes
        'Set firstEdgeAtt.RefEdgeFeature = pRefEdge
        Dim SplitID As Long
        SplitID = JunctionAtt.PSRCJunctID

        pPolyCurve = pRefEdge.Shape
        pPoint = pJunction.Shape
        Dim bSplit As Boolean, lPart As Long, lSeg As Long
        pPolyCurve.SplitAtPoint(pPoint, False, True, bSplit, lPart, lSeg)
        ' MsgBox bSplit
        pGeoColl = pPolyCurve
        'MessageBox.Show(pGeoColl.GeometryCount)

        Dim originalI As Long
        Dim originalJ As Long
        index = pRefEdge.Fields.FindField("INode")
        originalI = pRefEdge.Value(index)
        index = pRefEdge.Fields.FindField("JNode")
        originalJ = pRefEdge.Value(index)
        indexEdgeID = pRefEdge.Fields.FindField("PSRCEdgeID")



        Dim i As Integer
        For i = 0 To 1
            If i = 0 Then
                'StartSDEWorkspaceEditorOperation(workspace, app)
                pRefEdge.Shape = BUILDPOLYLINE(pGeoColl.Geometry(i)) 'orig feature becomes first part
                'pRefEdge.Value(indexEdgeID) = lngNewEdgeID1
                StartSDEWorkspaceEditorOperation(workspace, app)
                pRefEdge.Store()
                'StopWorkspaceEditOperation(workspace, app)

                'Marshal.ReleaseComObject(RefEdges)



                'StopWorkspaceEditOperation(app, workspace)

                pNewFeature = pRefEdge
                pRefCollection.Add(pNewFeature)
                'Marshal.ReleaseComObject(pRefEdge)
            Else
                'StartSDEWorkspaceEditorOperation(workspace, app)
                pNewFeature = RefEdges.FeatureClass.CreateFeature
                pNewFeature.Shape = BUILDPOLYLINE(pGeoColl.Geometry(i))
                CopyAttributes(pRefEdge, pNewFeature)
                'pNewFeature.Value(indexEdgeID) = lngNewEdgeID2
                StartSDEWorkspaceEditorOperation(workspace, app)
                pNewFeature.Store()
                'StopWorkspaceEditOperation(workspace, app)


                'StopWorkspaceEditOperation(app, workspace)
                pRefCollection.Add(pNewFeature)
                'Marshal.ReleaseComObject(pNewFeature)
            End If
            pFilter = New SpatialFilter
            With pFilter
                .Geometry = pNewFeature.ShapeCopy
                .GeometryField = Junctions.FeatureClass.ShapeFieldName
                .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            End With
            If Junctions.FeatureClass.FeatureCount(pFilter) > 0 Then
                'MsgBox m_junctShp.featurecount(pFilter)
                pFC = pFeatLayer.Search(pFilter, False)
                pJFeat = pFC.NextFeature
                Do Until pJFeat Is Nothing
                    Dim nodecnt As Long
                    nodecnt = 0
                    'keep original direction
                    'index = pRefEdge.Fields.FindField("UseEmmeN")

                    'If pNewFeature.Value(index) = 0 Then
                    indexJ = pJFeat.Fields.FindField("PSRCJunctID")
                    'Else
                    'indexJ = pJFeat.Fields.FindField("Emme2nodeID")
                    'End If
                    If (pJFeat.Value(indexJ)) > 0 Then
                        If (originalI = pJFeat.Value(indexJ)) Then

                            index = pNewFeature.Fields.FindField("INode")
                            pNewFeature.Value(index) = pJFeat.Value(indexJ)
                            index = pNewFeature.Fields.FindField("JNode")
                            pNewFeature.Value(index) = SplitID

                            'index = pNewFeature.Fields.FindField("UseEmmeN")

                            'If pNewFeature.Value(index) > 0 Then pNewFeature.Value(index) = 3
                            'If pNewFeature.Value(index) = 0 Then pNewFeature.Value(index) = 2
                        ElseIf (originalJ = pJFeat.Value(indexJ)) Then

                            index = pNewFeature.Fields.FindField("JNode")
                            pNewFeature.Value(index) = pJFeat.Value(indexJ)
                            index = pNewFeature.Fields.FindField("INode")
                            pNewFeature.Value(index) = SplitID
                            'index = pNewFeature.Fields.FindField("UseEmmeN")

                            'If pNewFeature.Value(index) > 0 Then pNewFeature.Value(index) = 3
                            'If pNewFeature.Value(index) = 0 Then pNewFeature.Value(index) = 1
                        End If
                    Else 'its Null
                        'If fVerboseLog Then WriteLogLine "PSRCJunctionID is null"
                    End If

                    pJFeat = pFC.NextFeature
                Loop
                pNewFeature.Store()
            End If
            pJFeat = Nothing
        Next i
        SplitEdgeTopologyFix3 = pRefCollection

        pFeatLayer = Nothing
        'StopWorkspaceEditOperation(app)
        If Not pNewFeature Is Nothing Then
            Marshal.ReleaseComObject(pNewFeature)
        End If
        Marshal.ReleaseComObject(pRefEdge)
        GC.Collect()
        Exit Function


        'CloseLogFile "PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.SplitFeature"
        'MsgBox "PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.SplitFeature"

        'Print #1, Err.Description, vbInformation, "SplitFeatures"
        pFeatLayer = Nothing
    End Function

    Public Sub SplitEdgeByNewJunction(ByVal RefEdges As IFeatureLayer, ByVal Junctions As IFeatureLayer, ByVal app As IApplication)
        'this is when the user is adding just a junction, not a project




        Dim topoUiD As New UID
        Dim fClass As IFeatureClass
        'Dim pPolyline As IPolyline
        Dim topologyExt As ESRI.ArcGIS.EditorExt.ITopologyExtension
        Dim topology As ITopology
        Dim geoDS As IGeoDataset
        Dim errorContainer As IErrorFeatureContainer
        Dim eErrorFeat As IEnumTopologyErrorFeature
        Dim topoError As ITopologyErrorFeature
        Dim pFeature As IFeature
        Dim pJunction As IFeature
        Dim pRefEdge As IFeature
        Dim pFilter As ISpatialFilter
        Dim pFeatureCursor As IFeatureCursor
        Dim pQFilter As IQueryFilter


        SetEditlayer(RefEdges, app)
        fClass = RefEdges.FeatureClass



        'get the topology extension
        topoUiD.Value = "esriEditorExt.TopologyExtension"

        topologyExt = app.FindExtensionByCLSID(topoUiD)

        'get the current topology
        topology = topologyExt.CurrentTopology

        geoDS = topology

        errorContainer = topology

        eErrorFeat = errorContainer.ErrorFeaturesByRuleType(geoDS.SpatialReference, esriTopologyRuleType.esriTRTPointCoveredByLineEndpoint, Nothing, True, False)


        topoError = eErrorFeat.Next

        'MsgBox "1 " & topoError.TopologyRuleType
        If topoError Is Nothing Then
            Exit Sub
        End If


        'set the tempJunction to the topoError- QI to get the geometry of the error
        pFeature = topoError
        'MsgBox topoError.OriginOID
        'MsgBox topoError.DestinationOID
        'Dim intcounter As Integer
        'incounter = pFeature.Fields.FindField("SHAPE")
        'MsgBox incounter

        'Need to select the junction under the topoerror, which is has the geometry of the Junction:
        pQFilter = New QueryFilter
        pQFilter.WhereClause = "OBJECTID = " & topoError.OriginOID


        pFeatureCursor = Junctions.Search(pQFilter, False)

        pJunction = pFeatureCursor.NextFeature
        'MsgBox(pJunction.Fields.Field(4).AliasName)


        'Set up a spatial filter. esriSpatialContains will get the underlying edges and nothing else- what we want
        Dim pPolygon As IPolygon
        Dim pTopoOp As ITopologicalOperator
        pPolygon = New Polygon
        pTopoOp = pJunction.Shape 'make a buffer object
        pPolygon = pTopoOp.Buffer(1)  'map units only which are in feet and need points here
        pFilter = New SpatialFilter

        With pFilter
            .Geometry = pPolygon
            .GeometryField = RefEdges.FeatureClass.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
        End With



        'Set pFilter = New SpatialFilter
        'With pFilter
        '   Set .Geometry = pJunction.Shape
        '  .GeometryField = "SHAPE"
        ' .SpatialRel = esriSpatialRelIntersects

        'End With

        'Execute the filter:
        pFeatureCursor = RefEdges.Search(pFilter, False)
        'this should be the edge under the junction:
        pRefEdge = pFeatureCursor.NextFeature
        If pRefEdge Is Nothing Then
            'MsgBox("nothing")
        End If


        '*********Junction = tempJunction, Edge = pRefEdge
        'Dim pEnumVertex As IEnumVertex
        Dim pGeoColl As IGeometryCollection
        Dim pPolyCurve As IPolycurve2
        'Dim pEnumSplitPoint As IEnumSplitPoint
        Dim pNewFeature As IFeature
        'Dim PartCount As Integer
        Dim index As Long, indexJ As Long

        Dim pFeatLayer As IFeatureLayer
        pFeatLayer = New FeatureLayer
        pFeatLayer.FeatureClass = Junctions.FeatureClass
        Dim pPoint As IPoint

        'Dim pFilter As ISpatialFilter
        Dim pFC As IFeatureCursor
        Dim pJFeat As IFeature
        Dim indexEdgeID As Long

        'Dim lngNewEdgeID1 As Long
        'Dim lngNewEdgeID2 As Long
        Dim pTable As ITable
        pTable = RefEdges.FeatureClass
        'GetLargestID(pTable, "PSRCEdgeID", lngNewEdgeID1)
        'lngNewEdgeID1 = lngNewEdgeID1 + 1
        'lngNewEdgeID2 = lngNewEdgeID1 + 1


        Dim JunctionAtt As New JunctionAttributes(pJunction)

        'JunctionAtt.JunctionFeature = pJunction
        'Dim firstEdgeAtt As New RefEdgesAttributes
        'Set firstEdgeAtt.RefEdgeFeature = pRefEdge
        Dim SplitID As Long
        SplitID = JunctionAtt.PSRCJunctID

        pPolyCurve = pRefEdge.Shape
        pPoint = pJunction.Shape
        Dim bSplit As Boolean, lPart As Long, lSeg As Long
        pPolyCurve.SplitAtPoint(pPoint, False, True, bSplit, lPart, lSeg)
        ' MsgBox bSplit
        pGeoColl = pPolyCurve

        Dim originalI As Long
        Dim originalJ As Long
        index = pRefEdge.Fields.FindField("INode")
        originalI = pRefEdge.Value(index)
        index = pRefEdge.Fields.FindField("JNode")
        originalJ = pRefEdge.Value(index)
        indexEdgeID = pRefEdge.Fields.FindField("PSRCEdgeID")



        Dim i As Integer
        For i = 0 To 1
            If i = 0 Then
                pRefEdge.Shape = BUILDPOLYLINE(pGeoColl.Geometry(i)) 'orig feature becomes first part
                'pRefEdge.Value(indexEdgeID) = lngNewEdgeID1
                pRefEdge.Store()
                pNewFeature = pRefEdge
            Else
                pNewFeature = RefEdges.FeatureClass.CreateFeature
                pNewFeature.Shape = BUILDPOLYLINE(pGeoColl.Geometry(i))
                CopyAttributes(pRefEdge, pNewFeature)
                'pNewFeature.Value(indexEdgeID) = lngNewEdgeID2
                pNewFeature.Store()
            End If
            pFilter = New SpatialFilter
            With pFilter
                .Geometry = pNewFeature.ShapeCopy
                .GeometryField = Junctions.FeatureClass.ShapeFieldName
                .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            End With
            If Junctions.FeatureClass.FeatureCount(pFilter) > 0 Then
                'MsgBox m_junctShp.featurecount(pFilter)
                pFC = pFeatLayer.Search(pFilter, False)
                pJFeat = pFC.NextFeature
                Do Until pJFeat Is Nothing
                    Dim nodecnt As Long
                    nodecnt = 0
                    'keep original direction
                    'index = pRefEdge.Fields.FindField("UseEmmeN")

                    'If pNewFeature.Value(index) = 0 Then
                    indexJ = pJFeat.Fields.FindField("PSRCJunctID")
                    'Else
                    'indexJ = pJFeat.Fields.FindField("Emme2nodeID")
                    'End If
                    If (pJFeat.Value(indexJ)) > 0 Then
                        If (originalI = pJFeat.Value(indexJ)) Then

                            index = pNewFeature.Fields.FindField("INode")
                            pNewFeature.Value(index) = pJFeat.Value(indexJ)
                            index = pNewFeature.Fields.FindField("JNode")
                            pNewFeature.Value(index) = SplitID

                            'index = pNewFeature.Fields.FindField("UseEmmeN")

                            'If pNewFeature.Value(index) > 0 Then pNewFeature.Value(index) = 3
                            'If pNewFeature.Value(index) = 0 Then pNewFeature.Value(index) = 2
                        ElseIf (originalJ = pJFeat.Value(indexJ)) Then

                            index = pNewFeature.Fields.FindField("JNode")
                            pNewFeature.Value(index) = pJFeat.Value(indexJ)
                            index = pNewFeature.Fields.FindField("INode")
                            pNewFeature.Value(index) = SplitID
                            'index = pNewFeature.Fields.FindField("UseEmmeN")

                            'If pNewFeature.Value(index) > 0 Then pNewFeature.Value(index) = 3
                            'If pNewFeature.Value(index) = 0 Then pNewFeature.Value(index) = 1
                        End If
                    Else 'its Null
                        'If fVerboseLog Then WriteLogLine "PSRCJunctionID is null"
                    End If

                    pJFeat = pFC.NextFeature
                Loop
                pNewFeature.Store()
            End If
            pJFeat = Nothing
        Next i

        pFeatLayer = Nothing
        'StopWorkspaceEditOperation(app)

        Exit Sub

eh:
        'CloseLogFile "PROGRAM ERROR: " & Err.Number & ", " & Err.Description & "--GlobalMod.SplitFeature"
        'MsgBox "PROGRAM ERROR: " & Err.Number & ", " & Err.Description, , "GlobalMod.SplitFeature"

        'Print #1, Err.Description, vbInformation, "SplitFeatures"
        pFeatLayer = Nothing
    End Sub
    Public Sub RemoveLastPointInSketch(ByVal edSketch As IEditSketch2, ByVal editor As IEditor)
        'removes the most recent vertex in an editing sketch

        If edSketch Is Nothing Then
            Exit Sub
        End If

        Dim ptColl As IPointCollection4
        ptColl = edSketch.Geometry
        Dim sketchOp As ISketchOperation2
        sketchOp = New SketchOperation

        sketchOp.Start(editor)


        Dim pt As IPoint
        pt = ptColl.Point(ptColl.PointCount - 1)
        ptColl.RemovePoints(ptColl.PointCount - 1, 1)

        sketchOp.Finish(edSketch.Geometry.Envelope, esriSketchOperationType.esriSketchOperationVertexDeleted, pt)
    End Sub
    Public Function FirstVertexInSketch(ByVal edSketch As IEditSketch2, ByVal editor As IEditor) As Boolean
        If edSketch Is Nothing Then
            Exit Function
        End If
        Dim ptColl As IPointCollection4
        ptColl = edSketch.Geometry
        Dim sketchOp As ISketchOperation2
        sketchOp = New SketchOperation

        sketchOp.Start(editor)

        If ptColl.PointCount = 1 Then
            FirstVertexInSketch = True
        Else
            FirstVertexInSketch = False

        End If



    End Function
    Public Sub ValidateTopologyForCurrentExtent(ByVal app As IApplication)
        'calls the ValidateToplologyForCurrentExtent command
        Dim pUID As New UID
        pUID.Value = "{8FDE78D4-5160-4410-B611-C179F003680B}"


        Dim commandItem As ICommandItem
        commandItem = app.Document.CommandBars.Find(pUID)
        commandItem.Execute()

    End Sub
    Public Sub ValidateTopology(ByVal topo As ITopology, ByVal env As IEnvelope)

        topo.ValidateTopology(env)

    End Sub
    Public Function TopologyErrors(ByVal topo As ITopology) As Boolean
        Dim errorContainer As IErrorFeatureContainer
        Dim eErrorFeat As IEnumTopologyErrorFeature
        Dim topoError As ITopologyErrorFeature
        Dim geoDS As IGeoDataset

        geoDS = topo
        errorContainer = topo


        eErrorFeat = errorContainer.ErrorFeaturesByRuleType(geoDS.SpatialReference, esriTopologyRuleType.esriTRTAny, Nothing, True, False)


        topoError = eErrorFeat.Next
        If Not topoError Is Nothing Then
            TopologyErrors = True
        Else
            TopologyErrors = False
        End If
    End Function
    Public Sub SelectFeature(ByVal map As IMap, ByVal feature As IObject, ByVal mxdoc As IMxDocument, ByVal activeview As IActiveView)
        Dim a As Integer
        Dim flayer As IFeatureLayer


        For a = 0 To map.LayerCount - 1
            If TypeOf map.Layer(a) Is IFeatureLayer Then
                flayer = map.Layer(a)
                If flayer.FeatureClass Is feature.Table Then
                    map.ClearSelection()
                    map.SelectFeature(mxdoc.FocusMap.Layer(a), feature)
                    activeview.Refresh()

                    Exit For
                End If
            End If

        Next
    End Sub
    Public Sub SelectFeatures(ByVal map As IMap, ByVal features As Collection, ByVal mxdoc As IMxDocument, ByVal activeview As IActiveView)
        Dim a As Integer
        Dim x As Integer
        Dim flayer As IFeatureLayer
        'Dim pObject As IObject
        Dim pFeature As IFeature


        map.ClearSelection()
        For a = 0 To map.LayerCount - 1
            If TypeOf map.Layer(a) Is IFeatureLayer Then
                flayer = map.Layer(a)
                For x = 1 To features.Count

                    'pObject = CType(features(x), IObject)

                    pFeature = CType(features.Item(x), IFeature)

                    If flayer.FeatureClass Is pFeature.Table Then
                        map.SelectFeature(mxdoc.FocusMap.Layer(a), features(x))

                        Dim pFSel As IFeatureSelection
                        pFSel = flayer
                        pFSel.SelectionChanged()

                    End If
                Next
            End If
        Next
        ' activeview.Refresh()
                

    End Sub
    Public Function GetDirtyAreas(ByVal topo As ITopology) As IEnvelope
        Dim pPolygon As ESRI.ArcGIS.Geometry.IPolygon
        Dim pGeoDS As IGeoDataset
        Dim pLocation As ESRI.ArcGIS.Geometry.ISegmentCollection
        pGeoDS = topo

        pLocation = New Polygon
        pLocation.SetRectangle(pGeoDS.Extent)
        pPolygon = topo.DirtyArea(pLocation)




        GetDirtyAreas = pPolygon.Envelope
    End Function

    Public Sub SetSnappingTolerance(ByVal fEditor As IEditor, ByVal fSnappingTolerance As Double)

        'Dim pFtrSnapAgent As IFeatureSnapAgent
        Dim pSnapEnv As ISnapEnvironment
        Dim pSnapWindow As ISnappingWindow
        Dim pID As New UID

        pSnapEnv = fEditor
        'pFtrSnapAgent = New FeatureSnap
        'With pFtrSnapAgent
        '.FeatureClass = fRefEdges
        '.HitType = esriGeometryPartEndpoint  'And esriGeometryPartVertex
        '.HitType = esriGeometryPartBoundary

        'End With

        ' Clear all current snapping agents
        'pSnapEnv.ClearSnapAgents()

        ' Add the one we've just created
        ' pSnapEnv.AddSnapAgent(pFtrSnapAgent)
        pSnapEnv.SnapTolerance = fSnappingTolerance


        ' Refresh the contents of the Snapping window
        ' to reflect our changes
        pID.Value = "esriEditor.SnappingWindow"
        pSnapWindow = fEditor.FindExtension(pID)
        pSnapWindow.RefreshContents()




    End Sub
    Public Sub ZoomInCenter(ByVal app As IApplication, ByVal envelope As IEnvelope)
        Dim pMxDocument As IMxDocument
        Dim pActiveView As IActiveView
        Dim pDisplayTransform As ESRI.ArcGIS.Display.IDisplayTransformation
        Dim pEnvelope As IEnvelope
        Dim pCenterPoint As IPoint

        pMxDocument = app.Document
        pActiveView = pMxDocument.FocusMap
        pDisplayTransform = pActiveView.ScreenDisplay.DisplayTransformation
        'pEnvelope = pDisplayTransform.VisibleBounds
        pEnvelope = envelope
        'In this case, we could have set pEnvelope to IActiveView::Extent
        'Set pEnvelope = pActiveView.Extent
        pCenterPoint = New ESRI.ArcGIS.Geometry.Point

        pCenterPoint.X = ((pEnvelope.XMax - pEnvelope.XMin) / 2) + pEnvelope.XMin
        pCenterPoint.Y = ((pEnvelope.YMax - pEnvelope.YMin) / 2) + pEnvelope.YMin

        If pEnvelope.Width < 2000 Then
            pEnvelope.Width = 5000
        Else
            pEnvelope.Width = pEnvelope.Width + 2000
        End If
        If pEnvelope.Height < 2000 Then
            pEnvelope.Height = 5000
        Else
            pEnvelope.Height = pEnvelope.Height + 500
        End If
        pEnvelope.CenterAt(pCenterPoint)

        pDisplayTransform.VisibleBounds = pEnvelope
        pActiveView.Refresh()
    End Sub
    Public Function ExpandEnvelope(ByVal Height As Double, ByVal Width As Double, ByVal envelope As IEnvelope) As IEnvelope

        Dim pCenterPoint As IPoint
        Dim pEnvelope As IEnvelope
        pEnvelope = envelope
        pCenterPoint = New ESRI.ArcGIS.Geometry.Point

        pCenterPoint.X = ((pEnvelope.XMax - pEnvelope.XMin) / 2) + pEnvelope.XMin
        pCenterPoint.Y = ((pEnvelope.YMax - pEnvelope.YMin) / 2) + pEnvelope.YMin
        pEnvelope.Width = pEnvelope.Width + Width
        pEnvelope.Height = pEnvelope.Height + Height
        pEnvelope.CenterAt(pCenterPoint)
        ExpandEnvelope = pEnvelope



    End Function

    Public Sub SetSnappingEnv(ByVal fEditor As IEditor, ByVal fc As IFeatureClass, ByVal HitType As esriGeometryHitPartType)

        Dim pFtrSnapAgent As IFeatureSnapAgent
        Dim pSnapEnv As ISnapEnvironment
        Dim pSnapWindow As ISnappingWindow
        Dim pID As New UID

        pSnapEnv = fEditor
        pFtrSnapAgent = New FeatureSnap
        With pFtrSnapAgent
            .FeatureClass = fc
            .HitType = HitType   'And esriGeometryPartVertex
            '.HitType = esriGeometryPartBoundary

        End With

        ' Clear all current snapping agents
        'pSnapEnv.ClearSnapAgents()

        ' Add the one we've just created
        pSnapEnv.AddSnapAgent(pFtrSnapAgent)
        'pSnapEnv.SnapTolerance = fSnappingTolerance


        ' Refresh the contents of the Snapping window
        ' to reflect our changes
        pID.Value = "esriEditor.SnappingWindow"
        pSnapWindow = fEditor.FindExtension(pID)
        pSnapWindow.RefreshContents()




    End Sub

    Public Sub ClearSnapAgents(ByVal fEditor As IEditor)
        Dim pSnapEnv As ISnapEnvironment
        pSnapEnv = fEditor
        pSnapEnv.ClearSnapAgents()
    End Sub
    Public Sub SetSnappingTolerance2(ByVal envelope As IEnvelope, ByVal mxdoc As IMxDocument, ByVal editor As IEditor, ByRef refEnv As IEnvelope)


        Dim dblSnappingTolerance As Double

        dblSnappingTolerance = mxdoc.FocusMap.MapScale / 100

        SetSnappingTolerance(editor, dblSnappingTolerance)




        refEnv = mxdoc.ActiveView.Extent.Envelope

    End Sub

    Public Function GetFeatureINode(ByVal pFeature As IFeature, ByVal pJunctions As IFeatureLayer) As Long
        'created by:  Stefan Coe
        'created:     100507


        '*******inputs: These must be the FIRST TWO LAYERS IN THE TOC IN ORDER:
        '*******Edges to be checked must be SELECTED
        '   edges           =TransRefEdges
        '   junctions       =TransRefJunctions


        'Checks every ****SELECTED**** edge to see if the edge's digitized direction corresponds to its IJ Direction
        'Does this by commparing the Junction under the first Vertex of each Edge to the Edge's Inode
        'If the JunctionID and the Edge's Inode are the same, they are in same direction.
        'If not, the edge's digitized direciton needs to be flipped

        'Creates/Appends a Text File "c:\FlipEdgeLogFile.txt"  that lists edges that cannot be evaluted because more
        'than one junction was found under/near the edges first vertex. See below





        Dim pFilter As ISpatialFilter

        Dim pJunctionFeature As IFeature
        Dim pPolygon As IPolygon
        Dim pCurve As ICurve
        Dim pJCursor As IFeatureCursor
        Dim pTopoOp As ITopologicalOperator
        Dim pVertexPoint As IPoint
        Dim junctID As Long
        Dim JunctionIDIndex As Long



        pCurve = pFeature.Shape
        pVertexPoint = pCurve.FromPoint 'This is the edge's first vertex
        pTopoOp = pVertexPoint
        pPolygon = pTopoOp.Buffer(0.2) 'small buffer around the vertex
        pFilter = New SpatialFilter

        'Setting up a spatial filter that will select the junction(s) (should only be one. more than one or none
        'will result in an error and the edge is flagged and skipped) that intersect with the buffer (pPolygon)
        'of the first vertex.
        With pFilter
            .Geometry = pPolygon 'the vertex buffer
            .GeometryField = pJunctions.FeatureClass.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
        End With
        'EdgeIDIndex = pEdge.Fields.FindField("PSRCEdgeID")
        'EdgeID = pEdge.Value(EdgeIDIndex)
        If pJunctions.FeatureClass.FeatureCount(pFilter) > 1 Then 'The buffer selected more than one junction. Flag and go to the next Edge

            'MessageBox.Show("More than 1 possible INode Junction Found. Please enter the right JNode using the Attribute Editor", "Multiple Junctions Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            GetFeatureINode = -1

        ElseIf pJunctions.FeatureClass.FeatureCount(pFilter) = 0 Then 'The buffer selected no junctions. Flag and go to the next Edge

            'MessageBox.Show("An INode could not be found for this project. If an INode junction does not exist for this project, please create one", "No INode Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            GetFeatureINode = 0
        Else 'we have one junction. Evaluate to see if its PSRCJuntID is the same as the Inode of the edge
            pJCursor = pJunctions.FeatureClass.Search(pFilter, True)
            pJunctionFeature = pJCursor.NextFeature 'get the selected junction. Should only be one
            JunctionIDIndex = pJunctions.FeatureClass.FindField("PSRCJUNCTID")
            junctID = pJunctionFeature.Value(JunctionIDIndex) 'Get the Junction ID
            GetFeatureINode = junctID



        End If
        Marshal.ReleaseComObject(pFilter)

        pCurve = Nothing
        pVertexPoint = Nothing 'This is the edge's first vertex
        pTopoOp = Nothing
        pPolygon = Nothing 'small buffer around the vertex
        pFilter = Nothing




    End Function

    Public Function GetFeatureJNode(ByVal pFeature As IFeature, ByVal pJunctions As IFeatureLayer) As Single

        'created by:  Stefan Coe
        'created:     100507


        '*******inputs: These must be the FIRST TWO LAYERS IN THE TOC IN ORDER:
        '*******Edges to be checked must be SELECTED
        '   edges           =TransRefEdges
        '   junctions       =TransRefJunctions


        'Checks every ****SELECTED**** edge to see if the edge's digitized direction corresponds to its IJ Direction
        'Does this by commparing the Junction under the first Vertex of each Edge to the Edge's Inode
        'If the JunctionID and the Edge's Inode are the same, they are in same direction.
        'If not, the edge's digitized direciton needs to be flipped

        'Creates/Appends a Text File "c:\FlipEdgeLogFile.txt"  that lists edges that cannot be evaluted because more
        'than one junction was found under/near the edges first vertex. See below






        Dim pFilter As ISpatialFilter
        Dim pJunctionFeature As IFeature
        Dim pPolygon As IPolygon
        Dim pCurve As ICurve
        Dim pJCursor As IFeatureCursor
        Dim pTopoOp As ITopologicalOperator
        Dim pVertexPoint As IPoint
        Dim junctID As Long
        Dim JunctionIDIndex As Long




        pCurve = pFeature.Shape
        pVertexPoint = pCurve.ToPoint  'This is the edge's first vertex
        pTopoOp = pVertexPoint
        pPolygon = pTopoOp.Buffer(0.2) 'small buffer around the vertex
        pFilter = New SpatialFilter

        'Setting up a spatial filter that will select the junction(s) (should only be one. more than one or none
        'will result in an error and the edge is flagged and skipped) that intersect with the buffer (pPolygon)
        'of the first vertex.
        With pFilter
            .Geometry = pPolygon 'the vertex buffer
            .GeometryField = pJunctions.FeatureClass.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
        End With
        'EdgeIDIndex = pEdge.Fields.FindField("PSRCEdgeID")
        'EdgeID = pEdge.Value(EdgeIDIndex)
        If pJunctions.FeatureClass.FeatureCount(pFilter) > 1 Then 'The buffer selected more than one junction. Flag and go to the next Edge

            'MessageBox.Show("More than 1 possible JNode Junction Found. Please enter the right JNode using the Attribute Editor", "Multiple Junctions Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            GetFeatureJNode = -1
        ElseIf pJunctions.FeatureClass.FeatureCount(pFilter) = 0 Then 'The buffer selected no junctions. Flag and go to the next Edge

            'MessageBox.Show("A JNode could not be found for this project. If a JNode junction does not exist for this project, please create one", "No JNode Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            GetFeatureJNode = 0

        Else 'we have one junction. Evaluate to see if its PSRCJuntID is the same as the Inode of the edge
            pJCursor = pJunctions.FeatureClass.Search(pFilter, True)
            pJunctionFeature = pJCursor.NextFeature 'get the selected junction. Should only be one
            JunctionIDIndex = pJunctions.FeatureClass.FindField("PSRCJUNCTID")
            junctID = pJunctionFeature.Value(JunctionIDIndex) 'Get the Junction ID
            GetFeatureJNode = junctID



        End If

        Marshal.ReleaseComObject(pFilter)

        pCurve = Nothing
        pVertexPoint = Nothing 'This is the edge's first vertex
        pTopoOp = Nothing
        pPolygon = Nothing 'small buffer around the vertex
        pFilter = Nothing




    End Function

    Public Function GetTurnMovementEdgeIDs(ByVal pFeature As IFeature, ByVal Field As Integer, ByVal pEdges As IFeatureLayer, ByVal pJunctions As IFeatureLayer) As Long
        'created by:  Stefan Coe
        'created:     100507


        '*******inputs: These must be the FIRST TWO LAYERS IN THE TOC IN ORDER:
        '*******Edges to be checked must be SELECTED
        '   edges           =TransRefEdges
        '   junctions       =TransRefJunctions


        'Checks every ****SELECTED**** edge to see if the edge's digitized direction corresponds to its IJ Direction
        'Does this by commparing the Junction under the first Vertex of each Edge to the Edge's Inode
        'If the JunctionID and the Edge's Inode are the same, they are in same direction.
        'If not, the edge's digitized direciton needs to be flipped

        'Creates/Appends a Text File "c:\FlipEdgeLogFile.txt"  that lists edges that cannot be evaluted because more
        'than one junction was found under/near the edges first vertex. See below






        Dim pFilter As ISpatialFilter

        Dim pJunctionFeature As IFeature
        Dim pPolygon As IPolygon
        Dim pCurve As ICurve
        Dim pJCursor As IFeatureCursor
        Dim pTopoOp As ITopologicalOperator
        Dim pTMFromPoint As IPoint
        Dim pTMToPoint As IPoint
        Dim pToPoly As IPolygon
        Dim pFromPoly As IPolygon
        Dim junctID As Long
        Dim JunctionIDIndex As Long
        Dim pEdgeCursor As IFeatureCursor
        Dim pEdgeFeature As IFeature
        Dim pJunctCursor As IFeatureCursor
        Dim pJunctFeature As IFeature
        Dim pRO As IRelationalOperator
        Dim pEdgeAttributes As EdgeAttributes
        Dim pEdgeIDIndex As Integer
        Dim pEdgeID As Long
        Dim pCollection As New Collection

        Dim pJunctIndex As Integer
        Dim pJunctID As Long
        Dim pJunctPoint As IPoint




        'need to select the edges under the Turn feature:
        pCurve = pFeature.Shape
        pTMFromPoint = pCurve.FromPoint
        pTMToPoint = pCurve.ToPoint

        pTopoOp = pTMFromPoint
        pFromPoly = pTopoOp.Buffer(0.5)

        pTopoOp = pTMToPoint
        pToPoly = pTopoOp.Buffer(0.5)

        
        'get all the junctions that intersect with the turn feature. Exit if not 3. 
        If Field = 1 Then
            pFilter = New SpatialFilter
            With pFilter
                .Geometry = pCurve
                .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                .GeometryField = pJunctions.FeatureClass.ShapeFieldName
            End With
            If pJunctions.FeatureClass.FeatureCount(pFilter) < 3 Then
                GetTurnMovementEdgeIDs = -1
                Exit Function
            ElseIf pJunctions.FeatureClass.FeatureCount(pFilter) > 3 Then
                GetTurnMovementEdgeIDs = -2
                Exit Function
            End If



            pJunctCursor = pJunctions.FeatureClass.Search(pFilter, True)
            pJunctFeature = pJunctCursor.NextFeature
            'loop through each junction
            Do Until pJunctFeature Is Nothing
                pJunctIndex = pJunctFeature.Fields.FindField("PSRCJunctID")
                pJunctID = pJunctFeature.Value(pJunctIndex)
                'pJunctPoint = pJunctFeature.Shape
                'pTopoOp = pJunctPoint
                'pPolygon = pTopoOp.Buffer(0.3)

                pRO = pJunctFeature.Shape
                If pRO.Equals(pTMFromPoint) = False And pRO.Equals(pTMToPoint) = False Then
                    GetTurnMovementEdgeIDs = pJunctID


                End If
                pJunctFeature = pJunctCursor.NextFeature
            Loop
            'pRO = pFeature.Shape


        Else
            pFilter = New SpatialFilter
            With pFilter
                .Geometry = pCurve
                .SpatialRel = esriSpatialRelEnum.esriSpatialRelContains
                .GeometryField = pEdges.FeatureClass.ShapeFieldName
            End With
            If pEdges.FeatureClass.FeatureCount(pFilter) <> 2 Then
                GetTurnMovementEdgeIDs = -1
                Exit Function
            End If
            pEdgeCursor = pEdges.FeatureClass.Search(pFilter, True)
            pEdgeFeature = pEdgeCursor.NextFeature
            Do Until pEdgeFeature Is Nothing
                pEdgeIDIndex = pEdgeFeature.Fields.FindField("PSRCEdgeID")
                pEdgeID = pEdgeFeature.Value(pEdgeIDIndex)
                pRO = pEdgeFeature.Shape
                If pRO.Touches(pTMFromPoint) And Field = 2 Then
                    GetTurnMovementEdgeIDs = pEdgeID

                ElseIf pRO.Touches(pTMToPoint) And Field = 3 Then
                    GetTurnMovementEdgeIDs = pEdgeID
                End If


                pEdgeFeature = pEdgeCursor.NextFeature
            Loop

            'MessageBox.Show(pEdges.FeatureClass.FeatureCount(pFilter))

        End If





    End Function
    Public Function GetTurnMovementEdgeIDs2(ByVal pFeature As IFeature, ByVal pEdges As IFeatureLayer, ByVal pJunctions As IFeatureLayer) As Turns

        'created by:  Stefan Coe
        'created:     100507


        '*******inputs: These must be the FIRST TWO LAYERS IN THE TOC IN ORDER:
        '*******Edges to be checked must be SELECTED
        '   edges           =TransRefEdges
        '   junctions       =TransRefJunctions


        'Checks every ****SELECTED**** edge to see if the edge's digitized direction corresponds to its IJ Direction
        'Does this by commparing the Junction under the first Vertex of each Edge to the Edge's Inode
        'If the JunctionID and the Edge's Inode are the same, they are in same direction.
        'If not, the edge's digitized direciton needs to be flipped

        'Creates/Appends a Text File "c:\FlipEdgeLogFile.txt"  that lists edges that cannot be evaluted because more
        'than one junction was found under/near the edges first vertex. See below






        Dim pFilter As ISpatialFilter

        Dim pJunctionFeature As IFeature
        Dim pPolygon As IPolygon
        Dim pCurve As ICurve
        Dim pJCursor As IFeatureCursor
        Dim pTopoOp As ITopologicalOperator
        Dim pTMFromPoint As IPoint
        Dim pTMToPoint As IPoint
        Dim pToPoly As IPolygon
        Dim pFromPoly As IPolygon
        Dim junctID As Long
        Dim JunctionIDIndex As Long
        Dim pEdgeCursor As IFeatureCursor
        Dim pEdgeFeature As IFeature
        Dim pJunctCursor As IFeatureCursor
        Dim pJunctFeature As IFeature
        Dim pRO As IRelationalOperator
        Dim pEdgeAttributes As EdgeAttributes
        Dim pEdgeIDIndex As Integer
        Dim pEdgeID As Long
        Dim pCollection As New Collection
        Dim pJunctionIDCollection As New Collection
        Dim pTurnBuffer As IPolygon


        Dim pJunctIndex As Integer
        Dim pFromJunctID As Long
        Dim pJunctPoint As IPoint
        Dim x As Integer
        Dim fromEdgeID As Integer
        Dim toEdgeID As Integer
        Dim pTurnJunctID As Integer
        Dim pEdge As RefEdgeAttributes
        GetTurnMovementEdgeIDs2.errorCode = 0

        'get the endpoints of the turn feature. Select coresponding junctions. 
        pCurve = pFeature.Shape
        pTMFromPoint = pCurve.FromPoint
        pTMToPoint = pCurve.ToPoint

        pTopoOp = pTMFromPoint
        pFromPoly = pTopoOp.Buffer(0.5)
        'pCollection.Add(pFromPoly)

        'pTopoOp = pTMToPoint
        ' pToPoly = pTopoOp.Buffer(0.5)
        'pCollection.Add(pToPoly)

        'select junction underneath from point. Get junction id.  
        pFilter = New SpatialFilter
        With pFilter
            .Geometry = pFromPoly
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            .GeometryField = pJunctions.FeatureClass.ShapeFieldName
        End With

        If pJunctions.FeatureClass.FeatureCount(pFilter) <> 1 Then
            'error, no junction or too many junctions under turn feature from node

            GetTurnMovementEdgeIDs2.errorCode = 1

            Exit Function
        Else
            pJunctCursor = pJunctions.FeatureClass.Search(pFilter, True)
            pJunctFeature = pJunctCursor.NextFeature
            pJunctIndex = pJunctFeature.Fields.FindField("PSRCJunctID")
            pFromJunctID = pJunctFeature.Value(pJunctIndex)

        End If


        'select the edges under the Turn feature & get the edge ids:
        pFilter = New SpatialFilter
        pTopoOp = pCurve
        'pMap.MapUnits = esriUnits.esriPoints
        pTurnBuffer = pTopoOp.Buffer(5)
        With pFilter
            .Geometry = pTurnBuffer
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelContains
            .GeometryField = pEdges.FeatureClass.ShapeFieldName
        End With
        If pEdges.FeatureClass.FeatureCount(pFilter) <> 2 Then
            GetTurnMovementEdgeIDs2.errorCode = 2

            Exit Function
        End If
        pEdgeCursor = pEdges.FeatureClass.Search(pFilter, True)
        pEdgeFeature = pEdgeCursor.NextFeature
        Do Until pEdgeFeature Is Nothing
            pEdge = New RefEdgeAttributes(pEdgeFeature)
            pRO = pEdgeFeature.Shape
            'find the from edge
            If pRO.Touches(pTMFromPoint) Then
                'get the fromEdge ID
                fromEdgeID = pEdge.EdgeID
                If pEdge.INode = pFromJunctID Then
                    'the turn junction has to be it's JNode
                    pTurnJunctID = pEdge.JNode
                Else
                    'the turn junction has to be it's INode
                    pTurnJunctID = pEdge.INode

                End If
            Else
                'we have the toEdge, get the id
                toEdgeID = pEdge.EdgeID


            End If

            pEdgeFeature = pEdgeCursor.NextFeature
        Loop
        GetTurnMovementEdgeIDs2.PSRCJunctID = pTurnJunctID
        GetTurnMovementEdgeIDs2.FrEdgeID = fromEdgeID
        GetTurnMovementEdgeIDs2.ToEdgeID = toEdgeID





























    End Function
    Structure Turns
        'struct to hold some turn attributes. 
        Public PSRCJunctID As Long

        Public FrEdgeID As Long
        Public ToEdgeID As Long
        Public errorCode As Long
    End Structure
    Public Function UnionEnvelope(ByVal features As Collection) As IEnvelope
        Dim x As Integer
        Dim pFeature As IFeature
        For x = 1 To features.Count
            If x = 1 Then
                pFeature = CType(features.Item(x), IFeature)
                UnionEnvelope = pFeature.Shape.Envelope

            Else
                pFeature = CType(features.Item(x), IFeature)
                UnionEnvelope.Union(pFeature.Shape.Envelope)
            End If

            'pObject = CType(features(x), IObject)

           

        Next


    End Function

    Public Sub CheckIJNode2(ByVal pMap As IMap)
        Dim pEnumFeat As IEnumFeature
        Dim pSelection As ISelection
        Dim pFeature As IFeature
        pSelection = pMap.FeatureSelection
        pEnumFeat = pSelection
        pFeature = pEnumFeat.Next
        Do Until pFeature Is Nothing
            If pFeature.Class.AliasName = "TransRefEdges" Then
                MessageBox.Show("yes")
            End If




            pFeature = pEnumFeat.Next
        Loop
    End Sub


    




    Public Sub SwitchModeAttributes(ByVal ModeAttributes As ITable, ByVal EdgeID As Long)
        Dim pTable As ITable
        Dim pRow As IRow
        Dim pFilter As IQueryFilter
        Dim pCursor As ICursor
        Dim pClsModeAttributes As ModeAttributes


        pTable = ModeAttributes
        pFilter = New QueryFilter
        pFilter.WhereClause = "PSRCEDGEID = " & EdgeID

        pCursor = pTable.Search(pFilter, True)
        pRow = pCursor.NextRow

        pClsModeAttributes = New ModeAttributes(pRow)
        'pClsModeAttributes.ModeAttributeRow = pRow

        Dim IJBIKELANES As Long
        Dim IJFFS As Double
        Dim IJLANECAPGP As Long
        Dim IJLANECAPHOV As Long
        Dim IJLANESGPADJUST As Double
        Dim IJLANESGPAM As Long
        Dim IJLANESGPEV As Long
        Dim IJLANESGPMD As Long
        Dim IJLANESGPNI As Long
        Dim IJLANESGPPM As Long
        Dim IJLANESHOVAM As Long
        Dim IJLANESHOVEV As Long
        Dim IJLANESHOVMD As Long
        Dim IJLANESHOVNI As Long
        Dim IJLANESHOVPM As Long
        Dim IJLANESTK As Long
        Dim IJLANESTR As Long
        Dim IJSIDEWALKS As Long
        Dim IJSPEEDLIMIT As Long
        Dim IJVDFUNC As Long

        Dim JIBIKELANES As Long
        Dim JIFFS As Long
        Dim JILANECAPGP As Long
        Dim JILANECAPHOV As Long
        Dim JILANESGPADJUST As Double
        Dim JILANESGPAM As Long
        Dim JILANESGPEV As Long
        Dim JILANESGPMD As Long
        Dim JILANESGPNI As Long
        Dim JILANESGPPM As Long
        Dim JILANESHOVAM As Long
        Dim JILANESHOVEV As Long
        Dim JILANESHOVMD As Long
        Dim JILANESHOVNI As Long
        Dim JILANESHOVPM As Long
        Dim JILANESTK As Long
        Dim JILANESTR As Long
        Dim JISIDEWALKS As Long
        Dim JISPEEDLIMIT As Long
        Dim JIVDFUNC As Long


        IJBIKELANES = pClsModeAttributes.JIBIKELANES
        'IJFFS = pClsModeAttributes.JIFFS
        IJLANECAPGP = pClsModeAttributes.JILANECAPGP
        IJLANECAPHOV = pClsModeAttributes.JILANECAPHOV
        IJLANESGPADJUST = pClsModeAttributes.JILANESGPADJUST
        IJLANESGPAM = pClsModeAttributes.JILANESGPAM
        IJLANESGPEV = pClsModeAttributes.JILANESGPEV
        IJLANESGPMD = pClsModeAttributes.JILANESGPMD
        IJLANESGPNI = pClsModeAttributes.JILANESGPNI
        IJLANESGPPM = pClsModeAttributes.JILANESGPPM
        IJLANESHOVAM = pClsModeAttributes.JILANESHOVAM
        IJLANESHOVEV = pClsModeAttributes.JILANESHOVEV
        IJLANESHOVMD = pClsModeAttributes.JILANESHOVMD
        IJLANESHOVNI = pClsModeAttributes.JILANESHOVNI
        IJLANESHOVPM = pClsModeAttributes.JILANESHOVPM
        IJLANESTK = pClsModeAttributes.JILANESTK
        IJLANESTR = pClsModeAttributes.JILANESTR
        IJSIDEWALKS = pClsModeAttributes.JISIDEWALKS
        IJSPEEDLIMIT = pClsModeAttributes.JISPEEDLIMIT
        IJVDFUNC = pClsModeAttributes.JIVDFUNC


        JIBIKELANES = pClsModeAttributes.IJBIKELANES
        'JIFFS = pClsModeAttributes.IJFFS
        JILANECAPGP = pClsModeAttributes.IJLANECAPGP
        JILANECAPHOV = pClsModeAttributes.IJLANECAPHOV
        JILANESGPADJUST = pClsModeAttributes.IJLANESGPADJUST
        JILANESGPAM = pClsModeAttributes.IJLANESGPAM
        JILANESGPEV = pClsModeAttributes.IJLANESGPEV
        JILANESGPMD = pClsModeAttributes.IJLANESGPMD
        JILANESGPNI = pClsModeAttributes.IJLANESGPNI
        JILANESGPPM = pClsModeAttributes.IJLANESGPPM
        JILANESHOVAM = pClsModeAttributes.IJLANESHOVAM
        JILANESHOVEV = pClsModeAttributes.IJLANESHOVEV
        JILANESHOVMD = pClsModeAttributes.IJLANESHOVMD
        JILANESHOVNI = pClsModeAttributes.IJLANESHOVNI
        JILANESHOVPM = pClsModeAttributes.IJLANESHOVPM
        JILANESTK = pClsModeAttributes.IJLANESTK
        JILANESTR = pClsModeAttributes.IJLANESTR
        JISIDEWALKS = pClsModeAttributes.IJSIDEWALKS
        JISPEEDLIMIT = pClsModeAttributes.IJSPEEDLIMIT
        JIVDFUNC = pClsModeAttributes.IJVDFUNC




        pClsModeAttributes.IJBIKELANES = IJBIKELANES
        'pClsModeAttributes.IJFFS = IJFFS
        pClsModeAttributes.IJLANECAPGP = IJLANECAPGP
        pClsModeAttributes.IJLANECAPHOV = IJLANECAPHOV
        pClsModeAttributes.IJLANESGPADJUST = IJLANESGPADJUST
        pClsModeAttributes.IJLANESGPAM = IJLANESGPAM
        pClsModeAttributes.IJLANESGPEV = IJLANESGPEV
        pClsModeAttributes.IJLANESGPMD = IJLANESGPMD
        pClsModeAttributes.IJLANESGPNI = IJLANESGPNI
        pClsModeAttributes.IJLANESGPPM = IJLANESGPPM
        pClsModeAttributes.IJLANESHOVAM = IJLANESHOVAM
        pClsModeAttributes.IJLANESHOVEV = IJLANESHOVEV
        pClsModeAttributes.IJLANESHOVMD = IJLANESHOVMD
        pClsModeAttributes.IJLANESHOVNI = IJLANESHOVNI
        pClsModeAttributes.IJLANESHOVPM = IJLANESHOVPM
        pClsModeAttributes.IJLANESTK = IJLANESTK
        pClsModeAttributes.IJLANESTR = IJLANESTR
        pClsModeAttributes.IJSIDEWALKS = IJSIDEWALKS
        pClsModeAttributes.IJSPEEDLIMIT = IJSPEEDLIMIT
        pClsModeAttributes.IJVDFUNC = IJVDFUNC


        pClsModeAttributes.JIBIKELANES = JIBIKELANES
        'pClsModeAttributes.JIFFS = JIFFS
        pClsModeAttributes.JILANECAPGP = JILANECAPGP
        pClsModeAttributes.JILANECAPHOV = JILANECAPHOV
        pClsModeAttributes.JILANESGPADJUST = JILANESGPADJUST
        pClsModeAttributes.JILANESGPAM = JILANESGPAM
        pClsModeAttributes.JILANESGPEV = JILANESGPEV
        pClsModeAttributes.JILANESGPMD = JILANESGPMD
        pClsModeAttributes.JILANESGPNI = JILANESGPNI
        pClsModeAttributes.JILANESGPPM = JILANESGPPM
        pClsModeAttributes.JILANESHOVAM = JILANESHOVAM
        pClsModeAttributes.JILANESHOVEV = JILANESHOVEV
        pClsModeAttributes.JILANESHOVMD = JILANESHOVMD
        pClsModeAttributes.JILANESHOVNI = JILANESHOVNI
        pClsModeAttributes.JILANESHOVPM = JILANESHOVPM
        pClsModeAttributes.JILANESTK = JILANESTK
        pClsModeAttributes.JILANESTR = JILANESTR
        pClsModeAttributes.JISIDEWALKS = JISIDEWALKS
        pClsModeAttributes.JISPEEDLIMIT = JISPEEDLIMIT
        pClsModeAttributes.JIVDFUNC = JIVDFUNC







    End Sub
    Public Sub WriteLogLine(ByVal strOutput As String)
        'writes strOutput as a line to the debug log file
        FileOpen(1, "c:\SumStats.txt", OpenMode.Append, OpenAccess.ReadWrite, OpenShare.Default)
        PrintLine(1, strOutput)
        FileClose(1)
    End Sub
    Public Sub CopyModeAttributes(ByVal edgeColl As Collection, ByVal m_application As IApplication)
        Dim pEdge1 As IFeature
        Dim pedge2 As IFeature
        Dim pEdge1Attributes As EdgeAttributes
        Dim pEdge2Attributes As EdgeAttributes
        Dim pEdge1ModeAttributes As ModeAttributes
        Dim pEdge2ModeAttributes As ModeAttributes
        Dim m_ModeAttributes As IStandaloneTable
        Dim pRow1, pRow2 As IRow
        Dim pFilter As IQueryFilter
        Dim pCursor As ICursor




        pEdge1 = CType(edgeColl.Item(1), IFeature)
        pedge2 = CType(edgeColl.Item(2), IFeature)
        pEdge1Attributes = New EdgeAttributes(pEdge1)
        pEdge2Attributes = New EdgeAttributes(pedge2)

        pFilter = New QueryFilter
        pFilter.WhereClause = "PSRCEDGEID = " & pEdge1Attributes.PSRCEdgeID
        m_ModeAttributes = getStandaloneTable(g_Schema & g_ModeAttributes, m_application)
        pCursor = m_ModeAttributes.Table.Search(pFilter, True)
        pRow1 = pCursor.NextRow
        pEdge1ModeAttributes = New ModeAttributes(pRow1)

        pFilter = New QueryFilter
        pFilter.WhereClause = "PSRCEDGEID = " & pEdge2Attributes.PSRCEdgeID
        pCursor = m_ModeAttributes.Table.Search(pFilter, True)
        pRow2 = pCursor.NextRow
        pEdge2ModeAttributes = New ModeAttributes(pRow2)

        pEdge2ModeAttributes.IJBIKELANES = pEdge1ModeAttributes.IJBIKELANES
        'pEdge2ModeAttributes.IJFFS = pEdge1ModeAttributes.IJFFS
        pEdge2ModeAttributes.IJLANECAPGP = pEdge1ModeAttributes.IJLANECAPGP
        pEdge2ModeAttributes.IJLANECAPHOV = pEdge1ModeAttributes.IJLANECAPHOV
        pEdge2ModeAttributes.IJLANESGPADJUST = pEdge1ModeAttributes.IJLANESGPADJUST
        pEdge2ModeAttributes.IJLANESGPAM = pEdge1ModeAttributes.IJLANESGPAM
        pEdge2ModeAttributes.IJLANESGPEV = pEdge1ModeAttributes.IJLANESGPEV
        pEdge2ModeAttributes.IJLANESGPMD = pEdge1ModeAttributes.IJLANESGPMD
        pEdge2ModeAttributes.IJLANESGPNI = pEdge1ModeAttributes.IJLANESGPNI
        pEdge2ModeAttributes.IJLANESGPPM = pEdge1ModeAttributes.IJLANESGPPM
        pEdge2ModeAttributes.IJLANESHOVAM = pEdge1ModeAttributes.IJLANESHOVAM
        pEdge2ModeAttributes.IJLANESHOVEV = pEdge1ModeAttributes.IJLANESHOVEV
        pEdge2ModeAttributes.IJLANESHOVMD = pEdge1ModeAttributes.IJLANESHOVMD
        pEdge2ModeAttributes.IJLANESHOVNI = pEdge1ModeAttributes.IJLANESHOVNI
        pEdge2ModeAttributes.IJLANESHOVPM = pEdge1ModeAttributes.IJLANESHOVPM
        pEdge2ModeAttributes.IJLANESTK = pEdge1ModeAttributes.IJLANESTK
        pEdge2ModeAttributes.IJLANESTR = pEdge1ModeAttributes.IJLANESTR
        pEdge2ModeAttributes.IJSIDEWALKS = pEdge1ModeAttributes.IJSIDEWALKS
        pEdge2ModeAttributes.IJSPEEDLIMIT = pEdge1ModeAttributes.IJSPEEDLIMIT
        pEdge2ModeAttributes.IJVDFUNC = pEdge1ModeAttributes.IJVDFUNC
        pEdge2ModeAttributes.JIBIKELANES = pEdge1ModeAttributes.JIBIKELANES
        'pEdge2ModeAttributes.JIFFS = pEdge1ModeAttributes.JIFFS
        pEdge2ModeAttributes.JILANECAPGP = pEdge1ModeAttributes.JILANECAPGP
        pEdge2ModeAttributes.JILANECAPHOV = pEdge1ModeAttributes.JILANECAPHOV
        pEdge2ModeAttributes.JILANESGPADJUST = pEdge1ModeAttributes.JILANESGPADJUST
        pEdge2ModeAttributes.JILANESGPAM = pEdge1ModeAttributes.JILANESGPAM
        pEdge2ModeAttributes.JILANESGPEV = pEdge1ModeAttributes.JILANESGPEV
        pEdge2ModeAttributes.JILANESGPMD = pEdge1ModeAttributes.JILANESGPMD
        pEdge2ModeAttributes.JILANESGPNI = pEdge1ModeAttributes.JILANESGPNI
        pEdge2ModeAttributes.JILANESGPPM = pEdge1ModeAttributes.JILANESGPPM
        pEdge2ModeAttributes.JILANESHOVAM = pEdge1ModeAttributes.JILANESHOVAM
        pEdge2ModeAttributes.JILANESHOVEV = pEdge1ModeAttributes.JILANESHOVEV
        pEdge2ModeAttributes.JILANESHOVMD = pEdge1ModeAttributes.JILANESHOVMD
        pEdge2ModeAttributes.JILANESHOVNI = pEdge1ModeAttributes.JILANESHOVNI
        pEdge2ModeAttributes.JILANESHOVPM = pEdge1ModeAttributes.JILANESHOVPM
        pEdge2ModeAttributes.JILANESTK = pEdge1ModeAttributes.JILANESTK
        pEdge2ModeAttributes.JILANESTR = pEdge1ModeAttributes.JILANESTR
        pEdge2ModeAttributes.JISIDEWALKS = pEdge1ModeAttributes.JISIDEWALKS
        pEdge2ModeAttributes.JISPEEDLIMIT = pEdge1ModeAttributes.JISPEEDLIMIT
        pEdge2ModeAttributes.JIVDFUNC = pEdge1ModeAttributes.JIVDFUNC
        'pEdge2ModeAttributes.NEW1 = pEdge1ModeAttributes.NEW1
        'pEdge2ModeAttributes.OLD_EDGEID = pEdge1ModeAttributes.OLD_EDGEID
        'pEdge2ModeAttributes.PSRC_E2ID = pEdge1ModeAttributes.PSRC_E2ID
































    End Sub
    Public Sub TempModeAttributes(ByVal fromRow As IRow, ByVal toRow As IRow, ByVal m_application As IApplication)
        Dim pEdge1 As IFeature
        Dim pedge2 As IFeature
        Dim pEdge1Attributes As EdgeAttributes
        Dim pEdge2Attributes As EdgeAttributes
        Dim pEdge1ModeAttributes As ModeAttributes
        Dim pEdge2ModeAttributes As ModeAttributes
        Dim m_ModeAttributes As IStandaloneTable
        Dim pRow1, pRow2 As IRow
        Dim pFilter As IQueryFilter
        Dim pCursor As ICursor




        'pEdge1 = CType(edgeColl.Item(1), IFeature)
        'pedge2 = CType(edgeColl.Item(2), IFeature)
        'pEdge1Attributes = New EdgeAttributes(pEdge1)
        'pEdge2Attributes = New EdgeAttributes(pedge2)

        'pFilter = New QueryFilter
        'pFilter.WhereClause = "PSRCEDGEID = " & pEdge1Attributes.PSRCEdgeID
        'm_ModeAttributes = getStandaloneTable(g_Schema & g_ModeAttributes, m_application)
        'pCursor = m_ModeAttributes.Table.Search(pFilter, True)
        'pRow1 = pCursor.NextRow
        pEdge1ModeAttributes = New ModeAttributes(fromRow)

        'pFilter = New QueryFilter
        'pFilter.WhereClause = "PSRCEDGEID = " & pEdge2Attributes.PSRCEdgeID
        'pCursor = m_ModeAttributes.Table.Search(pFilter, True)
        'pRow2 = pCursor.NextRow
        pEdge2ModeAttributes = New ModeAttributes(toRow)
        pEdge2ModeAttributes.PSRCEDGEID = pEdge1ModeAttributes.PSRCEDGEID

        pEdge2ModeAttributes.IJBIKELANES = pEdge1ModeAttributes.IJBIKELANES
        'pEdge2ModeAttributes.IJFFS = pEdge1ModeAttributes.IJFFS
        pEdge2ModeAttributes.IJLANECAPGP = pEdge1ModeAttributes.IJLANECAPGP
        pEdge2ModeAttributes.IJLANECAPHOV = pEdge1ModeAttributes.IJLANECAPHOV
        pEdge2ModeAttributes.IJLANESGPADJUST = pEdge1ModeAttributes.IJLANESGPADJUST
        pEdge2ModeAttributes.IJLANESGPAM = pEdge1ModeAttributes.IJLANESGPAM
        pEdge2ModeAttributes.IJLANESGPEV = pEdge1ModeAttributes.IJLANESGPEV
        pEdge2ModeAttributes.IJLANESGPMD = pEdge1ModeAttributes.IJLANESGPMD
        pEdge2ModeAttributes.IJLANESGPNI = pEdge1ModeAttributes.IJLANESGPNI
        pEdge2ModeAttributes.IJLANESGPPM = pEdge1ModeAttributes.IJLANESGPPM
        pEdge2ModeAttributes.IJLANESHOVAM = pEdge1ModeAttributes.IJLANESHOVAM
        pEdge2ModeAttributes.IJLANESHOVEV = pEdge1ModeAttributes.IJLANESHOVEV
        pEdge2ModeAttributes.IJLANESHOVMD = pEdge1ModeAttributes.IJLANESHOVMD
        pEdge2ModeAttributes.IJLANESHOVNI = pEdge1ModeAttributes.IJLANESHOVNI
        pEdge2ModeAttributes.IJLANESHOVPM = pEdge1ModeAttributes.IJLANESHOVPM
        pEdge2ModeAttributes.IJLANESTK = pEdge1ModeAttributes.IJLANESTK
        pEdge2ModeAttributes.IJLANESTR = pEdge1ModeAttributes.IJLANESTR
        pEdge2ModeAttributes.IJSIDEWALKS = pEdge1ModeAttributes.IJSIDEWALKS
        pEdge2ModeAttributes.IJSPEEDLIMIT = pEdge1ModeAttributes.IJSPEEDLIMIT
        pEdge2ModeAttributes.IJVDFUNC = pEdge1ModeAttributes.IJVDFUNC
        pEdge2ModeAttributes.JIBIKELANES = pEdge1ModeAttributes.JIBIKELANES
        'pEdge2ModeAttributes.JIFFS = pEdge1ModeAttributes.JIFFS
        pEdge2ModeAttributes.JILANECAPGP = pEdge1ModeAttributes.JILANECAPGP
        pEdge2ModeAttributes.JILANECAPHOV = pEdge1ModeAttributes.JILANECAPHOV
        pEdge2ModeAttributes.JILANESGPADJUST = pEdge1ModeAttributes.JILANESGPADJUST
        pEdge2ModeAttributes.JILANESGPAM = pEdge1ModeAttributes.JILANESGPAM
        pEdge2ModeAttributes.JILANESGPEV = pEdge1ModeAttributes.JILANESGPEV
        pEdge2ModeAttributes.JILANESGPMD = pEdge1ModeAttributes.JILANESGPMD
        pEdge2ModeAttributes.JILANESGPNI = pEdge1ModeAttributes.JILANESGPNI
        pEdge2ModeAttributes.JILANESGPPM = pEdge1ModeAttributes.JILANESGPPM
        pEdge2ModeAttributes.JILANESHOVAM = pEdge1ModeAttributes.JILANESHOVAM
        pEdge2ModeAttributes.JILANESHOVEV = pEdge1ModeAttributes.JILANESHOVEV
        pEdge2ModeAttributes.JILANESHOVMD = pEdge1ModeAttributes.JILANESHOVMD
        pEdge2ModeAttributes.JILANESHOVNI = pEdge1ModeAttributes.JILANESHOVNI
        pEdge2ModeAttributes.JILANESHOVPM = pEdge1ModeAttributes.JILANESHOVPM
        pEdge2ModeAttributes.JILANESTK = pEdge1ModeAttributes.JILANESTK
        pEdge2ModeAttributes.JILANESTR = pEdge1ModeAttributes.JILANESTR
        pEdge2ModeAttributes.JISIDEWALKS = pEdge1ModeAttributes.JISIDEWALKS
        pEdge2ModeAttributes.JISPEEDLIMIT = pEdge1ModeAttributes.JISPEEDLIMIT
        pEdge2ModeAttributes.JIVDFUNC = pEdge1ModeAttributes.JIVDFUNC
        'pEdge2ModeAttributes.NEW1 = pEdge1ModeAttributes.NEW1
        'pEdge2ModeAttributes.OLD_EDGEID = pEdge1ModeAttributes.OLD_EDGEID
        'pEdge2ModeAttributes.PSRC_E2ID = pEdge1ModeAttributes.PSRC_E2ID
































    End Sub


End Module
