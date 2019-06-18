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
Public Class frmNodeScanner
    Public m_application As IApplication

    Private WithEvents m_pEditEvents As Editor
    'Private WithEvents m_pEditEvents2 As EditEvents2
    Private m_editSketch As IEditSketch
    Private m_editSketch2 As IEditSketch2
    Private m_editor As Editor
    Public m_HasClicked As Boolean
    Public m_IsChecked As Boolean

    Private m_mxDoc As IMxDocument
    Private m_activeView As IActiveView
    Public m_map As IMap
    Public m_RefEdges As IFeatureLayer
    Public m_Projects As IFeatureLayer
    Public m_Junctions As IFeatureLayer
    Public m_TurnMovements As IFeatureLayer
    Public m_ModeAttributes As ITable
    Public m_TransitPoints As IFeatureLayer


    Private test(150) As String
    Public x As Integer

    Public m_pRefEdgesTable As ITable
    Public m_pRefJunctionsTable As ITable
    Public m_pProjectsTable As ITable

    Public m_TLComboBoxControl As New PSRCEditingTools.TLComboBoxControl
    Private m_lngVertexCount As Long
    Private pSnap As ISnapEnvironment
    Private boolLastVertexSnapped As Boolean
    Private objCollection As New Collection

    'Private m_cmdStartEditing As New PSRCEditingTools.cmdStartEditing
    Private m_FeatAdded As IFeature
    Private refedgeFeature As IFeature
    Private pDataset As IDataset
    Private pWrkspc As IWorkspace
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        m_mxDoc = m_application.Document
        Dim uid2 As New UID
        uid2.Value = "esriEditor.Editor"
        m_editor = CType(m_application.FindExtensionByCLSID(uid2), IEditor)
        m_activeView = m_mxDoc.ActiveView

        m_map = m_activeView.FocusMap
        g_Schema = checkWS(m_editor, m_application)

        m_RefEdges = GetFeatureLayer(GlobalConstants.g_Schema & g_RefEdge, m_application)
        m_Junctions = GetFeatureLayer(g_Schema & g_RefJunct, m_application)
        m_Projects = GetFeatureLayer(g_Schema & g_ProjectRoutes, m_application)
        m_ModeAttributes = getStandaloneTable(g_Schema & g_ModeAttributes, m_application)
        m_TurnMovements = GetFeatureLayer(GlobalConstants.g_Schema & g_TurnMovements, m_application)


        pDataset = CType(m_RefEdges.FeatureClass, IDataset)

        pWrkspc = pDataset.Workspace
        CheckIJNode(m_map)

    End Sub
    Public Sub CheckIJNode(ByVal pMap As IMap)
        Dim pEnumFeat As IEnumFeature
        Dim pSelection As ISelection
        Dim pFeature As IFeature
        Dim pINode As Long
        Dim pJNode As Long
        Dim TurnCollection As New Collection

        pSelection = pMap.FeatureSelection
        pEnumFeat = pSelection
        pFeature = pEnumFeat.Next
        Do Until pFeature Is Nothing


            If pFeature.Class.AliasName = g_Schema & "ProjectRoutes" Then
                Dim pProject As New ProjectAttributes(pFeature)

                'If pProject.HasINode = False Then

                pINode = GetFeatureINode(pFeature, m_Junctions)
                If pINode = 0 Then
                    WriteLogLine("An INode Junction could not be found for Project ID: " & pProject.PROJRTEID)
                Else
                    StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
                    pProject.ProjectINode = pINode
                    StopWorkspaceEditOperation(m_application, pWrkspc)
                End If

                'If

                'If pProject.HasJNode = False Then
                pJNode = GetFeatureJNode(pFeature, m_Junctions)
                If pJNode = 0 Then
                    WriteLogLine("An JNode Junction could not be found for Project ID: " & pProject.PROJRTEID)
                Else
                    StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
                    pProject.ProjectJNode = pJNode
                    StopWorkspaceEditOperation(m_application, pWrkspc)
                End If

                'End If
            ElseIf pFeature.Class.AliasName = g_Schema & "TurnMovements" Then
                Dim pTurnAttributes As New TurnAttributes(pFeature)
                Dim pTurnStruct As New Turns
                Dim fromID As Long
                Dim toID As Long
                Dim JunctionID As Long


                pTurnStruct = GetTurnMovementEdgeIDs2(pFeature, m_RefEdges, m_Junctions)
                If pTurnStruct.errorCode = 1 Then
                    WriteLogLine("Error Finding More than one junction under from node" & pTurnAttributes.TurnID)
                ElseIf pTurnStruct.errorCode = 2 Then
                    WriteLogLine("Error finding exactly two edges under turn feature " & pTurnAttributes.TurnID)
                Else
                    StartSDEWorkspaceEditorOperation(pWrkspc, m_application)

                    pTurnAttributes.PSRCJunctID = pTurnStruct.PSRCJunctID
                    pTurnAttributes.FrEdgeID = pTurnStruct.FrEdgeID
                    pTurnAttributes.ToEdgeID = pTurnStruct.ToEdgeID
                    StopWorkspaceEditOperation(m_application, pWrkspc)
                End If



            ElseIf pFeature.Class.AliasName = g_Schema & "TransRefEdges" Then
                Dim pEdge As New RefEdgeAttributes(pFeature)

                'If pEdge.HasINode = False Then
                pINode = GetFeatureINode(pFeature, m_Junctions)
                pJNode = GetFeatureJNode(pFeature, m_Junctions)
                If pINode > 0 And pJNode > 0 And pINode = pJNode Then
                    WriteLogLine("I and J Junction same for edge " & pEdge.EdgeID)

                Else


                    If pINode = -1 Then
                        WriteLogLine("found more than 1 I junction for edge " & pEdge.EdgeID)
                        'MessageBox.Show("An INode Junction could not be found for Edge ID: " & pEdge.EdgeID)
                    End If
                    If pJNode = -1 Then
                        WriteLogLine("found more than 1 J junction for edge " & pEdge.EdgeID)
                        'MessageBox.Show("An INode Junction could not be found for Edge ID: " & pEdge.EdgeID)
                    End If

                    If pINode = 0 Then
                        'FileClose(300)
                        'Make sure log file is not already open

                        WriteLogLine("could not find I junction for edge " & pEdge.EdgeID)
                        'MessageBox.Show("An INode Junction could not be found for Edge ID: " & pEdge.EdgeID)
                    End If
                    If pJNode = 0 Then
                        'FileClose(300)
                        'Make sure log file is not already ope
                        WriteLogLine("could not find I junction for edge " & pEdge.EdgeID)
                        'MessageBox.Show("An INode Junction could not be found for Edge ID: " & pEdge.EdgeID)
                    End If

                    'If pEdge.INode <> pINode Then
                    'Nodes are switched:
                    If pEdge.INode = pJNode And pEdge.JNode = pINode Then
                        'FileOpen(1, "c:\SumStats.txt", OpenMode.Append, OpenAccess.ReadWrite, OpenShare.Default)
                        'PrintLine(1, "INode & JNode are Switched for edge " & pEdge.EdgeID)
                        'FileClose(1)
                        WriteLogLine("INode & JNode should be Switched for edge " & pEdge.EdgeID)
                        If chkSwitchNodes.Checked Then
                            pEdge.INode = pINode
                            pEdge.JNode = pJNode
                            'If chkSwitchAttributes.Checked Then
                            '    SwitchModeAttributes(m_ModeAttributes, pEdge.EdgeID)
                            'WriteLogLine("I & J Nodes + Attributes Switched for edge " & pEdge.EdgeID)
                            'End If

                            'FileOpen(1, "c:\SumStats.txt", OpenMode.Append, OpenAccess.ReadWrite, OpenShare.Default)
                            WriteLogLine("I & J Nodes Switched for edge " & pEdge.EdgeID)
                            'FileClose(1)
                        End If
                    Else
                        If pINode > 0 And pEdge.INode <> pINode Then
                            WriteLogLine("INode wrong for edge " & pEdge.EdgeID)
                            If chkFixIJNode.Checked Then
                                StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
                                pEdge.INode = pINode
                                StopWorkspaceEditOperation(m_application, pWrkspc)
                                WriteLogLine("INode changed for edge " & pEdge.EdgeID)
                            End If
                        End If

                        If pJNode > 0 And pEdge.JNode <> pJNode Then
                            WriteLogLine("JNode wrong for edge " & pEdge.EdgeID)
                            If chkFixIJNode.Checked Then
                                StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
                                pEdge.JNode = pJNode
                                StopWorkspaceEditOperation(m_application, pWrkspc)
                                WriteLogLine("JNode changed for edge " & pEdge.EdgeID)
                            End If
                        End If
                        'FileClose(1)
                        'Make sure log file is not already open


                        'StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
                        'pEdge.INode = pINode
                        'StopWorkspaceEditOperation(m_application, pWrkspc)
                        'End If

                        'End If
                        'End If



                        'If pEdge.HasJNode = False Then
                    End If


                End If



            End If




            pFeature = pEnumFeat.Next

            GC.Collect()
        Loop
        MessageBox.Show("IJ Nodes Checked for Selected TransRefEdges and Project Features", "IJ Nodes Checked", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub
   
    Public Sub CheckTransitPoints(ByVal pMap As IMap)
        Dim pEnumFeat As IEnumFeature
        Dim pSelection As ISelection
        Dim pFeature As IFeature
        Dim pINode As Long
        Dim pJNode As Long
        Dim TurnCollection As New Collection
        Dim pTP As TransitPointAttributes
        Dim pPoint As IPoint
        Dim pJunctID As Long

        pSelection = pMap.FeatureSelection
        pEnumFeat = pSelection
        pFeature = pEnumFeat.Next
        Try

        
            Do Until pFeature Is Nothing

                If pFeature.Class.AliasName = g_Schema & "TransitPoints" Then
                    pTP = New TransitPointAttributes(pFeature)



                    pPoint = pFeature.Shape
                    'getJunction(pPoint)
                    pJunctID = getJunctionID(pPoint, pFeature.OID)

                    If pJunctID <> pTP.PSRCJunctionID Then
                        WriteLogLine("JunctionID wrong for TransitPoint " & pFeature.OID)
                        StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
                        pTP.PSRCJunctionID = pJunctID
                        'pFeature.Store()
                        StopWorkspaceEditOperation(m_application, pWrkspc)
                        WriteLogLine("JunctionID changed for TransitPoint " & pFeature.OID)
                    End If

                    pFeature = pEnumFeat.Next
                End If

            Loop

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try







    End Sub
    Public Function getJunction(ByVal pPoint As IPoint) As IFeature
        Dim spatialFilter As ESRI.ArcGIS.Geodatabase.ISpatialFilter = New ESRI.ArcGIS.Geodatabase.SpatialFilterClass

        Dim pFCls As IFeatureClass
        pFCls = m_Junctions.FeatureClass
        'pFCls = GetFeatureLayer(g_Schema & g_RefJunct, m_application).FeatureClass
        'spatialFilter = New SpatialFilter
        With spatialFilter
            .Geometry = pPoint
            .GeometryField = pFCls.ShapeFieldName
            .SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
        End With

        Dim pFCS As IFeatureCursor
        Dim pFT As IFeature
        pFCS = pFCls.Search(spatialFilter, False)
        pFT = pFCS.NextFeature
        getJunction = pFT
        Marshal.ReleaseComObject(pFCS)
        pFCS = Nothing
        spatialFilter = Nothing
    End Function

    Private Function getJunctionID(ByVal pPoint As IPoint, tranistPointOID As Long) As Long
        Try


            Dim pFT As IFeature
            pFT = getJunction(pPoint)

            If Not pFT Is Nothing Then
                getJunctionID = pFT.Value(pFT.Fields.FindField(g_fldJctID))
            Else
                WriteLogLine("Could not find JunctionID for TransitPoint " & tranistPointOID)
            End If


            pFT = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        m_mxDoc = m_application.Document
        Dim uid2 As New UID
        uid2.Value = "esriEditor.Editor"
        m_editor = CType(m_application.FindExtensionByCLSID(uid2), IEditor)
        m_activeView = m_mxDoc.ActiveView

        m_map = m_activeView.FocusMap
        g_Schema = checkWS(m_editor, m_application)

        m_RefEdges = GetFeatureLayer(GlobalConstants.g_Schema & g_RefEdge, m_application)
        m_Junctions = GetFeatureLayer(g_Schema & g_RefJunct, m_application)
        m_Projects = GetFeatureLayer(g_Schema & g_ProjectRoutes, m_application)
        m_ModeAttributes = getStandaloneTable(g_Schema & g_ModeAttributes, m_application)
        m_TurnMovements = GetFeatureLayer(GlobalConstants.g_Schema & g_TurnMovements, m_application)
        m_TransitPoints = GetFeatureLayer(GlobalConstants.g_Schema & g_TransitPoints, m_application)


        pDataset = CType(m_RefEdges.FeatureClass, IDataset)

        pWrkspc = pDataset.Workspace
        'CheckIJNode(m_map)







        CheckTransitPoints(m_map)
    End Sub

    Private Sub btnFlipEdges_Click(sender As System.Object, e As System.EventArgs) Handles btnFlipEdges.Click

        m_mxDoc = m_application.Document
        Dim uid2 As New UID
        uid2.Value = "esriEditor.Editor"
        m_editor = CType(m_application.FindExtensionByCLSID(uid2), IEditor)
        m_activeView = m_mxDoc.ActiveView

        m_map = m_activeView.FocusMap
        g_Schema = checkWS(m_editor, m_application)

        m_RefEdges = GetFeatureLayer(GlobalConstants.g_Schema & g_RefEdge, m_application)
        m_Junctions = GetFeatureLayer(g_Schema & g_RefJunct, m_application)
        m_Projects = GetFeatureLayer(g_Schema & g_ProjectRoutes, m_application)
        m_ModeAttributes = getStandaloneTable(g_Schema & g_ModeAttributes, m_application)
        m_TurnMovements = GetFeatureLayer(GlobalConstants.g_Schema & g_TurnMovements, m_application)


        pDataset = CType(m_RefEdges.FeatureClass, IDataset)

        pWrkspc = pDataset.Workspace

        FlipEdges(m_map)
    End Sub
    Public Sub FlipEdges(ByVal pMap As IMap)
        Dim pEnumFeat As IEnumFeature
        Dim pSelection As ISelection
        Dim pFeature As IFeature
        Dim pINode As Long
        Dim pJNode As Long
        Dim TurnCollection As New Collection

        pSelection = pMap.FeatureSelection
        pEnumFeat = pSelection
        pFeature = pEnumFeat.Next
        Do Until pFeature Is Nothing
            If pFeature.Class.AliasName = g_Schema & "TransRefEdges" Then
                Dim pEdge As New RefEdgeAttributes(pFeature)
                Dim curve As ESRI.ArcGIS.Geometry.ICurve = TryCast(pFeature.Shape, ESRI.ArcGIS.Geometry.ICurve)
                StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
                curve.ReverseOrientation()
                pFeature.Shape = curve
                pFeature.Store()
                StopWorkspaceEditOperation(m_application, pWrkspc)
                WriteLogLine("Edge Flipped:  " & pEdge.EdgeID)
                pFeature = pEnumFeat.Next
            End If

        Loop
    End Sub

 

End Class