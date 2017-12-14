Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Controls
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geodatabase

<ComClass(cmdTopoFix.ClassId, cmdTopoFix.InterfaceId, cmdTopoFix.EventsId), _
 ProgId("PSRCEditingTools.cmdTopoFix")> _
Public NotInheritable Class cmdTopoFix
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "bdef4d84-8b18-4455-ab1c-e8b118ae4151"
    Public Const InterfaceId As String = "c02e0c00-d47d-49f5-92dc-c8a6a4948147"
    Public Const EventsId As String = "b7cc477e-fbb5-49a5-9093-1574c6696b6f"
#End Region

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region


    Private m_application As IApplication
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
    
    Private m_cmdStartEditing As New PSRCEditingTools.cmdStartEditing
    Private m_FeatAdded As IFeature
    Private refedgeFeature As IFeature


    



    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "PSRCEditingTools"  'localizable text 
        MyBase.m_caption = "Topology Validate and Fix"   'localizable text 
        MyBase.m_message = "Topology Validate and Fix"   'localizable text 
        MyBase.m_toolTip = "Use to Validate and Fix Topology Errors" 'localizable text 
        MyBase.m_name = "PSRCEditingTools.cmdTopoFix"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        Try
            'TODO: change bitmap name if necessary
            Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try


    End Sub


    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            m_application = CType(hook, IApplication)

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If
        'Get the editor, set editing events interfaces to it 
        Dim uid As New UID
        uid.Value = "esriEditor.Editor"
        m_editor = CType(m_application.FindExtensionByCLSID(uid), IEditor)
        m_pEditEvents = m_editor

        'AddHandler (CType(m_editor, IEditEvents2_Event).OnVertexAdded), AddressOf OnVertexAdded

        'm_pEditEvents2 = m_editor

        m_editSketch = m_editor

        m_editSketch2 = m_editor

        pSnap = m_editor

        Dim uidCmd As New UIDClass
        Dim cmd As ICommandItem

        uidCmd.Value = "{001386c4-ce20-4882-bd6c-9826b25015ff}"
        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)
        m_cmdStartEditing = cmd.Command


        ' TODO:  Add other initialization code
    End Sub

    Public Overrides ReadOnly Property Enabled() As Boolean

        Get

            'If m_editor.EditState = esriEditState.esriStateEditing Then
            If m_cmdStartEditing.m_HasClicked = True Then
                Return True

                'ElseIf m_editor.EditState = esriEditState.esriStateNotEditing Then
            ElseIf m_cmdStartEditing.m_HasClicked = False Then
                Return False

            End If



        End Get
    End Property

    Public Overrides Sub OnClick()
        'get the ComboBox
        Dim uidCmd As New UIDClass
        Dim cmd As ICommandItem

        'get access to the combo box
        uidCmd.Value = "{20A94A18-61FC-4516-A924-3C7AD4C9EC8E}"
        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)
        m_TLComboBoxControl = cmd.Command

        'get access to cmdStartEditing (must do this to access public variables declared in that class)
        uidCmd.Value = "{001386c4-ce20-4882-bd6c-9826b25015ff}"
        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)
        m_cmdStartEditing = cmd.Command

        m_mxDoc = m_application.Document

        m_activeView = m_mxDoc.ActiveView

        m_map = m_activeView.FocusMap
        g_Schema = checkWS(m_editor, m_application)

        m_RefEdges = GetFeatureLayer(GlobalConstants.g_Schema & g_RefEdge, m_application)
        m_Junctions = GetFeatureLayer(g_Schema & g_RefJunct, m_application)
        m_Projects = GetFeatureLayer(g_Schema & g_ProjectRoutes, m_application)

        m_pRefEdgesTable = m_RefEdges.FeatureClass
        m_pRefJunctionsTable = m_Junctions.FeatureClass
        m_pProjectsTable = m_Projects.FeatureClass










        Topologyfix3()
        'End If
        ' m_application = Nothing

        'TODO: Add cmdTopoFix.OnClick implementation
    End Sub
    Private Sub Topologyfix3()
        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor


        m_cmdStartEditing.m_boolSketchEdit = False

        x = 1
        Dim Done As Boolean
        Done = False
        Dim pFeature As IFeature
        'Dim msgResponse As MessageDialog
        Dim topoRuleContainter As ITopologyRuleContainer
        Dim env As ESRI.ArcGIS.Geometry.IEnvelope

        Dim pID As New UID
        Dim pCmdItm As ICommandItem






        Dim topoUiD As New UID
        Dim SaveEditsUID As New UID
        'Dim cmd As ICommandItem
        Dim refEdgeCollection As New Collection
        Dim p_obj As IObject
        'System.Diagnostics.Debugger.Break()



        topoUiD.Value = "esriEditorExt.TopologyExtension"
        Dim topologyExt As ESRI.ArcGIS.EditorExt.ITopologyExtension
        topologyExt = m_application.FindExtensionByCLSID(topoUiD)

        'MsgBox topologyExt.ActiveErrorCount
        Dim topology As ESRI.ArcGIS.Geodatabase.ITopology
        topology = topologyExt.CurrentTopology
        topoRuleContainter = topology

        'StopWorkspaceEditOperation(m_application, m_cmdStartEditing.pWrkspc)

        StartSDEWorkspaceEditorOperation(m_cmdStartEditing.pWrkspc, m_application)





        Dim geoDS As ESRI.ArcGIS.Geodatabase.IGeoDataset

        'Dim boolTopoErrors As Boolean

        geoDS = topology

        'validate topology to the extent of the digitized feature 
        ' env = topology.DirtyArea(topology.FeatureDataset


        'env = GetDirtyAreas(topology)


        'If env.IsEmpty Then
        'MessageBox.Show("There are no more errors")
        'm_activeView.Refresh()
        'Exit Sub
        'End If
        'MessageBox.Show(m_cmdStartEditing.m_AreaToValidate.Height & " " & m_cmdStartEditing.m_AreaToValidate.Width)
        If m_cmdStartEditing.m_AreaToValidate Is Nothing Then
            MessageBox.Show("This tool is used to fix topology errors when a Project, Edge or Junction has been added", "Nothing Digitized to Check", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            m_cmdStartEditing.m_AreaToValidate = ExpandEnvelope(500, 500, m_cmdStartEditing.m_AreaToValidate)
        ElseIf m_cmdStartEditing.m_AreaToValidate.Width < 1 And m_cmdStartEditing.m_AreaToValidate.Height < 1 Then
            m_cmdStartEditing.m_AreaToValidate = ExpandEnvelope(500, 500, m_cmdStartEditing.m_AreaToValidate)
            'Else
            ' m_cmdStartEditing.m_AreaToValidate = ExpandEnvelope(10, 10, m_cmdStartEditing.m_AreaToValidate)
        End If
        env = m_cmdStartEditing.m_AreaToValidate
        StartSDEWorkspaceEditorOperation(m_cmdStartEditing.pWrkspc, m_application)

        ValidateTopology(topology, env)


        ' m_activeView.Refresh()

        'StopWorkspaceEditOperation(m_cmdStartEditing.m_application)


        Dim errorContainer As ESRI.ArcGIS.Geodatabase.IErrorFeatureContainer

        errorContainer = topology

        Dim eErrorFeat As IEnumTopologyErrorFeature

        'Set eErrorFeat = errorContainer.ErrorFeaturesByRuleType(geoDS.SpatialReference, esriTopologyRuleType.esriTRTLineEndpointCoveredByPoint, Nothing, True, False)



        eErrorFeat = errorContainer.ErrorFeaturesByRuleType(geoDS.SpatialReference, esriTopologyRuleType.esriTRTAny, env, True, False)



        'MessageBox.Show(ex.ToString & vbCrLf & vbCrLf & "Please contact Stefan Coe if problem persists.", "Data Catalog", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)




        Dim topoError As ITopologyErrorFeature
        topoError = eErrorFeat.Next

        If topoError Is Nothing Then
            MessageBox.Show("There are no more errors")
            m_activeView.Refresh()
            System.Windows.Forms.Cursor.Current = Cursors.Default
            Exit Sub

        End If


        'We want to take care of junctions first, edges second and split edges last
        Do Until topoError Is Nothing
            'MsgBox(topoError.TopologyRuleType)
            If topoError.TopologyRuleType = esriTopologyRuleType.esriTRTLineEndpointCoveredByPoint Then
                pFeature = topoError
                env = pFeature.Extent.Envelope
                topologyExt.ClearActiveErrors(ESRI.ArcGIS.EditorExt.esriTEEventHint.esriTENone)

                topologyExt.AddActiveError(topoError, ESRI.ArcGIS.EditorExt.esriTEEventHint.esriTENone)
                ZoomInCenter(m_application, env)
                If MessageBox.Show("A junction needs to added. Click Yes to Add or No to mark error as an exception", "Junction Needed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    'pFeature = topoError
                    'm_mxDoc.ActiveView.Extent = pFeature.Extent.Envelope
                    'MessageBox.Show("zoomed in")

                    'Dim lngNewJunctionID As Long
                    'GetLargestID(m_pRefJunctionsTable, "PSRCJunctID", lngNewJunctionID)
                    'lngNewJunctionID = lngNewJunctionID + 1
                    'StartEditor(m_Junctions, m_application)

                    p_obj = EdgesEndPointMustBeCoveredByJunctionsTopologyFix(m_Junctions, pFeature, m_application)
                    StopWorkspaceEditOperation(m_application, m_cmdStartEditing.pWrkspc)
                    'test(x) = "JCT" & CStr(lngNewJunctionID)
                    x = x + 1
                    'MsgBox m_pAttDict.Count

                    'reset the target layer
                    SetTargetLayer()
                    SelectFeature(m_map, p_obj, m_mxDoc, m_activeView)
                    pID.Value = "{44276914-98C1-11D1-8464-0000F875B9C6}:0"

                    pCmdItm = m_application.Document.CommandBars.Find(pID)

                    pCmdItm.Execute()

                    CheckIJNode()
                    'ValidateTopology(topology, env)
                    'OpenAttributeEditor()

                    m_mxDoc.ActiveView.Refresh()
                    System.Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                Else


                    topoRuleContainter.PromoteToRuleException(topoError)

                    System.Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            End If
            topoError = eErrorFeat.Next
        Loop

        eErrorFeat = errorContainer.ErrorFeaturesByRuleType(geoDS.SpatialReference, esriTopologyRuleType.esriTRTAny, env, True, False)
        topoError = eErrorFeat.Next

        'Try


        Do Until topoError Is Nothing

            'MsgBox(topoError.TopologyRuleType)
            If topoError.TopologyRuleType = esriTopologyRuleType.esriTRTLineCoveredByLineClass Then
                pFeature = topoError


                Dim indexDateCreated As Integer
                'Dim intOriginOID As Long
                Dim pQF As IQueryFilter
                'Dim pFC As IFeatureCursor
                pQF = New QueryFilter
                pQF.WhereClause = "OID = " & topoError.OriginOID
                indexDateCreated = m_Projects.FeatureClass.GetFeature(topoError.OriginOID).Fields.FindField("DateCreated")

                env = pFeature.Extent.Envelope

                topologyExt.ClearActiveErrors(ESRI.ArcGIS.EditorExt.esriTEEventHint.esriTENone)

                topologyExt.AddActiveError(topoError, ESRI.ArcGIS.EditorExt.esriTEEventHint.esriTENone)
                ZoomInCenter(m_application, env)


                'Format(Now, "dddd, d MMM yyyy")
                'indexDateCreated = topoError.OriginOIDpFeature.Fields.FindField("DATECREATED")
                'MessageBox.Show(CType(m_Projects.FeatureClass.GetFeature(topoError.OriginOID).Value(indexDateCreated), Date).Date)
                Dim mdate As Date

                If Not m_Projects.FeatureClass.GetFeature(topoError.OriginOID).Value(indexDateCreated) Is DBNull.Value Then

                    mdate = CType(m_Projects.FeatureClass.GetFeature(topoError.OriginOID).Value(indexDateCreated), Date)
                    If Format(mdate, "dddd, d MMM yyyy") <> Format(Now, "dddd, d MMM yyyy") Then
                        topoRuleContainter.PromoteToRuleException(topoError)
                    Else
                        If MessageBox.Show("A Reference Edge needs to added. Click Yes to Add or No to mark error as an exception", "Reference Edge Needed", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then



                            'Dim lngNewEdgeID As Long
                            'GetLargestID(m_pRefEdgesTable, "PSRCEdgeID", lngNewEdgeID)
                            'lngNewEdgeID = lngNewEdgeID + 1

                            p_obj = ProjectsMustCoverRefEdgesTopologyFix(m_RefEdges, pFeature, m_application, m_cmdStartEditing.pWrkspc)
                            StopWorkspaceEditOperation(m_application, m_cmdStartEditing.pWrkspc)
                            refedgeFeature = CType(p_obj, IFeature)

                            'test(x) = "REF" & CStr(lngNewEdgeID)
                            'x = x + 1
                            'MsgBox m_pAttDict.Count
                            'reset the target layer
                            SetTargetLayer()
                            SelectFeature(m_map, p_obj, m_mxDoc, m_activeView)
                            pID.Value = "{44276914-98C1-11D1-8464-0000F875B9C6}:0"

                            pCmdItm = m_application.Document.CommandBars.Find(pID)

                            pCmdItm.Execute()
                            CheckIJNode()
                            'ValidateTopology(topology, env)
                            'OpenAttributeEditor()
                            m_mxDoc.ActiveView.Refresh()
                            System.Windows.Forms.Cursor.Current = Cursors.Default
                            Exit Sub
                        Else
                            topoRuleContainter.PromoteToRuleException(topoError)
                            System.Windows.Forms.Cursor.Current = Cursors.Default
                            Exit Sub

                        End If
                    End If
                Else
                    If m_cmdStartEditing.boolOutsideError = True Then
                        MessageBox.Show("An error was found with a project unrelated to the current Edits. Please run the Topology Fix button again. If the same message appears, all of the errors associated with the current edits should be fixed.", "Topology Error Unrelated to Current Edits", MessageBoxButtons.OK, MessageBoxIcon.Warning)


                    End If
                    m_cmdStartEditing.boolOutsideError = True
                    'MessageBox.Show(topoError.IsException)
                    topoRuleContainter.PromoteToRuleException(topoError)
                    Exit Do
                End If
               




               

                
            End If
            topoError = eErrorFeat.Next
        Loop

        eErrorFeat = errorContainer.ErrorFeaturesByRuleType(geoDS.SpatialReference, esriTopologyRuleType.esriTRTAny, env, True, False)
        topoError = eErrorFeat.Next

        Do Until topoError Is Nothing
            pFeature = topoError
            env = pFeature.Extent.Envelope
            topologyExt.ClearActiveErrors(ESRI.ArcGIS.EditorExt.esriTEEventHint.esriTENone)

            topologyExt.AddActiveError(topoError, ESRI.ArcGIS.EditorExt.esriTEEventHint.esriTENone)
            ZoomInCenter(m_application, env)
            If topoError.TopologyRuleType = esriTopologyRuleType.esriTRTPointCoveredByLineEndpoint Then
                If MessageBox.Show("A Reference Edge needs to be split. Clicking Yes will split the edge and automatically save exisitng edits. Clicking No will mark error as an exception", "Edge Needs to be Split", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then


                    refEdgeCollection = SplitEdgeTopologyFix2(m_RefEdges, m_Junctions, m_cmdStartEditing.m_AreaToValidate, m_application, m_cmdStartEditing.pWrkspc)


                    CopyModeAttributes(refEdgeCollection, m_application)
                    StopWorkspaceEditOperation(m_application, m_cmdStartEditing.pWrkspc)
                    'reset the target layer
                    SetTargetLayer()
                    'm_cmdStartEditing.m_AreaToValidate = UnionEnvelope(refEdgeCollection)
                    ' MessageBox.Show(m_cmdStartEditing.m_AreaToValidate.Height & m_cmdStartEditing.m_AreaToValidate.Width)
                    SelectFeatures(m_map, refEdgeCollection, m_mxDoc, m_activeView)

                    pID.Value = "{44276914-98C1-11D1-8464-0000F875B9C6}:0"

                    pCmdItm = m_application.Document.CommandBars.Find(pID)

                    pCmdItm.Execute()
                    CheckIJNode()


                    'ValidateTopology(topology, env)
                    'OpenAttributeEditor()
                    'SaveEditsUID.Value = "{59D2AFD2-9EA2-11D1-9165-0080C718DF97}"
                    'cmd = m_application.Document.CommandBars.Find(SaveEditsUID, False, False)
                    'cmd.Execute()
                    m_mxDoc.ActiveView.Refresh()
                    'MessageBox.Show("Rembember to Edit Attributes for both Edges if necessary", "Edit Reminder!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    System.Windows.Forms.Cursor.Current = Cursors.Default
                    'pCmdItm.Execute()

                    Exit Sub
                Else
                    topoRuleContainter.PromoteToRuleException(topoError)

                End If
            End If

            topoError = eErrorFeat.Next
        Loop


        'If Not m_cmdStartEditing.projFeature Is Nothing Then
        'StopWorkspaceEditOperation(m_application, pWrkspc)
        'StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
        'Dim pProject2 As New ProjectAttributes(m_cmdStartEditing.projFeature)
        'If pProject2.HasINode = False Then
        'GetFeatureINode(m_cmdStartEditing.projFeature, m_Junctions)
        ' End If
        'If pProject2.HasJNode = False Then
        'GetFeatureJNode(m_cmdStartEditing.projFeature, m_Junctions)
        'End If
        ' pProject2 = Nothing
        'StopWorkspaceEditOperation(m_application, pWrkspc)
        'End If
        'Catch ex As Exception
        'MessageBox.Show(ex.ToString & vbCrLf & vbCrLf & "Please contact Stefan Coe if problem persists.", "Data Catalog", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        'End Try
        System.Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub m_pEditEvents_OnChangeFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) Handles m_pEditEvents.OnChangeFeature

    End Sub


    Private Sub m_pEditEvents_OnCreateFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject) Handles m_pEditEvents.OnCreateFeature

        m_FeatAdded = obj

        'Dim pID As New UID
        'Dim pCmdItm As ICommandItem
        'Dim a As Integer
        'Dim flayer As IFeatureLayer
        'Dim pFeature As IFeature
        'pFeature = obj
        'm_mxDoc.ActiveView.Extent = pFeature.Extent.Envelope
        'm_mxDoc.ActiveView.Refresh()
        'MessageBox.Show("fired")



        'For a = 0 To m_map.LayerCount - 1
        'flayer = m_map.Layer(a)
        'If flayer.FeatureClass Is obj.Table Then

        'm_map.SelectFeature(m_mxDoc.FocusMap.Layer(a), obj)
        'pID.Value = "{44276914-98C1-11D1-8464-0000F875B9C6}:0"
        'pCmdItm = m_application.Document.CommandBars.Find(pID)
        'pCmdItm.Execute()

        'Exit For
        'End If

        'Next

    End Sub


    Public Sub OpenAttributeEditor()
        Dim pID As New UID
        Dim a As Integer
        Dim flayer As IFeatureLayer

        Dim pCmdItm As ICommandItem
        For a = 0 To m_map.LayerCount - 1
            flayer = m_map.Layer(a)
            If flayer.FeatureClass Is m_FeatAdded.Table Then
                m_map.ClearSelection()
                m_map.SelectFeature(m_mxDoc.FocusMap.Layer(a), m_FeatAdded)
                m_activeView.Refresh()

                Exit For
            End If
        Next

        pID.Value = "{44276914-98C1-11D1-8464-0000F875B9C6}:0"
        pCmdItm = m_application.Document.CommandBars.Find(pID)
        pCmdItm.Execute()

    End Sub
    Public Sub SetTargetLayer()
        Select Case m_TLComboBoxControl.ComboBox1.SelectedItem

            Case Is = "Add Projects"

                SetEditlayer(m_Projects, m_application)

            Case Is = "Add Junctions"

                SetEditlayer(m_Junctions, m_application)

            Case Is = "Edit Projects"

                SetEditlayer(m_Projects, m_application)

            Case Is = "Add TransRefEdges"

                SetEditlayer(m_RefEdges, m_application)
        End Select


    End Sub
    Public Sub CheckIJNode()
        If Not m_cmdStartEditing.projFeature Is Nothing Then
            'StopWorkspaceEditOperation(m_application, pWrkspc)
            'StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
            Dim pProject2 As New ProjectAttributes(m_cmdStartEditing.projFeature)

            If pProject2.HasINode = False Then
                StartSDEWorkspaceEditorOperation(m_cmdStartEditing.pWrkspc, m_application)
                pProject2.ProjectINode = GetFeatureINode(m_cmdStartEditing.projFeature, m_Junctions)
                StopWorkspaceEditOperation(m_application, m_cmdStartEditing.pWrkspc)
            End If
            If pProject2.HasJNode = False Then
                StartSDEWorkspaceEditorOperation(m_cmdStartEditing.pWrkspc, m_application)
                pProject2.ProjectJNode = GetFeatureJNode(m_cmdStartEditing.projFeature, m_Junctions)
                StopWorkspaceEditOperation(m_application, m_cmdStartEditing.pWrkspc)
            End If
            pProject2 = Nothing
            'StopWorkspaceEditOperation(m_application, pWrkspc)
        End If
        If Not refedgeFeature Is Nothing Then
            Dim pEdge2 As New RefEdgeAttributes(refedgeFeature)

            If pEdge2.HasINode = False Then
                StartSDEWorkspaceEditorOperation(m_cmdStartEditing.pWrkspc, m_application)
                pEdge2.INode = GetFeatureINode(refedgeFeature, m_Junctions)
                StopWorkspaceEditOperation(m_application, m_cmdStartEditing.pWrkspc)
            End If
            If pEdge2.HasJNode = False Then
                StartSDEWorkspaceEditorOperation(m_cmdStartEditing.pWrkspc, m_application)
                pEdge2.JNode = GetFeatureJNode(refedgeFeature, m_Junctions)
                StopWorkspaceEditOperation(m_application, m_cmdStartEditing.pWrkspc)
            End If
            pEdge2 = Nothing
        End If


    End Sub
   

End Class



