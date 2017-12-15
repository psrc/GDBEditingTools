Option Explicit On
Option Strict Off


Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Controls

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
Imports System.Reflection





<ComClass(cmdStartEditing.ClassId, cmdStartEditing.InterfaceId, cmdStartEditing.EventsId), _
 ProgId("PSRCEditingTools.cmdStartEditing")> _
Public NotInheritable Class cmdStartEditing
    Inherits BaseCommand
    'Implements IEditTask



#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "001386c4-ce20-4882-bd6c-9826b25015ff"
    Public Const InterfaceId As String = "31348630-4209-4c5f-bc79-f924a7600a9c"
    Public Const EventsId As String = "3565bda4-1cd7-4a80-bdd5-b7fb90a9b19a"
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
        ControlsCommands.Register(regKey)
        MxCommands.Register(regKey)
    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        ControlsCommands.Unregister(regKey)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region


    Public WithEvents PSRCStartEditor As TestEvents

    Private m_pEditEvents As Editor

    Private WithEvents m_pWatchDisplay As DisplayTransformation

    Private WithEvents m_pActiveView As Map

    Private dCreateFeature As IEditEvents_OnCreateFeatureEventHandler

    Private dVertexAdded As IEditEvents2_OnVertexAddedEventHandler

    'Private dSketchFinished As IEditEvents_OnSketchFinishedEventHandler

    'Private dEditingStopped As IEditEvents_OnStopEditingEventHandler






    Private m_editSketch2 As IEditSketch2

    Private m_application As IApplication

    Private m_editor As Editor

    Public m_HasClicked As Boolean

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

    Public m_tblBikeAttributes As ITable


    Public m_TLComboBoxControl As New PSRCEditingTools.TLComboBoxControl

    Private m_lngVertexCount As Long

    Private pSnap As ISnapEnvironment

    Private m_boolLastVertexSnapped As Boolean

    Public m_pEnvelope As ESRI.ArcGIS.Geometry.IEnvelope

    Public pWrkspc As IWorkspace

    Public m_topology As ESRI.ArcGIS.Geodatabase.ITopology

    Public m_FeatureAdded As IFeature

    Public m_boolSketchEdit As Boolean

    Public m_AreaToValidate As IEnvelope

    Private m_editorStopped As Boolean

    Public bProjHasINode As Boolean
    Public bProjHasJNode As Boolean
    Public projFeature As IFeature
    Public edgeFeature As IFeature
    Public boolOutsideError As Boolean
    Public m_ModeAttributes As IStandaloneTable
    Public sdePrefix As String





    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.

    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "PSRCEditingTools"  'localizable text 

        MyBase.m_caption = "Start Editing"   'localizable text 

        MyBase.m_message = "This should work in ArcMap/MapControl/PageLayoutControl"   'localizable text 

        MyBase.m_toolTip = "" 'localizable text 

        MyBase.m_name = "PSRCEditingTools.cmdStartEditing"  'unique id, non-localizable (e.g. "MyCategory_MyCommand")


    End Sub


    Public Overrides Sub OnCreate(ByVal hook As Object)
        'The following happpens as Arc Map is openign and creating components

        If Not hook Is Nothing Then
            m_application = CType(hook, IApplication)

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then

                MyBase.m_enabled = True

            Else

                MyBase.m_enabled = False

            End If
        End If

        'keeps track of start/stop editing
        m_HasClicked = False


    End Sub

    Public Overrides Sub OnClick()

        'Try


        SetUpEvents()

        'some boolean variables to keep track of state
        'this is used to make sure the sketch tool is not used to trace a line
        m_boolLastVertexSnapped = False

        'this is used to tell if the user is adding a feature by a sketch or a toplogical fix
        m_boolSketchEdit = True

        bProjHasINode = True
        bProjHasJNode = True


        'get the ComboBox so we know what kind of edits the user is going to make
        Dim uidCmd As New UIDClass

        Dim cmd As ICommandItem

        Dim pDataset As IDataset

        uidCmd.Value = "{20A94A18-61FC-4516-A924-3C7AD4C9EC8E}"

        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)

        m_TLComboBoxControl = CType(cmd.Command, TLComboBoxControl)



        'get a reference to the map
        m_mxDoc = CType(m_application.Document, IMxDocument)

        'Display transformation events
        m_pWatchDisplay = m_mxDoc.ActiveView.ScreenDisplay.DisplayTransformation

        m_activeView = m_mxDoc.ActiveView

        m_pEnvelope = m_activeView.Extent.Envelope



        m_map = m_activeView.FocusMap

        'ActiveView Events:
        m_pActiveView = m_map

        'get layers

        Dim pFWorkspaceLayer As IFeatureLayer
        Dim p
        Dim pFlayer As IFeatureLayer2
        Dim a As Integer
        For a = 0 To m_map.LayerCount - 1



            If TypeOf m_map.Layer(a) Is IFeatureLayer Then
                pFlayer = CType(m_map.Layer(a), IFeatureLayer2)
                If pFlayer.DataSourceType = "SDE Feature Class" Then
                    If m_map.Layer(a).Name.Contains(g_ProjectRoutes) Then 'make sure it is teh Project layer*****delete later?" Then
                        pFWorkspaceLayer = pFlayer
                        Exit For
                    End If
                End If
            End If

            'MessageBox.Show("Layer sde.PSRC.ProjectRoutes is missing. Please add it to the TOC.")
        Next

        If pFWorkspaceLayer Is Nothing Then
            MessageBox.Show("Layer sde.SDE.ProjectRoutes is missing. Please add it to the TOC.")
        End If


        sdePrefix = pFlayer.FeatureClass.AliasName.Replace("ProjectRoutes", "")


        pDataset = CType(pFWorkspaceLayer.FeatureClass, IDataset)

        pWrkspc = pDataset.Workspace

        'start the editor and/or a workspace edit operation 
        StartSDEWorkspaceEditorOperation(pWrkspc, m_application)

        SetSnappingTolerance2(m_pEnvelope, m_mxDoc, m_editor, m_pEnvelope)

        If checkWS(m_editor, m_application) = "-1" Then
            StopEditor(False, m_application)

        Else



            g_Schema = checkWS(m_editor, m_application)



            'm_ModeAttributes = getStandaloneTable(g_Schema & g_ModeAttributes, m_application)
            m_RefEdges = GetFeatureLayer(GlobalConstants.g_Schema & g_RefEdge, m_application)
            m_Junctions = GetFeatureLayer(g_Schema & g_RefJunct, m_application)
            m_Projects = GetFeatureLayer(g_Schema & g_ProjectRoutes, m_application)
            m_tblBikeAttributes = getStandaloneTable(g_Schema & g_BikeAttributes, m_application)

            SetEditlayer(m_Projects, m_application)

            ClearSnapAgents(m_editor)

            SetSnappingEnv(m_editor, m_RefEdges.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
            SetSnappingEnv(m_editor, m_Junctions.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)



            'layers, fix later
            'm_RefEdges = CType(m_map.Layer(0), IFeatureLayer)

            ' m_Junctions = CType(m_map.Layer(1), IFeatureLayer)

            'm_Projects = CType(m_map.Layer(2), IFeatureLayer)

            'tables, fix later
            'm_pRefEdgesTable = CType(m_RefEdges.FeatureClass, ITable)

            'm_pRefJunctionsTable = CType(m_Junctions.FeatureClass, ITable)

            'm_pProjectsTable = CType(m_Projects.FeatureClass, ITable)



            'select the target layer based on the selected item in the combo box
            Select Case m_TLComboBoxControl.ComboBox1.SelectedItem

                Case Is = "Add Projects"

                    SetEditlayer(m_Projects, m_application)
                    ClearSnapAgents(m_editor)
                    SetSnappingEnv(m_editor, m_RefEdges.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
                    SetSnappingEnv(m_editor, m_Junctions.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)



                Case Is = "Add Junctions"

                    SetEditlayer(m_Junctions, m_application)
                    ClearSnapAgents(m_editor)
                    SetSnappingEnv(m_editor, m_RefEdges.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
                    SetSnappingEnv(m_editor, m_Projects.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
                    SetSnappingEnv(m_editor, m_RefEdges.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
                    SetSnappingEnv(m_editor, m_Projects.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)


                Case Is = "Edit Projects"

                    SetEditlayer(m_Projects, m_application)

                Case Is = "Add TransRefEdges"

                    SetEditlayer(m_RefEdges, m_application)
                    ClearSnapAgents(m_editor)
                    SetSnappingEnv(m_editor, m_Projects.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
                    SetSnappingEnv(m_editor, m_RefEdges.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
                    SetSnappingEnv(m_editor, m_Junctions.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
                    SetEditlayer(m_RefEdges, m_application)
            End Select


            m_HasClicked = True





            bProjHasINode = True
            bProjHasJNode = True


        End If

        



    End Sub


    Private Sub OnCreateFeature(ByVal obj As ESRI.ArcGIS.Geodatabase.IObject)  'Do Something in responce to the event. 

        Dim pID As New UID

        Dim topoUiD As New UID

        Dim uidCmd As New UIDClass

        Dim cmd As ICommandItem

        Dim m_cmdSplitTool As New cmdSplitTool

        Dim m_modeAttributes As ModeAttributes

        Dim m_tblLineProjects As TblLineProjects

        Dim pRow As IRow

        Dim pEnvelope As IEnvelope = New Envelope

        Dim pRelationshipClass As IRelationshipClass

        Dim pFeatureWorkspace As IFeatureWorkspace

        pFeatureWorkspace = pWrkspc

        pRelationshipClass = pFeatureWorkspace.OpenRelationshipClass(sdePrefix & "TransRefEdges_tblBikeAttributes")

        Dim pDestinationTable As ITable

        pDestinationTable = pRelationshipClass.DestinationClass

        uidCmd.Value = "{2d07a9af-8f08-4c77-85e5-ea054352d454}"

        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)

        m_cmdSplitTool = CType(cmd.Command, cmdSplitTool)

        'get the Split Edge Tool so we know what kind of edits the user is going to make
        If Not m_cmdSplitTool.m_SplitEdge Is Nothing Then

            If obj.Class.AliasName = g_Schema & g_ModeAttributes Then

                pRow = CType(obj, IRow)

                m_modeAttributes = New ModeAttributes(pRow)

                GetSplitFeatureAttributes(pRow, m_modeAttributes, m_cmdSplitTool.m_EdgeModeAttributes)

                m_cmdSplitTool.m_SplitEdge = Nothing

                m_cmdSplitTool.m_edgeColl.Clear()

                Exit Sub

            End If

            If obj.Class.AliasName = g_Schema & g_ProjectAttributes Then

                pRow = CType(obj, IRow)

                m_tblLineProjects = New TblLineProjects(pRow)

                GetSplitFeatureAttributes(pRow, m_tblLineProjects, m_cmdSplitTool.m_ProjecAttributes)

                m_cmdSplitTool.m_SplitEdge = Nothing

                m_cmdSplitTool.m_edgeColl.Clear()

                Exit Sub

            End If

        End If


    End Sub
    Public Sub SetUpEvents()

        Dim uid As New UID

        uid.Value = "esriEditor.Editor"

        m_editor = CType(CType(m_application.FindExtensionByCLSID(uid), IEditor), Editor)

        m_pEditEvents = m_editor

        'set up some edit event handlers 

        dCreateFeature = New IEditEvents_OnCreateFeatureEventHandler(AddressOf OnCreateFeature)

        AddHandler m_pEditEvents.OnCreateFeature, dCreateFeature

    End Sub

    Public Sub CleanUp()

        RemoveHandler m_pEditEvents.OnCreateFeature, dCreateFeature

        RemoveHandler (CType(m_editor, IEditEvents2_Event).OnVertexAdded), dVertexAdded

        m_editor = Nothing

        m_pEditEvents = Nothing

        dCreateFeature = Nothing

        dVertexAdded = Nothing


    End Sub




    Private Sub WatchDisplay_VisibleBoundsUpdated(ByVal sender As ESRI.ArcGIS.Display.IDisplayTransformation, ByVal sizeChanged As Boolean) Handles m_pWatchDisplay.VisibleBoundsUpdated

        'test to see if map scale has changed (zoom in or out)
        'if yes, adjust snapping 
        If (m_activeView.Extent.Height * m_activeView.Extent.Width) <> (m_pEnvelope.Height * m_pEnvelope.Width) Then

            If m_editor.EditState = esriEditState.esriStateEditing Then

                SetSnappingTolerance2(m_pEnvelope, m_mxDoc, m_editor, m_pEnvelope)







                ' Dim dblSnappingTolerance As Double

                'dblSnappingTolerance = m_mxDoc.FocusMap.MapScale / 100

                'SetSnappingTolerance(m_editor, dblSnappingTolerance)
            End If

        End If

        'm_pEnvelope = m_activeView.Extent.Envelope

    End Sub

    Private Sub OnSketchFinished()

    End Sub

    Private Sub OnStopEditing(ByVal save As Boolean)

    End Sub

    Public Sub GetSplitFeatureAttributes(ByVal tableRow As IRow, ByVal toObject As Object, ByVal fromObject As Object)

        Dim _type As Type = toObject.GetType()

        Dim properties() As Reflection.PropertyInfo = _type.GetProperties()

        For Each _property As Reflection.PropertyInfo In properties

            If _property.Name <> "OID" And _property.Name <> "PSRCEDGEID" And _property.Name <> "projRteID" And _property.Name <> "CompletionDate" And _property.Name <> "EditNotes" And _property.Name <> "Processing" Then

                If tableRow.Fields.FindField(_property.Name) <> -1 Then

                    _property.SetValue(toObject, _property.GetValue(fromObject))

                End If
            End If

        Next

    End Sub

    Protected Overrides Sub Finalize()

        MyBase.Finalize()

    End Sub


End Class



