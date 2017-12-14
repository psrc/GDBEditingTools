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
<ComClass(cmdCheckIJNode2.ClassId, cmdCheckIJNode2.InterfaceId, cmdCheckIJNode2.EventsId), _
 ProgId("PSRCEditingTools.cmdCheckIJNode2")> _
Public NotInheritable Class cmdCheckIJNode2
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "a5b884f1-91b2-47c2-bfa0-f1a63754a901"
    Public Const InterfaceId As String = "d09e2447-7e2a-4cc1-ad04-075fce2c03da"
    Public Const EventsId As String = "ca0efe0b-96c0-4db4-961c-14838f355c9f"
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
    Public m_cmdStartEditing As New PSRCEditingTools.cmdStartEditing
    Private m_lngVertexCount As Long
    Private pSnap As ISnapEnvironment
    Private boolLastVertexSnapped As Boolean
    Private objCollection As New Collection

    'Private m_cmdStartEditing As New PSRCEditingTools.cmdStartEditing
    Private m_FeatAdded As IFeature
    Private refedgeFeature As IFeature
    Private pDataset As IDataset
    Private pWrkspc As IWorkspace
    Public WithEvents mEditorOn As TestEvents



    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "PSRCEditingTools"   'localizable text 
        MyBase.m_caption = "Adds IJ Attributes to Selected Projects and Edges"   'localizable text 
        MyBase.m_message = "Adds IJ Attributes to Selected Projects and Edges"    'localizable text 
        MyBase.m_toolTip = "Click to Add IJ Attributes to Selected Projects and Edges features" 'localizable text 
        MyBase.m_name = "PSRCEditingTools.cmdCheckIJNode2"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        mEditorOn = New TestEvents
        'TestAddHandler()

        ' TODO:  Add other initialization code
        ' TODO:  Add other initialization code
    End Sub

    Public Overrides Sub OnClick()
        mEditorOn = New TestEvents
        'TODO: Add cmdCheckIJNode2.OnClick implementation
        m_mxDoc = m_application.Document

        m_activeView = m_mxDoc.ActiveView

        m_map = m_activeView.FocusMap
        g_Schema = checkWS(m_editor, m_application)

        m_RefEdges = GetFeatureLayer(GlobalConstants.g_Schema & g_RefEdge, m_application)
        m_Junctions = GetFeatureLayer(g_Schema & g_RefJunct, m_application)
        m_Projects = GetFeatureLayer(g_Schema & g_ProjectRoutes, m_application)



        pDataset = CType(m_RefEdges.FeatureClass, IDataset)

        pWrkspc = pDataset.Workspace
        CheckIJNode3(m_map)
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
    Public Sub CheckIJNode3(ByVal pMap As IMap)
        Dim pEnumFeat As IEnumFeature
        Dim pSelection As ISelection
        Dim pFeature As IFeature
        Dim pINode As Long
        Dim pJNode As Long
        pSelection = pMap.FeatureSelection
        pEnumFeat = pSelection
        pFeature = pEnumFeat.Next
        If pFeature Is Nothing Then
            MessageBox.Show("No Features are selected. This tool only works on selected TransRefEdge and Project Features!", "No Features Selected!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else

            Do Until pFeature Is Nothing
                If pFeature.Class.AliasName.Contains("ProjectRoutes") Then
                    Dim pProject As New ProjectAttributes(pFeature)

                    'If pProject.HasINode = False Then

                    pINode = GetFeatureINode(pFeature, m_Junctions)
                    If pINode = 0 Then
                        MessageBox.Show("An INode Junction could not be found for Project ID: " & pProject.PROJRTEID)
                    Else
                        StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
                        pProject.ProjectINode = pINode
                        StopWorkspaceEditOperation(m_application, pWrkspc)
                    End If

                    'If

                    'If pProject.HasJNode = False Then
                    pJNode = GetFeatureJNode(pFeature, m_Junctions)
                    If pJNode = 0 Then
                        MessageBox.Show("An JNode Junction could not be found for Project ID: " & pProject.PROJRTEID)
                    Else
                        StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
                        pProject.ProjectJNode = pJNode
                        StopWorkspaceEditOperation(m_application, pWrkspc)
                    End If

                    'End If

                ElseIf pFeature.Class.AliasName.Contains("TransRefEdges") Then
                    Dim pEdge As New RefEdgeAttributes(pFeature)

                    'If pEdge.HasINode = False Then
                    pINode = GetFeatureINode(pFeature, m_Junctions)
                    If pINode = 0 Then
                        MessageBox.Show("An INode Junction could not be found for Edge ID: " & pEdge.EdgeID)
                    Else
                        StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
                        pEdge.INode = pINode
                        StopWorkspaceEditOperation(m_application, pWrkspc)
                    End If
                    'End If



                    'If pEdge.HasJNode = False Then
                    pJNode = GetFeatureJNode(pFeature, m_Junctions)
                    If pJNode = 0 Then
                        MessageBox.Show("An JNode Junction could not be found for Edge ID: " & pEdge.EdgeID)
                    Else
                        StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
                        pEdge.JNode = pJNode
                        StopWorkspaceEditOperation(m_application, pWrkspc)
                    End If

                    'End If



                End If




                pFeature = pEnumFeat.Next
            Loop
            MessageBox.Show("IJ Nodes Checked for Selected TransRefEdges and Project Features", "IJ Nodes Checked", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub



    'Public Sub EHandler(ByRef IsOn As Boolean)
    'MessageBox.Show(IsOn.ToString)

    'End Sub
    ' Public Sub TestAddHandler()
    'Dim CI As New cmdStartEditing
    ' AddHandler CI.EditorOn, AddressOf EHandler

    ' End Sub



    Private Sub mEditorOn_EditorOn(ByVal IsOn As Boolean) Handles mEditorOn.EditorOn
        If IsOn = True Then
            MessageBox.Show("is on")

        End If
    End Sub
End Class



