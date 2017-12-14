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
Imports ESRI.ArcGIS.GeoDatabaseUI

<ComClass(cmdPopulateBikeAttributes.ClassId, cmdPopulateBikeAttributes.InterfaceId, cmdPopulateBikeAttributes.EventsId), _
 ProgId("PSRCEditingTools.cmdPopulateBikeAttributes")> _
Public NotInheritable Class cmdPopulateBikeAttributes
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "c292936f-0644-40d9-a9e5-8a9c48f8ad25"
    Public Const InterfaceId As String = "a3e05be1-8537-4ff2-a210-823b45f1328b"
    Public Const EventsId As String = "da6302d2-b3f3-4d70-bfee-e307d97e6aa3"
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
    Public m_BikeAttributes As ITable


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
        MyBase.m_caption = "Populates EdgeID in Bike Attribute Table"   'localizable text 
        MyBase.m_message = "Populates EdgeID in Bike Attribute Table"    'localizable text 
        MyBase.m_toolTip = "Click to Add EdgeID to Bike Attribute Table from Selected Edges features" 'localizable text 
        MyBase.m_name = "PSRCEditingTools.cmdPopulateBikeAttributes"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        m_BikeAttributes = getStandaloneTable(g_Schema & g_BikeAttributes, m_application)


        pDataset = CType(m_RefEdges.FeatureClass, IDataset)

        pWrkspc = pDataset.Workspace
        PopulateBikeAttributes(m_map, pWrkspc)
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
    Public Sub PopulateBikeAttributes(ByVal pMap As IMap, ByVal pWorkspace As IWorkspace)















        Dim pEnumFeat As IEnumFeature
        Dim pSelection As ISelection
        Dim pFeatureSelection As IFeatureSelection
        Dim pSelSet As ISelectionSet
        Dim pFtrCsr As IFeatureCursor
        Dim pRelationshipClass As IRelationshipClass
        Dim pFeatureWorkspace As IFeatureWorkspace
        pFeatureWorkspace = pWorkspace
        pRelationshipClass = pFeatureWorkspace.OpenRelationshipClass("sde.SDE.EdgeToBike")
        Dim pDestinationTable As ITable
        pDestinationTable = pRelationshipClass.DestinationClass
        Dim pDestinationRow As IRow





        Dim pFeature As IFeature

        pFeatureSelection = m_RefEdges
        pFeatureSelection.SelectionSet.Search(Nothing, False, pFtrCsr)

        pFeature = pFtrCsr.NextFeature
        Dim pEdgeAttributes As RefEdgeAttributes

        If pFeature Is Nothing Then
            MessageBox.Show("No Features are selected. This tool only works on selected TransRefEdge and Project Features!", "No Features Selected!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else

            While Not pFeature Is Nothing
                pEdgeAttributes = New RefEdgeAttributes(pFeature)
                If pRelationshipClass.GetObjectsRelatedToObject(pFeature).Count > 0 Then
                    pFeature = pFtrCsr.NextFeature
                Else
                    Dim rowBuffer As IRowBuffer = pDestinationTable.CreateRowBuffer
                    Dim insertCursor As ICursor = pDestinationTable.Insert(True)
                    Try
                        rowBuffer.Value(1) = pEdgeAttributes.EdgeID
                        insertCursor.InsertRow(rowBuffer)
                        insertCursor.Flush()
                        Marshal.ReleaseComObject(insertCursor)
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try

                    pFeature = pFtrCsr.NextFeature
                End If



            End While
        End If
        'Dim pTablewindow As ITableWindow
        ' pTablewindow = New TableWindow
        ' pTablewindow.Table = pDestinationTable
        'pTablewindow = pDestinationTable
        'pTablewindow.Refresh()
        'pTablewindow.Show(True)

        Dim pMxDocument As IMxDocument
        Dim pActiveView As IActiveView
        pMxDocument = m_application.Document
        pActiveView = pMxDocument.FocusMap

        pActiveView.Refresh()




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



