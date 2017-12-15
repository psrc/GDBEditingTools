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
Imports ESRI.ArcGIS.Geodatabase
Imports System.Windows.Forms




<ComClass(cmdSplitTool.ClassId, cmdSplitTool.InterfaceId, cmdSplitTool.EventsId), _
 ProgId("PSRCEditingTools.cmdSplitTool")> _
Public NotInheritable Class cmdSplitTool
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "2d07a9af-8f08-4c77-85e5-ea054352d454"
    Public Const InterfaceId As String = "e895d234-19d1-440f-bd9a-4b370de4fcd6"
    Public Const EventsId As String = "4427e4ca-0369-4b97-abe1-f443cf69879e"
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
    Public m_editor As Editor
    Private m_cmdStartEditing As New PSRCEditingTools.cmdStartEditing
    Public m_map As IMap
    Public m_RefEdges As IFeatureLayer
    Public m_Projects As IFeatureLayer
    Public m_Junctions As IFeatureLayer
    Public m_SplitEdge As IFeature
    Public m_SplitHappened As Boolean
    Public m_edgeColl As New Collection
    Public m_EdgeModeAttributes As ModeAttributes
    Public m_ProjecAttributes As TblLineProjects



    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties

        MyBase.m_category = "PSRCEditingTools"  'localizable text 
        MyBase.m_caption = "Split Tool"   'localizable text 
        MyBase.m_message = "Split Tool"   'localizable text 
        MyBase.m_toolTip = "Use to Split Edges" 'localizable text 
        MyBase.m_name = "PSRCEditingTools.SplitTool"
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

        m_SplitHappened = False

        ' TODO:  Add other initialization code
        Dim uid2 As New UID

        Dim uidCmd As New UIDClass

        Dim cmd As ICommandItem

        uid2.Value = "esriEditor.Editor"

        m_editor = CType(m_application.FindExtensionByCLSID(uid2), IEditor)

        'get access to  cmdStartEditing 
        uidCmd.Value = "{001386c4-ce20-4882-bd6c-9826b25015ff}"

        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)

        m_cmdStartEditing = cmd.Command


    End Sub
    Public Overrides ReadOnly Property Enabled() As Boolean

        Get

            If m_cmdStartEditing.m_HasClicked = True Then

                Return True

            ElseIf m_cmdStartEditing.m_HasClicked = False Then

                Return False

            End If

        End Get
    End Property

    Public Overrides Sub OnClick()

        Dim pRow As IRow

        Dim pFeatureWorkspace As IFeatureWorkspace

        g_Schema = checkWS(m_editor, m_application)

        m_RefEdges = GetFeatureLayer(GlobalConstants.g_Schema & g_RefEdge, m_application)

        m_Junctions = GetFeatureLayer(g_Schema & g_RefJunct, m_application)

        m_Projects = GetFeatureLayer(g_Schema & g_ProjectRoutes, m_application)

        m_SplitEdge = GetEditSelection()

        pFeatureWorkspace = m_cmdStartEditing.pWrkspc

        If m_SplitEdge.Class.AliasName = g_Schema & g_RefEdge Then

            pRow = GetRelatedRecord(g_Schema & g_edgeToAttribute, pFeatureWorkspace, m_SplitEdge)

            m_EdgeModeAttributes = New ModeAttributes(pRow)

        ElseIf m_SplitEdge.Class.AliasName = g_Schema & g_ProjectRoutes Then

            pRow = GetRelatedRecord(g_Schema & g_projRouteToAttribute, pFeatureWorkspace, m_SplitEdge)

            m_ProjecAttributes = New TblLineProjects(pRow)

        End If

        Dim uidCmd As New UIDClass

        Dim cmd As ICommandItem

        uidCmd.Value = "{5609B740-112F-11D2-84A9-0000F875B9C6}"

        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)

        m_application.CurrentTool = cmd

    End Sub

    Public Function GetEditSelection() As IFeature

        Dim pEnumFeat As IEnumFeature

        Dim pFeature As IFeature

        Dim Count As Integer

        'Get the selection
        pEnumFeat = m_editor.EditSelection

        pEnumFeat.Reset()

        'Loop through the selection and perform some action
        If m_editor.SelectionCount > 1 Then

            MessageBox.Show("Only 1 Edge can be selected to use the Split Tool!", "More than 1 Edge Selected", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)


        ElseIf m_editor.SelectionCount = 0 Then

            MessageBox.Show("An edge must be selected to use the Split Tool!", "No Edge Selected", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        Else

            pFeature = pEnumFeat.Next

            GetEditSelection = pFeature


        End If

    End Function

    


End Class



