Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.esriSystem





<ComClass(cmdAttributeUpdater.ClassId, cmdAttributeUpdater.InterfaceId, cmdAttributeUpdater.EventsId), _
 ProgId("PSRCEditingTools.cmdAttributeUpdater")> _
Public NotInheritable Class cmdAttributeUpdater
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "e424b89b-78da-4f62-ab2b-d97e51feb623"
    Public Const InterfaceId As String = "a537f938-1973-4989-957f-cd71eafccf93"
    Public Const EventsId As String = "e80d5db0-29f7-4126-ae7e-a46f203b67ba"
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

    Public m_frmAttributeUpdater As New frmAttributeUpdater

    Public m_editor As Editor
    Private m_cmdStartEditing As New PSRCEditingTools.cmdStartEditing

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "PSRCEditingTools" 'localizable text 
        MyBase.m_caption = "Edit Multiple ModeAttribute Records"   'localizable text 
        MyBase.m_message = "Edit Multiple ModeAttribute Records"   'localizable text 
        MyBase.m_toolTip = "Click to Edit Multiple ModeAttribute Records" 'localizable text 
        MyBase.m_name = "PSRCEditingTools.cmdAttributeUpdater"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        m_frmAttributeUpdater.m_application = m_application
        Dim uid2 As New UID
        uid2.Value = "esriEditor.Editor"
        m_editor = CType(m_application.FindExtensionByCLSID(uid2), IEditor)

        'get access to  cmdStartEditing 
        Dim uidCmd As New UIDClass
        Dim cmd As ICommandItem

        uidCmd.Value = "{001386c4-ce20-4882-bd6c-9826b25015ff}"
        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)
        m_cmdStartEditing = cmd.Command




        ' TODO:  Add other initialization code
    End Sub

    Public Overrides Sub OnClick()
        Dim tableSelect As ITableSelection
        Dim SelectionSet As ISelectionSet
        Dim m_ModeAttributes As ITable
        Dim m_RefEdges As IFeatureLayer
        Dim EnumRelationshipClass As IEnumRelationshipClass
        Dim RelationshipClass As IRelationshipClass



        m_RefEdges = GetFeatureLayer(GlobalConstants.g_Schema & g_RefEdge, m_application)
        m_ModeAttributes = getStandaloneTable(g_Schema & g_ModeAttributes, m_application)
        EnumRelationshipClass = m_RefEdges.FeatureClass.RelationshipClasses(esriRelRole.esriRelRoleAny)
        'RelationshipClass = EnumRelationshipClass.Next
        'Do Until RelationshipClass Is Nothing
        'If RelationshipClass.DestinationClass.AliasName = "sde.SDE.modeAttributes" Then
        'Exit Do
        ' Else
        'RelationshipClass = EnumRelationshipClass.Next
        'End If
        'Loop
        'relationshipclass.


        tableSelect = m_ModeAttributes
        SelectionSet = tableSelect.SelectionSet




        If SelectionSet.Count = 0 Then
            MessageBox.Show("No Records in ModeAttributes Selected", "No Records Selected", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If



        m_frmAttributeUpdater.ShowDialog()


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
End Class



