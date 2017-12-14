Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Controls
Imports System.Windows.Forms


Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto



<ComClass(SketchTool.ClassId, SketchTool.InterfaceId, SketchTool.EventsId), _
 ProgId("PSRCEditingTools.SketchTool")> _
Public NotInheritable Class SketchTool
    Inherits BaseTool

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "5320f0a1-6db0-4e8b-a749-f856de242211"
    Public Const InterfaceId As String = "797750db-d9f8-4cd6-b435-a786e4b576a6"
    Public Const EventsId As String = "750bed44-3391-4ca3-afa9-7d433d072c12"
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

    Private m_hookHelper As IHookHelper
    Private m_application As IApplication
    Public m_TLComboBoxControl As New PSRCEditingTools.TLComboBoxControl
    Public m_editor As Editor
    Private m_cmdStartEditing As New PSRCEditingTools.cmdStartEditing



    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "PSRCEditingTools"  'localizable text 
        MyBase.m_caption = "Sketch Tool"   'localizable text 
        MyBase.m_message = "Sketch Tool"   'localizable text 
        MyBase.m_toolTip = "Use to Digitize Features" 'localizable text 
        MyBase.m_name = "PSRCEditingTools.SketchTool"  'unique id, non-localizable (e.g. "MyCategory_MyTool")

        Try
            'TODO: change resource name if necessary
            m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("PSRCEditingTools.pencil_fat.bmp"))
            'Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            'MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
            MyBase.m_cursor = New System.Windows.Forms.Cursor(Me.GetType(), Me.GetType().Name + ".cur")
        Catch ex As Exception
            System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        End Try
    End Sub


    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            m_application = CType(hook, IApplication)

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = False
            Else
                MyBase.m_enabled = False
            End If
        End If
     




        'get the editor
        Dim uid2 As New UID
        uid2.Value = "esriEditor.Editor"
        m_editor = CType(m_application.FindExtensionByCLSID(uid2), IEditor)

        'get access to cmdStartEditing (must do this to access public variables declared in that class)
        Dim uidCmd As New UIDClass
        Dim cmd As ICommandItem

        uidCmd.Value = "{001386c4-ce20-4882-bd6c-9826b25015ff}"
        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)
        m_cmdStartEditing = cmd.Command








        'make the Tasks ComboBox active:

        ' m_TLComboBoxControl.ComboBox1.Enabled = True
        'm_TLComboBoxControl.LoadComboBox()
    End Sub

    Public Overrides Sub OnClick()

        Dim uidCmd As New UIDClass
        Dim cmd As ICommandItem
        'get access to cmdStartEditing (must do this to access public variables declared in that class)
        'uidCmd.Value = "{001386c4-ce20-4882-bd6c-9826b25015ff}"
        'cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)
        'm_cmdStartEditing = cmd.Command

        m_cmdStartEditing.m_boolSketchEdit = True




        uidCmd.Value = "{B479F48A-199D-11D1-9646-0000F8037368}"
        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)
        m_application.CurrentTool = cmd








    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add SketchTool.OnMouseDown implementation
    End Sub

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add SketchTool.OnMouseMove implementation
    End Sub

    Public Overrides Sub OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add SketchTool.OnMouseUp implementation
    End Sub
    Public Overrides ReadOnly Property Enabled() As Boolean
        'enabled if we are in editing mode
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


