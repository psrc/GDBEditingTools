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

<ComClass(TraceTool.ClassId, TraceTool.InterfaceId, TraceTool.EventsId), _
 ProgId("PSRCEditingTools.TraceTool")> _
Public NotInheritable Class TraceTool
    Inherits BaseTool

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "27cbe1f8-0c46-4cf3-8b32-3ff892767ed8"
    Public Const InterfaceId As String = "03f40ef7-b575-416f-92f2-aa8df1626599"
    Public Const EventsId As String = "70f60ca9-4bce-474c-a19e-9badde26ba99"
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
        MyBase.m_category = "PSRCEditingTools" 'localizable text 
        MyBase.m_caption = "Trace Projects"   'localizable text 
        MyBase.m_message = "Trace Projects"   'localizable text 
        MyBase.m_toolTip = "Use to Trace Projects" 'localizable text 
        MyBase.m_name = "PSRCEditingTools.TraceTool"  'unique id, non-localizable (e.g. "MyCategory_MyTool")

        Try
            'TODO: change resource name if necessary
            'Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
            'MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
            m_bitmap = New Bitmap(Me.GetType.Assembly.GetManifestResourceStream("PSRCEditingTools.TraceTool.bmp"))
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
        Dim uid As New UID
        Dim cmd As ICommandItem
        uid.Value = "{20A94A18-61FC-4516-A924-3C7AD4C9EC8E}"
        cmd = m_application.Document.CommandBars.Find(uid, False, False)
        m_TLComboBoxControl = cmd.Command

        Dim uid2 As New UID

        uid2.Value = "esriEditor.Editor"
        m_editor = CType(m_application.FindExtensionByCLSID(uid2), IEditor)
        'TODO: Add other initialization code

    End Sub

    Public Overrides Sub OnClick()
        Dim uidCmd As New UIDClass
        Dim cmd As ICommandItem

        'get access to cmdStartEditing (must do this to access public variables declared in that class)
        uidCmd.Value = "{001386c4-ce20-4882-bd6c-9826b25015ff}"
        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)
        m_cmdStartEditing = cmd.Command

        m_cmdStartEditing.m_boolSketchEdit = True

        'sketch tool 
        uidCmd.Value = "{E3B16B57-4D2E-4F49-A306-9DF7817F6D60}"
        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)
        m_application.CurrentTool = cmd

        'TODO: Add TraceTool.OnClick implementation

    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        If Button = 2 Then

        End If
        'TODO: Add TraceTool.OnMouseDown implementation
    End Sub

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add TraceTool.OnMouseMove implementation
    End Sub

    Public Overrides Sub OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add TraceTool.OnMouseUp implementation
    End Sub
    Public Overrides ReadOnly Property Enabled() As Boolean
        Get
            'If m_TLComboBoxControl.ComboBox1.SelectedItem = "Add Projects" Then
            If m_editor.EditState = esriEditState.esriStateEditing And m_TLComboBoxControl.ComboBox1.SelectedItem = "Add Projects" Then
                Return True
                ' End If


            Else
                Return False



            End If
        End Get
    End Property
End Class


