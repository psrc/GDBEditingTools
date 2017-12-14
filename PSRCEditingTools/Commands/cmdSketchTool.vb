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


<ComClass(cmdSketchTool.ClassId, cmdSketchTool.InterfaceId, cmdSketchTool.EventsId), _
 ProgId("PSRCEditingTools.cmdSketchTool")> _
Public NotInheritable Class cmdSketchTool
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "8109ef4a-f17c-4a3c-ba77-9c610e3f7757"
    Public Const InterfaceId As String = "0b64d277-e2f7-444f-bd9b-4b20a7db5249"
    Public Const EventsId As String = "45902b28-9782-415a-be78-e6d29def967d"
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

    'Public m_TLComboBoxControl As New PSRCEditingTools.TLComboBoxControl
    Public m_editor As Editor
    Private m_cmdStartEditing As New PSRCEditingTools.cmdStartEditing



    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = ""  'localizable text 
        MyBase.m_caption = ""   'localizable text 
        MyBase.m_message = ""   'localizable text 
        MyBase.m_toolTip = "" 'localizable text 
        MyBase.m_name = ""  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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


        'Get the Target Layer Combo Box (TLComboBoxControl)
        'Dim uid As New UID
        'Dim cmd As ICommandItem
        'uid.Value = "{20A94A18-61FC-4516-A924-3C7AD4C9EC8E}"
        'cmd = m_application.Document.CommandBars.Find(uid, False, False)
        'm_TLComboBoxControl = cmd.Command

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



    End Sub

    Public Overrides Sub OnClick()
        Dim uidCmd As New UIDClass
        Dim cmd As ICommandItem

       


        'Get the sketch tool and set it to ArcMap current tool (editing)
      
        uidCmd.Value = "{B479F48A-199D-11D1-9646-0000F8037368}"
        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)
        m_application.CurrentTool = cmd

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



