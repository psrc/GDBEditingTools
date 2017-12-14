Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Controls
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Framework

<ComClass(cmdStopEditing.ClassId, cmdStopEditing.InterfaceId, cmdStopEditing.EventsId), _
 ProgId("PSRCEditingTools.cmdStopEditing")> _
Public NotInheritable Class cmdStopEditing
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "1473738c-4ddb-4a3c-b6ef-d453491abc1d"
    Public Const InterfaceId As String = "bb51446e-03fa-4aaf-bd3a-0e833b9ce239"
    Public Const EventsId As String = "d7060523-42c3-46ef-ac9a-e0d0edef87f6"
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


    Private m_application As ESRI.ArcGIS.Framework.IApplication
    Private m_editor As Editor
    Private m_TLComboBoxControl As PSRCEditingTools.TLComboBoxControl
    Private m_mxDoc As IMxDocument
    Private m_cmdStartEditing As New PSRCEditingTools.cmdStartEditing


    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "PSRCEditingTools"  'localizable text 
        MyBase.m_caption = "Stop Editing"   'localizable text 
        MyBase.m_message = "Stop Editing"   'localizable text 
        MyBase.m_toolTip = "Click to Stop Editor" 'localizable text 
        MyBase.m_name = "PSRCEditingTools.cmdStopEditing"  'unique id, non-localizable (e.g. "MyCategory_MyCommand")

        'Try
        'TODO: change bitmap name if necessary
        'Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
        'MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        'Catch ex As Exception
        'System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        'End Try


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
        Dim uid As New UID

        uid.Value = "esriEditor.Editor"
        m_editor = CType(m_application.FindExtensionByCLSID(uid), IEditor)
        m_mxDoc = m_application.Document

        'get the ComboBox
        Dim uidCmd As New UIDClass
        Dim cmd As ICommandItem

        'get access to cmdStartEditing (must do this to access public variables declared in that class)
        uidCmd.Value = "{001386c4-ce20-4882-bd6c-9826b25015ff}"
        cmd = m_application.Document.CommandBars.Find(uidCmd, False, False)
        m_cmdStartEditing = cmd.Command



        'm_pPolyLine = m_pMap.Layer(3)
    End Sub

    Public Overrides Sub OnClick()

        If MessageBox.Show("Do you want to save your edits?", "Save Edits?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            StopEditor(True, m_application)
        Else
            StopEditor(False, m_application)

        End If
        m_mxDoc.ActiveView.Refresh()
        m_cmdStartEditing.m_HasClicked = False

        m_cmdStartEditing.CleanUp()

        'm_TLComboBoxControl.ComboBox1.Enabled = False

        'TODO: Add cmdStopEditing.OnClick implementation
    End Sub
    Public Overrides ReadOnly Property Checked() As Boolean

        'Enabled if this command (start psrc editing) has not been clicked yet or after 'stop editing' button is clicked 
        Get

            If m_editor.EditState = esriEditState.esriStateEditing Then


                Return False

            ElseIf m_editor.EditState = esriEditState.esriStateNotEditing Then


                Return True




            End If




        End Get
    End Property

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class



