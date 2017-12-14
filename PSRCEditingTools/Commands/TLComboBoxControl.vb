' Copyright 2006 ESRI
'
' All rights reserved under the copyright laws of the United States
' and applicable international laws, treaties, and conventions.
'
' You may freely redistribute and use this sample code, with or
' without modification, provided you include the original copyright
' notice and use restrictions.
'
' See use restrictions at /arcgis/developerkit/userestrictions.

Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.SystemUI
Imports System.Drawing.Text
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Editor

Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses

Imports ESRI.ArcGIS.Controls


Imports ESRI.ArcGIS.esriSystem

Imports System.Windows.Forms



<ComClass(TLComboBoxControl.ClassId, TLComboBoxControl.InterfaceId, TLComboBoxControl.EventsId), _
 ProgId("PSRCEditingTools.TLComboBoxControl")> _
Public Class TLComboBoxControl
    Implements ICommand
    Implements IToolControl
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
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "20A94A18-61FC-4516-A924-3C7AD4C9EC8E"
    Public Const InterfaceId As String = "1C3AB91A-3B3C-4fae-95E0-42C23F850BA9"
    Public Const EventsId As String = "7C107C99-AACA-469f-8C19-C4505FD8B8B4"
#End Region

    Private m_application As IApplication
    Private m_mxDoc As IMxDocument
    Private m_mxDocument As MxDocument
    Private m_map As Map
    Public m_pFLyLine As IFeatureLayer
    Public m_pFLyPoint As IFeatureLayer2
    Private m_pFClsLine As IFeatureClass
    Private m_PFCLsPoint As IFeatureClass
    Private m_dctLines As Collection


    'Private m_ifc As InstalledFontCollection
    Private m_hBitmap As IntPtr
    Private m_completeNotify As ICompletionNotify
    Public m_ComboBox1Selection As String
    Public m_enabled As Boolean
    Private m_editor As Editor
    'Public m_cmdStartEditing As New PSRCEditingTools.cmdStartEditing

    Private m_activeView As IActiveView

    Public m_RefEdges As IFeatureLayer
    Public m_Projects As IFeatureLayer
    Public m_Junctions As IFeatureLayer

    'Private m_cmdStartEditing As New PSRCEditingTools.cmdStartEditing





    <DllImport("gdi32.dll")> _
    Private Shared Function DeleteObject(ByVal hObject As IntPtr) As Boolean
    End Function

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        'Set up bitmap - Note clean up is done in Dispose method instead of Finalize
        'm_hBitmap = My.Resources.FontIcon.GetHbitmap(Drawing.Color.Magenta)
    End Sub





#Region "ICommand Members"
    Public ReadOnly Property Bitmap() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.Bitmap
        Get
            Return m_hBitmap.ToInt32()
        End Get
    End Property

    Public ReadOnly Property Caption() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Caption
        Get
            Return ""
        End Get
    End Property

    Public ReadOnly Property Category() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Category
        Get
            Return ".NET Samples"
        End Get
    End Property

    Public ReadOnly Property Checked() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Checked
        Get
            Return True
        End Get
    End Property

    Public ReadOnly Property Enabled1() As Boolean Implements ESRI.ArcGIS.SystemUI.ICommand.Enabled
        Get

            If m_editor.EditState = esriEditState.esriStateEditing Then
                Me.ComboBox1.BackColor = Color.White
                ' Me.ComboBox1.ForeColor = DefaultForeC

               

                Return True

            ElseIf m_editor.EditState = esriEditState.esriStateNotEditing Then

                Me.ComboBox1.BackColor = Color.LightGray
                Return False
                Me.Refresh()



            End If



        End Get
    End Property

    Public ReadOnly Property HelpContextID() As Integer Implements ESRI.ArcGIS.SystemUI.ICommand.HelpContextID
        Get
            Return 0
        End Get
    End Property

    Public ReadOnly Property HelpFile() As String Implements ESRI.ArcGIS.SystemUI.ICommand.HelpFile
        Get
            Return String.Empty
        End Get
    End Property

    Public ReadOnly Property Message() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Message
        Get
            Return ""
        End Get
    End Property

    Public ReadOnly Property Name1() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Name
        Get
            Return ""
        End Get
    End Property

    Public Sub OnClick1() Implements ESRI.ArcGIS.SystemUI.ICommand.OnClick

    End Sub

    Public Sub OnCreate(ByVal hook As Object) Implements ESRI.ArcGIS.SystemUI.ICommand.OnCreate
        'Set up data for the dropdown
        'm_ifc = New InstalledFontCollection()
        'cboFont.DataSource = m_ifc.Families
        'cboFont.ValueMember = "Name"

        m_application = TryCast(hook, IApplication)
        Me.Enabled = True
        Me.Visible = True
        LoadComboBox()
        Dim uid As New UID

        uid.Value = "esriEditor.Editor"
        m_editor = CType(m_application.FindExtensionByCLSID(uid), IEditor)
        Me.ComboBox1.BackColor = Color.LightGray

        'gain access to the StartEditing command
        'Dim uid2 As New UID
        'Dim cmd As ICommandItem
        'uid2.Value = "{001386c4-ce20-4882-bd6c-9826b25015ff}"
        'cmd = m_application.Document.CommandBars.Find(uid2, False, False)
        'm_cmdStartEditing = cmd.Command


        'm_enabled = False





        'TODO: Uncomment the following lines if you want the control to sync with default document font
        'OnDocumentSession()
        'SetUpDocumentEvent(m_application.Document)
        'AddHandler cboFont.SelectedValueChanged, AddressOf cboFont_SelectedValueChanged
    End Sub

    Public ReadOnly Property Tooltip() As String Implements ESRI.ArcGIS.SystemUI.ICommand.Tooltip
        Get
            Return ""
        End Get
    End Property
#End Region
#Region "IToolControl Members"
    Public ReadOnly Property hWnd() As Integer Implements ESRI.ArcGIS.SystemUI.IToolControl.hWnd
        Get
            Return Me.Handle.ToInt32()
        End Get
    End Property

    Public Function OnDrop(ByVal barType As ESRI.ArcGIS.SystemUI.esriCmdBarType) As Boolean Implements ESRI.ArcGIS.SystemUI.IToolControl.OnDrop
        ' OnDocumentSession() 'Initialize the font


        m_mxDoc = m_application.Document

        m_activeView = m_mxDoc.ActiveView

        m_map = m_activeView.FocusMap
        m_RefEdges = m_map.Layer(0)
        m_Junctions = m_map.Layer(1)
        m_Projects = m_map.Layer(2)
        Return True
    End Function

    Public Sub OnFocus(ByVal complete As ESRI.ArcGIS.SystemUI.ICompletionNotify) Implements ESRI.ArcGIS.SystemUI.IToolControl.OnFocus

        'LoadComboBox()


        m_completeNotify = complete


    End Sub

#End Region






    Public Sub LoadComboBox()

        Dim dct As Collection
        dct = New Collection
        dct.Add("Add Projects")
        'dct.Add("Edit Projects")
        dct.Add("Add TransRefEdges")
        dct.Add("Add Junctions")

        ComboBox1.DataSource = dct





    End Sub

    Private Sub ComboBox1_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.DropDown

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        m_ComboBox1Selection = ComboBox1.SelectedValue.ToString

    End Sub


    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        'prevent the user from typing in the combobox

        e.Handled = True

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        m_mxDoc = m_application.Document

        m_activeView = m_mxDoc.ActiveView

        m_map = m_activeView.FocusMap
       

        If Not m_map Is Nothing And Not m_editor Is Nothing Then
            m_map = m_activeView.FocusMap
            g_Schema = checkWS(m_editor, m_application)

            m_RefEdges = GetFeatureLayer(GlobalConstants.g_Schema & g_RefEdge, m_application)
            m_Junctions = GetFeatureLayer(g_Schema & g_RefJunct, m_application)
            m_Projects = GetFeatureLayer(g_Schema & g_ProjectRoutes, m_application)

           


            If Me.ComboBox1.SelectedItem = "Add Projects" Then

                'ClearSnapAgents(m_editor)
                'SetSnappingEnv(m_editor, m_RefEdges.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
                'SetSnappingEnv(m_editor, m_Junctions.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
                SetEditlayer(m_Projects, m_application)

            ElseIf Me.ComboBox1.SelectedItem = "Add Junctions" Then
                'ClearSnapAgents(m_editor)
                'SetSnappingEnv(m_editor, m_RefEdges.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
                'SetSnappingEnv(m_editor, m_Projects.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)

                SetEditlayer(m_Junctions, m_application)
            ElseIf Me.ComboBox1.SelectedItem = "Add TransRefEdges" Then
                'ClearSnapAgents(m_editor)
                'SetSnappingEnv(m_editor, m_Projects.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
                'SetSnappingEnv(m_editor, m_RefEdges.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
                'SetSnappingEnv(m_editor, m_Junctions.FeatureClass, ESRI.ArcGIS.Geometry.esriGeometryHitPartType.esriGeometryPartBoundary)
                SetEditlayer(m_RefEdges, m_application)
            End If

            'ElseIf Me.ComboBox1.SelectedItem = "Edit Junctions" Then
            'StartEditor(m_cmdStartEditing.m_Junctions, m_application)

        End If
    End Sub

    Private Sub TLComboBoxControl_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        ComboBox1.Refresh()
        Label1.Refresh()

    End Sub

    Private Sub ComboBox1_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.DropDownClosed
        If m_completeNotify IsNot Nothing Then
            m_completeNotify.SetComplete()
        End If
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
End Class
