Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports System.Runtime.InteropServices

<ComClass(PSRCEditingToolBar.ClassId, PSRCEditingToolBar.InterfaceId, PSRCEditingToolBar.EventsId), _
 ProgId("PSRCEditingTools.PSRCEditingToolBar")> _
Public NotInheritable Class PSRCEditingToolBar
    Inherits BaseToolbar

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
        MxCommandBars.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommandBars.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "09195090-3a7b-4176-bed3-c7c328c5cbd0"
    Public Const InterfaceId As String = "ff4475b7-f9bf-4998-9ada-b57c66316cbe"
    Public Const EventsId As String = "9468b71a-0ce9-4c3e-b282-deda7776e6f3"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()

        '
        'TODO: Define your toolbar here by adding items
        '

        'start editing drop down
        AddItem("{0550a6e2-c8d7-470d-8458-a3a23c88fe16}")
        BeginGroup()
        'Bike Attribute Updateer
        AddItem("{c292936f-0644-40d9-a9e5-8a9c48f8ad25}")
        AddItem("{e424b89b-78da-4f62-ab2b-d97e51feb623}")
        AddItem("{2d07a9af-8f08-4c77-85e5-ea054352d454}")
        AddItem("{a5b884f1-91b2-47c2-bfa0-f1a63754a901}")

        BeginGroup()
        AddItem("{5f246a88-4eec-4d51-95cf-7e7e02d1a072}")

        'edit task
        'AddItem("{20A94A18-61FC-4516-A924-3C7AD4C9EC8E}")
       
        'sketch tool
        'AddItem("{B479F48A-199D-11D1-9646-0000F8037368}")
       

        'Split tool
        ' AddItem("{27cbe1f8-0c46-4cf3-8b32-3ff892767ed8}")

        'topology tools
        'BeginGroup()
        ' AddItem("{bdef4d84-8b18-4455-ab1c-e8b118ae4151}")

        'AddItem("{0fd2b09a-122d-44aa-9bf8-25af626bddf2}")


        'AddItem("{a5b884f1-91b2-47c2-bfa0-f1a63754a901}")
        'AddItem("{5f246a88-4eec-4d51-95cf-7e7e02d1a072}")






        AddItem("{15280b31-8441-4eb7-a641-c758471b5465}")

    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            'TODO: Replace bar caption
            Return "PSRC Editing Tool Bar"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            'TODO: Replace bar ID
            Return "PSRCEditingToolBar"
        End Get
    End Property
End Class
