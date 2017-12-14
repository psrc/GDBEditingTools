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
Imports ESRI.ArcGIS.DataSourcesGDB
Public Class frmDeleteFeatures
    Public m_application As IApplication
    Public m_transitPoints As IFeatureLayer

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
    Public m_ModeAttributes As ITable


    Private test(150) As String
    Public x As Integer

    Public m_pRefEdgesTable As ITable
    Public m_pRefJunctionsTable As ITable
    Public m_pProjectsTable As ITable
    Public m_TLComboBoxControl As New PSRCEditingTools.TLComboBoxControl
    Private m_lngVertexCount As Long
    Private pSnap As ISnapEnvironment
    Private boolLastVertexSnapped As Boolean
    Private objCollection As New Collection

    'Private m_cmdStartEditing As New PSRCEditingTools.cmdStartEditing
    Private m_FeatAdded As IFeature
    Private refedgeFeature As IFeature
    Private pDataset As IDataset
    Private pWrkspc As IWorkspace

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        m_mxDoc = m_application.Document
        Dim uid2 As New UID
        Dim pTable As ITable
        Dim pfilter As IQueryFilter

        uid2.Value = "esriEditor.Editor"
        m_editor = CType(m_application.FindExtensionByCLSID(uid2), IEditor)
        m_activeView = m_mxDoc.ActiveView

        m_map = m_activeView.FocusMap
        g_Schema = checkWS(m_editor, m_application)

        pfilter = New QueryFilter
        pfilter.WhereClause = "POINTORDER >= 0"

        m_transitPoints = GetFeatureLayer(GlobalConstants.g_Schema & g_TransitPoints, m_application)
        pDataset = CType(m_transitPoints.FeatureClass, IDataset)

        pWrkspc = pDataset.Workspace

        pTable = m_transitPoints.FeatureClass

        'StopWorkspaceEditOperation(m_application, pWrkspc)
        StartSDEWorkspaceEditorOperation(pWrkspc, m_application)
        MessageBox.Show(pTable.RowCount(pfilter))


        pTable.DeleteSearchedRows(pfilter)
        StopWorkspaceEditOperation(m_application, pWrkspc)
        MessageBox.Show("done")


    End Sub
End Class