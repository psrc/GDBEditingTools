Imports ESRI.ArcGIS.Editor
Imports ESRI.ArcGIS.Geodatabase
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Controls
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports Microsoft.SqlServer

Module GeoDatabaseManagement
    Public Function GetFeatureLayer(ByVal sLayerName As String, ByVal app As IApplication, Optional ByVal bLayerName As Boolean = False) As IFeatureLayer2

        Dim pDoc As IMxDocument
        Dim pMap As IMap
        Dim pEnumLy As IEnumLayer


        pDoc = app.Document
        pMap = pDoc.FocusMap
        pEnumLy = pMap.Layers

        Dim pLy As ILayer, pFLy As IFeatureLayer
        pLy = pEnumLy.Next
        Do Until pLy Is Nothing

            If pLy.Valid Then
                If TypeOf pLy Is IFeatureLayer Then
                    pFLy = pLy
                    If bLayerName Then
                        If UCase(pLy.Name) = UCase(sLayerName) Then
                            GetFeatureLayer = pFLy
                            'GoTo ReleaseObjs
                        End If
                    Else
                        If Not TypeOf pFLy.FeatureClass Is IRelQueryTable Then
                            If UCase(pFLy.FeatureClass.AliasName) = UCase(sLayerName) Then
                                GetFeatureLayer = pFLy
                                'GoTo ReleaseObjs
                            End If
                        End If
                    End If
                End If
            End If
            pLy = pEnumLy.Next
        Loop

        'ReleaseObjs:
        pDoc = Nothing
        pMap = Nothing
        pEnumLy = Nothing
        pLy = Nothing
        'pFLy = Nothing
    End Function

    Public Function getStandaloneTable(ByVal sTableName As String, ByVal app As IApplication) As ITable
        Dim pDoc As IMxDocument
        Dim pMap As IMap



        pDoc = app.Document
        pMap = pDoc.FocusMap







        Dim pStTblCol As IStandaloneTableCollection
        pStTblCol = pMap

        Dim i As Integer
        For i = 0 To pStTblCol.StandaloneTableCount - 1
            If pStTblCol.StandaloneTable(i).Name = sTableName Then
                If pStTblCol.StandaloneTable(i).Valid = True Then
                    getStandaloneTable = pStTblCol.StandaloneTable(i)
                End If
            End If
        Next i

        pStTblCol = Nothing

    End Function
    Public Function checkWS(ByVal pEditor As IEditor, ByVal app As IApplication, Optional ByVal bTransit As Boolean = False) As String
        Dim pWS As IWorkspace
        Dim sSchema As String
        pWS = pEditor.EditWorkspace
        'MsgBox pWS.ConnectionProperties.GetProperty("Database")
        Dim i As Integer
        Dim pFLayer As IFeatureLayer
        If pWS.WorkspaceFactory.WorkspaceType = esriWorkspaceType.esriRemoteDatabaseWorkspace Then
            Dim pELy As IEditLayers, pLy As ILayer2
            Dim pos As Integer
            pELy = pEditor
            If pELy.CurrentLayer Is Nothing Then
                For i = 0 To pEditor.Map.LayerCount - 1
                    pFLayer = pEditor.Map.Layer(i)
                    If pFLayer.DataSourceType = "SDE Feature Class" Then
                        pLy = pFLayer
                        Exit For

                    End If
                Next
            Else
                pLy = pELy.CurrentLayer
            End If
            pos = InStrRev(pLy.Name, ".")
            If pos > 0 Then
                'the schema name is located

                sSchema = Left(pLy.Name, pos)
            Else
                sSchema = ""
            End If
        Else
            sSchema = ""
        End If

            checkWS = sSchema
            Dim sMissingLayer As String
            If GetFeatureLayer(sSchema & g_RefEdge, app) Is Nothing Then sMissingLayer = sMissingLayer & sSchema & g_RefEdge & vbNewLine
            If GetFeatureLayer(sSchema & g_RefJunct, app) Is Nothing Then sMissingLayer = sMissingLayer & sSchema & g_RefJunct & vbNewLine
            If GetFeatureLayer(sSchema & g_TransitLines, app) Is Nothing Then sMissingLayer = sMissingLayer & sSchema & g_TransitLines & vbNewLine
            If GetFeatureLayer(sSchema & g_TransitPoints, app) Is Nothing Then sMissingLayer = sMissingLayer & sSchema & g_TransitPoints & vbNewLine
            If getStandaloneTable(sSchema & g_ModeAttributes, app) Is Nothing Then sMissingLayer = sMissingLayer & sSchema & g_ModeAttributes & vbNewLine
            If Not bTransit Then
                If GetFeatureLayer(sSchema & g_TurnMovements, app) Is Nothing Then sMissingLayer = sMissingLayer & sSchema & g_TurnMovements & vbNewLine
                If GetFeatureLayer(sSchema & g_NamedRoutes, app) Is Nothing Then sMissingLayer = sMissingLayer & sSchema & g_NamedRoutes & vbNewLine
                If GetFeatureLayer(sSchema & g_ProjectRoutes, app) Is Nothing Then sMissingLayer = sMissingLayer & sSchema & g_ProjectRoutes & vbNewLine
            End If
            '    If Not getTable(sSchema & g_ModeAttributes) Is Nothing Then sMissingLayer = sMissingLayer & sSchema & g_ModeAttributes & vbNewLine
            '    If Not getTable(sSchema & g_ModeTolls) Is Nothing Then sMissingLayer = sMissingLayer & sSchema & g_ModeAttributes & vbNewLine
            '    If Not getTable(sSchema & g_EdgeFac) Is Nothing Then sMissingLayer = sMissingLayer & sSchema & g_ModeAttributes & vbNewLine
            '    If Not getTable(sSchema & g_) Is Nothing Then sMissingLayer = sMissingLayer & sSchema & g_ModeAttributes & vbNewLine

            If sMissingLayer <> "" Then
                MsgBox("The following layers are missing, please load them first. Make sure they are all from the same Version" & vbNewLine & sMissingLayer, vbInformation)
                checkWS = "-1"
            End If
    End Function

End Module
