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

Public Class frmAttributeUpdater
    Public m_application As IApplication
    'Public m_application As IApplication

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
    Private myArrayList(50) As ArrayList

    Private CodedValueComboBox(50) As ComboBox
    Private ArrayListCounter As Integer
    Private tableWrapper As ArcDataBinding.TableWrapper




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim clsModeAttributes As ModeAttributes
        Dim clsModeAttributes2 As ModeAttributes
        Dim tableSelect As ITableSelection
        Dim selectionSet As ISelectionSet
        Dim enumRow As IEnumRow
        Dim lID As Long
        Dim row As IRow
        Dim row2 As IRow

        Dim pEnumIDs As IEnumIDs
        Dim a As Integer



        'txtIJLanesGPAM.Text = cboIJ.SelectedItem


        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor






                m_mxDoc = m_application.Document
                Dim uid2 As New UID
                uid2.Value = "esriEditor.Editor"
                m_editor = CType(m_application.FindExtensionByCLSID(uid2), IEditor)
                pWrkspc = m_editor.EditWorkspace
                m_activeView = m_mxDoc.ActiveView

                m_map = m_activeView.FocusMap
                g_Schema = checkWS(m_editor, m_application)

                'm_RefEdges = GetFeatureLayer(GlobalConstants.g_Schema & g_RefEdge, m_application)
                'm_Junctions = GetFeatureLayer(g_Schema & g_RefJunct, m_application)
                'm_Projects = GetFeatureLayer(g_Schema & g_ProjectRoutes, m_application)



                'pDataset = CType(m_RefEdges.FeatureClass, IDataset)

                ' pWrkspc = pDataset.Workspace
        'CheckIJNode(m_map)
        tableSelect = m_ModeAttributes
        selectionSet = tableSelect.SelectionSet
        pEnumIDs = selectionSet.IDs
        lID = pEnumIDs.Next
        Dim m
        If RadioButton1.Checked = True Then
            Do Until lID = -1
                row = m_ModeAttributes.GetRow(lID)
                UpdateAttributes(row)

                lID = pEnumIDs.Next
            Loop
        Else
            'a = CType(BindingNavigator1.PositionItem.Text, Integer)
            row = CType(BindingSource1.Item(BindingSource1.Position), IRow)
            clsModeAttributes = New ModeAttributes(row)
            Do Until lID = -1
                row2 = m_ModeAttributes.GetRow(lID)
                clsModeAttributes2 = New ModeAttributes(row2)
                If clsModeAttributes2.PSRCEDGEID = clsModeAttributes.PSRCEDGEID Then
                    UpdateAttributes(row2)
                End If


                lID = pEnumIDs.Next
            Loop

            UpdateAttributes(row)

            'NOT A Row from mode attributes!!!!!!!!!1
        End If
        SetBindingSource(GetTableSelection)

        SetComboBoxtoCurrentRecord()
        System.Windows.Forms.Cursor.Current = Cursors.Default






    End Sub
    Public Sub UpdateAttributes(ByVal pRow As IRow)
        Dim clsModeAttributes As ModeAttributes
        
        Dim lID As Long
        Dim row As IRow

        Dim pEnumIDs As IEnumIDs


        
        Try



            ' row = m_ModeAttributes.GetRow(lID)
            clsModeAttributes = New ModeAttributes(pRow)
            clsModeAttributes.IJLANESGPAM = cboIJLanesGPAM.Text
            'MessageBox.Show(clsModeAttributes.PSRCEDGEID.ToString)
            clsModeAttributes.IJLANESGPEV = cboIJLanesGPEV.Text
            clsModeAttributes.IJLANESGPMD = cboIJLanesGPMD.Text
            clsModeAttributes.IJLANESGPNI = cboIJLanesGPNI.Text
            clsModeAttributes.IJLANESGPPM = cboIJLanesGPPM.Text

            clsModeAttributes.JILANESGPAM = cboJILanesGPAM.Text
            clsModeAttributes.JILANESGPEV = cboJILanesGPEV.Text
            clsModeAttributes.JILANESGPMD = cboJILanesGPMD.Text
            clsModeAttributes.JILANESGPNI = cboJILanesGPNI.Text
            clsModeAttributes.JILANESGPPM = cboJILanesGPNI.Text

            clsModeAttributes.IJLANESHOVAM = cboIJLanesHOVAM.SelectedValue
            clsModeAttributes.IJLANESHOVEV = cboIJLanesHOVEV.SelectedValue
            clsModeAttributes.IJLANESHOVMD = cboIJLanesHOVMD.SelectedValue
            clsModeAttributes.IJLANESHOVNI = cboIJLanesHOVNI.SelectedValue
            clsModeAttributes.IJLANESHOVPM = cboIJLanesHOVPM.SelectedValue



            clsModeAttributes.JILANESHOVAM = cboJILanesHOVAM.SelectedValue
            clsModeAttributes.JILANESHOVEV = cboJILanesHOVEV.SelectedValue
            clsModeAttributes.JILANESHOVMD = cboJILanesHOVMD.SelectedValue
            clsModeAttributes.JILANESHOVNI = cboJILanesHOVNI.SelectedValue
            clsModeAttributes.JILANESHOVPM = cboJILanesHOVPM.SelectedValue

            clsModeAttributes.IJSPEEDLIMIT = cboIJSpeedLimit.SelectedValue
            clsModeAttributes.IJVDFUNC = cboIJVDFunc.Text
            clsModeAttributes.IJSIDEWALKS = cboIJSideWalks.SelectedValue
            clsModeAttributes.IJBIKELANES = cboIJBikeLanes.SelectedValue
            clsModeAttributes.IJLANESTR = cboIJLanesTR.Text
            clsModeAttributes.IJLANESTK = cboIJLanesTK.Text
            clsModeAttributes.IJLANECAPHOV = cboIJLaneCapHOV.Text
            clsModeAttributes.IJLANECAPGP = cboIJLaneCapGP.Text

            clsModeAttributes.JISPEEDLIMIT = cboJISpeedLimit.SelectedValue
            clsModeAttributes.JIVDFUNC = cboJIVDFunc.Text
            clsModeAttributes.JISIDEWALKS = cboJISideWalks.SelectedValue
            clsModeAttributes.JIBIKELANES = cboJIBikeLanes.SelectedValue
            clsModeAttributes.JILANESTR = cboJILanesTR.Text
            clsModeAttributes.JILANESTK = cboJILanesTK.Text
            clsModeAttributes.JILANECAPHOV = cboJILaneCapHOV.Text
            clsModeAttributes.JILANECAPGP = cboJILaneCapGP.Text


            ';txtIJLanesGPAM.Text = clsModeAttributes.IJLANESGPAM



        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try






    End Sub

    Private Sub txtIJLanesGPAM_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub frmAttributeUpdater_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate

    End Sub

    Private Sub frmAttributeUpdater_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Me.Dispose()


    End Sub

    Private Sub frmAttributeUpdater_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim tableSelect As ITableSelection
        Dim SelectionSet As ISelectionSet
        Dim RelationshipClass As IRelationshipClass

        RadioButton1.Checked = True
        'relationshipclass = 


        m_ModeAttributes = getStandaloneTable(g_Schema & g_ModeAttributes, m_application)

        tableSelect = m_ModeAttributes
        SelectionSet = tableSelect.SelectionSet




        If SelectionSet.Count = 0 Then
            MessageBox.Show("No Records in ModeAttributes Selected", "No Records Selected", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
       

        ArrayListCounter = 1
        m_mxDoc = m_application.Document
        Dim uid2 As New UID
        uid2.Value = "esriEditor.Editor"
        m_editor = CType(m_application.FindExtensionByCLSID(uid2), IEditor)
        pWrkspc = m_editor.EditWorkspace







        Dim control As Control
        Dim myComboBox As New ComboBox
        Dim mycoll As Collection




        Try



            mycoll = GetRangeDomain("dLanes")

            'ComboBox6.DataSource = mycoll
            LoadGPComboBoxes(GroupBox1, mycoll)
            LoadGPComboBoxes(GroupBox2, mycoll)

            ' mycoll = GetCodedValueDomain("dHOVlanes")
            'LoadGPComboBoxes(GroupBox3, mycoll)
            'LoadGPComboBoxes(GroupBox4, mycoll)

            For Each control In GroupBox3.Controls
                If TypeOf control Is ComboBox Then

                    myArrayList(ArrayListCounter) = GetCodedValueDomainArrayList("dHOVlanes")
                    LoadHOVComboBoxes(CType(control, ComboBox), myArrayList(ArrayListCounter))
                    CodedValueComboBox(ArrayListCounter) = control
                    ArrayListCounter = ArrayListCounter + 1
                End If



            Next

            For Each control In GroupBox4.Controls
                If TypeOf control Is ComboBox Then
                    myArrayList(ArrayListCounter) = GetCodedValueDomainArrayList("dHOVlanes")
                    LoadHOVComboBoxes(CType(control, ComboBox), myArrayList(ArrayListCounter))
                    CodedValueComboBox(ArrayListCounter) = control
                    ArrayListCounter = ArrayListCounter + 1

                End If
            Next
            'myArrayList(ArrayListCounter) = GetCodedValueDomainArrayList("dHOVlanes")
            'LoadHOVComboBoxes(GroupBox, myArrayList(ArrayListCounter))

            'LoadHOVComboBoxes(GroupBox4, myArrayList)


            myArrayList(ArrayListCounter) = GetCodedValueDomainArrayList("dSpeedlimit")
            LoadHOVComboBoxes(cboIJSpeedLimit, myArrayList(ArrayListCounter))
            CodedValueComboBox(ArrayListCounter) = cboIJSpeedLimit
            ArrayListCounter = ArrayListCounter + 1

            myArrayList(ArrayListCounter) = GetCodedValueDomainArrayList("dSpeedlimit")
            LoadHOVComboBoxes(cboJISpeedLimit, myArrayList(ArrayListCounter))
            CodedValueComboBox(ArrayListCounter) = cboJISpeedLimit
            ArrayListCounter = ArrayListCounter + 1


            'mycoll = GetCodedValueDomain("dSpeedlimit")
            'LoadSingleComboBox(cboIJSpeedLimit, mycoll)
            'LoadSingleComboBox(cboJISpeedLimit, mycoll)



            mycoll = GetRangeDomain("dLaneCapacity")
            LoadSingleComboBox(cboIJLaneCapGP, mycoll)
            LoadSingleComboBox(cboIJLaneCapHOV, mycoll)
            LoadSingleComboBox(cboJILaneCapGP, mycoll)
            LoadSingleComboBox(cboJILaneCapHOV, mycoll)

            mycoll = GetRangeDomain("dVDFuncInt")
            LoadSingleComboBox(cboIJVDFunc, mycoll)
            LoadSingleComboBox(cboJIVDFunc, mycoll)


            myArrayList(ArrayListCounter) = GetCodedValueDomainArrayList("dSideWalks")
            LoadHOVComboBoxes(cboIJSideWalks, myArrayList(ArrayListCounter))
            CodedValueComboBox(ArrayListCounter) = cboIJSideWalks
            ArrayListCounter = ArrayListCounter + 1

            myArrayList(ArrayListCounter) = GetCodedValueDomainArrayList("dSideWalks")
            LoadHOVComboBoxes(cboJISideWalks, myArrayList(ArrayListCounter))
            CodedValueComboBox(ArrayListCounter) = cboJISideWalks
            ArrayListCounter = ArrayListCounter + 1

            'mycoll = GetCodedValueDomain("dSideWalks")
            'LoadSingleComboBox(cboIJSideWalks, mycoll)
            'LoadSingleComboBox(cboJISideWalks, mycoll)


            myArrayList(ArrayListCounter) = GetCodedValueDomainArrayList("dBikeLanes")
            LoadHOVComboBoxes(cboIJBikeLanes, myArrayList(ArrayListCounter))
            CodedValueComboBox(ArrayListCounter) = cboIJBikeLanes
            ArrayListCounter = ArrayListCounter + 1

            myArrayList(ArrayListCounter) = GetCodedValueDomainArrayList("dBikeLanes")
            LoadHOVComboBoxes(cboJIBikeLanes, myArrayList(ArrayListCounter))
            CodedValueComboBox(ArrayListCounter) = cboJIBikeLanes
            'ArrayListCounter = ArrayListCounter + 1


            'mycoll = GetCodedValueDomain("dBikeLanes")
            'LoadSingleComboBox(cboIJBikeLanes, mycoll)
            'LoadSingleComboBox(cboJIBikeLanes, mycoll)

            mycoll = GetRangeDomain("dLanes")
            LoadSingleComboBox(cboIJLanesTR, mycoll)
            LoadSingleComboBox(cboIJLanesTK, mycoll)
            LoadSingleComboBox(cboJILanesTR, mycoll)
            LoadSingleComboBox(cboJILanesTK, mycoll)


            'cboIJLanesGPAM.DataSource = mycoll
            'cboIJLanesGPMD.DataSource = mycoll
            'cboIJLanesGPPM.DataSource = mycoll

            mycoll = Nothing
            'ArrayListCounter = 0

            '  For Each control In Panel1.Controls
            'If TypeOf control Is ComboBox Then
            'myComboBox = CType(control, ComboBox)
            ' myComboBox.DataSource = mycoll





            ' End If
            ' Next
            ' For Each control In GroupBox2.Controls
            'If TypeOf control Is TextBox Then
            'control.Text = cboJI.SelectedItem

            ' End If
            'Next


            'Dim tableSelect As ITableSelection
            'Dim selectionSet As ISelectionSet
           
            'txtSelectedRecords.Text = selectionSet.Count.ToString










            'Dim pEnumIDs As IEnumIDs
            'Dim pRow As IRow



            'pEnumIDs = selectionSet.IDs
            'Dim lID As Long

            'lID = pEnumIDs.Next


            'pRow = m_ModeAttributes.GetRow(lID)

            SetBindingSource(selectionSet)

           SetComboBoxtoCurrentRecord()



            Me.Refresh()

            'Dim mycoll As Collection
            'mycoll = dLanes()
            'cboIJLanesGPAM.DataSource = mycoll


            'mycoll = domains()
            'ComboBox3.DataSource = mycoll
        Catch ex As Exception
            MessageBox.Show(ex.ToString)


        End Try



    End Sub
   

    Private Sub cboIJ_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboIJ.SelectedIndexChanged
        Dim control As Control
        For Each control In GroupBox1.Controls
            If TypeOf control Is ComboBox Then
                control.Text = cboIJ.SelectedItem

            End If
        Next


    End Sub


    Private Sub frmAttributeUpdater_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        'cboIJ.SelectedItem = 0
        'cboJI.SelectedItem = 0
    End Sub

    Private Sub cboJI_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboJI.SelectedIndexChanged
        Dim control As Control
        For Each control In GroupBox2.Controls
            If TypeOf control Is ComboBox Then
                control.Text = cboJI.SelectedItem

            End If
        Next
    End Sub
    Private Function GetCodedValueDomain(ByVal DomainName As String) As Collection



        ' Dim pFact As IWorkspaceFactory
        'pFact = New SdeWorkspaceFactory


        'Dim pWorkspace As IWorkspace
        'pWorkspace = pFact.OpenFromFile("C:\Documents and Settings\stefan.PSRC\Application Data\ESRI\ArcCatalog\psrcsqlsde.sde", 0)

        Dim pFeatws As IFeatureWorkspace
        pFeatws = pWrkspc

        Dim pWSDomains As IWorkspaceDomains
        pWSDomains = pWrkspc
        Dim x As Integer
        Dim pEnumDomain As IEnumDomain
        Dim pDomain As IDomain


        GetCodedValueDomain = New Collection





        'pEnumDomain = pWSDomains.Domains

        'pDomain = pEnumDomain.Next()
        'Do Until pEnumDomain Is Nothing
        'MessageBox.Show(pDomain.Name.ToString)
        'If pDomain.Name.ToString = "dBikeLanes" Then
        'Exit Do
        'End If
        'pDomain = pEnumDomain.Next





        'Loop




        'MessageBox.Show(pWSDomains.Domains.Next.Name)

        'Dim pDomain As IDomain
        pDomain = pWSDomains.DomainByName(DomainName)
        Dim pCodedValueDomain As ICodedValueDomain
        pCodedValueDomain = pDomain



        Try
            'MessageBox.Show(pCodedValueDomain.CodeCount)


            For x = 0 To pCodedValueDomain.CodeCount - 1



                GetCodedValueDomain.Add(pCodedValueDomain.Value(x))

                'ComboBox3.Items.Add(pCodedValueDomain.Value(x).ToString)


            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try
    End Function
    Private Function GetRangeDomain(ByVal DomainName As String) As Collection
        Dim pFeatws As IFeatureWorkspace


        Dim pWSDomains As IWorkspaceDomains

        Dim x As Integer
        Dim pEnumDomain As IEnumDomain

        Dim pDomain As IDomain








        'pEnumDomain = pWSDomains.Domains

        'pDomain = pEnumDomain.Next()
        'Do Until pEnumDomain Is Nothing
        'MessageBox.Show(pDomain.Name.ToString)
        'If pDomain.Name.ToString = "dBikeLanes" Then
        'Exit Do
        'End If
        'pDomain = pEnumDomain.Next





        'Loop




        'MessageBox.Show(pWSDomains.Domains.Next.Name)

        'Dim pDomain As IDomain




        Try
            pFeatws = pWrkspc
            pWSDomains = pWrkspc
            GetRangeDomain = New Collection
            pDomain = pWSDomains.DomainByName(DomainName)
            Dim pRangeDomain As IRangeDomain
            pRangeDomain = pDomain


            For x = pRangeDomain.MinValue To pRangeDomain.MaxValue

                GetRangeDomain.Add(x)





                'ComboBox3.Items.Add(pCodedValueDomain.Value(x).ToString)


            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try
    End Function

    Private Sub cboIJLanesGPAM_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboIJLanesGPAM.SelectedIndexChanged


    End Sub


    Private Sub LoadGPComboBoxes(ByVal myGroupBox As GroupBox, ByVal dataCollection As Collection)
        Dim myControl As Control
        Dim myComboBox As New ComboBox
        Try

            For Each myControl In myGroupBox.Controls
                If TypeOf myControl Is ComboBox Then
                    myComboBox = CType(myControl, ComboBox)
                    'MessageBox.Show(Len(myComboBox.Name.ToString) - InStr(myComboBox.Name.ToString, "GP"))
                    myComboBox.Items.Clear()

                    For x = 1 To dataCollection.Count

                        myComboBox.Items.Add(dataCollection.Item(x).ToString)
                        ' myComboBox.ValueMember = dataCollection.Item(1).ToString
                        myComboBox.Text = dataCollection.Item(1).ToString


                    Next


                End If
            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try
    End Sub
    Private Sub LoadSingleComboBox(ByVal myComboBox As ComboBox, ByVal dataCollection As Collection)
        'Dim myControl As Control
        'Dim myComboBox As New ComboBox
        Try



            ' myComboBox = CType(myControl, ComboBox)
            'MessageBox.Show(Len(myComboBox.Name.ToString) - InStr(myComboBox.Name.ToString, "GP"))
            myComboBox.Items.Clear()

            For x = 1 To dataCollection.Count

                myComboBox.Items.Add(dataCollection.Item(x).ToString)
                myComboBox.Text = dataCollection.Item(1).ToString



            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try
    End Sub



    Private Sub cboHOVIJ_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboHOVIJ.SelectedIndexChanged

        'MessageBox.Show(sender.ToString)
        Dim control As Control
        Dim myComboBox As ComboBox
        For Each control In GroupBox3.Controls
            If TypeOf control Is ComboBox Then
                myComboBox = control
                If myComboBox.ValueMember = "" Then
                    Exit Sub
                Else

                    myComboBox.SelectedItem = cboHOVIJ.SelectedItem
                    myComboBox.SelectedValue = cboHOVIJ.SelectedValue


                End If

            End If
        Next
    End Sub


    Private Function GetCodedValueDomainArrayList(ByVal DomainName As String) As ArrayList

        ' Dim pFact As IWorkspaceFactory
        'pFact = New SdeWorkspaceFactory


        'Dim pWorkspace As IWorkspace
        'pWorkspace = pFact.OpenFromFile("C:\Documents and Settings\stefan.PSRC\Application Data\ESRI\ArcCatalog\psrcsqlsde.sde", 0)

        Dim pFeatws As IFeatureWorkspace
        pFeatws = pWrkspc

        Dim pWSDomains As IWorkspaceDomains
        pWSDomains = pWrkspc
        Dim x As Integer
        Dim pEnumDomain As IEnumDomain
        Dim pDomain As IDomain


        GetCodedValueDomainArrayList = New ArrayList





        'pEnumDomain = pWSDomains.Domains

        'pDomain = pEnumDomain.Next()
        'Do Until pEnumDomain Is Nothing
        'MessageBox.Show(pDomain.Name.ToString)
        'If pDomain.Name.ToString = "dBikeLanes" Then
        'Exit Do
        'End If
        'pDomain = pEnumDomain.Next





        'Loop




        'MessageBox.Show(pWSDomains.Domains.Next.Name)

        'Dim pDomain As IDomain
        pDomain = pWSDomains.DomainByName(DomainName)
        Dim pCodedValueDomain As ICodedValueDomain
        pCodedValueDomain = pDomain







        Dim myCodedValue As CodedValue


        Try
            'MessageBox.Show(pCodedValueDomain.CodeCount)


            For x = 0 To pCodedValueDomain.CodeCount - 1
                myCodedValue = New CodedValue(pCodedValueDomain.Name(x), pCodedValueDomain.Value(x))
                GetCodedValueDomainArrayList.Add(myCodedValue)
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try

    End Function
    Private Sub LoadHOVComboBoxes(ByVal myComboBox As ComboBox, ByVal dataCollection As ArrayList)
        Dim myControl As Control
        'Dim myComboBox As New ComboBox
        Try




            'MessageBox.Show(Len(myComboBox.Name.ToString) - InStr(myComboBox.Name.ToString, "GP"))
            'myComboBox.Items.Clear()

            For x = 1 To dataCollection.Count

                myComboBox.DataSource = dataCollection
                myComboBox.ValueMember = "Value"
                myComboBox.DisplayMember = "Name"

                ' myComboBox.ValueMember = dataCollection.Item(1).ToString
                'myComboBox.Text = dataCollection.Item(1).ToString




            Next









        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try
    End Sub

    Private Sub cboHOVJI_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboHOVJI.SelectedIndexChanged
        'MessageBox.Show(sender.ToString)
        Dim control As Control
        Dim myComboBox As ComboBox
        For Each control In GroupBox4.Controls
            If TypeOf control Is ComboBox Then
                myComboBox = control
                If myComboBox.ValueMember = "" Then
                    Exit Sub
                Else

                    myComboBox.SelectedItem = cboHOVJI.SelectedItem
                    myComboBox.SelectedValue = cboHOVJI.SelectedValue


                End If

            End If
        Next
    End Sub
    Public Sub SetComboBoxesDefault(ByVal pRow As IRow)
        Dim clsModeAttributes As ModeAttributes
        Dim tableSelect As ITableSelection
        Dim selectionSet As ISelectionSet
        Dim enumRow As IEnumRow
        Dim lID As Long
        Dim row As IRow = New Row
        Dim myOID As Integer



        Dim row2 As IRow = New Row
        'Dim MyObjectClass As IObjectClass = New ObjectClass
        Dim MyTable As ITable

        Dim MyField As IField


        Dim pEnumIDs As IEnumIDs





        Try







            'pEnumIDs = mySelectionSet.IDs
            'lID = pEnumIDs.Next


            'row = m_ModeAttributes.GetRow(lID)
            clsModeAttributes = New ModeAttributes(pRow)
            'MessageBox.Show(clsModeAttributes.PSRCEDGEID.ToString)
            TextBox1.Text = clsModeAttributes.PSRCEDGEID.ToString



            cboIJLanesGPAM.Text = clsModeAttributes.IJLANESGPAM
            cboIJLanesGPEV.Text = clsModeAttributes.IJLANESGPEV
            cboIJLanesGPMD.Text = clsModeAttributes.IJLANESGPMD
            cboIJLanesGPNI.Text = clsModeAttributes.IJLANESGPNI
            cboIJLanesGPPM.Text = clsModeAttributes.IJLANESGPPM

            cboJILanesGPAM.Text = clsModeAttributes.JILANESGPAM
            cboJILanesGPEV.Text = clsModeAttributes.JILANESGPEV
            cboJILanesGPMD.Text = clsModeAttributes.JILANESGPMD
            cboJILanesGPNI.Text = clsModeAttributes.JILANESGPNI
            cboJILanesGPPM.Text = clsModeAttributes.JILANESGPPM

            Dim obj As Object
            Dim test As CodedValue
            Dim x As Integer = 1

            For x = 1 To ArrayListCounter
                For Each obj In myArrayList(x)
                    test = obj
                    If CodedValueComboBox(x).Name <> "cboHOVIJ" And CodedValueComboBox(x).Name <> "cboHOVJI" And test.Value = clsModeAttributes.AttributeByString(Mid(CodedValueComboBox(x).Name, 4)) Then
                        CodedValueComboBox(x).SelectedValue = test.Value
                        CodedValueComboBox(x).SelectedItem = test.Name
                        CodedValueComboBox(x).Text = test.Name.ToString
                        CodedValueComboBox(x).Refresh()
                    End If
                Next
                'x = x + 1
            Next


            ' For Each obj In myArrayList(1)
            'test = obj
            'If test.Value = clsModeAttributes.IJLANESHOVAM Then
            'cboIJILanesHOVAM.SelectedValue = test.Value
            ' cboIJILanesHOVAM.SelectedItem = test.Name
            ' cboIJILanesHOVAM.Text = test.Name.ToString


            'cboIJLanesHOVAM.Refresh()


            'End If
            ' Next



            'cboIJILanesHOVEV.SelectedItem = clsModeAttributes.IJLANESHOVEV
            'cboIJILanesHOVMD.SelectedItem = clsModeAttributes.IJLANESHOVMD
            'cboIJILanesHOVNI.SelectedItem = clsModeAttributes.IJLANESHOVNI
            'cboIJILanesHOVPM.SelectedItem = clsModeAttributes.IJLANESHOVPM



            'cboJILanesHOVAM.SelectedItem = clsModeAttributes.JILANESHOVAM
            'cboJILanesHOVEV.SelectedItem = clsModeAttributes.JILANESHOVEV
            'cboJILanesHOVMD.SelectedItem = clsModeAttributes.JILANESHOVMD
            'cboJILanesHOVNI.SelectedItem = clsModeAttributes.JILANESHOVNI
            ' cboJILanesHOVPM.SelectedItem = clsModeAttributes.JILANESHOVPM

            'cboIJSpeedLimit.SelectedItem = clsModeAttributes.IJSPEEDLIMIT
            cboIJVDFunc.Text = clsModeAttributes.IJVDFUNC
            'cboIJSideWalks.SelectedItem = CType(clsModeAttributes.IJSIDEWALKS, Long)

            'cboIJBikeLanes.SelectedItem = clsModeAttributes.IJBIKELANES
            cboIJLanesTR.Text = clsModeAttributes.IJLANESTR
            cboIJLanesTK.Text = clsModeAttributes.IJLANESTK
            cboIJLaneCapHOV.Text = clsModeAttributes.IJLANECAPHOV
            cboIJLaneCapGP.Text = clsModeAttributes.IJLANECAPGP

            'cboJISpeedLimit.SelectedItem = clsModeAttributes.JISPEEDLIMIT
            cboJIVDFunc.Text = clsModeAttributes.JIVDFUNC
            'cboJISideWalks.SelectedItem = clsModeAttributes.JISIDEWALKS
            'cboJIBikeLanes.SelectedItem = clsModeAttributes.JIBIKELANES
            cboJILanesTR.Text = clsModeAttributes.JILANESTR
            cboJILanesTK.Text = clsModeAttributes.JILANESTK
            cboJILaneCapHOV.Text = clsModeAttributes.JILANECAPHOV
            cboJILaneCapGP.Text = clsModeAttributes.JILANECAPGP

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim obj As Object
        Dim test As CodedValue

        For Each obj In myArrayList(1)
            test = obj
            'MessageBox.Show(test.Name & "" & test.Value)


        Next
    End Sub

    Private Sub cboIJLanesHOVAM_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub BindingNavigator1_Scroll(ByVal sender As Object, ByVal e As System.Windows.Forms.ScrollEventArgs)

    End Sub
    Private Function CreateTable() As ITable

        Dim pTable As ITable
        Dim pInMemWS As IFeatureWorkspace
        pInMemWS = CreateWorkspace("MyWorkspace4")

        'Dim pSRF As ISpatialReferenceFactory2
        'pSRF = New SpatialReferenceEnvironment
        '
        ' Dim pSR As ISpatialReference
        'pSR = pSRF.CreateGeographicCoordinateSystem(esriSRGeoCSType.esriSRGeoCS_NAD1983)
        ' won't work if I comment out next line
        'pSR.SetDomain(-200, 200, -200, 200)

        'Dim pGeoDefEdit As IGeometryDefEdit
        'pGeoDefEdit = New GeometryDef
        'With pGeoDefEdit
        ' .SpatialReference = pSR
        ''.GeometryType = esriGeometryType.esriGeometryPolyline
        '.GridCount = 1
        '.GridSize(0) = 1000
        ' End With

        Dim pFldsEdit As IFieldsEdit
        pFldsEdit = New Fields

        'pFldsEdit.AddField(MakeField("ObjectID", esriFieldTypeOID))
        'pFldsEdit.AddField(MakeField("Shape", esriFieldTypeGeometry, 0, pGeoDefEdit))

        Dim pFWS As IFeatureWorkspace
        pFWS = pInMemWS

        Dim pUID As New UID
        pUID.Value = "esriGeodatabase.Object"

        'Dim pLineFC As IFeatureClass
        'pLineFC = pFWS.CreateFeatureClass("mylines2", pFldsEdit, pUID, Nothing, esriFTSimple, "Shape", "")



        Dim y As Integer
        For y = 0 To m_ModeAttributes.Fields.FieldCount - 1
            pFldsEdit.AddField(m_ModeAttributes.Fields.Field(y))


        Next



        'Dim pTable As ITable
        pTable = pFWS.CreateTable("SelecteModeAttributeRecords", pFldsEdit, pUID, Nothing, Nothing)
        CreateTable = pTable



    End Function

    Function CreateWorkspace(ByVal sName As String) As IWorkspace
        Dim pWSF As IWorkspaceFactory
        pWSF = New InMemoryWorkspaceFactory
        Dim pName As IName
        pName = pWSF.Create("", sName, Nothing, 0)
        CreateWorkspace = pName.Open
    End Function

    Private Sub BindingNavigator1_RefreshItems(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SetComboBoxtoCurrentRecord()

    End Sub

    Private Sub BindingNavigatorMoveNextItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        SetComboBoxtoCurrentRecord()





    End Sub
    Private Sub SetBindingSource(ByVal mySelectionSet As ISelectionSet)
        Dim clsModeAttributes As ModeAttributes
        Dim tableSelect As ITableSelection
        Dim selectionSet As ISelectionSet
        Dim enumRow As IEnumRow
        Dim lID As Long
        Dim row As IRow = New Row
        Dim myOID As Integer



        Dim row2 As IRow = New Row
        'Dim MyObjectClass As IObjectClass = New ObjectClass
        Dim MyTable As ITable

        Dim MyField As IField


        Dim pEnumIDs As IEnumIDs


        'txtIJLanesGPAM.Text = cboIJ.SelectedItem


        Try


            MyTable = CreateTable()
            pEnumIDs = mySelectionSet.IDs
            lID = pEnumIDs.Next
            Do Until lID = -1


                row = MyTable.CreateRow
                row.Store()





                'row = MyTable.GetRow(myOID)
                row2 = m_ModeAttributes.GetRow(lID)
                Dim test As New ModeAttributes(row2)
                'MessageBox.Show(test.PSRCEDGEID.ToString)


                TempModeAttributes(row2, row, m_application)






                lID = pEnumIDs.Next



            Loop

            tableWrapper = New ArcDataBinding.TableWrapper(MyTable)


            Me.BindingSource1.DataSource = tableWrapper

        Catch ex As Exception

        End Try
    End Sub
    Private Sub SetComboBoxtoCurrentRecord()
        Dim row As IRow
        Dim x As Integer
        Dim pEnumIDs As IEnumIDs
        Dim clsModeAttributes As ModeAttributes

        'pEnumIDs = mySelectionSet.IDs
        'row = BindingNavigator1.PositionItem.Text
        ' For x = 1 To BindingSource1.Count
        'BindingNavigator1.Update()

        'If CType(BindingNavigator1.PositionItem.Text, Integer) = x Then
        If Not BindingSource1.DataSource Is Nothing Then
            row = CType(BindingSource1.Item(BindingSource1.Position), IRow)
            SetBindingSource(GetTableSelection)

            SetComboBoxesDefault(row)
        End If
      
        'Exit Sub





        'End If

    End Sub

    Private Sub BindingNavigator1_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles BindingNavigator1.ItemClicked

    End Sub

   

  

   



    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub BindingNavigator1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles BindingNavigator1.TextChanged
        'MessageBox.Show("text changed")

    End Sub
    Public Function GetTableSelection() As SelectionSet
        Dim tableSelect As ITableSelection
        Dim selectionSet As ISelectionSet
        tableSelect = m_ModeAttributes
        selectionSet = tableSelect.SelectionSet
        GetTableSelection = selectionSet

    End Function

 

  


  

    Private Sub BindingNavigatorPositionItem_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles BindingNavigatorPositionItem.TextChanged
        SetComboBoxtoCurrentRecord()
    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click

    End Sub

    Private Sub BindingNavigatorMovePreviousItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorMovePreviousItem.Click

    End Sub

   

    Private Sub ToolStripLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel1.Click

    End Sub
End Class

