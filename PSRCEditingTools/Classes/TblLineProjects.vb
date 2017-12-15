Imports ESRI.ArcGIS.Geodatabase
Public Class TblLineProjects
    Inherits RowByEdgeID
    'Private modeAttributeRow As IRow
    Public Sub New(ByVal row As IRow)
        MyBase.New(row)
        'modeAttributeRow = row
    End Sub


    Public Property projRteID() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("projRteID")
            projRteID = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("projRteID")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property

    Public Property projDBS() As String
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("projDBS")
            projDBS = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("projDBS")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property
    Public Property projID() As String
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("projID")
            projID = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("projID")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property
    Public Property version() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("version")
            version = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("version")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property
    Public Property InServiceDate() As Short
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("InServiceDate")
            InServiceDate = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Short)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("InServiceDate")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property
    Public Property OutServiceDate() As Short
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("OutServiceDate")
            OutServiceDate = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Short)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("OutServiceDate")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property
    Public Property CompletionDate() As Short
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("CompletionDate")
            CompletionDate = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Short)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("CompletionDate")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property
    Public Property Modes() As String
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("Modes")
            Modes = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("Modes")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property
    Public Property EditNotes() As String
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("EditNotes")
            EditNotes = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As String)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("EditNotes")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property
    Public Property Processing() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("Processing")
            Processing = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("Processing")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property
End Class
