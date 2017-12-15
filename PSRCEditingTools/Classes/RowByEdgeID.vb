Imports ESRI.ArcGIS.Geodatabase
Public Class RowByEdgeID
    Public modeAttributeRow As IRow
    Public Sub New(ByVal row As IRow)
        modeAttributeRow = row
    End Sub
    Public ReadOnly Property OID() As Long
        Get
            'Dim intPos As Integer
            'intPos = modeAttributeRow.Fields.FindField("PSRCEDGEID")
            OID = modeAttributeRow.Value(0)
        End Get
    End Property


    

    Public Property IJLANESGPAM() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESGPAM")
            IJLANESGPAM = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESGPAM")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property






    Public Property IJLANESGPMD() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESGPMD")
            IJLANESGPMD = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESGPMD")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property IJLANESGPPM() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESGPPM")
            IJLANESGPPM = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESGPPM")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property IJLANESGPEV() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESGPEV")
            IJLANESGPEV = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESGPEV")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property



    Public Property IJLANESGPNI() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESGPNI")
            IJLANESGPNI = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESGPNI")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property



    Public Property JILANESGPAM() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESGPAM")
            JILANESGPAM = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESGPAM")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property JILANESGPMD() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESGPMD")
            JILANESGPMD = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESGPMD")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property JILANESGPPM() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESGPPM")
            JILANESGPPM = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESGPPM")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property

    Public Property JILANESGPEV() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESGPEV")
            JILANESGPEV = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intpos As Integer
            intpos = modeAttributeRow.Fields.FindField("JILANESGPEV")
            modeAttributeRow.Value(intpos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property JILANESGPNI() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESGPNI")
            JILANESGPNI = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESGPNI")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property IJLANESGPADJUST() As Double
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESGPADJUST")
            IJLANESGPADJUST = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Double)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESGPADJUST")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property



    Public Property JILANESGPADJUST() As Double
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESGPADJUST")
            JILANESGPADJUST = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Double)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESGPADJUST")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property



    Public Property IJLANESHOVAM() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESHOVAM")
            IJLANESHOVAM = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESHOVAM")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property IJLANESHOVMD() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESHOVMD")
            IJLANESHOVMD = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESHOVMD")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property IJLANESHOVPM() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESHOVPM")
            IJLANESHOVPM = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESHOVPM")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property




    Public Property IJLANESHOVEV() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESHOVEV")
            IJLANESHOVEV = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESHOVEV")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property IJLANESHOVNI() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESHOVNI")
            IJLANESHOVNI = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESHOVNI")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property

    Public Property JILANESHOVAM() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESHOVAM")
            JILANESHOVAM = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESHOVAM")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property JILANESHOVMD() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESHOVMD")
            JILANESHOVMD = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESHOVMD")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property JILANESHOVPM() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESHOVPM")
            JILANESHOVPM = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESHOVPM")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property JILANESHOVEV() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESHOVEV")
            JILANESHOVEV = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESHOVEV")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property JILANESHOVNI() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESHOVNI")
            JILANESHOVNI = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESHOVNI")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property IJSPEEDLIMIT() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJSPEEDLIMIT")
            IJSPEEDLIMIT = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJSPEEDLIMIT")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property JISPEEDLIMIT() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JISPEEDLIMIT")
            JISPEEDLIMIT = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JISPEEDLIMIT")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property

    Public Property IJFFS() As Double
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJFFS")
            IJFFS = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Double)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJFFS")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property



    Public Property JIFFS() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JIFFS")
            JIFFS = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JIFFS")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property IJVDFUNC() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJVDFUNC")
            IJVDFUNC = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJVDFUNC")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property



    Public Property JIVDFUNC() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JIVDFUNC")
            JIVDFUNC = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JIVDFUNC")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property IJLANECAPGP() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANECAPGP")
            IJLANECAPGP = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANECAPGP")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property

    Public Property IJLANECAPHOV() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANECAPHOV")
            IJLANECAPHOV = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANECAPHOV")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property JILANECAPGP() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANECAPGP")
            JILANECAPGP = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANECAPGP")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property JILANECAPHOV() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANECAPHOV")
            JILANECAPHOV = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANECAPHOV")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property IJSIDEWALKS() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJSIDEWALKS")
            IJSIDEWALKS = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJSIDEWALKS")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property JISIDEWALKS() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JISIDEWALKS")
            JISIDEWALKS = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JISIDEWALKS")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property IJBIKELANES() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJBIKELANES")
            IJBIKELANES = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJBIKELANES")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property



    Public Property JIBIKELANES() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JIBIKELANES")
            JIBIKELANES = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JIBIKELANES")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property IJLANESTR() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESTR")
            IJLANESTR = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESTR")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property JILANESTR() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESTR")
            JILANESTR = modeAttributeRow.Value(intPos)

        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESTR")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property IJLANESTK() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESTK")
            IJLANESTK = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("IJLANESTK")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property


    Public Property JILANESTK() As Long
        Get
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESTK")
            JILANESTK = modeAttributeRow.Value(intPos)
        End Get
        Set(ByVal value As Long)
            Dim intPos As Integer
            intPos = modeAttributeRow.Fields.FindField("JILANESTK")
            modeAttributeRow.Value(intPos) = value
            modeAttributeRow.Store()
        End Set
    End Property
End Class
