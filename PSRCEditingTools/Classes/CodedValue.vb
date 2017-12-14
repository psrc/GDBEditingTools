Public Class CodedValue
    Private myStrName As String
    Private myLngValue As String

    Public Sub New(ByVal strName As String, ByVal lngValue As Long)
        Me.myStrName = strName
        Me.myLngValue = lngValue
    End Sub 'New

    Public ReadOnly Property Name() As String
        Get
            Return myStrName
        End Get
    End Property

    Public ReadOnly Property Value() As String
        Get
            Return myLngValue
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return Me.myStrName + " - " + Me.myLngValue
    End Function 'ToString
End Class 'USState


