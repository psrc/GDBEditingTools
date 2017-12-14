Public Class TestEvents

    Public Event EditorOn(ByVal IsOn As Boolean)

    Public Sub New()

    End Sub

    Public Sub Trigger(ByVal test As Boolean)
        If test = True Then
            RaiseEvent EditorOn(True)
        Else
            RaiseEvent EditorOn(False)
        End If
    End Sub





End Class
