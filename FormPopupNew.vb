Public Class FormPopupNew
    Public Values0() As System.Windows.Forms.Label
    Public Values() As System.Windows.Forms.Label
    Public Labels() As System.Windows.Forms.Label

    Private Sub FormPopupNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim s As System.Windows.Forms.Screen = System.Windows.Forms.Screen.FromControl(Me)
        Me.Top = mouseY
        'Me.Left = mouseX
    End Sub

End Class