Public Class Form3

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then
            MsgBox("Enter the server connection string.")
            Exit Sub
        End If
        StrServerConnection = TextBox2.Text
        Form1.SaveConfigData()
        Form1.LoadLoad()
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox2.Text = StrServerConnection
    End Sub
End Class