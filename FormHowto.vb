Public Class FormHowto

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim proc As New Process()                 ' （1）
        proc.StartInfo.FileName = StrCDir & "\グラフを表示する.xlsx"
        proc.Start()

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim proc As New Process()                 ' （1）
        proc.StartInfo.FileName = StrCDir & "\データを更新する.xlsx"
        proc.Start()
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim proc As New Process()                 ' （1）
        proc.StartInfo.FileName = StrCDir & "\グラフの拡大縮小を行う.xlsx"
        proc.Start()
    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim proc As New Process()                 ' （1）
        proc.StartInfo.FileName = StrCDir & "\アラームコメントの入力・閲覧を行う.xlsx"
        proc.Start()
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim proc As New Process()                 ' （1）
        proc.StartInfo.FileName = StrCDir & "\SPCルール・管理値を変更する.xlsx"
        proc.Start()
    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click

        Dim strFilePath As String = StrCDir & "\Manual\Manual1.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label42.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual1.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual2.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual2.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual2.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual2.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual3.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual3.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label14.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual5.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label13.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual5.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label16.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual4.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual4.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label18.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual6.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label17.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual6.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label20.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual7.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label19.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual7.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual8.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual8.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual9.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub Label4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        Dim strFilePath As String = StrCDir & "\Manual\Manual9.png"
        FormPicture.PictureBox1.Image = System.Drawing.Image.FromFile(strFilePath)
        FormPicture.Show()
    End Sub

    Private Sub FormHowto_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class