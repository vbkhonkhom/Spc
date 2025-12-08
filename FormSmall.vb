Public Class FormSmall
    'ラベルコントロール配列のフィールドを作成

    Public LabXBar_Small() As System.Windows.Forms.Label
    Public LabR_Small() As System.Windows.Forms.Label
    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label17.Click

    End Sub

    Private Sub Label36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label16.Click

    End Sub

    Private Sub Label30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label14.Click

    End Sub

    Private Sub Label19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label19.Click

    End Sub

    Private Sub Label20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label20.Click

    End Sub

    Private Sub Label29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label29.Click

    End Sub

    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click

    End Sub

    Private Sub FormSmall_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Form1.Close()
    End Sub

    Private Sub FormSmall_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '日本語・英語表記の切り替えを行う
        FormAlarmInput.Translation_AlarmInput()
        FormAlarmDisp.Translation_AlarmDisp()
        FormControl.Translation_FormControl()

        'X側の軸ラベルを配列化する
        Me.LabXBar_Small = New System.Windows.Forms.Label(10) {}
        Me.LabXBar_Small(0) = Me.Label13
        Me.LabXBar_Small(1) = Me.Label12
        Me.LabXBar_Small(2) = Me.Label3
        Me.LabXBar_Small(3) = Me.Label4
        Me.LabXBar_Small(4) = Me.Label5
        Me.LabXBar_Small(5) = Me.Label6
        Me.LabXBar_Small(6) = Me.Label7
        Me.LabXBar_Small(7) = Me.Label8
        Me.LabXBar_Small(8) = Me.Label9
        Me.LabXBar_Small(9) = Me.Label10
        Me.LabXBar_Small(10) = Me.Label11

        'R側の軸ラベルを配列化する
        Me.LabR_Small = New System.Windows.Forms.Label(7) {}
        Me.LabR_Small(0) = Me.Label29
        Me.LabR_Small(1) = Me.Label20
        Me.LabR_Small(2) = Me.Label19
        Me.LabR_Small(3) = Me.Label14
        Me.LabR_Small(4) = Me.Label15
        Me.LabR_Small(5) = Me.Label16
        Me.LabR_Small(6) = Me.Label17
        Me.LabR_Small(7) = Me.Label18

        Me.Top = 0
        Me.Left = 0
    End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Form1.TreeView_AfterSelect(TreeView1)
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    'PopUpを表示する
    Private Sub _MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove, PictureBox2.MouseMove
        popUp(e.X, e.Y, sender.name)
    End Sub
    Private Sub _MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseLeave, PictureBox2.MouseLeave
        i_old = 1000
        FormPopupNew.Close()
    End Sub
    'アラームコメントを表示・入力する
    Private Sub _MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDown, PictureBox2.MouseDown
        alarmInfo(e.X, e.Y, sender.name, e.Button.ToString)
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click

    End Sub



  

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        If StrServerConnection = "" Then
            Dim dr As DialogResult
            Dim frm As New Form3
            dr = frm.ShowDialog
            If dr = Windows.Forms.DialogResult.OK Then
                Form1.LoadLoad()
            ElseIf dr = Windows.Forms.DialogResult.Cancel Then
                Me.Close()
            End If

        Else
            Form1.LoadLoad()
        End If
        MsgBox("Tree Updated")
    End Sub

    Private Sub ExToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExToolStripMenuItem.Click
        'Form1.SPCData_Export() '表示中のSPCデータを出力する
    End Sub

    Private Sub DToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DToolStripMenuItem.Click
        FormProperty.Show() 'プロパティ画面を表示する
    End Sub

    Private Sub LanguageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LanguageToolStripMenuItem.Click

    End Sub

    Private Sub JapaneseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JapaneseToolStripMenuItem.Click
        StrLanguage = "Japanese"
        Form1.SaveConfigData()
        FormAlarmInput.Translation_AlarmInput()
        FormAlarmDisp.Translation_AlarmDisp()
        FormControl.Translation_FormControl()
        MsgBox("表示言語を日本語に変更しました")
    End Sub

    Private Sub EnglishToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnglishToolStripMenuItem.Click
        StrLanguage = "English"
        Form1.SaveConfigData()
        FormAlarmInput.Translation_AlarmInput()
        FormAlarmDisp.Translation_AlarmDisp()
        FormControl.Translation_FormControl()
        MsgBox("Changed display language to English")
    End Sub

    Private Sub HowToUseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HowToUseToolStripMenuItem.Click
        FormHowto.Show() '使用マニュアルを表示する
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        StrResolution = "MAX"
        Form1.SaveConfigData()
        Form1.Show()
        Me.Hide()
    End Sub

    Private Sub CreateNewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreateNewToolStripMenuItem.Click
        FormControl.Show() 'グラフ初期設定画面を表示する
    End Sub

    Private Sub SPCRuleToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SPCRuleToolStripMenuItem1.Click
        FormSPCRule.Show() 'SPC詳細設定画面を表示する
    End Sub



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

    End Sub

    Private Sub Button2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button2.MouseDown
        'グラフ左移動START
        Form1.Timer1.Enabled = True
    End Sub

    Private Sub Button2_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button2.MouseUp
        'グラフ左移動STOP
        Form1.Timer1.Enabled = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

    End Sub

    Private Sub Button3_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button3.MouseDown
        'グラフ右移動START
        Form1.Timer2.Enabled = True
    End Sub

    Private Sub Button3_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button3.MouseUp
        'グラフ右移動START
        Form1.Timer2.Enabled = False
    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click

    End Sub

    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click

    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        StrResolution = "Middle"
        Form1.SaveConfigData()
        FormMiddle.Show()
        Me.Hide()
    End Sub



    Private Sub TreeView1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TreeView1.MouseMove

    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        MRFlag = True
        GraphDisp()

        Label1.Text = "MR"
        GroupBox2.Text = "MR"

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        MRFlag = False
        GraphDisp()
        Label1.Text = "R"
        GroupBox2.Text = "R"

    End Sub


    Private Sub ServerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ServerToolStripMenuItem.Click
        Dim dr As DialogResult
        Dim frm As New Form3
        dr = frm.ShowDialog
    End Sub

    Private Sub DeleteAlarmTableToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteAlarmTableToolStripMenuItem.Click
        Form1.Delete_AlarmTable()
    End Sub

    Private Sub StandardNoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StandardNoToolStripMenuItem.Click
        Form1.StandardNo()
    End Sub

    Private Sub UserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserToolStripMenuItem.Click
        Form1.User()
    End Sub
End Class