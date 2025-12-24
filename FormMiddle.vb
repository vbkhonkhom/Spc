
Imports System.IO
Public Class FormMiddle
    Public LabXBar_Middle() As System.Windows.Forms.Label
    Public LabR_Middle() As System.Windows.Forms.Label

    Private Sub FormMiddle_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Form1.Close()
    End Sub
    Private Sub FormMiddle_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '日本語・英語表記の切り替えを行う
        FormAlarmInput.Translation_AlarmInput()
        FormAlarmDisp.Translation_AlarmDisp()
        FormControl.Translation_FormControl()
        'X側の軸ラベルを配列化する
        Me.LabXBar_Middle = New System.Windows.Forms.Label(10) {}
        Me.LabXBar_Middle(0) = Me.Label13
        Me.LabXBar_Middle(1) = Me.Label12
        Me.LabXBar_Middle(2) = Me.Label3
        Me.LabXBar_Middle(3) = Me.Label4
        Me.LabXBar_Middle(4) = Me.Label5
        Me.LabXBar_Middle(5) = Me.Label6
        Me.LabXBar_Middle(6) = Me.Label7
        Me.LabXBar_Middle(7) = Me.Label8
        Me.LabXBar_Middle(8) = Me.Label9
        Me.LabXBar_Middle(9) = Me.Label10
        Me.LabXBar_Middle(10) = Me.Label11

        'R側の軸ラベルを配列化する
        Me.LabR_Middle = New System.Windows.Forms.Label(8) {}
        Me.LabR_Middle(0) = Me.Label29
        Me.LabR_Middle(1) = Me.Label20
        Me.LabR_Middle(2) = Me.Label19
        Me.LabR_Middle(3) = Me.Label14
        Me.LabR_Middle(4) = Me.Label15
        Me.LabR_Middle(5) = Me.Label16
        Me.LabR_Middle(6) = Me.Label17
        Me.LabR_Middle(7) = Me.Label18
        Me.LabR_Middle(8) = Me.Label2

        Me.Top = 0
        Me.Left = 0
    End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Form1.TreeView_AfterSelect(TreeView1)
        'StrSerch1 = ""
        'StrSerch2 = ""
        'StrSerch3 = ""

        ''親ノードがない場合(ノードを開いていない場合)この関数を終了する
        'If TreeView1.SelectedNode.Parent Is Nothing Then Exit Sub
        'If TreeView1.SelectedNode.Parent.Parent Is Nothing Then Exit Sub
        'If TreeView1.SelectedNode.Parent.Parent.Parent Is Nothing Then Exit Sub
        'If TreeView1.SelectedNode.Parent.Parent.Parent.Parent Is Nothing Then Exit Sub

        'If TreeView1.SelectedNode.Parent.Parent.Parent.Parent.Parent Is Nothing Then
        '    '子ノード名を取得する
        '    Dim SelectNode As TreeNode = TreeView1.SelectedNode
        '    SPCMcNo = SelectNode.Text
        '    '子ノードの親ノード名を取得する
        '    Dim SelectNodeparent As TreeNode = TreeView1.SelectedNode.Parent
        '    SPCkoumoku = SelectNodeparent.Text
        '    '親ノードの親ノードの親ノード名を取得する
        '    Dim SelectNodeparent1 As TreeNode = TreeView1.SelectedNode.Parent.Parent
        '    StrSelectMc = SelectNodeparent1.Text
        '    '親ノードの親ノードの親ノードの親ノード名を取得する
        '    Dim SelectNodeparent2 As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent
        '    StrSerchMonitor = SelectNodeparent2.Text
        '    '親ノードの親ノードの親ノードの親ノードの親ノード名を取得する
        '    Dim SelectNodehost As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent.Parent
        '    StrSerchHostItem = SelectNodehost.Text
        'Else
        '    If TreeView1.SelectedNode.Parent.Parent.Parent.Parent.Parent.Parent Is Nothing Then
        '        '絞込み条件が一つの場合
        '        Dim Serch1Node As TreeNode = TreeView1.SelectedNode
        '        StrSerch1 = Serch1Node.Text
        '        '子ノード名を取得する
        '        Dim SelectNode As TreeNode = TreeView1.SelectedNode.Parent
        '        SPCMcNo = SelectNode.Text
        '        '子ノードの親ノード名を取得する
        '        Dim SelectNodeparent As TreeNode = TreeView1.SelectedNode.Parent.Parent
        '        SPCkoumoku = SelectNodeparent.Text
        '        '親ノードの親ノードの親ノード名を取得する
        '        Dim SelectNodeparent1 As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent
        '        StrSelectMc = SelectNodeparent1.Text
        '        '親ノードの親ノードの親ノードの親ノード名を取得する
        '        Dim SelectNodeparent2 As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent.Parent
        '        StrSerchMonitor = SelectNodeparent2.Text
        '        '親ノードの親ノードの親ノードの親ノードの親ノード名を取得する
        '        Dim SelectNodehost As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent.Parent.Parent
        '        StrSerchHostItem = SelectNodehost.Text
        '    Else
        '        If TreeView1.SelectedNode.Parent.Parent.Parent.Parent.Parent.Parent.Parent Is Nothing Then
        '            '絞込み条件が二つの場合
        '            Dim Serch1Node As TreeNode = TreeView1.SelectedNode.Parent
        '            StrSerch1 = Serch1Node.Text
        '            Dim Serch2Node As TreeNode = TreeView1.SelectedNode
        '            StrSerch2 = Serch2Node.Text
        '            '子ノード名を取得する
        '            Dim SelectNode As TreeNode = TreeView1.SelectedNode.Parent.Parent
        '            SPCMcNo = SelectNode.Text
        '            '子ノードの親ノードの親ノード名を取得する
        '            Dim SelectNodeparent As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent
        '            SPCkoumoku = SelectNodeparent.Text
        '            '親ノードの親ノードの親ノードの親ノード名を取得する
        '            Dim SelectNodeparent2 As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent.Parent
        '            StrSelectMc = SelectNodeparent2.Text
        '            '親ノードの親ノードの親ノードの親ノードの親ノード名を取得する
        '            Dim SelectNodeparent3 As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent.Parent.Parent
        '            StrSerchMonitor = SelectNodeparent3.Text
        '            '親ノードの親ノードの親ノードの親ノードの親ノードの親ノード名を取得する
        '            Dim SelectNodehost As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent.Parent.Parent.Parent
        '            StrSerchHostItem = SelectNodehost.Text
        '        Else
        '            '絞込み条件が三つの場合
        '            Dim Serch1Node As TreeNode = TreeView1.SelectedNode.Parent.Parent
        '            StrSerch1 = Serch1Node.Text
        '            Dim Serch2Node As TreeNode = TreeView1.SelectedNode.Parent
        '            StrSerch2 = Serch2Node.Text
        '            Dim Serch3Node As TreeNode = TreeView1.SelectedNode
        '            StrSerch3 = Serch3Node.Text
        '            '子ノード名を取得する
        '            Dim SelectNode As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent
        '            SPCMcNo = SelectNode.Text
        '            '子ノードの親ノード名を取得する
        '            Dim SelectNodeparent1 As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent.Parent
        '            SPCkoumoku = SelectNodeparent1.Text
        '            '親ノードの親ノードの親ノード名を取得する
        '            Dim SelectNodeparent2 As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent.Parent.Parent
        '            StrSelectMc = SelectNodeparent2.Text
        '            '親ノードの親ノードの親ノードの親ノード名を取得する
        '            Dim SelectNodeparent3 As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent.Parent.Parent.Parent
        '            StrSerchMonitor = SelectNodeparent3.Text
        '            '親ノードの親ノードの親ノードの親ノードの親ノード名を取得する
        '            Dim SelectNodehost As TreeNode = TreeView1.SelectedNode.Parent.Parent.Parent.Parent.Parent.Parent.Parent
        '            StrSerchHostItem = SelectNodehost.Text
        '        End If

        '    End If
        'End If

        'Me.ComboDevice.Text = SPCMcNo
        'Me.ComboItem1.Text = StrSerch1
        'Me.TextItem2.Text = StrSerch2
        'Me.TextItem3.Text = StrSerch3
        'Graphsmallcount = 1
        ''上記の条件でSPCデータを取得する
        'MRFlag = False
        'Label1.Text = "R"
        'GroupBox2.Text = "R"
        'Form1.GraphDisp()
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
        'グラフ右移動STOP
        Form1.Timer2.Enabled = False
    End Sub


    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click

    End Sub

    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click

    End Sub

    Private Sub ExToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExToolStripMenuItem.Click
        'Form1.SPCData_Export() '表示中のSPCデータを出力する
    End Sub

    Private Sub DToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DToolStripMenuItem.Click
        FormProperty.Show() 'プロパティ画面を表示する
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

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        StrResolution = "MIN"
        Form1.SaveConfigData()
        FormSmall.Show()
        Me.Hide()
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

    Private Sub SPCRuleToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SPCRuleToolStripMenuItem1.Click
        FormSPCRule.Show() 'SPC詳細設定画面を表示する
    End Sub

    Private Sub CreateNewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreateNewToolStripMenuItem.Click
        FormControl.Show() 'グラフ初期設定画面を表示する
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        If StrServerConnection = "" Then
            Dim dr As DialogResult
            Dim frm As New Form3
            dr = frm.ShowDialog
            If dr = System.Windows.Forms.DialogResult.OK Then
                Form1.LoadLoad()
            ElseIf dr = System.Windows.Forms.DialogResult.Cancel Then
                Me.Close()
            End If

        Else
            Form1.LoadLoad()
        End If
        MsgBox("Tree Updated")
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

    Private Sub UserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserToolStripMenuItem.Click
        Form1.User()
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

    Private Sub ButtonLoad_Click(sender As Object, e As EventArgs) Handles ButtonLoad.Click
        LoadDataFromTextFile()
        GraphDisp()
    End Sub

End Class