Imports System
Imports Microsoft.VisualBasic
Imports System.IO ' <--- [1] เพิ่ม Library สำหรับอ่านไฟล์
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Net
Imports System.Data.SqlClient

Public Class Form1
    Dim myHostName As String
    Dim FROMHOST_SPC_Yukuhashi As String
    '受信用のイベント 

    Private WithEvents objEv As New ClassLibrary1.RemoteHttp
    'COPYDATASTRUCT構造体
    Public Structure COPYDATASTRUCT
        Public dwData As Int32      '送信するビット値
        Public cbData As Int32      'lpDataのバイト数
        Public lpData As String     '送信するデータへのポインタ(0も可能)
    End Structure


    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindow(
            ByVal lpClassName As String,
            ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function SendMessage(
                           ByVal hWnd As IntPtr,
                           ByVal wMsg As Int32,
                           ByVal wParam As Int32,
                           ByVal lParam As Int32) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function SendMessage(
                            ByVal hWnd As IntPtr,
                            ByVal wMsg As Int32,
                            ByVal wParam As Int32,
                            ByRef lParam As COPYDATASTRUCT) As Integer
    End Function

    Public Const WM_COPYDATA As Int32 = &H4A
    Public Const WM_USER As Int32 = &H400

    Private Sub SendCommand(ByVal strMes As String)
        Dim result As Int32 = 0

        '相手のウィンドウハンドルを取得します
        Dim hWnd As Int32 = FindWindow(Nothing, "YPCS Client")
        If hWnd = 0 Then
            'ハンドルが取得できなかった
            If MessageBoxShowFlag = False Then
                MessageBoxShowFlag = True
                MessageBox.Show("YPCS Clientアプリが動作していません。")
            End If

            Return
        End If

        MessageBoxShowFlag = False
        '文字列メッセージを送信します
        If strMes <> String.Empty Then
            '送信データをByte配列に格納
            Dim bytearry() As Byte =
                         System.Text.Encoding.Default.GetBytes(strMes)
            Dim len As Int32 = bytearry.Length
            Dim cds As COPYDATASTRUCT
            cds.dwData = 0        '使用しない
            cds.lpData = strMes 'テキストのポインターをセット
            cds.cbData = len + 1   '長さをセット
            '文字列を送る
            result = SendMessage(hWnd, WM_COPYDATA, 0, cds)
        End If

    End Sub

    Public Sub WSSendToSelcom(ByVal HEADER As String, ByVal FROMHOST As String, ByVal TOHOST As String, ByVal CONTENT1 As String, ByVal CONTENT2 As String, ByVal CONTENT3 As String, ByVal CONTENT4 As String)
        On Error Resume Next

        Dim wsSendData As String
        wsSendData = ""
        wsSendData = HEADER & "|" & FROMHOST & "|" & TOHOST & "|" & CONTENT1 & "|" & CONTENT2 & "|" & CONTENT3 & "|" & CONTENT4

        SendCommand(wsSendData)
        'Timer6.Enabled = True
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        Select Case m.Msg

            Case WM_USER
                '数値が送信されて来た

            Case WM_COPYDATA
                '文字が送信されて来た
                Dim mystr As COPYDATASTRUCT = New COPYDATASTRUCT()
                Dim mytype As Type = mystr.GetType()
                mystr = CType(m.GetLParam(mytype), COPYDATASTRUCT)

                RxTextBox.Text = mystr.lpData
        End Select
        MyBase.WndProc(m)
    End Sub

    'ラベルコントロール配列のフィールドを作成
    Public LabXBar() As System.Windows.Forms.Label
    Public LabR() As System.Windows.Forms.Label

    'Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Integer)
    'unicodeのEcodingクラスに作成
    Dim encUni As Encoding = Encoding.GetEncoding("utf-16")
    's-jisのEncodingクラスの作成
    Dim encSjis As Encoding = Encoding.GetEncoding("shift-jis")


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If CInt(DateTime.Now.Year) > 2500 Then
            MsgBox("Change the date format to A.D. in control panel." & Environment.NewLine & "( DateTime.Now.Year = " & DateTime.Now.Year & " )")
            End
        End If
        StrCDir = System.IO.Directory.GetCurrentDirectory
        'コンフィグデータを取得する
        ReadConfigData()

        If StrServerConnection = "" Then
            Dim dr As DialogResult
            Dim frm As New Form3
            dr = frm.ShowDialog
            If dr = System.Windows.Forms.DialogResult.OK Then
                LoadLoad()
            ElseIf dr = System.Windows.Forms.DialogResult.Cancel Then
                Me.Close()
            End If

        Else
            LoadLoad()
        End If

    End Sub

    Public Sub LoadLoad()
        'If IO.Directory.Exists(StrRootFolder) = False Then
        '    IO.Directory.CreateDirectory(StrRootFolder)
        'Else
        GetStandardNo()
        GetTreeList_Server()
        'End If
        'SPCリストを取得する

        '日本語・英語表記の切り替えを行う
        FormAlarmInput.Translation_AlarmInput()
        FormAlarmDisp.Translation_AlarmDisp()
        FormControl.Translation_FormControl()

        'X側の軸ラベルを配列化する
        Me.LabXBar = New System.Windows.Forms.Label(10) {}
        Me.LabXBar(0) = Me.Label13
        Me.LabXBar(1) = Me.Label12
        Me.LabXBar(2) = Me.Label3
        Me.LabXBar(3) = Me.Label4
        Me.LabXBar(4) = Me.Label5
        Me.LabXBar(5) = Me.Label6
        Me.LabXBar(6) = Me.Label7
        Me.LabXBar(7) = Me.Label8
        Me.LabXBar(8) = Me.Label9
        Me.LabXBar(9) = Me.Label10
        Me.LabXBar(10) = Me.Label11

        'R側の軸ラベルを配列化する
        Me.LabR = New System.Windows.Forms.Label(9) {}
        Me.LabR(0) = Me.Label36
        Me.LabR(1) = Me.Label30
        Me.LabR(2) = Me.Label29
        Me.LabR(3) = Me.Label20
        Me.LabR(4) = Me.Label19
        Me.LabR(5) = Me.Label14
        Me.LabR(6) = Me.Label15
        Me.LabR(7) = Me.Label16
        Me.LabR(8) = Me.Label17
        Me.LabR(9) = Me.Label18

        Me.Top = 0
        Me.Left = 0


        If StrLanguage = "Japanese" Then
            gType(0) = "上下限"
            gType(1) = "上限のみ"
            gType(2) = "下限のみ"
        ElseIf StrLanguage = "English" Then
            gType(0) = "Upper and Lower"
            gType(1) = "Upper only"
            gType(2) = "Lower only"
        End If


        'ディスプレイの大きさに合わせてフォームサイズの設定
        'ディスプレイの高さ
        Dim h As Integer = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height
        'ディスプレイの幅
        Dim w As Integer = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width

        If w >= 1920 Then
            StrResolution = "MAX"
        ElseIf w >= 1280 And w < 1920 Then
            If h >= 1024 Then
                StrResolution = "Middle"
            Else
                StrResolution = "MIN"
            End If
        ElseIf w <= 1024 Then
            StrResolution = "MIN"
        End If

        SaveConfigData()

        Timer3.Enabled = True 'フォームサイズを画面サイズで変更
    End Sub
#Region "コンフィグデータを取得"
    'コンフィグデータを取得する
    Private Sub ReadConfigData()
        Dim strFileName As String = StrCDir & "\Config.csv"
        Dim temp() As String
        Dim sr As New System.IO.StreamReader(strFileName, System.Text.Encoding.Default)

        'ファイルの最後までループ
        Do Until sr.Peek() = -1
            temp = Split(sr.ReadLine(), ",")
            If temp(0).Trim(Chr(34)) <> "" Then  '
                'SPCデータの保存先フォルダを取得
                If temp(0) = "RootFolder" Then
                    StrRootFolder = temp(1)
                    'SPCデータのネットワーク上の保存先フォルダを取得
                ElseIf temp(0) = "NetworkFolder" Then
                    StrNetworkFolder = temp(1)
                    'SPCの表示開始日時を取得
                ElseIf temp(0).Trim(Chr(34)) = "StartDay" Then
                    strStartDate = temp(1)
                    'SPCのアラーム監視開始日時を取得
                ElseIf temp(0).Trim(Chr(34)) = "AlarmStartDay" Then
                    strAlarmStartDate = temp(1)
                    'Host名を取得
                ElseIf temp(0).Trim(Chr(34)) = "FromHost" Then
                    myHostName = temp(1)
                    'サーバー名を取得
                ElseIf temp(0).Trim(Chr(34)) = "ToHost" Then
                    ServerName = temp(1)
                    '大項目名を取得
                ElseIf temp(0).Trim(Chr(34)) = "MajorItem" Then
                    MajorItem = temp(1)
                    'Host or Subを取得
                ElseIf temp(0).Trim(Chr(34)) = "HostSub" Then
                    HostSub = temp(1)
                    '翻訳言語を取得
                ElseIf temp(0).Trim(Chr(34)) = "Language" Then
                    StrLanguage = temp(1)
                    '解像度を取得
                ElseIf temp(0).Trim(Chr(34)) = "Resolution" Then
                    StrResolution = temp(1)
                    '接続文字列を取得
                ElseIf temp(0).Trim(Chr(34)) = "Connectionstring" Then
                    StrServerConnection = temp(1)
                End If
            End If
        Loop
        sr.Close()     'ファイルを閉じる

    End Sub
#End Region

#Region "コンフィグデータを保存"
    'コンフィグデータを保存する
    Public Sub SaveConfigData()
        Dim str1 As String
        Dim strFileName As String = StrCDir & "\Config.csv"

        Dim sw1 As New System.IO.StreamWriter(strFileName, False, System.Text.Encoding.Default)  '上書き
        str1 = Err.Description
        If str1 = "" Then
            'SPCデータの保存先フォルダ名を保存
            sw1.WriteLine("RootFolder," & StrRootFolder)
            'SPCデータのネットワーク上の保存先フォルダ名を保存
            sw1.WriteLine("NetworkFolder," & StrNetworkFolder)
            '現在の表示開始日時を保存
            sw1.WriteLine("StartDay," & strStartDate)
            '現在のアラーム監視開始日時を保存
            sw1.WriteLine("AlarmStartDay," & strAlarmStartDate)
            'Host名を保存
            sw1.WriteLine("FromHost," & myHostName)
            'サーバー名を保存
            sw1.WriteLine("ToHost," & ServerName)
            '大項目名を保存
            sw1.WriteLine("MajorItem," & MajorItem)
            'Host or Subを保存
            sw1.WriteLine("HostSub," & HostSub)
            '翻訳言語を保存
            sw1.WriteLine("Language," & StrLanguage)
            '解像度を保存
            sw1.WriteLine("Resolution," & StrResolution)
            '接続文字列を保存
            sw1.WriteLine("Connectionstring," & StrServerConnection)
            sw1.Close()  '
        End If

    End Sub
#End Region

#Region "ツリー作成"


    'サーバーよりツリー項目一覧を取得する
    Public Sub GetTreeList_Server()

        Dim TreeRist(,) As String
        TreeRist = getTreeData()
        If TreeRist Is Nothing Then
            Exit Sub
        End If

        TreeDisp_Server_New(TreeRist, TreeView1)
        TreeDisp_Server_New(TreeRist, FormMiddle.TreeView1)
        TreeDisp_Server_New(TreeRist, FormSmall.TreeView1)
    End Sub



    Private Sub TreeDisp_Server_New(ByVal _TreeRist(,) As String, ByVal _TreeView As TreeView)


        _TreeView.Nodes.Clear()

        Dim temp() As String
        temp = Split(_TreeRist(0, 0), ",")

        Dim maxLen As Integer = 0
        For i As Integer = 0 To UBound(_TreeRist, 1)
            If _TreeRist(i, 0) IsNot Nothing Then
                Dim parts() As String = Split(_TreeRist(i, 0), ",")
                If parts.Length > maxLen Then maxLen = parts.Length
            End If
        Next
        Dim TRist(UBound(_TreeRist, 1), maxLen) As String

        Dim Icons As New ImageList
        Icons.Images.Add("Yellow", Image.FromFile(StrCDir & "\Picture\Folder.jpg")) 'SPCアラームがない場合のフォルダアイコン
        Icons.Images.Add("Red", Image.FromFile(StrCDir & "\Picture\FolderRed.jpg")) 'SPCアラームがある場合のフォルダアイコン(赤)
        Icons.Images.Add("Blue", Image.FromFile(StrCDir & "\Picture\FolderBlue.jpg")) 'QC承認が入っていない場合のフォルダアイコン(青)

        _TreeView.ImageList = Icons 'アイコンイメージ

        Dim Tree(UBound(temp, 1))() As String
        Dim RootNode(UBound(temp, 1))() As TreeNode

        For i As Integer = 0 To UBound(_TreeRist, 1)
            temp = Split(_TreeRist(i, 0), ",")
            For j As Integer = 0 To UBound(temp, 1)
                TRist(i, j) = temp(j)
            Next
            TRist(i, UBound(TRist, 2)) = _TreeRist(i, 1)
        Next


        Dim Buf(UBound(TRist, 1)) As String
        Dim Buf2() As String
        Dim _al As System.Collections.ArrayList
        Dim _j As String

        For i As Integer = 0 To UBound(temp, 1)
            For j As Integer = 0 To UBound(Buf, 1)
                Buf(j) = ""
                For k As Integer = 0 To i
                    Buf(j) &= TRist(j, k)
                    If Not k = i Then
                        Buf(j) &= ","
                    End If
                Next

            Next

            _al = New System.Collections.ArrayList(UBound(Buf, 1))

            '重複削除
            For Each _j In Buf
                If Not _al.Contains(_j) And Not _j = "" Then
                    _al.Add(_j)
                End If
            Next

            Buf2 = DirectCast(_al.ToArray(GetType(String)), String())

            ReDim Tree(i)(UBound(Buf2, 1))
            ReDim RootNode(i)(UBound(Buf2, 1))

            For j As Integer = 0 To UBound(Buf2, 1)
                Tree(i)(j) = Buf2(j)
            Next


        Next


        Dim Kfile As Integer = 0
        Dim OyaNodeBuf As String = ""
        For i As Integer = 0 To UBound(Tree, 1)
            For j As Integer = 0 To UBound(Tree(i), 1)
                For k As Integer = 0 To UBound(_TreeRist, 1) 'アラームありか確認
                    If InStr(_TreeRist(k, 0), Tree(i)(j)) Then
                        If _TreeRist(k, 1) = "1" Then
                            Kfile = 1
                        End If
                    End If
                Next
                temp = Split(Tree(i)(j), ",")
                If Not i = 0 Then
                    For k As Integer = 0 To UBound(temp, 1) - 1
                        OyaNodeBuf &= temp(k)
                        If Not k = UBound(temp, 1) - 1 Then
                            OyaNodeBuf &= ","
                        End If
                    Next
                End If

                If Not temp(UBound(temp, 1)) = "" Then
                    RootNode(i)(j) = New TreeNode(temp(UBound(temp, 1)), Kfile, Kfile)
                    If i = 0 Then
                        _TreeView.Nodes.Add(RootNode(i)(j))
                    Else
                        For k As Integer = 0 To UBound(Tree(i - 1), 1)
                            If Tree(i - 1)(k) = OyaNodeBuf Then
                                RootNode(i - 1)(k).Nodes.Add(RootNode(i)(j))
                            End If
                        Next
                    End If
                End If


                Kfile = 0
                OyaNodeBuf = ""
            Next
        Next


    End Sub

#End Region


    'サーバーよりユーザーリストを取得する
    Public Function GetUserList() As Boolean
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim table As New DataTable
        Dim n As Integer
        Try
            GetUserList = False

            Cn.ConnectionString = StrServerConnection
            table.Clear()

            strSQL = "SELECT *"
            strSQL &= " FROM SPC_User"
            strSQL &= " ORDER BY cUserID "

            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)
            n = table.Rows.Count

            Adapter.Dispose()
            Cn.Dispose()
            table.Dispose()

            If n > 0 Then
                GetUserList = True
                Form_User.ComboBox1.Items.Clear()
                For i = 0 To n - 1
                    If Not table.Rows(i)("cUserID") = "Stan_No" Then
                        Form_User.ComboBox1.Items.Add(table.Rows(i)("cUserID"))
                    End If

                Next
            End If

        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()
            table.Dispose()
            GetUserList = False
            StrErrMes = "ユーザーリスト取得エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Function
        End Try
    End Function

    'パスワードの照合を実施する
    Public Function Check_UserPass(ByVal _User As String, ByVal _Password As String, ByVal QC As Integer, ByVal DisOK As Integer) As String
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim Adapter As New SqlDataAdapter
        Dim strSQL As String = ""
        Dim table As New DataTable
        Dim n As Integer
        Dim msg As String = ""
        Try
            Check_UserPass = ""

            Cn.ConnectionString = StrServerConnection
            table.Clear()

            strSQL = "SELECT *"
            strSQL &= " FROM SPC_User"
            strSQL &= " WHERE cUserID = '" & _User & "' and  cPassword = '" & _Password & "'"

            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)
            n = table.Rows.Count

            Adapter.Dispose()
            Cn.Dispose()
            table.Dispose()
            'MsgBox(table.Rows(0)("bApprover"))
            If n = 1 Then '該当するのは一人のはず
                If QC = 1 Then
                    If table.Rows(0)("bApprover") = True Then
                        Check_UserPass = table.Rows(0)("cUserID")
                    Else
                        msg = "入力されたIDは、QC承認者ではありません。"
                    End If
                Else
                    Check_UserPass = table.Rows(0)("cUserID")
                End If
            Else
                msg = "IDまたはPasswordが間違っています。" & Environment.NewLine & "Password verification NG"
            End If


            If _User = "Stan_No" Then
                msg = "IDまたはPasswordが間違っています。" & Environment.NewLine & "Password verification NG"
                Check_UserPass = ""
            End If

            If Not msg = "" Then
                MsgBox(msg)
            Else
                If DisOK = 1 Then
                    If StrLanguage = "Japanese" Then
                        MsgBox("パスワード照合OK")
                    ElseIf StrLanguage = "English" Then
                        MsgBox("Password verification OK")
                    End If
                End If
            End If

        Catch ex As System.Exception
            Adapter.Dispose()
            Cn.Dispose()
            table.Dispose()
            Check_UserPass = ""
            StrErrMes = "パスワード照合エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Function
        End Try
    End Function
    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect

        TreeView_AfterSelect(TreeView1)
    End Sub


    Public Sub TreeView_AfterSelect(ByVal _TreeView As TreeView)

        'ツリーデータセット
        TreeName = treeInfo(_TreeView)

        'ツリーデータからプロパティデータセット(QC承認も確認)
        PropertyTable = getProperty()

        If PropertyTable Is Nothing Then
            Exit Sub
        End If

        'データ取得とグラフ用データ準備
        M_Data = getSPCMaster()
        If M_Data Is Nothing Then
            Exit Sub
        End If


        'SPCアラーム判定を行う
        GetAlarmData_kai()

        M_Alarm = getAlarmMaster()

        If M_Alarm Is Nothing Then
            Exit Sub
        End If

        MRFlag = False
        If StrResolution = "MAX" Then
            Me.Label1.Text = "R"
            Me.GroupBox2.Text = "R"
        ElseIf StrResolution = "MIN" Then
            FormMiddle.Label1.Text = "R"
            FormMiddle.GroupBox2.Text = "R"
        ElseIf StrResolution = "Middle" Then
            FormSmall.Label1.Text = "R"
            FormSmall.GroupBox2.Text = "R"
        End If


        If QCNotCheckFlag = False Then 'QC未承認の場合
            If HostSub = "Sub" Then
                'If AllPropertyNo <> 0 Then
                If PropertyTable.Rows.Count <> 0 Then

                    If StrResolution = "MAX" Then
                        Me.LabelQC.Visible = True
                    ElseIf StrResolution = "MIN" Then
                        FormMiddle.LabelQC.Visible = True
                    ElseIf StrResolution = "Middle" Then
                        FormSmall.LabelQC.Visible = True
                    End If
                    If StrLanguage = "Japanese" Then

                        If StrResolution = "MAX" Then
                            Me.LabelQC.Text = "QC未承認"
                        ElseIf StrResolution = "MIN" Then
                            FormMiddle.LabelQC.Text = "QC未承認"
                        ElseIf StrResolution = "Middle" Then
                            FormSmall.LabelQC.Text = "QC未承認"
                        End If

                        MsgBox("QC承認待ちです。")
                    ElseIf StrLanguage = "English" Then

                        If StrResolution = "MAX" Then
                            Me.LabelQC.Text = "QC not approved"
                        ElseIf StrResolution = "MIN" Then
                            FormMiddle.LabelQC.Text = "QC not approved"
                        ElseIf StrResolution = "Middle" Then
                            FormSmall.LabelQC.Text = "QC not approved"
                        End If

                        MsgBox("Waiting for QC approval")
                    End If

                End If
            End If
        Else

            If StrResolution = "MAX" Then
                Me.LabelQC.Visible = False
            ElseIf StrResolution = "MIN" Then
                FormMiddle.LabelQC.Visible = False
            ElseIf StrResolution = "Middle" Then
                FormSmall.LabelQC.Visible = False
            End If

        End If

        'GraphDisp_Server() 'グラフを表示する
        GraphDisp() 'グラフを表示する
    End Sub


    Private Sub Button2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button2.MouseDown
        'グラフ左移動START
        Timer1.Enabled = True
    End Sub

    Private Sub Button3_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button3.MouseDown
        'グラフ右移動START
        Timer2.Enabled = True
    End Sub
    'グラフ左移動
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim dt As Date

        dt = Now()
        If DispStartPosition > 0 Then
            DispStartPosition -= 1

            GraphDisp()


        End If

    End Sub
    'グラフ右移動
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Dim dt As Date
        dt = Now()

        If SPCDataNum - 30 > DispStartPosition Then
            DispStartPosition += 1

            GraphDisp()

        End If

    End Sub

    Private Sub Button2_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button2.MouseUp
        'グラフ左移動STOP
        Timer1.Enabled = False
    End Sub

    Private Sub Button3_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Button3.MouseUp
        'グラフ右移動STOP
        Timer2.Enabled = False
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





    Public Sub UPDATE_d_a_StartDate(ByVal dDay As Date, ByVal aDay As Date, ByVal ALL As Integer)

        Dim Cn As New SqlConnection
        Dim strSQL As String
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim trans As SqlTransaction 'トランザクション定義
        Dim temp() As String
        Dim Updateflag As Boolean = False

        Try

            Cn.ConnectionString = StrServerConnection

            Cn.Open()
            trans = Cn.BeginTransaction
            SQLCm.Transaction = trans


            strSQL = ""
            strSQL = "UPDATE SPC_Property SET "

            strSQL &= " dStartDate= '" & Format(dDay, "yyyy-MM-dd 00:00:00.000") & "'"
            strSQL &= ","
            strSQL &= " aStartDate= '" & Format(aDay, "yyyy-MM-dd 00:00:00.000") & "'"



            If ALL = 0 Then

                For i As Integer = 0 To UBound(TreeName, 1)
                    If i = 0 Then
                        strSQL &= " WHERE"
                    Else
                        strSQL &= " AND"
                    End If

                    strSQL &= " cTreeName" & i + 1 & " = '" & TreeName(i) & "'"
                Next
            End If

            SQLCm.CommandText = strSQL
            SQLCm.ExecuteNonQuery()

            trans.Commit()
            Cn.Close()

            GetTreeList_Server()

            FormProperty.Close()
            MsgBox("Change OK")

        Catch ex As Exception
            If IsNothing(trans) = False Then
                trans.Rollback()
            End If
            MsgBox("Change NG")
            StrErrMes = "アラームコメントアップデートエラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try

    End Sub



    Public Function Check_Textdata() As Boolean

        Check_Textdata = True

        Dim eMas As String = ""

        If IsNumeric(FormSPCRule.TextX_CL_After.Text) = False Then
            eMas &= " ・No value entered for XBar_CL." & Environment.NewLine
        End If
        If IsNumeric(FormSPCRule.TextX_UCL_After.Text) = False Then
            eMas &= " ・No value entered for XBar_UCL." & Environment.NewLine
        End If
        If IsNumeric(FormSPCRule.TextX_LCL_After.Text) = False Then
            eMas &= " ・No value entered for XBar_LCL." & Environment.NewLine
        End If
        If IsNumeric(FormSPCRule.TextX_S_After.Text) = False Then
            eMas &= " ・No value entered for XBar_σ." & Environment.NewLine
        End If

        If IsNumeric(FormSPCRule.TextR_CL_After.Text) = False Then
            eMas &= " ・No value entered for R_CL." & Environment.NewLine
        End If
        If IsNumeric(FormSPCRule.TextR_UCL_After.Text) = False Then
            eMas &= " ・No value entered for R_UCL." & Environment.NewLine
        End If
        If IsNumeric(FormSPCRule.TextR_LCL_After.Text) = False Then
            eMas &= " ・No value entered for R_LCL." & Environment.NewLine
        End If
        If IsNumeric(FormSPCRule.TextR_S_After.Text) = False Then
            eMas &= " ・No value entered for R_σ." & Environment.NewLine
        End If

        If IsNumeric(FormSPCRule.TextMR_CL_After.Text) = False Then
            eMas &= " ・No value entered for MR_CL." & Environment.NewLine
        End If
        If IsNumeric(FormSPCRule.TextMR_UCL_After.Text) = False Then
            eMas &= " ・No value entered for MR_UCL." & Environment.NewLine
        End If
        If IsNumeric(FormSPCRule.TextMR_LCL_After.Text) = False Then
            eMas &= " ・No value entered for MR_LCL." & Environment.NewLine
        End If
        If IsNumeric(FormSPCRule.TextMR_S_After.Text) = False Then
            eMas &= " ・No value entered for MR_σ." & Environment.NewLine
        End If

        If Not eMas = "" Then
            Check_Textdata = False
            MsgBox("<Error>" & Environment.NewLine & eMas)
        End If


    End Function
    '管理値を変更する
    Public Sub UpdateControlLine()

        Dim strReason As String = ""
        Dim Buf(0) As String

        If StrLanguage = "Japanese" Then '行橋のみ
            If StrNetworkFolder = "L:\" Then
                MsgBox("この設定ではSPCルールを変更することは出来ません。")
                Exit Sub
            End If
        End If


        UserName = ""
        JP_Message = "SPCルールを変更します。管理者パスワードを入力して下さい。"
        EN_Message = "Change SPC rules. Enter administrator password."
        UserName = FormAlarmInput.Input_Pass_to_Get_UserName(JP_Message, EN_Message, 0, 0) '0:IDPass合ってれば名前取得　1:QC承認者であるかも判定 

        If UserName <> "" Then

            If StrLanguage = "Japanese" Then
                strReason = InputBox("パスワード照合OK。変更理由を記入してください。")
            ElseIf StrLanguage = "English" Then
                strReason = InputBox("Password verification OK. Please enter the reason for the change")
            End If

            If strReason = "" Then
                If StrLanguage = "Japanese" Then
                    MsgBox("変更理由が記入されていません。もう一度やり直してください。")
                ElseIf StrLanguage = "English" Then
                    MsgBox("The reason for the change has not been entered. Please try again.")
                End If
                Exit Sub
            End If

            Input_Chenged_Property(Buf, strReason, "Control limit change", PropertyTable.Rows.Count - 1)
            PropertyTable = getProperty()
        End If



    End Sub
    Public Sub Input_Chenged_Property(ByRef _Buf() As String, ByRef Reason As String, ByVal _Mode As String, ByVal _p As String)
        ReDim _Buf(PropertyTable.Columns.Count - 1)

        '最新のPropertyのデータ取得
        For i As Integer = 0 To UBound(_Buf, 1)
            If IsDBNull(PropertyTable.Rows(_p)(i)) Then
                _Buf(i) = "Null"
            Else
                _Buf(i) = PropertyTable.Rows(_p)(i)
            End If
        Next


        If _Mode = "Control limit change" Then


            _Buf(0) = FormControl.Get_PropertyNo() 'iGraphNo
            'Buf(1) = "" 'cProcessName
            'Buf(2) = "" 'cMachineNo
            'Buf(3) = "" 'cControlItem
            'Buf(4) = "" 'cDeviceName
            'Buf(5) = "" 'cFilter_1
            'Buf(6) = "" 'cFilter_2
            'Buf(7) = "" 'cFilter_3
            'Buf(8) = "" 'cFilter_4
            'Buf(9) = "" 'cFilter_5
            'Buf(10) = "" 'cFilter_6
            'Buf(11) = "" 'cFilter_7
            'Buf(12) = "" 'cFilter_8
            'Buf(13) = "" 'cFilter_9
            'Buf(14) = "" 'cFilter_10
            'Buf(15) = "" 'cTreeName1
            'Buf(16) = "" 'cTreeName2
            'Buf(17) = "" 'cTreeName3
            'Buf(18) = "" 'cTreeName4
            'Buf(19) = "" 'cTreeName5
            'Buf(20) = "" 'cTreeName6
            'Buf(21) = "" 'cTreeName7
            'Buf(22) = "" 'cTreeName8
            'Buf(23) = "" 'cTreeName9
            'Buf(24) = "" 'cTreeName10
            'Buf(25) = "" 'cUnit
            'Buf(26) = "" 'cLimitType
            'Buf(27) = "" 'cScl
            'Buf(28) = "" 'cTolerance
            'Buf(29) = "" 'cUsl
            'Buf(30) = "" 'cLsl

            _Buf(31) = FormSPCRule.TextX_CL_After.Text 'cXcl
            _Buf(32) = FormSPCRule.TextX_S_After.Text 'cXdev
            _Buf(33) = FormSPCRule.TextX_UCL_After.Text 'cXucl
            _Buf(34) = FormSPCRule.TextX_LCL_After.Text 'cXlcl
            _Buf(35) = FormSPCRule.TextR_CL_After.Text 'cRcl
            _Buf(36) = FormSPCRule.TextR_S_After.Text 'cRdev
            _Buf(37) = FormSPCRule.TextR_UCL_After.Text 'cRucl
            _Buf(38) = FormSPCRule.TextR_LCL_After.Text 'cRlcl
            _Buf(39) = FormSPCRule.TextMR_CL_After.Text 'cMRcl
            _Buf(40) = FormSPCRule.TextMR_S_After.Text 'cMRdev
            _Buf(41) = FormSPCRule.TextMR_UCL_After.Text 'cMRucl
            _Buf(42) = FormSPCRule.TextMR_LCL_After.Text 'cMRlcl

            If FormSPCRule.CheckBox9.Checked = True Then
                _Buf(43) = "1" 'cMR
            Else
                _Buf(43) = "0" 'cMR
            End If
            If FormSPCRule.CheckBox1.Checked = True Then
                _Buf(44) = "1" 'cSpcrule1
            Else
                _Buf(44) = "0" 'cSpcrule1
            End If
            If FormSPCRule.CheckBox2.Checked = True Then
                _Buf(45) = "1" 'cSpcrule2
            Else
                _Buf(45) = "0" 'cSpcrule2
            End If
            If FormSPCRule.CheckBox3.Checked = True Then
                _Buf(46) = "1" 'cSpcrule3
            Else
                _Buf(46) = "0" 'cSpcrule3
            End If
            If FormSPCRule.CheckBox4.Checked = True Then
                _Buf(47) = "1" 'cSpcrule4
            Else
                _Buf(47) = "0" 'cSpcrule4
            End If
            If FormSPCRule.CheckBox5.Checked = True Then
                _Buf(48) = "1" 'cSpcrule5
            Else
                _Buf(48) = "0" 'cSpcrule5
            End If
            If FormSPCRule.CheckBox6.Checked = True Then
                _Buf(49) = "1" 'cSpcrule6
            Else
                _Buf(49) = "0" 'cSpcrule6
            End If
            If FormSPCRule.CheckBox7.Checked = True Then
                _Buf(50) = "1" 'cSpcrule7
            Else
                _Buf(50) = "0" 'cSpcrule7
            End If
            If FormSPCRule.CheckBox8.Checked = True Then
                _Buf(51) = "1" 'cSpcrule8
            Else
                _Buf(51) = "0" 'cSpcrule8
            End If



            _Buf(52) = DateTime.Now.ToString 'cUpdateDate
            _Buf(53) = "Control limit change" 'cUpdateContent
            _Buf(54) = Reason 'cUpdateReason
            _Buf(55) = UserName 'cIncharge
            _Buf(56) = "Null" 'cApprovalDate
            _Buf(57) = "Null" 'cApproverName

            _Buf(58) = DateTime.Parse(_Buf(58)).ToString 'dStartDate
            _Buf(59) = DateTime.Parse(_Buf(59)).ToString 'aStartDate

            To_Server(_Buf, _Mode)

        ElseIf _Mode = "Control limit approval" Then


            _Buf(56) = DateTime.Now.ToString 'cApprovalDate
            _Buf(57) = UserName  'cApproverName

            To_Server(_Buf, _Mode)

        End If

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        If StrServerConnection = "" Then
            Dim dr As DialogResult
            Dim frm As New Form3
            dr = frm.ShowDialog
            If dr = System.Windows.Forms.DialogResult.OK Then
                LoadLoad()
            ElseIf dr = System.Windows.Forms.DialogResult.Cancel Then
                Me.Close()
            End If

        Else
            LoadLoad()
        End If
        MsgBox("Tree Updated")
    End Sub
    'SPCデータの自動更新を行う
    Private Sub UpdateTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateTimer.Tick
        'UpdateTimer.Enabled = False
        'System.IO.File.Copy("R:\" & UpdateFileName, "C:\" & UpdateFileName, True) 'Rドライブからrootフォルダへコピー
        'System.IO.File.Copy("R:\" & UpdatePropertyName, "C:\" & UpdatePropertyName, True) 'Rドライブからrootフォルダへコピー
        'System.IO.File.Copy("R:\SPCData\Alarm.csv", "C:\SPCData\Alarm.csv", True) 'Rドライブからrootフォルダへコピー
        'Me.TreeView1.Nodes.Clear()
        'GetHostItemList() 'ツリービュー再描画
    End Sub



    Private Sub UpdateTimerHost_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateTimerHost.Tick
        'UpdateTimerHost.Enabled = False
        'If System.IO.File.Exists("R:\" & UpdateFileName) Then
        '    System.IO.File.Copy("R:\" & UpdateFileName, "C:\" & UpdateFileName, True) 'Rドライブからrootフォルダへコピー
        '    System.IO.File.Copy("R:\" & UpdatePropertyName, "C:\" & UpdatePropertyName, True) 'Rドライブからrootフォルダへコピー

        'End If

        'System.IO.File.Copy("R:\SPCData\Alarm.csv", "C:\SPCData\Alarm.csv", True) 'Rドライブからrootフォルダへコピー
    End Sub


    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

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


    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click

    End Sub

    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click

    End Sub

    Private Sub ExToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExToolStripMenuItem.Click
        'Me.SPCData_Export() '表示中のSPCデータを出力する
    End Sub

    Private Sub DToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DToolStripMenuItem.Click
        FormProperty.Show() 'プロパティ画面を表示する
    End Sub

    Private Sub CreateNewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreateNewToolStripMenuItem.Click
        FormControl.Show() 'グラフ初期設定画面を表示する
    End Sub

    Private Sub SPCRuleToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SPCRuleToolStripMenuItem1.Click
        FormSPCRule.Show() 'SPC詳細設定画面を表示する
        FormSPCRule.Button2_Click(Nothing, Nothing)
    End Sub

    Private Sub JapaneseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JapaneseToolStripMenuItem.Click
        StrLanguage = "Japanese"
        SaveConfigData()
        FormAlarmInput.Translation_AlarmInput()
        FormAlarmDisp.Translation_AlarmDisp()
        FormControl.Translation_FormControl()
        MsgBox("表示言語を日本語に変更しました")
    End Sub

    Private Sub EnglishToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnglishToolStripMenuItem.Click
        StrLanguage = "English"
        SaveConfigData()
        FormAlarmInput.Translation_AlarmInput()
        FormAlarmDisp.Translation_AlarmDisp()
        FormControl.Translation_FormControl()
        MsgBox("Changed display language to English")
    End Sub

    Private Sub HowToUseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HowToUseToolStripMenuItem.Click
        FormHowto.Show() '使用マニュアルを表示する
    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        StrResolution = "MIN"
        SaveConfigData()
        FormSmall.Show()
        Me.Hide()
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        StrResolution = "MAX"
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click

    End Sub

    Private Sub FontDialog1_Apply(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontDialog1.Apply

    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        Timer3.Enabled = False
        If StrResolution = "MIN" Then
            FormSmall.Show()
            Me.Hide()
        ElseIf StrResolution = "Middle" Then
            FormMiddle.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        StrResolution = "Middle"
        SaveConfigData()
        FormMiddle.Show()
        Me.Hide()
    End Sub

    Private Sub DataConverterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataConverterToolStripMenuItem.Click
        'Dim proc As New Process()                 ' （1）
        'proc.StartInfo.FileName = "D:\SpcDataConverter\SPCDataConverter_Yuk\SPCDataConverter_Yuk\bin\X86\Release\SPCDataConverter_Yuk.exe"    ' （2）
        'proc.Start()                              ' （3）
    End Sub


    Private Sub Button10_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim proc As New Process()                 ' （1）
        proc.StartInfo.FileName = "C:\ProductMonitor_Converter\ProductMonitor_Converter\bin\X86\Release\ProductMonitor_Converter.exe"    ' （2）
        proc.Start()
    End Sub

    Private Sub UserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserToolStripMenuItem.Click
        User()
    End Sub

    Public Sub New()

        ' この呼び出しは、Windows フォーム デザイナで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。

    End Sub

    Public Sub To_Server(ByVal Array() As String, ByVal Mode As String)
        Dim Cn As New SqlConnection
        Dim strSQL As String
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim trans As SqlTransaction 'トランザクション定義

        Try

            Cn.ConnectionString = StrServerConnection

            Cn.Open()
            trans = Cn.BeginTransaction
            SQLCm.Transaction = trans


            '規格変更は1行増える。変更後の承認は、既存の行に承認者の名前が記入される。
            If Mode = "Control limit approval" Then
                SQLCm.CommandText = "DELETE FROM SPC_Property WHERE iGraphNo = '" & Array(0) & "'"
                SQLCm.ExecuteNonQuery()
            End If



            strSQL = ""
            strSQL = "INSERT INTO SPC_Property VALUES ("

            For i As Integer = 0 To UBound(Array, 1)
                If Array(i) = "Null" Then
                    strSQL &= "Null"
                Else
                    strSQL &= "'" & Array(i) & "'"
                End If
                If i = UBound(Array, 1) Then
                    strSQL &= ")"
                Else
                    strSQL &= ","
                End If
            Next

            SQLCm.CommandText = strSQL
            SQLCm.ExecuteNonQuery()


            trans.Commit()
            Cn.Close()

        Catch ex As Exception
            If IsNothing(trans) = False Then
                trans.Rollback()
            End If
            StrErrMes = "SPCアラームデータ更新エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try

    End Sub

    Private Sub ServerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ServerToolStripMenuItem.Click
        Dim dr As DialogResult
        Dim frm As New Form3
        dr = frm.ShowDialog
    End Sub

    Private Sub DeleteAlarmTableToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteAlarmTableToolStripMenuItem.Click
        Delete_AlarmTable()
    End Sub
    Public Sub Delete_AlarmTable()
        Dim result As DialogResult
        Dim eMsg As String = ""
        result = MessageBox.Show("Do you want to delete the data in the SPC_Alarm ?", "Delete Data", MessageBoxButtons.OKCancel)
        If result = System.Windows.Forms.DialogResult.OK Then
            eMsg = Input_CSV_to_Server(StrCDir & "\SPC_Alarm.csv", 1)
            If Not eMsg = "" Then
                Write_Error("error Delete_AlarmTable" & eMsg)
                MsgBox("Error: Deletion failed.")
                Exit Sub
            End If
            LoadLoad()
            MsgBox("Deleted the data in the SPC_Alarm.")
        End If
    End Sub
    Public Function Input_CSV_to_Server(ByVal CsvName As String, ByVal Del As Integer) As String

        Input_CSV_to_Server = ""

        Dim CSVData(0, 0) As String
        Dim CSVData_Header(0, 0) As String
        Dim TableName As String = ""
        Input_CSV_to_Server = Input_CSV(CsvName, CSVData, CSVData_Header, TableName)
        If Not Input_CSV_to_Server = "" Then
            Exit Function
        End If
        Input_CSV_to_Server = To_Server0(CSVData, CSVData_Header, TableName, Del)
        If Not Input_CSV_to_Server = "" Then
            Exit Function
        End If
    End Function
    Public Function Input_CSV(ByVal Filename As String, ByRef Array(,) As String, ByRef Array_Header(,) As String, ByRef TName As String) As String 'CSVデータを配列化(1行目のヘッダーは除く)
        Input_CSV = ""
        Dim eCode As String = ""
        Try
            Dim Buf(,) As String
            Dim temp0() As String
            Dim sr0 As System.IO.StreamReader
            Dim CoMax As Integer = 0
            Dim Gyo As Integer = 0
            '-------------------------------------------------------------------------CSVの行列大きさ確認----------------------------------------------------------
            eCode = "a"
            sr0 = New System.IO.StreamReader(Filename, System.Text.Encoding.Default)
            Do Until sr0.Peek() = -1
                temp0 = Split(sr0.ReadLine(), ",")
                If CoMax < temp0.Length Then
                    CoMax = temp0.Length
                End If
                Gyo += 1
            Loop
            sr0.Close()     'ファイルを閉じる
            '-------------------------------------------------------------------------CSVの行列大きさ確認----------------------------------------------------------
            eCode = "b"
            '-------------------------------------------------------------------------その大きさ分定義してBufに格納-------------------------------------------------------------------------
            ReDim Buf(Gyo - 1, CoMax - 1)

            Gyo = 0
            sr0 = New System.IO.StreamReader(Filename, System.Text.Encoding.Default)
            'temp0 = Split(sr0.ReadLine(), ",")
            'For i As Integer = 0 To Array_Header.Length - 1
            '    Array_Header(i) = temp0(i)
            'Next
            Do Until sr0.Peek() = -1
                temp0 = Split(sr0.ReadLine(), ",")
                For i As Integer = 0 To CoMax - 1
                    If i < temp0.Length Then
                        Buf(Gyo, i) = temp0(i)
                    Else
                        Buf(Gyo, i) = ""
                    End If
                Next
                Gyo += 1
            Loop
            sr0.Close()     'ファイルを閉じる
            '-------------------------------------------------------------------------その大きさ分定義してBufに格納-------------------------------------------------------------------------
            eCode = "c"
            ReDim Array_Header(3 - 1, UBound(Buf, 2) - 1) '列名、データ型、
            ReDim Array(UBound(Buf, 1) - (UBound(Array_Header, 1) + 1) - 1, UBound(Buf, 2) - 1)
            TName = Buf(0, 1)
            For i As Integer = 1 To 1 + UBound(Array_Header, 1)
                For j As Integer = 1 To UBound(Buf, 2)
                    Array_Header(i - 1, j - 1) = Buf(i, j)
                Next
            Next
            eCode = "d"
            For i As Integer = 1 + UBound(Array_Header, 1) + 1 To UBound(Buf, 1)
                For j As Integer = 1 To UBound(Buf, 2)
                    Array(i - (1 + UBound(Array_Header, 1) + 1), j - 1) = Buf(i, j)
                Next
            Next
            eCode = "e"
        Catch ex As Exception
            Input_CSV = "Input_CSV " & eCode & Environment.NewLine & ex.Message & Environment.NewLine & ex.StackTrace
        End Try



    End Function
    Public Function To_Server0(ByVal Array(,) As String, ByVal Array_Header(,) As String, ByVal TName As String, ByVal Del As Integer) As String

        To_Server0 = ""
        Dim eCode As String = ""
        Dim Cn As New SqlConnection
        Dim strSQL As String = ""
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim trans As SqlTransaction = Nothing  'トランザクション定義

        Try

            eCode = "a"

            Cn.ConnectionString = StrServerConnection

            Cn.Open()
            trans = Cn.BeginTransaction
            SQLCm.Transaction = trans

            If 1 Then

                eCode = "b"

                If Del = 0 Then '0:消さずにスルー、1:消して新規作成
                    MsgBox("Table:" & TName & " already exists.")
                    Cn.Close()
                    Exit Function
                End If
                eCode = "c"
                SQLCm.CommandText = "DROP TABLE " & TName '消す
                SQLCm.ExecuteNonQuery()
                eCode = "d"
                strSQL = "Create Table " & TName & "("
                For j As Integer = 0 To UBound(Array_Header, 2)
                    strSQL &= Array_Header(0, j) & " " & Array_Header(1, j)
                    If Array_Header(2, j) = "NO" Then
                        strSQL &= " Not Null"
                    End If
                    If j = UBound(Array_Header, 2) Then
                        'strSQL &= ",PRIMARY KEY (" & Array_Header(0, 0) & ")"
                        strSQL &= ")"
                    Else
                        strSQL &= ","
                    End If
                Next
                eCode = "e"
                SQLCm.CommandText = strSQL
                SQLCm.ExecuteNonQuery()

                eCode = "f"
                SQLCm.CommandText = "DELETE FROM " & TName
                SQLCm.ExecuteNonQuery()
                eCode = "g"
                For i As Integer = 0 To UBound(Array, 1)

                    strSQL = ""
                    strSQL = "INSERT INTO " & TName & " VALUES ("

                    For j As Integer = 0 To UBound(Array, 2)
                        If Array(i, j) = "Null" Then
                            strSQL &= "Null"
                        Else
                            strSQL &= "'" & Array(i, j) & "'"
                        End If
                        If j = UBound(Array, 2) Then
                            strSQL &= ")"
                        Else
                            strSQL &= ","
                        End If
                        eCode = "h " & i & " " & j
                    Next

                    SQLCm.CommandText = strSQL
                    SQLCm.ExecuteNonQuery()

                Next

            End If

            trans.Commit()
            Cn.Close()

        Catch ex As Exception
            If IsNothing(trans) = False Then
                trans.Rollback()
            End If
            To_Server0 = "To_Server0 " & eCode & Environment.NewLine & strSQL & Environment.NewLine & ex.Message & Environment.NewLine & ex.StackTrace
        End Try

    End Function
    Private Sub Write_Error(ByVal strE As String)
        Dim TextFile As IO.StreamWriter = New IO.StreamWriter(StrCDir & "\Emsg.txt", True, System.Text.Encoding.Default)
        TextFile.Write(strE)
        TextFile.WriteLine()
        TextFile.Close()
    End Sub

    Private Sub StandardNoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StandardNoToolStripMenuItem.Click
        StandardNo()
    End Sub
    Public Sub StandardNo()
        Form_User.TabControl1.SelectedIndex = 2
        Form_User.Show()
    End Sub
    Public Sub User()
        Form_User.TabControl1.SelectedIndex = 0
        Form_User.Show()
    End Sub
    Public Sub GetStandardNo()

        txtStan_No.Text = "No Data"
        FormMiddle.txtStan_No.Text = "No Data"
        FormSmall.txtStan_No.Text = "No Data"
        Dim Cn As New SqlConnection
        Dim strSQL As String
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim Adapter As New SqlDataAdapter
        Dim table As New DataTable
        Dim n As Integer
        Try

            Cn.ConnectionString = StrServerConnection

            strSQL = "SELECT *"
            strSQL &= " FROM SPC_User"
            strSQL &= " WHERE cUserID = '" & "Stan_No" & "'"

            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)
            n = table.Rows.Count
            Adapter.Dispose()

            If Not n = 0 Then
                txtStan_No.Text = table.Rows(0)("cPassword")
                FormMiddle.txtStan_No.Text = table.Rows(0)("cPassword")
                FormSmall.txtStan_No.Text = table.Rows(0)("cPassword")
            End If





        Catch ex As Exception
            StrErrMes = "ユーザー情報更新エラー" + ", " + ex.Message & ex.StackTrace
            Call SaveLog(Now(), StrErrMes)
            Exit Sub
        End Try

    End Sub


    ' [3] ส่วนที่เพิ่มใหม่: ฟังก์ชันอ่านข้อมูลจากไฟล์ testdata.txt
    Private Sub LoadDataFromTextFile()
        ' แก้ Path ให้ตรงกับที่อยู่ไฟล์ testdata ของคุณ
        Dim filePath As String = "C:\testdata.txt"

        If Not File.Exists(filePath) Then
            MsgBox("File not found: " & filePath)
            Exit Sub
        End If

        Dim lines() As String = File.ReadAllLines(filePath)

        ' รีเซ็ตค่าตัวนับข้อมูล
        SPCDataNum = 0

        Graphsmallcount = 1

        ' เคลียร์ข้อมูลเก่าใน M_Data
        ReDim M_Data(0)

        ReDim M_Alarm(lines.Length)

        ' วนลูปอ่านข้อมูล (เริ่มที่ 1 เพื่อข้าม Header)
        For i As Integer = 1 To lines.Length - 1
            Dim line As String = lines(i)
            If line.Trim() = "" Then Continue For

            Dim cols() As String = line.Split(vbTab) ' ข้อมูลในไฟล์คั่นด้วย Tab

            If cols.Length > 20 Then
                SPCDataNum += 1

                ' 1. เก็บค่าลงตัวแปร M_Data (จำลองโครงสร้างข้อมูลให้เหมือนมาจาก Server)
                ' Format: ID, Date, Mean(X), Range(R), MR, Operator, Lot, ...
                ' Index อ้างอิงจากไฟล์ testdata: Date=6, Time=7, Mean=5, Range=20
                Dim rawRow As String = ""
                rawRow &= i & ","                 ' Index (ID)
                rawRow &= cols(6) & " " & cols(7) & "," ' Date Time
                rawRow &= cols(5) & ","           ' Mean (X) -> Column 5
                rawRow &= cols(20) & ","          ' Range (R) -> Column 20
                rawRow &= "0,"                    ' MR (สมมติเป็น 0)
                rawRow &= cols(21) & ","          ' Operator -> Column 21
                rawRow &= cols(0) & ","           ' Lot Number -> Column 0
                rawRow &= "Pass"                  ' Status

                ' ขยายขนาด Array และเก็บข้อมูล
                ReDim Preserve M_Data(SPCDataNum)
                M_Data(SPCDataNum) = rawRow

                ReDim Preserve M_Alarm(SPCDataNum)
                ReDim M_Alarm(SPCDataNum)(2)
                M_Alarm(SPCDataNum)(0) = "0,00000000"
                M_Alarm(SPCDataNum)(1) = "0,00000000"
                M_Alarm(SPCDataNum)(2) = "0,00000000"

                ' 2. เก็บค่าลง Buffer (เผื่อ GraphDisp ใช้ตัวนี้ด้วย)
                If SPCDataNum < MesureValueBuf.Length Then
                    MesureValueBuf(SPCDataNum) = cols(5) ' เก็บค่า Mean
                End If


                ' 3. ดึงค่า Control Limits (ใช้ค่าจากไฟล์)
                ' UCL=9, LCL=10, R_UCL=14, R_LCL=15
                Try
                    X_UCL = Val(cols(9))   ' Upper Control Limit (X-Bar)
                    X_LCL = Val(cols(10))  ' Lower Control Limit (X-Bar)
                    X_CL = (X_UCL + X_LCL) / 2 ' คำนวณเส้นกลางคร่าวๆ
                    R_UCL = Val(cols(14))  ' Upper Range Limit
                    R_LCL = Val(cols(15))  ' Lower Range Limit
                    R_CL = (R_UCL + R_LCL) / 2
                    X_USL = Val(cols(12))
                    X_LSL = Val(cols(13))
                Catch ex As Exception
                End Try
            End If
        Next

        Dim dt As New DataTable()
        Dim col_cMRdev As String = "cMRdev"
        Dim col_cMR As String = "cMR"
        Dim col_cSpcRule As String = "cSpcRule"

        dt.Columns.Add("cScl", GetType(Double))
        dt.Columns.Add("cTolerance", GetType(Double))
        dt.Columns.Add("cUnit", GetType(String))
        dt.Columns.Add("cLimitType", GetType(String))
        dt.Columns.Add("cUsl", GetType(Double))
        dt.Columns.Add("cLsl", GetType(Double))
        dt.Columns.Add("cXcl", GetType(Double))
        dt.Columns.Add("cXucl", GetType(Double))
        dt.Columns.Add("cXlcl", GetType(Double))
        dt.Columns.Add("cXdev", GetType(Double))
        dt.Columns.Add("cRucl", GetType(Double))
        dt.Columns.Add("cRcl", GetType(Double))
        dt.Columns.Add("cRdev", GetType(Double))
        dt.Columns.Add("cMRucl", GetType(Double))
        dt.Columns.Add("cMRcl", GetType(Double))
        dt.Columns.Add(col_cMRdev, GetType(Double))
        dt.Columns.Add(col_cMR, GetType(String))


        For k As Integer = 1 To 8
            dt.Columns.Add(col_cSpcRule & k, GetType(Boolean))
        Next

        dt.Columns.Add("cApprovalDate", GetType(DateTime))
        dt.Columns.Add("cMachineNo", GetType(String))
        dt.Columns.Add("cControlItem", GetType(String))


        Try
            Dim dr As DataRow = dt.NewRow()
            If Not dt.Columns.Contains(col_cMRdev) Then
                MsgBox("System Error: Column cMRdev failed to create!")
                Exit Sub
            End If
            dr("cScl") = (X_UCL + X_LCL) / 2
            dr("cTolerance") = Math.Abs(X_USL - X_LSL)
            If IsDBNull(dr("cTolerance")) OrElse dr("cTolerance") = 0 Then dr("cTolerance") = 1
            dr("cUnit") = "Unit"
            dr("cLimitType") = "UpperLower"
            dr("cUsl") = X_USL
            dr("cLsl") = X_LSL
            dr("cXcl") = X_CL
            dr("cXucl") = X_UCL
            dr("cXlcl") = X_LCL
            dr("cXdev") = 0
            dr("cRucl") = R_UCL
            dr("cRcl") = R_CL
            dr("cRdev") = 0
            dr("cMRucl") = 0
            dr("cMRcl") = 0
            dr(col_cMRdev) = 0
            dr(col_cMR) = "0"

            For k As Integer = 1 To 8
                dr(col_cSpcRule & k) = False
            Next

            dr("cApprovalDate") = DateTime.Now.AddYears(-10)
            dr("cMachineNo") = "TextFile"
            dr("cControlItem") = "Data"

            dt.Rows.Add(dr)
            PropertyTable = dt
        Catch ex As Exception
            Dim colist As String = ""
            For Each c As DataColumn In dt.Columns
                colist &= c.ColumnName & ", "
            Next
            MsgBox("Critical Error building table: " & ex.Message)
            Exit Sub
        End Try
        ReDim TreeName(0)
        TreeName(0) = "TextFile Data"
        ' ตั้งค่าการแสดงผลเบื้องต้น
        DispStartPosition = 0
        If SPCDataNum > 30 Then DispStartPosition = SPCDataNum - 30 ' ให้กราฟแสดงช่วงท้ายๆ

        ' สั่งวาดกราฟ
        GraphDisp()

        MsgBox("Loaded " & SPCDataNum & " points from file.")
    End Sub

    ' [2] ส่วนที่เพิ่มใหม่: Event Handler สำหรับปุ่ม ButtonLoad
    ' หมายเหตุ: ต้องสร้างปุ่มชื่อ ButtonLoad ใน Form Designer ก่อน
    Private Sub ButtonLoad_Click_1(sender As Object, e As EventArgs) Handles ButtonLoad.Click
        LoadDataFromTextFile()
    End Sub
End Class